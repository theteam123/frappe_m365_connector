"""Remembered sign-in for the Send-to-ERP Outlook add-in.

The add-in signs a user in with Microsoft 365, but that identity token lives for
about an hour and, in Outlook on the web, can only be renewed by opening a
sign-in window — so users were asked to sign in every time the panel opened.

After a successful Microsoft sign-in the add-in calls ``CreateAddinSession`` and
receives an opaque ERP-issued session token, which it keeps in the panel's own
storage and sends as ``Authorization: Bearer sterp_...`` from then on. The auth
hook in ``addin_sso`` recognises the prefix and validates it here.

Token discipline:
- Random, high-entropy, never stored — only its SHA-256 hash is kept.
- Sliding expiry (``SESSION_SLIDING_DAYS``) while the add-in is being used,
  with an absolute cap (``SESSION_ABSOLUTE_MAX_DAYS``) after which Microsoft
  sign-in is required again.
- Revocable from the desk (tick *Revoked* on the M365 Addin Session record),
  or by the user via the add-in's *Sign out*.
- Can only be minted by a request authenticated with a Microsoft token, so a
  leaked session token cannot mint replacements for itself.
"""
import hashlib
import secrets
from datetime import datetime
from typing import Optional

import frappe
from frappe import _
from frappe.utils import add_days, get_datetime, now_datetime


SESSION_DOCTYPE = "M365 Addin Session"
SESSION_TOKEN_PREFIX = "sterp_"
SESSION_TOKEN_BYTES = 32
SESSION_SLIDING_DAYS = 30
SESSION_ABSOLUTE_MAX_DAYS = 90
# The expiry is only rewritten once it has drifted at least this far behind the
# full sliding window, so a busy user does not cause a write on every request.
SESSION_SLIDE_GRANULARITY_DAYS = 1
MAX_PLATFORM_CHARS = 80

# How the current request was authenticated; set on frappe.local by the auth hook.
AUTH_METHOD_M365 = "m365"
AUTH_METHOD_SESSION = "addin_session"



@frappe.whitelist()
def CreateAddinSession(platform: str = "") -> dict:
	"""Mint a remembered sign-in for the current user.

	Only a request authenticated with a Microsoft 365 token may mint one: the
	user has just proven their identity to Microsoft, and an existing session
	token cannot be used to extend itself beyond its own limits.

	Args:
		platform: Free-text hint of the Outlook host (shown to admins only).

	Returns:
		{token, expires_on, user}. The token is shown to its owner exactly once
		and is never stored server-side (only its hash is).
	"""
	user = frappe.session.user
	if not user or user == "Guest":
		frappe.throw(_("Sign in with Microsoft 365 first."), frappe.AuthenticationError)

	if getattr(frappe.local, "addin_auth_method", None) != AUTH_METHOD_M365:
		frappe.throw(_("Sign in with Microsoft 365 first."), frappe.PermissionError)

	PurgeDeadSessions(user)

	token = SESSION_TOKEN_PREFIX + secrets.token_urlsafe(SESSION_TOKEN_BYTES)
	now = now_datetime()
	expiresOn = add_days(now, SESSION_SLIDING_DAYS)

	session = frappe.get_doc({
		"doctype": SESSION_DOCTYPE,
		"user": user,
		"platform": (platform or "").strip()[:MAX_PLATFORM_CHARS],
		"token_hash": HashToken(token),
		"expires_on": expiresOn,
		"last_used_on": now,
	})
	# Users have no create permission on sessions (in_create, System Manager
	# only); minting is gated on the Microsoft-token check above instead.
	session.insert(ignore_permissions=True)

	# The token is the credential being issued to its owner, not a stored secret
	# being disclosed (SEC-003): it exists nowhere else and is returned once.
	return {"token": token, "expires_on": expiresOn, "user": user}



@frappe.whitelist()
def RevokeAddinSession() -> dict:
	"""Sign the current add-in session out (the add-in's *Sign out* button).

	Returns:
		{revoked}: False when the request was not made with a session token.
	"""
	sessionName = getattr(frappe.local, "addin_session_name", None)
	if not sessionName:
		return {"revoked": False}

	frappe.db.set_value(SESSION_DOCTYPE, sessionName, "revoked", 1, update_modified=False)
	return {"revoked": True}



def ValidateSessionToken(token: str) -> Optional[str]:
	"""Resolve a session token to its user, sliding the expiry while in use.

	Called by the auth hook. The token is looked up by the hash of a 256-bit
	random value, so the comparison is an indexed equality on a digest rather
	than a byte-wise secret comparison (SEC-004 does not apply).

	Returns:
		The user id when the token is live, otherwise None.
	"""
	if not token or not token.startswith(SESSION_TOKEN_PREFIX):
		return None

	session = frappe.db.get_value(
		SESSION_DOCTYPE,
		{"token_hash": HashToken(token)},
		["name", "user", "expires_on", "revoked", "creation"],
		as_dict=True,
	)
	if not session or session.revoked:
		return None

	now = now_datetime()
	if get_datetime(session.expires_on) <= now:
		return None
	if add_days(get_datetime(session.creation), SESSION_ABSOLUTE_MAX_DAYS) <= now:
		return None
	if not frappe.db.get_value("User", {"name": session.user, "enabled": 1}, "name"):
		return None

	SlideExpiry(session, now)
	frappe.local.addin_session_name = session.name
	return session.user



def SlideExpiry(session: dict, now: datetime) -> None:
	"""Push the expiry out to a full window again, at most once a day."""
	fullWindowEnd = add_days(now, SESSION_SLIDING_DAYS)
	rewriteThreshold = add_days(fullWindowEnd, -SESSION_SLIDE_GRANULARITY_DAYS)
	if get_datetime(session.expires_on) > rewriteThreshold:
		return

	frappe.db.set_value(
		SESSION_DOCTYPE,
		session.name,
		{"expires_on": fullWindowEnd, "last_used_on": now},
		update_modified=False,
	)
	# The auth hook runs before the request body, and GET requests are never
	# committed by the framework, so persist the slide explicitly.
	frappe.db.commit()



def PurgeDeadSessions(user: str) -> None:
	"""Delete the user's expired and revoked sessions (housekeeping on sign-in)."""
	frappe.db.delete(SESSION_DOCTYPE, {"user": user, "expires_on": ["<", now_datetime()]})
	frappe.db.delete(SESSION_DOCTYPE, {"user": user, "revoked": 1})



def HashToken(token: str) -> str:
	"""SHA-256 hex digest of a session token — the only form ever stored."""
	return hashlib.sha256(token.encode("utf-8")).hexdigest()
