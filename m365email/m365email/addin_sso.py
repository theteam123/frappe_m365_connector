"""Microsoft 365 single sign-on for the Send-to-ERP Outlook add-in.

Registered as a Frappe ``auth_hook``. When a request carries an Azure AD bearer
token (obtained by the add-in via Office SSO / MSAL), this validates the token
against the issuing tenant's public signing keys and authenticates the request
as the matching Frappe user. Identity comes from the user's existing Microsoft
365 sign-in — no ERP passwords are involved.

Only Microsoft-issued JWTs whose audience matches a configured M365 app are
handled; anything else is ignored so normal Frappe auth still applies.
"""
import frappe
from typing import Optional

import jwt
from jwt import PyJWKClient


BEARER_PREFIX = "Bearer "
ACCEPTED_ALGORITHMS = ["RS256"]
JWKS_URI_TEMPLATE = "https://login.microsoftonline.com/{tenant}/discovery/v2.0/keys"
ACCEPTED_ISSUER_HOSTS = ("login.microsoftonline.com", "sts.windows.net")
JWKS_CLIENT_CACHE: dict = {}  # tenant_id -> PyJWKClient (per-process)
USER_EMAIL_CLAIMS = ("preferred_username", "upn", "email", "unique_name")



def ValidateAddinToken() -> None:
	"""Auth hook: authenticate a request bearing an Azure AD add-in token."""
	if frappe.local.login_manager.user not in ("", "Guest", None):
		return

	token = _ExtractBearerToken()
	if not token:
		return

	apps = _GetConfiguredApps()
	if not apps:
		return

	claims = _ValidateToken(token, apps)
	if not claims:
		return

	user = _ResolveUser(claims)
	if not user:
		return

	# set_user resets form_dict, so preserve the request args around it
	# (mirrors frappe's own validate_api_key_secret).
	formDict = frappe.local.form_dict
	frappe.set_user(user)
	frappe.local.login_manager.user = user
	frappe.local.form_dict = formDict



def _ExtractBearerToken() -> Optional[str]:
	"""Return the bearer token from the Authorization header, if JWT-shaped."""
	header = frappe.get_request_header("Authorization") or ""
	if not header.startswith(BEARER_PREFIX):
		return None

	token = header[len(BEARER_PREFIX):].strip()
	# A JWT has three dot-separated segments; ignore anything else (e.g. opaque tokens).
	if token.count(".") != 2:
		return None

	return token



def _GetConfiguredApps() -> list[dict]:
	"""Return enabled M365 apps as {tenant_id, client_id} for token validation."""
	apps = frappe.get_all(
		"M365 Email Service Principal Settings",
		filters={"enabled": 1},
		fields=["tenant_id", "client_id"],
	)
	return [a for a in apps if a.get("tenant_id") and a.get("client_id")]



def _ValidateToken(token: str, apps: list[dict]) -> Optional[dict]:
	"""Validate a Microsoft JWT against the configured apps.

	Args:
		token: The raw JWT from the Authorization header.
		apps: Enabled M365 apps providing accepted tenants and audiences.

	Returns:
		The decoded claims if valid, otherwise None.
	"""
	unverified = _DecodeUnverified(token)
	if not unverified:
		return None

	tenantId = _MatchTenant(unverified.get("iss", ""), apps)
	if not tenantId:
		return None

	signingKey = _GetSigningKey(tenantId, token)
	if not signingKey:
		return None

	try:
		claims = jwt.decode(
			token,
			signingKey,
			algorithms=ACCEPTED_ALGORITHMS,
			options={"verify_aud": False, "require": ["exp", "iss"]},
		)
	except jwt.InvalidTokenError:
		return None

	if not _IsAcceptedAudience(claims.get("aud", ""), apps):
		return None

	return claims



def _DecodeUnverified(token: str) -> Optional[dict]:
	"""Decode JWT claims without signature verification (for routing only)."""
	try:
		return jwt.decode(token, options={"verify_signature": False})
	except jwt.InvalidTokenError:
		return None



def _MatchTenant(issuer: str, apps: list[dict]) -> Optional[str]:
	"""Return the tenant id if the issuer is a Microsoft issuer for a known app."""
	if not any(host in issuer for host in ACCEPTED_ISSUER_HOSTS):
		return None

	for app in apps:
		tenantId = app["tenant_id"]
		if tenantId in issuer:
			return tenantId

	return None



def _GetSigningKey(tenantId: str, token: str):
	"""Return the RSA signing key for the token from the tenant's JWKS endpoint."""
	client = JWKS_CLIENT_CACHE.get(tenantId)
	if client is None:
		client = PyJWKClient(JWKS_URI_TEMPLATE.format(tenant=tenantId), cache_keys=True)
		JWKS_CLIENT_CACHE[tenantId] = client

	try:
		return client.get_signing_key_from_jwt(token).key
	except Exception:
		return None



def _IsAcceptedAudience(audience: str, apps: list[dict]) -> bool:
	"""Check the token audience matches a configured app's client id.

	Office SSO tokens carry an audience of either the bare client id, the
	``api://<client_id>`` form, or the domain-qualified Application ID URI
	``api://<host>/<client_id>``; accept all three forms.
	"""
	if not audience:
		return False

	for app in apps:
		clientId = app["client_id"]
		if audience in (clientId, f"api://{clientId}") or audience.endswith(f"/{clientId}"):
			return True

	return False



def _ResolveUser(claims: dict) -> Optional[str]:
	"""Map validated token claims to an enabled Frappe user via email."""
	email = _ExtractEmail(claims)
	if not email:
		return None

	user = frappe.db.get_value("User", {"email": email, "enabled": 1}, "name")
	return user



def _ExtractEmail(claims: dict) -> Optional[str]:
	"""Return the first usable email-like claim from the token."""
	for claimName in USER_EMAIL_CLAIMS:
		value = claims.get(claimName)
		if value and "@" in value:
			return value.lower()

	return None
