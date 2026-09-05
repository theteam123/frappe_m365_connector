# Copyright (c) 2026, TierneyMorris Pty Ltd and contributors
# For license information, please see license.txt

"""
Create outgoing M365 Email Accounts for people whose mailbox exists in Microsoft 365.

Two callers share this: the lazy path in send.py (a user holding the 'M365 User' role
sends for the first time) and any app that walks its own people list and asks for an
account once it knows the mailbox is live (the hr app's daily mailbox sync does this).
"""

import frappe
from frappe import _
from frappe.model.document import Document
from m365email.m365email.auth import get_access_token
from m365email.m365email.graph_api import MailboxExists

SERVICE_PRINCIPAL_DOCTYPE = "M365 Email Service Principal Settings"
EMAIL_ACCOUNT_DOCTYPE = "Email Account"
USER_DOCTYPE = "User"
M365_SERVICE = "M365"
USER_MAILBOX_TYPE = "User Mailbox"
SENDER_PLACEHOLDER = "<!--Sender-->"

PROVISION_CREATED = "created"
PROVISION_EXISTS = "exists"
PROVISION_NO_MAILBOX = "no-mailbox"
PROVISION_NOT_CONFIGURED = "not-configured"
PROVISION_NO_USER = "no-user"



def NormalizeDomain(domain: str) -> str:
	"""Lower-case with no leading @, so "@Example.com" and "example.com" compare equal."""
	return (domain or "").strip().lower().lstrip("@")



def ExtractDomain(emailAddress: str) -> str:
	address = (emailAddress or "").strip().lower()
	if "@" not in address:
		return ""
	return address.rsplit("@", 1)[1]



def FindProvisioningServicePrincipal(domain: str) -> Document | None:
	"""The enabled Service Principal that auto-provisions this domain, or None.

	NOTE: Compared in Python rather than filtered in SQL so the Domain field tolerates a
	leading @ or mixed case, which is how people type it.
	"""
	wantedDomain = NormalizeDomain(domain)
	if not wantedDomain:
		return None

	filters = {"enabled": 1, "enable_auto_provision": 1}
	fields = ["name", "domain"]
	candidates = frappe.get_all(SERVICE_PRINCIPAL_DOCTYPE, filters=filters, fields=fields)
	for candidate in candidates:
		if NormalizeDomain(candidate.domain) == wantedDomain:
			return frappe.get_doc(SERVICE_PRINCIPAL_DOCTYPE, candidate.name)

	return None



def ProbeMailbox(emailAddress: str, servicePrincipalName: str) -> bool:
	"""Ask Microsoft 365 whether a mailbox exists at this address.

	Raises on anything other than a clear yes or no; callers must treat an exception as
	"unknown", never as "no mailbox".
	"""
	accessToken = get_access_token(servicePrincipalName)
	return MailboxExists(emailAddress, accessToken)



def BuildDefaultSignature(servicePrincipal: Document, userFullName: str) -> str:
	if not servicePrincipal.default_footer:
		return ""
	return servicePrincipal.default_footer.replace(SENDER_PLACEHOLDER, userFullName)



def ResolveAccountName(userFullName: str, emailAddress: str) -> str:
	"""The person's name, as the hand-made accounts are named; name plus address only
	when two people share a name."""
	if not frappe.db.exists(EMAIL_ACCOUNT_DOCTYPE, userFullName):
		return userFullName
	return f"{userFullName} ({emailAddress})"



def CreateM365Account(servicePrincipal: Document, user: str, emailAddress: str) -> Document:
	"""Insert an outgoing-only M365 Email Account owned by this user. Does not commit."""
	if frappe.db.exists(EMAIL_ACCOUNT_DOCTYPE, {"email_id": emailAddress}):
		frappe.throw(_("An Email Account already exists for {0}").format(emailAddress))

	userDoc = frappe.get_doc(USER_DOCTYPE, user)
	userFullName = userDoc.full_name or userDoc.first_name or user

	account = frappe.get_doc({
		"doctype": EMAIL_ACCOUNT_DOCTYPE,
		"email_account_name": ResolveAccountName(userFullName, emailAddress),
		"service": M365_SERVICE,
		"email_id": emailAddress,
		"linked_user": user,
		"m365_account_type": USER_MAILBOX_TYPE,
		"m365_service_principal": servicePrincipal.name,
		"enable_incoming": 0,
		"enable_outgoing": 1,
		# NOTE: default_outgoing is deliberately not set. It is a site-wide flag and Frappe
		# permits only one holder, so setting it per user would strip it from the site's
		# real default account on every provision.
		"signature": BuildDefaultSignature(servicePrincipal, userFullName),
	})
	account.insert(ignore_permissions=True)
	return account



def EnsureM365AccountForUser(
	user: str, emailAddress: str = "", isMailboxLive: bool | None = None
) -> str:
	"""Give a user an outgoing M365 Email Account if their mailbox exists. Idempotent.

	Args:
		user: the User name (login ID)
		emailAddress: the mailbox to send from; defaults to the login ID
		isMailboxLive: pass the answer when the caller has already asked Microsoft 365,
			so the mailbox is not probed twice in one run

	Returns one of the PROVISION_* outcomes.
	"""
	userRow = frappe.db.get_value(USER_DOCTYPE, user, ["name", "enabled"], as_dict=True)
	if not userRow or not userRow.enabled:
		return PROVISION_NO_USER

	address = (emailAddress or userRow.name).strip().lower()
	if frappe.db.exists(EMAIL_ACCOUNT_DOCTYPE, {"email_id": address}):
		return PROVISION_EXISTS

	servicePrincipal = FindProvisioningServicePrincipal(ExtractDomain(address))
	if servicePrincipal is None:
		return PROVISION_NOT_CONFIGURED

	if isMailboxLive is None:
		isMailboxLive = ProbeMailbox(address, servicePrincipal.name)
	if not isMailboxLive:
		return PROVISION_NO_MAILBOX

	CreateM365Account(servicePrincipal, user, address)
	return PROVISION_CREATED
