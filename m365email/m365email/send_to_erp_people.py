"""Send to ERP: who an email is really from, and which Contact it belongs with.

Paul, 2026-11-07: emails are often forwarded internally, so the add-in's sender is then a
colleague, not the person the email is about. Every Send to ERP option that ties an email to
a person now pre-fills a best guess and lets the user pick any other Contact:

- the original sender, read from the forwarded-message header in the body
  ("From: Name <email>"), skipping our own staff, else the actual sender;
- the existing Contact holding that address, if the user may read it;
- for procurement, the supplier: the Contact's company, else the email domain.

The body is attacker-controlled. It is reduced to plain text, capped, and read with one
bounded pattern; addresses are only ever looked up as bound filter values.
"""

import html
import re
from types import ModuleType
from typing import Optional

import frappe
from frappe import _

CONTACT = "Contact"
CONTACT_EMAIL = "Contact Email"
USER = "User"
SYSTEM_USER = "System User"
ORGANISATION = "Organisation"

MAX_BODY_CHARS = 50000
MAX_HTML_CHARS = 2_000_000  # Outlook HTML is heavy; the headers sit near the top
LINE_BREAK_TAG_PATTERN = re.compile(r"<(?:br|/p|/div|/tr|/li|hr)\b[^<>]{0,200}>", re.IGNORECASE)
# NOTE: Our own tag remover, not frappe.utils.strip_html: that one's "<[^>]*>" retries at
#       every unmatched "<", so a body of 2 MB of "<" took seconds (Codex review, 2026-11-07).
#       A tag here cannot contain "<", so each attempt stops at the next "<" or ">" and the
#       whole pass is linear; it has no length cap, so a long attribute is still removed.
ANY_TAG_PATTERN = re.compile(r"<[^<>]*>")
MAX_FORWARD_HEADERS = 10

# NOTE: Outlook's forwarded-message header, once the HTML is stripped:
#       "From: Jun Adolfo <jun@example.com>" or "From: jun@example.com". Every run is bounded
#       so a crafted body cannot make matching slow.
FORWARD_FROM_PATTERN = re.compile(
	r"^[ \t>*]{0,6}From:\*{0,2}[ \t]{0,6}"
	r"(?:(?P<name>[^<\n\r@]{1,120}?)[ \t]{0,6}<[ \t]{0,3})?"
	r"(?:mailto:)?(?P<email>[A-Za-z0-9._%+'-]{1,64}"
	r"@[A-Za-z0-9-]{1,63}(?:\.[A-Za-z0-9-]{1,63}){1,6})",
	re.IGNORECASE | re.MULTILINE,
)



@frappe.whitelist()
def GuessPeople(sender_name: str = "", sender_email: str = "", body_html: str = "") -> dict:
	"""The best guess at who an email is from, for the add-in to pre-fill.

	Returns:
		{"person": {name, email, is_forwarded, is_internal},
		 "contact": {value, label, sublabel} or None,
		 "supplier": {value, label} or None}
	"""
	person = GuessPerson(sender_name, sender_email, body_html)
	contactName = FindContactForEmail(person["email"])
	return {
		"person": person,
		"contact": FormatContact(contactName),
		"supplier": GuessSupplier(contactName, person["email"]),
	}



def GuessPerson(senderName: str, senderEmail: str, bodyHTML: str) -> dict:
	"""The sender, unless they are a colleague forwarding someone else's email."""
	sender = {"name": (senderName or "").strip(), "email": NormaliseEmail(senderEmail)}
	if not IsInternalSender(sender["email"]):
		return dict(sender, is_forwarded=False, is_internal=False)

	for forwarded in ReadForwardedSenders(bodyHTML):
		if not IsInternalSender(forwarded["email"]):
			return dict(forwarded, is_forwarded=True, is_internal=False)

	# NOTE: Only colleagues in the chain: say so, so the add-in asks rather than guessing.
	return dict(sender, is_forwarded=False, is_internal=bool(sender["email"]))



def ReadForwardedSenders(bodyHTML: str) -> list[dict]:
	"""Every "From:" in the forwarded-message headers, newest first."""
	# NOTE: Line breaks and block ends become new lines first, so a "From:" that follows a
	#       <br> or a closing paragraph starts a line, as the pattern requires.
	withLines = LINE_BREAK_TAG_PATTERN.sub("\n", (bodyHTML or "")[:MAX_HTML_CHARS])
	text = html.unescape(ANY_TAG_PATTERN.sub("", withLines))[:MAX_BODY_CHARS]
	senders = []
	for match in FORWARD_FROM_PATTERN.finditer(text):
		email = NormaliseEmail(match.group("email"))
		if any(sender["email"] == email for sender in senders):
			continue
		senders.append({"name": (match.group("name") or "").strip(" \"'"), "email": email})
		if len(senders) >= MAX_FORWARD_HEADERS:
			break
	return senders



def IsInternalSender(senderEmail: str) -> bool:
	"""True when the address belongs to one of our own ERP users (a colleague)."""
	emailAddress = NormaliseEmail(senderEmail)
	if not emailAddress:
		return False
	# NOTE: Desk (System) users only: a supplier with a website login is not staff.
	if frappe.db.exists(USER, {"name": emailAddress, "user_type": SYSTEM_USER}):
		return True
	return bool(frappe.db.exists(USER, {"email": emailAddress, "user_type": SYSTEM_USER}))



def IsColleagueContact(contactName: Optional[str]) -> bool:
	"""True when any address on the Contact belongs to one of our own staff."""
	if not contactName:
		return False
	addresses = frappe.get_all(CONTACT_EMAIL, filters={"parent": contactName}, pluck="email_id")
	addresses.append(frappe.db.get_value(CONTACT, contactName, "email_id") or "")
	return any(IsInternalSender(address) for address in addresses if address)



def FindContactForEmail(email: str) -> Optional[str]:
	"""The Contact holding this address, if it is not a colleague's and the user may read it."""
	emailAddress = NormaliseEmail(email)
	if not emailAddress or IsInternalSender(emailAddress):
		return None

	contactName = frappe.db.get_value(CONTACT_EMAIL, {"email_id": emailAddress}, "parent")
	if not contactName or not frappe.has_permission(CONTACT, "read", contactName):
		return None
	if IsColleagueContact(contactName):
		return None
	return contactName



def ResolveChosenContact(contactName: str) -> Optional[str]:
	"""Check a Contact the user picked: it must exist and be readable. "" means none."""
	if not contactName:
		return None
	# NOTE: One message for "missing" and "not allowed", so it never reveals that a contact exists.
	isUsable = frappe.db.exists(CONTACT, contactName)
	if not isUsable or not frappe.has_permission(CONTACT, "read", contactName):
		frappe.throw(_("The contact you picked cannot be used."), frappe.PermissionError)
	return contactName



def AssertNotColleague(contactName: Optional[str]) -> None:
	"""Refuse a colleague's Contact where the supplier's contact is meant (quote, invoice)."""
	if IsColleagueContact(contactName):
		frappe.throw(_("Pick the supplier's contact, not one of our own staff."))



def HasAnyContactForEmail(email: str) -> bool:
	"""True when any Contact holds this address, readable or not."""
	emailAddress = NormaliseEmail(email)
	return bool(emailAddress) and bool(frappe.db.exists(CONTACT_EMAIL, {"email_id": emailAddress}))



def FormatContact(contactName: Optional[str]) -> Optional[dict]:
	"""A Contact as the add-in shows it: full name, with the email beneath."""
	if not contactName:
		return None

	row = frappe.db.get_value(
		CONTACT, contactName, ["full_name", "first_name", "last_name", "email_id"], as_dict=True,
	)
	if not row:
		return None
	label = row.full_name or " ".join(part for part in (row.first_name, row.last_name) if part)
	email = row.email_id or ""
	return {"value": contactName, "label": label or email, "sublabel": email}



def GuessSupplier(contactName: Optional[str], email: str) -> Optional[dict]:
	"""The supplier this person works for: the Contact's company, else their email domain.

	Only on a site with the erp app, only Organisations ticked as Supplier, and only when
	the user may read them. None whenever it cannot be told.
	"""
	if "erp" not in frappe.get_installed_apps():
		return None
	if not frappe.has_permission(ORGANISATION, "read"):
		return None

	from erp.procurement.doctype.supplier_quote.extractors import (
		supplier_detect as supplierDetect,
	)

	organisation = None
	if contactName:
		company = frappe.db.get_value(CONTACT, contactName, "company_name")
		if company and frappe.db.get_value(ORGANISATION, company, "supplier"):
			organisation = company
	if not organisation:
		organisation = SupplierForDomain(supplierDetect, email)
	if not organisation or not frappe.has_permission(ORGANISATION, "read", organisation):
		return None

	label = frappe.db.get_value(ORGANISATION, organisation, "organisation_name")
	return {"value": organisation, "label": label or ""}



def SupplierForDomain(supplierDetect: ModuleType, email: str) -> Optional[str]:
	"""The one supplier using this email domain, or None (public mail, none, or several)."""
	domain = email.rpartition("@")[2] if "@" in email else ""
	if not domain or domain in supplierDetect.PUBLIC_EMAIL_DOMAINS:
		return None

	matches = supplierDetect.FindSuppliersByEmailDomain({domain})
	return matches[0] if len(matches) == 1 else None



def LinkEmailToContact(communicationName: Optional[str], contactName: Optional[str]) -> None:
	"""Show the filed email on the chosen Contact's timeline too."""
	if not communicationName or not contactName:
		return

	communication = frappe.get_doc("Communication", communicationName)
	communication.add_link(CONTACT, contactName)
	# NOTE: The Contact was checked readable by ResolveChosenContact, and the email itself was
	#       filed under the user's write permission on its record.
	communication.save(ignore_permissions=True)



def DescribeContact(contactName: Optional[str]) -> str:
	""""Name (email)" for a Contact, for a ticket's description."""
	formatted = FormatContact(contactName)
	if not formatted:
		return ""
	if formatted["sublabel"] and formatted["sublabel"] != formatted["label"]:
		return f"{formatted['label']} ({formatted['sublabel']})"
	return formatted["label"]



def NormaliseEmail(email: str) -> str:
	return (email or "").strip().lower()
