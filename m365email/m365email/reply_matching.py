"""File a reply on the record its original email was sent from.

Every email we send carries a Message-ID, and a reply repeats it in In-Reply-To, so no
record key needs to appear in the message. Microsoft keeps our Message-ID on send.
"""

import frappe

from m365email.m365email.graph_api import make_graph_request

IN_REPLY_TO_HEADER = "in-reply-to"
MAX_MESSAGE_ID_LENGTH = 255
# NOTE: Only mail we sent counts. A received Communication holds an outsider's Message-ID,
# so an outsider could otherwise quote it to file onto that record.
SENT_MAIL_SOURCES = (
	("Email Queue", {}),
	("Communication", {"sent_or_received": "Sent"}),
)



def GetInReplyTo(userEmail: str, messageID: str, accessToken: str) -> str | None:
	"""The bare Message-ID this message replies to, or None. The delta sync omits headers."""
	endpoint = f"/users/{userEmail}/messages/{messageID}"
	params = {"$select": "internetMessageHeaders"}
	details = make_graph_request(endpoint, accessToken, params=params) or {}
	for header in details.get("internetMessageHeaders") or []:
		if (header.get("name") or "").lower() != IN_REPLY_TO_HEADER:
			continue
		value = (header.get("value") or "").strip().split(" ")[0].strip("<>")
		return value if value and len(value) <= MAX_MESSAGE_ID_LENGTH else None
	return None



def FindSentMailReference(inReplyTo: str | None) -> tuple:
	"""The (doctype, name) our sent email belongs to, or (None, None)."""
	if not inReplyTo:
		return None, None

	idForms = [inReplyTo, f"<{inReplyTo}>"]
	for doctype, sentFilter in SENT_MAIL_SOURCES:
		rows = frappe.get_all(
			doctype,
			filters={"message_id": ["in", idForms], "reference_name": ["is", "set"], **sentFilter},
			fields=["reference_doctype", "reference_name"],
			limit=1,
		)
		if not rows:
			continue
		row = rows[0]
		if row.reference_doctype and frappe.db.exists(row.reference_doctype, row.reference_name):
			return row.reference_doctype, row.reference_name
	return None, None



def GetReplyReference(userEmail: str, messageID: str, accessToken: str) -> tuple:
	"""Linking must never break the sync, so any failure leaves the email unlinked."""
	try:
		inReplyTo = GetInReplyTo(userEmail, messageID, accessToken)
		return FindSentMailReference(inReplyTo)
	except Exception as error:
		frappe.logger("m365email").debug(f"GetReplyReference failed: {type(error).__name__}")
		return None, None
