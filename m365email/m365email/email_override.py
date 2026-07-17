# Copyright (c) 2025, Zeke Tierney and contributors
# For license information, please see license.txt

"""
Override Frappe's core email sending to support M365
This allows sending emails without requiring a default Email Account
"""

import json
import frappe
from frappe import _
from frappe.email.email_body import get_message_id
from frappe.utils import (
	cint,
	get_formatted_email,
	get_string_between,
	list_to_str,
)
from typing import TYPE_CHECKING

if TYPE_CHECKING:
	from frappe.core.doctype.communication.communication import Communication


@frappe.whitelist()
def make(
	doctype=None,
	name=None,
	content=None,
	subject=None,
	sent_or_received="Sent",
	sender=None,
	sender_full_name=None,
	recipients=None,
	communication_medium="Email",
	send_email=False,
	print_html=None,
	print_format=None,
	attachments=None,
	send_me_a_copy=False,
	cc=None,
	bcc=None,
	read_receipt=None,
	print_letterhead=True,
	email_template=None,
	communication_type=None,
	send_after=None,
	print_language=None,
	now=False,
	**kwargs,
) -> dict[str, str]:
	"""
	Override of frappe.core.doctype.communication.email.make
	Supports M365 email sending without requiring a default Email Account
	"""
	if kwargs:
		from frappe.utils.commands import warn

		warn(
			f"Options {kwargs} used in frappe.core.doctype.communication.email.make "
			"are deprecated or unsupported",
			category=DeprecationWarning,
		)

	if doctype and name and not frappe.has_permission(doctype=doctype, ptype="email", doc=name):
		frappe.throw(
			_(
				"You are not allowed to send emails related to {0} {1}. "
				"Your account is missing the 'Email' permission on {0} — "
				"ask your administrator to grant it on a role you hold."
			).format(doctype, name),
			frappe.PermissionError,
			title=_("Email Permission Required"),
		)

	return _make(
		doctype=doctype,
		name=name,
		content=content,
		subject=subject,
		sent_or_received=sent_or_received,
		sender=sender,
		sender_full_name=sender_full_name,
		recipients=recipients,
		communication_medium=communication_medium,
		send_email=send_email,
		print_html=print_html,
		print_format=print_format,
		attachments=attachments,
		send_me_a_copy=cint(send_me_a_copy),
		cc=cc,
		bcc=bcc,
		read_receipt=cint(read_receipt),
		print_letterhead=print_letterhead,
		email_template=email_template,
		communication_type=communication_type,
		add_signature=False,
		send_after=send_after,
		print_language=print_language,
		now=now,
	)


def _make(
	doctype=None,
	name=None,
	content=None,
	subject=None,
	sent_or_received="Sent",
	sender=None,
	sender_full_name=None,
	recipients=None,
	communication_medium="Email",
	send_email=False,
	print_html=None,
	print_format=None,
	attachments=None,
	send_me_a_copy=False,
	cc=None,
	bcc=None,
	read_receipt=None,
	print_letterhead=True,
	email_template=None,
	communication_type=None,
	add_signature=True,
	send_after=None,
	print_language=None,
	now=False,
) -> dict[str, str]:
	"""
	Override of frappe.core.doctype.communication.email._make
	Bypasses email account check when M365 sending is available
	"""
	from m365email.m365email.send import can_send_via_m365
	from frappe.core.doctype.communication.email import add_attachments

	sender = sender or get_formatted_email(frappe.session.user)
	recipients = list_to_str(recipients) if isinstance(recipients, list) else recipients
	cc = list_to_str(cc) if isinstance(cc, list) else cc
	bcc = list_to_str(bcc) if isinstance(bcc, list) else bcc

	comm: "Communication" = frappe.get_doc(
		{
			"doctype": "Communication",
			"subject": subject,
			"content": content,
			"sender": sender,
			"sender_full_name": sender_full_name,
			"recipients": recipients,
			"cc": cc or None,
			"bcc": bcc or None,
			"communication_medium": communication_medium,
			"sent_or_received": sent_or_received,
			"reference_doctype": doctype,
			"reference_name": name,
			"email_template": email_template,
			"message_id": get_string_between("<", get_message_id(), ">"),
			"read_receipt": read_receipt,
			"has_attachment": 1 if attachments else 0,
			"communication_type": communication_type,
			"send_after": send_after,
		}
	)
	comm.flags.skip_add_signature = not add_signature
	comm.insert(ignore_permissions=True)

	# if not committed, delayed task doesn't find the communication
	if attachments:
		if isinstance(attachments, str):
			attachments = json.loads(attachments)
		add_attachments(comm.name, attachments)

	if cint(send_email):
		# Check if M365 sending is available
		m365_available = can_send_via_m365()
		
		# Only check for email account if M365 is NOT available
		if not m365_available and not comm.get_outgoing_email_account():
			frappe.throw(
				_(
					"Unable to send mail because of a missing email account. Please setup default Email Account from Settings > Email Account"
				),
				exc=frappe.OutgoingEmailError,
			)

		comm.send_email(
			print_html=print_html,
			print_format=print_format,
			send_me_a_copy=send_me_a_copy,
			print_letterhead=print_letterhead,
			print_language=print_language,
			now=now,
		)

		# User-initiated sends should deliver immediately, not wait for the
		# scheduled queue flush. Scheduled (send_after) emails stay queued.
		if not now and not send_after:
			frappe.enqueue(
				"m365email.m365email.email_override.FlushCommunicationEmailQueue",
				queue="short",
				enqueue_after_commit=True,
				communicationName=comm.name,
			)

	emails_not_sent_to = comm.exclude_emails_list(include_sender=send_me_a_copy)

	return {"name": comm.name, "emails_not_sent_to": ", ".join(emails_not_sent_to)}


def FlushCommunicationEmailQueue(communicationName: str) -> None:
	"""Send a communication's queued emails now instead of at the next scheduled flush."""
	queuedNames = frappe.get_all(
		"Email Queue",
		filters={"communication": communicationName, "status": "Not Sent"},
		pluck="name",
	)
	for queuedName in queuedNames:
		queuedDoc = frappe.get_doc("Email Queue", queuedName)
		# Re-check under this job in case the scheduled flush got there first.
		if queuedDoc.status != "Not Sent":
			continue
		try:
			queuedDoc.send()
		except Exception:
			frappe.log_error(
				title=f"Instant email send failed: {queuedName}",
				message=frappe.get_traceback(),
			)

