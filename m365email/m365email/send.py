# Copyright (c) 2025, Zeke Tierney and contributors
# For license information, please see license.txt

"""
Email sending module for M365 Email Integration
Handles sending emails via Microsoft Graph API instead of SMTP
"""

import frappe
from frappe import _
from email import message_from_bytes, policy
from email.utils import parseaddr
from frappe.email.doctype.email_queue.email_queue import SendMailContext
from m365email.m365email.auth import get_access_token
from m365email.m365email.graph_api import send_mime_as_user
from m365email.m365email.provisioning import (
	CreateM365Account,
	ExtractDomain,
	FindProvisioningServicePrincipal,
)


def auto_provision_m365_account(sender_email):
	"""
	Auto-provision Email Account with service='M365' for eligible users

	Eligibility criteria:
	- User has 'M365 User' role
	- User's email domain matches a Service Principal's domain
	- Service Principal has auto-provisioning enabled

	NOTE: The account itself is built by provisioning.CreateM365Account, shared with the
	proactive path other apps call once they know a mailbox is live.

	Args:
		sender_email: Email address of the sender

	Returns:
		Email Account doc if auto-provisioned, None otherwise
	"""
	try:
		user = frappe.db.get_value("User", {"email": sender_email}, "name")
		if not user:
			return None

		if not frappe.db.exists("Has Role", {"parent": user, "role": "M365 User"}):
			return None

		servicePrincipal = FindProvisioningServicePrincipal(ExtractDomain(sender_email))
		if servicePrincipal is None:
			return None

		account = CreateM365Account(servicePrincipal, user, sender_email)
		frappe.db.commit()

		provisionedMessage = f"M365 Email: Auto-provisioned account for {sender_email}"
		print(f"{provisionedMessage} using Service Principal {servicePrincipal.name}")
		return account

	except Exception as e:
		frappe.log_error(
			title=f"M365 Auto-Provision Failed: {sender_email}",
			message=f"Error: {str(e)}\n\nTraceback: {frappe.get_traceback()}"
		)
		return None



def GetUserDefaultEmailAccount(user: str = None):
	"""
	Get the user's default email account for sending.

	Finds an M365 Email Account linked to the user with default_outgoing enabled.

	Args:
		user: User ID. Defaults to current session user.

	Returns:
		Email Account doc if found, None otherwise
	"""
	if not user:
		user = frappe.session.user

	if not user or user == "Guest":
		return None

	# NOTE: Find email account linked to this user with user_default checked
	accountName = frappe.db.get_value(
		"Email Account",
		{
			"linked_user": user,
			"user_default": 1,
			"enable_outgoing": 1
		},
		"name"
	)

	if accountName:
		return frappe.get_doc("Email Account", accountName)

	# NOTE: Fall back to any account linked to the user with outgoing enabled
	accountName = frappe.db.get_value(
		"Email Account",
		{
			"linked_user": user,
			"enable_outgoing": 1
		},
		"name"
	)

	if accountName:
		return frappe.get_doc("Email Account", accountName)

	return None



def get_sending_account_for_sender(sender_email):
	"""
	Get the Email Account for a specific sender email.

	Checks in order:
	1. Email Account with service='M365' matching the explicitly requested sender
	2. User's linked email account (via linked_user field)
	3. Auto-provision if eligible
	4. Default M365 outgoing account

	NOTE (2026-08-25): step 1 used to come after step 2, so a user with their own
	linked mailbox could never send as a shared mailbox they had picked in the From
	field; the choice was silently overridden. An explicit, enabled sender now wins;
	everything without one behaves exactly as before.

	Args:
		sender_email: Email address of the sender

	Returns:
		tuple: (account_doc, is_matched)
			- account_doc: Email Account doc with service='M365'
			- is_matched: True if sender matched an account
	"""
	from email.utils import parseaddr

	# Parse email address from "Name <email>" format
	if sender_email:
		_, sender_email = parseaddr(sender_email)

	# Respect an explicitly requested sender that is a configured, enabled M365 mailbox
	if sender_email and '@' in sender_email:
		account_name = frappe.db.get_value(
			"Email Account",
			{"service": "M365", "email_id": sender_email, "enable_outgoing": 1}
		)
		if account_name:
			return frappe.get_doc("Email Account", account_name), True

	# NOTE: Then the user's linked email account
	userLinkedAccount = GetUserDefaultEmailAccount()
	if userLinkedAccount and userLinkedAccount.service == "M365":
		return userLinkedAccount, True

	# Try auto-provisioning if no account matched the sender
	if sender_email and '@' in sender_email:
		auto_provisioned_account = auto_provision_m365_account(sender_email)
		if auto_provisioned_account:
			return auto_provisioned_account, True

	# Fall back to default outgoing account
	account_name = frappe.db.get_value(
		"Email Account",
		{"service": "M365", "default_outgoing": 1, "enable_outgoing": 1}
	)
	if account_name:
		return frappe.get_doc("Email Account", account_name), False

	return None, False


def can_send_via_m365():
	"""
	Check if M365 sending is available

	Returns:
		bool: True if M365 sending is configured (has default outgoing account)
	"""
	account_name = frappe.db.get_value(
		"Email Account",
		{"service": "M365", "default_outgoing": 1, "enable_outgoing": 1}
	)
	return account_name is not None


def intercept_email_queue(doc, method=None):
	"""
	Hook: Email Queue before_insert
	Check if we should send via M365 instead of SMTP
	Automatically matches sender email to M365 account or uses default

	Args:
		doc: Email Queue document
		method: Hook method name (unused)
	"""
	# Get sender email from Email Queue
	sender_email = getattr(doc, 'sender', None) or frappe.session.user

	# Find the appropriate M365 account for this sender
	sending_account, is_matched = get_sending_account_for_sender(sender_email)

	if not sending_account:
		# No M365 sending configured, use default SMTP
		return

	# Mark this email for M365 sending
	doc.m365_send = 1
	doc.m365_account = sending_account.name


def send_via_m365(email_queue_doc):
	"""
	Send an Email Queue entry via Graph, one message per recipient.

	The message itself is built by frappe's own SendMailContext, exactly as it
	would be for SMTP, so attachments, inline images, unsubscribe links, tracking
	and threading headers all come from core rather than being reassembled here.

	Graph takes the recipients from the message headers, and core writes a single
	recipient into those headers (see EMail.make), so CC addresses can only
	ever be delivered by their own pass through this loop — never as a copy on
	someone else's.

	Args:
		email_queue_doc: Email Queue document

	Returns:
		bool: True if sent successfully to at least one recipient
	"""
	try:
		# Get the sending account
		sending_account = frappe.get_doc("Email Account", email_queue_doc.m365_account)

		if sending_account.service != "M365":
			frappe.log_error(
				title="M365 Email Send Failed: Invalid Account",
				message=f"Account {email_queue_doc.m365_account} is not an M365 email account"
			)
			return False

		service_principal = sending_account.m365_service_principal

		# Get access token
		access_token = get_access_token(service_principal)

		if not access_token:
			frappe.log_error(
				title="M365 Email Send Failed: No Access Token",
				message=f"Could not get access token for {sending_account.service_principal}"
			)
			return False

		sender_mailbox = sending_account.email_id

		# Core's context builds the message; only delivery differs here. Not used
		# as a context manager — M365EmailQueue.send() owns the queue status.
		ctx = SendMailContext(email_queue_doc)
		ctx.email_account_doc = sending_account

		if email_queue_doc.expose_recipients == "header":
			return send_single_mime_to_all(email_queue_doc, ctx, sender_mailbox, access_token)

		# Track if we sent to at least one recipient
		sent_to_at_least_one = False

		# Send to each recipient individually
		for recipient in email_queue_doc.recipients:
			# Skip if already sent
			if recipient.is_mail_sent():
				continue

			try:
				mime_message = build_mime_for_recipient(ctx, recipient.recipient, sender_mailbox)
				result = send_mime_as_user(sender_mailbox, mime_message, access_token)

				if result.get("success"):
					# Update recipient status
					recipient.update_db(status="Sent", commit=True)
					sent_to_at_least_one = True
					print(f"M365 Email: Sent to {recipient.recipient}")
				else:
					# Mark as error
					recipient.update_db(status="Not Sent", error=result.get("message"), commit=True)
					print(f"M365 Email: Failed to send to {recipient.recipient}")
			except Exception as e:
				# Log error for this recipient but continue with others
				recipient.update_db(status="Not Sent", error=str(e), commit=True)
				frappe.log_error(
					title=f"M365 Email: Failed to send to {recipient.recipient}",
					message=f"Email Queue: {email_queue_doc.name}\nRecipient: {recipient.recipient}\nError: {str(e)}"
				)
				print(f"M365 Email: Exception sending to {recipient.recipient}: {str(e)}")

		# Update Communication if linked and we sent to at least one recipient
		if sent_to_at_least_one and email_queue_doc.communication:
			comm = frappe.get_doc("Communication", email_queue_doc.communication)
			comm.db_set("sent_or_received", "Sent")
			comm.db_set("delivery_status", "Sent")

		return sent_to_at_least_one

	except Exception as e:
		frappe.log_error(
			title="M365 Email Send Exception",
			message=f"Email: {email_queue_doc.name}\n\nError: {str(e)}"
		)
		return False


def build_mime_for_recipient(ctx, recipient_email, sender_mailbox):
	"""
	Build frappe's per-recipient message, addressed to the sending mailbox.

	Core writes this one recipient into the To header and no CC header at all
	(show_as_cc is a display list, not an address list), so the envelope Graph
	derives from the headers is exactly this recipient.
	"""
	mime_message = ctx.build_message(recipient_email)
	return apply_sending_mailbox(mime_message, sender_mailbox)


def apply_sending_mailbox(mime_message, sender_mailbox):
	"""
	Pin the From header to the mailbox Graph is sending as, which Graph requires.
	Reply-To is left as frappe set it so replies still reach the real author.
	"""
	message_object = message_from_bytes(mime_message, policy=policy.SMTP)
	current_sender = parseaddr(message_object.get("From", ""))[1]

	if current_sender.lower() == sender_mailbox.lower():
		return mime_message

	del message_object["From"]
	message_object["From"] = sender_mailbox
	return message_object.as_bytes()


def send_single_mime_to_all(email_queue_doc, ctx, sender_mailbox, access_token):
	"""
	Send one message covering every pending recipient.

	Only for expose_recipients="header", where core puts the real To and CC lists
	into the headers: Graph would then deliver to all of them on every pass, so
	the loop has to become a single send.
	"""
	pending_recipients = [r for r in email_queue_doc.recipients if not r.is_mail_sent()]
	if not pending_recipients:
		return False

	mime_message = build_mime_for_recipient(ctx, pending_recipients[0].recipient, sender_mailbox)
	result = send_mime_as_user(sender_mailbox, mime_message, access_token)
	sent = bool(result.get("success"))

	for recipient in pending_recipients:
		if sent:
			recipient.update_db(status="Sent", commit=True)
		else:
			recipient.update_db(status="Not Sent", error=result.get("message"), commit=True)

	if sent and email_queue_doc.communication:
		comm = frappe.get_doc("Communication", email_queue_doc.communication)
		comm.db_set("sent_or_received", "Sent")
		comm.db_set("delivery_status", "Sent")

	return sent


def process_email_queue_m365():
	"""
	Process Email Queue items marked for M365 sending
	This can be called from a scheduled task or manually
	"""
	# Get all pending emails marked for M365 sending
	email_queue_list = frappe.get_all(
		"Email Queue",
		filters={
			"status": ["in", ["Not Sent", "Sending"]],
			"m365_send": 1
		},
		limit=100
	)
	
	sent_count = 0
	failed_count = 0
	
	for email_queue in email_queue_list:
		doc = frappe.get_doc("Email Queue", email_queue.name)
		
		# Update status to Sending
		doc.db_set("status", "Sending")
		frappe.db.commit()
		
		# Try to send
		if send_via_m365(doc):
			doc.db_set("status", "Sent")
			sent_count += 1
		else:
			doc.db_set("status", "Error")
			doc.db_set("error", "Failed to send via M365")
			failed_count += 1
		
		frappe.db.commit()
	
	if sent_count > 0 or failed_count > 0:
		print(f"M365 Email Queue: Sent {sent_count}, Failed {failed_count}")
	
	return {
		"sent": sent_count,
		"failed": failed_count
	}

