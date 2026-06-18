# Copyright (c) 2025, Zeke Tierney and contributors
# For license information, please see license.txt

"""
Email synchronization module for M365 Email Integration
Core logic for syncing emails from M365 to Frappe Communications
"""

import json
import base64
import frappe
from frappe import _
from datetime import datetime
from m365email.m365email.auth import get_access_token
from m365email.m365email.graph_api import (
	get_messages_delta,
	get_message_details,
	get_message_attachments,
	download_attachment
)
from m365email.m365email.utils import (
	should_sync_message,
	parse_recipients,
	format_email_body,
	create_sync_log,
	update_sync_log,
	get_or_create_contact,
	sanitize_subject,
	get_communication_reference,
	parse_m365_datetime
)


def sync_email_account(email_account_name, folder_name=None):
	"""
	Main sync function for an email account (user or shared mailbox)
	Fetches emails from M365 and creates Communications
	Handles both initial sync and incremental delta sync

	Args:
		email_account_name: Name of Email Account with service='M365'
		folder_name: Specific folder to sync (None = sync all enabled folders)

	Returns:
		dict: Sync results
	"""
	email_account = frappe.get_doc("Email Account", email_account_name)

	if email_account.service != "M365":
		return {"success": False, "message": "Not an M365 email account"}

	return sync_email_account_by_doc(email_account, folder_name)


def sync_email_account_by_doc(email_account, folder_name=None):
	"""
	Sync function for Email Account with service='M365'.
	Used by the overridden Email Account class.

	Args:
		email_account: Email Account doc with service='M365'
		folder_name: Specific folder to sync (None = sync all enabled folders)

	Returns:
		dict: Sync results including list of communications created
	"""
	if email_account.service != "M365":
		return {"success": False, "message": "Not an M365 account"}

	if not email_account.enable_incoming:
		return {"success": False, "message": "Incoming is not enabled"}

	# Create sync log (if utils supports it for Email Account)
	try:
		sync_log = create_sync_log(email_account, sync_type="Delta Sync")
	except Exception:
		sync_log = None

	try:
		# Get access token using the service principal
		access_token = get_access_token(email_account.m365_service_principal)

		# Determine which folders to sync
		folders_to_sync = []
		if folder_name:
			folders_to_sync = [{"folder_name": folder_name}]
		elif email_account.m365_folder_filter:
			folders_to_sync = [f for f in email_account.m365_folder_filter if f.sync_enabled]
		else:
			folders_to_sync = [{"folder_name": "Inbox"}]

		total_fetched = 0
		total_created = 0
		total_updated = 0
		total_failed = 0
		communications = []

		# Sync each folder
		for folder in folders_to_sync:
			fname = folder.get("folder_name") or folder.folder_name
			result = sync_folder_for_email_account(email_account, fname, access_token, sync_log)

			total_fetched += result.get("fetched", 0)
			total_created += result.get("created", 0)
			total_updated += result.get("updated", 0)
			total_failed += result.get("failed", 0)
			communications.extend(result.get("communications", []))

		# Update email account status
		email_account.db_set("m365_last_sync_time", datetime.now(), update_modified=False)
		email_account.db_set("m365_last_sync_status", "Success", update_modified=False)
		email_account.db_set("m365_sync_error_message", None, update_modified=False)

		if sync_log:
			update_sync_log(
				sync_log,
				status="Success",
				messages_fetched=total_fetched,
				messages_created=total_created,
				messages_updated=total_updated,
				messages_failed=total_failed
			)

		frappe.db.commit()

		return {
			"success": True,
			"fetched": total_fetched,
			"created": total_created,
			"updated": total_updated,
			"failed": total_failed,
			"communications": communications
		}

	except Exception as e:
		error_msg = str(e)
		frappe.log_error(
			title=f"M365 Email Sync Failed: {email_account.name}",
			message=error_msg
		)

		email_account.db_set("m365_last_sync_status", "Failed", update_modified=False)
		email_account.db_set("m365_sync_error_message", error_msg[:500], update_modified=False)

		if sync_log:
			update_sync_log(sync_log, status="Failed", error_message=error_msg)

		frappe.db.commit()

		return {"success": False, "message": error_msg, "communications": []}


def sync_folder_for_email_account(email_account, folder_name, access_token, sync_log):
	"""
	Sync specific folder for Email Account with service='M365'.
	Wrapper around sync_folder adapted for Email Account doctype.

	Args:
		email_account: Email Account doc with service='M365'
		folder_name: Folder name to sync
		access_token: Access token
		sync_log: Sync log doc (can be None)

	Returns:
		dict: Sync results for this folder
	"""
	# Get delta token for this folder from m365_delta_tokens
	delta_tokens = {}
	if email_account.m365_delta_tokens:
		try:
			delta_tokens = json.loads(email_account.m365_delta_tokens)
		except:
			delta_tokens = {}

	delta_token = delta_tokens.get(folder_name)

	# Fetch messages using delta query
	response = get_messages_delta(
		email_account.email_id,  # Email Account uses email_id, not email_address
		access_token,
		folder=folder_name,
		delta_token=delta_token
	)

	messages = response.get("value", [])
	fetched = len(messages)
	created = 0
	updated = 0
	failed = 0
	communications = []

	# Process each message
	for message in messages:
		try:
			# Check if this is a deletion marker
			if message.get("@removed"):
				continue

			# Check sync_from_date filter
			if email_account.m365_sync_from_date:
				received_date = message.get("receivedDateTime")
				if received_date:
					msg_date = parse_m365_datetime(received_date)
					if msg_date and msg_date.date() < email_account.m365_sync_from_date:
						continue

			# Check if message already exists
			message_id = message.get("id")
			existing = frappe.db.exists(
				"Communication",
				{"m365_message_id": message_id}
			)

			if existing:
				updated += 1
				continue

			# Create communication from message
			comm = create_communication_from_message_for_email_account(
				message, email_account, access_token
			)
			if comm:
				created += 1
				communications.append(comm)
			else:
				failed += 1

		except Exception as e:
			failed += 1
			frappe.log_error(
				title=f"M365 Message Sync Failed: {message.get('id', 'Unknown')}",
				message=str(e)
			)

	# Save new delta token
	new_delta_token = response.get("@odata.deltaLink", "").split("$deltatoken=")[-1] if "@odata.deltaLink" in response else None
	if new_delta_token:
		delta_tokens[folder_name] = new_delta_token
		email_account.db_set("m365_delta_tokens", json.dumps(delta_tokens), update_modified=False)

	# Update folder filter last sync time
	if email_account.m365_folder_filter:
		for folder in email_account.m365_folder_filter:
			if folder.folder_name == folder_name:
				folder.db_set("last_sync_time", datetime.now(), update_modified=False)
				break

	return {
		"fetched": fetched,
		"created": created,
		"updated": updated,
		"failed": failed,
		"communications": communications
	}


def create_communication_from_message_for_email_account(message, email_account, access_token):
	"""
	Create Communication document from M365 message for Email Account.
	Adapted version for Email Account doctype.

	Args:
		message: M365 message data
		email_account: Email Account doc
		access_token: Access token

	Returns:
		Communication doc name or None
	"""
	try:
		# Get full message details if needed
		message_id = message.get("id")

		# Parse sender
		sender = message.get("from", {}).get("emailAddress", {})
		sender_email = sender.get("address", "")
		sender_name = sender.get("name", "")

		# Parse recipients
		to_recipients = message.get("toRecipients", [])
		cc_recipients = message.get("ccRecipients", [])

		recipients = parse_recipients(to_recipients)
		cc = parse_recipients(cc_recipients)

		# Get subject and body
		subject = sanitize_subject(message.get("subject", ""))
		body = message.get("body", {})
		content = body.get("content", "")
		content_type = body.get("contentType", "text")

		# Determine communication type
		is_incoming = email_account.email_id.lower() not in sender_email.lower()

		# Try to link to reference document
		reference_doctype, reference_name = get_communication_reference(
			subject, sender_email, recipients, cc
		)

		# Create Contact if enabled
		if email_account.m365_auto_create_contact and is_incoming:
			get_or_create_contact(sender_email, sender_name)

		# Create Communication
		comm = frappe.new_doc("Communication")
		comm.communication_type = "Communication"
		comm.communication_medium = "Email"
		comm.subject = subject
		comm.content = format_email_body(content, content_type)
		comm.sender = sender_email
		comm.sender_full_name = sender_name
		comm.recipients = recipients
		comm.cc = cc
		comm.sent_or_received = "Received" if is_incoming else "Sent"
		comm.email_account = email_account.name  # Link to Email Account
		comm.m365_message_id = message_id
		comm.m365_email_account = email_account.name  # Custom field for M365 tracking

		# Set communication date
		received_date = message.get("receivedDateTime")
		if received_date:
			comm.communication_date = parse_m365_datetime(received_date)

		# Link to reference if found
		if reference_doctype and reference_name:
			comm.reference_doctype = reference_doctype
			comm.reference_name = reference_name

		# NOTE: Set owner and linked_user for access control
		linkedUser = getattr(email_account, 'linked_user', None)
		if linkedUser:
			comm.owner = linkedUser
			comm.linked_user = linkedUser

		comm.insert(ignore_permissions=True)

		# Handle attachments if enabled
		if email_account.m365_sync_attachments and message.get("hasAttachments"):
			sync_attachments_for_communication(
				comm.name, email_account, message_id, access_token
			)

		return comm.name

	except Exception as e:
		frappe.log_error(
			title="M365 Create Communication Failed",
			message=f"Message: {message.get('id')}\nError: {str(e)}"
		)
		return None


def sync_attachments_for_communication(comm_name, email_account, message_id, access_token):
	"""
	Sync attachments for a communication.

	Args:
		comm_name: Communication name
		email_account: Email Account doc
		message_id: M365 message ID
		access_token: Access token
	"""
	try:
		attachments = get_message_attachments(
			email_account.email_id,
			message_id,
			access_token
		)

		max_size = (email_account.m365_max_attachment_size or 10) * 1024 * 1024  # MB to bytes

		for attachment in attachments.get("value", []):
			try:
				# Skip if too large
				if attachment.get("size", 0) > max_size:
					continue

				# Download attachment content
				# Graph API returns attachment object with contentBytes field
				attachment_data = download_attachment(
					email_account.email_id,
					message_id,
					attachment.get("id"),
					access_token
				)

				# Extract contentBytes from the response
				content_bytes = attachment_data.get("contentBytes") if isinstance(attachment_data, dict) else None

				if content_bytes:
					# Create Frappe file - contentBytes is base64 encoded
					file_doc = frappe.get_doc({
						"doctype": "File",
						"file_name": attachment.get("name"),
						"attached_to_doctype": "Communication",
						"attached_to_name": comm_name,
						"content": base64.b64decode(content_bytes)
					})
					file_doc.insert(ignore_permissions=True)

			except Exception as e:
				frappe.log_error(
					title="M365 Attachment Sync Failed",
					message=f"Attachment: {attachment.get('name')}\nError: {str(e)}"
				)
	except Exception as e:
		frappe.log_error(
			title="M365 Attachments Fetch Failed",
			message=f"Message: {message_id}\nError: {str(e)}"
		)


