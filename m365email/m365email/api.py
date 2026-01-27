# Copyright (c) 2025, Zeke Tierney and contributors
# For license information, please see license.txt

"""
API endpoints for M365 Email Integration
Whitelisted functions for frontend/client access
"""

import frappe
from frappe import _
from m365email.m365email.auth import test_connection, get_access_token
from m365email.m365email.sync import sync_email_account
from m365email.m365email.event_sync import sync_calendar_events
from m365email.m365email.graph_api import get_mail_folders
from m365email.m365email.utils import user_can_configure_account


@frappe.whitelist()
def enable_email_sync(email_address, service_principal, account_type="User Mailbox"):
	"""
	Enable email sync for current user's mailbox

	Creates an Email Account with service='M365' (new approach).

	Args:
		email_address: User's email address
		service_principal: Service principal name
		account_type: "User Mailbox" or "Shared Mailbox"

	Returns:
		dict: Created email account details
	"""
	user = frappe.session.user

	# For Shared Mailbox, only System Manager can create
	if account_type == "Shared Mailbox":
		if "System Manager" not in frappe.get_roles(user):
			frappe.throw(_("Only System Manager can configure shared mailboxes"))

	# Check if account already exists in Email Account
	existing_email_account = frappe.db.exists(
		"Email Account",
		{
			"email_id": email_address,
			"service": "M365"
		}
	)

	if existing_email_account:
		frappe.throw(_("Email account already exists for this email address"))

	# Create Email Account with service='M365'
	account = frappe.get_doc({
		"doctype": "Email Account",
		"email_account_name": f"M365-{email_address}",
		"email_id": email_address,
		"service": "M365",
		"enable_incoming": 1,
		"enable_outgoing": 1,
		"default_incoming": 0,
		"default_outgoing": 0,
		# M365-specific fields
		"m365_service_principal": service_principal,
		"m365_account_type": account_type,
		"m365_sync_attachments": 1,
		"m365_max_attachment_size": 10,
		# Folder filters as child table
		"m365_folder_filter": [
			{"folder_name": "Inbox", "sync_enabled": 1},
			{"folder_name": "Sent Items", "sync_enabled": 0}
		]
	})
	account.insert()
	frappe.db.commit()

	return {
		"success": True,
		"account_name": account.name,
		"doctype": "Email Account",
		"message": _("Email sync enabled successfully")
	}


@frappe.whitelist()
def disable_email_sync(email_account_name):
	"""
	Disable email sync for an account

	Args:
		email_account_name: Name of Email Account with service='M365'

	Returns:
		dict: Success message
	"""
	account = frappe.get_doc("Email Account", email_account_name)

	if account.service != "M365":
		frappe.throw(_("Not an M365 email account"))

	# Check permissions
	if not frappe.has_permission("Email Account", "write", email_account_name):
		frappe.throw(_("You don't have permission to configure this email account"))

	account.enable_incoming = 0
	account.save()
	frappe.db.commit()

	return {
		"success": True,
		"message": _("Email sync disabled successfully")
	}


@frappe.whitelist()
def trigger_manual_sync(email_account_name):
	"""
	Manually trigger sync for an email account

	Args:
		email_account_name: Name of Email Account with service='M365'

	Returns:
		dict: Sync results
	"""
	account = frappe.get_doc("Email Account", email_account_name)

	if account.service != "M365":
		frappe.throw(_("Not an M365 email account"))

	# Check permissions
	if not frappe.has_permission("Email Account", "read", email_account_name):
		frappe.throw(_("You don't have permission to sync this email account"))

	return sync_email_account(email_account_name)


@frappe.whitelist()
def trigger_manual_event_sync(email_account_name):
	"""
	Manually trigger calendar event sync for an email account

	Args:
		email_account_name: Name of Email Account with service='M365'

	Returns:
		dict: Sync results
	"""
	account = frappe.get_doc("Email Account", email_account_name)

	if account.service != "M365":
		frappe.throw(_("Not an M365 email account"))

	# Check permissions
	if not frappe.has_permission("Email Account", "read", email_account_name):
		frappe.throw(_("You don't have permission to sync this email account"))

	return sync_calendar_events(email_account_name)


@frappe.whitelist()
def get_sync_status(email_account_name=None):
	"""
	Get sync status and recent logs

	Args:
		email_account_name: Name of Email Account with service='M365' (optional)

	Returns:
		dict: Sync status information
	"""
	user = frappe.session.user

	if email_account_name:
		account = frappe.get_doc("Email Account", email_account_name)

		if account.service != "M365":
			frappe.throw(_("Not an M365 email account"))

		# Check permissions
		if not frappe.has_permission("Email Account", "read", email_account_name):
			frappe.throw(_("You don't have permission to view this email account"))

		# Get recent logs
		logs = frappe.get_all(
			"M365 Email Sync Log",
			filters={"email_account": email_account_name},
			fields=["name", "sync_type", "status", "start_time", "end_time", "messages_fetched", "messages_created"],
			order_by="start_time desc",
			limit=10
		)

		return {
			"account": account.as_dict(),
			"logs": logs
		}
	else:
		# Get all M365 accounts
		all_accounts = []

		# Get Email Accounts with service='M365'
		if "System Manager" in frappe.get_roles(user):
			m365_email_accounts = frappe.get_all(
				"Email Account",
				filters={"service": "M365"},
				fields=["name", "email_account_name", "email_id", "enable_incoming", "m365_last_sync_time", "m365_last_sync_status"]
			)
		else:
			# For non-admin users, they see their linked email accounts
			m365_email_accounts = frappe.get_all(
				"Email Account",
				filters={"service": "M365"},
				fields=["name", "email_account_name", "email_id", "enable_incoming", "m365_last_sync_time", "m365_last_sync_status"]
			)

		for acc in m365_email_accounts:
			all_accounts.append({
				"name": acc.name,
				"account_name": acc.email_account_name,
				"email_address": acc.email_id,
				"enabled": acc.enable_incoming,
				"last_sync_time": acc.m365_last_sync_time,
				"last_sync_status": acc.m365_last_sync_status,
				"doctype": "Email Account"
			})

		return {
			"accounts": all_accounts
		}


@frappe.whitelist()
def test_service_principal_connection(service_principal_name):
	"""
	Test service principal credentials
	System Manager only
	
	Args:
		service_principal_name: Name of service principal
		
	Returns:
		dict: Connection test results
	"""
	if "System Manager" not in frappe.get_roles(frappe.session.user):
		frappe.throw(_("Only System Manager can test service principal connections"))
	
	result = test_connection(service_principal_name)
	return result


@frappe.whitelist()
def get_available_service_principals():
	"""
	Get list of enabled service principals
	For user to select when enabling email sync
	
	Returns:
		list: Available service principals
	"""
	service_principals = frappe.get_all(
		"M365 Email Service Principal Settings",
		filters={"enabled": 1},
		fields=["name", "service_principal_name", "tenant_name", "tenant_id"]
	)
	
	return service_principals


@frappe.whitelist()
def get_shared_mailboxes():
	"""
	Get all configured shared mailboxes
	System Manager sees all, others see none

	Returns:
		list: Shared mailboxes (Email Accounts with service='M365' and m365_account_type='Shared Mailbox')
	"""
	# Only System Manager can see shared mailbox configurations
	if "System Manager" not in frappe.get_roles(frappe.session.user):
		return []

	# Get Email Accounts with service='M365' and m365_account_type='Shared Mailbox'
	m365_email_accounts = frappe.get_all(
		"Email Account",
		filters={"service": "M365", "m365_account_type": "Shared Mailbox"},
		fields=["name", "email_account_name", "email_id", "enable_incoming", "m365_last_sync_time", "m365_last_sync_status"]
	)

	shared_mailboxes = []
	for acc in m365_email_accounts:
		shared_mailboxes.append({
			"name": acc.name,
			"account_name": acc.email_account_name,
			"email_address": acc.email_id,
			"enabled": acc.enable_incoming,
			"last_sync_time": acc.m365_last_sync_time,
			"last_sync_status": acc.m365_last_sync_status,
			"doctype": "Email Account"
		})

	return shared_mailboxes


@frappe.whitelist()
def get_available_folders(email_account_name):
	"""
	Get available mail folders for an email account

	Args:
		email_account_name: Name of Email Account with service='M365'

	Returns:
		list: Available folders
	"""
	account = frappe.get_doc("Email Account", email_account_name)

	if account.service != "M365":
		frappe.throw(_("Not an M365 email account"))

	# Check permissions
	if not frappe.has_permission("Email Account", "read", email_account_name):
		frappe.throw(_("You don't have permission to view this email account"))

	email_address = account.email_id
	service_principal = account.m365_service_principal

	# Get access token
	access_token = get_access_token(service_principal)

	# Get folders from Graph API
	response = get_mail_folders(email_address, access_token)
	folders = response.get("value", [])

	# Format folder list
	folder_list = []
	for folder in folders:
		folder_list.append({
			"id": folder.get("id"),
			"displayName": folder.get("displayName"),
			"totalItemCount": folder.get("totalItemCount"),
			"unreadItemCount": folder.get("unreadItemCount")
		})

	return folder_list


@frappe.whitelist()
def update_folder_filters(email_account_name, folders):
	"""
	Update folder filters for an email account

	Args:
		email_account_name: Name of Email Account with service='M365'
		folders: List of folder dicts with folder_name and sync_enabled

	Returns:
		dict: Success message
	"""
	import json

	# Parse folders if string
	if isinstance(folders, str):
		folders = json.loads(folders)

	account = frappe.get_doc("Email Account", email_account_name)

	if account.service != "M365":
		frappe.throw(_("Not an M365 email account"))

	# Check permissions
	if not frappe.has_permission("Email Account", "write", email_account_name):
		frappe.throw(_("You don't have permission to configure this email account"))

	# Clear existing folder filters
	account.m365_folder_filter = []

	# Add new filters using child table
	for folder in folders:
		account.append("m365_folder_filter", {
			"folder_name": folder.get("folder_name"),
			"sync_enabled": folder.get("sync_enabled", 1)
		})

	account.save()
	frappe.db.commit()

	return {
		"success": True,
		"message": _("Folder filters updated successfully")
	}




@frappe.whitelist()
def SendEmailQueueNow(queues: str | list) -> dict:
	"""
	Immediately process and send selected emails from the Email Queue.

	Unlike retry_sending which just marks emails for later processing,
	this function immediately sends the emails.

	Args:
		queues: List of Email Queue names or JSON string of names

	Returns:
		dict: Results with sent count, failed count, and any errors
	"""
	import json

	if not frappe.has_permission("Email Queue"):
		frappe.throw(_("You don't have permission to access Email Queue"))

	if isinstance(queues, str):
		queues = json.loads(queues)

	if not queues:
		return {"sent": 0, "failed": 0, "errors": []}

	sentCount = 0
	failedCount = 0
	errors = []

	for queueName in queues:
		try:
			queueDoc = frappe.get_doc("Email Queue", queueName)

			if queueDoc.status not in ["Not Sent", "Partially Sent", "Error"]:
				continue

			queueDoc.send(force_send=True)
			queueDoc.reload()

			if queueDoc.status == "Sent":
				sentCount += 1
			elif queueDoc.status in ["Error", "Partially Sent"]:
				failedCount += 1
				errors.append({
					"name": queueName,
					"error": queueDoc.error or "Unknown error"
				})

		except Exception as e:
			failedCount += 1
			errors.append({
				"name": queueName,
				"error": str(e)
			})

	frappe.db.commit()

	return {
		"sent": sentCount,
		"failed": failedCount,
		"errors": errors
	}

