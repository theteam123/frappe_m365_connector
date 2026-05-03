# Copyright (c) 2025, Zeke Tierney and contributors
# For license information, please see license.txt

"""
Custom fields for M365 Email Integration

This module adds M365-specific fields to the Email Account doctype,
allowing M365 to be used as a service type alongside SMTP/IMAP.
"""

import frappe
from frappe.custom.doctype.custom_field.custom_field import create_custom_fields
from frappe.custom.doctype.property_setter.property_setter import make_property_setter


# Timezone options for M365 accounts
M365_TIMEZONE_OPTIONS = """Australia/Perth
Australia/Sydney
Australia/Brisbane
Australia/Adelaide
Australia/Hobart
Australia/Darwin
America/Los_Angeles
America/New_York
America/Chicago
America/Denver
America/Sao_Paulo
America/Cuiaba
America/Fortaleza
America/Rio_Branco
America/Bahia
Europe/London
Europe/Paris
Africa/Johannesburg
UTC"""


def create_m365_custom_fields():
	"""
	Create custom fields for M365 Email Integration

	This adds M365-specific fields to Email Account and updates
	Communication/Email Queue/Event to link to Email Account.
	"""

	# First, add M365 as a service option to Email Account
	add_m365_service_option()

	custom_fields = {
		# ============================================
		# Email Account - M365 Settings Section
		# ============================================
		"Email Account": [
			# ============================================
			# Linked User (Details Tab - Account Section)
			# ============================================
			{
				"fieldname": "linked_user",
				"label": "Linked User",
				"fieldtype": "Link",
				"options": "User",
				"insert_after": "email_id",
				"description": "The Frappe user who owns this email account. Only this user can view emails synced from this account."
			},

			# ============================================
			# User Default (Outgoing Tab)
			# ============================================
			{
				"fieldname": "user_default",
				"label": "User Default",
				"fieldtype": "Check",
				"insert_after": "track_email_status",
				"depends_on": "enable_outgoing",
				"default": "0",
				"description": "When checked, this email account becomes the default for sending emails when logged in as the Linked User"
			},

			# M365 Settings Section Break
			{
				"fieldname": "m365_settings_section",
				"label": "M365 Settings",
				"fieldtype": "Section Break",
				"insert_after": "service",
				"depends_on": "eval:doc.service=='M365'",
				"collapsible": 0
			},
			# Service Principal Link
			{
				"fieldname": "m365_service_principal",
				"label": "Service Principal",
				"fieldtype": "Link",
				"options": "M365 Email Service Principal Settings",
				"insert_after": "m365_settings_section",
				"depends_on": "eval:doc.service=='M365'",
				"mandatory_depends_on": "eval:doc.service=='M365'",
				"description": "Azure AD Service Principal for authentication"
			},
			# Account Type (User Mailbox vs Shared Mailbox)
			{
				"fieldname": "m365_account_type",
				"label": "M365 Account Type",
				"fieldtype": "Select",
				"options": "User Mailbox\nShared Mailbox",
				"default": "User Mailbox",
				"insert_after": "m365_service_principal",
				"depends_on": "eval:doc.service=='M365'"
			},
			# Column break for M365 settings
			{
				"fieldname": "m365_column_break_1",
				"fieldtype": "Column Break",
				"insert_after": "m365_account_type"
			},
			# Sync calendar events
			{
				"fieldname": "m365_sync_events",
				"label": "Sync Calendar Events",
				"fieldtype": "Check",
				"default": "0",
				"insert_after": "m365_column_break_1",
				"depends_on": "eval:doc.service=='M365'",
				"description": "Sync calendar events from M365 to Frappe Events"
			},

			# ============================================
			# M365 Sync Settings Section
			# ============================================
			{
				"fieldname": "m365_sync_settings_section",
				"label": "M365 Sync Settings",
				"fieldtype": "Section Break",
				"insert_after": "m365_sync_events",
				"depends_on": "eval:doc.service=='M365' && doc.enable_incoming",
				"collapsible": 1
			},
			# Sync from date
			{
				"fieldname": "m365_sync_from_date",
				"label": "Sync From Date",
				"fieldtype": "Date",
				"insert_after": "m365_sync_settings_section",
				"depends_on": "eval:doc.service=='M365'",
				"description": "Only sync emails received after this date"
			},
			# User timezone
			{
				"fieldname": "m365_user_timezone",
				"label": "User Timezone",
				"fieldtype": "Select",
				"options": M365_TIMEZONE_OPTIONS,
				"default": "Australia/Perth",
				"insert_after": "m365_sync_from_date",
				"depends_on": "eval:doc.service=='M365'",
				"description": "Timezone for calendar events"
			},
			# Auto create contact
			{
				"fieldname": "m365_auto_create_contact",
				"label": "Auto Create Contact",
				"fieldtype": "Check",
				"default": "1",
				"insert_after": "m365_user_timezone",
				"depends_on": "eval:doc.service=='M365'"
			},
			# Column break
			{
				"fieldname": "m365_sync_column_break",
				"fieldtype": "Column Break",
				"insert_after": "m365_auto_create_contact"
			},
			# Sync attachments
			{
				"fieldname": "m365_sync_attachments",
				"label": "Sync Attachments",
				"fieldtype": "Check",
				"default": "1",
				"insert_after": "m365_sync_column_break",
				"depends_on": "eval:doc.service=='M365'"
			},
			# Max attachment size
			{
				"fieldname": "m365_max_attachment_size",
				"label": "Max Attachment Size (MB)",
				"fieldtype": "Int",
				"default": "10",
				"insert_after": "m365_sync_attachments",
				"depends_on": "eval:doc.service=='M365'"
			},

			# ============================================
			# M365 Folder Sync Section
			# ============================================
			{
				"fieldname": "m365_folder_section",
				"label": "M365 Folder Sync",
				"fieldtype": "Section Break",
				"insert_after": "m365_max_attachment_size",
				"depends_on": "eval:doc.service=='M365' && doc.enable_incoming",
				"collapsible": 1
			},
			# Folder filter table
			{
				"fieldname": "m365_folder_filter",
				"label": "Folders to Sync",
				"fieldtype": "Table",
				"options": "M365 Email Folder Filter",
				"insert_after": "m365_folder_section",
				"depends_on": "eval:doc.service=='M365'"
			},

			# ============================================
			# M365 Sync Status Section
			# ============================================
			{
				"fieldname": "m365_status_section",
				"label": "M365 Sync Status",
				"fieldtype": "Section Break",
				"insert_after": "m365_folder_filter",
				"depends_on": "eval:doc.service=='M365'",
				"collapsible": 1
			},
			# Last sync time
			{
				"fieldname": "m365_last_sync_time",
				"label": "Last Sync Time",
				"fieldtype": "Datetime",
				"insert_after": "m365_status_section",
				"read_only": 1,
				"no_copy": 1
			},
			# Last sync status
			{
				"fieldname": "m365_last_sync_status",
				"label": "Last Sync Status",
				"fieldtype": "Select",
				"options": "\nSuccess\nFailed\nIn Progress",
				"insert_after": "m365_last_sync_time",
				"read_only": 1,
				"no_copy": 1
			},
			# Column break
			{
				"fieldname": "m365_status_column_break",
				"fieldtype": "Column Break",
				"insert_after": "m365_last_sync_status"
			},
			# Sync error message
			{
				"fieldname": "m365_sync_error_message",
				"label": "Sync Error Message",
				"fieldtype": "Small Text",
				"insert_after": "m365_status_column_break",
				"read_only": 1,
				"no_copy": 1
			},
			# Delta tokens (hidden, for internal use)
			{
				"fieldname": "m365_delta_tokens",
				"label": "Delta Tokens",
				"fieldtype": "Long Text",
				"insert_after": "m365_sync_error_message",
				"read_only": 1,
				"hidden": 1,
				"no_copy": 1
			}
		],

		# ============================================
		# Communication - M365 fields
		# ============================================
		"Communication": [
			{
				"fieldname": "m365_message_id",
				"label": "M365 Message ID",
				"fieldtype": "Data",
				"length": 500,
				"insert_after": "uid",
				"read_only": 1,
				"no_copy": 1,
				"hidden": 1,
				"unique": 1
			},
			{
				"fieldname": "m365_email_account",
				"label": "M365 Email Account",
				"fieldtype": "Link",
				"options": "Email Account",
				"insert_after": "m365_message_id",
				"read_only": 1,
				"no_copy": 1,
				"hidden": 1
			},
			{
				"fieldname": "linked_user",
				"label": "Linked User",
				"fieldtype": "Link",
				"options": "User",
				"insert_after": "m365_email_account",
				"read_only": 1,
				"no_copy": 1,
				"hidden": 1,
				"description": "User who owns this email for privacy filtering"
			}
		],

		# ============================================
		# Email Queue - M365 fields
		# ============================================
		"Email Queue": [
			{
				"fieldname": "m365_send",
				"label": "Send via M365",
				"fieldtype": "Check",
				"insert_after": "status",
				"read_only": 1,
				"no_copy": 1,
				"default": "0"
			},
			{
				"fieldname": "m365_account",
				"label": "M365 Account",
				"fieldtype": "Link",
				"options": "Email Account",  # Changed from M365 Email Account
				"insert_after": "m365_send",
				"read_only": 1,
				"no_copy": 1
			}
		],

		# ============================================
		# Event - M365 calendar sync fields
		# ============================================
		"Event": [
			{
				"fieldname": "m365_event_id",
				"label": "M365 Event ID",
				"fieldtype": "Data",
				"length": 500,
				"insert_after": "event_type",
				"read_only": 1,
				"no_copy": 1,
				"hidden": 1,
				"unique": 1
			},
			{
				"fieldname": "m365_email_account",
				"label": "M365 Email Account",
				"fieldtype": "Link",
				"options": "Email Account",  # Changed from M365 Email Account
				"insert_after": "m365_event_id",
				"read_only": 1,
				"no_copy": 1,
				"hidden": 1
			},
			{
				"fieldname": "m365_icaluid",
				"label": "M365 iCalUId",
				"fieldtype": "Data",
				"length": 500,
				"insert_after": "m365_email_account",
				"read_only": 1,
				"no_copy": 1,
				"hidden": 1
			},
			{
				"fieldname": "m365_timezone",
				"label": "M365 Timezone",
				"fieldtype": "Data",
				"length": 100,
				"insert_after": "m365_icaluid",
				"read_only": 1,
				"no_copy": 1,
				"hidden": 0,
				"description": "Original timezone from Microsoft 365"
			}
		]
	}

	create_custom_fields(custom_fields, update=True)
	frappe.db.commit()


def add_m365_service_option():
	"""
	Add 'M365' as a service option to Email Account's service field
	"""
	# Get current options for the service field
	current_options = frappe.db.get_value(
		"DocField",
		{"parent": "Email Account", "fieldname": "service"},
		"options"
	)

	if current_options and "M365" not in current_options:
		# Add M365 to the options
		new_options = current_options + "\nM365"

		# Use property setter to add the option
		make_property_setter(
			"Email Account",
			"service",
			"options",
			new_options,
			"Text",
			for_doctype=False,
			validate_fields_for_doctype=False
		)

		frappe.db.commit()
		print("✅ Added 'M365' to Email Account service options")
	elif current_options and "M365" in current_options:
		print("ℹ️ 'M365' already exists in Email Account service options")
	else:
		# If no options exist, create with M365
		make_property_setter(
			"Email Account",
			"service",
			"options",
			"M365",
			"Text",
			for_doctype=False,
			validate_fields_for_doctype=False
		)
		frappe.db.commit()
		print("✅ Created Email Account service options with 'M365'")


def execute():
	"""
	Execute function for patches
	"""
	create_m365_custom_fields()

