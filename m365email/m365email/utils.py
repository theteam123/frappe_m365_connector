# Copyright (c) 2025, Zeke Tierney and contributors
# For license information, please see license.txt

"""
Utility functions for M365 Email Integration
"""

import json
import re
import frappe
from frappe import _
from datetime import datetime
from email.utils import parseaddr
from dateutil import parser as dateutil_parser
import pytz


# NOTE: Inbound emails forwarded to a tagged address like "project+3074@domain"
# are auto-linked to the matching record. The local-part keyword (before "+")
# selects the DocType; the plus-tag (after "+") is the human-facing identifier,
# resolved to the real docname per DocType. Add a keyword here to support more.
TAGGED_ADDRESS_DOCTYPES = {"project": "Project"}

# NOTE: Matches "<keyword>+<identifier>@..." capturing the keyword and the
# identifier. Anchored to the local part so it ignores the domain.
TAGGED_ADDRESS_PATTERN = re.compile(r"^(?P<keyword>[a-z]+)\+(?P<identifier>[^@]+)@", re.IGNORECASE)


def parse_m365_datetime(datetime_string, from_timezone_name=None, to_timezone_name=None):
	"""
	Parse M365 datetime string to naive datetime for Frappe
	M365 returns ISO 8601 format with timezone (e.g., "2025-10-24T07:22:14Z")
	Frappe/MySQL expects naive datetime without timezone

	Args:
		datetime_string: ISO 8601 datetime string from M365 (usually in UTC)
		from_timezone_name: Source timezone name (e.g., "UTC", "W. Australia Standard Time")
		to_timezone_name: Target timezone name for conversion (optional)

	Returns:
		datetime: Naive datetime object (timezone removed, optionally converted)
	"""
	if not datetime_string:
		return None

	# Parse the datetime string (handles timezone)
	dt = dateutil_parser.parse(datetime_string)

	# If timezone conversion is requested
	if from_timezone_name and to_timezone_name:
		try:
			# Convert Windows timezone names to IANA format
			from_tz = get_timezone(from_timezone_name)
			to_tz = get_timezone(to_timezone_name)

			# Make datetime timezone-aware if it isn't already
			if dt.tzinfo is None:
				dt = from_tz.localize(dt)
			else:
				dt = dt.astimezone(from_tz)

			# Convert to target timezone
			dt = dt.astimezone(to_tz)
		except Exception as e:
			print(f"Timezone conversion failed: {str(e)}")

	# Convert to naive datetime (remove timezone info)
	# This keeps the time value but removes the timezone awareness
	return dt.replace(tzinfo=None)


def get_timezone(timezone_name):
	"""
	Convert timezone name to pytz timezone object
	Handles both IANA and Windows timezone names

	Args:
		timezone_name: Timezone name (IANA or Windows format)

	Returns:
		pytz timezone object
	"""
	# Map common Windows timezone names to IANA
	windows_to_iana = {
		# Australia
		"W. Australia Standard Time": "Australia/Perth",
		"AUS Eastern Standard Time": "Australia/Sydney",
		"E. Australia Standard Time": "Australia/Brisbane",
		"Cen. Australia Standard Time": "Australia/Adelaide",
		"Tasmania Standard Time": "Australia/Hobart",
		"AUS Central Standard Time": "Australia/Darwin",

		# Americas
		"Pacific Standard Time": "America/Los_Angeles",
		"Eastern Standard Time": "America/New_York",
		"Central Standard Time": "America/Chicago",
		"Mountain Standard Time": "America/Denver",

		# Brazil - Multiple timezones
		"E. South America Standard Time": "America/Sao_Paulo",  # São Paulo, Rio (UTC-3)
		"Central Brazilian Standard Time": "America/Cuiaba",    # Cuiabá (UTC-4)
		"SA Eastern Standard Time": "America/Fortaleza",        # Fortaleza (UTC-3)
		"SA Pacific Standard Time": "America/Rio_Branco",       # Acre (UTC-5)
		"Bahia Standard Time": "America/Bahia",                 # Bahia (UTC-3)

		# Europe & Africa
		"GMT Standard Time": "Europe/London",
		"Central European Standard Time": "Europe/Paris",
		"South Africa Standard Time": "Africa/Johannesburg",

		# UTC
		"UTC": "UTC",
	}

	# Try to get IANA timezone name
	iana_name = windows_to_iana.get(timezone_name, timezone_name)

	try:
		return pytz.timezone(iana_name)
	except pytz.exceptions.UnknownTimeZoneError:
		# Fallback to UTC if timezone is unknown
		print(f"Unknown timezone: {timezone_name}, falling back to UTC")
		return pytz.UTC


def parse_email_address(email_string):
	"""
	Parse email address from string (handles "Name <email@domain.com>" format)

	Args:
		email_string: Email string to parse

	Returns:
		tuple: (name, email)
	"""
	if not email_string:
		return None, None

	name, email = parseaddr(email_string)
	return name or None, email or None


def parse_recipients(recipients_list):
	"""
	Parse list of recipients from M365 message format

	Args:
		recipients_list: List of recipient dicts from Graph API

	Returns:
		str: Comma-separated string of email addresses (for Communication doctype)
	"""
	if not recipients_list:
		return ""

	emails = []
	for recipient in recipients_list:
		email_address = recipient.get("emailAddress", {}).get("address")
		if email_address:
			emails.append(email_address)

	return ", ".join(emails) if emails else ""


def should_sync_message(message, email_account):
	"""
	Determine if message should be synced based on filters

	Args:
		message: Message dict from Graph API
		email_account: Email Account doc with service='M365'

	Returns:
		bool: True if message should be synced
	"""
	# Check sync_from_date filter
	# Use m365_sync_from_date for Email Account
	sync_from_date = getattr(email_account, 'm365_sync_from_date', None)
	if sync_from_date:
		received_datetime = message.get("receivedDateTime")
		if received_datetime:
			received_date = frappe.utils.get_datetime(received_datetime).date()
			if received_date < sync_from_date:
				return False

	return True


def user_can_configure_account(user, email_account):
	"""
	Check if user can configure email account settings

	Args:
		user: Frappe user
		email_account: Email Account doc or name (with service='M365')

	Returns:
		bool: True if user can configure
	"""
	if isinstance(email_account, str):
		email_account = frappe.get_doc("Email Account", email_account)

	# System Manager can configure all
	if "System Manager" in frappe.get_roles(user):
		return True

	# For User Mailbox: user can configure their own
	account_type = getattr(email_account, 'm365_account_type', 'User Mailbox')
	if account_type == "User Mailbox" and email_account.owner == user:
		return True

	# For Shared Mailbox: only System Manager
	return False


def parse_tagged_recipient(address: str) -> tuple | None:
	"""
	Parse a tagged recipient address like "project+3074@domain" into a
	(reference_doctype, identifier) pair.

	Args:
		address: A single bare email address

	Returns:
		tuple: (reference_doctype, identifier) if the address matches a known
		keyword, otherwise None
	"""
	if not address:
		return None

	match = TAGGED_ADDRESS_PATTERN.match(address.strip())
	if not match:
		return None

	keyword = match.group("keyword").lower()
	reference_doctype = TAGGED_ADDRESS_DOCTYPES.get(keyword)
	if not reference_doctype:
		return None

	identifier = match.group("identifier").strip()
	if not identifier:
		return None

	return reference_doctype, identifier


def resolve_project_reference(project_code: str) -> str | None:
	"""
	Resolve a human-facing Project code (e.g. "3074") to the Project docname.

	NOTE: Project.autoname is UUID, so the code is held in the project_code
	field, NOT the docname. The identifier from the email must be resolved here.

	Args:
		project_code: The project_code value from the tagged address

	Returns:
		str: The Project docname (UUID) if found, otherwise None
	"""
	if not project_code:
		return None

	return frappe.db.get_value("Project", {"project_code": project_code}, "name")


def get_communication_reference(subject: str, sender_email: str, recipients: str, cc: str | None = None) -> tuple:
	"""
	Determine reference_doctype and reference_name for a Communication by
	scanning recipient addresses for a tagged alias like "project+3074@domain".

	Args:
		subject: Email subject line (unused today, kept for signature stability)
		sender_email: Sender's email address (unused today)
		recipients: Comma-separated string of To recipient addresses
		cc: Comma-separated string of CC recipient addresses (optional)

	Returns:
		tuple: (reference_doctype, reference_name) or (None, None)
	"""
	# NOTE: Linking must never break the sync run; any failure falls through to
	# (None, None) so the Communication is still created, just unlinked.
	try:
		candidate_addresses = []
		for address_string in (recipients, cc):
			if address_string:
				candidate_addresses.extend(part.strip() for part in address_string.split(","))

		for address in candidate_addresses:
			parsed = parse_tagged_recipient(address)
			if not parsed:
				continue

			reference_doctype, identifier = parsed

			# NOTE: Only Project is wired today; its code needs UUID resolution.
			# Other future doctypes can match the identifier as the docname.
			if reference_doctype == "Project":
				reference_name = resolve_project_reference(identifier)
			else:
				reference_name = identifier if frappe.db.exists(reference_doctype, identifier) else None

			if reference_name:
				return reference_doctype, reference_name

			frappe.logger("m365email").debug(
				f"Tagged address matched {reference_doctype} '{identifier}' but no record was found"
			)

		return None, None

	except Exception as error:
		frappe.logger("m365email").debug(f"get_communication_reference failed: {error}")
		return None, None


def format_email_body(body_content, content_type="html"):
	"""
	Format email body content for Communication doctype
	
	Args:
		body_content: Email body content
		content_type: Content type (html or text)
		
	Returns:
		str: Formatted body content
	"""
	if not body_content:
		return ""
	
	# For HTML content, return as-is
	if content_type.lower() == "html":
		return body_content
	
	# For text content, wrap in <pre> tags to preserve formatting
	return f"<pre>{body_content}</pre>"


def create_sync_log(email_account, sync_type="Delta Sync"):
	"""
	Create a new sync log entry

	Args:
		email_account: Email Account name or doc (with service='M365')
		sync_type: Type of sync (Full Sync, Delta Sync, Manual Sync)

	Returns:
		M365 Email Sync Log doc
	"""
	if isinstance(email_account, str):
		email_account_name = email_account
	else:
		email_account_name = email_account.name
	
	log = frappe.get_doc({
		"doctype": "M365 Email Sync Log",
		"email_account": email_account_name,
		"sync_type": sync_type,
		"status": "In Progress",
		"start_time": datetime.now()
	})
	log.insert(ignore_permissions=True)
	frappe.db.commit()
	
	return log


def update_sync_log(log, status, **kwargs):
	"""
	Update sync log with results
	
	Args:
		log: M365 Email Sync Log doc or name
		status: Status (Success, Failed, Partial Success)
		**kwargs: Additional fields to update
	"""
	if isinstance(log, str):
		log = frappe.get_doc("M365 Email Sync Log", log)
	
	log.status = status
	log.end_time = datetime.now()
	
	# Calculate duration
	if log.start_time and log.end_time:
		duration = (log.end_time - log.start_time).total_seconds()
		log.duration = duration
	
	# Update additional fields
	for key, value in kwargs.items():
		if hasattr(log, key):
			setattr(log, key, value)
	
	log.save(ignore_permissions=True)
	frappe.db.commit()


def get_or_create_contact(email_address, name=None):
	"""
	Get existing contact or create new one from email address
	
	Args:
		email_address: Email address
		name: Contact name (optional)
		
	Returns:
		str: Contact name
	"""
	if not email_address:
		return None
	
	# Check if contact exists with this email
	contact = frappe.db.get_value(
		"Contact Email",
		{"email_id": email_address},
		"parent"
	)
	
	if contact:
		return contact
	
	# Create new contact
	try:
		contact_doc = frappe.get_doc({
			"doctype": "Contact",
			"first_name": name or email_address.split("@")[0],
			"email_ids": [{
				"email_id": email_address,
				"is_primary": 1
			}]
		})
		contact_doc.insert(ignore_permissions=True)
		frappe.db.commit()
		
		return contact_doc.name
	except Exception as e:
		frappe.log_error(
			title=f"Failed to create contact for {email_address}",
			message=str(e)
		)
		return None


def sanitize_subject(subject):
	"""
	Sanitize email subject for use in Communication
	
	Args:
		subject: Email subject
		
	Returns:
		str: Sanitized subject
	"""
	if not subject:
		return _("(No Subject)")
	
	# Limit length
	max_length = 140
	if len(subject) > max_length:
		subject = subject[:max_length] + "..."
	
	return subject

