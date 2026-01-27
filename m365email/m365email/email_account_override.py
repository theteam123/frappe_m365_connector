# Copyright (c) 2025, Zeke Tierney and contributors
# For license information, please see license.txt

"""
Email Account doctype override for M365 integration.

This module extends the standard Email Account class to support M365
service type, handling:
- M365-specific validation (skip SMTP/IMAP validation)
- Email receiving via Graph API delta sync
- Access token management through service principal
- Override of find_outgoing to include M365 accounts
"""

import frappe
from frappe import _
from frappe.email.doctype.email_account.email_account import EmailAccount
from frappe.utils import now_datetime


class M365EmailAccount(EmailAccount):
	"""
	Extended Email Account class with M365 support.
	
	When service='M365', this class:
	- Skips SMTP/IMAP validation
	- Uses Graph API for receiving emails
	- Gets access tokens from linked service principal
	"""
	
	def validate(self):
		"""
		Validate Email Account settings.
		For M365 accounts, skip SMTP/IMAP validation.
		"""
		if self.service == "M365":
			self._validate_m365()
		else:
			# NOTE: Use standard validation for non-M365 accounts
			super().validate()

		# NOTE: Validate linked user for all account types
		self._validateLinkedUser()

	def _validate_m365(self):
		"""Validate M365-specific fields."""
		from frappe.utils import validate_email_address
		
		# Validate email address
		if self.email_id:
			validate_email_address(self.email_id, True)
		
		# Validate service principal is set
		if not self.m365_service_principal:
			frappe.throw(_("Service Principal is required for M365 accounts"))
		
		# Validate service principal exists and is enabled
		try:
			sp = frappe.get_doc("M365 Email Service Principal Settings", self.m365_service_principal)
			if not sp.enabled:
				frappe.throw(
					_("Service Principal {0} is not enabled").format(self.m365_service_principal)
				)
		except frappe.DoesNotExistError:
			frappe.throw(
				_("Service Principal {0} not found").format(self.m365_service_principal)
			)
		
		# Check for duplicate email per service principal
		duplicate = frappe.db.exists(
			"Email Account",
			{
				"name": ["!=", self.name],
				"service": "M365",
				"m365_service_principal": self.m365_service_principal,
				"email_id": self.email_id
			}
		)
		if duplicate:
			frappe.throw(
				_("An M365 Email Account with email {0} already exists for this service principal").format(
					self.email_id
				)
			)
		
		# Validate default_outgoing requires enable_outgoing
		if self.default_outgoing and not self.enable_outgoing:
			frappe.throw(_("Default Outgoing requires Enable Outgoing to be checked"))
		
		# Validate notification settings
		if self.notify_if_unreplied:
			if not self.send_notification_to:
				frappe.throw(_("{0} is mandatory").format(self.meta.get_label("send_notification_to")))


	def _validateLinkedUser(self):
		"""
		Validate linked user settings.
		Ensures only one email account can be the user default for a given linked user.
		"""
		linkedUser = getattr(self, 'linked_user', None)

		if not linkedUser:
			return

		# NOTE: Validate that the linked user exists
		if not frappe.db.exists("User", linkedUser):
			frappe.throw(_("Linked User '{0}' does not exist").format(linkedUser))

		# NOTE: If user_default is checked, ensure no other account is user default for this user
		isUserDefault = getattr(self, 'user_default', 0)
		if not isUserDefault:
			return

		existingDefault = frappe.db.get_value(
			"Email Account",
			{
				"name": ["!=", self.name],
				"linked_user": linkedUser,
				"user_default": 1,
				"enable_outgoing": 1
			},
			"name"
		)

		if existingDefault:
			frappe.throw(
				_("Email Account '{0}' is already set as the User Default for '{1}'. "
				  "Please uncheck User Default on that account first.").format(
					existingDefault, linkedUser
				)
			)


	def get_incoming_server(self, in_receive=False, email_sync_rule="UNSEEN"):
		"""
		Get incoming server connection.
		For M365 accounts, return None as we use Graph API instead.
		"""
		if self.service == "M365":
			return None
		return super().get_incoming_server(in_receive=in_receive, email_sync_rule=email_sync_rule)
	
	def validate_smtp_conn(self):
		"""
		Validate SMTP connection.
		For M365 accounts, skip SMTP validation.
		"""
		if self.service == "M365":
			return True
		return super().validate_smtp_conn()
	
	def receive(self):
		"""
		Receive emails from this Email account.
		For M365 accounts, use Graph API delta sync.
		For other accounts, use standard POP3/IMAP.
		"""
		if self.service == "M365":
			return self._receive_m365()
		return super().receive()
	
	def _receive_m365(self):
		"""Receive emails using M365 Graph API delta sync."""
		from m365email.m365email.sync import sync_email_account_by_doc
		
		if not self.enable_incoming:
			return []
		
		try:
			result = sync_email_account_by_doc(self)
			return result.get("communications", [])
		except Exception as e:
			# Update sync status
			self.db_set("m365_last_sync_time", now_datetime(), update_modified=False)
			self.db_set("m365_last_sync_status", "Failed", update_modified=False)
			self.db_set("m365_sync_error_message", str(e)[:500], update_modified=False)
			
			frappe.log_error(
				title=f"M365 Email Sync Failed: {self.name}",
				message=str(e)
			)
			return []
	
	def get_m365_access_token(self, force_refresh=False):
		"""
		Get access token for this M365 Email Account.

		Args:
			force_refresh: Force token refresh even if cached token is valid

		Returns:
			str: Access token
		"""
		if self.service != "M365":
			frappe.throw(_("Access token is only available for M365 accounts"))

		if not self.m365_service_principal:
			frappe.throw(_("Service Principal is not configured"))

		from m365email.m365email.auth import get_access_token
		return get_access_token(self.m365_service_principal, force_refresh=force_refresh)

	@classmethod
	def find_outgoing(cls, match_by_email=None, match_by_doctype=None, _raise_error=False):
		"""
		Find the outgoing Email account to use.
		Extended to properly handle M365 accounts.

		:param match_by_email: Find account using emailID
		:param match_by_doctype: Find account by matching `Append To` doctype
		:param _raise_error: Raise error if no account found
		"""
		from frappe.utils import parse_addr

		if match_by_email:
			match_by_email = parse_addr(match_by_email)[1]
			# First check for M365 account with this email
			doc = cls.find_one_by_filters(enable_outgoing=1, service="M365", email_id=match_by_email)
			if doc:
				return {match_by_email: doc}
			# Then check standard accounts
			doc = cls.find_one_by_filters(enable_outgoing=1, email_id=match_by_email)
			if doc:
				return {match_by_email: doc}

		if match_by_doctype:
			doc = cls.find_one_by_filters(enable_outgoing=1, enable_incoming=1, append_to=match_by_doctype)
			if doc:
				return {match_by_doctype: doc}

		# Find default - check M365 defaults first, then standard
		doc = cls.find_one_by_filters(enable_outgoing=1, default_outgoing=1, service="M365")
		if doc:
			return {"default": doc}

		doc = cls.find_default_outgoing()
		if doc:
			return {"default": doc}

		if _raise_error:
			frappe.throw(
				_("Please setup a default Email Account from Settings > Email Account"),
				frappe.OutgoingEmailError,
			)

	@classmethod
	def find_m365_account_for_user(cls, user: str = None):
		"""
		Find an M365 Email Account linked to the given user.

		Args:
			user: User ID. Defaults to current session user.

		Returns:
			Email Account doc or None
		"""
		if not user:
			user = frappe.session.user

		if not user or user == "Guest":
			return None

		# NOTE: Find email account linked to this user (prefer default_outgoing)
		account = frappe.db.get_value(
			"Email Account",
			{
				"service": "M365",
				"linked_user": user,
				"default_outgoing": 1,
				"enable_outgoing": 1
			},
			"name"
		)

		if account:
			return frappe.get_doc("Email Account", account)

		# NOTE: Fall back to any M365 account linked to user
		account = frappe.db.get_value(
			"Email Account",
			{
				"service": "M365",
				"linked_user": user,
				"enable_outgoing": 1
			},
			"name"
		)

		if account:
			return frappe.get_doc("Email Account", account)

		return None

