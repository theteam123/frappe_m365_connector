# Copyright (c) 2025, Zeke Tierney and contributors
# For license information, please see license.txt

"""
Email sending module for M365 Email Integration
Handles sending emails via Microsoft Graph API instead of SMTP
"""

import frappe
from frappe import _
import base64
from m365email.m365email.auth import get_access_token
from m365email.m365email.graph_api import send_email_as_user


def auto_provision_m365_account(sender_email):
	"""
	Auto-provision Email Account with service='M365' for eligible users

	Eligibility criteria:
	- User has 'M365 User' role
	- User's email domain matches a Service Principal's domain
	- Service Principal has auto-provisioning enabled

	Args:
		sender_email: Email address of the sender

	Returns:
		Email Account doc if auto-provisioned, None otherwise
	"""
	try:
		# Get the user from email
		user = frappe.db.get_value("User", {"email": sender_email}, "name")
		if not user:
			return None

		# Check if user has M365 User role
		if not frappe.db.exists("Has Role", {"parent": user, "role": "M365 User"}):
			return None

		# Extract domain from email
		if '@' not in sender_email:
			return None
		domain = sender_email.split('@')[1].lower()

		# Find a Service Principal with matching domain and auto-provisioning enabled
		service_principals = frappe.get_all(
			"M365 Email Service Principal Settings",
			filters={
				"enabled": 1,
				"enable_auto_provision": 1,
				"domain": domain
			},
			limit=1
		)

		if not service_principals:
			return None

		sp_doc = frappe.get_doc("M365 Email Service Principal Settings", service_principals[0].name)

		# Get user's full name
		user_doc = frappe.get_doc("User", user)
		user_full_name = user_doc.full_name or user_doc.first_name or user

		# Prepare signature with replaced placeholder
		signature = ""
		if sp_doc.default_footer:
			signature = sp_doc.default_footer.replace("<!--Sender-->", user_full_name)

		# Create Email Account with service='M365'
		account = frappe.get_doc({
			"doctype": "Email Account",
			"email_account_name": f"{user_full_name} ({sender_email})",
			"service": "M365",
			"email_id": sender_email,
			"linked_user": user,
			"m365_account_type": "User Mailbox",
			"m365_service_principal": sp_doc.name,
			"enable_incoming": 0,
			"enable_outgoing": 1,
			"default_outgoing": 1,
			"signature": signature
		})

		account.insert(ignore_permissions=True)
		frappe.db.commit()

		print(f"M365 Email: Auto-provisioned account for {sender_email} using Service Principal {sp_doc.name}")

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
	1. User's linked email account (via linked_user field)
	2. Email Account with service='M365' matching sender email
	3. Auto-provision if eligible
	4. Default M365 outgoing account

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

	# NOTE: First check for user's linked email account
	userLinkedAccount = GetUserDefaultEmailAccount()
	if userLinkedAccount and userLinkedAccount.service == "M365":
		return userLinkedAccount, True

	# Try to find an Email Account with service='M365' matching the sender
	if sender_email and '@' in sender_email:
		account_name = frappe.db.get_value(
			"Email Account",
			{"service": "M365", "email_id": sender_email, "enable_outgoing": 1}
		)
		if account_name:
			return frappe.get_doc("Email Account", account_name), True

		# Try auto-provisioning if no account found
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
	Send emails via M365 Graph API to each recipient individually
	This mimics Frappe's standard Email Queue behavior:
	- Sends separate email to each recipient (privacy)
	- Personalizes message with unsubscribe links, tracking, etc.
	- Updates individual recipient status

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

		# Create M365 send context helper
		ctx = M365SendContext(email_queue_doc, sending_account, access_token)

		# Track if we sent to at least one recipient
		sent_to_at_least_one = False

		# Send to each recipient individually
		for recipient in email_queue_doc.recipients:
			# Skip if already sent
			if recipient.is_mail_sent():
				continue

			try:
				# Send to this recipient
				if ctx.send_to_recipient(recipient):
					# Update recipient status
					recipient.update_db(status="Sent", commit=True)
					sent_to_at_least_one = True
					print(f"M365 Email: Sent to {recipient.recipient}")
				else:
					# Mark as error
					recipient.update_db(status="Not Sent", error="Failed to send via M365", commit=True)
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


class M365SendContext:
	"""
	Helper class for building and sending personalized M365 emails
	Similar to Frappe's SendMailContext but for M365 Graph API
	"""

	def __init__(self, queue_doc, sending_account, access_token):
		self.queue_doc = queue_doc
		self.sending_account = sending_account
		self.access_token = access_token

		# Get account email address
		self.account_email = sending_account.email_id

		# Parse the base MIME message once
		import email
		from email import policy
		self.msg = email.message_from_string(queue_doc.message, policy=policy.default)
		self.subject = self.msg.get('Subject', 'No Subject')

		# Extract base body
		self.base_body = self._extract_body()

		# Parse sender
		from email.utils import parseaddr
		sender_email = getattr(queue_doc, 'sender', None) or frappe.session.user
		if sender_email:
			_, sender_email = parseaddr(sender_email)

		# Override sender if using default account
		if sender_email != self.account_email:
			print(f"M365 Email: Overriding sender from {sender_email} to {self.account_email}")
			sender_email = self.account_email

		self.sender_email = sender_email

	def _extract_body(self):
		"""Extract HTML or text body from MIME message"""
		body = None
		if self.msg.is_multipart():
			for part in self.msg.walk():
				content_type = part.get_content_type()
				if content_type == 'text/html':
					body = part.get_content()
					break
				elif content_type == 'text/plain' and not body:
					body = part.get_content()
		else:
			body = self.msg.get_content()

		return body or "No content"

	def build_message_for_recipient(self, recipient_email):
		"""
		Build personalized message for a specific recipient
		Replaces placeholders like <!--unsubscribe_url-->, <!--email_open_check-->, etc.
		"""
		import quopri
		from frappe.utils import get_url
		from frappe.utils.verified_command import get_signed_params
		from frappe.email.queue import get_unsubcribed_url

		message = self.base_body

		# Replace unsubscribe URL placeholder
		if self.queue_doc.add_unsubscribe_link and self.queue_doc.reference_doctype:
			unsubscribe_url = get_unsubcribed_url(
				reference_doctype=self.queue_doc.reference_doctype,
				reference_name=self.queue_doc.reference_name,
				email=recipient_email,
				unsubscribe_method=self.queue_doc.get("unsubscribe_method"),
				unsubscribe_params=self.queue_doc.get("unsubscribe_param"),
			)
			unsubscribe_str = quopri.encodestring(unsubscribe_url.encode()).decode()
			message = message.replace("<!--unsubscribe_url-->", unsubscribe_str)
		else:
			# Remove placeholder if no unsubscribe link
			message = message.replace("<!--unsubscribe_url-->", "")

		# Replace email tracking pixel placeholder
		tracker_url = ""
		if self.queue_doc.communication:
			# Use communication-based email tracking
			tracker_url = f"{get_url()}/api/method/frappe.core.doctype.communication.email.mark_email_as_seen?name={self.queue_doc.communication}"

		if tracker_url:
			tracker_html = f'<img src="{tracker_url}"/>'
			tracker_str = quopri.encodestring(tracker_html.encode()).decode()
			message = message.replace("<!--email_open_check-->", tracker_str)
		else:
			message = message.replace("<!--email_open_check-->", "")

		# Replace CC message placeholder
		cc_message = ""
		if self.queue_doc.get("expose_recipients") == "footer":
			# Get TO recipients from recipients child table
			to_list = [r.recipient for r in self.queue_doc.recipients if r.recipient]
			to_str = ", ".join(to_list) if to_list else ""

			# Get CC from show_as_cc field
			cc_str = self.queue_doc.get("show_as_cc") or ""

			if to_str:
				cc_message = f"This email was sent to {to_str}"
				cc_message = f"{cc_message} and copied to {cc_str}" if cc_str else cc_message
		message = message.replace("<!--cc_message-->", cc_message)

		# Replace recipient placeholder
		recipient_str = recipient_email if self.queue_doc.get("expose_recipients") != "header" else ""
		message = message.replace("<!--recipient-->", recipient_str)

		# Append signature if configured
		footer = getattr(self.sending_account, 'signature', None)

		if footer:
			footer_html = f"""
<div class="m365-email-footer" style="margin-top: 12px;">
{footer}
</div>
"""
			message = message + footer_html

		return message

	def get_attachments(self):
		"""Get attachments from Email Queue"""
		attachments = None
		if hasattr(self.queue_doc, 'attachments') and self.queue_doc.attachments:
			try:
				import json
				# Parse attachments JSON
				attachments_data = self.queue_doc.attachments
				if isinstance(attachments_data, str):
					attachments_data = json.loads(attachments_data)

				if attachments_data:
					attachments = []
					for attachment in attachments_data:
						try:
							# Handle print format attachments (Attach Document Print)
							if attachment.get('print_format_attachment') == 1:
								attachment_copy = attachment.copy()
								attachment_copy.pop('print_format_attachment', None)

								if 'print_letterhead' in attachment_copy:
									print_letterhead = attachment_copy['print_letterhead']
									if isinstance(print_letterhead, str):
										attachment_copy['print_letterhead'] = print_letterhead == '1' or print_letterhead.lower() == 'true'

								print_format_file = frappe.attach_print(**attachment_copy)
								base64_content = base64.b64encode(print_format_file['fcontent']).decode('utf-8')

								attachments.append({
									"name": print_format_file['fname'],
									"content": base64_content
								})
								continue

							# Handle regular file attachments
							file_identifier = attachment.get('file_url') or attachment.get('fid')
							if not file_identifier:
								continue

							file_doc = None
							try:
								file_doc = frappe.get_doc("File", file_identifier)
							except frappe.DoesNotExistError:
								file_list = frappe.get_all("File", filters={"file_url": file_identifier}, limit=1)
								if file_list:
									file_doc = frappe.get_doc("File", file_list[0].name)

							if not file_doc:
								frappe.log_error(
									title="M365 Email: Attachment Not Found",
									message=f"Email: {self.queue_doc.name}\nFile identifier: {file_identifier}"
								)
								continue

							file_content = file_doc.get_content()
							base64_content = base64.b64encode(file_content).decode('utf-8')

							attachments.append({
								"name": attachment.get('file_name') or file_doc.file_name,
								"content": base64_content
							})
						except Exception as attach_error:
							frappe.log_error(
								title="M365 Email: Single Attachment Failed",
								message=f"Email: {self.queue_doc.name}\nAttachment: {attachment}\nError: {str(attach_error)}"
							)
							continue
			except Exception as e:
				frappe.log_error(
					title="M365 Email: Attachment Processing Failed",
					message=f"Email: {self.queue_doc.name}\nError: {str(e)}"
				)
				attachments = None

		return attachments

	def send_to_recipient(self, recipient):
		"""
		Send email to a single recipient

		Args:
			recipient: Email Queue Recipient object

		Returns:
			bool: True if sent successfully
		"""
		try:
			# Build personalized message for this recipient
			body = self.build_message_for_recipient(recipient.recipient)

			# Get attachments (same for all recipients)
			attachments = self.get_attachments()

			# Get CC list (if any)
			cc = None
			show_as_cc = self.queue_doc.get('show_as_cc')
			if show_as_cc and isinstance(show_as_cc, str):
				cc = [r.strip() for r in show_as_cc.split(",") if r.strip()]

			# Send via Graph API to this single recipient
			result = send_email_as_user(
				sender_email=self.sender_email,
				recipients=[recipient.recipient],  # Single recipient
				subject=self.subject,
				body=body,
				access_token=self.access_token,
				cc=cc,
				bcc=None,  # BCC not stored in Email Queue
				attachments=attachments,
				is_html=True
			)

			return result.get("success", False)

		except Exception as e:
			frappe.log_error(
				title=f"M365 Email: Failed to send to {recipient.recipient}",
				message=f"Email Queue: {self.queue_doc.name}\nError: {str(e)}"
			)
			return False


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

