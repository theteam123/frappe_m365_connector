# Copyright (c) 2025, Zeke Tierney and contributors
# For license information, please see license.txt

"""
Microsoft Graph API integration module
Handles all Graph API requests for M365 Email Integration
"""

import requests
import frappe
from frappe import _
import time


# NOTE: Map display folder names to Microsoft Graph API well-known folder names
# Graph API expects lowercase, specific naming (e.g., 'sentitems' not 'Sent Items')
FOLDER_NAME_MAP = {
    "inbox": "inbox",
    "sent": "sentitems",
    "sent items": "sentitems",
    "sentitems": "sentitems",
    "drafts": "drafts",
    "draft": "drafts",
    "deleted": "deleteditems",
    "deleted items": "deleteditems",
    "deleteditems": "deleteditems",
    "trash": "deleteditems",
    "junk": "junkemail",
    "junk email": "junkemail",
    "junkemail": "junkemail",
    "spam": "junkemail",
    "archive": "archive",
    "outbox": "outbox",
    "scheduled": "scheduled",
}


def NormalizeFolderName(folderName: str) -> str:
    """
    Convert display folder name to Microsoft Graph API well-known folder name.

    Microsoft Graph API expects specific well-known folder names:
    - inbox, sentitems, drafts, deleteditems, junkemail, archive, outbox, etc.

    Args:
        folderName: Display name like "Inbox", "Sent Items", "Draft", etc.

    Returns:
        str: Graph API well-known folder name or original if not mapped
    """
    if not folderName:
        return "inbox"

    normalizedName = folderName.lower().strip()
    return FOLDER_NAME_MAP.get(normalizedName, folderName)


def make_graph_request(endpoint, access_token, method='GET', data=None, params=None):
	"""
	Generic Graph API request handler
	Handles pagination, rate limiting, error handling

	Args:
		endpoint: Graph API endpoint (full URL or path)
		access_token: Access token for authentication
		method: HTTP method (GET, POST, PATCH, etc.)
		data: Request body data (for POST/PATCH)
		params: Query parameters

	Returns:
		dict: Response data
	"""
	# Ensure endpoint is a full URL
	if not endpoint.startswith("http"):
		endpoint = f"https://graph.microsoft.com/v1.0{endpoint}"

	headers = {
		"Authorization": f"Bearer {access_token}",
		"Content-Type": "application/json"
	}

	try:
		response = requests.request(
			method=method,
			url=endpoint,
			headers=headers,
			json=data,
			params=params,
			timeout=30
		)

		# Handle rate limiting (429)
		if response.status_code == 429:
			retry_after = int(response.headers.get("Retry-After", 60))
			frappe.log_error(
				title="M365 Graph API Rate Limited",
				message=f"Rate limited. Retry after {retry_after} seconds. Endpoint: {endpoint}"
			)
			time.sleep(retry_after)
			# Retry the request
			return make_graph_request(endpoint, access_token, method, data, params)

		# Handle errors
		if response.status_code >= 400:
			error_data = response.json() if response.content else {}
			error_msg = error_data.get("error", {}).get("message", response.text)
			frappe.log_error(
				title=f"M365 Graph API Error: {response.status_code}",
				message=f"Endpoint: {endpoint}\nError: {error_msg}\nResponse: {response.text}"
			)
			frappe.throw(_("Graph API request failed: {0}").format(error_msg))

		return response.json() if response.content else {}

	except requests.exceptions.RequestException as e:
		frappe.log_error(
			title="M365 Graph API Request Failed",
			message=f"Endpoint: {endpoint}\nError: {str(e)}"
		)
		frappe.throw(_("Graph API request failed: {0}").format(str(e)))


def get_user_messages(user_email, access_token, folder='inbox', top=50, select=None):
	"""
	Get messages for specific user

	Args:
		user_email: User's email address
		access_token: Access token
		folder: Folder name (inbox, sentitems, etc.) or folder ID
		top: Number of messages to retrieve
		select: Fields to select (comma-separated string)

	Returns:
		dict: Response with messages
	"""
	# NOTE: Normalize folder name for Graph API compatibility
	normalizedFolder = NormalizeFolderName(folder)
	endpoint = f"/users/{user_email}/mailFolders/{normalizedFolder}/messages"

	params = {
		"$top": top,
		"$orderby": "receivedDateTime DESC"
	}

	if select:
		params["$select"] = select

	return make_graph_request(endpoint, access_token, params=params)


def get_messages_delta(user_email, access_token, folder='inbox', delta_token=None):
	"""
	Get messages using delta query for incremental sync

	Args:
		user_email: User's email address
		access_token: Access token
		folder: Folder name or ID
		delta_token: Delta token from previous sync (None for initial sync)

	Returns:
		dict: Response with messages and delta link
	"""
	if delta_token and delta_token.startswith("http"):
		# Use the stored delta link URL directly
		endpoint = delta_token
	else:
		# Initial delta query. Also self-heals if a corrupt non-URL token was
		# stored: ignore it and start a fresh delta rather than building an
		# invalid ".../v1.0<token>" URL.
		# NOTE: Normalize folder name for Graph API compatibility
		normalizedFolder = NormalizeFolderName(folder)
		endpoint = f"/users/{user_email}/mailFolders/{normalizedFolder}/messages/delta"

	return make_graph_request(endpoint, access_token)


def get_message_details(user_email, message_id, access_token):
	"""
	Get full message details including body

	Args:
		user_email: User's email address
		message_id: Message ID
		access_token: Access token

	Returns:
		dict: Message details
	"""
	endpoint = f"/users/{user_email}/messages/{message_id}"
	return make_graph_request(endpoint, access_token)


def get_message_attachments(user_email, message_id, access_token):
	"""
	Get all attachments for a message

	Args:
		user_email: User's email address
		message_id: Message ID
		access_token: Access token

	Returns:
		dict: Response with attachments
	"""
	endpoint = f"/users/{user_email}/messages/{message_id}/attachments"
	return make_graph_request(endpoint, access_token)


def download_attachment(user_email, message_id, attachment_id, access_token):
	"""
	Download specific attachment

	Args:
		user_email: User's email address
		message_id: Message ID
		attachment_id: Attachment ID
		access_token: Access token

	Returns:
		dict: Attachment data
	"""
	endpoint = f"/users/{user_email}/messages/{message_id}/attachments/{attachment_id}"
	return make_graph_request(endpoint, access_token)


def mark_message_as_read(user_email, message_id, access_token):
	"""
	Mark message as read in M365

	Args:
		user_email: User's email address
		message_id: Message ID
		access_token: Access token

	Returns:
		dict: Response
	"""
	endpoint = f"/users/{user_email}/messages/{message_id}"
	data = {"isRead": True}
	return make_graph_request(endpoint, access_token, method='PATCH', data=data)


def get_mail_folders(user_email, access_token):
	"""
	Get all mail folders for a user

	Args:
		user_email: User's email address
		access_token: Access token

	Returns:
		dict: Response with folders
	"""
	endpoint = f"/users/{user_email}/mailFolders"
	return make_graph_request(endpoint, access_token)


def get_mailbox_settings(user_email, access_token):
	"""
	Get mailbox settings (to identify shared mailboxes)

	Args:
		user_email: User's email address
		access_token: Access token

	Returns:
		dict: Mailbox settings
	"""
	endpoint = f"/users/{user_email}/mailboxSettings"
	return make_graph_request(endpoint, access_token)


def list_all_users(access_token, top=100):
	"""
	List all users in the tenant (for discovering shared mailboxes)

	Args:
		access_token: Access token
		top: Number of users to retrieve

	Returns:
		dict: Response with users
	"""
	endpoint = "/users"
	params = {"$top": top}
	return make_graph_request(endpoint, access_token, params=params)


def get_all_pages(initial_response, access_token):
	"""
	Helper function to get all pages from a paginated response

	Args:
		initial_response: Initial response dict
		access_token: Access token

	Returns:
		list: All items from all pages
	"""
	all_items = initial_response.get("value", [])
	next_link = initial_response.get("@odata.nextLink")

	while next_link:
		response = make_graph_request(next_link, access_token)
		all_items.extend(response.get("value", []))
		next_link = response.get("@odata.nextLink")

	return all_items



def get_calendar_events_delta(user_email, access_token, delta_token=None):
	"""
	Get calendar events using delta query for incremental sync

	Note: Delta queries do not support $select, $filter, $orderby, etc.
	They return all fields by default.

	Args:
		user_email: User's email address
		access_token: Access token
		delta_token: Delta token from previous sync (None for initial sync)

	Returns:
		dict: Response with events and delta link
	"""
	if delta_token and delta_token.startswith("http"):
		# Use the stored delta link URL directly
		endpoint = delta_token
	else:
		# Initial delta query - returns all fields by default. Also self-heals if
		# a corrupt non-URL token was stored rather than building an invalid URL.
		endpoint = f"/users/{user_email}/calendar/events/delta"

	return make_graph_request(endpoint, access_token)


def get_calendar_event_details(user_email, event_id, access_token):
	"""
	Get full event details

	Args:
		user_email: User's email address
		event_id: Event ID
		access_token: Access token

	Returns:
		dict: Event details
	"""
	endpoint = f"/users/{user_email}/calendar/events/{event_id}"
	return make_graph_request(endpoint, access_token)


def SetReadStatus(
	userEmail: str,
	messageID: str,
	accessToken: str,
	isRead: bool
) -> dict:
	"""
	Set read/unread status on a message in M365.

	Args:
		userEmail: User's email address
		messageID: Graph API message ID
		accessToken: Access token
		isRead: True to mark as read, False to mark as unread

	Returns:
		dict: Response data
	"""
	endpoint = f"/users/{userEmail}/messages/{messageID}"
	data = {"isRead": isRead}
	return make_graph_request(endpoint, accessToken, method="PATCH", data=data)



def SetReadStatusSafe(
	userEmail: str,
	messageID: str,
	accessToken: str,
	isRead: bool
) -> dict:
	"""
	Set read/unread status on a message in M365 (non-throwing).
	Logs errors but does not raise exceptions.

	Args:
		userEmail: User's email address
		messageID: Graph API message ID
		accessToken: Access token
		isRead: True to mark as read, False to mark as unread

	Returns:
		dict: {"success": bool, "message": str}
	"""
	statusLabel = "read" if isRead else "unread"
	try:
		SetReadStatus(userEmail, messageID, accessToken, isRead)
		return {
			"success": True,
			"message": f"Marked as {statusLabel}",
		}
	except Exception as e:
		frappe.log_error(
			title="M365 Read Status Sync Failed",
			message=(
				f"Failed to mark message {messageID} as {statusLabel} "
				f"for {userEmail}: {str(e)}"
			)
		)
		return {
			"success": False,
			"message": str(e),
		}



def MoveMessage(
	userEmail: str,
	messageID: str,
	accessToken: str,
	destinationFolder: str
) -> dict:
	"""
	Move a message to a different folder in M365.

	Args:
		userEmail: User's email address
		messageID: Graph API message ID
		accessToken: Access token
		destinationFolder: Target folder name (archive, trash, spam, inbox)

	Returns:
		dict: Response with moved message data (includes new message ID)
	"""
	normalizedFolder = NormalizeFolderName(destinationFolder)
	endpoint = f"/users/{userEmail}/messages/{messageID}/move"
	data = {"destinationId": normalizedFolder}
	return make_graph_request(endpoint, accessToken, method="POST", data=data)



def MoveMessageSafe(
	userEmail: str,
	messageID: str,
	accessToken: str,
	destinationFolder: str
) -> dict:
	"""
	Move a message to a different folder in M365 (non-throwing).
	Logs errors but does not raise exceptions.

	Args:
		userEmail: User's email address
		messageID: Graph API message ID
		accessToken: Access token
		destinationFolder: Target folder name

	Returns:
		dict: {"success": bool, "message": str, "data": dict or None}
	"""
	try:
		responseData = MoveMessage(
			userEmail, messageID, accessToken, destinationFolder
		)
		return {
			"success": True,
			"message": f"Moved to {destinationFolder}",
			"data": responseData
		}
	except Exception as e:
		frappe.log_error(
			title="M365 Provider Sync Failed",
			message=(
				f"Failed to move message {messageID} to {destinationFolder} "
				f"for {userEmail}: {str(e)}"
			)
		)
		return {
			"success": False,
			"message": str(e),
			"data": None
		}



def send_email_as_user(sender_email, recipients, subject, body, access_token, cc=None, bcc=None, attachments=None, is_html=True):
	"""
	Send email as a specific user using Microsoft Graph API
	Requires Mail.Send application permission

	Args:
		sender_email: Email address to send from
		recipients: List of recipient email addresses
		subject: Email subject
		body: Email body content
		access_token: Access token with Mail.Send permission
		cc: List of CC email addresses (optional)
		bcc: List of BCC email addresses (optional)
		attachments: List of attachment dicts with 'name' and 'content' (base64) (optional)
		is_html: Whether body is HTML (default: True)

	Returns:
		dict: Response with success status and message ID
	"""
	# Build recipient list
	to_recipients = [{"emailAddress": {"address": email}} for email in recipients]
	cc_recipients = [{"emailAddress": {"address": email}} for email in (cc or [])]
	bcc_recipients = [{"emailAddress": {"address": email}} for email in (bcc or [])]

	# Build message
	message = {
		"subject": subject,
		"body": {
			"contentType": "HTML" if is_html else "Text",
			"content": body
		},
		"toRecipients": to_recipients
	}

	if cc_recipients:
		message["ccRecipients"] = cc_recipients

	if bcc_recipients:
		message["bccRecipients"] = bcc_recipients

	# Add attachments if provided
	if attachments:
		message["attachments"] = []
		for attachment in attachments:
			message["attachments"].append({
				"@odata.type": "#microsoft.graph.fileAttachment",
				"name": attachment.get("name"),
				"contentBytes": attachment.get("content")  # Base64 encoded
			})

	# Send email
	endpoint = f"/users/{sender_email}/sendMail"
	data = {
		"message": message,
		"saveToSentItems": True
	}

	try:
		response = make_graph_request(
			endpoint=endpoint,
			access_token=access_token,
			method="POST",
			data=data
		)

		return {
			"success": True,
			"message": "Email sent successfully"
		}
	except Exception as e:
		return {
			"success": False,
			"message": str(e)
		}
