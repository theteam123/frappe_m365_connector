# Copyright (c) 2026, TierneyMorris Pty Ltd and contributors
# For license information, please see license.txt

"""
Place a ready-made, received-looking email straight into a named folder of a mailbox.

Why this exists: some events that are not email still belong in a team's mailbox next to
its email. The first is an SMS reply on the recruitment line, which recruiters have always
found in the recruitment mailbox under Inbox/Bulk SMS (ERP13 emailed it there and a mail
rule filed it). Writing it straight into the folder needs no mail rule and sends nothing.

Graph creates every message posted to a folder as an unsent draft. Setting the MAPI
message-flags property at creation is what makes it an ordinary received message instead:
the flags are written once, at creation, and can never be changed afterwards.

Needs the Mail.ReadWrite application permission on the mailbox's app login. Not
whitelisted: callers are server-side jobs only.
"""

import datetime
from urllib.parse import quote

import frappe
from frappe import _

from m365email.m365email.auth import get_access_token
from m365email.m365email.draft_mail import ResolveServicePrincipal
from m365email.m365email.graph_api import FOLDER_NAME_MAP, NormalizeFolderName, make_graph_request

FOLDER_PATH_SEPARATOR = "/"
ROOT_FOLDERS_ENDPOINT = "/users/{mailbox}/mailFolders"
CHILD_FOLDERS_ENDPOINT = "/users/{mailbox}/mailFolders/{folderID}/childFolders"
WELL_KNOWN_FOLDER_ENDPOINT = "/users/{mailbox}/mailFolders/{folderName}"
FOLDER_MESSAGES_ENDPOINT = "/users/{mailbox}/mailFolders/{folderID}/messages"

# NOTE: MAPI property tags as Graph names them. PR_MESSAGE_FLAGS (0x0E07) decides draft or
# received; 0 means "received, not yet read", so it shows bold like any new email.
# PR_MESSAGE_DELIVERY_TIME (0x0E06) and PR_CLIENT_SUBMIT_TIME (0x0039) set the date Outlook
# shows and sorts by; without them the message is dated when it was written.
MESSAGE_FLAGS_PROPERTY = "Integer 0x0E07"
MESSAGE_FLAGS_RECEIVED_UNREAD = "0"
DELIVERY_TIME_PROPERTY = "SystemTime 0x0E06"
SUBMIT_TIME_PROPERTY = "SystemTime 0x0039"
GRAPH_UTC_FORMAT = "%Y-%m-%dT%H:%M:%SZ"



def QuotePath(value: str) -> str:
	"""A value made safe as one URL path segment: every '/' and '?' is encoded too."""
	return quote(value, safe="")



def FindFolderByPath(mailbox: str, folderPath: str, accessToken: str) -> str:
	"""The Graph id of a folder named by path, e.g. "Inbox/Bulk SMS". Throws when absent.

	The first segment may be a well-known folder (Inbox, Archive, ...) or any top-level
	folder's display name; each later segment is a child folder's display name. Names are
	matched without regard to case, as Outlook itself does.
	"""
	segments = [segment.strip() for segment in (folderPath or "").split(FOLDER_PATH_SEPARATOR)]
	segments = [segment for segment in segments if segment]
	if not segments:
		frappe.throw(_("No mailbox folder given"))

	folderID = FindRootFolder(mailbox, segments[0], accessToken)
	for segment in segments[1:]:
		folderSegment = QuotePath(folderID)
		endpoint = CHILD_FOLDERS_ENDPOINT.format(mailbox=QuotePath(mailbox), folderID=folderSegment)
		folderID = FindChildFolder(endpoint, segment, accessToken)
		if not folderID:
			frappe.throw(_("Folder {0} not found in {1}").format(folderPath, mailbox))

	return folderID



def FindRootFolder(mailbox: str, folderName: str, accessToken: str) -> str:
	"""A top-level folder's id: a well-known name directly, otherwise by display name."""
	if folderName.lower() in FOLDER_NAME_MAP:
		wellKnownName = NormalizeFolderName(folderName)
		mailboxSegment = QuotePath(mailbox)
		nameSegment = QuotePath(wellKnownName)
		endpoint = WELL_KNOWN_FOLDER_ENDPOINT.format(mailbox=mailboxSegment, folderName=nameSegment)
		return make_graph_request(endpoint, accessToken, params={"$select": "id"})["id"]

	endpoint = ROOT_FOLDERS_ENDPOINT.format(mailbox=QuotePath(mailbox))
	folderID = FindChildFolder(endpoint, folderName, accessToken)
	if not folderID:
		frappe.throw(_("Folder {0} not found in {1}").format(folderName, mailbox))
	return folderID



def FindChildFolder(endpoint: str, folderName: str, accessToken: str) -> str | None:
	"""The id of the folder under `endpoint` whose display name matches, or None."""
	# NOTE: OData string literals escape a single quote by doubling it.
	escapedName = folderName.replace("'", "''")
	params = {"$filter": f"displayName eq '{escapedName}'", "$select": "id,displayName"}
	matches = make_graph_request(endpoint, accessToken, params=params).get("value") or []
	if not matches:
		return None
	return matches[0]["id"]



def FileMessageInFolder(mailbox: str, folderPath: str, message: dict) -> str:
	"""Create `message` in the folder as a received, unread email. Returns its Graph id.

	`message` carries subject, html, senderName, senderAddress and receivedAt (a timezone-
	aware datetime). Nothing is sent: the message only ever exists in this one mailbox.
	"""
	servicePrincipal = ResolveServicePrincipal(mailbox)
	accessToken = get_access_token(servicePrincipal)
	folderID = FindFolderByPath(mailbox, folderPath, accessToken)

	folderSegment = QuotePath(folderID)
	endpoint = FOLDER_MESSAGES_ENDPOINT.format(mailbox=QuotePath(mailbox), folderID=folderSegment)
	payload = BuildMessagePayload(mailbox, message)
	created = make_graph_request(endpoint, accessToken, method="POST", data=payload)
	return created.get("id") or ""



def BuildMessagePayload(mailbox: str, message: dict) -> dict:
	"""The Graph message body for FileMessageInFolder."""
	if message["receivedAt"].tzinfo is None:
		frappe.throw(_("receivedAt must carry a timezone, or its time would be misread"))

	# NOTE: Graph reads a SystemTime property as UTC, so the time is converted, not relabelled.
	receivedAt = message["receivedAt"].astimezone(datetime.timezone.utc).strftime(GRAPH_UTC_FORMAT)
	sender = {"emailAddress": {"name": message["senderName"], "address": message["senderAddress"]}}

	return {
		"subject": message["subject"],
		"body": {"contentType": "HTML", "content": message["html"]},
		"from": sender,
		"sender": sender,
		"toRecipients": [{"emailAddress": {"address": mailbox}}],
		"singleValueExtendedProperties": [
			{"id": MESSAGE_FLAGS_PROPERTY, "value": MESSAGE_FLAGS_RECEIVED_UNREAD},
			{"id": DELIVERY_TIME_PROPERTY, "value": receivedAt},
			{"id": SUBMIT_TIME_PROPERTY, "value": receivedAt},
		],
	}
