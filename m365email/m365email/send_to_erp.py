"""Send-to-ERP API for the Outlook add-in.

Lets an Outlook add-in file an email (and/or its attachments) against any
permitted ERP record. The add-in panel is served same-origin from this site,
so these endpoints rely on the standard Frappe login session for auth.

All public endpoints enforce normal Frappe document permissions — nothing is
inserted with ignore_permissions, so a user can only file against records they
are allowed to see and attach to.
"""
import base64
import binascii
import hashlib
import json
import re
from typing import Callable, Optional

import frappe
from frappe import _

from m365email.m365email import send_to_erp_people as people


# Doctypes a user may file an email against, in picker order. Each user only
# sees the subset they have read access to (see GetPickableDoctypes).
# NOTE: Candidate and Employee added for recruitment (Paul, 2026-10-23); sites without hr skip
# them, because GetPickableDoctypes drops doctypes that do not exist. Supplier Quote added for
# procurement (Paul, 2026-11-06); it exists only where the erp app is installed.
DEFAULT_PICKABLE_DOCTYPES = [
	"Contact",
	"Candidate",
	"Employee",
	"Project",
	"Purchase Order",
	"Supplier Quote",
	"Supplier",
	"Sales Invoice",
	"Purchase Invoice",
]

MAX_SEARCH_RESULTS = 20
MAX_ATTACHMENT_BYTES = 35 * 1024 * 1024  # 35 MB — above Outlook's own attach limit
MAX_TOTAL_ATTACHMENT_BYTES = 100 * 1024 * 1024  # all attachments of one filing together
BASE64_TEXT = re.compile(r"[A-Za-z0-9+/]*={0,2}")  # matched after line breaks are removed
CONTACT_EXTRACTION_MAX_CHARS = 8000  # email body chars sent to AI (covers most signatures)
CONTACT_EXTRACTION_MAX_TOKENS = 400  # AI response cap for contact-field extraction

# Secondary detail shown beneath the name in the picker, to disambiguate records
# that share a title (e.g. two contacts both named "Paul Johnson").
SECONDARY_FIELDS = {
	"Contact": "email_id",
	"Candidate": "email_address",
	"Employee": "work_email_address",
}

# Special pseudo-target: rather than an existing record, the user picks a person
# and a new ToDo (task) is created for them, with the email filed against it.
TASK_TARGET = "ToDo"
TASK_TARGET_LABEL = "Task / To-Do"

# Special pseudo-target: raise a new Support Ticket from the email. The "record"
# the user picks is the ticket type (Issue/Bug, Feature Request, ...). Tickets go
# through the support app's CreateTicket, so a remote (proxy-mode) site such as
# staging forwards them to the primary helpdesk exactly as the helpdesk UI does.
# NOTE: Added for Paul, 2026-11-04.
TICKET_TARGET = "Support Ticket"
TICKET_TARGET_LABEL = "Create support ticket"
TICKET_PRIORITY = "Medium"  # triage sets the real priority on the helpdesk
TICKET_BODY_MAX_CHARS = 20000  # email text copied into the ticket description
TICKET_UPLOAD_TIMEOUT_SECONDS = 120  # per attachment, when forwarding to the primary site
TICKET_UPLOAD_METHOD = "support.support.page.helpdesk.data.ticket.UploadTicketAttachment"

# Targets that are not existing records to search, but something new to create.
PSEUDO_TARGETS = (TASK_TARGET, TICKET_TARGET)

# "Create new" on a Supplier Quote raises a quote from the email (send_to_erp_procurement).
SUPPLIER_QUOTE_TARGET = "Supplier Quote"

# Procurement records show their supplier's name beneath their number in the picker. The
# supplier field is a Link to Organisation, whose id is an opaque key, so the name is looked up.
SUPPLIER_SUBLABEL_DOCTYPES = ("Purchase Order", "Supplier Quote", "Purchase Invoice")

# A "create" action from one email runs once: a second click, or a colleague opening the same
# email in a shared mailbox, gets the record already made instead of a duplicate. The record is
# remembered for this long, and a click waits this long for one still being created.
CREATE_GUARD_SECONDS = 30 * 60
CREATE_GUARD_LOCK_TIMEOUT_SECONDS = 60
CREATE_GUARD_PREFIX = "sendtoerp-create"



@frappe.whitelist()
def GetSession() -> dict:
	"""Return the logged-in user and a CSRF token for the add-in panel.

	The panel is served same-origin and authenticates via the standard login
	session; it calls this (GET, no CSRF needed) after login to obtain the
	token required for the state-changing FileEmailToRecord POST.

	Returns:
		{user, csrf_token} for the current session.
	"""
	return {
		"user": frappe.session.user,
		"csrf_token": frappe.sessions.get_csrf_token(),
	}



@frappe.whitelist()
def GetPickableDoctypes() -> list[dict]:
	"""Return the doctypes the current user may file an email against.

	Returns:
		List of {doctype, label} the user has read permission on, in picker order.
	"""
	pickable = []
	for doctype in DEFAULT_PICKABLE_DOCTYPES:
		if not frappe.db.exists("DocType", doctype):
			continue

		if not frappe.has_permission(doctype, "read"):
			continue

		pickable.append({"doctype": doctype, "label": _(doctype)})

	if frappe.has_permission(TASK_TARGET, "create"):
		pickable.append({"doctype": TASK_TARGET, "label": _(TASK_TARGET_LABEL)})

	# Sites without the support app have no Support Ticket doctype, so no option.
	hasTicketDoctype = frappe.db.exists("DocType", TICKET_TARGET)
	if hasTicketDoctype and frappe.has_permission(TICKET_TARGET, "create"):
		pickable.append({"doctype": TICKET_TARGET, "label": _(TICKET_TARGET_LABEL)})

	return pickable



@frappe.whitelist()
def SearchTargets(target_doctype: Optional[str] = None, txt: str = "") -> list[dict]:
	"""Search records of a doctype for the add-in's record picker.

	Args:
		target_doctype: Target doctype — must be in the pickable allow-list.
			(``doctype`` is a Frappe-reserved request key, so it is read from the
			raw form_dict as a fallback for backward compatibility.)
		txt: Partial text to match against name or the doctype's title field.

	Returns:
		List of {value, label, sublabel} matches, capped at MAX_SEARCH_RESULTS.
	"""
	doctype = target_doctype or frappe.form_dict.get("doctype")
	if not doctype:
		frappe.throw(_("No target type was provided."))

	_GuardPickableDoctype(doctype)

	# The Task/To-Do target searches people to assign to, not existing records.
	if doctype == TASK_TARGET:
		return _SearchUsers(txt)

	# The support-ticket target lists ticket types, not existing tickets.
	if doctype == TICKET_TARGET:
		return SearchTicketTypes(txt)

	if not frappe.has_permission(doctype, "read"):
		frappe.throw(_("You do not have access to {0}").format(_(doctype)), frappe.PermissionError)

	titleField = _GetTitleField(doctype)
	fields = ["name"]
	if titleField and titleField != "name":
		fields.append(titleField)

	secondaryField = _GetSecondaryField(doctype)
	if secondaryField and secondaryField not in fields:
		fields.append(secondaryField)
	if doctype in SUPPLIER_SUBLABEL_DOCTYPES:
		fields.append("supplier")

	searchText = (txt or "").strip()
	orFilters = {}
	if searchText:
		orFilters["name"] = ["like", f"%{searchText}%"]
		if titleField and titleField != "name":
			orFilters[titleField] = ["like", f"%{searchText}%"]

	records = frappe.get_list(
		doctype,
		or_filters=orFilters or None,
		fields=fields,
		limit_page_length=MAX_SEARCH_RESULTS,
		order_by="modified desc",
	)

	results = [_FormatSearchRow(row, titleField, secondaryField) for row in records]
	if doctype in SUPPLIER_SUBLABEL_DOCTYPES:
		AddSupplierSublabels(doctype, records, results)
	return results



@frappe.whitelist()
def FileEmailToRecord(
	target_doctype: str,
	target_name: str,
	subject: str = "",
	sender: str = "",
	recipients: str = "",
	sent_date: Optional[str] = None,
	body_html: str = "",
	save_email: int = 1,
	comment: str = "",
	create_new: int = 0,
	new_name: str = "",
	new_email: str = "",
	attachments: str = "[]",
	project: str = "",
	supplier: str = "",
	message_id: str = "",
	contact: str = "",
) -> dict:
	"""File an email and/or its attachments against a record or a new task.

	Args:
		target_doctype: Filing target — a pickable doctype, ``ToDo`` to create
			a new task (``target_name`` is the assignee's user id), or
			``Support Ticket`` to raise a ticket (``target_name`` is the ticket type).
		target_name: Record to file against, the assignee for a Task/To-Do, or
			the ticket type for a Support Ticket.
		subject: Email subject line.
		sender: Sender email address.
		recipients: Recipient email addresses (comma separated).
		sent_date: ISO timestamp the email was sent/received.
		body_html: HTML body of the email.
		save_email: 1 to record the email itself on the timeline.
		comment: Optional note — a timeline comment on a record, or the task's
			description when creating a Task/To-Do.
		create_new: 1 to create a new Contact or Candidate (from new_name/new_email)
			and file against it, instead of using an existing record.
		new_name: Sender's display name for the new record when create_new is set.
		new_email: Sender's email address for the new record when create_new is set.
		attachments: JSON list of {file_name, content_base64, content_type}.
		project: For "create new" Supplier Quote, the job the quote is for.
		supplier: For "create new" Supplier Quote, a supplier the user picked (optional:
			otherwise it is taken from the sender's contact, then the quote PDF).
		message_id: The email's Internet message id. With it, a "create new" runs once per
			email: repeating it returns the record already created.
		contact: A Contact the user picked as the person the email is about (Paul,
			2026-11-07: often a forwarded email's original sender). The filed email also shows
			on that Contact's timeline; a new Supplier Quote takes it as supplier contact; a
			support ticket names it.

	Returns:
		{success, communication, files, message, record, record_url} describing what was
		created or filed to.
	"""
	_GuardPickableDoctype(target_doctype)

	if int(create_new or 0) and message_id:
		def createOnce() -> dict:
			return FileEmailToRecord(
				target_doctype, target_name, subject, sender, recipients, sent_date, body_html,
				save_email, comment, create_new, new_name, new_email, attachments, project, supplier,
				"", contact,
			)
		# NOTE: The job is part of the key, so a quote for a second job is not mistaken for a repeat.
		return RunOncePerEmail(f"create-{target_doctype}-{project}", message_id, createOnce)

	shouldSaveEmail = bool(int(save_email or 0))
	attachmentList = _ParseAttachments(attachments)
	commentText = (comment or "").strip()
	isTask = target_doctype == TASK_TARGET
	contactName = people.ResolveChosenContact(contact)
	shouldCreateNew = bool(int(create_new or 0))

	emailFields = {
		"subject": subject,
		"sender": sender,
		"recipients": recipients,
		"sent_date": sent_date,
		"body_html": body_html,
	}

	# A ticket may live on another site (proxy mode), so it takes its own path.
	if target_doctype == TICKET_TARGET:
		aboutWhom = people.DescribeContact(contactName)
		if aboutWhom:
			aboutLine = _("About: {0}").format(aboutWhom)
			commentText = "\n".join(part for part in (aboutLine, commentText) if part)
		return FileEmailAsTicket(
			target_name, emailFields, shouldSaveEmail, commentText, attachmentList,
		)

	communicationName = None
	creationNote = ""
	if isTask:
		recordDoctype, recordName, communicationName = CreateTaskFromEmail(
			target_name, commentText, emailFields, shouldSaveEmail,
		)
	elif shouldCreateNew:
		if target_doctype == SUPPLIER_QUOTE_TARGET:
			from m365email.m365email.send_to_erp_procurement import CreateSupplierQuoteFromEmail

			# NOTE: Not called "sender": that is this endpoint's sender address, used below.
			senderDetails = {
				"name": new_name, "email": new_email, "body_html": body_html, "sent_date": sent_date,
				"contact": contactName or "",
			}
			recordDoctype, recordName, creationNote = CreateSupplierQuoteFromEmail(
				project, supplier, senderDetails, attachmentList,
			)
		else:
			recordDoctype, recordName = CreateRecordFromSender(
				target_doctype, new_name, new_email, body_html,
			)
		# "Create new" may reuse a record already holding the sender's email, and
		# filing then writes to it, so apply the same write check as a picked record.
		# The message names no record: the user may not be allowed to read it.
		# A refusal rolls back a record created in this request.
		if not frappe.has_permission(recordDoctype, "write", recordName):
			message = _("You do not have permission to add to the {0} record for this sender.")
			frappe.throw(message.format(_(recordDoctype)), frappe.PermissionError)
	else:
		_GuardTargetExists(target_doctype, target_name)
		if not frappe.has_permission(target_doctype, "write", target_name):
			frappe.throw(
				_("You do not have permission to attach to {0} {1}").format(
					_(target_doctype), _ResolveRecordLabel(target_doctype, target_name)
				),
				frappe.PermissionError,
			)
		if not shouldSaveEmail and not attachmentList and not commentText:
			frappe.throw(_("Add a comment, or select the email and/or an attachment to file."))
		recordDoctype, recordName = target_doctype, target_name

	# Attach files to the target record itself so they show in the record's
	# standard attachments, and capture their URLs for the email body.
	createdFiles = _AttachFiles(attachmentList, recordDoctype, recordName)

	# NOTE: A new Supplier Quote: import its PDF's lines when nothing needs a person, else
	#       send "Open in ERP" to the import page with the PDF loaded (Paul, 2026-11-07).
	#       Imported here, not at the top, because the procurement module imports this one.
	recordURL = frappe.utils.get_url_to_form(recordDoctype, recordName)
	if creationNote and recordDoctype == SUPPLIER_QUOTE_TARGET:
		from m365email.m365email.send_to_erp_procurement import AutoImportNewQuote

		importOutcome = AutoImportNewQuote(recordName, createdFiles)
		creationNote = " ".join(part for part in (creationNote, importOutcome["note"]) if part)
		recordURL = importOutcome["record_url"] or recordURL

	# Build the email body: on a record, a filing note rides at the top so the
	# comment shares the email's timeline entry; then the attachment links.
	# (A task's email already exists — it was filed before the task so the task
	# could reference it — so only the attachment links are added to it here.)
	if isTask:
		AppendAttachmentLinksToCommunication(communicationName, createdFiles)
	elif shouldSaveEmail:
		emailBody = body_html
		if commentText:
			emailBody = _PrependNote(emailBody, commentText)
		communicationName = _CreateCommunication(
			recordDoctype, recordName, subject, sender, recipients, sent_date,
			_AppendAttachmentLinks(emailBody, createdFiles),
		)
		if recordDoctype != people.CONTACT:
			people.LinkEmailToContact(communicationName, contactName)

	# A comment with no email to ride along becomes a standalone timeline comment.
	# (A task already carries the comment as its description.)
	if commentText and not isTask and not shouldSaveEmail:
		_AddComment(recordDoctype, recordName, commentText)

	frappe.db.commit()

	# For a task, name the assignee in the message (not the new ToDo's id).
	summaryName = target_name if isTask else _ResolveRecordLabel(recordDoctype, recordName)
	hasComment = bool(commentText)
	summaryParts = (isTask, shouldSaveEmail, createdFiles, recordDoctype, summaryName, hasComment)
	summary = _BuildSummary(*summaryParts)
	if creationNote:
		created = _("Created {0} {1}.").format(_(recordDoctype), summaryName)
		summary = " ".join((created, creationNote, summary))
	return {
		"success": True,
		"communication": communicationName,
		"files": createdFiles,
		"record": {"doctype": recordDoctype, "name": recordName},
		"record_url": recordURL,
		"message": summary,
	}



def RunOncePerEmail(action: str, messageID: str, createRecord: Callable[[], dict]) -> dict:
	"""Run a "create" action at most once per email, returning the earlier record on a repeat.

	A filesystem lock serialises two clicks arriving together; the record made is remembered
	in the cache, keyed by the action and a hash of the email's message id (never the raw id).
	The guard is shared by everyone, so a colleague opening the same email in a shared mailbox
	is not able to create a duplicate either; they are only told the record's name when they
	may read it.

	Args:
		action: What is being created, e.g. "create-Supplier Quote".
		messageID: The email's Internet message id; without one there is nothing to key on
			and the action simply runs.
		createRecord: Creates the record, commits, and returns the endpoint's result, which
			carries {"record": {doctype, name}}.

	Returns:
		The new result, or an "already created" result naming the earlier record.
	"""
	from frappe.utils.synchronization import filelock

	if not messageID:
		return createRecord()

	guardKey = BuildCreateGuardKey(action, messageID)
	with filelock(guardKey, timeout=CREATE_GUARD_LOCK_TIMEOUT_SECONDS):
		earlier = frappe.cache.get_value(guardKey)
		if earlier and frappe.db.exists(earlier.get("doctype"), earlier.get("name")):
			return BuildAlreadyCreatedResult(earlier)

		result = createRecord()
		record = result.get("record") or {}
		if record.get("doctype") and record.get("name"):
			frappe.cache.set_value(guardKey, record, expires_in_sec=CREATE_GUARD_SECONDS)
		return result



def BuildCreateGuardKey(action: str, messageID: str) -> str:
	"""A filesystem- and cache-safe key: the action and the message id, hashed together."""
	digest = hashlib.sha256(f"{action}\n{messageID}".encode("utf-8")).hexdigest()
	return f"{CREATE_GUARD_PREFIX}-{digest}"



def BuildAlreadyCreatedResult(record: dict) -> dict:
	"""The reply to a repeated "create": name the earlier record only to someone who may read it."""
	doctype, name = record["doctype"], record["name"]
	if not frappe.has_permission(doctype, "read", name):
		message = _("A {0} has already been created from this email.").format(_(doctype))
		return {
			"success": True,
			"communication": None,
			"files": [],
			"message": message,
			"already_created": True,
		}

	message = _("{0} {1} was already created from this email, so nothing new was made.")
	return {
		"success": True,
		"communication": None,
		"files": [],
		"record": record,
		"record_url": frappe.utils.get_url_to_form(doctype, name),
		"message": message.format(_(doctype), _ResolveRecordLabel(doctype, name)),
		"already_created": True,
	}



def AddSupplierSublabels(doctype: str, records: list[dict], results: list[dict]) -> None:
	"""Show each procurement record's supplier name beneath its number in the picker."""
	suppliers = {row.get("supplier") for row in records if row.get("supplier")}
	if not suppliers:
		return
	# NOTE: A user who may not read suppliers still gets their search results, just unlabelled.
	if not frappe.has_permission("Organisation", "read"):
		return

	names = dict(frappe.get_list(
		"Organisation",
		filters={"name": ["in", list(suppliers)]},
		fields=["name", "organisation_name"],
		as_list=True,
	))
	for row, result in zip(records, results):
		supplierName = names.get(row.get("supplier"))
		if supplierName:
			result["sublabel"] = supplierName



def _GuardPickableDoctype(doctype: str) -> None:
	"""Reject any target not in the pickable allow-list (Task/To-Do included)."""
	if doctype not in PSEUDO_TARGETS and doctype not in DEFAULT_PICKABLE_DOCTYPES:
		frappe.throw(_("{0} is not available as a filing target.").format(doctype), frappe.PermissionError)


def _SearchUsers(txt: str) -> list[dict]:
	"""Search enabled system users to assign a task to, for the add-in picker.

	Returns:
		List of {value, label, sublabel} where value is the user id (email),
		label the full name, and sublabel the email — capped at MAX_SEARCH_RESULTS.
	"""
	searchText = (txt or "").strip()
	orFilters = {}
	if searchText:
		orFilters["full_name"] = ["like", f"%{searchText}%"]
		orFilters["name"] = ["like", f"%{searchText}%"]

	users = frappe.get_list(
		"User",
		filters={"enabled": 1, "user_type": "System User"},
		or_filters=orFilters or None,
		fields=["name", "full_name"],
		limit_page_length=MAX_SEARCH_RESULTS,
		order_by="full_name asc",
	)
	return [{"value": user["name"], "label": user.get("full_name") or user["name"], "sublabel": user["name"]} for user in users]



def _GuardTargetExists(doctype: str, name: str) -> None:
	"""Reject a missing target record."""
	if not name or not frappe.db.exists(doctype, name):
		frappe.throw(_("{0} was not found.").format(_(doctype)))



def _ResolveRecordLabel(doctype: str, name: str) -> str:
	"""Return the human-readable title for a filing target, never the raw id when a title exists."""
	if not doctype or not name:
		return name or ""
	titleField = _GetTitleField(doctype)
	if not titleField or titleField == "name":
		return name
	title = frappe.utils.cstr(frappe.get_cached_value(doctype, name, titleField)).strip()
	return title or name



def _GetTitleField(doctype: str) -> Optional[str]:
	"""Return the doctype's configured title field, if any."""
	meta = frappe.get_meta(doctype)
	return meta.get_title_field() if meta.title_field else None



def _GetSecondaryField(doctype: str) -> Optional[str]:
	"""Return a disambiguating sub-field for the doctype, if it has one.

	Some doctypes (notably Contact) can have many records sharing a title, so the
	picker shows a secondary detail beneath the name. Only a field that actually
	exists on the doctype is returned.
	"""
	field = SECONDARY_FIELDS.get(doctype)
	if field and frappe.get_meta(doctype).has_field(field):
		return field
	return None



def _FormatSearchRow(row: dict, titleField: Optional[str], secondaryField: Optional[str] = None) -> dict:
	"""Format one search result as {value, label, sublabel} for the picker.

	The label shows the human-readable title (e.g. project number/name) and
	never the raw record id, which for several doctypes is an opaque UUID. The
	sublabel carries a disambiguating detail (e.g. a contact's email) so records
	with identical names can be told apart.
	"""
	name = row.get("name")
	title = row.get(titleField) if titleField else None
	label = title if title else name
	sublabel = row.get(secondaryField) if secondaryField else None
	return {"value": name, "label": label, "sublabel": sublabel or ""}



def _ParseAttachments(attachments: str) -> list[dict]:
	"""Parse and size-check the attachments JSON payload.

	Args:
		attachments: JSON list of {file_name, content_base64, content_type}.

	Returns:
		List of attachment dicts with decoded-size validated.
	"""
	try:
		parsed = json.loads(attachments or "[]")
	except (ValueError, TypeError):
		frappe.throw(_("Attachment payload was not valid."))

	if not isinstance(parsed, list):
		frappe.throw(_("Attachment payload was not valid."))

	totalBytes = 0
	for item in parsed:
		if not isinstance(item, dict):
			frappe.throw(_("Attachment payload was not valid."))
		content = item.get("content_base64", "")
		# NOTE: Checked here so a malformed attachment is a clear message, never a server error.
		# A length that is not a multiple of 4 cannot be base64 the browser produced.
		if not isinstance(content, str):
			frappe.throw(_("Attachment {0} could not be read.").format(item.get("file_name", "")))
		compact = content.replace("\r", "").replace("\n", "")
		if len(compact) % 4 or not BASE64_TEXT.fullmatch(compact):
			frappe.throw(_("Attachment {0} could not be read.").format(item.get("file_name", "")))
		# base64 expands ~4/3; estimate decoded size without decoding twice.
		estimatedBytes = (len(content) * 3) // 4
		if estimatedBytes > MAX_ATTACHMENT_BYTES:
			frappe.throw(_("Attachment {0} is too large.").format(item.get("file_name", "")))
		totalBytes += estimatedBytes
		if totalBytes > MAX_TOTAL_ATTACHMENT_BYTES:
			frappe.throw(_("The attachments are too large to file together. File fewer at once."))

	return parsed



def _AppendAttachmentLinks(body_html: str, files: list[dict]) -> str:
	"""Append a clickable list of the filed attachments to the email body.

	The record timeline renders the Communication's HTML, so listing the files
	here surfaces them directly beneath the email — in addition to their place in
	the record's own attachments.
	"""
	if not files:
		return body_html or ""

	htmlItems = "".join(
		'<li><a href="{0}">{1}</a></li>'.format(
			frappe.utils.escape_html(file["file_url"]),
			frappe.utils.escape_html(file["file_name"]),
		)
		for file in files
	)
	heading = _("Attachments filed to this record")
	return (body_html or "") + "<hr><p><strong>{0}</strong></p><ul>{1}</ul>".format(heading, htmlItems)


def _PrependNote(body_html: str, comment: str) -> str:
	"""Prepend a filing note to the email body (escaped, styled) so the comment
	sits inside the email's own timeline entry, directly above the message."""
	note = frappe.utils.escape_html(comment).replace("\n", "<br>")
	block = (
		'<div style="border-left:4px solid #3fd921;background:#f0fdf4;'
		'padding:8px 12px;margin:0 0 12px 0;border-radius:4px;">'
		"<strong>{0}</strong><br>{1}</div>"
	).format(_("Note added when filing"), note)
	return block + (body_html or "")



def _CreateCommunication(
	target_doctype: str,
	target_name: str,
	subject: str,
	sender: str,
	recipients: str,
	sent_date: Optional[str],
	body_html: str,
) -> str:
	"""Create a received-email Communication linked to the target record.

	Returns:
		The new Communication's name.
	"""
	communication = frappe.get_doc({
		"doctype": "Communication",
		"communication_type": "Communication",
		"communication_medium": "Email",
		"sent_or_received": "Received",
		"subject": subject or _("(no subject)"),
		"content": body_html or "",
		"sender": sender or "",
		"recipients": recipients or "",
		"reference_doctype": target_doctype,
		"reference_name": target_name,
		"communication_date": _ParseEmailDate(sent_date),
	})
	# Authorization is enforced upstream via write permission on the target
	# record; the linked Communication is created on the user's behalf.
	communication.insert(ignore_permissions=True)
	return communication.name



def _ParseEmailDate(value: Optional[str]):
	"""Convert an email ISO timestamp to a system-timezone datetime.

	Outlook sends an ISO 8601 UTC string like ``2026-06-17T11:18:54.000Z``,
	which the database cannot store directly. Parse it, convert UTC to the
	site timezone, and fall back to now if it cannot be parsed.
	"""
	if not value:
		return frappe.utils.now_datetime()

	import datetime as dtmod
	try:
		parsed = dtmod.datetime.fromisoformat(value.replace("Z", "+00:00"))
	except (ValueError, TypeError):
		return frappe.utils.now_datetime()

	if parsed.tzinfo is None:
		return parsed

	utcNaive = parsed.astimezone(dtmod.timezone.utc).replace(tzinfo=None)
	return frappe.utils.convert_utc_to_system_timezone(utcNaive).replace(tzinfo=None)



def _AttachFiles(attachmentList: list[dict], parentDoctype: str, parentName: str) -> list[dict]:
	"""Create File records for each attachment under the given parent.

	Args:
		attachmentList: Parsed attachment dicts.
		parentDoctype: Doctype to attach the files to.
		parentName: Record name to attach the files to.

	Returns:
		List of {file_name, file_url} created.
	"""
	created = []
	for item in attachmentList:
		fileName = item.get("file_name") or "attachment"
		content = item.get("content_base64")
		if not content:
			continue

		try:
			decoded = base64.b64decode(content)
		except (binascii.Error, ValueError):
			frappe.throw(_("Attachment {0} could not be read.").format(fileName))

		fileDoc = frappe.get_doc({
			"doctype": "File",
			"file_name": fileName,
			"attached_to_doctype": parentDoctype,
			"attached_to_name": parentName,
			"is_private": 1,
			"content": decoded,
		})
		fileDoc.insert(ignore_permissions=True)
		created.append({"file_name": fileDoc.file_name, "file_url": fileDoc.file_url})

	return created



def _BuildSummary(isTask: bool, savedEmail: bool, createdFiles: list[dict], targetDoctype: str, targetName: str, hasComment: bool) -> str:
	"""Build a human-readable confirmation message."""
	parts = []
	if savedEmail:
		parts.append(_("the email"))
	if createdFiles:
		parts.append(_("{0} attachment(s)").format(len(createdFiles)))
	if hasComment and not isTask:
		parts.append(_("a comment"))

	if isTask:
		if parts:
			return _("Created a task for {0} with {1}.").format(targetName, _(" and ").join(parts))
		return _("Created a task for {0}.").format(targetName)

	what = _(" and ").join(parts) if parts else _("nothing")
	return _("Filed {0} to {1} {2}.").format(what, _(targetDoctype), targetName)


def CreateTaskFromEmail(
	assignee: str,
	comment: str,
	email: dict,
	shouldSaveEmail: bool,
) -> tuple[str, str, Optional[str]]:
	"""Create a ToDo for the assignee, filed from the open email.

	The email is saved FIRST so the task can carry it as its reference: the
	"ToDo Assigned" notification renders when the task is inserted and reads the
	source email (subject, sender, body) through ``reference_type`` /
	``reference_name``. The email is then pointed back at the task so it shows
	on the task's own timeline — the same two-way link Frappe's inbox uses when
	an email is assigned to someone.

	Args:
		assignee: User id (email) to assign the task to.
		comment: The user's note, used as the task description when provided.
		email: {subject, sender, recipients, sent_date, body_html} of the email.
		shouldSaveEmail: Whether the user chose to file the email itself.

	Returns:
		("ToDo", task name, communication name or None).
	"""
	if not assignee or not frappe.db.exists("User", assignee):
		frappe.throw(_("Select a person to assign the task to."))

	sourceCommunication = None
	if shouldSaveEmail:
		sourceCommunication = _CreateCommunication(
			"", "", email["subject"], email["sender"], email["recipients"],
			email["sent_date"], email["body_html"],
		)

	taskName = _CreateTask(assignee, email["subject"], comment, sourceCommunication)

	if sourceCommunication:
		communication = frappe.get_doc("Communication", sourceCommunication)
		communication.reference_doctype = "ToDo"
		communication.reference_name = taskName
		communication.save(ignore_permissions=True)

	return ("ToDo", taskName, sourceCommunication)



def AppendAttachmentLinksToCommunication(communicationName: Optional[str], files: list) -> None:
	"""Add the filed-attachment links to an already-saved email's body."""
	if not communicationName or not files:
		return

	content = frappe.db.get_value("Communication", communicationName, "content") or ""
	newContent = _AppendAttachmentLinks(content, files)
	frappe.db.set_value(
		"Communication", communicationName, "content", newContent, update_modified=False,
	)



def _CreateTask(
	assignee: str,
	subject: str,
	comment: str,
	sourceCommunication: Optional[str] = None,
) -> str:
	"""Create a ToDo (task) assigned to the given user.

	Args:
		assignee: User id (email) to assign the task to — already validated.
		subject: Email subject, used as a fallback task description.
		comment: The user's note, used as the task description when provided.
		sourceCommunication: The filed email's Communication name, stored as the
			task's reference so notifications and the task board can show it.

	Returns:
		The new ToDo's name, so the attachments can be filed against it like
		any other record.
	"""
	# Store the description as plain text: the ToDo description shows in plain
	# fields (task title / edit box), so HTML-escaping it would surface literal
	# entities like &apos;. Frappe sanitises this field when rendered as HTML.
	description = comment or subject or _("Email filed from Outlook")
	task = frappe.get_doc({
		"doctype": "ToDo",
		"allocated_to": assignee,
		"assigned_by": frappe.session.user,
		"description": description,
		"priority": "Medium",
		"status": "Open",
		"reference_type": "Communication" if sourceCommunication else None,
		"reference_name": sourceCommunication,
	})
	# Authorization is enforced upstream: the Task/To-Do option is only offered
	# when the user has create permission on ToDo (see GetPickableDoctypes).
	task.insert(ignore_permissions=True)
	return task.name


def _CreateContact(target_doctype: str, name: str, email: str, body_html: str = "") -> tuple[str, str]:
	"""Create a Contact from an email sender, enriched from the signature.

	Only Contact may be created this way; other targets must already exist. An
	existing contact with the same email is reused rather than duplicated. Phone,
	title and company are pulled from the signature via the bench's AI extractor.

	Returns:
		(doctype, name) of the Contact (existing or new).
	"""
	if target_doctype != "Contact":
		frappe.throw(_("Creating a new {0} is not supported.").format(_(target_doctype)))

	emailAddress = (email or "").strip().lower()

	# Reuse an existing contact with this email instead of creating a duplicate.
	if emailAddress:
		existingName = frappe.db.get_value("Contact Email", {"email_id": emailAddress}, "parent")
		if existingName:
			return ("Contact", existingName)

	if not frappe.has_permission("Contact", "create"):
		frappe.throw(_("You do not have permission to create a Contact."), frappe.PermissionError)

	extracted = _ExtractContactFields(body_html, name, emailAddress)

	firstName = extracted.get("first_name") or (name.split(" ", 1)[0] if name else "") \
		or (emailAddress.split("@")[0] if emailAddress else "")
	if not firstName:
		frappe.throw(_("A name or email is required to create a contact."))

	contact = frappe.get_doc({"doctype": "Contact", "first_name": firstName})
	lastName = extracted.get("last_name") or (name.split(" ", 1)[1] if name and " " in name else "")
	if lastName:
		contact.last_name = lastName
	if extracted.get("designation"):
		contact.designation = extracted["designation"]
	if extracted.get("company_name"):
		# company_name is a Link to Organisation; only set it when that org already
		# exists, otherwise the insert fails with a link-validation error.
		orgName = frappe.db.get_value("Organisation", extracted["company_name"], "name")
		if orgName:
			contact.company_name = orgName

	finalEmail = (extracted.get("email") or emailAddress or "").strip().lower()
	if finalEmail:
		contact.append("email_ids", {"email_id": finalEmail, "is_primary": 1})
		contact.email_id = finalEmail  # populate the shown field, not just the child row

	if extracted.get("phone"):
		contact.append("phone_nos", {"phone": extracted["phone"], "is_primary_phone": 1})
		contact.phone = extracted["phone"]
	if extracted.get("mobile"):
		contact.append("phone_nos", {"phone": extracted["mobile"], "is_primary_mobile_no": 1})
		contact.mobile_no = extracted["mobile"]

	contact.insert(ignore_permissions=True)
	return ("Contact", contact.name)


def _ExtractContactFields(body_html: str, name: str, email: str) -> dict[str, str]:
	"""Pull the sender's contact details from the email signature via the bench's
	Claude integration. Returns {} on any failure so contact creation proceeds.
	"""
	try:
		from anthropic import Anthropic
		from claude_agent.claude_agent.model_config import GetActiveModel

		apiKey = frappe.get_single("Agent Settings").get_password("anthropic_api_key", raise_exception=False)
		if not apiKey:
			return {}

		plainBody = frappe.utils.strip_html(body_html or "")[:CONTACT_EXTRACTION_MAX_CHARS]
		if not plainBody.strip():
			return {}

		tool = {
			"name": "contact_details",
			"description": "Record the sender's contact details taken from their email signature.",
			"input_schema": {
				"type": "object",
				"properties": {
					"first_name": {"type": "string"},
					"last_name": {"type": "string"},
					"email": {"type": "string"},
					"phone": {"type": "string", "description": "Landline or office phone"},
					"mobile": {"type": "string", "description": "Mobile / cell number"},
					"designation": {"type": "string", "description": "Job title or position"},
					"company_name": {"type": "string"},
				},
			},
		}
		systemPrompt = (
			"Extract the contact details of the person who SENT this email, from their "
			"signature. Only fill a field when you are confident; otherwise leave it out. "
			"Never invent values."
		)
		prompt = f"Sender name: {name}\nSender email: {email}\n\nEmail:\n{plainBody}"

		response = Anthropic(api_key=apiKey).messages.create(
			model=GetActiveModel(),
			max_tokens=CONTACT_EXTRACTION_MAX_TOKENS,
			system=systemPrompt,
			tools=[tool],
			tool_choice={"type": "tool", "name": "contact_details"},
			messages=[{"role": "user", "content": prompt}],
		)
		for block in response.content:
			if getattr(block, "type", None) == "tool_use":
				return {k: v.strip() for k, v in (block.input or {}).items() if isinstance(v, str) and v.strip()}
		return {}
	except Exception:
		frappe.log_error(frappe.get_traceback(), "Send-to-ERP contact extraction failed")
		return {}


def _AddComment(doctype: str, name: str, text: str) -> None:
	"""Add a timeline comment to the target record (text escaped to HTML)."""
	content = frappe.utils.escape_html(text).replace("\n", "<br>")
	doc = frappe.get_doc(doctype, name)
	doc.add_comment("Comment", content)



def CreateRecordFromSender(
	targetDoctype: str, name: str, email: str, bodyHTML: str,
) -> tuple[str, str]:
	"""Create (or reuse) a Contact or Candidate for the open email's sender.

	Returns:
		(doctype, name) of the record to file the email against.
	"""
	if targetDoctype == "Candidate":
		return CreateCandidate(name, email, bodyHTML)

	return _CreateContact(targetDoctype, name, email, bodyHTML)



def CreateCandidate(name: str, email: str, bodyHTML: str) -> tuple[str, str]:
	"""Create a Candidate from an email sender, enriched from the signature.

	An existing candidate with the same email address is reused rather than
	duplicated. Candidate requires both a first name and a surname, so when
	neither the sender's display name nor the signature gives both, the user is
	asked to create the candidate in ERP instead of getting a half-filled record.

	Returns:
		("Candidate", name) of the existing or new candidate.
	"""
	emailAddress = (email or "").strip().lower()

	if emailAddress:
		existingName = frappe.db.get_value("Candidate", {"email_address": emailAddress}, "name")
		if existingName:
			return ("Candidate", existingName)

	if not frappe.has_permission("Candidate", "create"):
		frappe.throw(_("You do not have permission to create a Candidate."), frappe.PermissionError)

	extracted = _ExtractContactFields(bodyHTML, name, emailAddress)
	firstName, lastName = SplitSenderName(name, extracted)
	if not firstName or not lastName:
		who = name or emailAddress or _("this sender")
		message = _(
			"Couldn't work out both a first name and a surname for {0}. "
			"Please create this candidate in ERP."
		)
		frappe.throw(message.format(who))

	candidate = frappe.get_doc({
		"doctype": "Candidate",
		"first_name": firstName,
		"last_name": lastName,
		"email_address": (extracted.get("email") or emailAddress or "").strip().lower(),
	})
	phoneNumber = extracted.get("mobile") or extracted.get("phone")
	if phoneNumber:
		candidate.phone_number = phoneNumber

	# Normal permissions apply: the user was checked for Candidate create above.
	candidate.insert()
	return ("Candidate", candidate.name)



def SplitSenderName(name: str, extracted: dict) -> tuple[str, str]:
	"""Return (first name, surname), preferring the signature over the display name."""
	nameParts = (name or "").strip().split(" ", 1)
	fallbackFirst = nameParts[0] if nameParts else ""
	fallbackLast = nameParts[1].strip() if len(nameParts) > 1 else ""

	firstName = extracted.get("first_name") or fallbackFirst
	lastName = extracted.get("last_name") or fallbackLast
	return (firstName.strip(), lastName.strip())



def GetTicketTypes() -> list[str]:
	"""Return the Support Ticket types, in the order the doctype lists them."""
	options = frappe.get_meta(TICKET_TARGET).get_field("support_type").options or ""
	return [option.strip() for option in options.split("\n") if option.strip()]



def SearchTicketTypes(txt: str) -> list[dict]:
	"""List the ticket types matching the typed text, for the add-in picker."""
	searchText = (txt or "").strip().lower()
	return [
		{"value": ticketType, "label": _(ticketType), "sublabel": ""}
		for ticketType in GetTicketTypes()
		if searchText in ticketType.lower()
	]



def FileEmailAsTicket(
	supportType: str, email: dict, shouldSaveEmail: bool, comment: str, attachmentList: list[dict],
) -> dict:
	"""Raise a Support Ticket from the open email.

	The ticket goes through the support app's own CreateTicket, so on a remote
	(proxy-mode) site it is created on the primary helpdesk, exactly like a
	ticket raised from the helpdesk page. The email's text is copied into the
	ticket description, and the chosen attachments are uploaded to the ticket.

	Args:
		supportType: Ticket type the user picked, e.g. "Issue/Bug".
		email: {subject, sender, recipients, sent_date, body_html} of the email.
		shouldSaveEmail: Whether to copy the email itself into the ticket.
		comment: The user's note, placed at the top of the description.
		attachmentList: Parsed attachments to upload to the ticket.

	Returns:
		{success, communication, files, message} like FileEmailToRecord.
	"""
	from support.support.page.helpdesk.data.ticket import CreateTicket

	if not frappe.has_permission(TICKET_TARGET, "create"):
		message = _("You do not have permission to raise a support ticket.")
		frappe.throw(message, frappe.PermissionError)
	if supportType not in GetTicketTypes():
		frappe.throw(_("Choose a ticket type."))
	if not shouldSaveEmail and not comment:
		frappe.throw(_("Add a comment, or include the email, so the ticket says what is needed."))

	user = frappe.session.user
	ticket = CreateTicket(
		subject=email["subject"] or _("(no subject)"),
		description=BuildTicketDescription(email, comment, shouldSaveEmail),
		supportType=supportType,
		priorityLevel=TICKET_PRIORITY,
		raised_by_user=user,
		raised_by_email=frappe.db.get_value("User", user, "email") or user,
	)
	ticketName = (ticket or {}).get("name")
	if not ticketName:
		frappe.throw(_("The support ticket could not be created."))

	canAttach = CanAttachToTicket(ticketName)
	attachedFiles = AttachFilesToTicket(ticketName, attachmentList) if canAttach else []
	frappe.db.commit()

	return {
		"success": True,
		"communication": None,
		"files": attachedFiles,
		"message": BuildTicketSummary(ticketName, attachmentList, attachedFiles, canAttach),
	}



def CanAttachToTicket(ticketName: str) -> bool:
	"""True when the user may add files to the ticket, under the helpdesk's own rule.

	The helpdesk's upload refuses anyone without write on the ticket, and the
	"All" role can create and read its own tickets but not write them. The
	add-in applies the same rule rather than slipping files in on its behalf.
	On a proxy-mode site the ticket lives on the primary, so the check is the
	user's write permission on Support Tickets in general.
	"""
	from support.support.utils.remote import IsProxyMode

	if IsProxyMode():
		return bool(frappe.has_permission(TICKET_TARGET, "write"))
	return bool(frappe.has_permission(TICKET_TARGET, "write", ticketName))



def BuildTicketDescription(email: dict, comment: str, shouldSaveEmail: bool) -> str:
	"""Build the ticket description HTML: who raised it, their note, then the email.

	The email body is reduced to plain text (Outlook HTML is heavy and full of
	styling) and every value is escaped before it goes into the HTML.
	"""
	escape = frappe.utils.escape_html
	fullName = frappe.utils.get_fullname(frappe.session.user)
	parts = ["<p>{0}</p>".format(escape(_("Raised from Outlook by {0}.").format(fullName)))]

	if comment:
		parts.append("<p>{0}</p>".format(escape(comment).replace("\n", "<br>")))

	if not shouldSaveEmail:
		return "".join(parts)

	from frappe.core.utils import html2text

	sentOn = frappe.utils.format_datetime(_ParseEmailDate(email["sent_date"]))
	headerLines = [
		"<strong>{0}</strong>".format(escape(_("Original email"))),
		escape(_("From: {0}").format(email["sender"])),
		escape(_("To: {0}").format(email["recipients"])),
		escape(_("Sent: {0}").format(sentOn)),
		escape(_("Subject: {0}").format(email["subject"])),
	]
	bodyText = html2text(email["body_html"] or "")[:TICKET_BODY_MAX_CHARS]
	parts.append("<hr><p>{0}</p>".format("<br>".join(headerLines)))
	parts.append("<p>{0}</p>".format(escape(bodyText).replace("\n", "<br>")))
	return "".join(parts)



def AttachFilesToTicket(ticketName: str, attachmentList: list[dict]) -> list[dict]:
	"""Attach files to the ticket, forwarding them when the ticket lives on the primary site.

	A failed upload is logged and skipped, not raised: the ticket already exists
	by then, and the summary tells the user which files did not make it.

	Returns:
		List of {file_name, file_url} that were attached.
	"""
	from support.support.utils.remote import GetRemoteConfig, IsProxyMode

	if not attachmentList:
		return []
	if not IsProxyMode():
		return _AttachFiles(attachmentList, TICKET_TARGET, ticketName)

	config = GetRemoteConfig()
	attached = []
	for item in attachmentList:
		fileName = item.get("file_name") or "attachment"
		content = item.get("content_base64")
		if not content:
			continue
		if UploadFileToPrimaryTicket(config, ticketName, fileName, base64.b64decode(content)):
			attached.append({"file_name": fileName, "file_url": ""})

	return attached



def UploadFileToPrimaryTicket(config: dict, ticketName: str, fileName: str, content: bytes) -> bool:
	"""Upload one file to a ticket on the primary helpdesk site. Returns True on success."""
	import requests

	url = "https://{0}/api/method/{1}".format(config["remoteSiteUrl"], TICKET_UPLOAD_METHOD)
	headers = {"Authorization": "token {0}:{1}".format(config["apiKey"], config["apiSecret"])}
	try:
		response = requests.post(
			url,
			headers=headers,
			files={"file": (fileName, content)},
			data={"ticketName": ticketName},
			timeout=TICKET_UPLOAD_TIMEOUT_SECONDS,
		)
	except requests.RequestException as uploadError:
		# Log the error class only: some requests errors (e.g. InvalidHeader) quote
		# the Authorization header, which carries the support API key and secret.
		message = "Upload of {0} to {1} failed: {2}".format(fileName, ticketName, type(uploadError).__name__)
		frappe.log_error(message, "Send-to-ERP ticket attachment upload failed")
		return False

	if response.ok:
		return True

	# Log the status only: the response body may echo request details.
	frappe.log_error(
		"Upload of {0} to {1} returned HTTP {2}".format(fileName, ticketName, response.status_code),
		"Send-to-ERP ticket attachment upload failed",
	)
	return False



def BuildTicketSummary(
	ticketName: str, attachmentList: list[dict], attachedFiles: list[dict], canAttach: bool,
) -> str:
	"""Build the confirmation shown in the panel after a ticket is raised."""
	wantedCount = len([item for item in attachmentList if item.get("content_base64")])
	attachedCount = len(attachedFiles)

	message = _("Created support ticket {0}.").format(ticketName)
	if wantedCount and not canAttach:
		notAllowedText = _("Attachments were not added: only support staff can attach files to tickets.")
		return message + " " + notAllowedText
	if attachedCount:
		message += " " + _("{0} attachment(s) added.").format(attachedCount)
	if attachedCount < wantedCount:
		missedCount = wantedCount - attachedCount
		missedText = _("{0} attachment(s) could not be added; please add them on the ticket.")
		message += " " + missedText.format(missedCount)
	return message
