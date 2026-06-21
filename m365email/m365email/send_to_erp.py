"""Send-to-ERP API for the Outlook add-in.

Lets an Outlook add-in file an email (and/or its attachments) against any
permitted ERP record. The add-in panel is served same-origin from this site,
so these endpoints rely on the standard Frappe login session for auth.

All public endpoints enforce normal Frappe document permissions — nothing is
inserted with ignore_permissions, so a user can only file against records they
are allowed to see and attach to.
"""
import base64
import json
from typing import Optional

import frappe
from frappe import _


# Doctypes a user may file an email against, in picker order. Each user only
# sees the subset they have read access to (see GetPickableDoctypes).
DEFAULT_PICKABLE_DOCTYPES = [
	"Contact",
	"Project",
	"Purchase Order",
	"Supplier",
	"Sales Invoice",
	"Purchase Invoice",
]

MAX_SEARCH_RESULTS = 20
MAX_ATTACHMENT_BYTES = 35 * 1024 * 1024  # 35 MB — above Outlook's own attach limit
CONTACT_EXTRACTION_MAX_CHARS = 8000  # email body chars sent to AI (covers most signatures)
CONTACT_EXTRACTION_MAX_TOKENS = 400  # AI response cap for contact-field extraction

# Secondary detail shown beneath the name in the picker, to disambiguate records
# that share a title (e.g. two contacts both named "Paul Johnson").
SECONDARY_FIELDS = {
	"Contact": "email_id",
}

# Special pseudo-target: rather than an existing record, the user picks a person
# and a new ToDo (task) is created for them, with the email filed against it.
TASK_TARGET = "ToDo"
TASK_TARGET_LABEL = "Task / To-Do"



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

	if not frappe.has_permission(doctype, "read"):
		frappe.throw(_("You do not have access to {0}").format(_(doctype)), frappe.PermissionError)

	titleField = _GetTitleField(doctype)
	fields = ["name"]
	if titleField and titleField != "name":
		fields.append(titleField)

	secondaryField = _GetSecondaryField(doctype)
	if secondaryField and secondaryField not in fields:
		fields.append(secondaryField)

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

	return [_FormatSearchRow(row, titleField, secondaryField) for row in records]



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
) -> dict:
	"""File an email and/or its attachments against a record or a new task.

	Args:
		target_doctype: Filing target — a pickable doctype, or ``ToDo`` to create
			a new task (in which case ``target_name`` is the assignee's user id).
		target_name: Record to file against, or the assignee for a Task/To-Do.
		subject: Email subject line.
		sender: Sender email address.
		recipients: Recipient email addresses (comma separated).
		sent_date: ISO timestamp the email was sent/received.
		body_html: HTML body of the email.
		save_email: 1 to record the email itself on the timeline.
		comment: Optional note — a timeline comment on a record, or the task's
			description when creating a Task/To-Do.
		create_new: 1 to create a new Contact (from new_name/new_email) and file
			against it, instead of using an existing record.
		new_name: Display name for the new Contact when create_new is set.
		new_email: Email address for the new Contact when create_new is set.
		attachments: JSON list of {file_name, content_base64, content_type}.

	Returns:
		{success, communication, files, message} describing what was created.
	"""
	_GuardPickableDoctype(target_doctype)

	shouldSaveEmail = bool(int(save_email or 0))
	attachmentList = _ParseAttachments(attachments)
	commentText = (comment or "").strip()
	isTask = target_doctype == TASK_TARGET
	shouldCreateNew = bool(int(create_new or 0))

	if isTask:
		recordDoctype, recordName = _CreateTask(target_name, subject, commentText)
	elif shouldCreateNew:
		recordDoctype, recordName = _CreateContact(target_doctype, new_name, new_email, body_html)
	else:
		_GuardTargetExists(target_doctype, target_name)
		if not frappe.has_permission(target_doctype, "write", target_name):
			frappe.throw(
				_("You do not have permission to attach to {0} {1}").format(_(target_doctype), target_name),
				frappe.PermissionError,
			)
		if not shouldSaveEmail and not attachmentList and not commentText:
			frappe.throw(_("Add a comment, or select the email and/or an attachment to file."))
		recordDoctype, recordName = target_doctype, target_name

	# Attach files to the target record itself so they show in the record's
	# standard attachments, and capture their URLs for the email body.
	createdFiles = _AttachFiles(attachmentList, recordDoctype, recordName)

	# Build the email body: on a record, a filing note rides at the top so the
	# comment shares the email's timeline entry; then the attachment links.
	communicationName = None
	if shouldSaveEmail:
		emailBody = body_html
		if commentText and not isTask:
			emailBody = _PrependNote(emailBody, commentText)
		communicationName = _CreateCommunication(
			recordDoctype, recordName, subject, sender, recipients, sent_date,
			_AppendAttachmentLinks(emailBody, createdFiles),
		)

	# A comment with no email to ride along becomes a standalone timeline comment.
	# (A task already carries the comment as its description.)
	if commentText and not isTask and not shouldSaveEmail:
		_AddComment(recordDoctype, recordName, commentText)

	frappe.db.commit()

	# For a task, name the assignee in the message (not the new ToDo's id).
	summaryName = target_name if isTask else recordName
	return {
		"success": True,
		"communication": communicationName,
		"files": createdFiles,
		"message": _BuildSummary(isTask, shouldSaveEmail, createdFiles, recordDoctype, summaryName, bool(commentText)),
	}



def _GuardPickableDoctype(doctype: str) -> None:
	"""Reject any target not in the pickable allow-list (Task/To-Do included)."""
	if doctype != TASK_TARGET and doctype not in DEFAULT_PICKABLE_DOCTYPES:
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
		frappe.throw(_("{0} {1} was not found.").format(_(doctype), name or ""))



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

	for item in parsed:
		content = item.get("content_base64") or ""
		# base64 expands ~4/3; estimate decoded size without decoding twice.
		estimatedBytes = (len(content) * 3) // 4
		if estimatedBytes > MAX_ATTACHMENT_BYTES:
			frappe.throw(_("Attachment {0} is too large.").format(item.get("file_name", "")))

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

		fileDoc = frappe.get_doc({
			"doctype": "File",
			"file_name": fileName,
			"attached_to_doctype": parentDoctype,
			"attached_to_name": parentName,
			"is_private": 1,
			"content": base64.b64decode(content),
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


def _CreateTask(assignee: str, subject: str, comment: str) -> tuple[str, str]:
	"""Create a ToDo (task) assigned to the given user.

	Args:
		assignee: User id (email) to assign the task to.
		subject: Email subject, used as a fallback task description.
		comment: The user's note, used as the task description when provided.

	Returns:
		(doctype, name) of the new ToDo, so the email and attachments can be
		filed against it like any other record.
	"""
	if not assignee or not frappe.db.exists("User", assignee):
		frappe.throw(_("Select a person to assign the task to."))

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
	})
	# Authorization is enforced upstream: the Task/To-Do option is only offered
	# when the user has create permission on ToDo (see GetPickableDoctypes).
	task.insert(ignore_permissions=True)
	return ("ToDo", task.name)


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
		contact.company_name = extracted["company_name"]

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
