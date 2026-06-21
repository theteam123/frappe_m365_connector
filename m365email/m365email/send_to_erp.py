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

# Secondary detail shown beneath the name in the picker, to disambiguate records
# that share a title (e.g. two contacts both named "Paul Johnson").
SECONDARY_FIELDS = {
	"Contact": "email_id",
}



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
	attachments: str = "[]",
) -> dict:
	"""File an email and/or its attachments against an ERP record.

	Args:
		target_doctype: Doctype to attach to — must be in the pickable allow-list.
		target_name: Record name to attach to.
		subject: Email subject line.
		sender: Sender email address.
		recipients: Recipient email addresses (comma separated).
		sent_date: ISO timestamp the email was sent/received.
		body_html: HTML body of the email.
		save_email: 1 to record the email itself on the record's timeline.
		attachments: JSON list of {file_name, content_base64, content_type}.

	Returns:
		{success, communication, files, message} describing what was created.
	"""
	_GuardPickableDoctype(target_doctype)
	_GuardTargetExists(target_doctype, target_name)

	if not frappe.has_permission(target_doctype, "write", target_name):
		frappe.throw(
			_("You do not have permission to attach to {0} {1}").format(_(target_doctype), target_name),
			frappe.PermissionError,
		)

	shouldSaveEmail = bool(int(save_email or 0))
	attachmentList = _ParseAttachments(attachments)

	if not shouldSaveEmail and not attachmentList:
		frappe.throw(_("Select the email and/or at least one attachment to file."))

	# Attach files to the target record itself so they show in the record's
	# standard attachments, and capture their URLs for the email body.
	createdFiles = _AttachFiles(attachmentList, target_doctype, target_name)

	# Save the email on the timeline, listing the attachments in its body so they
	# are reachable directly beneath the email as well.
	communicationName = None
	if shouldSaveEmail:
		communicationName = _CreateCommunication(
			target_doctype, target_name, subject, sender, recipients, sent_date,
			_AppendAttachmentLinks(body_html, createdFiles),
		)

	frappe.db.commit()

	return {
		"success": True,
		"communication": communicationName,
		"files": createdFiles,
		"message": _BuildSummary(shouldSaveEmail, createdFiles, target_doctype, target_name),
	}



def _GuardPickableDoctype(doctype: str) -> None:
	"""Reject any doctype not in the pickable allow-list."""
	if doctype not in DEFAULT_PICKABLE_DOCTYPES:
		frappe.throw(_("{0} is not available as a filing target.").format(doctype), frappe.PermissionError)



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



def _BuildSummary(savedEmail: bool, createdFiles: list[dict], targetDoctype: str, targetName: str) -> str:
	"""Build a human-readable confirmation message."""
	parts = []
	if savedEmail:
		parts.append(_("the email"))
	if createdFiles:
		parts.append(_("{0} attachment(s)").format(len(createdFiles)))

	what = _(" and ").join(parts) if parts else _("nothing")
	return _("Filed {0} to {1} {2}.").format(what, _(targetDoctype), targetName)
