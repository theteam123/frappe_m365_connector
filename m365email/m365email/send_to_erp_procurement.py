"""Send to ERP: procurement actions for the Outlook add-in.

Paul, 2026-11-06. Lets procurement turn supplier emails into ERP records in one or two clicks:

- Create new Supplier Quote: from a supplier's quote email, a Draft quote for the job they pick,
  with the sender as the supplier contact, the supplier worked out for them, and the PDF attached.
- Purchase Order suggestion: reads the PO number from the email (subject, then body, then the
  PDFs) and preselects that order.
- Enter Supplier Invoice: from a supplier's invoice email, a Draft Purchase Invoice against that
  order carrying its not-yet-invoiced lines, with the invoice PDF and the email attached.

The business rules live in the erp app (po_number_detect, purchase_invoice.from_email,
supplier_quote.from_email). This module only reads the email and wires those to the add-in, so
it is imported only on a site that has erp installed.

Everything here handles untrusted input: the subject, body and attachments come from whoever
sent the email. Text is only ever matched with fixed patterns and looked up as bound filter
values; PDFs are read by pdfplumber with a page cap and a size cap; nothing from the email is
executed, rendered unescaped, or used to pick a doctype. All records are created with the
user's own permissions.
"""

import base64
from typing import Optional

import frappe
from frappe import _

from m365email.m365email import send_to_erp as filing

SUPPLIER_QUOTE = "Supplier Quote"
PURCHASE_ORDER = "Purchase Order"
PURCHASE_INVOICE = "Purchase Invoice"
ORGANISATION = "Organisation"
CONTACT = "Contact"
USER = "User"

# NOTE: Detection reads at most this much of each PDF. Larger files are still attached, they
#       are just not read, so a huge scan cannot tie up a web worker.
MAX_DETECTION_PDF_BYTES = 10 * 1024 * 1024
MAX_DETECTION_PDFS = 5
MAX_EMAIL_TEXT_CHARS = 50000
MAX_SUBJECT_CHARS = 1000
PDF_SIGNATURE = b"%PDF"
PDF_EXTENSION = ".pdf"
INVOICE_ATTACH_FIELD = "purchase_invoice_attach"
# NOTE: base64 turns every 3 bytes into 4 characters, so the encoded length gives the size.
BASE64_RATIO = 4 / 3
MAX_SUPPLIER_RESULTS = 20

GUARD_ACTION_INVOICE = "enter-supplier-invoice"



@frappe.whitelist()
def SuggestTarget(
	target_doctype: str = "",
	subject: str = "",
	body_html: str = "",
	attachments: str = "[]",
) -> dict:
	"""Suggest the record an email is about, for the add-in to preselect.

	Only Purchase Order is supported: the PO number is read from the subject, then the body,
	then any PDF attachments, and the supplier's invoice number and dates are read alongside.

	Args:
		target_doctype: The type picked in the add-in.
		subject: Email subject.
		body_html: Email body.
		attachments: JSON list of {file_name, content_base64}; only PDFs are read.

	Returns:
		{match, choices, invoice}: match is the one order to preselect (or None), choices the
		orders to offer when several are named, invoice the {supplier_invoice_number,
		invoice_date, due_date} read from the email ("" where not found).
	"""
	empty = {"match": None, "choices": [], "invoice": {}}
	if target_doctype != PURCHASE_ORDER or not IsErpInstalled():
		return empty
	if not frappe.has_permission(PURCHASE_ORDER, "read"):
		return empty

	from erp.accounting.doctype.purchase_invoice.from_email import ReadInvoiceDetails
	from erp.accounting.doctype.purchase_order.po_number_detect import SuggestPurchaseOrder

	sources = BuildEmailTextSources(subject, body_html, filing._ParseAttachments(attachments))
	suggestion = SuggestPurchaseOrder(sources)
	suggestion["invoice"] = ReadInvoiceDetails(sources)
	return suggestion



@frappe.whitelist()
def GetSenderSupplier(sender_email: str = "") -> dict:
	"""The supplier a quote from this sender would be filed under, for the add-in to show.

	Returns:
		{"supplier": {value, label} or None, "has_contact": bool}. The supplier comes from
		the sender's Contact; when it is None the add-in asks the person to pick one, or to
		leave it for the PDF to decide.
	"""
	if not IsErpInstalled() or not frappe.has_permission(SUPPLIER_QUOTE, "create"):
		return {"supplier": None, "has_contact": False}
	if not frappe.has_permission(ORGANISATION, "read"):
		return {"supplier": None, "has_contact": False}

	from erp.procurement.doctype.supplier_quote.from_email import SupplierFromContact

	contactName = FindSenderContact(sender_email)
	detection = SupplierFromContact(contactName)
	if detection and not frappe.has_permission(ORGANISATION, "read", detection["organisation"]):
		detection = None
	return {"supplier": FormatSupplier(detection), "has_contact": bool(contactName)}



@frappe.whitelist()
def SearchSuppliers(txt: str = "") -> list[dict]:
	"""Search supplier Organisations by name, for the add-in's supplier picker."""
	if not frappe.has_permission(ORGANISATION, "read"):
		frappe.throw(_("You do not have access to suppliers."), frappe.PermissionError)

	searchText = (txt or "").strip()
	filters = {"supplier": 1}
	if searchText:
		filters["organisation_name"] = ["like", f"%{searchText}%"]

	rows = frappe.get_list(
		ORGANISATION,
		filters=filters,
		fields=["name", "organisation_name"],
		limit_page_length=MAX_SUPPLIER_RESULTS,
		order_by="organisation_name asc",
	)
	return [FormatSupplierRow(row) for row in rows]



def FormatSupplierRow(row: dict) -> dict:
	return {"value": row.name, "label": row.organisation_name or "", "sublabel": ""}



def CreateSupplierQuoteFromEmail(
	project: str,
	supplier: str,
	sender: dict,
	attachmentList: list[dict],
) -> tuple[str, str, str]:
	"""Create the Draft Supplier Quote for "Create new" on a quote email.

	The supplier is worked out before anything is written, so a quote whose supplier cannot be
	told leaves nothing behind (not even a new Contact).

	Args:
		project: The job the user picked.
		supplier: A supplier the user picked, or "" to work it out.
		sender: {name, email, body_html, sent_date} of the email.
		attachmentList: The parsed attachments; PDFs among them are read for the supplier.

	Returns:
		("Supplier Quote", quote name, a sentence saying how the supplier was found).
	"""
	from erp.procurement.doctype.supplier_quote.from_email import (
		CreateDraftQuoteFromEmail,
		LinkContactToSupplier,
		ResolveQuoteSupplier,
	)

	if not frappe.has_permission(SUPPLIER_QUOTE, "create"):
		message = _("You do not have permission to create a Supplier Quote.")
		frappe.throw(message, frappe.PermissionError)

	contactName = FindSenderContact(sender["email"])
	documents = DecodeDetectionPDFs(attachmentList)
	detection = ResolveQuoteSupplier(supplier or None, contactName, documents)
	if not detection:
		message = _("Couldn't tell which supplier this quote is from. Pick the supplier.")
		frappe.throw(message)

	# NOTE: A quote forwarded by a colleague comes "from" them. Our own staff are never the
	#       supplier's contact, so no contact is created for them (the supplier still comes
	#       from a pick or the PDF), and FindSenderContact never returns theirs.
	isInternal = IsInternalSender(sender["email"])
	if not contactName and not isInternal and frappe.has_permission(CONTACT, "create"):
		senderName, senderEmail, bodyHTML = sender["name"], sender["email"], sender["body_html"]
		newContact = filing._CreateContact(CONTACT, senderName, senderEmail, bodyHTML)
		contactName = newContact[1]

	organisation = detection["organisation"]
	quoteDate = filing._ParseEmailDate(sender["sent_date"]).date()
	quote = CreateDraftQuoteFromEmail(project, organisation, contactName, quoteDate)
	LinkContactToSupplier(contactName, organisation)

	return (SUPPLIER_QUOTE, quote.name, BuildSupplierNote(detection))



def BuildSupplierNote(detection: dict) -> str:
	"""Say which supplier was used and how; the name only to someone who may read it."""
	source = _(detection["matched_on"])
	if not frappe.has_permission(ORGANISATION, "read", detection["organisation"]):
		return _("Supplier set from the {0}.").format(source)
	return _("Supplier {0} (from the {1}).").format(detection["organisation_name"], source)



@frappe.whitelist()
def EnterSupplierInvoice(
	purchase_order: str,
	supplier_invoice_number: str = "",
	invoice_date: str = "",
	due_date: str = "",
	subject: str = "",
	sender: str = "",
	recipients: str = "",
	sent_date: Optional[str] = None,
	body_html: str = "",
	save_email: int = 1,
	comment: str = "",
	attachments: str = "[]",
	message_id: str = "",
) -> dict:
	"""Raise a Draft Purchase Invoice against a PO from a supplier's invoice email.

	The invoice carries the order's not-yet-invoiced lines; the chosen attachments are attached
	to it (the first PDF also fills its Invoice attachment field) and the email is filed on its
	timeline. A second click on the same email returns the invoice already raised instead of
	making another.

	Returns:
		{success, message, record_url, files} like FileEmailToRecord.
	"""
	if not IsErpInstalled():
		frappe.throw(_("Supplier invoices are not available on this site."))

	attachmentList = filing._ParseAttachments(attachments)
	email = {
		"subject": subject, "sender": sender, "recipients": recipients,
		"sent_date": sent_date, "body_html": body_html,
	}
	shouldSaveEmail = bool(int(save_email or 0))
	invoiceFields = (purchase_order, supplier_invoice_number, invoice_date, due_date)

	commentText = (comment or "").strip()

	def createInvoice() -> dict:
		emailParts = (email, shouldSaveEmail, commentText, attachmentList)
		return RecordSupplierInvoice(invoiceFields, *emailParts)

	# NOTE: The order is part of the key: a second try against a different order is a
	#       correction, not a repeat, and must not be answered with the first invoice.
	guardAction = f"{GUARD_ACTION_INVOICE}-{purchase_order}"
	return filing.RunOncePerEmail(guardAction, message_id, createInvoice)



def RecordSupplierInvoice(
	invoiceFields: tuple,
	email: dict,
	shouldSaveEmail: bool,
	comment: str,
	attachmentList: list[dict],
) -> dict:
	"""Create the invoice, attach the files and file the email. Commits on success."""
	from erp.accounting.doctype.purchase_invoice.from_email import CreateDraftInvoiceFromOrder
	from erp.accounting.doctype.purchase_invoice.purchase_invoice import ROW_TYPE_LINE_ITEM

	invoice = CreateDraftInvoiceFromOrder(*invoiceFields)
	# NOTE: Files and the email are added with ignore_permissions, so the user must be allowed to
	#       edit the invoice they just made. A refusal rolls the invoice back with the request.
	if not frappe.has_permission(PURCHASE_INVOICE, "write", invoice.name):
		message = _("You can create supplier invoices but not add files to them.")
		frappe.throw(message, frappe.PermissionError)
	createdFiles = filing._AttachFiles(attachmentList, PURCHASE_INVOICE, invoice.name)
	SetInvoiceAttachment(invoice.name, createdFiles)

	if shouldSaveEmail:
		emailBody = email["body_html"]
		if comment:
			emailBody = filing._PrependNote(emailBody, comment)
		filing._CreateCommunication(
			PURCHASE_INVOICE, invoice.name, email["subject"], email["sender"], email["recipients"],
			email["sent_date"], filing._AppendAttachmentLinks(emailBody, createdFiles),
		)
	elif comment:
		filing._AddComment(PURCHASE_INVOICE, invoice.name, comment)

	frappe.db.commit()

	lineCount = len([row for row in invoice.items if row.row_type == ROW_TYPE_LINE_ITEM])
	message = _("Created draft supplier invoice {0} with {1} line(s) from the purchase order.")
	return {
		"success": True,
		"communication": None,
		"files": createdFiles,
		"record": {"doctype": PURCHASE_INVOICE, "name": invoice.name},
		"record_url": frappe.utils.get_url_to_form(PURCHASE_INVOICE, invoice.name),
		"message": message.format(invoice.supplier_invoice_number, lineCount),
	}



def SetInvoiceAttachment(invoiceName: str, createdFiles: list[dict]) -> None:
	"""Point the invoice's Invoice attachment field at the first PDF filed with it."""
	pdfFiles = [file for file in createdFiles if file["file_name"].lower().endswith(PDF_EXTENSION)]
	if not pdfFiles:
		return

	fileURL = pdfFiles[0]["file_url"]
	frappe.db.set_value(
		PURCHASE_INVOICE, invoiceName, INVOICE_ATTACH_FIELD, fileURL, update_modified=False
	)



def BuildEmailTextSources(subject: str, bodyHTML: str, attachmentList: list[dict]) -> list[str]:
	"""The email's texts in the order evidence is trusted: subject, body, then each PDF."""
	from frappe.core.utils import html2text

	subjectText = (subject or "")[:MAX_SUBJECT_CHARS]
	bodyText = html2text(bodyHTML or "")[:MAX_EMAIL_TEXT_CHARS]
	pdfTexts = [ReadPDFText(content) for content in DecodeDetectionPDFs(attachmentList)]
	return [subjectText, bodyText] + pdfTexts



def ReadPDFText(content: bytes) -> str:
	from erp.procurement.doctype.supplier_quote.extractors import pdf_table_reader

	return pdf_table_reader.ExtractDocumentText(content)[:MAX_EMAIL_TEXT_CHARS]



def DecodeDetectionPDFs(attachmentList: list[dict]) -> list[bytes]:
	"""Decode the attachments worth reading: real PDFs, under the size cap, the first few only."""
	documents = []
	for item in attachmentList:
		if len(documents) >= MAX_DETECTION_PDFS:
			break
		content = DecodeAttachment(item)
		if not content or len(content) > MAX_DETECTION_PDF_BYTES:
			continue
		if not content.startswith(PDF_SIGNATURE):
			continue
		documents.append(content)
	return documents



def DecodeAttachment(item: dict) -> bytes:
	"""Decode one attachment's base64 content; b"" when it is missing, malformed or too big."""
	encoded = item.get("content_base64") if isinstance(item, dict) else None
	if not encoded or not isinstance(encoded, str):
		return b""
	# NOTE: Skip an oversized file before decoding it, so it never costs the memory.
	if len(encoded) > MAX_DETECTION_PDF_BYTES * BASE64_RATIO:
		return b""
	try:
		return base64.b64decode(encoded, validate=False)
	except (ValueError, TypeError):
		return b""



def FindSenderContact(senderEmail: str) -> Optional[str]:
	"""The existing Contact holding the sender's email address, if the user may read it."""
	emailAddress = (senderEmail or "").strip().lower()
	if not emailAddress:
		return None

	if IsInternalSender(emailAddress):
		return None

	contactName = frappe.db.get_value("Contact Email", {"email_id": emailAddress}, "parent")
	if not contactName or not frappe.has_permission(CONTACT, "read", contactName):
		return None
	return contactName



def IsInternalSender(senderEmail: str) -> bool:
	"""True when the address belongs to one of our own ERP users (a colleague, not a supplier).

	Paul, 2026-11-07: procurement often forwards a supplier's quote internally, so the
	email's sender is then a colleague. Treating them as the supplier contact would file the
	quote under the wrong person and could write the supplier onto a staff contact.
	"""
	emailAddress = (senderEmail or "").strip().lower()
	if not emailAddress:
		return False
	if frappe.db.exists(USER, emailAddress):
		return True
	return bool(frappe.db.exists(USER, {"email": emailAddress}))



def FormatSupplier(detection: Optional[dict]) -> Optional[dict]:
	if not detection:
		return None
	return {"value": detection["organisation"], "label": detection["organisation_name"]}



def IsErpInstalled() -> bool:
	return "erp" in frappe.get_installed_apps()
