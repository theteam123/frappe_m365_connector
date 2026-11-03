# Copyright (c) 2026, TierneyMorris Pty Ltd and contributors
# For license information, please see license.txt

"""
Prepare an Outlook DRAFT with an HTML body and inline pictures. Nothing here ever sends.

Why this exists: the hosted Outlook connector Claude sessions use rejects <img> tags and
has no attachment option, so any email that needs a screenshot came out without it.

Two routes, same inputs:

- CreateDraft writes the draft straight into the person's Drafts folder through Microsoft
  Graph (POST /users/{mailbox}/messages, which only ever creates a draft). The pictures go
  in as inline file attachments whose contentId matches a cid: reference in the HTML.
  Needs the app registration to hold the Mail.ReadWrite application permission.
- WriteDraftEML builds the same message as a .eml file marked X-Unsent, for when Graph is
  not available. The person opens it, checks it and presses Send themselves.

The HTML refers to each picture as <img src="cid:NAME">, and images maps NAME to a file
path. Both are checked against each other, so a missing or unused picture is an error
rather than a broken email.

NOTE: Deliberately not whitelisted. It runs from the command line (see RunFromCommandLine)
or the console only, by someone who already has shell access to the bench. The checks
below therefore guard against mistakes (a wrong path, a stray file, an unknown mailbox),
not against an attacker: every picture must decode as a real PNG, JPEG or GIF, so a
mistyped path cannot put a config file into someone's mailbox.
"""

import argparse
import base64
import io
import json
import os
import re
from email.message import EmailMessage
from email.policy import SMTP
from email.utils import formatdate
from urllib.parse import quote

import frappe
from frappe import _
from frappe.utils import validate_email_address
from PIL import Image, ImageSequence

from m365email.m365email.auth import get_access_token
from m365email.m365email.graph_api import make_graph_request

EMAIL_ACCOUNT_DOCTYPE = "Email Account"
M365_SERVICE = "M365"
DRAFT_ENDPOINT = "/users/{mailbox}/messages"
DRAFT_RESULT_FIELDS = ("id", "webLink", "isDraft", "subject")
REQUIRED_GRAPH_ROLE = "Mail.ReadWrite"

# NOTE: Graph refuses a create-message request over 4 MB. Base64 grows the pictures by a
# third, so 2.9 MB of pictures leaves room for the body and the JSON around it; the whole
# request is measured as well before it goes. Larger needs Graph's upload sessions, which
# this tool does not do.
BYTES_PER_MEGABYTE = 1000 * 1000
GRAPH_MAX_REQUEST_BYTES = 4 * BYTES_PER_MEGABYTE
MAX_TOTAL_IMAGE_BYTES = 2_900_000

IMAGE_MIME_TYPES = {"PNG": "image/png", "JPEG": "image/jpeg", "GIF": "image/gif"}
IMAGE_DECODE_ERRORS = (OSError, SyntaxError, ValueError, EOFError, Image.DecompressionBombError)

# NOTE: A few kilobytes of PNG can claim a picture of a hundred million pixels, which then
# takes hundreds of megabytes to decode. 40 megapixels is far beyond any screenshot or
# photo worth emailing, so the size is checked from the header before anything decodes.
MAX_IMAGE_PIXELS = 40_000_000
HTML_EXTENSIONS = (".html", ".htm")
EML_EXTENSION = ".eml"
CONTENT_ID_PATTERN = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._-]{0,63}$")
CID_REFERENCE_PATTERN = re.compile(r"""cid:([^"'\s>)]+)""", re.IGNORECASE)

# NOTE: A token is three base64url segments. One of any other shape is refused before use,
# because requests puts the whole Authorization header into the error it raises for a
# malformed value, and that error would then land in the Error Log.
TOKEN_PATTERN = re.compile(r"[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+\.[A-Za-z0-9_-]*")
JWT_CLAIMS_INDEX = 1
BASE64_BLOCK = 4
JSON_INDENT = 1



def CreateDraft(
	mailbox: str,
	subject: str,
	htmlBody: str,
	images: dict | None = None,
	to: list | None = None,
	cc: list | None = None,
) -> dict:
	"""Create a draft in the mailbox's Drafts folder through Graph. Never sends.

	Args:
		mailbox: Address whose Drafts folder receives the message; must have an M365
			Email Account on this site
		subject: Subject line
		htmlBody: HTML body; pictures referenced as <img src="cid:NAME">
		images: {NAME: file path} for every cid: in the body
		to: Recipient addresses (optional; the person can add them before sending)
		cc: Copy addresses (optional)

	Returns:
		dict: mailbox, id, webLink (opens the draft in Outlook on the web), isDraft,
		subject, imageCount
	"""
	servicePrincipal = ResolveServicePrincipal(mailbox)
	inlineImages = LoadInlineImages(images or {})
	CheckImageReferences(htmlBody, inlineImages)
	payload = BuildDraftPayload(subject, htmlBody, inlineImages, to, cc)
	CheckRequestSize(payload)

	accessToken = get_access_token(servicePrincipal)
	CheckDraftPermission(servicePrincipal, accessToken)

	endpoint = DRAFT_ENDPOINT.format(mailbox=quote(mailbox.strip().lower(), safe="@"))
	created = make_graph_request(endpoint, accessToken, method="POST", data=payload)

	result = {field: created.get(field) for field in DRAFT_RESULT_FIELDS}
	result["mailbox"] = mailbox
	result["imageCount"] = len(inlineImages)
	return result



def CreateDraftFromFile(
	mailbox: str,
	subject: str,
	htmlPath: str,
	images: dict | None = None,
	to: list | None = None,
	cc: list | None = None,
) -> dict:
	"""CreateDraft with the HTML read from a .html file."""
	return CreateDraft(mailbox, subject, ReadHTMLFile(htmlPath), images, to, cc)



def WriteDraftEML(
	outputPath: str,
	subject: str,
	htmlBody: str,
	images: dict | None = None,
	to: list | None = None,
	cc: list | None = None,
	sender: str | None = None,
) -> str:
	"""Write the same message as an unsent .eml file, the fallback when Graph cannot be used.

	X-Unsent: 1 asks Outlook to open it as a new message ready to send rather than as a
	received one. Returns the absolute path written.
	"""
	if not outputPath.lower().endswith(EML_EXTENSION):
		frappe.throw(_("The output file must end in {0}: {1}").format(EML_EXTENSION, outputPath))

	inlineImages = LoadInlineImages(images or {})
	CheckImageReferences(htmlBody, inlineImages)

	message = EmailMessage(policy=SMTP)
	message["Subject"] = subject
	message["Date"] = formatdate(localtime=True)
	message["X-Unsent"] = "1"
	AddEMLAddresses(message, "From", [sender] if sender else [])
	AddEMLAddresses(message, "To", to)
	AddEMLAddresses(message, "Cc", cc)

	message.set_content(_("This email needs an HTML-capable mail program."))
	message.add_alternative(htmlBody, subtype="html")
	htmlPart = message.get_payload()[-1]
	for image in inlineImages:
		AddEMLPicture(htmlPart, image)

	absolutePath = os.path.abspath(outputPath)
	with open(absolutePath, "wb") as emlFile:
		emlFile.write(message.as_bytes())
	return absolutePath



def WriteDraftEMLFromFile(
	outputPath: str,
	subject: str,
	htmlPath: str,
	images: dict | None = None,
	to: list | None = None,
	cc: list | None = None,
	sender: str | None = None,
) -> str:
	"""WriteDraftEML with the HTML read from a .html file."""
	return WriteDraftEML(outputPath, subject, ReadHTMLFile(htmlPath), images, to, cc, sender)



def ResolveServicePrincipal(mailbox: str) -> str:
	"""The app login that serves this mailbox, read off its M365 Email Account.

	Restricting drafts to mailboxes that already have an account here keeps the tool to
	known SGC people, rather than any address the tenant happens to hold.
	"""
	address = (mailbox or "").strip().lower()
	if not validate_email_address(address):
		frappe.throw(_("Not a valid mailbox address: {0}").format(mailbox))

	accountFilters = {"service": M365_SERVICE, "email_id": address}
	accountFields = ["name", "m365_service_principal"]
	accounts = frappe.get_all(EMAIL_ACCOUNT_DOCTYPE, filters=accountFilters, fields=accountFields)
	if not accounts:
		frappe.throw(_("No M365 Email Account on this site for {0}").format(address))
	if not accounts[0].m365_service_principal:
		frappe.throw(_("Email Account {0} has no Service Principal").format(accounts[0].name))

	return accounts[0].m365_service_principal



def CheckDraftPermission(servicePrincipal: str, accessToken: str) -> None:
	"""Stop before Graph when the token is malformed or cannot create drafts."""
	if not TOKEN_PATTERN.fullmatch(accessToken or ""):
		message = _("The {0} app login returned a token of an unexpected shape; it was not used.")
		frappe.throw(message.format(servicePrincipal))

	if REQUIRED_GRAPH_ROLE not in ReadTokenRoles(accessToken):
		message = _("The {0} app login lacks the {1} permission, so it cannot create drafts. "
			"Use WriteDraftEML instead, or grant {1} (application) in Azure.")
		frappe.throw(message.format(servicePrincipal, REQUIRED_GRAPH_ROLE))



def ReadTokenRoles(accessToken: str) -> list:
	"""The application permissions (roles claim) the token carries.

	Reads the token's middle segment without verifying the signature: it is only used to
	give a clear message before Graph would refuse anyway, never to grant anything.
	"""
	segments = accessToken.split(".")
	if len(segments) <= JWT_CLAIMS_INDEX:
		return []

	claimsSegment = segments[JWT_CLAIMS_INDEX]
	paddedClaims = claimsSegment + "=" * (-len(claimsSegment) % BASE64_BLOCK)
	try:
		claims = json.loads(base64.urlsafe_b64decode(paddedClaims))
	except ValueError:
		return []
	if not isinstance(claims, dict):
		return []
	return claims.get("roles") or []



def LoadInlineImages(images: dict) -> list:
	"""Read and check each picture: a safe content id, a real PNG/JPEG/GIF, within size."""
	inlineImages = []
	totalBytes = 0
	sizeLimit = MAX_TOTAL_IMAGE_BYTES / BYTES_PER_MEGABYTE

	for contentID, imagePath in images.items():
		if not CONTENT_ID_PATTERN.match(contentID or ""):
			message = _("Picture name {0} must be letters, digits, dot, dash or underscore")
			frappe.throw(message.format(contentID))

		with open(imagePath, "rb") as imageFile:
			content = imageFile.read(MAX_TOTAL_IMAGE_BYTES + 1)

		totalBytes += len(content)
		if totalBytes > MAX_TOTAL_IMAGE_BYTES:
			message = _("Pictures exceed {0:.1f} MB in total; shrink or crop them")
			frappe.throw(message.format(sizeLimit))

		inlineImage = {
			"contentID": contentID,
			"fileName": os.path.basename(imagePath),
			"mimeType": DetectImageType(content, imagePath),
			"content": content,
		}
		inlineImages.append(inlineImage)

	return inlineImages



def DetectImageType(content: bytes, imagePath: str) -> str:
	"""The picture's MIME type, after decoding every frame so only real pictures get through.

	A file that merely starts like a picture fails here, because every frame's pixels have
	to decode. Residual: bytes after a complete picture (data appended after a PNG's last
	chunk, say) are not inspected; that needs a file built on purpose, not a mistyped path.
	"""
	try:
		imageFormat = DecodeAllFrames(content)
	except IMAGE_DECODE_ERRORS:
		imageFormat = None

	if imageFormat not in IMAGE_MIME_TYPES:
		frappe.throw(_("{0} is not a readable PNG, JPEG or GIF picture").format(imagePath))

	return IMAGE_MIME_TYPES[imageFormat]



def DecodeAllFrames(content: bytes) -> str:
	"""Decode every frame and return Pillow's format name; raises when anything is broken.

	Each frame's size is checked before it decodes, because a GIF can open as one pixel and
	declare a far larger canvas on a later frame. Residual: Pillow may allocate while moving
	to a frame, before this check sees it; its own bomb limit still caps that.
	"""
	with Image.open(io.BytesIO(content), formats=list(IMAGE_MIME_TYPES)) as picture:
		for frame in ImageSequence.Iterator(picture):
			width, height = frame.size
			if width * height > MAX_IMAGE_PIXELS:
				raise ValueError("picture has too many pixels")
			frame.load()
		return picture.format



def CheckImageReferences(htmlBody: str, inlineImages: list) -> None:
	"""Every cid: in the HTML has a picture, and every picture is used in the HTML."""
	referenced = set(CID_REFERENCE_PATTERN.findall(htmlBody or ""))
	supplied = {image["contentID"] for image in inlineImages}

	missing = sorted(referenced - supplied)
	if missing:
		frappe.throw(_("The HTML refers to pictures with no file: {0}").format(", ".join(missing)))

	unused = sorted(supplied - referenced)
	if unused:
		frappe.throw(_("Pictures not used in the HTML: {0}").format(", ".join(unused)))



def CheckRequestSize(payload: dict) -> None:
	"""Refuse a draft Graph would reject for size, measured as requests will serialise it."""
	requestBytes = len(json.dumps(payload).encode("utf-8"))
	if requestBytes > GRAPH_MAX_REQUEST_BYTES:
		requestSize = requestBytes / BYTES_PER_MEGABYTE
		sizeLimit = GRAPH_MAX_REQUEST_BYTES / BYTES_PER_MEGABYTE
		message = _("The draft is {0:.1f} MB; Graph takes at most {1:.1f} MB. Shrink the pictures")
		frappe.throw(message.format(requestSize, sizeLimit))



def BuildDraftPayload(
	subject: str,
	htmlBody: str,
	inlineImages: list,
	to: list | None,
	cc: list | None,
) -> dict:
	"""The Graph message body: HTML content plus one inline file attachment per picture."""
	payload = {
		"subject": subject,
		"body": {"contentType": "HTML", "content": htmlBody},
		"toRecipients": BuildGraphRecipients(to),
		"ccRecipients": BuildGraphRecipients(cc),
	}

	payload["attachments"] = [{
		"@odata.type": "#microsoft.graph.fileAttachment",
		"name": image["fileName"],
		"contentType": image["mimeType"],
		"contentId": image["contentID"],
		"isInline": True,
		"contentBytes": base64.b64encode(image["content"]).decode("ascii"),
	} for image in inlineImages]

	return payload



def BuildGraphRecipients(addresses: list | None) -> list:
	return [{"emailAddress": {"address": address}} for address in CheckAddresses(addresses)]



def AddEMLAddresses(message: EmailMessage, header: str, addresses: list | None) -> None:
	checkedAddresses = CheckAddresses(addresses)
	if not checkedAddresses:
		return

	message[header] = ", ".join(checkedAddresses)



def AddEMLPicture(htmlPart: EmailMessage, image: dict) -> None:
	"""Attach one picture to the HTML part as a related, inline part with its Content-ID."""
	mainType, subType = image["mimeType"].split("/")
	contentIDHeader = f"<{image['contentID']}>"
	htmlPart.add_related(
		image["content"],
		maintype=mainType,
		subtype=subType,
		cid=contentIDHeader,
		filename=image["fileName"],
		disposition="inline",
	)



def CheckAddresses(addresses: list | None) -> list:
	"""Trimmed addresses, each confirmed to be a valid email address."""
	if isinstance(addresses, str):
		addresses = [addresses]

	checkedAddresses = []
	for address in addresses or []:
		trimmedAddress = (address or "").strip()
		if not validate_email_address(trimmedAddress):
			frappe.throw(_("Not a valid email address: {0}").format(address))
		checkedAddresses.append(trimmedAddress)
	return checkedAddresses



def ReadHTMLFile(htmlPath: str) -> str:
	"""The body text of a .html file; other file types are refused."""
	if not htmlPath.lower().endswith(HTML_EXTENSIONS):
		frappe.throw(_("The body file must be .html or .htm: {0}").format(htmlPath))

	with open(htmlPath, encoding="utf-8") as htmlFile:
		return htmlFile.read()



def ParseImageArguments(imageArguments: list) -> dict:
	"""--image NAME=PATH pairs as {NAME: PATH}."""
	images = {}
	for imageArgument in imageArguments or []:
		contentID, separator, imagePath = imageArgument.partition("=")
		if not separator or not imagePath:
			raise SystemExit(f"--image must look like NAME=PATH, got: {imageArgument}")
		images[contentID] = imagePath
	return images



def BuildArgumentParser() -> argparse.ArgumentParser:
	description = "Prepare an Outlook draft with inline pictures. Never sends."
	parser = argparse.ArgumentParser(description=description)
	parser.add_argument("--site", required=True, help="Frappe site to read the app login from")
	parser.add_argument("--subject", required=True)
	parser.add_argument("--html", required=True, help="Path to the .html body")
	parser.add_argument("--image", action="append", default=[], help="NAME=PATH, repeatable")
	parser.add_argument("--to", action="append", default=[], help="Recipient, repeatable")
	parser.add_argument("--cc", action="append", default=[], help="Copy recipient, repeatable")
	parser.add_argument("--mailbox", help="Graph route: whose Drafts folder gets the draft")
	parser.add_argument("--eml", help="Fallback route: write an unsent .eml file here instead")
	parser.add_argument("--sender", help="From address written into the .eml (optional)")
	return parser



def RunRoute(arguments: argparse.Namespace, images: dict) -> dict:
	"""Create the draft by whichever route the arguments chose."""
	if arguments.mailbox:
		return CreateDraftFromFile(
			arguments.mailbox,
			arguments.subject,
			arguments.html,
			images,
			arguments.to,
			arguments.cc,
		)

	emlPath = WriteDraftEMLFromFile(
		arguments.eml,
		arguments.subject,
		arguments.html,
		images,
		arguments.to,
		arguments.cc,
		arguments.sender,
	)
	return {"eml": emlPath}



def RunFromCommandLine() -> None:
	"""Entry point: run from the bench's sites folder with `../env/bin/python -m`."""
	arguments = BuildArgumentParser().parse_args()
	if bool(arguments.mailbox) == bool(arguments.eml):
		raise SystemExit("Give exactly one of --mailbox (Graph draft) or --eml (file).")

	images = ParseImageArguments(arguments.image)
	try:
		frappe.init(site=arguments.site, sites_path=os.getcwd())
		frappe.connect()
		result = RunRoute(arguments, images)
	except frappe.ValidationError as error:
		raise SystemExit(f"Not created: {error}")
	finally:
		frappe.destroy()

	print(json.dumps(result, indent=JSON_INDENT))



if __name__ == "__main__":
	RunFromCommandLine()
