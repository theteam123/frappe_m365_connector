# Copyright (c) 2026, TierneyMorris Pty Ltd and contributors
# For license information, please see license.txt

"""Unit tests for drafts with inline pictures.

NOTE: Microsoft Graph and the token are both patched, so these run on any bench without a
network call or a write. Run them directly:

    cd sites && ../env/bin/python -m unittest m365email.m365email.test_draft_mail
"""

import base64
import email
import io
import json
import os
import shutil
import struct
import tempfile
import unittest
from email.policy import default as DEFAULT_POLICY
from unittest.mock import patch

import frappe
from PIL import Image

from m365email.m365email import draft_mail

MAILBOX = "someone@example.com.au"
SERVICE_PRINCIPAL = "Test Tenant"
HTML_WITH_PICTURE = '<p>Look:</p><img src="cid:shot1" alt="A screenshot">'
MODULE = "m365email.m365email.draft_mail"
PICTURE_SIZE = (4, 4)
PICTURE_COLOUR = (0, 82, 147)
HUGE_PICTURE_SIZE = (10_000, 5_000)
HUGE_GIF_FRAME_SIZE = (7_000, 7_000)

# NOTE: The pieces of a hand-built GIF, so a test can declare any frame size it likes.
GIF_HEADER = b"GIF89a"
GIF_SCREEN_FLAGS = b"\x80\x00\x00"  # global two-colour palette follows
GIF_PALETTE = b"\x00\x00\x00" + b"\xff\xff\xff"
GIF_IMAGE_SEPARATOR = b"\x2c"
GIF_IMAGE_DATA = b"\x00\x02\x02\x44\x01\x00"  # no local palette, one LZW-coded pixel
GIF_TRAILER = b"\x3b"
TRUNCATED_FRACTION = 2
FAKE_GIF_BYTES = b'GIF89a{"client_secret": "SYNTHETIC"}'



def RaiseValidationError(message: str, *args: object, **kwargs: object) -> None:
	"""Stands in for frappe.throw, which needs a loaded site to show its message."""
	raise frappe.ValidationError(message)



def MakePNGBytes() -> bytes:
	output = io.BytesIO()
	Image.new("RGB", PICTURE_SIZE, PICTURE_COLOUR).save(output, format="PNG")
	return output.getvalue()



def MakeGIFFrame(width: int, height: int) -> bytes:
	descriptor = struct.pack("<HHHH", 0, 0, width, height)
	return GIF_IMAGE_SEPARATOR + descriptor + GIF_IMAGE_DATA



def MakeExpandingGIF() -> bytes:
	"""A one-pixel GIF whose second frame declares a canvas far over the pixel cap."""
	header = GIF_HEADER + struct.pack("<HH", 1, 1) + GIF_SCREEN_FLAGS + GIF_PALETTE
	hugeWidth, hugeHeight = HUGE_GIF_FRAME_SIZE
	return header + MakeGIFFrame(1, 1) + MakeGIFFrame(hugeWidth, hugeHeight) + GIF_TRAILER



def MakeToken(roles: list) -> str:
	claims = base64.urlsafe_b64encode(json.dumps({"roles": roles}).encode()).decode()
	return f"header.{claims.rstrip('=')}.signature"



PNG_BYTES = MakePNGBytes()



class DraftMailTestCase(unittest.TestCase):

	def setUp(self) -> None:
		throwPatcher = patch("frappe.throw", side_effect=RaiseValidationError)
		throwPatcher.start()
		self.addCleanup(throwPatcher.stop)
		self.folder = tempfile.mkdtemp()
		self.picturePath = self.writeFile("shot1.png", PNG_BYTES)


	def tearDown(self) -> None:
		shutil.rmtree(self.folder)


	def writeFile(self, fileName: str, content: bytes) -> str:
		filePath = os.path.join(self.folder, fileName)
		with open(filePath, "wb") as outputFile:
			outputFile.write(content)
		return filePath



class TestImageChecks(DraftMailTestCase):

	def testPayloadCarriesInlinePicture(self) -> None:
		images = draft_mail.LoadInlineImages({"shot1": self.picturePath})
		body = HTML_WITH_PICTURE
		payload = draft_mail.BuildDraftPayload("Subject", body, images, [MAILBOX], None)

		attachment = payload["attachments"][0]
		self.assertTrue(attachment["isInline"])
		self.assertEqual(attachment["contentId"], "shot1")
		self.assertEqual(attachment["contentType"], "image/png")
		self.assertEqual(base64.b64decode(attachment["contentBytes"]), PNG_BYTES)
		self.assertEqual(payload["body"]["contentType"], "HTML")


	def testFileThatIsNotAPictureIsRefused(self) -> None:
		configPath = self.writeFile("site_config.png", b'{"db_password": "x"}')
		with self.assertRaises(frappe.ValidationError):
			draft_mail.LoadInlineImages({"shot1": configPath})


	def testFileThatOnlyStartsLikeAPictureIsRefused(self) -> None:
		fakePath = self.writeFile("fake.gif", FAKE_GIF_BYTES)
		with self.assertRaises(frappe.ValidationError):
			draft_mail.LoadInlineImages({"shot1": fakePath})


	def testTruncatedPictureIsRefused(self) -> None:
		truncatedPath = self.writeFile("cut.png", PNG_BYTES[:len(PNG_BYTES) // TRUNCATED_FRACTION])
		with self.assertRaises(frappe.ValidationError):
			draft_mail.LoadInlineImages({"shot1": truncatedPath})


	def testPictureWithTooManyPixelsIsRefused(self) -> None:
		output = io.BytesIO()
		Image.new("1", HUGE_PICTURE_SIZE).save(output, format="PNG")
		hugePath = self.writeFile("huge.png", output.getvalue())
		with self.assertRaises(frappe.ValidationError):
			draft_mail.LoadInlineImages({"huge": hugePath})


	def testGIFWithHugeLaterFrameIsRefused(self) -> None:
		bombPath = self.writeFile("bomb.gif", MakeExpandingGIF())
		with self.assertRaises(frappe.ValidationError):
			draft_mail.LoadInlineImages({"bomb": bombPath})


	def testOversizePicturesAreRefused(self) -> None:
		bigPath = self.writeFile("big.png", PNG_BYTES + b"0" * draft_mail.MAX_TOTAL_IMAGE_BYTES)
		with self.assertRaises(frappe.ValidationError):
			draft_mail.LoadInlineImages({"big": bigPath})


	def testRequestOverGraphLimitIsRefused(self) -> None:
		oversizeBody = "x" * draft_mail.GRAPH_MAX_REQUEST_BYTES
		payload = draft_mail.BuildDraftPayload("Subject", oversizeBody, [], None, None)
		with self.assertRaises(frappe.ValidationError):
			draft_mail.CheckRequestSize(payload)


	def testUnsafePictureNameIsRefused(self) -> None:
		with self.assertRaises(frappe.ValidationError):
			draft_mail.LoadInlineImages({'shot1"><script>': self.picturePath})


	def testMissingPictureIsRefused(self) -> None:
		with self.assertRaises(frappe.ValidationError):
			draft_mail.CheckImageReferences(HTML_WITH_PICTURE, [])


	def testUnusedPictureIsRefused(self) -> None:
		images = draft_mail.LoadInlineImages({"shot1": self.picturePath})
		with self.assertRaises(frappe.ValidationError):
			draft_mail.CheckImageReferences("<p>No pictures here</p>", images)


	def testInvalidRecipientIsRefused(self) -> None:
		with self.assertRaises(frappe.ValidationError):
			draft_mail.CheckAddresses(["not an address"])



class TestCreateDraft(DraftMailTestCase):

	def createDraft(self, accessToken: str) -> tuple:
		graphResponse = {"id": "AAMk", "isDraft": True}
		resolvePatcher = patch(f"{MODULE}.ResolveServicePrincipal", return_value=SERVICE_PRINCIPAL)
		tokenPatcher = patch(f"{MODULE}.get_access_token", return_value=accessToken)
		graphPatcher = patch(f"{MODULE}.make_graph_request", return_value=graphResponse)
		for patcher in (resolvePatcher, tokenPatcher):
			patcher.start()
			self.addCleanup(patcher.stop)
		graphRequest = graphPatcher.start()
		self.addCleanup(graphPatcher.stop)

		images = {"shot1": self.picturePath}
		result = draft_mail.CreateDraft(MAILBOX, "Subject", HTML_WITH_PICTURE, images, [MAILBOX])
		return result, graphRequest


	def testDraftIsCreatedNeverSent(self) -> None:
		result, graphRequest = self.createDraft(MakeToken(["Mail.Send", "Mail.ReadWrite"]))

		endpoint = graphRequest.call_args.args[0]
		self.assertEqual(endpoint, f"/users/{MAILBOX}/messages")
		self.assertEqual(graphRequest.call_args.kwargs["method"], "POST")
		self.assertEqual(graphRequest.call_count, 1)
		self.assertTrue(result["isDraft"])
		self.assertEqual(result["imageCount"], 1)


	def testMissingPermissionStopsBeforeGraph(self) -> None:
		with self.assertRaises(frappe.ValidationError):
			self.createDraft(MakeToken(["Mail.Send", "Mail.Read"]))


	def testMalformedTokenIsNeverUsed(self) -> None:
		malformedToken = MakeToken(["Mail.ReadWrite"]) + "\nSYNTHETIC"
		with self.assertRaises(frappe.ValidationError) as raised:
			self.createDraft(malformedToken)
		self.assertNotIn("SYNTHETIC", str(raised.exception))


	def testTokenWithTrailingLineBreakIsNeverUsed(self) -> None:
		with self.assertRaises(frappe.ValidationError):
			self.createDraft(MakeToken(["Mail.ReadWrite"]) + "\n")


	def testUnreadableTokenHasNoRoles(self) -> None:
		self.assertEqual(draft_mail.ReadTokenRoles("not-a-token"), [])
		self.assertEqual(draft_mail.ReadTokenRoles("a.!!!.c"), [])



class TestWriteDraftEML(DraftMailTestCase):

	def testEMLIsUnsentWithRelatedPicture(self) -> None:
		outputPath = os.path.join(self.folder, "draft.eml")
		images = {"shot1": self.picturePath}
		draft_mail.WriteDraftEML(outputPath, "Subject", HTML_WITH_PICTURE, images, [MAILBOX])

		with open(outputPath, "rb") as emlFile:
			message = email.message_from_binary_file(emlFile, policy=DEFAULT_POLICY)

		self.assertEqual(message["X-Unsent"], "1")
		pictures = [part for part in message.walk() if part.get_content_type() == "image/png"]
		self.assertEqual(len(pictures), 1)
		self.assertEqual(pictures[0]["Content-ID"], "<shot1>")
		self.assertEqual(pictures[0].get_content(), PNG_BYTES)


	def testOutputMustBeEML(self) -> None:
		textPath = os.path.join(self.folder, "draft.txt")
		images = {"shot1": self.picturePath}
		with self.assertRaises(frappe.ValidationError):
			draft_mail.WriteDraftEML(textPath, "Subject", HTML_WITH_PICTURE, images)
