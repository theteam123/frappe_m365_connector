# Copyright (c) 2026, TierneyMorris Pty Ltd and contributors
# For license information, please see license.txt

"""Unit tests for the Send to ERP double-create guard.

A second click on "Create" for the same email, or a colleague doing the same in a shared
mailbox, must return the record already made, never create a second one. The cache, the file
lock and the database are patched, so these run without a site. Run them directly:

    cd sites && ../env/bin/python -m unittest m365email.m365email.test_send_to_erp_guard
"""

import contextlib
import json
import unittest
from unittest.mock import MagicMock, patch

from m365email.m365email import send_to_erp
from m365email.m365email.send_to_erp import BuildCreateGuardKey, RunOncePerEmail

MESSAGE_ID = "<abc123@mail.example.com>"
ACTION = "create-Supplier Quote"
RECORD = {"doctype": "Supplier Quote", "name": "quote-1"}
RECORD_URL = "https://erp/app/x"



def RaiseValidationError(message: str, *args: object, **kwargs: object) -> None:
	"""Stands in for frappe.throw, which needs a loaded site to show its message."""
	raise send_to_erp.frappe.ValidationError(message)



class FakeCache:
	def __init__(self) -> None:
		self.values = {}


	def get_value(self, key: str) -> object:
		return self.values.get(key)


	def set_value(self, key: str, value: object, expires_in_sec: int = 0) -> None:
		self.values[key] = value



@contextlib.contextmanager
def NoLock(name: str, timeout: int = 0) -> object:
	yield



class TestRunOncePerEmail(unittest.TestCase):
	def setUp(self) -> None:
		self.cache = FakeCache()
		self.recordExists = True
		patches = [
			patch.object(send_to_erp.frappe, "cache", self.cache, create=True),
			patch("frappe.utils.synchronization.filelock", NoLock),
			patch.object(send_to_erp.frappe.db, "exists", side_effect=self.fakeExists),
			patch.object(send_to_erp.frappe, "has_permission", return_value=True),
			patch.object(send_to_erp, "_ResolveRecordLabel", return_value="PROC-SQTE-3285-0001"),
			patch.object(send_to_erp.frappe.utils, "get_url_to_form", return_value=RECORD_URL),
			patch.object(send_to_erp, "_", side_effect=lambda text: text),
		]
		for patcher in patches:
			patcher.start()
			self.addCleanup(patcher.stop)
		self.create = MagicMock(return_value={"success": True, "record": RECORD, "message": "made"})


	def fakeExists(self, *args: object) -> bool:
		return self.recordExists


	def test_first_click_creates(self) -> None:
		result = RunOncePerEmail(ACTION, MESSAGE_ID, self.create)
		self.assertEqual(result["message"], "made")
		self.create.assert_called_once()


	def test_second_click_returns_earlier_record(self) -> None:
		RunOncePerEmail(ACTION, MESSAGE_ID, self.create)
		repeat = RunOncePerEmail(ACTION, MESSAGE_ID, self.create)
		self.create.assert_called_once()
		self.assertTrue(repeat["already_created"])
		self.assertEqual(repeat["record"], RECORD)


	def test_deleted_earlier_record_allows_a_new_one(self) -> None:
		RunOncePerEmail(ACTION, MESSAGE_ID, self.create)
		self.recordExists = False
		RunOncePerEmail(ACTION, MESSAGE_ID, self.create)
		self.assertEqual(self.create.call_count, 2)


	def test_other_action_on_same_email_is_separate(self) -> None:
		RunOncePerEmail(ACTION, MESSAGE_ID, self.create)
		RunOncePerEmail("enter-supplier-invoice-po-b", MESSAGE_ID, self.create)
		self.assertEqual(self.create.call_count, 2)


	def test_without_message_id_always_runs(self) -> None:
		RunOncePerEmail(ACTION, "", self.create)
		RunOncePerEmail(ACTION, "", self.create)
		self.assertEqual(self.create.call_count, 2)


	def test_failed_create_is_not_remembered(self) -> None:
		failing = MagicMock(side_effect=ValueError("no supplier"))
		with self.assertRaises(ValueError):
			RunOncePerEmail(ACTION, MESSAGE_ID, failing)
		RunOncePerEmail(ACTION, MESSAGE_ID, self.create)
		self.create.assert_called_once()


	def test_unreadable_earlier_record_is_not_named(self) -> None:
		RunOncePerEmail(ACTION, MESSAGE_ID, self.create)
		with patch.object(send_to_erp.frappe, "has_permission", return_value=False):
			repeat = RunOncePerEmail(ACTION, MESSAGE_ID, self.create)
		self.assertNotIn("record", repeat)
		self.assertNotIn("PROC-SQTE", repeat["message"])


	def test_key_hides_message_id(self) -> None:
		key = BuildCreateGuardKey(ACTION, MESSAGE_ID)
		self.assertNotIn("abc123", key)
		self.assertRegex(key, r"^[a-z-]+-[0-9a-f]{64}$")



class TestParseAttachments(unittest.TestCase):
	def setUp(self) -> None:
		patcher = patch.object(send_to_erp.frappe, "throw", side_effect=RaiseValidationError)
		patcher.start()
		self.addCleanup(patcher.stop)


	def test_non_text_content_is_refused(self) -> None:
		with self.assertRaises(send_to_erp.frappe.ValidationError):
			send_to_erp._ParseAttachments('[{"file_name": "a.pdf", "content_base64": 1}]')


	def test_non_base64_content_is_refused(self) -> None:
		with self.assertRaises(send_to_erp.frappe.ValidationError):
			send_to_erp._ParseAttachments('[{"file_name": "a.pdf", "content_base64": "<script>"}]')


	def test_valid_content_is_accepted(self) -> None:
		payload = '[{"file_name": "a.pdf", "content_base64": "JVBERg=="}]'
		parsed = send_to_erp._ParseAttachments(payload)
		self.assertEqual(len(parsed), 1)


	def test_total_size_is_capped(self) -> None:
		chunk = "A" * (30 * 1024 * 1024)
		attachments = [{"file_name": f"{index}.pdf", "content_base64": chunk} for index in range(5)]
		payload = json.dumps(attachments)
		with self.assertRaises(send_to_erp.frappe.ValidationError):
			send_to_erp._ParseAttachments(payload)


	def test_bad_length_is_refused(self) -> None:
		with self.assertRaises(send_to_erp.frappe.ValidationError):
			send_to_erp._ParseAttachments('[{"file_name": "a.pdf", "content_base64": "A"}]')


	def test_missing_content_is_allowed_and_skipped(self) -> None:
		parsed = send_to_erp._ParseAttachments('[{"file_name": "a.pdf"}]')
		self.assertEqual(len(parsed), 1)


	def test_line_breaks_in_base64_are_accepted(self) -> None:
		payload = '[{"file_name": "a.txt", "content_base64": "YQ==\\r\\n"}]'
		self.assertEqual(len(send_to_erp._ParseAttachments(payload)), 1)



class TestInternalSender(unittest.TestCase):
	def setUp(self) -> None:
		from m365email.m365email import send_to_erp_people

		self.people = send_to_erp_people
		frappeDB = send_to_erp_people.frappe.db
		existsPatch = patch.object(frappeDB, "exists", side_effect=self.fakeExists)
		existsPatch.start()
		self.addCleanup(existsPatch.stop)


	def fakeExists(self, doctype: str, filters: object) -> bool:
		staff = {"angela.obrien@sgcaustralia.com.au"}
		if not isinstance(filters, dict) or filters.get("user_type") != "System User":
			return False
		return (filters.get("email") or filters.get("name")) in staff


	def test_colleague_is_internal(self) -> None:
		self.assertTrue(self.people.IsInternalSender(" Angela.OBrien@sgcaustralia.com.au "))


	def test_supplier_is_not_internal(self) -> None:
		self.assertFalse(self.people.IsInternalSender("sales@miw.net.au"))


	def test_colleague_never_yields_a_contact(self) -> None:
		self.assertIsNone(self.people.FindContactForEmail("angela.obrien@sgcaustralia.com.au"))


	def test_match_on_user_email_field(self) -> None:
		self.assertTrue(self.people.IsInternalSender("ANGELA.OBRIEN@sgcaustralia.com.au"))


	def test_empty_sender_is_not_internal(self) -> None:
		self.assertFalse(self.people.IsInternalSender(""))
