# Copyright (c) 2026, TierneyMorris Pty Ltd and contributors
# For license information, please see license.txt

"""Unit tests for the add-in's automatic import of a new supplier quote's PDF.

Paul, 2026-11-07: when everything checks out the lines are imported; otherwise "Open in ERP"
goes to the import page with the PDF loaded. A failure part-way must undo only the import.
The erp gate is patched, so these run without records:

    cd sites && ../env/bin/python -m unittest m365email.m365email.test_send_to_erp_autoimport
"""

import unittest
from unittest.mock import patch
from urllib.parse import parse_qs, urlparse

from m365email.m365email import send_to_erp_procurement as procurement

QUOTE = "quote-1"
PDF = {"file_name": "Quote MBL102897.pdf", "file_url": "/private/files/Quote MBL102897.pdf"}
SPREADSHEET = {"file_name": "lines.xlsx", "file_url": "/private/files/lines.xlsx"}
GATE = "erp.procurement.doctype.supplier_quote.from_email.TryAutoImportQuoteFile"



def SiteURL(path: str) -> str:
	return "https://erp" + path



class TestAutoImportNewQuote(unittest.TestCase):
	def setUp(self) -> None:
		patches = [
			patch.object(procurement.frappe.db, "savepoint"),
			patch.object(procurement.frappe.db, "rollback"),
			patch.object(procurement.frappe, "log_error"),
			patch.object(procurement.frappe.utils, "get_url", side_effect=SiteURL),
			patch.object(procurement, "_", side_effect=lambda text: text),
		]
		self.mocks = [patcher.start() for patcher in patches]
		for patcher in patches:
			self.addCleanup(patcher.stop)


	def test_no_pdf_does_nothing(self) -> None:
		with patch(GATE) as gate:
			outcome = procurement.AutoImportNewQuote(QUOTE, [SPREADSHEET])
		gate.assert_not_called()
		self.assertEqual(outcome, {"note": "", "record_url": None})


	def test_imported_keeps_the_quote_link(self) -> None:
		with patch(GATE, return_value={"status": "imported", "imported": 7}):
			outcome = procurement.AutoImportNewQuote(QUOTE, [SPREADSHEET, PDF])
		self.assertIn("7", outcome["note"])
		self.assertIsNone(outcome["record_url"])


	def test_needs_review_opens_the_import_page_with_the_pdf(self) -> None:
		reason = "some lines did not match an item exactly"
		with patch(GATE, return_value={"status": "needs_review", "reason": reason}):
			outcome = procurement.AutoImportNewQuote(QUOTE, [PDF])
		self.assertIn(reason, outcome["note"])
		link = urlparse(outcome["record_url"])
		self.assertEqual(link.path, "/app/supplier-quote-import")
		self.assertEqual(parse_qs(link.query), {"quote": [QUOTE], "file": [PDF["file_url"]]})


	def test_failure_undoes_only_the_import(self) -> None:
		rollback = self.mocks[1]
		with patch(GATE, side_effect=ValueError("bad pdf")):
			outcome = procurement.AutoImportNewQuote(QUOTE, [PDF])
		rollback.assert_called_once_with(save_point=procurement.AUTO_IMPORT_SAVEPOINT)
		self.assertTrue(outcome["record_url"])



	def test_lost_savepoint_falls_back_to_plain_rollback(self) -> None:
		rollback = self.mocks[1]
		rollback.side_effect = [RuntimeError("SAVEPOINT does not exist"), None]
		with patch(GATE, side_effect=ValueError("bad pdf")):
			with patch.object(procurement.frappe.db, "exists", return_value=True):
				outcome = procurement.AutoImportNewQuote(QUOTE, [PDF])
		self.assertEqual(rollback.call_count, 2)
		self.assertEqual(rollback.call_args_list[1].kwargs, {})
		self.assertTrue(outcome["record_url"])



	def test_lost_quote_fails_the_request(self) -> None:
		rollback = self.mocks[1]
		rollback.side_effect = [RuntimeError("SAVEPOINT does not exist"), None]
		with patch(GATE, side_effect=ValueError("bad pdf")):
			with patch.object(procurement.frappe.db, "exists", return_value=False):
				with patch.object(procurement.frappe, "throw", side_effect=RuntimeError("failed")):
					with self.assertRaises(RuntimeError):
						procurement.AutoImportNewQuote(QUOTE, [PDF])
