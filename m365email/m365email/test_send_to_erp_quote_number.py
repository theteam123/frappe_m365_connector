# Copyright (c) 2026, TierneyMorris Pty Ltd and contributors
# For license information, please see license.txt

"""Unit tests for the supplier quote number the add-in requires on "Create supplier quote".

The supplier's own quote number is required (Paul, 2026-11-26). The permission check and the
ERP app's module are patched, so these run without a site. Run them directly:

    cd sites && ../env/bin/python -m unittest m365email.m365email.test_send_to_erp_quote_number
"""

import sys
import unittest
from unittest.mock import MagicMock, patch

from m365email.m365email import send_to_erp_procurement as procurement

ERP_MODULE = "erp.procurement.doctype.supplier_quote.from_email"
BLANK_NUMBERS = ["", "   ", None]
SENDER = {"name": "A", "email": "a@x.com", "body_html": "", "sent_date": "", "contact": ""}


def RaiseValidationError(message: str, *args: object, **kwargs: object) -> None:
	"""Stands in for frappe.throw, which needs a loaded site to show its message."""
	raise procurement.frappe.ValidationError(message)


class TestSupplierQuoteNumberRequired(unittest.TestCase):
	def test_blankNumberIsRefusedBeforeAnythingIsWritten(self) -> None:
		fakeModule = MagicMock()
		with patch.dict(sys.modules, {ERP_MODULE: fakeModule}), patch.object(
			procurement.frappe, "has_permission", return_value=True
		), patch.object(procurement.frappe, "throw", side_effect=RaiseValidationError):
			for blank in BLANK_NUMBERS:
				with self.assertRaises(procurement.frappe.ValidationError):
					procurement.CreateSupplierQuoteFromEmail(
						"JOB-1", "SUP-1", SENDER, [], None, blank
					)
		fakeModule.CreateDraftQuoteFromEmail.assert_not_called()


if __name__ == "__main__":
	unittest.main()
