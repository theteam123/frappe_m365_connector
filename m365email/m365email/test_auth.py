# Copyright (c) 2026, TierneyMorris Pty Ltd and contributors
# For license information, please see license.txt

"""Unit tests for what a failed sign-in writes to the Error Log.

NOTE: MSAL and the database are both patched, so these run on any bench without a network
call or a write. Run them directly:

    cd sites && ../env/bin/python -m unittest m365email.m365email.test_auth
"""

import json
import unittest
from contextlib import ExitStack
from unittest.mock import MagicMock, patch

import frappe

from m365email.m365email import auth

MODULE = "m365email.m365email.auth"
SERVICE_PRINCIPAL = "Test Tenant"
SYNTHETIC_SECRET = "SYNTHETIC-SECRET-VALUE"
FAILED_REPLY = {
	"error": "invalid_client",
	"error_description": "AADSTS7000215: Invalid client secret provided.",
	"error_codes": [7000215],
	"trace_id": "trace",
	"correlation_id": "correlation",
	"unexpected_field": SYNTHETIC_SECRET,
}



def RaiseValidationError(message: str, *args: object, **kwargs: object) -> None:
	"""Stands in for frappe.throw, which needs a loaded site to show its message."""
	raise frappe.ValidationError(message)



class TestTokenErrorLogging(unittest.TestCase):

	def testSummaryKeepsOnlyNamedFields(self) -> None:
		summary = auth.SummarizeTokenError(FAILED_REPLY)
		self.assertIn("AADSTS7000215", summary)
		self.assertIn("correlation", summary)
		self.assertNotIn(SYNTHETIC_SECRET, summary)


	def testNestedValuesInsideNamedFieldsAreOmitted(self) -> None:
		nestedReply = {"trace_id": {"client_secret": SYNTHETIC_SECRET}, "error_codes": [1, b"x"]}
		summary = auth.SummarizeTokenError(nestedReply)
		self.assertNotIn(SYNTHETIC_SECRET, summary)
		self.assertIn(auth.OMITTED_VALUE, summary)


	def testLongValuesAreCut(self) -> None:
		longReply = {"error_description": "x" * (auth.MAX_LOGGED_FIELD_CHARACTERS * 2)}
		summary = json.loads(auth.SummarizeTokenError(longReply))
		self.assertEqual(len(summary["error_description"]), auth.MAX_LOGGED_FIELD_CHARACTERS)


	def testNestedDescriptionNeverReachesTheScreen(self) -> None:
		nestedReply = {"error": "invalid_client", "error_description": {"secret": SYNTHETIC_SECRET}}
		raisedError = self.signInWith(nestedReply)
		self.assertNotIn(SYNTHETIC_SECRET, str(raisedError))


	def signInWith(self, reply: dict) -> Exception:
		"""Run a sign-in that fails with this MSAL reply and return what it raised."""
		settings = MagicMock(token_expires_at=None, scopes="https://graph.microsoft.com/.default")
		msalApp = MagicMock()
		msalApp.acquire_token_for_client.return_value = reply
		with ExitStack() as patches:
			patches.enter_context(patch(f"{MODULE}.frappe.get_doc", return_value=settings))
			patches.enter_context(patch(f"{MODULE}.get_msal_app", return_value=msalApp))
			patches.enter_context(patch("frappe.throw", side_effect=RaiseValidationError))
			patches.enter_context(patch(f"{MODULE}.frappe.log_error"))
			with self.assertRaises(frappe.ValidationError) as raised:
				auth.get_access_token(SERVICE_PRINCIPAL)
		return raised.exception


	def testFailedSignInLogsNoUnexpectedField(self) -> None:
		settings = MagicMock(token_expires_at=None, scopes="https://graph.microsoft.com/.default")
		msalApp = MagicMock()
		msalApp.acquire_token_for_client.return_value = FAILED_REPLY

		with ExitStack() as patches:
			patches.enter_context(patch(f"{MODULE}.frappe.get_doc", return_value=settings))
			patches.enter_context(patch(f"{MODULE}.get_msal_app", return_value=msalApp))
			patches.enter_context(patch("frappe.throw", side_effect=RaiseValidationError))
			logError = patches.enter_context(patch(f"{MODULE}.frappe.log_error"))
			with self.assertRaises(frappe.ValidationError) as raised:
				auth.get_access_token(SERVICE_PRINCIPAL)

		self.assertIn("AADSTS7000215", str(raised.exception))

		loggedMessage = logError.call_args.kwargs["message"]
		self.assertIn("invalid_client", loggedMessage)
		self.assertNotIn(SYNTHETIC_SECRET, loggedMessage)
