# Copyright (c) 2026, TierneyMorris Pty Ltd and contributors
# For license information, please see license.txt

"""Unit tests for the welcome email. Settings, the log, the token and Graph are all patched.

    cd sites && ../env/bin/python -m unittest m365email.m365email.test_welcome
"""

import unittest
from types import SimpleNamespace
from unittest.mock import MagicMock, patch

import frappe
import requests

from m365email.m365email import entra, welcome

PASSWORD = "Xk9-Unique-Secret-Pw"
UPN = "jane.smith@example.com.au"
PERSONAL = "jane@gmail.com"
REFERENCE = {"doctype": "Employee", "name": "HR-EMP-1"}
USER_ID = "44444444-4444-4444-4444-444444444444"



def MakeResult(mode: str, password: str | None = PASSWORD) -> dict:
	return {
		"success": True, "mode": mode, "userPrincipalName": UPN, "userID": USER_ID,
		"steps": [], "isResume": False, "warnings": [],
		"temporaryPassword": entra.SecretValue(password) if password else None,
	}



class TestWelcomeEmail(unittest.TestCase):

	def setUp(self) -> None:
		self.logRows: list = []
		self.settings = SimpleNamespace(
			welcome_sender="support@sgcaustralia.com.au",
			welcome_mail_service_principal="Mail SP",
		)
		self.token = MagicMock(return_value="tok")
		self.post = MagicMock(return_value=SimpleNamespace(status_code=202, text=""))
		self.sendmail = MagicMock()
		patches = [
			patch.object(entra, "GetSettings", return_value=self.settings),
			patch.object(entra, "WriteLog", side_effect=self.writeLog),
			patch.object(welcome, "get_access_token", self.token),
			patch.object(welcome.requests, "post", self.post),
			patch.object(welcome, "_", lambda text: text),
			patch.object(frappe, "sendmail", self.sendmail, create=True),
		]
		for patcher in patches:
			patcher.start()
			self.addCleanup(patcher.stop)


	def writeLog(self, context, action, status, detail, *args) -> None:
		self.logRows.append({"action": action, "status": status, "detail": detail,
			"mode": context.mode})


	def testLiveSendsFromSupportWithoutSentItemsCopy(self) -> None:
		result = MakeResult(entra.MODE_LIVE)
		outcome = welcome.SendWelcomeEmail(result, "Jane", PERSONAL, REFERENCE)
		self.assertEqual(outcome["outcome"], welcome.OUTCOME_SENT)
		url = self.post.call_args.args[0]
		self.assertEqual(url, f"{entra.GRAPH_BASE_URL}/users/support@sgcaustralia.com.au/sendMail")
		payload = self.post.call_args.kwargs["json"]
		self.assertIs(payload["saveToSentItems"], False)
		message = payload["message"]
		self.assertEqual(message["toRecipients"][0]["emailAddress"]["address"], PERSONAL)
		self.assertIn(PASSWORD, message["body"]["content"])
		self.assertIn(UPN, message["body"]["content"])
		self.assertIn("Welcome to SGC", message["subject"])
		self.token.assert_called_once_with("Mail SP")
		self.assertEqual(self.logRows[0]["action"], "Send Welcome Email")
		self.assertEqual(self.logRows[0]["status"], entra.STATUS_SUCCESS)
		self.assertEqual(result["steps"][-1]["step"], welcome.STEP_LABEL)


	def testBodyEscapesInsertedValues(self) -> None:
		body = welcome.BuildWelcomeBody("<b>Jo</b>", UPN, "a<script>")
		self.assertNotIn("<script>", body)
		self.assertIn("&lt;script&gt;", body)


	def testSimulationSendsNothing(self) -> None:
		outcome = welcome.SendWelcomeEmail(
			MakeResult(entra.MODE_SIMULATION, "SIMULATION-NOT-A-REAL-PASSWORD"),
			"Jane", PERSONAL, REFERENCE,
		)
		self.assertEqual(outcome["outcome"], welcome.OUTCOME_SIMULATED)
		self.post.assert_not_called()
		self.token.assert_not_called()
		self.assertEqual(self.logRows[0]["status"], entra.STATUS_SKIPPED)
		self.assertEqual(self.logRows[0]["mode"], entra.MODE_SIMULATION)


	def testNoPersonalEmailIsSkipped(self) -> None:
		for missing in ("", None, "not an email"):
			outcome = welcome.SendWelcomeEmail(
				MakeResult(entra.MODE_LIVE), "Jane", missing, REFERENCE
			)
			self.assertEqual(outcome["outcome"], welcome.OUTCOME_SKIPPED)
		self.post.assert_not_called()


	def testGraphRefusalIsFailedWithFixedText(self) -> None:
		self.post.return_value = SimpleNamespace(status_code=403, text=PASSWORD)
		outcome = welcome.SendWelcomeEmail(
			MakeResult(entra.MODE_LIVE), "Jane", PERSONAL, REFERENCE
		)
		self.assertEqual(outcome["outcome"], welcome.OUTCOME_FAILED)
		self.assertIn("403", outcome["reason"])
		self.assertEqual(self.logRows[0]["status"], entra.STATUS_FAILED)


	def testNetworkAndTokenErrorsNeverCarryTheirText(self) -> None:
		self.post.side_effect = requests.ConnectTimeout(f"boom {PASSWORD}")
		outcome = welcome.SendWelcomeEmail(
			MakeResult(entra.MODE_LIVE), "Jane", PERSONAL, REFERENCE
		)
		self.assertEqual(outcome["outcome"], welcome.OUTCOME_FAILED)
		self.assertEqual(outcome["reason"], welcome.SEND_FAILED_REASON)
		self.post.side_effect = None
		self.token.side_effect = RuntimeError(f"token {PASSWORD}")
		outcome = welcome.SendWelcomeEmail(
			MakeResult(entra.MODE_LIVE), "Jane", PERSONAL, REFERENCE
		)
		self.assertEqual(outcome["reason"], welcome.TOKEN_FAILED_REASON)


	def testTimeoutOrResetAfterSendingIsDeliveryUnknown(self) -> None:
		errors = (
			requests.ReadTimeout(PASSWORD),
			requests.ConnectionError(PASSWORD),
			requests.exceptions.ChunkedEncodingError(PASSWORD),
			requests.exceptions.ContentDecodingError(PASSWORD),
		)
		for error in errors:
			self.post.side_effect = error
			self.logRows.clear()
			result = MakeResult(entra.MODE_LIVE)
			outcome = welcome.SendWelcomeEmail(result, "Jane", PERSONAL, REFERENCE)
			self.assertEqual(outcome["outcome"], welcome.OUTCOME_UNKNOWN)
			self.assertEqual(outcome["reason"], "delivery unknown")
			self.assertEqual(self.logRows[0]["status"], entra.STATUS_FAILED)
			self.assertEqual(self.logRows[0]["detail"], "delivery unknown")
			self.assertNotIn(PASSWORD, repr(self.logRows) + repr(outcome) + repr(result))


	def testEmailTellsReaderNotToReplyAndSetsNoReplyTo(self) -> None:
		welcome.SendWelcomeEmail(MakeResult(entra.MODE_LIVE), "Jane", PERSONAL, REFERENCE)
		message = self.post.call_args.kwargs["json"]["message"]
		self.assertNotIn("replyTo", message)
		self.assertIn("do not reply", message["body"]["content"])
		self.assertIn("support@sgcaustralia.com.au in a new email", message["body"]["content"])


	def testMissingPrincipalIsFailed(self) -> None:
		self.settings.welcome_mail_service_principal = ""
		outcome = welcome.SendWelcomeEmail(
			MakeResult(entra.MODE_LIVE), "Jane", PERSONAL, REFERENCE
		)
		self.assertEqual(outcome["reason"], welcome.NO_PRINCIPAL_REASON)


	def testLogFailureDoesNotRaise(self) -> None:
		with patch.object(entra, "WriteLog", side_effect=RuntimeError(PASSWORD)):
			outcome = welcome.SendWelcomeEmail(
				MakeResult(entra.MODE_LIVE), "Jane", PERSONAL, REFERENCE
			)
		self.assertEqual(outcome["outcome"], welcome.OUTCOME_SENT)
		self.assertTrue(outcome["warnings"])


	def testPasswordNeverInLogsOrStepsOrOutcome(self) -> None:
		for status in (200, 202):
			self.post.return_value = SimpleNamespace(status_code=status, text=PASSWORD)
			result = MakeResult(entra.MODE_LIVE)
			outcome = welcome.SendWelcomeEmail(result, "Jane", PERSONAL, REFERENCE)
			dump = repr(self.logRows) + repr(result["steps"]) + repr(outcome) + repr(result)
			self.assertNotIn(PASSWORD, dump)
		self.sendmail.assert_not_called()



if __name__ == "__main__":
	unittest.main()
