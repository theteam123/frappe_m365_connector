# Copyright (c) 2026, TierneyMorris Pty Ltd and contributors
# For license information, please see license.txt

"""Unit tests for the mailbox probe and account provisioning.

NOTE: Microsoft Graph and the database are both patched, so these run on any bench without
a network call or a write. Run them directly:

    cd sites && ../env/bin/python -m unittest m365email.m365email.test_provisioning
"""

import unittest
from unittest.mock import MagicMock, patch

import frappe

from m365email.m365email import graph_api, provisioning

TOKEN = "token"
LIVE_ADDRESS = "someone@example.com.au"



def MakeResponse(statusCode: int, errorCode: str = "", retryAfter: str = "") -> MagicMock:
	response = MagicMock()
	response.status_code = statusCode
	response.headers = {"Retry-After": retryAfter} if retryAfter else {}
	if errorCode:
		response.json.return_value = {"error": {"code": errorCode, "message": "x"}}
	else:
		response.json.return_value = {"value": []}
	return response



class TestMailboxExists(unittest.TestCase):

	def testLiveMailbox(self) -> None:
		with patch("m365email.m365email.graph_api.requests.get", return_value=MakeResponse(200)):
			self.assertTrue(graph_api.MailboxExists(LIVE_ADDRESS, TOKEN))


	def testUnknownUserMeansNoMailbox(self) -> None:
		response = MakeResponse(404, "ErrorInvalidUser")
		with patch("m365email.m365email.graph_api.requests.get", return_value=response):
			self.assertFalse(graph_api.MailboxExists(LIVE_ADDRESS, TOKEN))


	def testPermissionProblemIsRaisedNotTreatedAsMissing(self) -> None:
		response = MakeResponse(403, "ErrorAccessDenied")
		with patch("m365email.m365email.graph_api.requests.get", return_value=response):
			with self.assertRaises(frappe.ValidationError):
				graph_api.MailboxExists(LIVE_ADDRESS, TOKEN)


	def testThrottlingIsRetriedThenAnswered(self) -> None:
		responses = [MakeResponse(429, retryAfter="0"), MakeResponse(200)]
		with patch("m365email.m365email.graph_api.requests.get", side_effect=responses) as getMock:
			with patch("m365email.m365email.graph_api.time.sleep"):
				self.assertTrue(graph_api.MailboxExists(LIVE_ADDRESS, TOKEN))
		self.assertEqual(getMock.call_count, 2)



class TestDomainHelpers(unittest.TestCase):

	def testNormalizeDomainTolerantOfTyping(self) -> None:
		self.assertEqual(provisioning.NormalizeDomain(" @Example.COM.au "), "example.com.au")


	def testExtractDomain(self) -> None:
		self.assertEqual(provisioning.ExtractDomain("A.B@Example.com"), "example.com")
		self.assertEqual(provisioning.ExtractDomain("not-an-address"), "")



class TestEnsureM365AccountForUser(unittest.TestCase):

	def setUp(self) -> None:
		self.getValue = patch("m365email.m365email.provisioning.frappe.db.get_value").start()
		self.exists = patch("m365email.m365email.provisioning.frappe.db.exists").start()
		findSPTarget = "m365email.m365email.provisioning.FindProvisioningServicePrincipal"
		self.findSP = patch(findSPTarget).start()
		self.probe = patch("m365email.m365email.provisioning.ProbeMailbox").start()
		self.create = patch("m365email.m365email.provisioning.CreateM365Account").start()
		self.addCleanup(patch.stopall)
		self.getValue.return_value = frappe._dict(name=LIVE_ADDRESS, enabled=1)
		self.exists.return_value = False
		self.findSP.return_value = MagicMock(name="SP")


	def testDisabledUserGetsNothing(self) -> None:
		self.getValue.return_value = frappe._dict(name=LIVE_ADDRESS, enabled=0)
		outcome = provisioning.EnsureM365AccountForUser(LIVE_ADDRESS)
		self.assertEqual(outcome, provisioning.PROVISION_NO_USER)
		self.create.assert_not_called()


	def testExistingAccountIsLeftAlone(self) -> None:
		self.exists.return_value = True
		outcome = provisioning.EnsureM365AccountForUser(LIVE_ADDRESS)
		self.assertEqual(outcome, provisioning.PROVISION_EXISTS)
		self.probe.assert_not_called()


	def testNoServicePrincipalForDomain(self) -> None:
		self.findSP.return_value = None
		outcome = provisioning.EnsureM365AccountForUser(LIVE_ADDRESS)
		self.assertEqual(outcome, provisioning.PROVISION_NOT_CONFIGURED)


	def testProbesWhenCallerDoesNotKnow(self) -> None:
		self.probe.return_value = False
		outcome = provisioning.EnsureM365AccountForUser(LIVE_ADDRESS)
		self.assertEqual(outcome, provisioning.PROVISION_NO_MAILBOX)
		self.probe.assert_called_once()


	def testCallerAnswerSkipsTheProbe(self) -> None:
		outcome = provisioning.EnsureM365AccountForUser(LIVE_ADDRESS, LIVE_ADDRESS, True)
		self.assertEqual(outcome, provisioning.PROVISION_CREATED)
		self.probe.assert_not_called()
		self.create.assert_called_once()



if __name__ == "__main__":
	unittest.main()
