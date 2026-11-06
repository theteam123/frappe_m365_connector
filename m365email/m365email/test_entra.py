# Copyright (c) 2026, TierneyMorris Pty Ltd and contributors
# For license information, please see license.txt

"""Unit tests for the Entra provisioning engine.

NOTE: Settings, the log and Microsoft Graph are all patched, so these need no site, no
network and no database. Run them directly:

    cd sites && ../env/bin/python -m unittest m365email.m365email.test_entra
"""

import re
import unittest
from types import SimpleNamespace
from unittest.mock import MagicMock, patch

import frappe

from m365email.m365email import entra

UPN = "jane.smith@example.com.au"
SKU = "11111111-1111-1111-1111-111111111111"
GROUP_ONE = "22222222-2222-2222-2222-222222222222"
GROUP_TWO = "33333333-3333-3333-3333-333333333333"
USER_ID = "44444444-4444-4444-4444-444444444444"
MODE_SIM = entra.MODE_SIMULATION
PERSON = {"firstName": "Jane", "lastName": "Smith", "jobTitle": "Estimator", "department": ""}
REFERENCE = {"doctype": "Employee", "name": "HR-EMP-1"}
PROFILE = {
	"name": "Office",
	"skuID": SKU,
	"licenceName": "Business Premium",
	"groups": [{"id": GROUP_ONE, "name": "All Staff"}, {"id": GROUP_TWO, "name": "Office"}],
}



REAL_HAS_LOGGED = entra.HasLogged
REAL_GET_OPEN_TARGETS = entra.GetOpenTargets



def RaiseValidationError(message: str, *args: object, **kwargs: object) -> None:
	raise frappe.ValidationError(message)



class FakeLog:
	"""In-memory stand-in for the log table, so the orchestrators' reads and writes line up."""

	def __init__(self) -> None:
		self.rows: list = []


	def write(
		self,
		context: entra.StepContext,
		action: str,
		status: str,
		detail: str,
		targetID: str = "",
		targetName: str = "",
	) -> None:
		self.rows.append({
			"action": action, "status": status, "mode": context.mode, "detail": detail,
			"user_principal_name": context.userPrincipalName, "target_id": targetID,
			"target_name": targetName, "entra_user_id": context.userID,
			"reference_doctype": context.referenceDoctype,
			"reference_name": context.referenceName,
		})


	def hasLogged(self, context: entra.StepContext, action: str, targetID: str = "") -> bool:
		if not (context.userID and context.referenceDoctype and context.referenceName):
			return False
		for row in self.rows:
			isMatch = row["action"] == action and row["entra_user_id"] == context.userID
			isMatch = isMatch and row["reference_doctype"] == context.referenceDoctype
			isMatch = isMatch and row["reference_name"] == context.referenceName
			isMatch = isMatch and row["mode"] == context.mode
			isMatch = isMatch and row["status"] == entra.STATUS_SUCCESS
			if isMatch and (not targetID or row["target_id"] == targetID):
				return True
		return False


	def hasCreateRecord(self, context: entra.StepContext, userID: str) -> bool:
		for row in self.rows:
			isMatch = row["action"] == entra.ACTION_CREATE and row["mode"] == context.mode
			isMatch = isMatch and row["status"] == entra.STATUS_SUCCESS
			isMatch = isMatch and row["entra_user_id"] == userID and bool(userID)
			isMatch = isMatch and row["reference_doctype"] == context.referenceDoctype
			isMatch = isMatch and row["reference_name"] == context.referenceName
			if isMatch:
				return True
		return False


	def openTargets(
		self,
		addAction: str,
		removeAction: str,
		mode: str,
		userID: str,
		createdFor: dict,
	) -> list:
		def ids(action: str) -> list:
			if not userID:
				return []
			return [
				row for row in self.rows
				if row["action"] == action
				and row["mode"] == mode
				and row["status"] == entra.STATUS_SUCCESS
				and row["entra_user_id"] == userID
			]
		removed = {row["target_id"] for row in ids(removeAction)}
		return [
			(row["target_id"], row["target_name"])
			for row in ids(addAction)
			if row["target_id"] not in removed
			and row["reference_name"] == createdFor.get("name")
		]



class EngineTestCase(unittest.TestCase):

	mode = entra.MODE_SIMULATION
	realHasLogged = staticmethod(REAL_HAS_LOGGED)
	realGetOpenTargets = staticmethod(REAL_GET_OPEN_TARGETS)

	def setUp(self) -> None:
		self.log = FakeLog()
		settings = SimpleNamespace(mode=self.mode, usage_location="AU", force_password_change=1)
		settings.get_password = lambda fieldName: "not-a-real-secret"
		settings.tenant_id = "55555555-5555-5555-5555-555555555555"
		settings.client_id = "66666666-6666-6666-6666-666666666666"
		for target, replacement in (
			("GetSettings", lambda: settings),
			("WriteLog", self.log.write),
			("HasLogged", self.log.hasLogged),
			("HasCreateRecord", self.log.hasCreateRecord),
			("GetOpenTargets", self.log.openTargets),
			("LoadProfile", lambda name: PROFILE),
		):
			patcher = patch.object(entra, target, replacement)
			patcher.start()
			self.addCleanup(patcher.stop)
		throwPatcher = patch("frappe.throw", side_effect=RaiseValidationError)
		throwPatcher.start()
		self.addCleanup(throwPatcher.stop)



class TestAddressProposal(EngineTestCase):

	def testSanitisesNames(self) -> None:
		address = entra.ProposeAddress("Seán", "O'Brien-Smith", "@Example.com.au", MODE_SIM)
		self.assertEqual(address, "sean.obrien-smith@example.com.au")


	def testSpacesAndDigitsRemoved(self) -> None:
		self.assertEqual(entra.SanitiseNamePart("Mary Ann 2"), "maryann")


	def testCollisionsGetNumericSuffix(self) -> None:
		taken = {"jane.smith@example.com.au", "jane.smith2@example.com.au"}
		isFree = lambda address, mode: address not in taken
		with patch.object(entra, "IsAddressAvailable", side_effect=isFree):
			proposed = entra.ProposeAddress("Jane", "Smith", "example.com.au", MODE_SIM)
			self.assertEqual(proposed, "jane.smith3@example.com.au")


	def testEmptyNameRefused(self) -> None:
		with self.assertRaises(entra.ProvisioningError):
			entra.ProposeAddress("???", "Smith", "example.com.au", MODE_SIM)


	def testAddressValidation(self) -> None:
		for bad in ("a b@example.com", "x@y", "../x@example.com", "a@example.com/../b", ""):
			with self.assertRaises(entra.ProvisioningError):
				entra.ValidateAddress(bad)



class TestSimulation(EngineTestCase):

	def testFullProvisionLogsEveryStepAsSimulation(self) -> None:
		with patch.object(entra.requests, "request", side_effect=AssertionError("network call")):
			result = entra.ProvisionAccount(UPN, PERSON, "Office", REFERENCE)
		self.assertTrue(result["success"])
		self.assertTrue(result["userID"].startswith(entra.SIMULATION_ID_PREFIX))
		self.assertIsInstance(result["temporaryPassword"], entra.SecretValue)
		self.assertEqual(result["temporaryPassword"].reveal(), entra.SIMULATION_PASSWORD)
		actions = [row["action"] for row in self.log.rows]
		expected = [
			entra.ACTION_LOOKUP,
			entra.ACTION_CREATE,
			entra.ACTION_LICENCE,
			entra.ACTION_ADD_GROUP,
			entra.ACTION_ADD_GROUP,
		]
		self.assertEqual(actions, expected)
		self.assertTrue(all(row["mode"] == entra.MODE_SIMULATION for row in self.log.rows))


	def testDeprovisionUndoesOnlyWhatWasLogged(self) -> None:
		first = entra.ProvisionAccount(UPN, PERSON, "Office", REFERENCE)
		result = entra.DeprovisionAccount(UPN, first["userID"], REFERENCE, mode=MODE_SIM)
		labels = [step["step"] for step in result["steps"]]
		expected = [
			"Disable account",
			"Remove from group: All Staff",
			"Remove from group: Office",
			"Release licence: Business Premium",
		]
		self.assertEqual(labels, expected)
		self.assertTrue(result["success"])


	def testDeprovisionWithNothingLoggedOnlyDisables(self) -> None:
		result = entra.DeprovisionAccount(UPN, "sim-1", REFERENCE, mode=MODE_SIM)
		self.assertEqual([step["step"] for step in result["steps"]], ["Disable account"])


	def testNoPasswordInAnyLogRow(self) -> None:
		result = entra.ProvisionAccount(UPN, PERSON, "Office", REFERENCE)
		entra.DeprovisionAccount(UPN, result["userID"], REFERENCE, mode=MODE_SIM)
		password = result["temporaryPassword"].reveal()
		for row in self.log.rows:
			for value in row.values():
				self.assertNotIn(password, str(value))



class TestModeAndIdentity(EngineTestCase):
	"""Findings: mode is read once per run, and step history is keyed on the object id."""

	mode = entra.MODE_SIMULATION

	def testModeReadOnceAndPassedToEveryPrimitive(self) -> None:
		modes = iter([entra.MODE_SIMULATION, entra.MODE_LIVE, entra.MODE_LIVE, entra.MODE_LIVE])
		seen: list = []
		def record(*args: object) -> None:
			seen.append(args[-1])
		created = {"id": USER_ID, "password": "Secret-Pass-123!"}
		def recordCreate(*args: object) -> dict:
			seen.append(args[-1])
			return created
		def recordGroup(*args: object) -> str:
			seen.append(args[-1])
			return ""
		with patch.object(entra, "GetMode", side_effect=lambda: next(modes)) as getMode, \
			patch.object(entra, "FindUser", side_effect=lambda upn, mode: seen.append(mode)), \
			patch.object(entra, "GetLicenceAvailability", side_effect=lambda sku, mode: 3), \
			patch.object(entra, "CreateUser", side_effect=recordCreate), \
			patch.object(entra, "AssignLicence", side_effect=record), \
			patch.object(entra, "AddToGroup", side_effect=recordGroup):
			result = entra.ProvisionAccount(UPN, PERSON, "Office", REFERENCE)
		self.assertEqual(getMode.call_count, 1)
		self.assertEqual(result["mode"], entra.MODE_SIMULATION)
		self.assertTrue(seen)
		self.assertEqual(set(seen), {entra.MODE_SIMULATION})
		self.assertTrue(all(row["mode"] == entra.MODE_SIMULATION for row in self.log.rows))


	def testDeprovisionNeverReadsModeItself(self) -> None:
		with patch.object(entra, "GetMode", side_effect=AssertionError("GetMode called")):
			result = entra.DeprovisionAccount(UPN, "sim-1", REFERENCE, mode=MODE_SIM)
		self.assertEqual(result["mode"], MODE_SIM)
		self.assertTrue(all(row["mode"] == MODE_SIM for row in self.log.rows))


	def testDeprovisionRefusesOffMode(self) -> None:
		with self.assertRaises(frappe.ValidationError):
			entra.DeprovisionAccount(UPN, "sim-1", REFERENCE, mode=entra.MODE_OFF)
		self.assertEqual(self.log.rows, [])


	def testLogFailureAfterCreateStillReturnsPassword(self) -> None:
		realWrite = self.log.write
		def failOnCreate(context, action, status, *args):
			if action == entra.ACTION_CREATE:
				raise RuntimeError("Secret-Pass-123! db down")
			realWrite(context, action, status, *args)
		created = {"id": USER_ID, "password": "Secret-Pass-123!"}
		with patch.object(entra, "WriteLog", side_effect=failOnCreate), \
			patch.object(entra, "CreateUser", return_value=created), \
			patch.object(entra, "AssignLicence") as assignLicence:
			result = entra.ProvisionAccount(UPN, PERSON, "Office", REFERENCE)
		self.assertEqual(result["temporaryPassword"].reveal(), "Secret-Pass-123!")
		self.assertFalse(result["success"])
		self.assertEqual(result["warnings"], [entra.LOG_WRITE_WARNING])
		assignLicence.assert_not_called()
		self.assertNotIn("Secret-Pass", repr(result) + str(result["warnings"]))
		self.assertNotIn("db down", str(result["steps"]))


	def testLogFailureOnLaterStepAlsoKeepsPassword(self) -> None:
		realWrite = self.log.write
		def failOnGroup(context, action, status, *args):
			if action == entra.ACTION_ADD_GROUP and status == entra.STATUS_SUCCESS:
				raise RuntimeError("boom")
			realWrite(context, action, status, *args)
		with patch.object(entra, "WriteLog", side_effect=failOnGroup):
			result = entra.ProvisionAccount(UPN, PERSON, "Office", REFERENCE)
		self.assertEqual(result["temporaryPassword"].reveal(), entra.SIMULATION_PASSWORD)
		self.assertIn(entra.LOG_WRITE_WARNING, result["warnings"])


	def testErrorAfterCreateFailsTheRunAndKeepsPassword(self) -> None:
		errors = [frappe.PermissionError("secret-in-text"), RuntimeError("db deadlock")]
		for error in errors:
			with self.subTest(error=type(error).__name__):
				self.log.rows.clear()
				with patch.object(entra, "FinishProvision", side_effect=error):
					result = entra.ProvisionAccount(UPN, PERSON, "Office", REFERENCE)
				self.assertFalse(result["success"])
				self.assertEqual(
					result["temporaryPassword"].reveal(), entra.SIMULATION_PASSWORD
				)
				steps = result["steps"]
				failed = [step for step in steps if step["status"] == entra.STATUS_FAILED]
				self.assertEqual(len(failed), 1)
				self.assertNotIn("secret-in-text", str(result["steps"]) + str(result["warnings"]))


	def testLogFailureBeforeAccountExistsStillRaises(self) -> None:
		with patch.object(entra, "WriteLog", side_effect=RuntimeError("down")):
			with self.assertRaises(entra.ProvisioningError):
				entra.ProvisionAccount(UPN, PERSON, "Office", REFERENCE)


	def testPrimitivesNeverReadSettingsMode(self) -> None:
		with patch.object(entra, "GetMode", side_effect=AssertionError("GetMode called")):
			self.assertIsNone(entra.FindUser(UPN, MODE_SIM))
			self.assertEqual(entra.ListSubscribedSkus(MODE_SIM), [])
			self.assertEqual(
				entra.GetLicenceAvailability(SKU, MODE_SIM), entra.SIMULATION_FREE_SEATS
			)
			self.assertTrue(entra.CreateUser(UPN, PERSON, MODE_SIM)["id"])
			self.assertIsNone(entra.AssignLicence(USER_ID, SKU, MODE_SIM))
			self.assertIsNone(entra.ReleaseLicence(USER_ID, SKU, MODE_SIM))
			self.assertEqual(entra.AddToGroup(USER_ID, GROUP_ONE, MODE_SIM), "")
			self.assertIsNone(entra.RemoveFromGroup(USER_ID, GROUP_ONE, MODE_SIM))
			self.assertIsNone(entra.DisableUser(USER_ID, MODE_SIM))


	def testOldRowsFromReusedAddressDoNotSuppressNewAccountSteps(self) -> None:
		oldID = "99999999-9999-9999-9999-999999999999"
		oldContext = entra.StepContext(
			entra.MODE_SIMULATION, UPN, REFERENCE["doctype"], REFERENCE["name"], oldID
		)
		self.log.write(oldContext, entra.ACTION_ADD_GROUP, entra.STATUS_SUCCESS, "", GROUP_ONE)
		self.log.write(oldContext, entra.ACTION_LICENCE, entra.STATUS_SUCCESS, "", SKU)
		newContext = entra.StepContext(
			entra.MODE_SIMULATION, UPN, REFERENCE["doctype"], REFERENCE["name"], USER_ID
		)
		self.assertFalse(entra.HasLogged(newContext, entra.ACTION_ADD_GROUP, GROUP_ONE))
		self.assertTrue(self.log.hasLogged(oldContext, entra.ACTION_ADD_GROUP, GROUP_ONE))
		self.assertFalse(self.log.hasLogged(newContext, entra.ACTION_ADD_GROUP, GROUP_ONE))
		result = entra.ProvisionAccount(UPN, PERSON, "Office", REFERENCE)
		groupSteps = [step for step in result["steps"] if step["step"].startswith("Add to group")]
		self.assertEqual([step["status"] for step in groupSteps], ["Success", "Success"])


	def testHasLoggedQueryKeysOnObjectIDReferenceAndMode(self) -> None:
		context = entra.StepContext(
			entra.MODE_LIVE, UPN, REFERENCE["doctype"], REFERENCE["name"], USER_ID
		)
		with patch.object(entra.frappe, "get_all", return_value=[]) as getAll:
			self.realHasLogged(context, entra.ACTION_ADD_GROUP, GROUP_ONE)
		filters = getAll.call_args.kwargs["filters"]
		self.assertEqual(filters["entra_user_id"], USER_ID)
		self.assertEqual(filters["reference_name"], REFERENCE["name"])
		self.assertEqual(filters["mode"], entra.MODE_LIVE)
		self.assertNotIn("user_principal_name", filters)


	def testHasLoggedWithoutObjectIDIsFalse(self) -> None:
		context = entra.StepContext(
			entra.MODE_LIVE, UPN, REFERENCE["doctype"], REFERENCE["name"], ""
		)
		self.assertFalse(self.realHasLogged(context, entra.ACTION_ADD_GROUP, GROUP_ONE))


	def testGetOpenTargetsQueryHasNoAddress(self) -> None:
		with patch.object(entra.frappe, "get_all", return_value=[]) as getAll:
			self.realGetOpenTargets(
				entra.ACTION_ADD_GROUP, entra.ACTION_REMOVE_GROUP, entra.MODE_LIVE,
				USER_ID, REFERENCE,
			)
		for call in getAll.call_args_list:
			self.assertNotIn("user_principal_name", call.kwargs["filters"])
			self.assertEqual(call.kwargs["filters"]["entra_user_id"], USER_ID)


	def testPasswordStaysWrappedInEngineResult(self) -> None:
		result = entra.ProvisionAccount(UPN, PERSON, "Office", REFERENCE)
		password = result["temporaryPassword"]
		self.assertIsInstance(password, entra.SecretValue)
		self.assertNotIn(password.reveal(), repr(result) + str(result))



class TestOff(EngineTestCase):

	mode = entra.MODE_OFF

	def testEveryEntryPointRefuses(self) -> None:
		with self.assertRaises(frappe.ValidationError):
			entra.ProvisionAccount(UPN, PERSON, "Office", REFERENCE)
		with self.assertRaises(frappe.ValidationError):
			entra.DeprovisionAccount(UPN, USER_ID, REFERENCE, mode=entra.MODE_OFF)
		with self.assertRaises(frappe.ValidationError):
			entra.GetMode()
		self.assertEqual(self.log.rows, [])



class TestGuardsAndResume(EngineTestCase):

	mode = entra.MODE_LIVE

	def testNoFreeSeatStopsBeforeCreate(self) -> None:
		with patch.object(entra, "FindUser", return_value=None), \
			patch.object(entra, "GetLicenceAvailability", return_value=0), \
			patch.object(entra, "CreateUser") as createUser:
			result = entra.ProvisionAccount(UPN, PERSON, "Office", REFERENCE)
		self.assertFalse(result["success"])
		createUser.assert_not_called()
		self.assertIsNone(result["temporaryPassword"])
		self.assertIn("No free", result["steps"][-1]["detail"])


	def testOldLogRowForSameAddressDoesNotAuthoriseResume(self) -> None:
		oldContext = entra.StepContext(
			entra.MODE_LIVE, UPN, REFERENCE["doctype"], REFERENCE["name"], "someone-elses-id"
		)
		self.log.write(oldContext, entra.ACTION_CREATE, entra.STATUS_SUCCESS, "")
		existing = {"id": USER_ID, "assignedLicenses": []}
		with patch.object(entra, "FindUser", return_value=existing), \
			patch.object(entra, "AddToGroup") as addToGroup:
			result = entra.ProvisionAccount(UPN, PERSON, "Office", REFERENCE)
		self.assertFalse(result["success"])
		addToGroup.assert_not_called()
		self.assertIn("not created by the ERP", result["steps"][-1]["detail"])


	def testOtherEmployeesCreateRowDoesNotAuthoriseResume(self) -> None:
		other = {"doctype": "Employee", "name": "HR-EMP-2"}
		otherArgs = (entra.MODE_LIVE, UPN, other["doctype"], other["name"], USER_ID)
		otherContext = entra.StepContext(*otherArgs)
		self.log.write(otherContext, entra.ACTION_CREATE, entra.STATUS_SUCCESS, "")
		existing = {"id": USER_ID, "assignedLicenses": []}
		with patch.object(entra, "FindUser", return_value=existing):
			result = entra.ProvisionAccount(UPN, PERSON, "Office", REFERENCE)
		self.assertFalse(result["success"])


	def testCreateRowStoresNewObjectIDAndReference(self) -> None:
		created = {"id": USER_ID, "password": "Secret-Pass-123!"}
		with patch.object(entra, "FindUser", return_value=None), \
			patch.object(entra, "GetLicenceAvailability", return_value=3), \
			patch.object(entra, "CreateUser", return_value=created), \
			patch.object(entra, "AssignLicence"), \
			patch.object(entra, "AddToGroup", return_value=""):
			entra.ProvisionAccount(UPN, PERSON, "Office", REFERENCE)
		createRows = [row for row in self.log.rows if row["action"] == entra.ACTION_CREATE]
		self.assertEqual(len(createRows), 1)
		self.assertEqual(createRows[0]["entra_user_id"], USER_ID)
		self.assertEqual(createRows[0]["reference_name"], REFERENCE["name"])


	def testAlreadyMemberLoggedSkippedAndNeverRemoved(self) -> None:
		created = {"id": USER_ID, "password": "Secret-Pass-123!"}
		answers = {GROUP_ONE: entra.ALREADY_MEMBER_DETAIL, GROUP_TWO: ""}
		with patch.object(entra, "FindUser", return_value=None), \
			patch.object(entra, "GetLicenceAvailability", return_value=3), \
			patch.object(entra, "CreateUser", return_value=created), \
			patch.object(entra, "AssignLicence"), \
			patch.object(entra, "AddToGroup", side_effect=lambda user, group, mode: answers[group]):
			entra.ProvisionAccount(UPN, PERSON, "Office", REFERENCE)
		groupRows = {
			row["target_id"]: row["status"]
			for row in self.log.rows if row["action"] == entra.ACTION_ADD_GROUP
		}
		self.assertEqual(groupRows[GROUP_ONE], entra.STATUS_SKIPPED)
		self.assertEqual(groupRows[GROUP_TWO], entra.STATUS_SUCCESS)
		with patch.object(entra, "RemoveFromGroup") as removeFromGroup, \
			patch.object(entra, "ReleaseLicence"), \
			patch.object(entra, "DisableUser"):
			offboarding = {"doctype": "HR Employee Offboarding", "name": "OFF-1"}
			entra.DeprovisionAccount(UPN, USER_ID, offboarding, REFERENCE, mode=entra.MODE_LIVE)
		removeFromGroup.assert_called_once_with(USER_ID, GROUP_TWO, entra.MODE_LIVE)


	def testLicenceAlreadyOnAccountLoggedSkippedAndNeverReleased(self) -> None:
		context = entra.StepContext(
			entra.MODE_LIVE, UPN, REFERENCE["doctype"], REFERENCE["name"], USER_ID
		)
		self.log.write(context, entra.ACTION_CREATE, entra.STATUS_SUCCESS, "")
		existing = {"id": USER_ID, "assignedLicenses": [{"skuId": SKU}]}
		with patch.object(entra, "FindUser", return_value=existing), \
			patch.object(entra, "AssignLicence") as assignLicence, \
			patch.object(entra, "AddToGroup", return_value=""):
			entra.ProvisionAccount(UPN, PERSON, "Office", REFERENCE)
		assignLicence.assert_not_called()
		licenceRows = [row for row in self.log.rows if row["action"] == entra.ACTION_LICENCE]
		self.assertEqual([row["status"] for row in licenceRows], [entra.STATUS_SKIPPED])
		with patch.object(entra, "ReleaseLicence") as releaseLicence, \
			patch.object(entra, "RemoveFromGroup"), \
			patch.object(entra, "DisableUser"):
			offboarding = {"doctype": "HR Employee Offboarding", "name": "OFF-1"}
			entra.DeprovisionAccount(UPN, USER_ID, offboarding, REFERENCE, mode=entra.MODE_LIVE)
		releaseLicence.assert_not_called()


	def testUnexpectedErrorHasSafeMessageAndNoPassword(self) -> None:
		created = {"id": USER_ID, "password": "Secret-Pass-123!"}
		with patch.object(entra, "FindUser", return_value=None), \
			patch.object(entra, "GetLicenceAvailability", return_value=3), \
			patch.object(entra, "CreateUser", return_value=created), \
			patch.object(entra, "AssignLicence", side_effect=KeyError("Secret-Pass-123!")):
			result = entra.ProvisionAccount(UPN, PERSON, "Office", REFERENCE)
		self.assertFalse(result["success"])
		self.assertNotIn("Secret-Pass-123!", str(result["steps"]))
		self.assertNotIn("Secret-Pass-123!", str(self.log.rows))
		leakyError = RuntimeError("tok-abc Secret-Pass-123!")
		with patch.object(entra, "LookUpExisting", side_effect=leakyError):
			with self.assertRaises(entra.ProvisioningError) as caught:
				entra.ProvisionAccount(UPN, PERSON, "Office", REFERENCE)
		self.assertNotIn("Secret", str(caught.exception))
		self.assertNotIn("tok-abc", str(caught.exception))
		self.assertIsNone(caught.exception.__cause__)
		self.assertTrue(caught.exception.__suppress_context__)


	def testSecretValueNeverPrints(self) -> None:
		secret = entra.SecretValue("Secret-Pass-123!")
		self.assertNotIn("Secret", repr(secret) + str(secret) + f"{secret}")
		self.assertEqual(secret.reveal(), "Secret-Pass-123!")


	def testExistingForeignAccountRefused(self) -> None:
		existing = {"id": USER_ID, "assignedLicenses": []}
		with patch.object(entra, "FindUser", return_value=existing), \
			patch.object(entra, "CreateUser") as createUser:
			result = entra.ProvisionAccount(UPN, PERSON, "Office", REFERENCE)
		self.assertFalse(result["success"])
		createUser.assert_not_called()


	def testFailureMidwayStopsAndReportsSteps(self) -> None:
		def failLicence(userID: str, skuID: str, mode: str) -> None:
			raise entra.ProvisioningError("Microsoft Graph 400 Forbidden: boom")
		created = {"id": USER_ID, "password": "Secret-Pass-123!"}
		with patch.object(entra, "FindUser", return_value=None), \
			patch.object(entra, "GetLicenceAvailability", return_value=3), \
			patch.object(entra, "CreateUser", return_value=created), \
			patch.object(entra, "AssignLicence", side_effect=failLicence), \
			patch.object(entra, "AddToGroup") as addToGroup:
			result = entra.ProvisionAccount(UPN, PERSON, "Office", REFERENCE)
		self.assertFalse(result["success"])
		addToGroup.assert_not_called()
		statuses = [step["status"] for step in result["steps"]]
		self.assertEqual(statuses, ["Success", "Success", "Success", "Failed"])
		for row in self.log.rows:
			self.assertNotIn("Secret-Pass-123!", str(row))


	def testResumeFinishesMissingStepsWithoutRecreating(self) -> None:
		context = entra.StepContext(
			entra.MODE_LIVE, UPN, REFERENCE["doctype"], REFERENCE["name"], USER_ID
		)
		self.log.write(context, entra.ACTION_CREATE, entra.STATUS_SUCCESS, "")
		addGroup = entra.ACTION_ADD_GROUP
		self.log.write(context, addGroup, entra.STATUS_SUCCESS, "", GROUP_ONE, "All Staff")
		existing = {"id": USER_ID, "assignedLicenses": []}
		with patch.object(entra, "FindUser", return_value=existing), \
			patch.object(entra, "GetLicenceAvailability", return_value=2), \
			patch.object(entra, "CreateUser") as createUser, \
			patch.object(entra, "AssignLicence") as assignLicence, \
			patch.object(entra, "AddToGroup") as addToGroup:
			result = entra.ProvisionAccount(UPN, PERSON, "Office", REFERENCE)
		self.assertTrue(result["success"])
		self.assertTrue(result["isResume"])
		self.assertIsNone(result["temporaryPassword"])
		createUser.assert_not_called()
		assignLicence.assert_called_once_with(USER_ID, SKU, entra.MODE_LIVE)
		addToGroup.assert_called_once_with(USER_ID, GROUP_TWO, entra.MODE_LIVE)



class TestLiveHttp(EngineTestCase):

	mode = entra.MODE_LIVE

	def setUp(self) -> None:
		super().setUp()
		for target in ("GetProvisioningToken",):
			patcher = patch.object(entra, target, return_value="tok")
			patcher.start()
			self.addCleanup(patcher.stop)


	@staticmethod
	def response(status: int, payload: dict | None = None) -> MagicMock:
		reply = MagicMock()
		reply.status_code = status
		reply.content = b"{}" if payload is not None else b""
		reply.headers = {}
		reply.json.return_value = payload or {}
		return reply


	def testFindUserNotFound(self) -> None:
		notFound = self.response(404, {"error": {"code": "Request_ResourceNotFound"}})
		with patch.object(entra.requests, "request", return_value=notFound):
			self.assertIsNone(entra.FindUser(UPN, entra.MODE_LIVE))


	def testFindUserEncodesAddress(self) -> None:
		payload = {
			"id": USER_ID,
			"displayName": "Jane",
			"accountEnabled": True,
			"assignedLicenses": [],
		}
		found = self.response(200, payload)
		with patch.object(entra.requests, "request", return_value=found) as request:
			user = entra.FindUser("Jane.Smith@Example.com.au", entra.MODE_LIVE)
		self.assertEqual(user["id"], USER_ID)
		expectedUrl = f"{entra.GRAPH_BASE_URL}/users/jane.smith@example.com.au"
		self.assertEqual(request.call_args.args[1], expectedUrl)


	def testLicenceAvailabilityFollowsPaging(self) -> None:
		nextLink = f"{entra.GRAPH_BASE_URL}/subscribedSkus?skip=1"
		pageOne = {"value": [{"skuId": "x"}], "@odata.nextLink": nextLink}
		pageTwo = {"value": [{"skuId": SKU, "prepaidUnits": {"enabled": 10}, "consumedUnits": 7}]}
		pages = [self.response(200, pageOne), self.response(200, pageTwo)]
		with patch.object(entra.requests, "request", side_effect=pages):
			self.assertEqual(entra.GetLicenceAvailability(SKU, entra.MODE_LIVE), 3)


	def testForeignPagingLinkRefused(self) -> None:
		page = {"value": [], "@odata.nextLink": "https://evil.example.com/x"}
		with patch.object(entra.requests, "request", return_value=self.response(200, page)):
			with self.assertRaises(entra.ProvisioningError):
				entra.GraphGetAll("/subscribedSkus")


	def testNonIntegerRetryAfterFallsBackAndDoesNotRaise(self) -> None:
		for raw in ("Wed, 21 Oct 2026 07:28:00 GMT", "abc", "", None, "-5"):
			seconds = entra.ParseRetryAfter(raw)
			self.assertGreaterEqual(seconds, 0)
			self.assertLessEqual(seconds, entra.MAX_RETRY_AFTER_SECONDS)
		self.assertEqual(entra.ParseRetryAfter("not-a-number"), entra.DEFAULT_RETRY_AFTER_SECONDS)
		self.assertEqual(entra.ParseRetryAfter("7"), 7)
		self.assertEqual(entra.ParseRetryAfter("9999"), entra.MAX_RETRY_AFTER_SECONDS)


	def testThrottledWithBadRetryAfterRetriesOnce(self) -> None:
		throttled = self.response(429, {})
		throttled.headers = {"Retry-After": "Wed, 21 Oct 2026 07:28:00 GMT"}
		good = self.response(200, {"value": []})
		with patch.object(entra.requests, "request", side_effect=[throttled, good]), \
			patch.object(entra.time, "sleep") as sleep:
			status, _payload = entra.GraphRequest("GET", "/subscribedSkus")
		self.assertEqual(status, 200)
		sleep.assert_called_once_with(entra.DEFAULT_RETRY_AFTER_SECONDS)


	def testUnexpectedHttpErrorIsSafe(self) -> None:
		boom = ValueError("Bearer tok-abc Secret-Pass-123!")
		with patch.object(entra.requests, "request", side_effect=boom):
			with self.assertRaises(entra.ProvisioningError) as caught:
				entra.GraphRequest("POST", "/users", body={"passwordProfile": {"password": "x"}})
		message = str(caught.exception)
		for secret in ("tok-abc", "Secret-Pass-123!", "Bearer"):
			self.assertNotIn(secret, message)


	def testCreateUserFailureMessageHasNoPassword(self) -> None:
		boom = RuntimeError("body leaked")
		with patch.object(entra.requests, "request", side_effect=boom):
			with self.assertRaises(entra.ProvisioningError) as caught:
				entra.CreateUser(UPN, PERSON, entra.MODE_LIVE)
		self.assertNotIn("leaked", str(caught.exception))
		self.assertIsNone(caught.exception.__cause__)


	def testAlreadyMemberReturnsMarker(self) -> None:
		message = "One or more added object references already exist for 'members'."
		error = {"error": {"code": "Request_BadRequest", "message": message}}
		with patch.object(entra.requests, "request", return_value=self.response(400, error)):
			outcome = entra.AddToGroup(USER_ID, GROUP_ONE, entra.MODE_LIVE)
		self.assertEqual(outcome, entra.ALREADY_MEMBER_DETAIL)


	def testAlreadyMemberIsSuccess(self) -> None:
		message = (
			"One or more added object references already exist for the following "
			"modified properties: 'members'."
		)
		error = {"error": {"code": "Request_BadRequest", "message": message}}
		with patch.object(entra.requests, "request", return_value=self.response(400, error)):
			entra.AddToGroup(USER_ID, GROUP_ONE, entra.MODE_LIVE)


	def testOtherBadRequestRaises(self) -> None:
		error = {"error": {"code": "Request_BadRequest", "message": "nope"}}
		with patch.object(entra.requests, "request", return_value=self.response(400, error)):
			with self.assertRaises(entra.ProvisioningError):
				entra.AddToGroup(USER_ID, GROUP_ONE, entra.MODE_LIVE)


	def testRemoveFromGroupGoneIsSuccess(self) -> None:
		with patch.object(entra.requests, "request", return_value=self.response(404, {})):
			entra.RemoveFromGroup(USER_ID, GROUP_ONE, entra.MODE_LIVE)


	def testBadGuidNeverReachesUrl(self) -> None:
		with patch.object(entra.requests, "request") as request:
			with self.assertRaises(entra.ProvisioningError):
				entra.AddToGroup(USER_ID, "../../users", entra.MODE_LIVE)
		request.assert_not_called()


	def testCreateUserBodyAndPassword(self) -> None:
		createdResponse = self.response(201, {"id": USER_ID})
		with patch.object(entra.requests, "request", return_value=createdResponse) as request:
			created = entra.CreateUser(UPN, PERSON, entra.MODE_LIVE)
		body = request.call_args.kwargs["json"]
		self.assertEqual(created["id"], USER_ID)
		self.assertEqual(body["passwordProfile"]["password"], created["password"])
		self.assertTrue(body["passwordProfile"]["forceChangePasswordNextSignIn"])
		self.assertEqual(body["mailNickname"], "jane.smith")


	def testPasswordStrength(self) -> None:
		for _attempt in range(20):
			password = entra.GenerateTemporaryPassword()
			self.assertGreaterEqual(len(password), 16)
			for pattern in (r"[A-Z]", r"[a-z]", r"[0-9]", r"[^A-Za-z0-9]"):
				self.assertTrue(re.search(pattern, password))


if __name__ == "__main__":
	unittest.main()
