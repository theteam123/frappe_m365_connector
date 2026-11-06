# Copyright (c) 2026, TierneyMorris Pty Ltd and contributors
# For license information, please see license.txt

"""Microsoft 365 (Entra) account provisioning engine.

Creates an account, gives it a licence and adds it to groups; and the reverse (disable,
remove from the groups the ERP added, release the licence). It knows nothing about HR: a
caller passes names and a reference record.

It uses its OWN app registration and settings ("M365 Provisioning Settings"), deliberately
separate from the mail integration's credentials.

Three modes, read from the settings on every call:
	Off         every action is refused with a clear error.
	Simulation  no network call at all; deterministic fake answers; every step is still
	            logged (mode Simulation), so the whole flow can be rehearsed safely.
	Live        real Microsoft Graph v1.0 calls.

NOTE: The temporary password for a new account is generated here, handed back to the caller
exactly once in the result, and is never logged, stored or written into an error message.
"""

import re
import secrets
import string
import time
import unicodedata
import uuid
from dataclasses import dataclass, field
from urllib.parse import quote

import frappe
import msal
import requests
from frappe import _

SETTINGS_DOCTYPE = "M365 Provisioning Settings"
PROFILE_DOCTYPE = "M365 Provisioning Profile"
LOG_DOCTYPE = "M365 Provisioning Log"

MODE_OFF = "Off"
MODE_SIMULATION = "Simulation"
MODE_LIVE = "Live"

ACTION_CREATE = "Create Account"
ACTION_LICENCE = "Assign Licence"
ACTION_ADD_GROUP = "Add Group"
ACTION_REMOVE_GROUP = "Remove Group"
ACTION_DISABLE = "Disable Account"
ACTION_RELEASE = "Release Licence"
ACTION_LOOKUP = "Lookup"

STATUS_SUCCESS = "Success"
STATUS_FAILED = "Failed"
STATUS_SKIPPED = "Skipped"

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
GRAPH_SCOPE = "https://graph.microsoft.com/.default"
AUTHORITY_TEMPLATE = "https://login.microsoftonline.com/{0}"
TOKEN_CACHE_KEY = "m365_provisioning_token"
TOKEN_EXPIRY_MARGIN_SECONDS = 300
DEFAULT_TOKEN_LIFETIME_SECONDS = 3600
REQUEST_TIMEOUT_SECONDS = 30
MAX_RETRY_AFTER_SECONDS = 30
DEFAULT_RETRY_AFTER_SECONDS = 1
MAX_REQUEST_ATTEMPTS = 2
MIN_CACHE_SECONDS = 1
MAX_ERROR_CHARACTERS = 300
MAX_DETAIL_CHARACTERS = 500
HTTP_TOO_MANY_REQUESTS = 429
HTTP_NOT_FOUND = 404
HTTP_BAD_REQUEST = 400
HTTP_ERROR_FLOOR = 400
USER_FIELDS = "id,displayName,accountEnabled,assignedLicenses,userPrincipalName"

SIMULATION_FREE_SEATS = 5
SIMULATION_ID_PREFIX = "sim-"
SIMULATION_PASSWORD = "SIMULATION-NOT-A-REAL-PASSWORD"
PASSWORD_LENGTH = 18
MAX_ADDRESS_ATTEMPTS = 50
ADDRESS_SUFFIX_START = 2
ALREADY_MEMBER_MARKER = "already exist"
ALREADY_MEMBER_DETAIL = "already a member"
LICENCE_ALREADY_DETAIL = "licence was already on the account"
LOG_WRITE_WARNING = (
	"A step log row could not be written after the account was created. The account exists "
	"in Microsoft 365; ask a System Manager to check the M365 Provisioning Log."
)
LOG_WRITE_DETAIL = "Done in Microsoft 365, but its log row could not be written."
POST_CREATE_WARNING = (
	"A step after the account was created stopped unexpectedly. The account exists in "
	"Microsoft 365; run Create again to finish the remaining steps."
)
FINISH_SETUP_LABEL = "Finish setup"
UNEXPECTED_ERROR_MESSAGE = "An unexpected error stopped the Microsoft 365 step. Nothing was logged."
HIDDEN_SECRET_TEXT = "<hidden>"

GUID_PATTERN = re.compile(r"[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}")
EMAIL_PATTERN = re.compile(
	r"[A-Za-z0-9][A-Za-z0-9._-]{0,63}@[A-Za-z0-9]([A-Za-z0-9-]*[A-Za-z0-9])?"
	r"(\.[A-Za-z0-9]([A-Za-z0-9-]*[A-Za-z0-9])?)+"
)
NAME_STRIP_PATTERN = re.compile(r"[^a-z-]")



class ProvisioningError(Exception):
	"""A step failed. The message is safe to log and show: it never holds a secret."""



class SecretValue:
	"""Holds a secret so a traceback that prints local variables shows only a placeholder."""

	__slots__ = ("_value",)

	def __init__(self, value: str) -> None:
		self._value = value


	def __repr__(self) -> str:
		return HIDDEN_SECRET_TEXT


	def __str__(self) -> str:
		return HIDDEN_SECRET_TEXT


	def __bool__(self) -> bool:
		return bool(self._value)


	def reveal(self) -> str:
		return self._value



@dataclass
class StepContext:
	"""What every step of one run shares; each step is logged against it."""

	mode: str
	userPrincipalName: str
	referenceDoctype: str = ""
	referenceName: str = ""
	userID: str = ""
	steps: list = field(default_factory=list)
	warnings: list = field(default_factory=list)



# ---------------------------------------------------------------- settings and mode

def GetSettings() -> object:
	return frappe.get_single(SETTINGS_DOCTYPE)



def GetMode() -> str:
	"""The active mode. Off refuses here, so no code path can act while switched off."""
	mode = GetSettings().mode or MODE_OFF
	if mode == MODE_OFF:
		frappe.throw(_("Microsoft 365 account provisioning is switched Off in the settings."))
	return mode



def RequireMode(mode: str | None) -> str:
	"""A caller-supplied mode, checked to be an active one, or the settings mode if none given.

	Off refuses either way, so no code path can act while switched off."""
	if mode is None:
		return GetMode()
	if mode not in (MODE_SIMULATION, MODE_LIVE):
		frappe.throw(_("Microsoft 365 account provisioning is switched Off in the settings."))
	return mode



def ValidateGUID(value: str, label: str) -> str:
	if not GUID_PATTERN.fullmatch(value or ""):
		raise ProvisioningError(f"{label} is not a valid ID.")
	return value



def ValidateAddress(userPrincipalName: str) -> str:
	if not EMAIL_PATTERN.fullmatch(userPrincipalName or ""):
		raise ProvisioningError("That is not a valid sign-in address.")
	return userPrincipalName.lower()



# ---------------------------------------------------------------- token and HTTP

def GetProvisioningToken() -> str:
	"""Client-credentials token for the provisioning app registration (cached till near expiry)."""
	cached = frappe.cache().get_value(TOKEN_CACHE_KEY)
	if cached:
		return cached

	settings = GetSettings()
	if not (settings.tenant_id and settings.client_id):
		raise ProvisioningError("Tenant ID and Client ID are not set in the provisioning settings.")

	try:
		clientApp = msal.ConfidentialClientApplication(
			client_id=settings.client_id,
			client_credential=settings.get_password("client_secret"),
			authority=AUTHORITY_TEMPLATE.format(settings.tenant_id),
		)
		result = clientApp.acquire_token_for_client(scopes=[GRAPH_SCOPE])
	except Exception:
		# NOTE: Dropped before raising so the frame a traceback prints holds no client secret.
		clientApp = None
		settings = None
		raise ProvisioningError("Could not get a Microsoft token.") from None
	token = result.get("access_token")
	if not token:
		# NOTE: Only the short error code, never the full reply.
		errorCode = str(result.get("error", "unknown"))[:MAX_ERROR_CHARACTERS]
		raise ProvisioningError(f"Could not get a Microsoft token ({errorCode}).")

	lifetime = int(result.get("expires_in", DEFAULT_TOKEN_LIFETIME_SECONDS))
	cacheSeconds = max(lifetime - TOKEN_EXPIRY_MARGIN_SECONDS, MIN_CACHE_SECONDS)
	frappe.cache().set_value(TOKEN_CACHE_KEY, token, expires_in_sec=cacheSeconds)
	return token



def ExtractGraphError(response: requests.Response) -> str:
	try:
		error = response.json().get("error", {})
	except ValueError:
		error = {}
	code = error.get("code", "")
	message = error.get("message") or "no detail"
	return f"Microsoft Graph {response.status_code} {code}: {message}"[:MAX_ERROR_CHARACTERS]



def ParseRetryAfter(rawValue: object) -> int:
	"""Seconds to wait from a Retry-After header; anything not whole seconds gives the default."""
	try:
		seconds = int(str(rawValue).strip())
	except (TypeError, ValueError):
		return DEFAULT_RETRY_AFTER_SECONDS
	return max(0, min(seconds, MAX_RETRY_AFTER_SECONDS))



def SendGraphRequest(
	method: str,
	path: str,
	body: dict | None,
	params: dict | None,
	tolerated: tuple,
) -> tuple:
	"""The HTTP work for GraphRequest. Its locals (token, body) are blanked on the way out."""
	url = path if path.startswith(GRAPH_BASE_URL) else f"{GRAPH_BASE_URL}{path}"
	headers = {
		"Authorization": f"Bearer {GetProvisioningToken()}",
		"Content-Type": "application/json",
	}
	response = None
	try:
		for attempt in range(MAX_REQUEST_ATTEMPTS):
			try:
				response = requests.request(
					method,
					url,
					headers=headers,
					json=body,
					params=params,
					timeout=REQUEST_TIMEOUT_SECONDS,
				)
			except requests.exceptions.RequestException as error:
				errorName = type(error).__name__
				raise ProvisioningError(f"Could not reach Microsoft Graph ({errorName}).") from None

			if response.status_code == HTTP_TOO_MANY_REQUESTS and attempt == 0:
				time.sleep(ParseRetryAfter(response.headers.get("Retry-After")))
				continue
			break

		if response.status_code in tolerated:
			return response.status_code, {}
		if response.status_code >= HTTP_ERROR_FLOOR:
			raise ProvisioningError(ExtractGraphError(response))

		payload = response.json() if response.content else {}
		return response.status_code, payload
	finally:
		# NOTE: The request object on a response also carries the Authorization header.
		headers = None
		body = None
		response = None



def GraphRequest(
	method: str,
	path: str,
	body: dict | None = None,
	params: dict | None = None,
	tolerated: tuple = (),
) -> tuple:
	"""One Graph call. Returns (status code, parsed JSON). Statuses in `tolerated` are returned
	rather than raised; any other 4xx/5xx raises ProvisioningError (one retry on throttling).
	Any unexpected failure becomes a ProvisioningError with a fixed message: the original is
	dropped so a logged traceback cannot show the token or the request body."""
	try:
		return SendGraphRequest(method, path, body, params, tolerated)
	except ProvisioningError:
		raise
	except Exception:
		body = None
		raise ProvisioningError(UNEXPECTED_ERROR_MESSAGE) from None



def GraphGetAll(path: str, params: dict | None = None) -> list:
	"""Every item of a paged collection, following @odata.nextLink (Graph URLs only)."""
	items: list = []
	nextUrl = path
	nextParams = params
	while nextUrl:
		if nextUrl.startswith("http") and not nextUrl.startswith(GRAPH_BASE_URL):
			raise ProvisioningError("Microsoft returned an unexpected paging link.")
		_status, payload = GraphRequest("GET", nextUrl, params=nextParams)
		items.extend(payload.get("value", []))
		nextUrl = payload.get("@odata.nextLink")
		nextParams = None
	return items



# ---------------------------------------------------------------- primitives

def FindUser(userPrincipalName: str, mode: str) -> dict | None:
	"""The user with this sign-in address (id, displayName, accountEnabled, licences) or None."""
	address = ValidateAddress(userPrincipalName)
	if mode == MODE_SIMULATION:
		return None

	path = f"/users/{quote(address, safe='@')}"
	query = {"$select": USER_FIELDS}
	status, payload = GraphRequest("GET", path, params=query, tolerated=(HTTP_NOT_FOUND,))
	if status == HTTP_NOT_FOUND:
		return None
	fieldNames = [name for name in USER_FIELDS.split(",") if name != "userPrincipalName"]
	return {name: payload.get(name) for name in fieldNames}



def SanitiseNamePart(rawName: str) -> str:
	"""ASCII letters and hyphens only, lower case; apostrophes, spaces and accents removed."""
	normalised = unicodedata.normalize("NFKD", rawName or "")
	asciiName = normalised.encode("ascii", "ignore").decode("ascii")
	return NAME_STRIP_PATTERN.sub("", asciiName.lower())



def IsAddressAvailable(userPrincipalName: str, mode: str) -> bool:
	return FindUser(userPrincipalName, mode) is None



def ProposeAddress(firstName: str, lastName: str, domain: str, mode: str) -> str:
	"""firstname.lastname@domain, then firstname.lastname2@domain ... the first one free."""
	first = SanitiseNamePart(firstName)
	last = SanitiseNamePart(lastName)
	cleanDomain = (domain or "").strip().lstrip("@").lower()
	if not first or not last or not cleanDomain:
		raise ProvisioningError("A first name, last name and domain are needed for an address.")

	baseName = f"{first}.{last}"
	lastSuffix = ADDRESS_SUFFIX_START + MAX_ADDRESS_ATTEMPTS
	suffixes = [""] + [str(number) for number in range(ADDRESS_SUFFIX_START, lastSuffix)]
	for suffix in suffixes:
		candidate = f"{baseName}{suffix}@{cleanDomain}"
		if IsAddressAvailable(candidate, mode):
			return candidate
	raise ProvisioningError("No free sign-in address could be found for that name.")



def ListSubscribedSkus(mode: str) -> list:
	if mode == MODE_SIMULATION:
		return []
	return GraphGetAll("/subscribedSkus")



def GetLicenceAvailability(skuID: str, mode: str) -> int:
	"""Free seats for a licence: purchased (enabled) minus consumed."""
	if mode == MODE_SIMULATION:
		return SIMULATION_FREE_SEATS

	ValidateGUID(skuID, "The licence")
	for sku in ListSubscribedSkus(mode):
		if str(sku.get("skuId", "")).lower() != skuID.lower():
			continue
		purchased = int(sku.get("prepaidUnits", {}).get("enabled", 0))
		return purchased - int(sku.get("consumedUnits", 0))
	raise ProvisioningError("That licence was not found in the Microsoft 365 tenant.")



def GenerateTemporaryPassword() -> str:
	"""A random password with upper, lower, digit and symbol characters."""
	symbols = "!@#$%^&*-_=+"
	pools = [string.ascii_uppercase, string.ascii_lowercase, string.digits, symbols]
	characters = [secrets.choice(pool) for pool in pools]
	allCharacters = "".join(pools)
	fillCount = PASSWORD_LENGTH - len(characters)
	characters += [secrets.choice(allCharacters) for _index in range(fillCount)]
	secrets.SystemRandom().shuffle(characters)
	return "".join(characters)



def CreateUser(userPrincipalName: str, person: dict, mode: str) -> dict:
	"""Create the account. Returns {"id": ..., "password": ...}; the password is given ONCE."""
	address = ValidateAddress(userPrincipalName)
	if mode == MODE_SIMULATION:
		return {"id": f"{SIMULATION_ID_PREFIX}{uuid.uuid4()}", "password": SIMULATION_PASSWORD}

	settings = GetSettings()
	password = GenerateTemporaryPassword()
	firstName = (person.get("firstName") or "").strip()
	lastName = (person.get("lastName") or "").strip()
	body = {
		"accountEnabled": True,
		"displayName": f"{firstName} {lastName}".strip(),
		"mailNickname": address.split("@")[0],
		"userPrincipalName": address,
		"givenName": firstName,
		"surname": lastName,
		"usageLocation": settings.usage_location,
		"passwordProfile": {
			"password": password,
			"forceChangePasswordNextSignIn": bool(settings.force_password_change),
		},
	}
	for graphKey, personKey in (("jobTitle", "jobTitle"), ("department", "department")):
		if person.get(personKey):
			body[graphKey] = person[personKey]

	try:
		_status, payload = GraphRequest("POST", "/users", body=body)
	except Exception:
		# NOTE: Blanked before the error leaves this frame; the message is fixed text.
		body = None
		password = None
		raise ProvisioningError("Microsoft 365 did not create the account.") from None

	newID = payload.get("id", "")
	if not newID:
		body = None
		password = None
		raise ProvisioningError("Microsoft 365 did not return an ID for the new account.")
	return {"id": newID, "password": password}



def AssignLicence(userID: str, skuID: str, mode: str) -> None:
	if mode == MODE_SIMULATION:
		return
	validSku = ValidateGUID(skuID, "The licence")
	validUser = ValidateGUID(userID, "The user")
	body = {"addLicenses": [{"skuId": validSku, "disabledPlans": []}], "removeLicenses": []}
	GraphRequest("POST", f"/users/{validUser}/assignLicense", body=body)



def ReleaseLicence(userID: str, skuID: str, mode: str) -> None:
	if mode == MODE_SIMULATION:
		return
	validSku = ValidateGUID(skuID, "The licence")
	validUser = ValidateGUID(userID, "The user")
	body = {"addLicenses": [], "removeLicenses": [validSku]}
	GraphRequest("POST", f"/users/{validUser}/assignLicense", body=body)



def AddToGroup(userID: str, groupID: str, mode: str) -> str:
	"""Add a member. Returns "" when added, or ALREADY_MEMBER_DETAIL when they already were,
	so the caller can log that as Skipped: the ERP did not add it and must not remove it."""
	if mode == MODE_SIMULATION:
		return ""
	ValidateGUID(userID, "The user")
	ValidateGUID(groupID, "The group")
	body = {"@odata.id": f"{GRAPH_BASE_URL}/directoryObjects/{userID}"}
	try:
		GraphRequest("POST", f"/groups/{groupID}/members/$ref", body=body)
	except ProvisioningError as error:
		if ALREADY_MEMBER_MARKER not in str(error).lower():
			raise
		return ALREADY_MEMBER_DETAIL
	return ""



def RemoveFromGroup(userID: str, groupID: str, mode: str) -> None:
	"""Remove a member. Already gone (404) counts as success."""
	if mode == MODE_SIMULATION:
		return
	ValidateGUID(userID, "The user")
	ValidateGUID(groupID, "The group")
	path = f"/groups/{groupID}/members/{userID}/$ref"
	GraphRequest("DELETE", path, tolerated=(HTTP_NOT_FOUND,))



def DisableUser(userID: str, mode: str) -> None:
	if mode == MODE_SIMULATION:
		return
	validUser = ValidateGUID(userID, "The user")
	GraphRequest("PATCH", f"/users/{validUser}", body={"accountEnabled": False})



# ---------------------------------------------------------------- log

def WriteLog(
	context: StepContext,
	action: str,
	status: str,
	detail: str,
	targetID: str = "",
	targetName: str = "",
) -> None:
	"""Record one step. `detail` is built from our own text only, and cut short."""
	entry = {
		"doctype": LOG_DOCTYPE,
		"action": action,
		"mode": context.mode,
		"status": status,
		"user_principal_name": context.userPrincipalName,
		"entra_user_id": context.userID,
		"target_id": targetID,
		"target_name": targetName,
		"reference_doctype": context.referenceDoctype or None,
		"reference_name": context.referenceName or None,
		"detail": (detail or "")[:MAX_DETAIL_CHARACTERS],
		"performed_by": frappe.session.user,
	}
	frappe.get_doc(entry).insert(ignore_permissions=True)
	if context.mode == MODE_LIVE:
		# NOTE: Committed straight away so a later failure cannot roll back the record of a
		# real change in Microsoft 365: resuming and undoing both rely on this log.
		frappe.db.commit()



def HasLogged(context: StepContext, action: str, targetID: str = "") -> bool:
	"""True when this run's account (Entra object id), reference and mode already succeeded.

	Keyed on the object id, never the address: an address can be reused by a new account."""
	if not (context.userID and context.referenceDoctype and context.referenceName):
		return False
	filters = {
		"action": action,
		"entra_user_id": context.userID,
		"reference_doctype": context.referenceDoctype,
		"reference_name": context.referenceName,
		"mode": context.mode,
		"status": STATUS_SUCCESS,
	}
	if targetID:
		filters["target_id"] = targetID
	return bool(frappe.get_all(LOG_DOCTYPE, filters=filters, limit=1))



def HasCreateRecord(context: StepContext, userID: str) -> bool:
	"""True when a successful Create row names exactly this account and this reference.

	An old row for the same address proves nothing: the address may belong to someone else now."""
	if not (userID and context.referenceDoctype and context.referenceName):
		return False
	filters = {
		"action": ACTION_CREATE,
		"mode": context.mode,
		"status": STATUS_SUCCESS,
		"entra_user_id": userID,
		"reference_doctype": context.referenceDoctype,
		"reference_name": context.referenceName,
	}
	return bool(frappe.get_all(LOG_DOCTYPE, filters=filters, limit=1))



def GetOpenTargets(
	addAction: str,
	removeAction: str,
	mode: str,
	userID: str,
	createdFor: dict,
) -> list:
	"""[(target id, name)] that the ERP added (a Success row, never a Skipped one) to this
	account for the record that created it, and has not since removed.

	Keyed on the Entra object id, never the address, which a new account may reuse."""
	if not userID:
		return []
	baseFilters = {
		"mode": mode,
		"status": STATUS_SUCCESS,
		"entra_user_id": userID,
	}
	addFilters = dict(
		baseFilters,
		action=addAction,
		reference_doctype=createdFor.get("doctype", ""),
		reference_name=createdFor.get("name", ""),
	)
	removeFilters = dict(baseFilters, action=removeAction)
	added = frappe.get_all(
		LOG_DOCTYPE,
		filters=addFilters,
		fields=["target_id", "target_name"],
		order_by="creation asc",
	)
	removed = frappe.get_all(LOG_DOCTYPE, filters=removeFilters, fields=["target_id"])
	removedIDs = {row["target_id"] for row in removed}

	openTargets: list = []
	openIDs: set = set()
	for row in added:
		if row["target_id"] in removedIDs or row["target_id"] in openIDs:
			continue
		openIDs.add(row["target_id"])
		openTargets.append((row["target_id"], row["target_name"] or ""))
	return openTargets



def TryWriteLog(
	context: StepContext,
	action: str,
	status: str,
	detail: str,
	targetID: str = "",
	targetName: str = "",
) -> bool:
	"""WriteLog, except that once the account exists a failed write becomes a warning.

	Before an account exists the error is raised as before (nothing to protect). Afterwards
	the one-time password must still reach the caller, so the failure is swallowed with a
	fixed warning: never the exception text, which could carry anything."""
	try:
		WriteLog(context, action, status, detail, targetID, targetName)
	except Exception:
		if not context.userID:
			raise
		if LOG_WRITE_WARNING not in context.warnings:
			context.warnings.append(LOG_WRITE_WARNING)
		return False
	return True



def RecordStep(context: StepContext, label: str, status: str, detail: str) -> None:
	context.steps.append({"step": label, "status": status, "detail": detail})



def RunStep(
	context: StepContext,
	label: str,
	action: str,
	work: object,
	targetID: str = "",
	targetName: str = "",
) -> bool:
	"""Run one step, log it, and append its outcome to context.steps. True on success.

	`work` may return a text: that means nothing was changed (it was already so), and the step
	is logged Skipped, which offboarding never undoes."""
	try:
		skipDetail = work()
	except ProvisioningError as error:
		RecordStep(context, label, STATUS_FAILED, str(error))
		TryWriteLog(context, action, STATUS_FAILED, str(error), targetID, targetName)
		return False
	except Exception:
		RecordStep(context, label, STATUS_FAILED, UNEXPECTED_ERROR_MESSAGE)
		TryWriteLog(
			context, action, STATUS_FAILED, UNEXPECTED_ERROR_MESSAGE, targetID, targetName
		)
		return False

	if skipDetail:
		SkipStep(context, label, action, str(skipDetail), targetID, targetName)
		return True

	RecordStep(context, label, STATUS_SUCCESS, "")
	if TryWriteLog(context, action, STATUS_SUCCESS, "", targetID, targetName):
		return True
	# NOTE: The change happened but is unrecorded, so the step is shown as failed and the run
	# stops cleanly; the result (with the wrapped password) still goes back to the caller.
	context.steps[-1].update(status=STATUS_FAILED, detail=LOG_WRITE_DETAIL)
	return False



def SkipStep(
	context: StepContext,
	label: str,
	action: str,
	reason: str,
	targetID: str = "",
	targetName: str = "",
) -> None:
	RecordStep(context, label, STATUS_SKIPPED, reason)
	TryWriteLog(context, action, STATUS_SKIPPED, reason, targetID, targetName)



def BuildResult(
	context: StepContext,
	isResume: bool = False,
	password: SecretValue | None = None,
) -> dict:
	"""The run's outcome. `temporaryPassword` stays a SecretValue (or None): only the final
	caller that hands it to a person may reveal() it."""
	isSuccess = all(step["status"] != STATUS_FAILED for step in context.steps)
	return {
		"success": isSuccess,
		"mode": context.mode,
		"userPrincipalName": context.userPrincipalName,
		"userID": context.userID,
		"steps": context.steps,
		"isResume": isResume,
		"temporaryPassword": password,
		"warnings": list(context.warnings),
	}



# ---------------------------------------------------------------- orchestrators

def LoadProfile(profileName: str) -> dict:
	profile = frappe.get_doc(PROFILE_DOCTYPE, profileName)
	groups = [{"id": row.group_id, "name": row.group_name} for row in profile.groups]
	return {
		"name": profile.name,
		"skuID": profile.licence_sku_id or "",
		"licenceName": profile.licence_name or "",
		"groups": groups,
	}



def LookUpExisting(context: StepContext) -> tuple:
	"""Step 1. Returns (existing user or None, ok). Fails the run if the account is not ours."""
	try:
		existing = FindUser(context.userPrincipalName, context.mode)
	except ProvisioningError as error:
		RecordStep(context, "Look up account", STATUS_FAILED, str(error))
		WriteLog(context, ACTION_LOOKUP, STATUS_FAILED, str(error))
		return None, False

	if existing is None:
		RecordStep(context, "Look up account", STATUS_SUCCESS, "Address is free")
		WriteLog(context, ACTION_LOOKUP, STATUS_SUCCESS, "Address is free")
		return None, True

	isOurs = HasCreateRecord(context, existing.get("id", ""))
	if not isOurs:
		reason = (
			"This account already exists in Microsoft 365 and was not created by the ERP "
			"for this employee."
		)
		RecordStep(context, "Look up account", STATUS_FAILED, reason)
		WriteLog(context, ACTION_LOOKUP, STATUS_FAILED, reason)
		return existing, False

	resumeNote = "Created earlier by the ERP; finishing the remaining steps"
	RecordStep(context, "Look up account", STATUS_SUCCESS, resumeNote)
	WriteLog(context, ACTION_LOOKUP, STATUS_SUCCESS, "Resuming an account the ERP created")
	return existing, True



def CheckLicenceSeat(context: StepContext, profile: dict) -> bool:
	"""The guard that runs before anything is created: no free seat means stop."""
	skuID = profile["skuID"]
	licenceName = profile["licenceName"]
	try:
		freeSeats = GetLicenceAvailability(skuID, context.mode)
	except ProvisioningError as error:
		RecordStep(context, "Check licence seats", STATUS_FAILED, str(error))
		WriteLog(context, ACTION_LICENCE, STATUS_FAILED, str(error), skuID, licenceName)
		return False

	if freeSeats > 0:
		RecordStep(context, "Check licence seats", STATUS_SUCCESS, f"{freeSeats} free")
		return True

	reason = f"No free {licenceName or 'licence'} seats. Nothing was created."
	RecordStep(context, "Check licence seats", STATUS_FAILED, reason)
	WriteLog(context, ACTION_LICENCE, STATUS_FAILED, reason, skuID, licenceName)
	return False



def AddProfileGroups(context: StepContext, profile: dict) -> bool:
	for group in profile["groups"]:
		label = f"Add to group: {group['name']}"
		if HasLogged(context, ACTION_ADD_GROUP, group["id"]):
			SkipStep(context, label, ACTION_ADD_GROUP, "Already added", group["id"], group["name"])
			continue

		work = lambda groupID=group["id"]: AddToGroup(context.userID, groupID, context.mode)
		if not RunStep(context, label, ACTION_ADD_GROUP, work, group["id"], group["name"]):
			return False
	return True



def ProvisionAccount(
	userPrincipalName: str,
	person: dict,
	profileName: str,
	reference: dict,
	mode: str | None = None,
) -> dict:
	"""Create (or finish creating) an account: look up, licence guard, create, licence, groups.

	`person` holds firstName, lastName, jobTitle, department; `reference` holds doctype and
	name of the record that asked. A failure stops the run, is logged, and the result lists
	exactly which steps worked. Nothing is ever deleted automatically; run it again to resume.
	An unexpected error leaves as a ProvisioningError with fixed text and no original cause.
	"""
	try:
		return RunProvision(userPrincipalName, person, profileName, reference, mode)
	except (ProvisioningError, frappe.ValidationError, frappe.PermissionError):
		raise
	except Exception:
		raise ProvisioningError(UNEXPECTED_ERROR_MESSAGE) from None



def LogExistingLicence(context: StepContext, profile: dict) -> None:
	"""A licence already on the account that the ERP never assigned is logged Skipped, so
	offboarding (which only releases Success rows) leaves it alone."""
	skuID = profile["skuID"]
	if HasLogged(context, ACTION_LICENCE, skuID):
		return
	label = f"Assign licence: {profile['licenceName']}"
	licenceName = profile["licenceName"]
	SkipStep(context, label, ACTION_LICENCE, LICENCE_ALREADY_DETAIL, skuID, licenceName)



def FinishProvision(
	context: StepContext,
	profile: dict,
	hasLicenceAlready: bool,
	needsLicence: bool,
) -> None:
	"""The steps after the account exists: note an existing licence, assign it, add groups."""
	if hasLicenceAlready:
		LogExistingLicence(context, profile)
	if needsLicence:
		work = lambda: AssignLicence(context.userID, profile["skuID"], context.mode)
		licenceName = profile["licenceName"]
		label = f"Assign licence: {licenceName}"
		if not RunStep(context, label, ACTION_LICENCE, work, profile["skuID"], licenceName):
			return
	AddProfileGroups(context, profile)



def RunProvision(
	userPrincipalName: str,
	person: dict,
	profileName: str,
	reference: dict,
	mode: str | None = None,
) -> dict:
	# NOTE: The mode is read once (or handed in by the caller, who read it once) and carried
	# on the context, so a settings change mid-run can never make a Live call under
	# Simulation rows.
	mode = RequireMode(mode)
	address = ValidateAddress(userPrincipalName)
	profile = LoadProfile(profileName)
	context = StepContext(mode, address, reference.get("doctype", ""), reference.get("name", ""))

	existing, isOk = LookUpExisting(context)
	if not isOk:
		return BuildResult(context)

	isResume = existing is not None
	hasSku = bool(profile["skuID"])
	assignedLicences = (existing or {}).get("assignedLicenses") or []
	assigned = {str(item.get("skuId", "")).lower() for item in assignedLicences}
	needsLicence = hasSku and profile["skuID"].lower() not in assigned

	if needsLicence and not CheckLicenceSeat(context, profile):
		return BuildResult(context, isResume)

	# NOTE: The temporary password lives only inside a SecretValue, never in a plain variable
	# of this frame, which goes on to make calls that can raise.
	holder: dict = {}
	if isResume:
		context.userID = existing.get("id", "")
	else:
		def createWork() -> None:
			created = CreateUser(address, person, context.mode)
			# NOTE: The ID is set before the step is logged, so the Create row records it.
			context.userID = created.get("id", "")
			holder["secret"] = SecretValue(created.get("password", ""))
			created = None

		if not RunStep(context, "Create account", ACTION_CREATE, createWork):
			# NOTE: If only the log row failed the account exists, so its password goes back.
			return BuildResult(context, isResume, holder.get("secret"))

	try:
		FinishProvision(context, profile, hasSku and not needsLicence, needsLicence)
	except Exception:
		# NOTE: The account exists, so any error here (validation and permission errors
		# included) is recorded as a failed step and the password still goes back. The
		# failed step makes the result unsuccessful, so the Employee is not bound and a
		# second click resumes. Fixed text only: the exception may carry anything.
		RecordStep(context, FINISH_SETUP_LABEL, STATUS_FAILED, POST_CREATE_WARNING)
		context.warnings.append(POST_CREATE_WARNING)
	return BuildResult(context, isResume, holder.get("secret"))



def RemoveLoggedGroups(context: StepContext, createdFor: dict) -> None:
	groups = GetOpenTargets(
		ACTION_ADD_GROUP, ACTION_REMOVE_GROUP, context.mode, context.userID, createdFor
	)
	for groupID, groupName in groups:
		work = lambda targetID=groupID: RemoveFromGroup(context.userID, targetID, context.mode)
		label = f"Remove from group: {groupName}"
		RunStep(context, label, ACTION_REMOVE_GROUP, work, groupID, groupName)



def ReleaseLoggedLicences(context: StepContext, createdFor: dict) -> None:
	licences = GetOpenTargets(
		ACTION_LICENCE, ACTION_RELEASE, context.mode, context.userID, createdFor
	)
	for skuID, licenceName in licences:
		work = lambda targetID=skuID: ReleaseLicence(context.userID, targetID, context.mode)
		label = f"Release licence: {licenceName}"
		RunStep(context, label, ACTION_RELEASE, work, skuID, licenceName)



def DeprovisionAccount(
	userPrincipalName: str,
	userID: str,
	reference: dict,
	createdFor: dict | None = None,
	*,
	mode: str,
) -> dict:
	"""Disable the account, then take back only what the ERP gave it (per the log).

	`reference` is the record asking (the offboarding); `createdFor` is the record the account
	was created for (the Employee), whose Success rows say what the ERP added. If disabling
	fails nothing else is attempted; later failures do not stop the remaining steps.

	`mode` is required: the caller reads it ONCE and uses the same value to find the Create
	record, so the mode can never differ within one request. It is never re-read here.
	"""
	try:
		return RunDeprovision(userPrincipalName, userID, reference, createdFor or reference, mode)
	except (ProvisioningError, frappe.ValidationError, frappe.PermissionError):
		raise
	except Exception:
		raise ProvisioningError(UNEXPECTED_ERROR_MESSAGE) from None



def RunDeprovision(
	userPrincipalName: str,
	userID: str,
	reference: dict,
	createdFor: dict,
	mode: str,
) -> dict:
	mode = RequireMode(mode)
	address = ValidateAddress(userPrincipalName)
	referenceDoctype = reference.get("doctype", "")
	referenceName = reference.get("name", "")
	context = StepContext(mode, address, referenceDoctype, referenceName, userID)

	disableWork = lambda: DisableUser(userID, context.mode)
	if not RunStep(context, "Disable account", ACTION_DISABLE, disableWork):
		return BuildResult(context)

	RemoveLoggedGroups(context, createdFor)
	ReleaseLoggedLicences(context, createdFor)
	return BuildResult(context)



# ---------------------------------------------------------------- settings page test

@frappe.whitelist()
def TestConnection() -> dict:
	"""System Manager only. Reports counts, never any tenant detail or secret."""
	frappe.only_for("System Manager")
	mode = GetMode()
	if mode == MODE_SIMULATION:
		note = _("Simulation mode: no connection was made. A real test needs Live mode.")
		return {"ok": True, "message": note}

	try:
		skus = ListSubscribedSkus(mode)
	except ProvisioningError as error:
		return {"ok": False, "message": str(error)}
	except Exception:
		return {"ok": False, "message": UNEXPECTED_ERROR_MESSAGE}
	note = _("Connected. {0} licence types are visible to the app.").format(len(skus))
	return {"ok": True, "message": note}
