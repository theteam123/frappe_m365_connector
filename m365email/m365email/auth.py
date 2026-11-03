# Copyright (c) 2025, Zeke Tierney and contributors
# For license information, please see license.txt

"""
Authentication module for M365 Email Integration
Handles MSAL authentication and token management for multiple tenants
"""

import json
import frappe
from frappe import _
from frappe.utils.password import encrypt, decrypt
from datetime import datetime, timedelta
import msal


def get_msal_app(service_principal_name):
	"""
	Initialize MSAL ConfidentialClientApplication for specified tenant
	
	Args:
		service_principal_name: Name of the M365 Email Service Principal Settings
		
	Returns:
		msal.ConfidentialClientApplication instance
	"""
	# Get service principal settings
	sp_settings = frappe.get_doc("M365 Email Service Principal Settings", service_principal_name)
	
	if not sp_settings.enabled:
		frappe.throw(_("Service Principal {0} is not enabled").format(service_principal_name))
	
	# Initialize MSAL app
	app = msal.ConfidentialClientApplication(
		client_id=sp_settings.client_id,
		client_credential=sp_settings.get_password("client_secret"),
		authority=sp_settings.authority_url,
		token_cache=_get_token_cache(sp_settings)
	)
	
	return app


def get_access_token(service_principal_name, force_refresh=False):
	"""
	Get access token for specified service principal
	Uses cached token if available and not expired, otherwise acquires new token
	
	Args:
		service_principal_name: Name of the M365 Email Service Principal Settings
		force_refresh: Force token refresh even if cached token is valid
		
	Returns:
		str: Access token
	"""
	sp_settings = frappe.get_doc("M365 Email Service Principal Settings", service_principal_name)
	
	# Check if we have a valid cached token
	if not force_refresh and sp_settings.token_expires_at:
		expires_at = frappe.utils.get_datetime(sp_settings.token_expires_at)
		# Add 5 minute buffer before expiration
		if datetime.now() < (expires_at - timedelta(minutes=5)):
			# Token is still valid, try to get from cache
			cache = _get_token_cache(sp_settings)
			accounts = cache.find(msal.TokenCache.CredentialType.ACCESS_TOKEN)
			if accounts:
				return accounts[0]["secret"]
	
	# Need to acquire new token
	app = get_msal_app(service_principal_name)
	
	# Use client credentials flow (application permissions)
	scopes = sp_settings.scopes.split("\n") if "\n" in sp_settings.scopes else [sp_settings.scopes]
	
	result = app.acquire_token_for_client(scopes=scopes)
	
	if "access_token" in result:
		# Update token cache and expiration
		_save_token_cache(sp_settings, app.token_cache)
		
		# Calculate expiration time
		expires_in = result.get("expires_in", 3600)  # Default 1 hour
		expires_at = datetime.now() + timedelta(seconds=expires_in)
		
		sp_settings.db_set("last_token_refresh", datetime.now(), update_modified=False)
		sp_settings.db_set("token_expires_at", expires_at, update_modified=False)
		
		frappe.db.commit()
		
		return result["access_token"]
	else:
		raw_error = result.get("error_description", result.get("error", "Unknown error"))
		error_msg = FlattenTokenErrorValue(raw_error)
		frappe.log_error(
			title=f"M365 Token Acquisition Failed: {service_principal_name}",
			message=f"Error: {error_msg}\nResponse: {SummarizeTokenError(result)}"
		)
		frappe.throw(_("Failed to acquire access token: {0}").format(error_msg))


# NOTE: Only these named fields of a failed sign-in reply reach the Error Log. Logging the
# whole reply would copy whatever Microsoft or MSAL chooses to put in it into a table any
# System Manager can read; these are enough to look the failure up with Microsoft.
LOGGED_TOKEN_ERROR_FIELDS = (
	"error",
	"error_description",
	"error_codes",
	"timestamp",
	"trace_id",
	"correlation_id",
)



# NOTE: Each field is also flattened to plain text and numbers and cut to length, so a field
# that arrives holding a nested block of data cannot carry that block into the log.
MAX_LOGGED_FIELD_CHARACTERS = 1000
LOGGED_SCALAR_TYPES = (str, int, float)
OMITTED_VALUE = "<omitted: not plain text>"
JSON_INDENT = 2



def SummarizeTokenError(result: dict) -> str:
	"""The diagnostic fields of a failed MSAL token reply, as JSON, and nothing else."""
	summary = {}
	for field in LOGGED_TOKEN_ERROR_FIELDS:
		if field in result:
			summary[field] = FlattenTokenErrorValue(result[field])
	return json.dumps(summary, indent=JSON_INDENT)



def FlattenTokenErrorValue(value: object) -> str | list:
	"""Plain text (or a list of plain text) cut to length; anything nested is omitted."""
	if isinstance(value, list):
		return [FlattenScalar(item) for item in value]
	return FlattenScalar(value)



def FlattenScalar(value: object) -> str:
	if isinstance(value, bool) or not isinstance(value, LOGGED_SCALAR_TYPES):
		return OMITTED_VALUE
	return str(value)[:MAX_LOGGED_FIELD_CHARACTERS]



def get_service_principal_for_email_account(email_account_name):
	"""
	Get the service principal associated with an email account

	Args:
		email_account_name: Name of Email Account with service='M365'

	Returns:
		M365 Email Service Principal Settings doc
	"""
	email_account = frappe.get_doc("Email Account", email_account_name)
	if email_account.service != "M365":
		frappe.throw(_("Not an M365 email account"))
	return frappe.get_doc("M365 Email Service Principal Settings", email_account.m365_service_principal)


def refresh_token(service_principal_name):
	"""
	Refresh access token for specified service principal
	
	Args:
		service_principal_name: Name of the M365 Email Service Principal Settings
		
	Returns:
		bool: True if refresh successful, False otherwise
	"""
	try:
		get_access_token(service_principal_name, force_refresh=True)
		return True
	except Exception as e:
		frappe.log_error(
			title=f"Token Refresh Failed: {service_principal_name}",
			message=str(e)
		)
		return False


def _get_token_cache(sp_settings):
	"""
	Get MSAL token cache from service principal settings
	
	Args:
		sp_settings: M365 Email Service Principal Settings doc
		
	Returns:
		msal.SerializableTokenCache instance
	"""
	cache = msal.SerializableTokenCache()
	
	if sp_settings.token_cache:
		try:
			# Decrypt and deserialize token cache
			decrypted_cache = decrypt(sp_settings.token_cache)
			cache.deserialize(decrypted_cache)
		except Exception as e:
			frappe.log_error(
				title=f"Token Cache Deserialization Failed: {sp_settings.name}",
				message=str(e)
			)
	
	return cache


def _save_token_cache(sp_settings, token_cache):
	"""
	Save MSAL token cache to service principal settings
	
	Args:
		sp_settings: M365 Email Service Principal Settings doc
		token_cache: msal.SerializableTokenCache instance
	"""
	if token_cache.has_state_changed:
		try:
			# Serialize and encrypt token cache
			serialized_cache = token_cache.serialize()
			encrypted_cache = encrypt(serialized_cache)

			sp_settings.db_set("token_cache", encrypted_cache, update_modified=False)
		except Exception as e:
			frappe.log_error(
				title=f"Token Cache Serialization Failed: {sp_settings.name}",
				message=str(e)
			)


def test_connection(service_principal_name):
	"""
	Test service principal connection by attempting to acquire token
	
	Args:
		service_principal_name: Name of the M365 Email Service Principal Settings
		
	Returns:
		dict: {"success": bool, "message": str, "token_expires_at": datetime}
	"""
	try:
		token = get_access_token(service_principal_name, force_refresh=True)
		sp_settings = frappe.get_doc("M365 Email Service Principal Settings", service_principal_name)
		
		return {
			"success": True,
			"message": _("Connection successful! Token acquired."),
			"token_expires_at": sp_settings.token_expires_at
		}
	except Exception as e:
		return {
			"success": False,
			"message": str(e),
			"token_expires_at": None
		}

