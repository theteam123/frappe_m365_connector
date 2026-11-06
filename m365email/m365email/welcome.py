# Copyright (c) 2026, TierneyMorris Pty Ltd and contributors
# For license information, please see license.txt

"""Welcome email for a new Microsoft 365 account, carrying the temporary password.

The email goes to the person's PERSONAL address, from a shared mailbox (Support), through one
Microsoft Graph `sendMail` call using the existing mail service principal. The provisioning app
registration has no Mail.Send, and its credentials are deliberately kept separate.

NOTE: The password is in the message body, so this path never uses frappe.sendmail, the Email
Queue or a Communication (each would store the body in the ERP), never logs or raises the body,
and asks Graph not to keep a copy in the sender's Sent Items. Failures leave as fixed text.

The recipient is whatever the caller passes, so the caller must pass only the address stored on
the Employee that HR confirmed. A timeout or reset after the request left is reported as
"Unknown" (delivery not confirmed), never as a plain failure.

NOTE: A reply to this email could quote the password into a mailbox that is synced into the ERP
(a Communication). No replyTo is set and the body says not to reply. The password must be
changed at first sign-in, so a quoted copy stops working. The remaining risk is accepted.
"""

from urllib.parse import quote

import requests
from frappe import _
from frappe.utils import escape_html

from m365email.m365email import entra
from m365email.m365email.auth import get_access_token

ACTION_WELCOME = "Send Welcome Email"
DEFAULT_SENDER = "support@sgcaustralia.com.au"
SEND_MAIL_PATH = "/users/{0}/sendMail"
HTTP_ACCEPTED = 202
REQUEST_TIMEOUT_SECONDS = 30
# NOTE: Raised before any byte of the request leaves, so nothing can have been sent.
NOT_SENT_ERRORS = (
	requests.ConnectTimeout,
	requests.exceptions.InvalidURL,
	requests.exceptions.MissingSchema,
	requests.exceptions.InvalidSchema,
)
PORTAL_URL = "https://portal.office.com"
ERP_URL = "https://ops.sgcloud.com.au"

OUTCOME_SENT = "Sent"
OUTCOME_FAILED = "Failed"
OUTCOME_SKIPPED = "Skipped"
OUTCOME_SIMULATED = "Simulated"
OUTCOME_UNKNOWN = "Unknown"
LOG_STATUS_FOR = {
	OUTCOME_SENT: entra.STATUS_SUCCESS,
	OUTCOME_FAILED: entra.STATUS_FAILED,
	OUTCOME_SKIPPED: entra.STATUS_SKIPPED,
	OUTCOME_SIMULATED: entra.STATUS_SKIPPED,
	OUTCOME_UNKNOWN: entra.STATUS_FAILED,
}
STEP_LABEL = "Send welcome email"

SIMULATION_REASON = "Simulation: no welcome email was sent."
NO_PASSWORD_REASON = "There is no new password to send."
NO_EMAIL_REASON = "The employee has no valid personal email address."
NO_PRINCIPAL_REASON = "No mail service principal is set for the welcome email in the settings."
BAD_SENDER_REASON = "The welcome email sender address in the settings is not valid."
TOKEN_FAILED_REASON = "Could not get a Microsoft mail token for the welcome email."
SEND_FAILED_REASON = "Could not reach Microsoft to send the welcome email."
UNKNOWN_REASON = "delivery unknown"
REFUSED_REASON = "Microsoft did not accept the welcome email (HTTP {0})."
UNEXPECTED_REASON = "An unexpected error stopped the welcome email."



class DeliveryUnknownError(entra.ProvisioningError):
	"""The request may have reached Microsoft, so the email may or may not have been sent."""



def BuildWelcomeSubject() -> str:
	return _("Welcome to SGC – your Microsoft 365 account")



def BuildWelcomeBody(
	firstName: str, userPrincipalName: str, password: str, supportAddress: str = DEFAULT_SENDER
) -> str:
	"""The email as simple HTML. Every inserted value is escaped. Edit the wording here."""
	paragraphs = [
		_("Hi {0},").format(escape_html(firstName)),
		_("Welcome to SGC. Your new SGC Microsoft 365 account has been set up."),
		_("Sign-in / email: <b>{0}</b>").format(escape_html(userPrincipalName)),
		_("Temporary password: <b>{0}</b>").format(escape_html(password)),
		_("Sign in at {0}. Use the same account to log in to the SGC ERP at {1}.").format(
			f'<a href="{PORTAL_URL}">{PORTAL_URL}</a>', f'<a href="{ERP_URL}">{ERP_URL}</a>'
		),
		_("You will be asked to choose a new password the first time you sign in."),
		_("If you were not expecting this email, please contact {0}.").format(
			escape_html(supportAddress)
		),
		_(
			"Please do not reply to this email. If you need help, contact {0} in a new email."
		).format(escape_html(supportAddress)),
	]
	return "".join(f"<p>{paragraph}</p>" for paragraph in paragraphs)



def BuildMessage(
	recipient: str,
	firstName: str,
	userPrincipalName: str,
	password: entra.SecretValue,
	supportAddress: str = DEFAULT_SENDER,
) -> dict:
	"""The Graph sendMail payload. saveToSentItems is False so the shared mailbox keeps no copy."""
	body = BuildWelcomeBody(firstName, userPrincipalName, password.reveal(), supportAddress)
	return {
		"message": {
			"subject": BuildWelcomeSubject(),
			"body": {"contentType": "HTML", "content": body},
			"toRecipients": [{"emailAddress": {"address": recipient}}],
		},
		"saveToSentItems": False,
	}



def RecipientDomain(address: str) -> str:
	return (address or "").rpartition("@")[2].lower()



def Deliver(
	settings: object,
	recipient: str,
	firstName: str,
	userPrincipalName: str,
	password: entra.SecretValue,
) -> None:
	"""One Graph sendMail call. Raises ProvisioningError with fixed text only."""
	principal = getattr(settings, "welcome_mail_service_principal", "")
	if not principal:
		raise entra.ProvisioningError(NO_PRINCIPAL_REASON)
	sender = getattr(settings, "welcome_sender", "") or DEFAULT_SENDER
	try:
		sender = entra.ValidateAddress(sender)
	except entra.ProvisioningError:
		raise entra.ProvisioningError(BAD_SENDER_REASON) from None
	try:
		token = get_access_token(principal)
	except Exception:
		raise entra.ProvisioningError(TOKEN_FAILED_REASON) from None

	payload = BuildMessage(recipient, firstName, userPrincipalName, password, sender)
	url = entra.GRAPH_BASE_URL + SEND_MAIL_PATH.format(quote(sender, safe="@"))
	headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
	try:
		response = requests.post(
			url, headers=headers, json=payload, timeout=REQUEST_TIMEOUT_SECONDS
		)
	except NOT_SENT_ERRORS:
		raise entra.ProvisioningError(SEND_FAILED_REASON) from None
	except requests.RequestException:
		# NOTE: Only failures that happen before the request leaves prove nothing was sent.
		# Everything else (read timeout, reset, truncated or undecodable reply) may follow a
		# message Microsoft already accepted, so delivery is unknown.
		raise DeliveryUnknownError(UNKNOWN_REASON) from None
	except Exception:
		# NOTE: Anything else (e.g. a malformed redirect) cannot prove the message was not sent,
		# and its traceback may hold the body, so it is replaced with fixed text.
		raise DeliveryUnknownError(UNKNOWN_REASON) from None
	finally:
		# NOTE: Dropped so a traceback printing locals shows neither the body nor the token.
		payload = None
		token = None
		headers = None
	if response.status_code != HTTP_ACCEPTED:
		raise entra.ProvisioningError(REFUSED_REASON.format(response.status_code))



def Attempt(result: dict, firstName: str, personalEmail: str) -> tuple[str, str]:
	"""(outcome, reason) for one welcome email. Never raises."""
	if result["mode"] == entra.MODE_SIMULATION:
		return OUTCOME_SIMULATED, SIMULATION_REASON
	password = result.get("temporaryPassword")
	if not password:
		return OUTCOME_SKIPPED, NO_PASSWORD_REASON
	try:
		recipient = entra.ValidateAddress((personalEmail or "").strip())
	except entra.ProvisioningError:
		return OUTCOME_SKIPPED, NO_EMAIL_REASON
	try:
		Deliver(entra.GetSettings(), recipient, firstName, result["userPrincipalName"], password)
	except DeliveryUnknownError:
		return OUTCOME_UNKNOWN, UNKNOWN_REASON
	except entra.ProvisioningError as error:
		return OUTCOME_FAILED, str(error)
	except Exception:
		return OUTCOME_FAILED, UNEXPECTED_REASON
	return OUTCOME_SENT, ""



def SendWelcomeEmail(result: dict, firstName: str, personalEmail: str, reference: dict) -> dict:
	"""Send (Live) or rehearse (Simulation) the welcome email, log it, and add it to the steps.

	`result` is the ProvisionAccount result. Returns {outcome, reason, recipient, warnings}.
	The account already exists, so nothing here ever raises or changes result["success"]."""
	outcome, reason = Attempt(result, firstName, personalEmail)
	context = entra.StepContext(
		result["mode"], result["userPrincipalName"],
		reference.get("doctype", ""), reference.get("name", ""), result.get("userID", ""),
	)
	domain = RecipientDomain(personalEmail)
	detail = f"Sent to an address at {domain}" if outcome == OUTCOME_SENT else reason
	entra.TryWriteLog(context, ACTION_WELCOME, LOG_STATUS_FOR[outcome], detail)
	result["steps"].append({
		"step": STEP_LABEL, "status": LOG_STATUS_FOR[outcome], "detail": detail,
	})
	return {
		"outcome": outcome,
		"reason": reason,
		"recipient": (personalEmail or "").strip(),
		"warnings": list(context.warnings),
	}
