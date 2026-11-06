# Copyright (c) 2026, TierneyMorris Pty Ltd and contributors
# For license information, please see license.txt

import re

import frappe
from frappe import _
from frappe.model.document import Document

GUID_PATTERN = re.compile(r"[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}")
DOMAIN_PATTERN = re.compile(
	r"[a-z0-9]([a-z0-9-]*[a-z0-9])?"
	r"(\.[a-z0-9]([a-z0-9-]*[a-z0-9])?)+"
)
COUNTRY_CODE_PATTERN = re.compile(r"[A-Z]{2}")
TOKEN_CACHE_KEY = "m365_provisioning_token"
MODE_LIVE = "Live"


class M365ProvisioningSettings(Document):
	"""Own credentials and mode for Entra account provisioning (separate from the mail app)."""

	def validate(self) -> None:
		for fieldName in ("tenant_id", "client_id"):
			value = (self.get(fieldName) or "").strip()
			self.set(fieldName, value)
			if value and not GUID_PATTERN.fullmatch(value):
				message = _("{0} must be a GUID such as 00000000-0000-0000-0000-000000000000.")
				frappe.throw(message.format(self.meta.get_label(fieldName)))

		domain = (self.default_domain or "").strip().lstrip("@").lower()
		self.default_domain = domain
		if domain and not DOMAIN_PATTERN.fullmatch(domain):
			frappe.throw(_("Default Domain must look like example.com.au."))

		location = (self.usage_location or "").strip().upper()
		self.usage_location = location
		if location and not COUNTRY_CODE_PATTERN.fullmatch(location):
			frappe.throw(_("Usage Location must be a two-letter country code such as AU."))

		hasLiveSettings = self.tenant_id and self.client_id and self.default_domain
		if self.mode == MODE_LIVE and not hasLiveSettings:
			frappe.throw(_("Live mode needs the Tenant ID, Client ID and Default Domain."))


	def on_update(self) -> None:
		# NOTE: A changed credential must never keep serving the old token.
		frappe.cache().delete_value(TOKEN_CACHE_KEY)
