# Copyright (c) 2026, TierneyMorris Pty Ltd and contributors
# For license information, please see license.txt

import re

import frappe
from frappe import _
from frappe.model.document import Document

GUID_PATTERN = re.compile(r"[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}")
SYSTEM_MANAGER_ROLE = "System Manager"
ADMINISTRATOR = "Administrator"
SETUP_CONTEXT_FLAGS = ("in_install", "in_migrate", "in_test")
EDIT_REFUSED_MESSAGE = (
	"Only a System Manager can change a Microsoft 365 provisioning profile: it decides "
	"which licence and groups a new account receives."
)


class M365ProvisioningProfile(Document):
	"""A licence plus a list of groups to give a new Microsoft 365 account."""

	def validate(self) -> None:
		self.guardEditor()
		skuID = (self.licence_sku_id or "").strip()
		self.licence_sku_id = skuID
		if skuID and not GUID_PATTERN.fullmatch(skuID):
			frappe.throw(_("Licence SKU ID must be a GUID."))

		seenGroups = set()
		for row in self.groups:
			row.group_id = (row.group_id or "").strip()
			if not GUID_PATTERN.fullmatch(row.group_id):
				message = _("Row {0}: the group ID for {1} must be a GUID.")
				frappe.throw(message.format(row.idx, row.group_name))
			if row.group_id.lower() in seenGroups:
				frappe.throw(_("Row {0}: {1} is listed twice.").format(row.idx, row.group_name))
			seenGroups.add(row.group_id.lower())


	def guardEditor(self) -> None:
		if any(frappe.flags.get(flagName) for flagName in SETUP_CONTEXT_FLAGS):
			return
		if frappe.session.user == ADMINISTRATOR:
			return
		if SYSTEM_MANAGER_ROLE in frappe.get_roles():
			return
		frappe.throw(_(EDIT_REFUSED_MESSAGE), frappe.PermissionError)
