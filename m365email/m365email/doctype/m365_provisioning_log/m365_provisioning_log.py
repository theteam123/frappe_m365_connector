# Copyright (c) 2026, TierneyMorris Pty Ltd and contributors
# For license information, please see license.txt

from frappe.model.document import Document

TITLE_SEPARATOR = " \u00b7 "


class M365ProvisioningLog(Document):
	"""One row per Microsoft 365 provisioning step. Written by server code only."""

	def before_insert(self) -> None:
		if not self.title:
			self.title = f"{self.action}{TITLE_SEPARATOR}{self.user_principal_name or '-'}"
