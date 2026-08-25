# Copyright (c) 2026, TierneyMorris Pty Ltd and contributors
# For license information, please see license.txt

from frappe.model.document import Document


class M365AddinSession(Document):
	"""A remembered Send-to-ERP add-in sign-in. Minted and validated by
	``m365email.m365email.addin_session``; the desk only lists and revokes."""
	pass
