# Copyright (c) 2025, Zeke Tierney and contributors
# For license information, please see license.txt

"""
Patch to add M365 custom fields to Email Queue
"""

import frappe
from frappe.custom.doctype.custom_field.custom_field import create_custom_fields


def execute():
	"""
	Add M365 custom fields to Email Queue
	"""
	custom_fields = {
		"Email Queue": [
			{
				"fieldname": "m365_send",
				"label": "Send via M365",
				"fieldtype": "Check",
				"insert_after": "status",
				"read_only": 1,
				"no_copy": 1,
				"default": "0"
			},
			{
				"fieldname": "m365_account",
				"label": "M365 Account",
				"fieldtype": "Link",
				"options": "Email Account",
				"insert_after": "m365_send",
				"read_only": 1,
				"no_copy": 1
			}
		]
	}
	
	create_custom_fields(custom_fields, update=True)
	frappe.db.commit()
	
	print("✅ Added M365 custom fields to Email Queue")

