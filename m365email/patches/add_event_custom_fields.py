# Copyright (c) 2025, Zeke Tierney and contributors
# For license information, please see license.txt

"""
Migration patch to add custom fields to Event doctype for M365 calendar sync
"""

import frappe
from frappe.custom.doctype.custom_field.custom_field import create_custom_fields


def execute():
	"""
	Add custom fields to Event doctype for M365 integration
	"""
	custom_fields = {
		"Event": [
			{
				"fieldname": "m365_event_id",
				"label": "M365 Event ID",
				"fieldtype": "Data",
				"length": 500,
				"insert_after": "event_type",
				"read_only": 1,
				"no_copy": 1,
				"hidden": 1,
				"unique": 1
			},
			{
				"fieldname": "m365_email_account",
				"label": "M365 Email Account",
				"fieldtype": "Link",
				"options": "Email Account",
				"insert_after": "m365_event_id",
				"read_only": 1,
				"no_copy": 1,
				"hidden": 1
			},
			{
				"fieldname": "m365_icaluid",
				"label": "M365 iCalUId",
				"fieldtype": "Data",
				"length": 500,
				"insert_after": "m365_email_account",
				"read_only": 1,
				"no_copy": 1,
				"hidden": 1
			}
		]
	}
	
	create_custom_fields(custom_fields, update=True)
	frappe.db.commit()
	
	print("M365 Email: Successfully added custom fields to Event doctype")

