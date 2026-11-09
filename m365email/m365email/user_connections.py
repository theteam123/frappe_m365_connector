# Copyright (c) 2026, TierneyMorris Pty Ltd and contributors
# For license information, please see license.txt

# NOTE: Connections shown on the User form for records this app owns. User is a core
# Frappe DocType, so this goes through `override_doctype_dashboards` rather than
# user.json (apps/frappe is not ours to edit) or a Customize Form link (database only,
# never reaches another bench). Other apps add their own User connections the same way.

# NOTE: (group label, linked DocType, the field on that DocType pointing at User).
# Email Account's `linked_user` is a custom field this app creates in custom_fields.py.
USER_CONNECTIONS = (
	("Email", "Email Account", "linked_user"),
)


def GetUserDashboardData(data: dict) -> dict:
	"""Add this app's User connections to the form dashboard data."""
	transactions = data.setdefault("transactions", [])
	nonStandardFieldnames = data.setdefault("non_standard_fieldnames", {})

	for groupLabel, linkDoctype, linkFieldname in USER_CONNECTIONS:
		AddConnection(transactions, groupLabel, linkDoctype)
		nonStandardFieldnames[linkDoctype] = linkFieldname

	return data



def AddConnection(transactions: list, groupLabel: str, linkDoctype: str) -> None:
	"""Append linkDoctype to its dashboard group, creating the group if absent."""
	for group in transactions:
		if group.get("label") != groupLabel:
			continue

		items = group.setdefault("items", [])
		if linkDoctype not in items:
			items.append(linkDoctype)
		return

	transactions.append({"label": groupLabel, "items": [linkDoctype]})
