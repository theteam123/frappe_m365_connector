# Copyright (c) 2025, Zeke Tierney and contributors
# For license information, please see license.txt

"""
Permission functions for M365 Email Integration.

Implements strict privacy controls for Communications synced from M365 accounts.
Only the linked user can see their emails - even System Managers cannot view
other users' emails.
"""

import frappe


PRIVACY_SETTINGS_DOCTYPE = "M365 Email Settings"
PRIVACY_SETTINGS_FIELD = "enable_email_privacy_filter"


def IsPrivacyFilterEnabled() -> bool:
    """
    Read the admin-controlled toggle for the M365 email privacy filter.

    The hooks for `permission_query_conditions` and `has_permission` are
    always registered, but both functions early-return when this returns
    False, so the filter is effectively off until an admin enables it
    from M365 Email Settings.

    Returns:
        True when the filter should be enforced, False to behave like
        standard Frappe Communication permissions.
    """
    return bool(frappe.db.get_single_value(PRIVACY_SETTINGS_DOCTYPE, PRIVACY_SETTINGS_FIELD))



def get_communication_permission_query_conditions(user: str) -> str | None:
    """
    Filter Communications so users only see emails from their linked accounts.

    This function is called by Frappe's permission system to generate SQL WHERE
    conditions for list queries on the Communication doctype.

    Access Rules:
    - Filter disabled in M365 Email Settings: no filtering (standard perms)
    - Non-email Communications: visible to all (standard Frappe permissions apply)
    - Email Communications with linked_user: only visible to that user
    - Email Communications without linked_user: visible to all (legacy data)

    Args:
        user: The user ID to check permissions for

    Returns:
        SQL WHERE clause string, or None for no filtering
    """
    if not IsPrivacyFilterEnabled():
        return None

    if not user:
        user = frappe.session.user

    # NOTE: Filter all users including System Manager
    # Only Administrator user bypasses this (Frappe core limitation)
    return """(
        `tabCommunication`.communication_medium != 'Email'
        OR `tabCommunication`.linked_user = {user}
        OR `tabCommunication`.linked_user IS NULL
        OR `tabCommunication`.linked_user = ''
    )""".format(user=frappe.db.escape(user))



def has_communication_permission(doc, ptype: str = "read", user: str = None) -> bool:
    """
    Check if user has permission to access a specific Communication.

    This function is called by Frappe when checking permission on individual
    Communication documents.

    Access Rules:
    - Non-email Communications: allow (standard permissions)
    - Email with linked_user set: only that user can access
    - Email without linked_user: allow (legacy data)

    Args:
        doc: The Communication document
        ptype: Permission type (read, write, etc.)
        user: The user to check permissions for

    Returns:
        True if user has permission, False otherwise
    """
    if not IsPrivacyFilterEnabled():
        return True

    if not user:
        user = frappe.session.user

    # NOTE: Non-email communications use standard Frappe permissions
    if doc.communication_medium != "Email":
        return True

    # NOTE: If no linked_user set, allow access (legacy data or non-M365 emails)
    linkedUser = getattr(doc, 'linked_user', None)
    if not linkedUser:
        return True

    # NOTE: Only the linked user can access their emails
    return linkedUser == user
