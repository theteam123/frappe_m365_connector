# Copyright (c) 2026, Team Group Pty Ltd and contributors
# For license information, please see license.txt

"""Compare two login IDs the way the database does: ignoring capitals.

Office 365 single sign-on hands Frappe the address exactly as the identity provider sends it, so
where a directory holds a capitalised userPrincipalName and the User record is lowercase, the two
differ: frappe/utils/oauth.py passes the raw string to login_as, while the password path
normalises it to the stored name. MariaDB's collation ignores capitals, so every SQL check still
passes and such a user lists and edits records normally - only PYTHON comparisons fail, silently.

Deliberately a local copy rather than an import: this decides access, and an authorisation check
must not depend on another app being installed. tg_custom/login_identity.py holds the same pair.
"""

import frappe


def SameLogin(leftLogin: str | None, rightLogin: str | None) -> bool:
    """True when both names identify the same account, whatever capitals either one uses."""
    # NOTE: An empty value matches nothing, including another empty value, as `==` never did.
    if not leftLogin or not rightLogin:
        return False

    return leftLogin.lower() == rightLogin.lower()



def IsSessionUser(login: str | None) -> bool:
    """True when the given login is the signed-in user, whatever capitals it uses."""
    return SameLogin(login, frappe.session.user)
