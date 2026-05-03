# M365 Email App - Bug Fixes Report

**Date:** 2026-01-26
**Reported by:** Paul Johnson
**Environment:** staging.theteam.net.au

---

## Summary

After installing the updated M365 Email app, several issues were discovered related to the migration from the standalone `M365 Email Account` doctype to the integrated `Email Account` doctype (with `service='M365'`). This document details each issue and the fixes implemented.

---

## Issue 1: Missing "Test Connection" Button

### Problem
The `M365 Email Service Principal Settings` form had no UI button to test the OAuth connection. The `test_connection()` function existed in Python ([auth.py](m365email/m365email/auth.py)) and the API endpoint `test_service_principal_connection` was available ([api.py](m365email/m365email/api.py)), but there was no JavaScript file to add the button to the form.

### Impact
Users had no way to verify their Azure AD credentials were working without attempting to sync emails.

### Fix
Created new file: `m365email/m365email/doctype/m365_email_service_principal_settings/m365_email_service_principal_settings.js`

```javascript
frappe.ui.form.on("M365 Email Service Principal Settings", {
    refresh(frm) {
        if (!frm.is_new()) {
            frm.add_custom_button(__("Test Connection"), function() {
                testConnection(frm);
            });
        }
    }
});

function testConnection(frm) {
    frappe.call({
        method: "m365email.m365email.api.test_service_principal_connection",
        args: { service_principal_name: frm.doc.name },
        freeze: true,
        freeze_message: __("Testing connection..."),
        callback: function(response) {
            // Display success/failure message
        }
    });
}
```

---

## Issue 2: Authority URL Placeholder Not Replaced

### Problem
The `authority_url` field in `M365 Email Service Principal Settings` had a default value of:
```
https://login.microsoftonline.com/{tenant_id}
```

This literal placeholder string was never replaced with the actual tenant ID. When MSAL attempted to authenticate, it failed with:
```
Unable to get authority configuration for https://login.microsoftonline.com/{tenant_id}.
Authority would typically be in a format of https://login.microsoftonline.com/your_tenant...
```

### Impact
All M365 authentication failed until the user manually corrected the Authority URL.

### Fix (3 parts)

#### Part A: Remove broken default from doctype JSON
**File:** `m365email/m365email/doctype/m365_email_service_principal_settings/m365_email_service_principal_settings.json`

```json
// BEFORE
{
    "default": "https://login.microsoftonline.com/{tenant_id}",
    "fieldname": "authority_url",
    "fieldtype": "Data",
    "label": "Authority URL"
}

// AFTER
{
    "description": "Auto-populated from Tenant ID",
    "fieldname": "authority_url",
    "fieldtype": "Data",
    "label": "Authority URL"
}
```

#### Part B: Server-side validation to auto-fix
**File:** `m365email/m365email/doctype/m365_email_service_principal_settings/m365_email_service_principal_settings.py`

```python
def validate(self):
    """Validate service principal settings"""
    # NOTE: Fix authority URL if it contains placeholder or is missing
    if self.tenant_id:
        if not self.authority_url or "{tenant_id}" in self.authority_url:
            self.authority_url = f"https://login.microsoftonline.com/{self.tenant_id}"
```

#### Part C: Client-side auto-fill when tenant ID is entered
**File:** `m365email/m365email/doctype/m365_email_service_principal_settings/m365_email_service_principal_settings.js`

```javascript
tenant_id(frm) {
    fixAuthorityUrl(frm);
}

function fixAuthorityUrl(frm) {
    let tenantId = frm.doc.tenant_id;
    let authorityUrl = frm.doc.authority_url;

    if (!tenantId) return;

    let hasPlaceholder = authorityUrl && authorityUrl.includes("{tenant_id}");
    let isEmpty = !authorityUrl;

    if (hasPlaceholder || isEmpty) {
        let correctUrl = `https://login.microsoftonline.com/${tenantId}`;
        frm.set_value("authority_url", correctUrl);
    }
}
```

---

## Issue 3: M365 Email Sync Log Links to Wrong Doctype

### Problem
The `M365 Email Sync Log` doctype had its `email_account` Link field configured with:
```json
"options": "M365 Email Account"
```

This referenced the old/deprecated doctype instead of the current `Email Account` doctype (with `service='M365'`).

When email sync attempted to create a sync log entry, it failed with:
```
Could not find Email Account: SGC - Paul Johnson
```

### Impact
The "Pull Emails" button failed completely. Email sync could not complete because it couldn't create sync log entries.

### Fix
**File:** `m365email/m365email/doctype/m365_email_sync_log/m365_email_sync_log.json`

```json
// BEFORE
{
    "fieldname": "email_account",
    "fieldtype": "Link",
    "options": "M365 Email Account",
    "reqd": 1
}

// AFTER
{
    "fieldname": "email_account",
    "fieldtype": "Link",
    "options": "Email Account",
    "reqd": 1
}
```

---

## Issue 4: Migration Patches Reference Wrong Doctype

### Problem
Two migration patch files were creating custom fields that linked to the old `M365 Email Account` doctype instead of `Email Account`.

### Impact
Future installations would have broken Link fields on Event and Email Queue doctypes.

### Fix

#### Patch A: Event Custom Fields
**File:** `m365email/patches/add_event_custom_fields.py`

```python
# BEFORE
{
    "fieldname": "m365_email_account",
    "options": "M365 Email Account",
    ...
}

# AFTER
{
    "fieldname": "m365_email_account",
    "options": "Email Account",
    ...
}
```

#### Patch B: Email Queue Custom Fields
**File:** `m365email/patches/add_email_queue_custom_fields.py`

```python
# BEFORE
{
    "fieldname": "m365_account",
    "options": "M365 Email Account",
    ...
}

# AFTER
{
    "fieldname": "m365_account",
    "options": "Email Account",
    ...
}
```

---

## Files Modified Summary

| File | Change |
|------|--------|
| `doctype/m365_email_service_principal_settings/m365_email_service_principal_settings.js` | **Created** - Test Connection button + Authority URL auto-fill |
| `doctype/m365_email_service_principal_settings/m365_email_service_principal_settings.json` | Removed broken `{tenant_id}` default |
| `doctype/m365_email_service_principal_settings/m365_email_service_principal_settings.py` | Added validate() to fix Authority URL |
| `doctype/m365_email_sync_log/m365_email_sync_log.json` | Changed Link from `M365 Email Account` to `Email Account` |
| `patches/add_event_custom_fields.py` | Changed Link from `M365 Email Account` to `Email Account` |
| `patches/add_email_queue_custom_fields.py` | Changed Link from `M365 Email Account` to `Email Account` |

---

## Root Cause Analysis

The app was migrated from using a standalone `M365 Email Account` doctype to integrating with Frappe's standard `Email Account` doctype (using `service='M365'`). However, several references to the old doctype name were not updated during this migration:

1. The `M365 Email Sync Log` doctype definition
2. Two migration patch files
3. The Authority URL default value used a placeholder syntax that was never implemented

---

## Recommendations

1. **Search for remaining references:** Run `grep -r "M365 Email Account" --include="*.py" --include="*.json"` to find any other references that may need updating.

2. **Add validation tests:** Consider adding automated tests that verify Link fields point to existing doctypes.

3. **Document the migration:** If the old `M365 Email Account` doctype is truly deprecated, consider adding a deprecation notice or migration guide.

---

## Testing Verification

After applying these fixes:
- ✅ Test Connection button appears and works
- ✅ Authority URL auto-populates correctly
- ✅ Pull Emails button works without errors
- ✅ Email sync completes successfully

---

## Issue 5: User-to-Email Account Linking and Privacy Controls

**Date:** 2026-01-26

### Problem
The M365 Email Account system lacked a clear mechanism for:
1. Linking an email account to a specific Frappe user
2. Controlling which user can send from which email account
3. **Privacy**: Preventing other users (including System Managers) from viewing emails synced to another user's account

The previous implementation had `user_default` and `user_default_for` fields on the Outgoing tab, but:
- These were confusing to configure
- They only controlled sending, not viewing
- There was no privacy filtering on Communications

### Impact
- Users could potentially see other users' synced emails
- System Managers had full access to all Communications
- No clear ownership model for email accounts

### Fix (Multi-part Implementation)

#### Part A: Replace User Default Fields with Linked User

**File:** `m365email/m365email/custom_fields.py`

Removed `user_default` and `user_default_for` fields. Added single `linked_user` field on the Details tab:

```python
{
    "fieldname": "linked_user",
    "label": "Linked User",
    "fieldtype": "Link",
    "options": "User",
    "insert_after": "email_id",
    "description": "The Frappe user who owns this email account. Only this user can view emails synced from this account."
}
```

#### Part B: Add Privacy Field to Communication

**File:** `m365email/m365email/custom_fields.py`

Added `linked_user` field to Communication doctype for privacy filtering:

```python
{
    "fieldname": "linked_user",
    "label": "Linked User",
    "fieldtype": "Link",
    "options": "User",
    "insert_after": "m365_email_account",
    "read_only": 1,
    "hidden": 1,
    "description": "User who owns this email for privacy filtering"
}
```

#### Part C: Create Permission Functions

**File:** `m365email/m365email/permissions.py` (NEW)

```python
def get_communication_permission_query_conditions(user: str) -> str | None:
    """
    Filter Communications so users only see emails from their linked accounts.

    Access Rules:
    - Non-email Communications: visible to all
    - Email Communications with linked_user: only visible to that user
    - Email Communications without linked_user: visible to all (legacy)
    """
    return """(
        `tabCommunication`.communication_medium != 'Email'
        OR `tabCommunication`.linked_user = {user}
        OR `tabCommunication`.linked_user IS NULL
        OR `tabCommunication`.linked_user = ''
    )""".format(user=frappe.db.escape(user))


def has_communication_permission(doc, ptype: str = "read", user: str = None) -> bool:
    """Check if user has permission to access a specific Communication."""
    if doc.communication_medium != "Email":
        return True

    linkedUser = getattr(doc, 'linked_user', None)
    if not linkedUser:
        return True

    return linkedUser == user
```

#### Part D: Register Permission Hooks

**File:** `m365email/hooks.py`

```python
permission_query_conditions = {
    "Communication": "m365email.m365email.permissions.get_communication_permission_query_conditions",
}

has_permission = {
    "Communication": "m365email.m365email.permissions.has_communication_permission",
}
```

#### Part E: Set linked_user When Syncing Emails

**File:** `m365email/m365email/sync.py`

```python
# NOTE: Set owner and linked_user for access control
linkedUser = getattr(email_account, 'linked_user', None)
if linkedUser:
    comm.owner = linkedUser
    comm.linked_user = linkedUser
```

#### Part F: Update Sending Logic

**File:** `m365email/m365email/send.py`

Updated `GetUserDefaultEmailAccount()` to use `linked_user` instead of `user_default_for`.

### Privacy Model

| User Role | Can See Own Emails | Can See Others' Emails |
|-----------|-------------------|----------------------|
| Regular User | Yes | No |
| System Manager | Yes | No |
| Administrator | Yes | Yes (Frappe limitation) |

**Note:** The `Administrator` user always has full access in Frappe - this cannot be overridden by any app.

### Files Modified Summary

| File | Change |
|------|--------|
| `m365email/permissions.py` | **Created** - Permission functions for Communication privacy |
| `hooks.py` | Added `permission_query_conditions` and `has_permission` hooks |
| `m365email/custom_fields.py` | Added `linked_user` to Email Account and Communication |
| `m365email/sync.py` | Set `linked_user` when creating Communications |
| `m365email/send.py` | Updated to use `linked_user` for account lookup |
| `m365email/email_account_override.py` | Updated validation for `linked_user` |

### Configuration

To link an email account to a user:
1. Open the Email Account (e.g., "SGC - Paul Johnson")
2. On the **Details** tab, set **Linked User** to the Frappe user
3. On the **Outgoing** tab, check **Default Outgoing** if this is their primary sending account

### Testing Verification

- ✅ `linked_user` field appears on Email Account Details tab
- ✅ `linked_user` field added to Communication (hidden)
- ✅ Permission hooks registered correctly
- ✅ Users can only see their own synced emails
- ✅ System Managers cannot see other users' emails in list view
