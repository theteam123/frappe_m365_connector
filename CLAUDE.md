# Claude Development Instructions — m365email

## Master Coding Standards

**All code MUST follow the SGC bench-wide standards (single source of truth):**
- `/home/frappe/frappe-bench/claude-config-docs/CodingStandards.md`
- `/home/frappe/frappe-bench/claude-config-docs/security-rules.md` (SEC-001…005)

Read them before writing any code. The rules below are app-specific
additions — they do not override the standards repo.

---

## CRITICAL: This Is NOT ERPNext

This platform runs on **Frappe Framework** — NOT ERPNext.

- **Do NOT** assume any ERPNext doctypes exist
- **Do NOT** import from `erpnext.*` — it is not installed
- All doctypes in this system are **custom-built** by Team Group
- When searching for solutions, search for **Frappe Framework** patterns,
  not ERPNext patterns

---

## App Identity

| Field | Value |
|-------|-------|
| **App name** | `m365email` |
| **Publisher** | `TierneyMorris Pty Ltd` |
| **Module** | `M365Email` |
| **Description** | Microsoft 365 email integration via Graph API |
| **Dependencies** | `msal>=1.24.0`, `requests>=2.31.0` |

---

## Architecture Overview

Bidirectional Microsoft 365 email integration using Azure AD Service Principals
(application permissions, no user passwords). Intercepts Frappe's standard email
pipeline for both incoming sync and outgoing send.

**Incoming:** Scheduled delta sync every 5 minutes via Microsoft Graph Delta API.
**Outgoing:** Intercepts Email Queue `before_insert`, marks for M365, sends via Graph API.

---

## DocTypes (4)

| DocType | Purpose |
|---------|---------|
| **M365 Email Service Principal Settings** | Azure AD credentials, OAuth token cache, auto-provisioning config |
| **M365 Email Account** | Per-mailbox config (user or shared), folder filters, sync state |
| **M365 Email Folder Filter** | Child table — per-folder sync enable/disable and delta tokens |
| **M365 Email Sync Log** | Audit trail for all sync operations (naming: `SYNC-LOG-{#####}`) |

---

## Core Modules

| Module | Lines | Purpose |
|--------|-------|---------|
| `auth.py` | ~210 | MSAL token acquisition, caching, encryption, refresh |
| `graph_api.py` | ~379 | Microsoft Graph API wrapper (messages, attachments, folders, events, send) |
| `sync.py` | ~406 | Incoming email sync engine — delta tokens, deduplication, Communication creation |
| `send.py` | ~561 | Outgoing email — intercepts Email Queue, auto-provisioning, MIME building |
| `event_sync.py` | ~344 | Calendar event sync with timezone handling |
| `tasks.py` | ~271 | Scheduled task orchestration |
| `custom_fields.py` | ~398 | Adds M365 fields to Email Account and Email Queue doctypes |
| `utils.py` | ~371 | Helpers — message filtering, contact creation, datetime parsing |
| `email_account_override.py` | ~259 | Extends frappe Email Account for M365 sending |
| `email_queue_override.py` | ~1688 | Extends frappe Email Queue — routes M365 emails via Graph API |
| `email_override.py` | ~190 | Overrides `frappe.core.doctype.communication.email.make` |

---

## Hooks (Key Integration Points)

```python
# Overrides Frappe core doctype classes
override_doctype_class = {
    "Email Account": "m365email.m365email.email_account_override.ExtendedEmailAccount",
    "Email Queue": "m365email.m365email.email_queue_override.M365EmailQueue"
}

# Intercepts email queue creation
doc_events = {
    "Email Queue": {
        "before_insert": "m365email.m365email.send.intercept_email_queue"
    }
}

# Overrides Frappe's email.make for M365 routing
override_whitelisted_methods = {
    "frappe.core.doctype.communication.email.make": "m365email.m365email.email_override.make"
}

# Custom fields added to Email Account and Email Queue after install/migrate
after_install = "m365email.m365email.custom_fields.create_m365_custom_fields"
after_migrate = "m365email.m365email.custom_fields.create_m365_custom_fields"
```

---

## Scheduled Tasks

| Schedule | Task | Purpose |
|----------|------|---------|
| `*/5 * * * *` | `tasks.sync_all_email_accounts` | Delta sync all enabled email accounts |
| `*/5 * * * *` | `tasks.sync_all_calendar_events` | Delta sync all calendar events |
| Hourly | `tasks.refresh_all_tokens` | Refresh OAuth tokens before expiration |
| Daily | `tasks.cleanup_old_logs` | Delete sync logs older than 30 days |
| Daily | `tasks.validate_service_principals` | Test all SP connections |

---

## API Endpoints (10+)

All in `m365email.m365email.api`:

| Endpoint | Purpose |
|----------|---------|
| `enable_email_sync` | Create M365 Email Account |
| `disable_email_sync` | Disable incoming sync |
| `trigger_manual_sync` | Manual email sync |
| `trigger_manual_event_sync` | Manual calendar sync |
| `get_sync_status` | Sync status and recent logs |
| `test_service_principal_connection` | Test Azure AD credentials (System Manager) |
| `get_available_service_principals` | List enabled SPs for dropdown |
| `get_shared_mailboxes` | List shared mailboxes (System Manager) |
| `get_available_folders` | List mail folders from Graph API |
| `update_folder_filters` | Update per-folder sync config |

No CSRF-exempt endpoints.

---

## Security

- **Authentication:** MSAL client credentials flow (Service Principal — no user passwords)
- **Token storage:** Encrypted in database, auto-refreshed 5 minutes before expiry
- **Custom permissions:** `has_permission` on M365 Email Account — users can only access own mailbox
- **Role:** `M365 User` for non-admin access
- **Shared mailboxes:** System Manager only, unless role explicitly assigned

---

## Custom Fields Added to Frappe Core DocTypes

**Email Account:** `m365_service_principal`, `m365_account_type`, `m365_sync_events`,
`m365_sync_attachments`, `m365_initial_sync_days`, `m365_delta_token`,
`m365_calendar_delta_token`, `m365_folder_filter`, `m365_last_sync_time`,
`m365_last_sync_status`, `m365_sync_error_message`

**Email Queue:** `m365_send`, `m365_account`

---

## Existing Documentation

- `README.md` — Feature overview, architecture, module descriptions
- `README_SETUP.md` — Azure AD setup, Service Principal config, installation
- `SENDING_SETUP.md` — Email sending config, permissions, troubleshooting

---

## Patches

1. `add_email_queue_custom_fields` — M365 fields on Email Queue
2. `migrate_enable_fields` — Migrate enable_incoming/enable_outgoing flags
3. `add_event_custom_fields` — Calendar event fields
4. `add_event_timezone_field` — Timezone field for events

---

## Staging Note

M365 Email integration is auto-disabled after production database refresh on staging.
Re-enable from the Service Principal Settings page if testing on staging.

---

## Build

```bash
bench build --app m365email
```
