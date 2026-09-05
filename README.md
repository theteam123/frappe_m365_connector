# M365 Email Integration for Frappe

A comprehensive Frappe app that integrates Microsoft 365 email using Azure AD Service Principal authentication. This app provides **bidirectional email integration** - both syncing incoming emails and sending outgoing emails via Microsoft Graph API.

## 🌟 Features

### Email Syncing (Incoming)
- ✅ **Incremental Delta Sync** - Efficient syncing using Microsoft Graph Delta API
- ✅ **Multi-Tenant Support** - Connect multiple Azure AD tenants
- ✅ **User & Shared Mailboxes** - Sync both personal and shared mailboxes
- ✅ **Folder Filtering** - Choose which folders to sync (Inbox, Sent Items, etc.)
- ✅ **Automatic Deduplication** - Prevents duplicate emails using Message-ID
- ✅ **Attachment Support** - Downloads and stores email attachments
- ✅ **Scheduled Sync** - Scheduler runs every minute; each Service Principal syncs on its own configured interval (default 5 minutes)
- ✅ **Manual Sync** - Trigger sync on-demand via API
- ✅ **Calendar Event Sync** - Sync M365 calendar events to Frappe Events

### Email Sending (Outgoing)
- ✅ **Send via M365 Graph API** - No SMTP configuration required
- ✅ **Send As Any User** - Service Principal can send as any user in the organization
- ✅ **Automatic Queue Processing** - Emails are automatically sent via M365
- ✅ **Seamless Integration** - Works with `frappe.sendmail()` and Communication doctype
- ✅ **HTML Email Support** - Full HTML email formatting
- ✅ **Attachment Support** - Send emails with attachments

### Native Frappe Integration
- ✅ **Extends Email Account** - Uses standard Email Account doctype with `service='M365'`
- ✅ **Works with Notifications** - M365 accounts appear in all email selectors
- ✅ **Document Email** - Send emails from any document using M365
- ✅ **User Preferences** - Users can select M365 accounts in their email preferences

## 🏗️ Architecture

### How It Works

This app extends Frappe's standard **Email Account** doctype by adding `M365` as a service option. When an Email Account has `service='M365'`, it uses Microsoft Graph API instead of SMTP/IMAP.

#### Email Syncing Flow
```
┌─────────────────────────────────────────────────────────────────┐
│ 1. Scheduled Task (runs every minute; throttled per Service      │
│    Principal interval, default 5 minutes)                        │
│    └─> m365email.m365email.tasks.sync_all_email_accounts()     │
└─────────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────────┐
│ 2. For Each Email Account with service='M365' & enable_incoming │
│    └─> Get Access Token from Service Principal                  │
│    └─> Call Microsoft Graph Delta API                           │
│        /users/{email}/messages/delta                             │
└─────────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────────┐
│ 3. Process Each Email Message                                   │
│    └─> Check if already synced (by Message-ID)                  │
│    └─> Create Communication document                            │
│    └─> Download attachments (if enabled)                        │
│    └─> Store delta token for next sync                          │
└─────────────────────────────────────────────────────────────────┘
```

#### Email Sending Flow
```
┌─────────────────────────────────────────────────────────────────┐
│ 1. User/System calls frappe.sendmail() or Communication.make() │
└─────────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────────┐
│ 2. Email Account Override (find_outgoing)                       │
│    └─> Check for M365 account matching sender email             │
│    └─> Fall back to default_outgoing M365 account               │
└─────────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────────┐
│ 3. Frappe's QueueBuilder Creates Email Queue                    │
│    └─> Formats HTML using templates/emails/standard.html        │
│    └─> Adds brand logo, header, footer from Email Account       │
│    └─> Adds unsubscribe link & tracking pixel placeholders      │
│    └─> Stores complete MIME message in Email Queue              │
└─────────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────────┐
│ 4. before_insert Hook Marks for M365                            │
│    └─> Set m365_send = 1                                        │
│    └─> Set m365_account = sending account                       │
└─────────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────────┐
│ 5. Frappe's Standard Queue Processing                           │
│    └─> frappe.email.queue.flush() processes pending emails      │
│    └─> Calls EmailQueue.send() on each email                    │
│    └─> Our M365EmailQueue.send() override intercepts            │
└─────────────────────────────────────────────────────────────────┘
                              ↓
┌─────────────────────────────────────────────────────────────────┐
│ 6. Send via Microsoft Graph API                                 │
│    └─> Parse MIME message from Email Queue                      │
│    └─> Extract subject, body, recipients                        │
│    └─> Replace placeholders (unsubscribe, tracking)             │
│    └─> Call /users/{sender}/sendMail                            │
│    └─> Update Email Queue status                                │
└─────────────────────────────────────────────────────────────────┘
```

## 📦 Components

### DocTypes

#### Email Account (Extended)
The standard Frappe Email Account doctype is extended with M365 support via custom fields.

**M365 Custom Fields Added:**
- `service` - Select field with 'M365' option (added via Property Setter)
- `linked_user` - Frappe user who owns this account (controls who can view its synced emails)
- `m365_service_principal` - Link to Service Principal Settings
- `m365_account_type` - User Mailbox or Shared Mailbox
- `m365_sync_events` - Enable/disable calendar event syncing
- `m365_sync_attachments` - Enable/disable attachment downloads
- `m365_sync_from_date` - Date to start syncing history from
- `m365_delta_tokens` - JSON store of per-folder delta sync state
- `m365_folder_filter` - Child table of folders to sync
- `m365_last_sync_time`, `m365_last_sync_status`, `m365_sync_error_message` - Sync status fields

**Note:** Shared mailbox access is managed via Frappe's standard User Email table (on User doctype), not custom roles.

**Standard Fields Used:**
- `email_id` - Email address
- `enable_incoming` - Enable email syncing
- `enable_outgoing` - Enable email sending
- `default_outgoing` - Use as default for sending

#### M365 Email Service Principal Settings
Stores Azure AD Service Principal credentials for authentication.

**Fields:**
- `service_principal_name` - Unique identifier
- `tenant_id` - Azure AD Tenant ID
- `tenant_name` - Friendly tenant name
- `client_id` - Azure AD Application (Client) ID
- `client_secret` - Azure AD Client Secret (encrypted)
- `enabled` - Enable/disable this service principal
- `token_cache` - Encrypted OAuth token cache

**Permissions:**
- Only System Managers can create/edit

#### M365 Email Sync Log
Tracks sync operations and errors.

**Fields:**
- `email_account` - Link to Email Account
- `sync_type` - Manual, Scheduled, or Initial
- `status` - Success, Failed, or Partial Success
- `start_time`, `end_time`, `duration` - Timing of the sync run
- `messages_fetched`, `messages_created`, `messages_updated`, `messages_failed` - Per-run counts
- `error_message` / `details` - Error details if failed

Naming: `SYNC-LOG-{#####}`

#### M365 Addin Session

A remembered sign-in for the **Send to ERP** Outlook add-in. Created when a user signs
in with Microsoft 365 from the add-in; lets later panel opens skip the sign-in.

**Fields:** `user`, `platform`, `expires_on` (30-day sliding, 90-day absolute cap),
`last_used_on`, `revoked`, `token_hash` (SHA-256 only — the token itself is never stored).

**Permissions:** System Manager (list, revoke, delete). Records are only created by the
add-in (`addin_session.CreateAddinSession`), which requires a request authenticated
with a validated Microsoft token.

### Core Modules

#### `auth.py` - Authentication
- `get_access_token(service_principal_name)` - Get OAuth access token
- Token caching with encryption

#### `graph_api.py` - Microsoft Graph API
- `make_graph_request(endpoint, access_token, method, data, params)` - Generic API wrapper (handles delta queries, pagination, rate limiting)
- `send_email_as_user(sender_email, recipients, subject, body, access_token, ...)` - Send email

#### `sync.py` - Email Syncing
- `sync_email_account(email_account_name)` - Sync a single account
- `create_communication_from_message_for_email_account()` - Process emails
- Delta token management for incremental sync (stores the full `@odata.deltaLink` URL per folder)

#### `provisioning.py` - Account Provisioning
- `FindProvisioningServicePrincipal(domain)` - The enabled Service Principal with auto-provisioning on for a domain
- `ProbeMailbox(email, service_principal)` - Ask Microsoft 365 whether a mailbox exists (see `graph_api.MailboxExists`)
- `CreateM365Account(sp, user, email)` - Insert an outgoing-only M365 Email Account owned by a user
- `EnsureM365AccountForUser(user, email="", isMailboxLive=None)` - Idempotent: create the account if the mailbox exists

Auto-provisioning is switched on per Service Principal (**Enable Auto-Provisioning** + **Domain**). Two paths use it:
the lazy one in `send.py` (a user with the `M365 User` role sends for the first time) and any app that walks its own
people list and calls `EnsureM365AccountForUser` once it knows the mailbox is live. On the SGC benches the `hr` app's
daily `SyncEmployeeMailboxes` task does exactly that, and the `tg_custom` signature hook fills in the signature on insert.

#### `send.py` - Email Sending
- `get_sending_account_for_sender(sender_email)` - Find M365 account for a sender
- `GetUserDefaultEmailAccount(user)` - Find a user's default outgoing account (via `linked_user`)
- `intercept_email_queue(doc, method)` - Hook to mark emails for M365
- `send_via_m365(email_queue_doc)` - Send email via Graph API
- `M365SendContext` - Helper class for building personalized M365 emails

#### `email_account_override.py` - Email Account Class Override
- Extends Email Account doctype class
- Custom `send()` method for M365 emails
- Overrides `find_outgoing()` to support M365 accounts
- Registered via `override_doctype_class` hook

#### `email_queue_override.py` - Email Queue Class Override
- Extends Email Queue doctype class
- Overrides `send()` method to intercept M365 emails
- If `m365_send=1`, sends via Graph API instead of SMTP
- Otherwise calls parent `send()` for standard SMTP
- Integrates with Frappe's standard queue processing

#### `custom_fields.py` - Custom Fields
- Adds M365 fields to Email Account doctype
- Adds `m365_send` and `m365_account` fields to Email Queue
- Applied during installation via `after_migrate` hook

#### `tasks.py` - Scheduled Tasks
- `sync_all_email_accounts()` - Sync all enabled accounts (scheduler runs every minute; per-SP interval throttling)
- `sync_all_calendar_events()` - Sync all enabled calendars (same schedule)
- `refresh_all_tokens()` - Refresh tokens (hourly)
- `cleanup_old_logs()` - Delete old logs (daily)
- `validate_service_principals()` - Check credentials (daily)

#### `api.py` - Whitelisted API Endpoints
- `enable_email_sync()` - Create / enable an M365 Email Account
- `disable_email_sync()` - Disable sync
- `trigger_manual_sync()` - Manual email sync trigger
- `trigger_manual_event_sync()` - Manual calendar sync trigger
- `get_sync_status()` - Get sync statistics
- `test_service_principal_connection()` - Test credentials
- `get_available_service_principals()`, `get_shared_mailboxes()`, `get_available_folders()`, `update_folder_filters()`, `SendEmailQueueNow()`

### Hooks Configuration

**`hooks.py`** registers:

```python
# Override Email Account and Email Queue doctype classes for M365 support
override_doctype_class = {
    "Email Account": "m365email.m365email.email_account_override.ExtendedEmailAccount",
    "Email Queue": "m365email.m365email.email_queue_override.M365EmailQueue"
}

# Hook into Email Queue creation to mark M365 emails
doc_events = {
    "Email Queue": {
        "before_insert": "m365email.m365email.send.intercept_email_queue"
    }
}

# Route Frappe's core email.make through M365 when applicable
override_whitelisted_methods = {
    "frappe.core.doctype.communication.email.make": "m365email.m365email.email_override.make"
}

# Filter Communications so users only see emails from their linked accounts
permission_query_conditions = {
    "Communication": "m365email.m365email.permissions.get_communication_permission_query_conditions"
}
has_permission = {
    "M365 Email Account": "m365email.m365email.doctype.m365_email_account.m365_email_account.has_permission",
    "Communication": "m365email.m365email.permissions.has_communication_permission"
}

# Validate Microsoft 365 tokens for the Send-to-ERP Outlook add-in
auth_hooks = [
    "m365email.m365email.addin_sso.ValidateAddinToken"
]

# Apply custom fields after install and after migration
after_install = "m365email.m365email.custom_fields.create_m365_custom_fields"
after_migrate = "m365email.m365email.custom_fields.create_m365_custom_fields"

# Scheduled tasks
# Note: M365 email sending is handled by Frappe's standard queue processing
# via our M365EmailQueue.send() override - no separate task needed
scheduler_events = {
    "cron": {
        # Scheduler fires every minute; each task throttles per Service
        # Principal interval (default 5 minutes)
        "* * * * *": [
            "m365email.m365email.tasks.sync_all_email_accounts",
            "m365email.m365email.tasks.sync_all_calendar_events"
        ]
    },
    "hourly": [
        "m365email.m365email.tasks.refresh_all_tokens"
    ],
    "daily": [
        "m365email.m365email.tasks.cleanup_old_logs",
        "m365email.m365email.tasks.validate_service_principals"
    ]
}
```

## 🚀 Installation

See [README_SETUP.md](README_SETUP.md) for detailed setup instructions.

**Quick Start:**

1. Install dependencies:
   ```bash
   bench --site your-site pip install msal
   ```

2. Run migrations (this adds M365 fields to Email Account):
   ```bash
   bench --site your-site migrate
   ```

3. Restart bench:
   ```bash
   bench restart
   ```

4. Configure Azure AD (see setup guide)

5. Create Service Principal Settings in Frappe

6. Create Email Accounts with service='M365' and configure

## 📧 Email Sending Setup

See [SENDING_SETUP.md](SENDING_SETUP.md) for detailed sending configuration.

**Quick Start:**

1. Ensure Service Principal has `Mail.Send` permission in Azure AD

2. Configure an Email Account for M365 sending:
   - Open Email Account > New
   - Set Service = "M365"
   - Set Email Address
   - Link to Service Principal
   - Check "Enable Outgoing"
   - Check "Default Outgoing" (for fallback)
   - Save

3. Send emails normally:
   ```python
   frappe.sendmail(
       recipients=["user@example.com"],
       sender="your-email@company.com",
       subject="Test Email",
       message="<p>Hello from M365!</p>"
   )
   ```

4. Emails are automatically sent via M365 Graph API!

## 📬 Shared Mailbox Setup

To give users access to a shared mailbox:

1. **Create the Shared Mailbox Email Account:**
   - Open Email Account > New
   - Set Service = "M365"
   - Set Email Address = shared mailbox email
   - Set Account Type = "Shared Mailbox"
   - Link to Service Principal
   - Check "Enable Outgoing" and/or "Enable Incoming"
   - Save

2. **Grant User Access:**
   - Go to the User's profile (User doctype)
   - Scroll to "User Emails" section
   - Add a new row with the shared mailbox Email Account
   - Save

3. Users will now see the shared mailbox in email dialogs and can send from it.

This uses Frappe's standard User Email table - no custom roles needed!

## 🔒 Security

- **Client Secrets**: Encrypted by Frappe's encryption system
- **Token Cache**: Encrypted before storage in database
- **Service Principal Auth**: Uses application permissions (no user passwords)
- **User Email Access**: Shared mailbox access controlled via User Email table
- **Communication Permissions**: Standard Frappe permission system

## 🧪 Testing & Debugging

### Debug Helpers

```python
from m365email.m365email.debug_helpers import *

# Check M365 sending configuration
check_m365_sending_config()

# Check Email Queue status
check_email_queue_status()

# Manually process email queue
manually_process_queue()

# Check recent errors
check_recent_errors()
```

### Manual Sync

```python
from m365email.m365email.sync import sync_email_account

# Sync a specific Email Account (with service='M365')
sync_email_account("Your Email Account Name")
```

## 📊 Monitoring

- **M365 Email Sync Log** - View sync history and errors
- **Email Queue** - Monitor outgoing emails (check `m365_send` field)
- **Communication** - View all synced emails
- **Error Log** - Check for M365-related errors

## 🤝 Contributing

Contributions are welcome! Please ensure:
- Code follows Frappe coding standards
- All new features include documentation
- Test thoroughly before submitting PR

## 📄 License

MIT License - See LICENSE file for details

## 🆘 Support

- **Setup Guide**: [README_SETUP.md](README_SETUP.md)
- **Sending Guide**: [SENDING_SETUP.md](SENDING_SETUP.md)
- **Frappe Forum**: https://discuss.frappe.io
- **GitHub Issues**: Report bugs and feature requests

---

**Built with ❤️ for the Frappe community**

