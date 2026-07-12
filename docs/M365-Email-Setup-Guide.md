# Microsoft 365 Email Setup Guide for Frappe

This guide explains how to configure Microsoft 365 email integration for **sending and receiving emails via Microsoft Graph API** using Azure AD Service Principal authentication.

---

## Prerequisites

- A Frappe site with the `m365email` app installed
- A Microsoft 365 account with admin access to Azure Portal
- Administrator access to your Frappe site

---

## Part 1: Azure Portal Setup

### Step 1: Create App Registration

1. Go to [Azure Portal](https://portal.azure.com/)
2. Navigate to **Microsoft Entra ID** (formerly Azure Active Directory)
3. Click **App registrations** > **New registration**
4. Configure:
   - **Name**: `Frappe M365 Email Integration` (or your preferred name)
   - **Supported account types**: "Accounts in this organizational directory only"
   - **Redirect URI**: Leave blank (not needed for client credentials flow)
5. Click **Register**

### Step 2: Configure API Permissions

1. In your new App Registration, go to **API permissions**
2. Click **Add a permission** > **Microsoft Graph** > **Application permissions**
3. Add these permissions:

   | Permission | Purpose |
   |------------|---------|
   | `Mail.Read` | Read mail in all mailboxes |
   | `Mail.ReadWrite` | Read/write mail (if modifying emails) |
   | `Mail.Send` | Send email as any user |
   | `User.Read.All` | Read user profiles |
   | `MailboxSettings.Read` | Read mailbox settings |
   | `Calendars.Read` | Read calendar events (optional) |

4. **Critical**: Click **Grant admin consent for [Your Tenant]** and confirm
   - You must see green checkmarks next to each permission

### Step 3: Create Client Secret

1. Go to **Certificates & secrets** > **Client secrets**
2. Click **New client secret**
3. Configure:
   - **Description**: `Frappe Integration` (or descriptive name)
   - **Expires**: Select expiry (recommend 12 or 24 months)
4. Click **Add**
5. **Important**: Copy the **Value** immediately - it won't be shown again!

### Step 4: Gather Credentials

From the **Overview** page, note these values:

| Field | Where to Find |
|-------|---------------|
| **Tenant ID** | Directory (tenant) ID on Overview page |
| **Client ID** | Application (client) ID on Overview page |
| **Client Secret** | The value you copied in Step 3 |

---

## Part 2: Frappe Configuration

### Step 1: Create M365 Email Service Principal Settings

1. In Frappe, search for **M365 Email Service Principal Settings** in the awesomebar
2. Click **+ Add M365 Email Service Principal Settings**
3. Fill in:

   | Field | Value |
   |-------|-------|
   | **Service Principal Name** | A unique identifier (e.g., "SGC Australia" or "Main Tenant") |
   | **Tenant ID** | Paste from Azure |
   | **Tenant Name** | Your organization name (optional, for reference) |
   | **Client ID** | Paste Application (client) ID from Azure |
   | **Client Secret** | Paste the secret value from Azure |
   | **Enabled** | ✓ Check this box |

4. Leave these fields as defaults (auto-populated):
   - **Authority URL**: `https://login.microsoftonline.com/{tenant_id}`
   - **Graph API Endpoint**: `https://graph.microsoft.com/v1.0`
   - **Scopes**: `https://graph.microsoft.com/.default`

5. Click **Save**

### Step 2: Test the Connection

After saving, verify the credentials work:

1. Open browser Developer Tools (F12) > Console tab
2. Run this JavaScript:

   ```javascript
   frappe.call({
       method: 'm365email.m365email.api.test_service_principal_connection',
       args: {
           service_principal_name: 'YOUR_SERVICE_PRINCIPAL_NAME'
       },
       callback: function(r) {
           console.log(r.message);
       }
   });
   ```

3. **Expected Success Response**:
   ```json
   {"success": true, "message": "Connection successful! Token acquired.", "token_expires_at": "2026-01-24 09:31:59"}
   ```

4. If you see a `token_cache` value populated in the Service Principal Settings, authentication is working.

### Step 3: Create M365 Email Account

1. Search for **M365 Email Account** in the awesomebar
2. Click **+ Add M365 Email Account**
3. Fill in:

   | Field | Value |
   |-------|-------|
   | **Account Name** | Identifier (e.g., user's name) |
   | **Account Type** | "User Mailbox" (or "Shared Mailbox" if applicable) |
   | **Email Address** | The M365 email address to sync (e.g., paul@company.com) |
   | **User** | Select the Frappe user who will own this account |
   | **Service Principal** | Select the Service Principal you created |
   | **Enable Incoming** | ✓ Check to sync incoming emails |
   | **Enable Outgoing** | ✓ Check to send emails via M365 |

4. Optional settings:
   - **Sync Calendar Events**: Check to sync calendar events
   - **Sync From Date**: Only sync emails after this date
   - **Auto Create Contact**: Auto-create Contact records from email senders
   - **Sync Attachments**: Download and store email attachments
   - **Max Attachment Size**: Limit attachment size (default 10 MB)

5. Click **Save**

**Note**: You may see an orange warning if the email address doesn't match the Frappe user's email. This is just a warning - the configuration will still work if the M365 email is correct.

---

## Part 3: Testing the Integration

### Test Email Sync (Incoming)

Trigger a manual sync to verify incoming email works:

```javascript
frappe.call({
    method: 'm365email.m365email.api.trigger_manual_sync',
    args: {
        email_account_name: 'YOUR_ACCOUNT_NAME'
    },
    callback: function(r) {
        console.log(r.message);
    }
});
```

**Expected Success Response**:
```json
{"success": true, "fetched": 10, "created": 5, "updated": 0, "failed": 0}
```

- **fetched**: Messages retrieved from M365
- **created**: New Communications created in Frappe
- **failed**: Should be 0

### Test Email Sending (Outgoing)

1. Open any document in Frappe
2. Click **Menu (...)** > **Email**
3. Compose and send a test email
4. Check **Email Queue** to verify status is "Sent"

### Check Sync Status

View sync history and status:

```javascript
frappe.call({
    method: 'm365email.m365email.api.get_sync_status',
    args: {
        email_account_name: 'YOUR_ACCOUNT_NAME'
    },
    callback: function(r) {
        console.log(r.message);
    }
});
```

---

## Part 4: How It Works

### Authentication Flow

The integration uses **Azure AD Service Principal authentication** (OAuth2 Client Credentials flow):

1. Service Principal Settings stores encrypted credentials
2. MSAL library acquires access token from Azure AD
3. Token is cached and automatically refreshed
4. All API calls use the Bearer token

**No user login required** - this is a server-to-server authentication flow.

### Email Sync (Incoming)

- A scheduled task runs **every minute** and syncs each account once its configured **Email Sync Interval** (default 5 minutes) has elapsed
- Uses Microsoft Graph **Delta Queries** for incremental sync (only fetches new/changed emails)
- Creates **Communication** documents in Frappe for each email
- Stores `m365_message_id` to prevent duplicate syncing
- Attachments are downloaded and stored as File documents

### Email Sending (Outgoing)

- Hooks into Frappe's **Email Queue** system
- Intercepts emails and routes them through Graph API instead of SMTP
- Uses `Mail.Send` permission to send as the specified user
- Automatically saves to Sent Items in M365

### Scheduled Tasks

| Task | Frequency | Purpose |
|------|-----------|---------|
| `sync_all_email_accounts` | Every minute (gated by **Email Sync Interval**, default 5 min) | Sync incoming emails |
| `sync_all_calendar_events` | Every minute (gated by **Calendar Sync Interval**, default 5 min) | Sync calendar events |
| `refresh_all_tokens` | Hourly | Keep tokens fresh |
| `cleanup_old_logs` | Daily | Delete old sync logs |
| `validate_service_principals` | Daily | Check credentials still valid |

> **Note:** Outgoing email is **not** a scheduled task — it is sent through Frappe's
> standard Email Queue via the `M365EmailQueue.send()` override (see *Email Sending
> (Outgoing)* above), so there is no separate "send" task.

---

## Troubleshooting

### "Incoming is not enabled" error

**Cause**: The sync code checks the `enable_incoming` field.

**Fix**: Ensure **Enable Incoming** is checked on the M365 Email Account.

### "Connection successful" but no emails syncing

1. Check that **Enable Incoming** is checked on the M365 Email Account
2. Verify the email address is correct and has emails in the Inbox
3. Check **Sync From Date** isn't set to a future date
4. Wait for the scheduled task (runs every 5 minutes) or trigger manual sync

### Token acquisition failed

1. Verify **Client Secret** hasn't expired in Azure Portal
2. Check **Client ID** and **Tenant ID** are correct
3. Ensure **Admin Consent** was granted for all permissions in Azure
4. Re-save the Service Principal Settings to clear cached token

### "Insufficient permissions" or 403 errors

1. Go to Azure Portal > App Registration > API permissions
2. Verify all required permissions are added (Mail.Read, Mail.Send, etc.)
3. Click **Grant admin consent** - you must see green checkmarks
4. Wait a few minutes for permissions to propagate

### Emails not appearing in Frappe

1. Check **Communication** list - synced emails appear there
2. Verify the email account's `last_sync_status` is "Success"
3. Check `sync_error_message` field for any errors
4. Look in **M365 Email Sync Log** for detailed sync history

### Changes not taking effect

After modifying Python code, restart the workers:

```bash
bench restart
```

Or if using development server, stop and restart `bench start`.

---

## Key Configuration Summary

### Azure Portal

| Setting | Value |
|---------|-------|
| App Type | App Registration |
| Account Types | Single tenant |
| Permissions | Application permissions (not Delegated) |
| Admin Consent | Required and granted |

### Frappe DocTypes

| DocType | Purpose |
|---------|---------|
| M365 Email Service Principal Settings | Stores Azure credentials |
| M365 Email Account | Configures which mailbox to sync |
| M365 Email Sync Log | Audit trail of sync operations |

### Required Azure Permissions

| Permission | Type | Purpose |
|------------|------|---------|
| Mail.Read | Application | Read emails |
| Mail.Send | Application | Send emails |
| User.Read.All | Application | Read user profiles |
| MailboxSettings.Read | Application | Read mailbox settings |
| Calendars.Read | Application | Read calendar (optional) |

---

## Quick Reference: API Endpoints

| Endpoint | Purpose |
|----------|---------|
| `test_service_principal_connection` | Test Azure credentials |
| `trigger_manual_sync` | Trigger immediate email sync |
| `get_sync_status` | Get account status and logs |
| `get_available_folders` | List mailbox folders |

---

## Key Lessons

1. **Service Principal Authentication**: Uses client credentials flow - no user interaction needed. The app authenticates as itself, not as a user.

2. **Application vs Delegated Permissions**: Always use **Application permissions** for server-to-server integration. Delegated permissions require user sign-in.

3. **Admin Consent Required**: Application permissions require admin consent before they work. Look for green checkmarks in Azure Portal.

4. **Delta Queries for Efficiency**: The sync uses Microsoft's delta query feature to only fetch changed emails, not the entire mailbox each time.

5. **One Service Principal, Multiple Accounts**: A single Service Principal can be used for multiple M365 Email Accounts. Each account specifies which mailbox to access.
