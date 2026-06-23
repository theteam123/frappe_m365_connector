# M365 Email Integration - Setup Guide

## Installation

1. **Install the app dependencies:**
   ```bash
   cd /path/to/frappe-bench
   bench --site your-site.local pip install msal
   ```

2. **Run migrations to create DocTypes:**
   ```bash
   bench --site your-site.local migrate
   ```

3. **Restart bench:**
   ```bash
   bench restart
   ```

## Azure AD Configuration

### 1. Create Azure AD App Registration

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** > **App registrations**
3. Click **New registration**
4. Enter a name (e.g., "Frappe M365 Email Integration")
5. Select **Accounts in this organizational directory only**
6. Click **Register**

### 2. Configure API Permissions

1. In your app registration, go to **API permissions**
2. Click **Add a permission** > **Microsoft Graph** > **Application permissions**
3. Add the following permissions:
   - `Mail.Read` - Read mail in all mailboxes
   - `Mail.ReadWrite` - Read and write mail in all mailboxes
   - `User.Read.All` - Read all users' profiles
   - `MailboxSettings.Read` - Read all mailbox settings
4. Click **Grant admin consent** for your organization

### 3. Create Client Secret

1. Go to **Certificates & secrets**
2. Click **New client secret**
3. Enter a description and select expiration period
4. Click **Add**
5. **Copy the secret value immediately** (you won't be able to see it again)

### 4. Note Your Credentials

You'll need:
- **Tenant ID** (from Overview page)
- **Client ID** (Application ID from Overview page)
- **Client Secret** (the value you just copied)

## Frappe Configuration

### 1. Create Service Principal Settings

1. In Frappe, go to **M365 Email Service Principal Settings**
2. Click **New**
3. Fill in the details:
   - **Service Principal Name**: A unique name (e.g., "Main Tenant")
   - **Tenant ID**: Your Azure AD Tenant ID
   - **Tenant Name**: Friendly name (e.g., "Contoso Corp")
   - **Client ID**: Application (client) ID from Azure
   - **Client Secret**: The secret value you copied
   - **Enabled**: Check this box
4. Click **Save**
5. Click **Test Connection** to verify credentials

### 2. Configure Email Accounts

M365 mailboxes are configured on Frappe's standard **Email Account** doctype with
**Service = "M365"** (this app extends Email Account; there is no separate "M365 Email
Account" doctype). The M365-specific fields appear once Service is set to M365.

#### For User Mailboxes:

1. Go to **Email Account**
2. Click **New**
3. Fill in:
   - **Service**: M365
   - **Email Address**: User's M365 email
   - **Linked User**: The Frappe user who owns this mailbox (controls who can view its synced emails)
   - **M365 Service Principal**: Select the service principal
   - **M365 Account Type**: User Mailbox
   - **Enable Incoming**: Check to sync mail
4. Configure folder filters (Inbox, Sent Items, etc.) and the **Sync From Date**
5. Click **Save**

#### For Shared Mailboxes (System Manager only):

1. Go to **Email Account**
2. Click **New**
3. Fill in:
   - **Service**: M365
   - **Email Address**: Shared mailbox email (e.g., support@company.com)
   - **Linked User**: An admin who has access to the shared mailbox in M365
   - **M365 Service Principal**: Select the service principal
   - **M365 Account Type**: Shared Mailbox
   - **Enable Incoming**: Check to sync mail
4. Configure folder filters
5. Click **Save**

### 3. Grant Access to Synced Emails

Synced emails are stored as **Communication** records owned by the account's **Linked User**.
This app filters Communications so users only see emails from accounts linked to them
(via `permission_query_conditions` on Communication). To let another user view a mailbox's
emails, set them as the **Linked User**, or grant a role with broad Communication read access.

## Scheduled Tasks

The following tasks run automatically:

- **Every minute** (scheduler cron): Sync all enabled email accounts and calendars. Each Service Principal is throttled to its own interval (default 5 minutes), so a mailbox is not actually polled every minute.
- **Hourly**: Refresh access tokens for all service principals
- **Daily**:
  - Cleanup old sync logs (older than 30 days)
  - Validate service principal credentials

## API Endpoints

Available whitelisted API endpoints:

- `m365email.m365email.api.enable_email_sync` - Enable email sync
- `m365email.m365email.api.disable_email_sync` - Disable email sync
- `m365email.m365email.api.trigger_manual_sync` - Manually trigger sync
- `m365email.m365email.api.get_sync_status` - Get sync status
- `m365email.m365email.api.test_service_principal_connection` - Test connection
- `m365email.m365email.api.get_available_service_principals` - List service principals
- `m365email.m365email.api.get_shared_mailboxes` - List shared mailboxes
- `m365email.m365email.api.get_available_folders` - Get mail folders
- `m365email.m365email.api.update_folder_filters` - Update folder filters

## Troubleshooting

### Sync Not Working

1. Check **M365 Email Sync Log** for errors
2. Verify service principal credentials are correct
3. Ensure Azure AD permissions are granted
4. Check that the email account is enabled
5. Verify the scheduler is running: `bench doctor`

### Token Errors

1. Go to **M365 Email Service Principal Settings**
2. Click **Test Connection**
3. If it fails, verify:
   - Client ID is correct
   - Client Secret is correct and not expired
   - Tenant ID is correct
   - Admin consent was granted for API permissions

### Emails Not Appearing

1. Confirm you are signed in as the account's **Linked User** (Communications are filtered per linked user)
2. Verify folder filters are configured correctly
3. Check **Communication** list for synced emails
4. Review **M365 Email Sync Log** for sync statistics

## Security Notes

- Client secrets are encrypted by Frappe
- Token cache is encrypted before storage
- Only System Managers can configure service principals
- Only System Managers can configure shared mailboxes
- Users can only configure their own user mailboxes
- Communication permissions control who can view emails

## Multi-Tenant Support

You can configure multiple Azure AD tenants:

1. Create separate **M365 Email Service Principal Settings** for each tenant
2. Each email account links to a specific service principal
3. Tokens are managed independently per tenant

## Support

For issues or questions, please refer to:
- Feature documentation: `docs/feature/desc.md`
- Frappe Forum: https://discuss.frappe.io
- GitHub Issues: [Your repository URL]

