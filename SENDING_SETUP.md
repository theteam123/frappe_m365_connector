# M365 Email Sending Setup Guide

## Overview

The M365 Email app can send emails via Microsoft Graph API instead of SMTP. This allows you to:
- Send emails as any user in your organization
- Send from shared mailboxes
- Bypass SMTP configuration entirely
- Use Microsoft's infrastructure for better deliverability

## How It Works

1. **M365 Email Accounts**: Sending uses the standard **Email Account** doctype with Service = "M365" and **Enable Outgoing** checked. Mark one as **Default Outgoing** to act as the fallback sender.
2. **Send As Any User**: The Service Principal has `Mail.Send` permission to send as any user
3. **Automatic Interception**: When Frappe creates an Email Queue entry, the `before_insert` hook marks it for M365 sending (`m365_send = 1`)
4. **Graph API Sending**: Frappe's standard email queue processing calls our `M365EmailQueue.send()` override, which sends via Microsoft Graph API instead of SMTP. There is no separate sending task.

## Azure AD Setup

### Required Permission

Your Service Principal needs the **`Mail.Send`** application permission:

1. Go to **Azure Portal** → **App Registrations** → Your App
2. Click **API Permissions**
3. Click **Add a permission** → **Microsoft Graph** → **Application permissions**
4. Search for and add: **`Mail.Send`**
5. Click **Grant admin consent**

### Permission Details

- **Mail.Send**: Allows the app to send mail as any user without a signed-in user

## Frappe Setup

### 1. Run Migration

After pulling the latest code, run migration to add custom fields:

```bash
bench --site your-site.local migrate
```

This adds:
- `m365_send` field to Email Queue
- `m365_account` field to Email Queue

### 2. Restart Bench

**IMPORTANT**: The override of Frappe's core email function requires a bench restart:

```bash
bench restart
```

Or if running in development mode, stop and restart:

```bash
# Press Ctrl+C to stop
bench start
```

### 3. Configure an Account for Sending

1. Go to **Email Account** list
2. Open (or create) the M365 account you want to send from (Service = "M365")
3. Check **Enable Outgoing**
4. Optionally check **Default Outgoing** to make it the fallback sender for emails whose sender does not match another M365 account
5. Set **Linked User** if this should be a specific user's default outgoing account
6. Save

**Note**: Multiple M365 accounts can send. The account is chosen per sender (matching sender email or the sender's linked account), falling back to the Default Outgoing M365 account.

### 4. Test Sending

Try sending an email from Frappe:

```python
# In bench console
frappe.sendmail(
    recipients=["test@example.com"],
    sender="your-email@domain.com",
    subject="Test M365 Send",
    message="This email was sent via Microsoft Graph API!"
)
```

Check the Email Queue:
- The email should have `m365_send = 1`
- It is processed by Frappe's standard email queue flush (our `M365EmailQueue.send()` override)
- Status should change to "Sent"

## How Sending Works

### Flow Diagram

```
User/System sends email
    ↓
Email Queue created
    ↓
before_insert hook checks for M365 sending account
    ↓
If found: Mark email with m365_send=1
    ↓
Frappe's standard email queue flush calls M365EmailQueue.send()
    ↓
Get access token from Service Principal
    ↓
Call Graph API: POST /users/{sender}/sendMail
    ↓
Email sent as the specified sender
    ↓
Update Email Queue status to "Sent"
```

### Send As Different Users

The beauty of this approach is that you can send as ANY user:

```python
# Send as user1@domain.com
frappe.sendmail(
    recipients=["recipient@example.com"],
    sender="user1@domain.com",
    subject="From User 1",
    message="This is sent as user1@domain.com"
)

# Send as shared mailbox
frappe.sendmail(
    recipients=["recipient@example.com"],
    sender="support@domain.com",
    subject="From Support",
    message="This is sent as support@domain.com"
)
```

As long as the Service Principal has `Mail.Send` permission, it can send as any of these users!

## Monitoring

### Check Email Queue

```sql
SELECT name, sender, recipients, status, m365_send, m365_account
FROM `tabEmail Queue`
WHERE m365_send = 1
ORDER BY creation DESC
LIMIT 10;
```

### Check Logs

```bash
# In bench console
frappe.get_all("Error Log", 
    filters={"error": ["like", "%M365 Email Send%"]},
    fields=["name", "creation", "error"],
    order_by="creation desc",
    limit=10
)
```

### Manual Processing

If you need to manually process the M365 email queue:

```python
# In bench console
from m365email.m365email.send import process_email_queue_m365

result = process_email_queue_m365()
print(result)  # {'sent': 5, 'failed': 0}
```

## Troubleshooting

### Email Not Marked for M365 Sending

**Problem**: Email Queue entries have `m365_send = 0`

**Solution**:
1. Check that you have an Email Account with Service = "M365", `enable_outgoing = 1`, and (for fallback) `default_outgoing = 1`
2. Restart bench after making changes to hooks.py

### Email Stuck in "Sending" Status

**Problem**: Email Queue status is "Sending" but never completes

**Solution**:
1. Check Error Log for M365 send errors
2. Verify Service Principal has `Mail.Send` permission
3. Check that access token is being generated correctly
4. Manually process the queue to see detailed errors

### Permission Denied Error

**Problem**: Error: "Insufficient privileges to complete the operation"

**Solution**:
1. Verify `Mail.Send` permission is added in Azure AD
2. Ensure admin consent is granted
3. Wait a few minutes for permission changes to propagate
4. Refresh the access token

### Sender Email Not Valid

**Problem**: Error: "The specified object was not found in the store"

**Solution**:
1. Verify the sender email exists in your M365 tenant
2. Check that the email address is correctly formatted
3. For shared mailboxes, ensure they're properly configured in M365

## Advanced Configuration

### Disable M365 Sending Temporarily

To temporarily disable M365 sending without removing the configuration:

1. Open the M365 Email Account used for sending
2. Uncheck **Enable Outgoing** (and **Default Outgoing**)
3. Save

New emails will then fall back to Frappe's standard (SMTP) sending.

### Fallback to SMTP

Currently, if M365 sending fails, the email is marked as "Error" and does NOT fall back to SMTP. This prevents duplicate sends.

If you want to retry via SMTP:
1. Find the failed email in Email Queue
2. Uncheck `m365_send`
3. Change status back to "Not Sent"
4. The standard email queue processor will send via SMTP

## Performance

- **Processing**: Emails are sent as part of Frappe's standard email queue flush (no separate M365 sending schedule)
- **Rate Limiting**: Microsoft Graph API has rate limits (handled automatically with retries)
- **Attachments**: Supported up to Microsoft's limits (typically 150 MB total message size)

## Security

- **Access Tokens**: Cached for 50 minutes, refreshed automatically
- **Client Secret**: Stored encrypted in Frappe using `get_password()`
- **Permissions**: Only System Managers can configure M365 Email Accounts
- **Audit Trail**: All sends are logged in Email Queue and Communication

## Next Steps

1. ✅ Run migration to add custom fields
2. ✅ Add `Mail.Send` permission in Azure AD
3. ✅ Mark one M365 Email Account for sending
4. ✅ Test sending an email
5. ✅ Monitor Email Queue and Error Log

Happy sending! 🚀

