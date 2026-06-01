/**
 * M365 Email Feature Guide
 *
 * Auto-generated feature reference for the `m365email` app, rendered in the same
 * sticky-TOC + scroll-spy format as the other Team Group workflow / feature
 * guides. Fully self-contained: all helpers are closure-scoped to this page so
 * the per-app guide pages never collide. Section data is embedded below.
 *
 * Static reference page. No backend calls.
 */

frappe.pages["m365email-feature-guide"].on_page_load = function(wrapper) {
	// NOTE: CONFIG is generated from a source audit of the app and injected here.
	const CONFIG = {"slug": "m365email-feature-guide", "title": "M365 Email Feature Guide", "subtitle": "A complete reference to the <strong>M365 Email</strong> app — its doctypes, pages, tools, integrations and customizations.", "laneOrder": ["M365Email", "M365email", "Pages & Tools", "Workspaces", "Integrations & Hooks", "Client-side", "Configuration"], "sections": [{"lane": "M365Email", "anchor": "m365-email-service-principal-settings", "title": "M365 Email Service Principal Settings", "status": "live", "bodyHTML": "<p>Stores the Azure AD <strong>app registration credentials</strong> used to authenticate to Microsoft Graph via the MSAL client-credentials flow (no user passwords). One record per tenant; multiple service principals are supported. This is the root of all M365 connectivity — email accounts link to a service principal, and the scheduled sync/send paths acquire tokens from here.</p><p>Open it at <a href=\"/app/m365-email-service-principal-settings\">M365 Email Service Principal Settings</a>. Permissions: <strong>System Manager only</strong>.</p><h4>Credential fields</h4><table class=\"tfg-fields\"><thead><tr><th>Field</th><th>Type</th><th>Notes</th></tr></thead><tbody><tr><td>service_principal_name</td><td>Data</td><td>Required, unique. Used as the record name (autoname <code>field:service_principal_name</code>).</td></tr><tr><td>tenant_id</td><td>Data</td><td>Required. Azure AD Directory (tenant) ID.</td></tr><tr><td>tenant_name</td><td>Data</td><td>Optional label.</td></tr><tr><td>client_id</td><td>Data</td><td>Required. Application (client) ID.</td></tr><tr><td>client_secret</td><td>Password</td><td>Required. Encrypted at rest.</td></tr><tr><td>authority_url</td><td>Data</td><td>Auto-populated as <code>https://login.microsoftonline.com/{tenant_id}</code> by <code>validate()</code> if blank or still containing the <code>{tenant_id}</code> placeholder.</td></tr><tr><td>graph_api_endpoint</td><td>Data</td><td>Defaults to <code>https://graph.microsoft.com/v1.0</code>.</td></tr><tr><td>scopes</td><td>Small Text</td><td>Defaults to <code>https://graph.microsoft.com/.default</code>.</td></tr><tr><td>enabled</td><td>Check</td><td>Off by default. Only enabled principals are picked up by scheduled tasks and the dropdown API.</td></tr></tbody></table><h4>Sync interval fields</h4><table class=\"tfg-fields\"><thead><tr><th>Field</th><th>Default</th><th>Notes</th></tr></thead><tbody><tr><td>email_sync_interval_minutes</td><td>5</td><td>How often incoming email sync runs. Minimum enforced is 1 minute.</td></tr><tr><td>calendar_sync_interval_minutes</td><td>5</td><td>How often calendar event sync runs. Minimum 1 minute.</td></tr><tr><td>last_email_sync / last_calendar_sync</td><td>—</td><td>Read-only timestamps updated by the scheduler after each successful run; gate whether the next minute's run actually executes.</td></tr></tbody></table><h4>Auto-provisioning (collapsible)</h4><table class=\"tfg-fields\"><thead><tr><th>Field</th><th>Notes</th></tr></thead><tbody><tr><td>enable_auto_provision</td><td>When on, an M365-backed Email Account is auto-created for senders in the configured domain who hold the <code>M365 User</code> role when they first send.</td></tr><tr><td>domain</td><td>Domain to match (e.g. <code>tierneymorris.com.au</code>). Shown only when auto-provision is on.</td></tr><tr><td>default_footer</td><td>Footer template applied to auto-provisioned accounts; supports a <code>&lt;!--Sender--&gt;</code> placeholder for the user's full name.</td></tr></tbody></table><h4>Token cache (read-only)</h4><p><code>token_cache</code>, <code>last_token_refresh</code>, and <code>token_expires_at</code> hold the MSAL token state. <code>on_update()</code> wipes all three whenever <code>client_id</code>, <code>client_secret</code>, or <code>tenant_id</code> changes, forcing re-authentication.</p><h4>Test Connection</h4><p>For saved records the form shows a <strong>Test Connection</strong> button (client script) that calls the API below and reports the acquired token's expiry. The form also auto-fixes the authority URL on load and whenever <code>tenant_id</code> changes.</p><table class=\"tfg-fields\"><thead><tr><th>API method</th><th>Access</th><th>Purpose</th></tr></thead><tbody><tr><td>m365email.m365email.api.test_service_principal_connection</td><td>System Manager</td><td>Acquires a token via <code>auth.test_connection</code> and returns success + expiry, or the error message.</td></tr></tbody></table><div class=\"tfg-callout tfg-callout-warn\"><strong>Staging:</strong> integrations are auto-disabled after a production refresh. Re-enable the service principal here before testing on staging.</div><h4>Key files</h4><ul><li>m365email/doctype/m365_email_service_principal_settings/m365_email_service_principal_settings.json</li><li>m365email/doctype/m365_email_service_principal_settings/m365_email_service_principal_settings.py</li><li>m365email/doctype/m365_email_service_principal_settings/m365_email_service_principal_settings.js</li><li>m365email/auth.py (token acquisition, refresh, encryption)</li></ul>"}, {"lane": "M365Email", "anchor": "m365-email-account", "title": "M365 Email Account (legacy / do-not-use)", "status": "mixed", "bodyHTML": "<p>This DocType configured a single mailbox (user or shared) for M365 sync. <strong>It is now deprecated</strong> — the form opens with a prominent red banner reading \"M365 EMAIL ACCOUNT - DO NOT USE\", directing admins to configure M365 mailboxes on the standard <a href=\"/app/email-account\">Email Account</a> instead (with <code>service = M365</code> and the <code>m365_*</code> custom fields). The DocType is retained for backwards compatibility.</p><p>Some server-side logic still references it as a fallback (see <code>has_permission</code> below and the validation rules). Permissions: <strong>System Manager</strong> (full) and <strong>M365 User</strong> (read/write).</p><h4>Notable fields (legacy schema)</h4><table class=\"tfg-fields\"><thead><tr><th>Field</th><th>Type</th><th>Notes</th></tr></thead><tbody><tr><td>account_name</td><td>Data</td><td>Required, unique, record name.</td></tr><tr><td>account_type</td><td>Select</td><td><code>User Mailbox</code> or <code>Shared Mailbox</code>.</td></tr><tr><td>email_address / user</td><td>Data / Link User</td><td>Both required. For User Mailbox, a mismatch between the address and the user's known emails raises an orange alert (warning, not blocking).</td></tr><tr><td>role</td><td>Link Role</td><td>Optional. Grants that role's holders access (used for shared mailboxes).</td></tr><tr><td>service_principal</td><td>Link</td><td>Required → M365 Email Service Principal Settings.</td></tr><tr><td>enable_incoming / enable_outgoing / sync_events</td><td>Check</td><td>Toggle Graph permissions in use (Mail.Read / Mail.Send / Calendars.Read).</td></tr><tr><td>default_outgoing</td><td>Check</td><td>Fallback sender when no account matches. Requires <code>enable_outgoing</code>; only one account may be the default (both enforced in <code>validate()</code>).</td></tr><tr><td>footer</td><td>Text Editor</td><td>Appended to outgoing mail.</td></tr><tr><td>sync_from_date / user_timezone</td><td>Date / Select</td><td>Timezone defaults to <code>Australia/Perth</code>.</td></tr><tr><td>auto_create_contact / sync_attachments / max_attachment_size</td><td>Check / Check / Int</td><td>Contact creation, attachment sync, size cap (default 10 MB).</td></tr><tr><td>folder_filter</td><td>Table → M365 Email Folder Filter</td><td>Per-folder sync config.</td></tr><tr><td>last_sync_time / last_sync_status / sync_error_message</td><td>read-only</td><td>Sync state.</td></tr><tr><td>delta_tokens</td><td>Long Text (JSON)</td><td>Per-folder Graph delta cursors.</td></tr></tbody></table><h4>Validation rules (validate)</h4><ul><li>User is always mandatory.</li><li>Duplicate <code>email_address</code> under the same service principal is blocked.</li><li><code>default_outgoing</code> requires <code>enable_outgoing</code>; only one default-outgoing account is allowed.</li></ul><h4>Custom permission (has_permission)</h4><p>System Manager sees all. A User-Mailbox owner sees their own. A user holding the account's assigned <code>role</code> gets access. Otherwise shared mailboxes are System-Manager-only. Registered in hooks under <code>has_permission</code> for \"M365 Email Account\".</p><div class=\"tfg-callout tfg-callout-info\"><strong>Tip:</strong> New M365 mailboxes should be created as standard Email Accounts with <code>service = M365</code>; the API and schedulers operate on those, not on this legacy DocType.</div><h4>Key files</h4><ul><li>m365email/doctype/m365_email_account/m365_email_account.json</li><li>m365email/doctype/m365_email_account/m365_email_account.py</li><li>m365email/doctype/m365_email_account/m365_email_account.js (empty/commented)</li></ul>"}, {"lane": "M365email", "anchor": "m365-email-folder-filter", "title": "M365 Email Folder Filter (child table)", "status": "live", "bodyHTML": "<p>Child table embedded in M365 Email Account (and mirrored on the standard Email Account via the <code>m365_folder_filter</code> custom field). One row per Outlook mail folder, controlling whether that folder is synced and storing its delta cursor. Note the module value on this DocType is <code>M365email</code> (lowercase 'e') unlike the parent.</p><h4>Fields</h4><table class=\"tfg-fields\"><thead><tr><th>Field</th><th>Type</th><th>Notes</th></tr></thead><tbody><tr><td>folder_name</td><td>Data</td><td>Required. Outlook folder display name (e.g. <code>Inbox</code>, <code>Sent Items</code>).</td></tr><tr><td>sync_enabled</td><td>Check</td><td>Default on. When off the folder is skipped during sync.</td></tr><tr><td>last_sync_time</td><td>Datetime</td><td>Read-only.</td></tr><tr><td>delta_token</td><td>Data</td><td>Read-only. Per-folder Graph delta token for incremental sync.</td></tr></tbody></table><p>The controller class <code>M365EmailFolderFilter</code> is a pass-through (no custom logic). Rows are populated and updated programmatically via the folder APIs on the parent account (see <code>get_available_folders</code> / <code>update_folder_filters</code>). When sync is enabled via the API the default seed is Inbox=on, Sent Items=off.</p><h4>Key files</h4><ul><li>m365email/doctype/m365_email_folder_filter/m365_email_folder_filter.json</li><li>m365email/doctype/m365_email_folder_filter/m365_email_folder_filter.py</li></ul>"}, {"lane": "M365Email", "anchor": "m365-email-sync-log", "title": "M365 Email Sync Log", "status": "live", "bodyHTML": "<p>Read-only <strong>audit trail</strong> of every sync operation (scheduled or manual). Named <code>SYNC-LOG-{#####}</code>. View at <a href=\"/app/m365-email-sync-log\">M365 Email Sync Log</a>. Permissions: <strong>System Manager</strong> read/report/print/export/share only (no create/write/delete in the UI — records are written by the sync engine).</p><h4>Fields</h4><table class=\"tfg-fields\"><thead><tr><th>Field</th><th>Type</th><th>Notes</th></tr></thead><tbody><tr><td>email_account</td><td>Link → Email Account</td><td>Required. The mailbox synced.</td></tr><tr><td>sync_type</td><td>Select</td><td><code>Full Sync</code> / <code>Delta Sync</code> / <code>Manual Sync</code>.</td></tr><tr><td>status</td><td>Select</td><td><code>In Progress</code> / <code>Success</code> / <code>Failed</code> / <code>Partial Success</code>.</td></tr><tr><td>start_time / end_time / duration</td><td>Datetime / Float</td><td>Duration in seconds (read-only).</td></tr><tr><td>messages_fetched / created / updated / failed</td><td>Int</td><td>Per-run statistics.</td></tr><tr><td>error_message</td><td>Text</td><td>Top-level error.</td></tr><tr><td>details</td><td>Long Text (JSON)</td><td>Structured per-message detail.</td></tr></tbody></table><p>The controller class is a pass-through. The daily <code>cleanup_old_logs</code> task deletes logs older than <strong>30 days</strong>.</p><div class=\"tfg-callout tfg-callout-info\"><strong>Tip:</strong> The status indicators map naturally to brand pills — <span class=\"tfg-pill tfg-pill-green\">Success</span> <span class=\"tfg-pill tfg-pill-red\">Failed</span> <span class=\"tfg-pill tfg-pill-yellow\">Partial Success</span> <span class=\"tfg-pill tfg-pill-grey\">In Progress</span>.</div><h4>Key files</h4><ul><li>m365email/doctype/m365_email_sync_log/m365_email_sync_log.json</li><li>m365email/doctype/m365_email_sync_log/m365_email_sync_log.py</li><li>m365email/sync.py and m365email/event_sync.py (write log rows)</li></ul>"}, {"lane": "M365Email", "anchor": "m365-email-settings", "title": "M365 Email Settings (Single — privacy filter)", "status": "live", "bodyHTML": "<p>Single DocType holding one admin toggle: the <strong>Email Privacy Filter</strong>. Open at <a href=\"/app/m365-email-settings\">M365 Email Settings</a>. Permissions: <strong>System Manager</strong> (read/write/create/print/email; no delete/export/report/share).</p><h4>Field</h4><table class=\"tfg-fields\"><thead><tr><th>Field</th><th>Type</th><th>Notes</th></tr></thead><tbody><tr><td>enable_email_privacy_filter</td><td>Check</td><td>Default <strong>off</strong>. When on, each synced email Communication is visible only to its linked user.</td></tr></tbody></table><h4>What the filter does</h4><p>When enabled it drives the two permission hooks on <strong>Communication</strong> (see Integrations &amp; Hooks). Rules, per <code>permissions.py</code>:</p><ul><li>Non-email Communications: unaffected (standard Frappe permissions).</li><li>Email Communications with a <code>linked_user</code>: visible <strong>only</strong> to that user — even System Managers are blocked. Only the <code>Administrator</code> user bypasses this (a Frappe core limitation noted in the code).</li><li>Email Communications with no <code>linked_user</code> (legacy data, shared mailboxes): visible to all users with normal Communication permissions.</li></ul><p>When the toggle is off, both permission functions early-return (no filtering) so behaviour is identical to stock Frappe.</p><div class=\"tfg-callout tfg-callout-warn\"><strong>Warn:</strong> Turning this on hides other users' synced email from System Managers. Confirm this is intended before enabling.</div><h4>Key files</h4><ul><li>m365email/doctype/m365_email_settings/m365_email_settings.json</li><li>m365email/doctype/m365_email_settings/m365_email_settings.py (pass-through)</li><li>m365email/permissions.py (filter logic)</li></ul>"}, {"lane": "Workspaces", "anchor": "email-settings-workspace", "title": "Email Settings workspace", "status": "live", "bodyHTML": "<p>Public desk workspace titled <strong>Email Settings</strong> (icon <code>mail</code>, blue indicator) that groups standard email config and the M365 doctypes into two cards.</p><h4>Card: Email</h4><ul><li><a href=\"/app/email-account\">Email Account</a></li><li><a href=\"/app/email-queue\">Email Queue</a></li><li><a href=\"/app/email-template\">Email Template</a></li><li><a href=\"/app/notification\">Notification</a></li><li><a href=\"/app/communication\">Communication</a></li></ul><h4>Card: Microsoft 365</h4><ul><li><a href=\"/app/m365-email-settings\">M365 Email Settings</a></li><li><a href=\"/app/m365-email-service-principal-settings\">M365 Email Service Principal Settings</a></li><li><a href=\"/app/m365-email-sync-log\">M365 Email Sync Log</a></li></ul><p>The legacy <a href=\"/app/m365-email-account\">M365 Email Account</a> is deliberately <em>not</em> linked here, reflecting its do-not-use status.</p><h4>Key files</h4><ul><li>m365email/workspace/email_settings/email_settings.json</li></ul>"}, {"lane": "Integrations & Hooks", "anchor": "hooks-and-pipeline", "title": "Email pipeline hooks & overrides", "status": "live", "bodyHTML": "<p>The app intercepts Frappe's standard email pipeline so M365 mailboxes work transparently through the normal Email Account / Email Queue / Communication flow. All wiring lives in <code>hooks.py</code>.</p><h4>Doctype class overrides</h4><table class=\"tfg-fields\"><thead><tr><th>Core DocType</th><th>Replaced by</th><th>Effect</th></tr></thead><tbody><tr><td>Email Account</td><td><code>email_account_override.ExtendedEmailAccount</code></td><td>Adds <code>M365</code> as a service type and bypasses SMTP/IMAP validation for M365 accounts.</td></tr><tr><td>Email Queue</td><td><code>email_queue_override.M365EmailQueue</code></td><td>Routes queued emails flagged <code>m365_send</code> through Microsoft Graph instead of SMTP on <code>send()</code>.</td></tr></tbody></table><h4>Document events</h4><table class=\"tfg-fields\"><thead><tr><th>DocType</th><th>Event</th><th>Handler</th></tr></thead><tbody><tr><td>Email Queue</td><td>before_insert</td><td><code>send.intercept_email_queue</code> — resolves the sending account for the sender; if an M365 account (or auto-provisioned one) matches, sets <code>m365_send = 1</code> and <code>m365_account</code> so the queue override sends via Graph.</td></tr></tbody></table><h4>Whitelisted method override</h4><p><code>frappe.core.doctype.communication.email.make</code> is overridden by <code>email_override.make</code> so the standard \"reply/new email\" composer routes through M365 when an M365 default-outgoing account is available (<code>can_send_via_m365</code>).</p><h4>Permission hooks (Communication)</h4><table class=\"tfg-fields\"><thead><tr><th>Hook</th><th>Function</th></tr></thead><tbody><tr><td>permission_query_conditions[\"Communication\"]</td><td><code>permissions.get_communication_permission_query_conditions</code></td></tr><tr><td>has_permission[\"Communication\"]</td><td><code>permissions.has_communication_permission</code></td></tr><tr><td>has_permission[\"M365 Email Account\"]</td><td><code>m365_email_account.has_permission</code></td></tr></tbody></table><p>Both Communication hooks are always registered but no-op unless the privacy filter (M365 Email Settings) is on.</p><h4>Install / migrate &amp; deletion safety</h4><ul><li><code>after_install</code> and <code>after_migrate</code> → <code>custom_fields.create_m365_custom_fields</code> (idempotently creates/updates all <code>m365_*</code> custom fields and adds the <code>M365</code> service option).</li><li><code>ignore_links_on_delete = [\"M365 Email Sync Log\", \"Email Queue\", \"Communication\"]</code> so M365 accounts can be deleted despite historical links.</li><li>No <code>fixtures</code>, <code>jinja</code>, <code>boot_session</code>, or website routes are configured (all commented out).</li></ul><h4>Client includes</h4><ul><li><code>app_include_js</code>: <code>public/js/email_queue_list.js</code> (global).</li><li><code>doctype_js</code>: Email Account → <code>public/js/email_account.js</code> (toggles M365 vs SMTP/IMAP fields by service).</li><li><code>doctype_list_js</code>: Email Queue → <code>public/js/email_queue_list.js</code> (adds the <strong>Send Now</strong> bulk action).</li></ul><h4>Key files</h4><ul><li>m365email/hooks.py</li><li>m365email/send.py, email_account_override.py, email_queue_override.py, email_override.py</li><li>m365email/permissions.py, custom_fields.py</li></ul>"}, {"lane": "Integrations & Hooks", "anchor": "scheduled-tasks", "title": "Scheduled tasks (sync, tokens, cleanup)", "status": "live", "bodyHTML": "<p>Five scheduled jobs in <code>tasks.py</code> drive the integration. The two sync jobs are scheduled <strong>every minute</strong> but self-throttle: each checks the per-service-principal interval (<code>email_sync_interval_minutes</code> / <code>calendar_sync_interval_minutes</code>, default 5, minimum 1) against <code>last_email_sync</code> / <code>last_calendar_sync</code> and skips if the interval hasn't elapsed. A service principal is synced at most once per scheduler run.</p><table class=\"tfg-fields\"><thead><tr><th>Schedule</th><th>Task</th><th>Purpose</th></tr></thead><tbody><tr><td><code>* * * * *</code></td><td>sync_all_email_accounts</td><td>Delta-sync incoming email for every Email Account with <code>service=M365</code> and <code>enable_incoming=1</code>.</td></tr><tr><td><code>* * * * *</code></td><td>sync_all_calendar_events</td><td>Delta-sync calendar events for accounts with <code>m365_sync_events=1</code>.</td></tr><tr><td>hourly</td><td>refresh_all_tokens</td><td>Refresh MSAL tokens for all enabled service principals before expiry.</td></tr><tr><td>daily</td><td>cleanup_old_logs</td><td>Delete M365 Email Sync Log rows older than 30 days.</td></tr><tr><td>daily</td><td>validate_service_principals</td><td>Run <code>test_connection</code> on every enabled principal; log errors for invalid credentials.</td></tr></tbody></table><p>Outgoing email needs no dedicated task — it flows through Frappe's standard email queue via the <code>M365EmailQueue.send()</code> override. Failures in the sync/validate loops are written to the Error Log via <code>frappe.log_error</code>.</p><h4>Key files</h4><ul><li>m365email/tasks.py</li><li>m365email/sync.py, event_sync.py, auth.py</li></ul>"}, {"lane": "Pages & Tools", "anchor": "api-surface", "title": "Whitelisted API surface", "status": "live", "bodyHTML": "<p>All endpoints live in <code>m365email.m365email.api</code>. None are CSRF-exempt or guest-accessible. Most operate on standard Email Accounts where <code>service=M365</code> and enforce <code>frappe.has_permission(\"Email Account\", ...)</code>.</p><table class=\"tfg-fields\"><thead><tr><th>Method</th><th>Access</th><th>Purpose</th></tr></thead><tbody><tr><td>enable_email_sync</td><td>Any (Shared Mailbox → System Manager)</td><td>Creates an Email Account with <code>service=M365</code>, incoming+outgoing on, and seeds folder filters (Inbox on, Sent Items off). Blocks duplicates.</td></tr><tr><td>disable_email_sync</td><td>Email Account write</td><td>Sets <code>enable_incoming=0</code> on an M365 account.</td></tr><tr><td>trigger_manual_sync</td><td>Email Account read</td><td>Runs <code>sync.sync_email_account</code> immediately.</td></tr><tr><td>trigger_manual_event_sync</td><td>Email Account read</td><td>Runs <code>event_sync.sync_calendar_events</code> immediately.</td></tr><tr><td>get_sync_status</td><td>Email Account read</td><td>For one account: returns the account dict + last 10 sync logs. With no argument: lists all M365 accounts with last sync time/status.</td></tr><tr><td>test_service_principal_connection</td><td>System Manager</td><td>Tests Azure AD credentials; returns token expiry. Backs the form button.</td></tr><tr><td>get_available_service_principals</td><td>Any</td><td>Lists enabled service principals for selection.</td></tr><tr><td>get_shared_mailboxes</td><td>System Manager (else empty list)</td><td>Lists M365 accounts of type Shared Mailbox.</td></tr><tr><td>get_available_folders</td><td>Email Account read</td><td>Calls Graph <code>get_mail_folders</code>; returns folder names with item/unread counts.</td></tr><tr><td>update_folder_filters</td><td>Email Account write</td><td>Replaces the account's <code>m365_folder_filter</code> child rows from supplied folder list.</td></tr><tr><td>SendEmailQueueNow</td><td>Email Queue permission</td><td>Immediately sends selected Email Queue rows (status Not Sent / Partially Sent / Error) via <code>queue.send(force_send=True)</code>; returns sent/failed counts and errors. Backs the list <strong>Send Now</strong> action.</td></tr></tbody></table><h4>Key files</h4><ul><li>m365email/api.py</li><li>m365email/auth.py, sync.py, event_sync.py, graph_api.py, utils.py</li></ul>"}, {"lane": "Configuration", "anchor": "custom-fields", "title": "Custom fields injected into core DocTypes", "status": "live", "bodyHTML": "<p><code>custom_fields.create_m365_custom_fields()</code> runs on install and on every <code>bench migrate</code>, idempotently adding M365 fields to three core DocTypes and registering <code>M365</code> as a selectable <code>service</code> on Email Account. This is what lets M365 mailboxes live on the standard <a href=\"/app/email-account\">Email Account</a> rather than the legacy doctype.</p><h4>Email Account</h4><p>Adds <code>linked_user</code>, <code>user_default</code>, an M365 settings section, <code>m365_service_principal</code> (Link → service principal), <code>m365_account_type</code>, <code>m365_sync_events</code>, sync-settings block (<code>m365_sync_from_date</code>, <code>m365_user_timezone</code>, <code>m365_auto_create_contact</code>, <code>m365_sync_attachments</code>, <code>m365_max_attachment_size</code>), <code>m365_folder_filter</code> (Table → M365 Email Folder Filter), and a status block (<code>m365_last_sync_time</code>, <code>m365_last_sync_status</code>, <code>m365_sync_error_message</code>, <code>m365_delta_tokens</code>). The <code>email_account.js</code> client script shows/hides these vs the SMTP/IMAP fields based on the selected service.</p><h4>Communication</h4><p>Adds <code>m365_message_id</code> (Graph message id, for deduplication), <code>m365_email_account</code> (Link → Email Account), and <code>linked_user</code> (Link → User) — the field the privacy filter keys on.</p><h4>Email Queue</h4><p>Adds <code>m365_send</code> (Check — set by the before_insert interceptor) and <code>m365_account</code> (Link → Email Account). The queue override reads these to route the send via Graph.</p><h4>Event</h4><p>Adds <code>m365_event_id</code>, <code>m365_email_account</code> (Link → Email Account), <code>m365_icaluid</code>, and <code>m365_timezone</code> — used by calendar event sync for dedup and timezone display.</p><div class=\"tfg-callout tfg-callout-info\"><strong>Tip:</strong> Because creation is idempotent and hooked to <code>after_migrate</code>, these fields self-heal after any migrate; no manual fixture import is needed.</div><h4>Key files</h4><ul><li>m365email/custom_fields.py</li><li>m365email/patches/add_email_queue_custom_fields.py, add_event_custom_fields.py, add_event_timezone_field.py, migrate_enable_fields.py</li></ul>"}, {"lane": "Configuration", "anchor": "m365-user-role", "title": "M365 User role & access model", "status": "live", "bodyHTML": "<p>The app ships one role, <strong>M365 User</strong> (desk-access enabled), for non-admin mailbox owners. It grants read/write on the legacy M365 Email Account doctype and is the role auto-provisioning checks before creating an account for a sender.</p><h4>Effective access summary</h4><table class=\"tfg-fields\"><thead><tr><th>Area</th><th>Who</th></tr></thead><tbody><tr><td>Service Principal Settings, Sync Log, M365 Email Settings, shared-mailbox APIs</td><td>System Manager only</td></tr><tr><td>Legacy M365 Email Account</td><td>System Manager (all); owner of a User Mailbox; holder of the account's assigned role</td></tr><tr><td>Standard M365 Email Accounts</td><td>Governed by Email Account permissions (the API checks these)</td></tr><tr><td>Synced email Communications</td><td>Standard Frappe perms, unless the privacy filter restricts each email to its <code>linked_user</code></td></tr></tbody></table><h4>Auto-provisioning</h4><p>When a service principal has <code>enable_auto_provision=1</code> and a matching <code>domain</code>, a sender in that domain holding the <strong>M365 User</strong> role gets an Email Account auto-created (<code>send.auto_provision_m365_account</code>) with <code>default_outgoing=1</code> the first time they send, so outgoing mail routes through Graph without manual setup.</p><h4>Key files</h4><ul><li>m365email/role/m365_user/m365_user.json</li><li>m365email/send.py (auto_provision_m365_account, get_sending_account_for_sender)</li><li>m365email/permissions.py</li></ul>"}, {"lane": "Client-side", "anchor": "client-scripts", "title": "Client scripts (Email Account form & Email Queue list)", "status": "live", "bodyHTML": "<p>The app injects two client scripts plus a global include via hooks.</p><h4>Email Account form (email_account.js)</h4><p>Loaded for the Email Account doctype. Toggles visibility of the M365 field group versus the standard SMTP/IMAP fields depending on the selected <code>service</code>, and applies per-service defaults when the service changes. Ensures an M365 account form shows only the relevant M365 settings.</p><h4>Email Queue list (email_queue_list.js)</h4><p>Loaded both globally (<code>app_include_js</code>) and as the Email Queue list script. It:</p><ul><li>Adds a <strong>Send Now</strong> actions-menu item to the Email Queue list that calls <code>api.SendEmailQueueNow</code> on the checked rows and reports sent/failed counts (guards against adding the button twice on refresh).</li><li>Patches a list-view prototype method related to <code>frappe.email.get_contact_list</code> (compose autocomplete) to fix a core behaviour.</li></ul><p>A third asset, <code>public/js/compose_email_autocomplete.js</code>, supports email compose autocomplete; it is part of the app's public JS but not directly registered in hooks.</p><div class=\"tfg-callout tfg-callout-info\"><strong>Tip:</strong> Changes to these <code>public/js</code> files are served after <code>bench build --app m365email</code> + cache clear.</div><h4>Key files</h4><ul><li>m365email/public/js/email_account.js</li><li>m365email/public/js/email_queue_list.js</li><li>m365email/public/js/compose_email_autocomplete.js</li></ul>"}]};

	const CONTAINER_ID = "tfg-m365email-feature-guide-container";
	const STYLE_ID = "tfg-guide-styles";
	const MAX_WIDTH = "1100px";
	const SIDEBAR_WIDTH = "300px";
	const UNKNOWN_LANE_RANK = 999;

	const PALETTE = [
		{ bg: "#f3e8ff", text: "#6b21a8", border: "#e9d5ff" },
		{ bg: "#dbeafe", text: "#1d4ed8", border: "#bfdbfe" },
		{ bg: "#dcfce7", text: "#15803d", border: "#bbf7d0" },
		{ bg: "#fef3c7", text: "#92400e", border: "#fde68a" },
		{ bg: "#fee2e2", text: "#b91c1c", border: "#fecaca" },
		{ bg: "#ccfbf1", text: "#0f766e", border: "#99f6e4" },
		{ bg: "#ffe4e6", text: "#9f1239", border: "#fecdd3" },
		{ bg: "#e0e7ff", text: "#3730a3", border: "#c7d2fe" },
		{ bg: "#ecfccb", text: "#3f6212", border: "#d9f99d" },
		{ bg: "#e5e7eb", text: "#374151", border: "#d1d5db" }
	];

	const page = frappe.ui.make_app_page({
		parent: wrapper,
		title: __(CONFIG.title),
		single_column: true
	});

	function GetLaneOrder() {
		if (CONFIG.laneOrder && CONFIG.laneOrder.length) return CONFIG.laneOrder;
		const seen = [];
		CONFIG.sections.forEach(s => { if (seen.indexOf(s.lane) === -1) seen.push(s.lane); });
		return seen;
	}

	function LaneColour(lane) {
		const PALETTE_FALLBACK_INDEX = 0;
		const order = GetLaneOrder();
		const index = order.indexOf(lane);
		const safeIndex = index === -1 ? PALETTE_FALLBACK_INDEX : index;
		return PALETTE[safeIndex % PALETTE.length];
	}

	function GetOrderedSections() {
		const order = GetLaneOrder();
		const rank = {};
		order.forEach((lane, i) => { rank[lane] = i; });
		const ordered = CONFIG.sections.slice();
		ordered.sort((a, b) => {
			const ra = rank[a.lane] === undefined ? UNKNOWN_LANE_RANK : rank[a.lane];
			const rb = rank[b.lane] === undefined ? UNKNOWN_LANE_RANK : rank[b.lane];
			return ra - rb;
		});
		return ordered;
	}

	function Render() {
		const pageBody = page.body[0] || page.body;
		pageBody.innerHTML = "";
		InjectStyles();

		const container = document.createElement("div");
		container.id = CONTAINER_ID;
		container.className = "tfg-container";

		const sections = GetOrderedSections();
		container.appendChild(BuildHeader(sections));
		container.appendChild(BuildLaneOverview(sections));
		container.appendChild(BuildLayout(sections));
		pageBody.appendChild(container);

		ActivateScrollSpy();
	}

	function BuildHeader(sections) {
		const header = document.createElement("div");
		header.className = "tfg-header";
		const laneCount = GetLaneOrder().filter(lane => sections.some(section => section.lane === lane)).length;
		header.innerHTML =
			'<h1 class="tfg-title">' + CONFIG.title + '</h1>' +
			'<p class="tfg-subtitle">' + CONFIG.subtitle +
			' <strong>' + sections.length + '</strong> documented features across <strong>' +
			laneCount + '</strong> areas.</p>';
		return header;
	}

	function BuildLaneOverview(sections) {
		const flow = document.createElement("div");
		flow.className = "tfg-status-flow";
		GetLaneOrder().forEach(lane => {
			const laneSections = sections.filter(s => s.lane === lane);
			if (!laneSections.length) return;
			flow.appendChild(BuildLaneRow(lane, laneSections));
		});
		return flow;
	}

	function BuildLaneRow(laneName, laneSections) {
		const row = document.createElement("div");
		row.className = "tfg-status-row";
		const colour = LaneColour(laneName);

		const label = document.createElement("div");
		label.className = "tfg-status-row-label";
		label.style.color = colour.text;
		label.textContent = laneName;
		row.appendChild(label);

		const chipsWrap = document.createElement("div");
		chipsWrap.className = "tfg-status-row-chips";
		laneSections.forEach(section => {
			const chip = document.createElement("a");
			chip.className = "tfg-status-chip";
			chip.href = "#" + section.anchor;
			chip.dataset.jump = section.anchor;
			chip.style.background = colour.bg;
			chip.style.color = colour.text;
			chip.style.border = "1px solid " + colour.border;
			chip.textContent = section.title;
			chipsWrap.appendChild(chip);
		});
		row.appendChild(chipsWrap);
		return row;
	}

	function BuildLayout(sections) {
		const layout = document.createElement("div");
		layout.className = "tfg-layout";
		layout.appendChild(BuildSidebar(sections));
		layout.appendChild(BuildContent(sections));
		return layout;
	}

	function BuildSidebar(sections) {
		const sidebar = document.createElement("nav");
		sidebar.className = "tfg-sidebar";
		const list = document.createElement("ul");
		list.className = "tfg-toc";
		let currentLane = null;
		sections.forEach((section, index) => {
			if (section.lane !== currentLane) {
				currentLane = section.lane;
				const laneHeading = document.createElement("div");
				laneHeading.className = "tfg-toc-lane-heading";
				laneHeading.textContent = currentLane;
				list.appendChild(laneHeading);
			}
			const item = document.createElement("li");
			item.className = "tfg-toc-item";
			const link = document.createElement("a");
			link.href = "#" + section.anchor;
			link.className = "tfg-toc-link";
			link.dataset.anchor = section.anchor;
			link.innerHTML =
				'<span class="tfg-toc-index">' + String(index + 1).padStart(2, "0") + '</span>' +
				'<span class="tfg-toc-title">' + section.title + '</span>';
			item.appendChild(link);
			list.appendChild(item);
		});
		sidebar.appendChild(list);
		return sidebar;
	}

	function BuildContent(sections) {
		const content = document.createElement("div");
		content.className = "tfg-content";
		let currentLane = null;
		sections.forEach((section, index) => {
			if (section.lane !== currentLane) {
				currentLane = section.lane;
				content.appendChild(BuildLaneDivider(currentLane));
			}
			content.appendChild(BuildSection(section, index + 1));
		});
		return content;
	}

	function BuildLaneDivider(laneName) {
		const colour = LaneColour(laneName);
		const divider = document.createElement("div");
		divider.className = "tfg-lane-divider";
		divider.style.color = colour.text;
		divider.style.borderColor = colour.border;
		divider.textContent = laneName;
		return divider;
	}

	function BuildSection(section, displayNumber) {
		const sectionEl = document.createElement("section");
		sectionEl.className = "tfg-section";
		sectionEl.id = section.anchor;

		const header = document.createElement("div");
		header.className = "tfg-section-header";
		let statusPill = '<span class="tfg-pill tfg-pill-green tfg-section-status">Live</span>';
		if (section.status === "coming-soon") {
			statusPill = '<span class="tfg-pill tfg-pill-grey tfg-section-status">Stub</span>';
		} else if (section.status === "mixed") {
			statusPill = '<span class="tfg-pill tfg-pill-yellow tfg-section-status">Partial</span>';
		}
		header.innerHTML =
			'<span class="tfg-section-number">' + String(displayNumber).padStart(2, "0") + '</span>' +
			'<h2 class="tfg-section-title">' + section.title + '</h2>' + statusPill;

		const body = document.createElement("div");
		body.className = "tfg-section-body";
		body.innerHTML = section.bodyHTML;

		sectionEl.appendChild(header);
		sectionEl.appendChild(body);
		return sectionEl;
	}

	function ActivateScrollSpy() {
		const jumpLinks = document.querySelectorAll(".tfg-toc-link, .tfg-status-chip");
		jumpLinks.forEach(link => {
			link.addEventListener("click", event => {
				const anchor = link.dataset.anchor || link.dataset.jump;
				if (!anchor) return;
				const target = document.getElementById(anchor);
				if (!target) return;
				event.preventDefault();
				target.scrollIntoView({ behavior: "smooth", block: "start" });
				SetActive(anchor);
			});
		});
		const sectionEls = document.querySelectorAll(".tfg-section");
		if (!sectionEls.length) return;
		const observer = new IntersectionObserver(entries => {
			entries.forEach(entry => {
				if (entry.isIntersecting) SetActive(entry.target.id);
			});
		}, { rootMargin: "-30% 0px -60% 0px" });
		sectionEls.forEach(sectionEl => observer.observe(sectionEl));
	}

	function SetActive(anchor) {
		document.querySelectorAll(".tfg-toc-link").forEach(link => {
			if (link.dataset.anchor === anchor) {
				link.classList.add("tfg-toc-link-active");
			} else {
				link.classList.remove("tfg-toc-link-active");
			}
		});
	}

	function InjectStyles() {
		if (document.getElementById(STYLE_ID)) return;
		const style = document.createElement("style");
		style.id = STYLE_ID;
		style.textContent = TFG_STYLES.replace(/__MAXW__/g, MAX_WIDTH).replace(/__SBW__/g, SIDEBAR_WIDTH);
		document.head.appendChild(style);
	}

	const TFG_STYLES = `
.tfg-container { max-width: __MAXW__; margin: 0 auto; padding: 24px 28px 64px; color: #1f2937; font-size: 14px; line-height: 1.55; }
.tfg-header { margin-bottom: 20px; }
.tfg-title { font-size: 26px; font-weight: 700; margin: 0 0 6px; color: #111827; }
.tfg-subtitle { margin: 0; color: #4b5563; font-size: 14px; max-width: 820px; }
.tfg-status-flow { display: flex; flex-direction: column; gap: 12px; padding: 16px 0 24px; border-bottom: 1px solid #e5e7eb; margin-bottom: 24px; }
.tfg-status-row { display: flex; align-items: flex-start; gap: 16px; flex-wrap: wrap; }
.tfg-status-row-label { min-width: 160px; padding-top: 6px; font-size: 12px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.06em; }
.tfg-status-row-chips { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; flex: 1 1 auto; }
.tfg-status-chip { padding: 6px 11px; border-radius: 8px; font-size: 12.5px; font-weight: 600; text-decoration: none; transition: filter 0.12s ease; }
.tfg-status-chip:hover { filter: brightness(0.96); text-decoration: none; }
.tfg-layout { display: grid; grid-template-columns: __SBW__ 1fr; gap: 32px; align-items: start; }
.tfg-sidebar { position: sticky; top: 16px; max-height: calc(100vh - 80px); overflow-y: auto; padding-right: 8px; }
.tfg-toc { list-style: none; padding: 0; margin: 0; }
.tfg-toc-item + .tfg-toc-item { margin-top: 2px; }
.tfg-toc-lane-heading { margin: 14px 0 4px; padding: 0 10px; font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.08em; color: #6b7280; }
.tfg-toc-lane-heading:first-child { margin-top: 0; }
.tfg-toc-link { display: flex; align-items: baseline; gap: 8px; padding: 7px 10px; border-radius: 6px; color: #4b5563; text-decoration: none; font-size: 13px; border-left: 3px solid transparent; }
.tfg-toc-link:hover { background: #f3f4f6; color: #111827; text-decoration: none; }
.tfg-toc-link-active { background: #eff6ff; color: #1d4ed8; border-left-color: #1d4ed8; font-weight: 600; }
.tfg-toc-index { font-family: ui-monospace, SFMono-Regular, Menlo, monospace; font-size: 11px; color: #9ca3af; min-width: 22px; }
.tfg-toc-link-active .tfg-toc-index { color: #1d4ed8; }
.tfg-content { min-width: 0; }
.tfg-lane-divider { margin: 28px 0 4px; padding: 8px 0 6px; font-size: 13px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.08em; border-bottom: 2px solid; }
.tfg-lane-divider:first-child { margin-top: 0; }
.tfg-section { padding: 22px 0 26px; border-bottom: 1px solid #e5e7eb; scroll-margin-top: 16px; }
.tfg-section-header { display: flex; align-items: baseline; gap: 12px; margin-bottom: 12px; flex-wrap: wrap; }
.tfg-section-number { font-family: ui-monospace, SFMono-Regular, Menlo, monospace; font-size: 12px; color: #9ca3af; background: #f3f4f6; padding: 3px 8px; border-radius: 6px; }
.tfg-section-title { font-size: 19px; font-weight: 700; margin: 0; color: #111827; flex: 1 1 auto; }
.tfg-section-status { align-self: center; }
.tfg-section-body p { margin: 0 0 12px; }
.tfg-section-body h4 { margin: 16px 0 8px; font-size: 13px; font-weight: 700; color: #111827; text-transform: uppercase; letter-spacing: 0.04em; }
.tfg-section-body ul, .tfg-section-body ol { margin: 0 0 12px; padding-left: 22px; }
.tfg-section-body li { margin-bottom: 4px; }
.tfg-section-body code { background: #f3f4f6; padding: 1px 5px; border-radius: 4px; font-size: 12.5px; color: #be185d; word-break: break-word; }
.tfg-section-body a { color: #1d4ed8; }
.tfg-section-body a:hover { text-decoration: underline; }
.tfg-fields { width: 100%; border-collapse: collapse; margin: 8px 0 12px; font-size: 13px; }
.tfg-fields th, .tfg-fields td { text-align: left; padding: 8px 12px; border-bottom: 1px solid #e5e7eb; vertical-align: top; }
.tfg-fields th { background: #f9fafb; font-weight: 600; color: #374151; }
.tfg-fields tr:last-child td { border-bottom: none; }
.tfg-callout { padding: 12px 14px; border-radius: 8px; margin: 14px 0 4px; font-size: 13px; border: 1px solid; }
.tfg-callout-info { background: #eff6ff; border-color: #bfdbfe; color: #1e3a8a; }
.tfg-callout-warn { background: #fffbeb; border-color: #fde68a; color: #78350f; }
.tfg-pill { padding: 3px 10px; border-radius: 999px; font-size: 12px; font-weight: 600; display: inline-block; }
.tfg-pill-grey { background: #f3f4f6; color: #4b5563; }
.tfg-pill-blue { background: #dbeafe; color: #1d4ed8; }
.tfg-pill-green { background: #dcfce7; color: #15803d; }
.tfg-pill-yellow { background: #fef3c7; color: #92400e; }
.tfg-pill-red { background: #fee2e2; color: #b91c1c; }
@media (max-width: 900px) {
	.tfg-layout { grid-template-columns: 1fr; }
	.tfg-sidebar { position: relative; top: auto; max-height: none; padding: 8px 0; border-bottom: 1px solid #e5e7eb; }
	.tfg-status-row { gap: 8px; }
	.tfg-status-row-label { min-width: auto; }
}
`;

	wrapper.tfgRender = Render;
	wrapper.tfgContainerId = CONTAINER_ID;
	Render();
};

frappe.pages["m365email-feature-guide"].on_page_show = function(wrapper) {
	if (typeof wrapper.tfgRender !== "function") return;
	if (!document.getElementById(wrapper.tfgContainerId)) wrapper.tfgRender();
};
