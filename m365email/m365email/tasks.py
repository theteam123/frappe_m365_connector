# Copyright (c) 2025, Zeke Tierney and contributors
# For license information, please see license.txt

"""
Scheduled tasks for M365 Email Integration
"""

import frappe
from frappe import _
from frappe.utils import now_datetime, time_diff_in_seconds
from m365email.m365email.sync import sync_email_account
from m365email.m365email.event_sync import sync_calendar_events
from m365email.m365email.auth import refresh_token, test_connection


# NOTE: Default intervals used when Service Principal settings are not configured
DEFAULT_EMAIL_SYNC_INTERVAL_MINUTES = 5
DEFAULT_CALENDAR_SYNC_INTERVAL_MINUTES = 5


def ShouldSyncServicePrincipal(servicePrincipalName: str, syncType: str) -> bool:
	"""
	Check if enough time has passed since the last sync for a Service Principal.

	Args:
		servicePrincipalName: Name of the Service Principal
		syncType: Either 'email' or 'calendar'

	Returns:
		bool: True if sync should proceed, False if interval hasn't elapsed
	"""
	if not servicePrincipalName:
		return True

	try:
		spDoc = frappe.get_doc("M365 Email Service Principal Settings", servicePrincipalName)
	except frappe.DoesNotExistError:
		return True

	if syncType == "email":
		intervalMinutes = spDoc.email_sync_interval_minutes or DEFAULT_EMAIL_SYNC_INTERVAL_MINUTES
		lastSync = spDoc.last_email_sync
	else:
		intervalMinutes = spDoc.calendar_sync_interval_minutes or DEFAULT_CALENDAR_SYNC_INTERVAL_MINUTES
		lastSync = spDoc.last_calendar_sync

	# NOTE: Ensure minimum interval of 1 minute
	intervalMinutes = max(1, intervalMinutes)

	if not lastSync:
		return True

	secondsSinceLastSync = time_diff_in_seconds(now_datetime(), lastSync)
	intervalSeconds = intervalMinutes * 60

	return secondsSinceLastSync >= intervalSeconds


def UpdateServicePrincipalLastSync(servicePrincipalName: str, syncType: str) -> None:
	"""
	Update the last sync timestamp for a Service Principal.

	Args:
		servicePrincipalName: Name of the Service Principal
		syncType: Either 'email' or 'calendar'
	"""
	if not servicePrincipalName:
		return

	try:
		fieldName = "last_email_sync" if syncType == "email" else "last_calendar_sync"
		frappe.db.set_value(
			"M365 Email Service Principal Settings",
			servicePrincipalName,
			fieldName,
			now_datetime()
		)
		frappe.db.commit()
	except Exception as e:
		print(f"M365 Email: Failed to update last sync time for {servicePrincipalName}: {str(e)}")


def sync_all_email_accounts():
	"""
	Sync all M365 email accounts with incoming enabled.
	Respects the sync interval configured on each Service Principal.
	Scheduled to run every minute, but only syncs when interval has elapsed.
	"""
	print("M365 Email: Starting scheduled sync for all accounts with incoming enabled")

	successCount = 0
	failedCount = 0
	skippedCount = 0

	# NOTE: Track which Service Principals we've already synced this run
	syncedServicePrincipals = set()

	# Sync Email Accounts with service='M365'
	emailAccounts = frappe.get_all(
		"Email Account",
		filters={"service": "M365", "enable_incoming": 1},
		fields=["name", "email_account_name", "email_id", "m365_service_principal"]
	)

	if not emailAccounts:
		print("M365 Email: No accounts with incoming enabled found")
		return

	print(f"M365 Email: Found {len(emailAccounts)} Email Account(s) with service='M365'")

	for account in emailAccounts:
		accountName = account.email_account_name or account.name
		servicePrincipal = account.m365_service_principal

		# Check if we should sync based on the Service Principal's interval
		if servicePrincipal and servicePrincipal not in syncedServicePrincipals:
			if not ShouldSyncServicePrincipal(servicePrincipal, "email"):
				print(f"M365 Email: Skipping {accountName} - sync interval not elapsed")
				skippedCount += 1
				continue

		try:
			print(f"M365 Email: Syncing {accountName} ({account.email_id})")
			result = sync_email_account(account.name)

			if result.get("success"):
				successCount += 1
				print(
					f"M365 Email: Successfully synced {accountName} - "
					f"Fetched: {result.get('fetched', 0)}, "
					f"Created: {result.get('created', 0)}, "
					f"Updated: {result.get('updated', 0)}"
				)

				# NOTE: Mark this Service Principal as synced and update timestamp
				if servicePrincipal:
					syncedServicePrincipals.add(servicePrincipal)
					UpdateServicePrincipalLastSync(servicePrincipal, "email")
			else:
				failedCount += 1
				print(
					f"M365 Email: Failed to sync {accountName} - "
					f"Error: {result.get('message')}"
				)
		except Exception as e:
			failedCount += 1
			print(f"M365 Email: Exception syncing {accountName}: {str(e)}")
			frappe.log_error(
				title="M365 Email Sync Failed",
				message=f"Account: {accountName}\nEmail: {account.email_id}\n\nError: {str(e)}"
			)

	print(
		f"M365 Email: Scheduled sync completed - "
		f"Success: {successCount}, Failed: {failedCount}, Skipped: {skippedCount}"
	)


def sync_all_calendar_events():
	"""
	Sync calendar events for all accounts with event sync enabled.
	Respects the sync interval configured on each Service Principal.
	Scheduled to run every minute, but only syncs when interval has elapsed.
	"""
	print("M365 Email: Starting scheduled calendar event sync for all accounts with event sync enabled")

	successCount = 0
	failedCount = 0
	skippedCount = 0

	# NOTE: Track which Service Principals we've already synced this run
	syncedServicePrincipals = set()

	# Sync Email Accounts with service='M365' and m365_sync_events enabled
	emailAccounts = frappe.get_all(
		"Email Account",
		filters={"service": "M365", "m365_sync_events": 1},
		fields=["name", "email_account_name", "email_id", "m365_service_principal"]
	)

	if not emailAccounts:
		print("M365 Email: No accounts with event sync enabled found")
		return

	print(f"M365 Email: Found {len(emailAccounts)} Email Account(s) with M365 event sync enabled")

	for account in emailAccounts:
		accountName = account.email_account_name or account.name
		servicePrincipal = account.m365_service_principal

		# Check if we should sync based on the Service Principal's interval
		if servicePrincipal and servicePrincipal not in syncedServicePrincipals:
			if not ShouldSyncServicePrincipal(servicePrincipal, "calendar"):
				print(f"M365 Email: Skipping calendar sync for {accountName} - sync interval not elapsed")
				skippedCount += 1
				continue

		try:
			print(f"M365 Email: Syncing calendar events for {accountName} ({account.email_id})")
			result = sync_calendar_events(account.name)

			if result.get("success"):
				successCount += 1
				print(
					f"M365 Email: Successfully synced events for {accountName} - "
					f"Fetched: {result.get('fetched', 0)}, "
					f"Created: {result.get('created', 0)}, "
					f"Updated: {result.get('updated', 0)}, "
					f"Deleted: {result.get('deleted', 0)}, "
					f"Skipped: {result.get('skipped', 0)}"
				)

				# NOTE: Mark this Service Principal as synced and update timestamp
				if servicePrincipal:
					syncedServicePrincipals.add(servicePrincipal)
					UpdateServicePrincipalLastSync(servicePrincipal, "calendar")
			else:
				failedCount += 1
				print(
					f"M365 Email: Failed to sync events for {accountName} - "
					f"Error: {result.get('message')}"
				)
		except Exception as e:
			failedCount += 1
			print(f"M365 Email: Exception syncing events for {accountName}: {str(e)}")
			frappe.log_error(
				title="M365 Event Sync Failed",
				message=f"Account: {accountName}\nEmail: {account.email_id}\n\nError: {str(e)}"
			)

	print(
		f"M365 Email: Calendar event sync completed - "
		f"Success: {successCount}, Failed: {failedCount}, Skipped: {skippedCount}"
	)


def refresh_all_tokens():
	"""
	Refresh tokens for all enabled service principals
	Scheduled to run hourly
	"""
	print("M365 Email: Starting token refresh for all service principals")

	# Get all enabled service principals
	service_principals = frappe.get_all(
		"M365 Email Service Principal Settings",
		filters={"enabled": 1},
		fields=["name", "service_principal_name"]
	)

	if not service_principals:
		print("M365 Email: No enabled service principals found")
		return

	print(f"M365 Email: Found {len(service_principals)} enabled service principal(s)")

	success_count = 0
	failed_count = 0

	for sp in service_principals:
		try:
			print(f"M365 Email: Refreshing token for {sp.service_principal_name}")
			success = refresh_token(sp.name)

			if success:
				success_count += 1
				print(f"M365 Email: Successfully refreshed token for {sp.service_principal_name}")
			else:
				failed_count += 1
				print(f"M365 Email: Failed to refresh token for {sp.service_principal_name}")
		except Exception as e:
			failed_count += 1
			print(f"M365 Email: Exception refreshing token for {sp.service_principal_name}: {str(e)}")
			frappe.log_error(
				title="M365 Token Refresh Failed",
				message=f"Service Principal: {sp.service_principal_name}\n\nError: {str(e)}"
			)

	print(
		f"M365 Email: Token refresh completed - "
		f"Success: {success_count}, Failed: {failed_count}"
	)


def cleanup_old_logs():
	"""
	Cleanup old sync logs (older than 30 days)
	Scheduled to run daily
	"""
	print("M365 Email: Starting cleanup of old sync logs")

	from frappe.utils import add_days, now_datetime

	# Delete logs older than 30 days
	cutoff_date = add_days(now_datetime(), -30)

	old_logs = frappe.get_all(
		"M365 Email Sync Log",
		filters={
			"creation": ["<", cutoff_date]
		},
		pluck="name"
	)

	if not old_logs:
		print("M365 Email: No old logs to cleanup")
		return

	print(f"M365 Email: Deleting {len(old_logs)} old log(s)")

	for log_name in old_logs:
		try:
			frappe.delete_doc("M365 Email Sync Log", log_name, ignore_permissions=True)
		except Exception as e:
			print(f"M365 Email: Failed to delete log {log_name}: {str(e)}")

	frappe.db.commit()

	print("M365 Email: Cleanup completed")


def validate_service_principals():
	"""
	Validate all service principal credentials
	Scheduled to run daily
	"""
	print("M365 Email: Starting service principal validation")

	# Get all enabled service principals
	service_principals = frappe.get_all(
		"M365 Email Service Principal Settings",
		filters={"enabled": 1},
		fields=["name", "service_principal_name"]
	)

	if not service_principals:
		print("M365 Email: No enabled service principals to validate")
		return

	print(f"M365 Email: Validating {len(service_principals)} service principal(s)")

	valid_count = 0
	invalid_count = 0

	for sp in service_principals:
		try:
			print(f"M365 Email: Validating {sp.service_principal_name}")
			result = test_connection(sp.name)

			if result.get("success"):
				valid_count += 1
				print(f"M365 Email: {sp.service_principal_name} is valid")
			else:
				invalid_count += 1
				print(
					f"M365 Email: {sp.service_principal_name} validation failed - "
					f"Error: {result.get('message')}"
				)

				# Log error for admin attention
				frappe.log_error(
					title="M365 Service Principal Invalid",
					message=f"Service Principal: {sp.service_principal_name}\n\nError: {result.get('message')}"
				)
		except Exception as e:
			invalid_count += 1
			print(f"M365 Email: Exception validating {sp.service_principal_name}: {str(e)}")
			frappe.log_error(
				title="M365 Service Principal Validation Failed",
				message=f"Service Principal: {sp.service_principal_name}\n\nError: {str(e)}"
			)

	print(
		f"M365 Email: Validation completed - "
		f"Valid: {valid_count}, Invalid: {invalid_count}"
	)

