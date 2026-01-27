app_name = "m365email"
app_title = "M365Email"
app_publisher = "TierneyMorris Pty Ltd"
app_description = "M365 Email Integration"
app_email = "support@tierneymorris.com.au"
app_license = "mit"

# Apps
# ------------------

# required_apps = []

# Each item in the list will be shown as an app in the apps page
# add_to_apps_screen = [
# 	{
# 		"name": "m365email",
# 		"logo": "/assets/m365email/logo.png",
# 		"title": "M365Email",
# 		"route": "/m365email",
# 		"has_permission": "m365email.api.permission.has_app_permission"
# 	}
# ]

# Includes in <head>
# ------------------

# include js, css files in header of desk.html
# app_include_css = "/assets/m365email/css/m365email.css"
app_include_js = "/assets/m365email/js/email_queue_list.js"

# include js, css files in header of web template
# web_include_css = "/assets/m365email/css/m365email.css"
# web_include_js = "/assets/m365email/js/m365email.js"

# include custom scss in every website theme (without file extension ".scss")
# website_theme_scss = "m365email/public/scss/website"

# include js, css files in header of web form
# webform_include_js = {"doctype": "public/js/doctype.js"}
# webform_include_css = {"doctype": "public/css/doctype.css"}

# include js in page
# page_js = {"page" : "public/js/file.js"}

# include js in doctype views
doctype_js = {
	"Email Account": "public/js/email_account.js"
}
doctype_list_js = {
	"Email Queue": "public/js/email_queue_list.js"
}
# doctype_tree_js = {"doctype" : "public/js/doctype_tree.js"}
# doctype_calendar_js = {"doctype" : "public/js/doctype_calendar.js"}

# Svg Icons
# ------------------
# include app icons in desk
# app_include_icons = "m365email/public/icons.svg"

# Home Pages
# ----------

# application home page (will override Website Settings)
# home_page = "login"

# website user home page (by Role)
# role_home_page = {
# 	"Role": "home_page"
# }

# Generators
# ----------

# automatically create page for each record of this doctype
# website_generators = ["Web Page"]

# Jinja
# ----------

# add methods and filters to jinja environment
# jinja = {
# 	"methods": "m365email.utils.jinja_methods",
# 	"filters": "m365email.utils.jinja_filters"
# }

# Installation
# ------------

# before_install = "m365email.install.before_install"
after_install = "m365email.m365email.custom_fields.create_m365_custom_fields"

# Migration
# ---------
# Create/update custom fields on bench migrate
after_migrate = "m365email.m365email.custom_fields.create_m365_custom_fields"

# Boot Session
# ------------
# No longer needed - M365 accounts are now integrated into Email Account
# with service='M365' and appear automatically in Frappe's email lists
# boot_session = "m365email.m365email.boot_session"

# Uninstallation
# ------------

# before_uninstall = "m365email.uninstall.before_uninstall"
# after_uninstall = "m365email.uninstall.after_uninstall"

# Integration Setup
# ------------------
# To set up dependencies/integrations with other apps
# Name of the app being installed is passed as an argument

# before_app_install = "m365email.utils.before_app_install"
# after_app_install = "m365email.utils.after_app_install"

# Integration Cleanup
# -------------------
# To clean up dependencies/integrations with other apps
# Name of the app being uninstalled is passed as an argument

# before_app_uninstall = "m365email.utils.before_app_uninstall"
# after_app_uninstall = "m365email.utils.after_app_uninstall"

# Desk Notifications
# ------------------
# See frappe.core.notifications.get_notification_config

# notification_config = "m365email.notifications.get_notification_config"

# Permissions
# -----------
# Permissions evaluated in scripted ways

permission_query_conditions = {
	"Communication": "m365email.m365email.permissions.get_communication_permission_query_conditions",
}

has_permission = {
	"M365 Email Account": "m365email.m365email.doctype.m365_email_account.m365_email_account.has_permission",
	"Communication": "m365email.m365email.permissions.has_communication_permission",
}

# DocType Class
# ---------------
# Override standard doctype classes

# override_doctype_class = {
# 	"ToDo": "custom_app.overrides.CustomToDo"
# }

# Document Events
# ---------------
# Hook on document methods and events

doc_events = {
	"Email Queue": {
		"before_insert": "m365email.m365email.send.intercept_email_queue"
	}
}

# Override Email Queue send method to skip SMTP for M365 emails
# Override Email Account to support M365 service type
override_doctype_class = {
	"Email Account": "m365email.m365email.email_account_override.M365EmailAccount",
	"Email Queue": "m365email.m365email.email_queue_override.M365EmailQueue"
}

# Scheduled Tasks
# ---------------

scheduler_events = {
	"cron": {
		# NOTE: Runs every minute, but tasks check configured interval on Service Principal
		"* * * * *": [
			"m365email.m365email.tasks.sync_all_email_accounts",
			"m365email.m365email.tasks.sync_all_calendar_events"
		]
		# M365 emails are now processed by Frappe's standard email queue
		# via our M365EmailQueue.send() override - no separate task needed
	},
	"hourly": [
		"m365email.m365email.tasks.refresh_all_tokens"
	],
	"daily": [
		"m365email.m365email.tasks.cleanup_old_logs",
		"m365email.m365email.tasks.validate_service_principals"
	]
}

# Testing
# -------

# before_tests = "m365email.install.before_tests"

# Overriding Methods
# ------------------------------
#
# Override Frappe's core email sending to support M365
override_whitelisted_methods = {
	"frappe.core.doctype.communication.email.make": "m365email.m365email.email_override.make"
}
#
# each overriding function accepts a `data` argument;
# generated from the base implementation of the doctype dashboard,
# along with any modifications made in other Frappe apps
# override_doctype_dashboards = {
# 	"Task": "m365email.task.get_dashboard_data"
# }

# exempt linked doctypes from being automatically cancelled
#
# auto_cancel_exempted_doctypes = ["Auto Repeat"]

# Ignore links to specified DocTypes when deleting documents
# -----------------------------------------------------------
# Allow deleting M365 Email Accounts even if they have:
# - Sync logs (historical data)
# - Sent emails in Email Queue
# - Synced emails in Communication

ignore_links_on_delete = ["M365 Email Sync Log", "Email Queue", "Communication"]

# Request Events
# ----------------
# before_request = ["m365email.utils.before_request"]
# after_request = ["m365email.utils.after_request"]

# Job Events
# ----------
# before_job = ["m365email.utils.before_job"]
# after_job = ["m365email.utils.after_job"]

# User Data Protection
# --------------------

# user_data_fields = [
# 	{
# 		"doctype": "{doctype_1}",
# 		"filter_by": "{filter_by}",
# 		"redact_fields": ["{field_1}", "{field_2}"],
# 		"partial": 1,
# 	},
# 	{
# 		"doctype": "{doctype_2}",
# 		"filter_by": "{filter_by}",
# 		"partial": 1,
# 	},
# 	{
# 		"doctype": "{doctype_3}",
# 		"strict": False,
# 	},
# 	{
# 		"doctype": "{doctype_4}"
# 	}
# ]

# Authentication and authorization
# --------------------------------

# auth_hooks = [
# 	"m365email.auth.validate"
# ]

# Automatically update python controller files with type annotations for this app.
# export_python_type_annotations = True

# default_log_clearing_doctypes = {
# 	"Logging DocType Name": 30  # days to retain logs
# }

