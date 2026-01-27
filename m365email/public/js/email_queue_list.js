// Email Queue List customizations for M365Email
// Adds "Send Now" bulk action to immediately process selected emails
// Also includes fix for email autocomplete in Compose Email dialog

console.log("[M365Email] Email Queue list JS loaded");

// ============================================================================
// EMAIL AUTOCOMPLETE FIX FOR COMPOSE DIALOG
// ============================================================================
// The Frappe core CommunicationComposer has a bug where the autocomplete
// for recipients/cc/bcc fields doesn't work properly.
// This fix patches the prototype method directly.

(function() {
	"use strict";

	let patchApplied = false;

	function applyPrototypePatch() {
		if (patchApplied) return true;
		if (!frappe.views || !frappe.views.CommunicationComposer) return false;

		const Composer = frappe.views.CommunicationComposer;

		// NOTE: Store original method
		const originalMethod = Composer.prototype.setup_multiselect_queries;

		// NOTE: Replace with our fixed version
		Composer.prototype.setup_multiselect_queries = function() {
			console.log("[M365Email] setup_multiselect_queries called (patched)");
			const me = this;
			const EMAIL_FIELDS = ["recipients", "cc", "bcc"];

			EMAIL_FIELDS.forEach((fieldName) => {
				const field = this.dialog.fields_dict[fieldName];
				if (!field) return;

				console.log("[M365Email] Setting up autocomplete for:", fieldName);

				// NOTE: Set minimum characters to trigger autocomplete
				if (field.awesomplete) {
					field.awesomplete.minChars = 1;
				}

				// NOTE: Override input handler for proper async fetch
				field.$input.off("input.m365autocomplete").on("input.m365autocomplete", frappe.utils.debounce(function() {
					const inputValue = field.get_value() || "";
					const searchText = inputValue.match(/[^,\s]*$/)[0] || "";

					console.log("[M365Email] Input detected, search:", searchText);

					if (searchText.length < 1) return;

					const args = { txt: searchText };

					if (me.frm?.events?.get_email_recipient_filters) {
						const extraFilters = me.frm.events.get_email_recipient_filters(me.frm, fieldName);
						if (extraFilters) {
							args.extra_filters = JSON.stringify(extraFilters);
						}
					}

					console.log("[M365Email] Fetching contacts...");
					frappe.call({
						method: "frappe.email.get_contact_list",
						args: args,
						callback: function(r) {
							console.log("[M365Email] Got contacts:", r.message);
							if (r.message && r.message.length) {
								field.set_data(r.message);
								if (field.awesomplete) {
									field.awesomplete.evaluate();
								}
							}
						}
					});
				}, 300));
			});
		};

		patchApplied = true;
		console.log("[M365Email] Email autocomplete prototype patch applied");
		return true;
	}

	// NOTE: Try immediately
	if (!applyPrototypePatch()) {
		// NOTE: Poll until CommunicationComposer is available
		console.log("[M365Email] CommunicationComposer not ready, polling...");
		let attempts = 0;
		const maxAttempts = 100;

		const interval = setInterval(function() {
			attempts++;
			if (applyPrototypePatch() || attempts >= maxAttempts) {
				clearInterval(interval);
				if (attempts >= maxAttempts && !patchApplied) {
					console.warn("[M365Email] Could not patch after", maxAttempts, "attempts");
				}
			}
		}, 100);
	}
})();

// ============================================================================
// EMAIL QUEUE LIST CUSTOMIZATIONS
// ============================================================================

frappe.provide("frappe.listview_settings");

// NOTE: Store reference to add button function for use in multiple places
function addSendNowButton(listView) {
	// NOTE: Only show for users with Email Queue write permission
	if (!frappe.perm.has_perm("Email Queue", 0, "write")) {
		return;
	}

	// NOTE: Prevent duplicate buttons on refresh
	if (listView._sendNowButtonAdded) {
		return;
	}
	listView._sendNowButtonAdded = true;

	listView.page.add_actions_menu_item(__("Send Now"), () => {
		const selectedItems = listView.get_checked_items(true);

		if (!selectedItems || selectedItems.length === 0) {
			frappe.msgprint(__("Please select at least one email to send."));
			return;
		}

		// NOTE: Filter to only sendable statuses
		const sendableStatuses = ["Not Sent", "Partially Sent", "Error"];
		const selectedDocs = listView.get_checked_items();
		const sendableItems = selectedDocs.filter(
			(doc) => sendableStatuses.includes(doc.status)
		);

		if (sendableItems.length === 0) {
			frappe.msgprint(
				__("None of the selected emails can be sent. Only emails with status 'Not Sent', 'Partially Sent', or 'Error' can be processed.")
			);
			return;
		}

		const sendableNames = sendableItems.map((doc) => doc.name);

		frappe.confirm(
			__("Send {0} email(s) now?", [sendableNames.length]),
			() => {
				frappe.show_alert({
					message: __("Processing {0} email(s)...", [sendableNames.length]),
					indicator: "blue"
				});

				frappe.call({
					method: "m365email.m365email.api.SendEmailQueueNow",
					args: {
						queues: sendableNames
					},
					callback: (response) => {
						if (response.exc) {
							frappe.msgprint({
								title: __("Error"),
								message: __("An error occurred while sending emails."),
								indicator: "red"
							});
							return;
						}

						const result = response.message;
						let message = __("Sent: {0}", [result.sent]);

						if (result.failed > 0) {
							message += "<br>" + __("Failed: {0}", [result.failed]);
						}

						if (result.errors && result.errors.length > 0) {
							message += "<br><br><strong>" + __("Errors:") + "</strong><ul>";
							result.errors.forEach((err) => {
								message += `<li>${err.name}: ${err.error}</li>`;
							});
							message += "</ul>";
						}

						const hasFailures = result.failed > 0;
						frappe.msgprint({
							title: hasFailures ? __("Completed with Errors") : __("Success"),
							message: message,
							indicator: hasFailures ? "orange" : "green"
						});

						listView.refresh();
					}
				});
			}
		);
	});
}


// NOTE: Hook into listview_settings onload to add our button
(function() {
	const existingSettings = frappe.listview_settings["Email Queue"] || {};
	const originalOnload = existingSettings.onload;
	const originalRefresh = existingSettings.refresh;

	frappe.listview_settings["Email Queue"] = Object.assign({}, existingSettings, {
		onload: function(listview) {
			if (originalOnload) {
				originalOnload.call(this, listview);
			}
			addSendNowButton(listview);
		},
		refresh: function(listview) {
			if (originalRefresh) {
				originalRefresh.call(this, listview);
			}
			// NOTE: Also try adding on refresh in case onload was missed
			addSendNowButton(listview);
		}
	});
})();
