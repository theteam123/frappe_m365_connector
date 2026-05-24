// Copyright (c) 2025, TierneyMorris Pty Ltd and contributors
// For license information, please see license.txt

console.log("[M365Email] compose_email_autocomplete.js loading...");

/**
 * Fix for email autocomplete in Compose Email dialog
 *
 * The Frappe core CommunicationComposer has a bug where the autocomplete
 * for recipients/cc/bcc fields doesn't work properly because it uses
 * async get_data() incorrectly.
 */

(function() {
	"use strict";

	console.log("[M365Email] Checking for CommunicationComposer...");
	console.log("[M365Email] frappe.views exists:", !!frappe.views);
	console.log("[M365Email] CommunicationComposer exists:", !!(frappe.views && frappe.views.CommunicationComposer));

	if (!frappe.views || !frappe.views.CommunicationComposer) {
		console.warn("[M365Email] CommunicationComposer not found at load time, setting up delayed patch");

		// NOTE: Use a proxy to catch when CommunicationComposer is defined
		let originalDescriptor = Object.getOwnPropertyDescriptor(frappe.views || {}, 'CommunicationComposer');

		frappe.provide("frappe.views");

		let patchApplied = false;

		Object.defineProperty(frappe.views, 'CommunicationComposer', {
			get: function() {
				return this._CommunicationComposer;
			},
			set: function(value) {
				console.log("[M365Email] CommunicationComposer being set, applying patch...");
				this._CommunicationComposer = PatchCommunicationComposer(value);
				patchApplied = true;
			},
			configurable: true
		});

		// NOTE: If it was already set, apply the patch
		if (originalDescriptor && originalDescriptor.value) {
			frappe.views.CommunicationComposer = originalDescriptor.value;
		}

		return;
	}

	// NOTE: CommunicationComposer already exists, patch it now
	console.log("[M365Email] CommunicationComposer found, patching now...");
	frappe.views.CommunicationComposer = PatchCommunicationComposer(frappe.views.CommunicationComposer);
})();


function PatchCommunicationComposer(OriginalClass) {
	console.log("[M365Email] Patching CommunicationComposer class...");

	const PatchedClass = class extends OriginalClass {
		setup_multiselect_queries() {
			console.log("[M365Email] setup_multiselect_queries called");

			const me = this;
			const EMAIL_FIELDS = ["recipients", "cc", "bcc"];

			EMAIL_FIELDS.forEach((fieldName) => {
				const field = this.dialog.fields_dict[fieldName];
				if (!field) {
					console.log("[M365Email] Field not found:", fieldName);
					return;
				}

				console.log("[M365Email] Setting up autocomplete for:", fieldName);

				// NOTE: Set minimum characters to trigger autocomplete
				if (field.awesomplete) {
					field.awesomplete.minChars = 1;
				}

				// NOTE: Override the input handler to use proper async fetch
				field.$input.off("input.m365autocomplete").on("input.m365autocomplete", frappe.utils.debounce(function(e) {
					const inputValue = field.get_value() || "";
					const searchText = inputValue.match(/[^,\s]*$/)[0] || "";

					console.log("[M365Email] Input event, searchText:", searchText);

					if (searchText.length < 1) {
						return;
					}

					const args = { txt: searchText };

					if (me.frm?.events?.get_email_recipient_filters) {
						const extraFilters = me.frm.events.get_email_recipient_filters(
							me.frm,
							fieldName
						);
						if (extraFilters) {
							args.extra_filters = JSON.stringify(extraFilters);
						}
					}

					console.log("[M365Email] Fetching contacts with args:", args);

					frappe.call({
						method: "frappe.email.get_contact_list",
						args: args,
						callback: function(r) {
							console.log("[M365Email] Got response:", r.message);
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

			console.log("[M365Email] Autocomplete setup complete");
		}
	};

	console.log("[M365Email] CommunicationComposer patched successfully");
	return PatchedClass;
}
