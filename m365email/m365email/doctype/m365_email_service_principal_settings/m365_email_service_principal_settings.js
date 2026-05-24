// Copyright (c) 2025, Zeke Tierney and contributors
// For license information, please see license.txt

frappe.ui.form.on("M365 Email Service Principal Settings", {
	refresh(frm) {
		if (!frm.is_new()) {
			frm.add_custom_button(__("Test Connection"), function() {
				testConnection(frm);
			});
		}
		// NOTE: Fix authority URL placeholder on load if needed
		fixAuthorityUrl(frm);
	},


	tenant_id(frm) {
		// NOTE: Auto-update authority URL when tenant ID changes
		fixAuthorityUrl(frm);
	}
});


function testConnection(frm) {
	frappe.call({
		method: "m365email.m365email.api.test_service_principal_connection",
		args: {
			service_principal_name: frm.doc.name
		},
		freeze: true,
		freeze_message: __("Testing connection..."),
		callback: function(response) {
			let result = response.message;

			if (result.success) {
				frappe.msgprint({
					title: __("Connection Successful"),
					indicator: "green",
					message: __("Token acquired successfully.<br>Expires at: {0}", [result.token_expires_at])
				});
			} else {
				frappe.msgprint({
					title: __("Connection Failed"),
					indicator: "red",
					message: result.message
				});
			}
		}
	});
}


function fixAuthorityUrl(frm) {
	let tenantId = frm.doc.tenant_id;
	let authorityUrl = frm.doc.authority_url;

	if (!tenantId) {
		return;
	}

	// NOTE: Replace placeholder or set default if missing/contains placeholder
	let hasPlaceholder = authorityUrl && authorityUrl.includes("{tenant_id}");
	let isEmpty = !authorityUrl;

	if (hasPlaceholder || isEmpty) {
		let correctUrl = `https://login.microsoftonline.com/${tenantId}`;
		frm.set_value("authority_url", correctUrl);
	}
}
