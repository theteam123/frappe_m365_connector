// Copyright (c) 2026, TierneyMorris Pty Ltd and contributors
// For license information, please see license.txt

// Replaces core's "Create User Email" handler (the button calls frm.events.create_user_email,
// which resolves to the last registered handler) so a new account opens pre-filled as M365.

const CORE_HAS_EMAIL_ACCOUNT = "frappe.core.doctype.user.user.has_email_account";
const M365_ACCOUNT_DEFAULTS = "m365email.m365email.provisioning.GetUserEmailAccountDefaults";

// Frappe's own prefill, kept for users whose domain has no M365 Service Principal
const CORE_EMAIL_DEFAULTS = { awaiting_password: 1, enable_incoming: 1 };

frappe.ui.form.on("User", {
	create_user_email: async function (frm) {
		const existing = await frappe.xcall(CORE_HAS_EMAIL_ACCOUNT, { email: frm.doc.email });
		if (Array.isArray(existing) && existing.length) {
			frappe.route_flags.create_user_account = frm.doc.name;
			frappe.set_route("Form", "Email Account", existing[0].name);
			return;
		}

		const m365Defaults = await frappe.xcall(M365_ACCOUNT_DEFAULTS, { user: frm.doc.name });
		const hasM365Defaults = Boolean(m365Defaults) && Object.keys(m365Defaults).length > 0;
		const coreDefaults = { email_id: frm.doc.email, ...CORE_EMAIL_DEFAULTS };
		frappe.route_options = hasM365Defaults ? m365Defaults : coreDefaults;

		frappe.model.with_doctype("Email Account", function () {
			const newAccount = frappe.model.get_new_doc("Email Account");
			frappe.route_flags.linked_user = frm.doc.name;
			frappe.route_flags.delete_user_from_locals = true;
			frappe.set_route("Form", "Email Account", newAccount.name);
		});
	},
});
