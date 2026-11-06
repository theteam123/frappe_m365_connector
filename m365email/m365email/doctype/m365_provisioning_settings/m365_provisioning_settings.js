// Copyright (c) 2026, TierneyMorris Pty Ltd and contributors
// For license information, please see license.txt

const SETTINGS_TEST_METHOD = "m365email.m365email.entra.TestConnection";
const SIMULATION_MODE = "Simulation";

frappe.ui.form.on("M365 Provisioning Settings", {
	refresh(frm) {
		if (frm.doc.mode === SIMULATION_MODE) {
			frm.dashboard.set_headline_alert(
				__("Simulation mode: nothing is created in Microsoft 365."), "yellow"
			);
		}

		frm.add_custom_button(__("Test Connection"), () => {
			frappe.call({
				method: SETTINGS_TEST_METHOD,
				freeze: true,
				freeze_message: __("Testing connection..."),
				callback: (response) => {
					const result = response.message || {};
					const indicator = result.ok ? "green" : "red";
					frappe.msgprint({
						title: __("Test Connection"),
						message: result.message,
						indicator: indicator,
					});
				},
			});
		});
	},
});
