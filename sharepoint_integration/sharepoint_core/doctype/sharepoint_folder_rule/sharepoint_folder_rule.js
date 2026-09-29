// Copyright (c) 2026 8848 Digital LLP. All rights reserved.
// Proprietary and confidential. Unauthorized copying, distribution, or use
// of this file, via any medium, is strictly prohibited without prior
// written permission from 8848 Digital LLP.

frappe.ui.form.on("SharePoint Folder Rule", {
	/**
	 * Limit Reference DocType to normal, stored DocTypes.
	 *
	 * @param {object} frm - Current form.
	 * @returns {void}
	 */
	setup(frm) {
		frm.set_query("reference_doctype", () => ({
			filters: { istable: 0, issingle: 0, is_virtual: 0 },
		}));
		frm.set_query("connection", () => ({ filters: { enabled: 1 } }));
	},

	/**
	 * Add the "Resolve Site" button once saved.
	 *
	 * @param {object} frm - Current form.
	 * @returns {void}
	 */
	refresh(frm) {
		if (frm.is_new()) return;
		frm.add_custom_button(__("Resolve Site"), () =>
			frappe.call({
				method: "sharepoint_integration.sharepoint_core.api.v1.settings.resolve_rule_site",
				args: { rule: frm.doc.name },
				freeze: true,
				error: (r) =>
					frappe.msgprint({
						title: __("SharePoint"),
						indicator: "red",
						message: (r && r.message) || __("Request failed"),
					}),
				callback: () => frm.reload_doc(),
			})
		);
	},
});
