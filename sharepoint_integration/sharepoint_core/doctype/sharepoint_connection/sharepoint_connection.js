// Copyright (c) 2026 8848 Digital LLP. All rights reserved.
// Proprietary and confidential. Unauthorized copying, distribution, or use
// of this file, via any medium, is strictly prohibited without prior
// written permission from 8848 Digital LLP.

frappe.ui.form.on("SharePoint Connection", {
	/**
	 * Add the "Test Connection" button.
	 *
	 * @param {object} frm - Current form.
	 * @returns {void}
	 */
	refresh(frm) {
		if (frm.is_new() || !frm.doc.enabled) return;
		frm.add_custom_button(__("Test Connection"), () =>
			frappe.call({
				method: "sharepoint_integration.sharepoint_core.api.v1.settings.check_connection",
				args: { connection: frm.doc.name },
				freeze: true,
				freeze_message: __("Connecting to SharePoint..."),
				error: (r) =>
					frappe.msgprint({
						title: __("SharePoint"),
						indicator: "red",
						message: (r && r.message) || __("Request failed"),
					}),
				callback: (r) => {
					const data = r.data || {};
					frappe.msgprint({
						title: __("Connected"),
						indicator: "green",
						message: __("Root site: {0}{1}", [
							data.site_id,
							data.sandbox ? " (sandbox)" : "",
						]),
					});
				},
			})
		);
	},
});
