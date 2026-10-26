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
	 * Add the "Resolve Site" and "Create Folders for Existing Records" buttons once saved.
	 *
	 * @param {object} frm - Current form.
	 * @returns {void}
	 */
	refresh(frm) {
		if (frm.is_new()) return;
		frm.add_custom_button(__("Resolve Site"), () =>
			sharepoint_integration.call({
				method: "sharepoint_integration.sharepoint_core.api.v1.settings.resolve_rule_site",
				args: { rule: frm.doc.name },
				freeze: true,
				callback: () => frm.reload_doc(),
			})
		);
		if (frm.doc.enabled) {
			frm.add_custom_button(__("Create Folders for Existing Records"), () =>
				confirm_existing_records(frm)
			);
		}
	},
});

/**
 * Show how many records the rule covers, then start the background run.
 *
 * @param {object} frm - Current form.
 * @returns {void}
 */
function confirm_existing_records(frm) {
	const api = "sharepoint_integration.sharepoint_core.api.v1.folder";
	sharepoint_integration.call({
		method: `${api}.count_existing_records`,
		args: { rule: frm.doc.name },
		type: "GET",
		callback: (r) => {
			const counts = r.data || {};
			const uploads = frm.doc.upload_attachments
				? __("Attachments of these records are uploaded too.")
				: "";
			const message = __(
				"{0} has {1} record(s) for this rule; {2} already have a SharePoint folder. Folders are created in the background for records that match the condition and have none. {3} Continue?",
				[__(frm.doc.reference_doctype), counts.records, counts.with_folder, uploads]
			);
			frappe.confirm(message, () =>
				sharepoint_integration.call({
					method: `${api}.create_for_existing_records`,
					args: { rule: frm.doc.name },
					freeze: true,
					callback: () =>
						frappe.show_alert({
							message: __("Started; you will get a message when it finishes."),
							indicator: "green",
						}),
				})
			);
		},
	});
}
