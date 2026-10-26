// Copyright (c) 2026 8848 Digital LLP. All rights reserved.
// Proprietary and confidential. Unauthorized copying, distribution, or use
// of this file, via any medium, is strictly prohibited without prior
// written permission from 8848 Digital LLP.

frappe.ui.form.on("SharePoint Access", {
	/**
	 * Only DocTypes with a Folder Rule can be referenced.
	 *
	 * @param {object} frm - Current form.
	 * @returns {void}
	 */
	setup(frm) {
		frm.set_query("reference_doctype", () => ({
			filters: { name: ["in", frm.__rule_doctypes || []] },
		}));
		frappe.db
			.get_list("SharePoint Folder Rule", {
				filters: { enabled: 1 },
				fields: ["reference_doctype"],
			})
			.then((rows) => (frm.__rule_doctypes = rows.map((row) => row.reference_doctype)));
	},

	/**
	 * Add "Sync Now" and "Open in SharePoint" buttons.
	 *
	 * @param {object} frm - Current form.
	 * @returns {void}
	 */
	refresh(frm) {
		if (frm.is_new()) return;
		if (frm.doc.web_url) {
			frm.add_custom_button(__("Open in SharePoint"), () =>
				window.open(frm.doc.web_url, "_blank")
			);
		}
		frm.add_custom_button(__("Revoke Access"), () => revokeSelected(frm));
		frm.add_custom_button(__("Sync Now"), () =>
			sharepoint_integration.call({
				method: "sharepoint_integration.sharepoint_core.api.v1.access.sync_now",
				args: { access: frm.doc.name },
				callback: () =>
					frappe.show_alert({ message: __("Access sync queued"), indicator: "blue" }),
			})
		);
	},
});

/**
 * Ask for a reason and revoke the selected member rows. Rows stay in the
 * table as "Revoked" so the access history is kept.
 *
 * @param {object} frm - SharePoint Access form.
 * @returns {void}
 */
function revokeSelected(frm) {
	const rows = frm.fields_dict.members.grid.get_selected_children();
	const open = rows.filter((row) =>
		["Pending", "Granted", "Failed"].includes(row.access_status)
	);
	if (!open.length) {
		frappe.msgprint(
			__("Select one or more members that still have (or are waiting for) access.")
		);
		return;
	}
	if (frm.is_dirty()) {
		frappe.msgprint(__("Save the form before revoking access."));
		return;
	}
	frappe.prompt(
		{ fieldname: "remarks", fieldtype: "Small Text", label: __("Reason"), reqd: 1 },
		(values) =>
			sharepoint_integration.call({
				method: "sharepoint_integration.sharepoint_core.api.v1.access.revoke",
				args: {
					access: frm.doc.name,
					members: open.map((row) => row.name),
					remarks: values.remarks,
				},
				freeze: true,
				freeze_message: __("Removing access..."),
				callback: (r) => {
					const data = r.data || {};
					const errors = data.errors || [];
					if (errors.length) {
						frappe.msgprint({
							title: __("Some members could not be removed"),
							indicator: "orange",
							message: errors.join("<br>"),
						});
					} else {
						frappe.show_alert({
							message: data.queued ? __("Revoke queued") : __("Access revoked"),
							indicator: "green",
						});
					}
					frm.reload_doc();
				},
			}),
		__("Revoke Access for {0} member(s)", [open.length]),
		__("Revoke")
	);
}
