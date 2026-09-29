// Copyright (c) 2026 8848 Digital LLP. All rights reserved.
// Proprietary and confidential. Unauthorized copying, distribution, or use
// of this file, via any medium, is strictly prohibited without prior
// written permission from 8848 Digital LLP.

/**
 * Adds a "SharePoint" button group to any form whose DocType has a Folder Rule
 * (detected by the sharepoint_link custom field the rule adds).
 *
 * @param {object} frm - The refreshed form.
 * @returns {void}
 */
function addSharePointButtons(frm) {
	if (!frm.fields_dict.sharepoint_link || frm.is_new()) return;

	if (frm.doc.sharepoint_link) {
		frm.add_custom_button(
			__("Open in SharePoint"),
			() => window.open(frm.doc.sharepoint_link, "_blank"),
			__("SharePoint")
		);
		return;
	}
	frm.add_custom_button(
		__("Create Folder"),
		() =>
			frappe.call({
				method: "sharepoint_integration.sharepoint_core.api.v1.folder.create_folder",
				args: { doctype: frm.doctype, docname: frm.docname },
				freeze: true,
				error: (r) =>
					frappe.msgprint({
						title: __("SharePoint"),
						indicator: "red",
						message: (r && r.message) || __("Request failed"),
					}),
				callback: () =>
					frappe.show_alert({
						message: __("SharePoint folder queued; refresh in a moment"),
						indicator: "blue",
					}),
			}),
		__("SharePoint")
	);
}

$(document).on("form-refresh", (event, frm) => addSharePointButtons(frm));
