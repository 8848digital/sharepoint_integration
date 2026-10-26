// Copyright (c) 2026 8848 Digital LLP. All rights reserved.
// Proprietary and confidential. Unauthorized copying, distribution, or use
// of this file, via any medium, is strictly prohibited without prior
// written permission from 8848 Digital LLP.

frappe.provide("sharepoint_integration");

/**
 * frappe.call that shows the server's own error message. The app's API replies
 * in the standard envelope ({status, message, errors}), which Frappe's default
 * error handling does not display, so failures used to say only "Request failed".
 *
 * @param {object} opts - frappe.call options (method, args, callback, freeze ...).
 * @returns {Promise} The request.
 */
sharepoint_integration.call = function (opts) {
	return frappe.call({ ...opts, error: () => {} }).fail((xhr) => {
		const body = (xhr && xhr.responseJSON) || {};
		frappe.msgprint({
			title: __("SharePoint"),
			indicator: "red",
			message: body.message || __("Request failed"),
		});
	});
};

/**
 * Whether SharePoint is enabled and this DocType has an enabled Folder Rule
 * (sent once in the desk boot; changes apply after a page reload).
 *
 * @param {string} doctype - DocType of the form.
 * @returns {boolean} True when folders can be created for this DocType.
 */
function sharePointActive(doctype) {
	const boot = frappe.boot.sharepoint || {};
	return Boolean(boot.enabled && (boot.doctypes || []).includes(doctype));
}

/**
 * Adds a "SharePoint" button group to forms of DocTypes with an enabled
 * Folder Rule, only while SharePoint Settings is enabled. "Create Folder"
 * is offered only to users who can edit the document.
 *
 * @param {object} frm - The refreshed form.
 * @returns {void}
 */
function addSharePointButtons(frm) {
	if (!frm.fields_dict.sharepoint_link || frm.is_new() || !sharePointActive(frm.doctype)) return;

	if (frm.doc.sharepoint_link) {
		frm.add_custom_button(
			__("Open in SharePoint"),
			() => window.open(frm.doc.sharepoint_link, "_blank"),
			__("SharePoint")
		);
		return;
	}
	if (!frm.perm[0] || !frm.perm[0].write) return;

	frm.add_custom_button(
		__("Create Folder"),
		() =>
			sharepoint_integration.call({
				method: "sharepoint_integration.sharepoint_core.api.v1.folder.create_folder",
				args: { doctype: frm.doctype, docname: frm.docname },
				freeze: true,
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
