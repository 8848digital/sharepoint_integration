// Copyright (c) 2026 8848 Digital LLP. All rights reserved.
// Proprietary and confidential. Unauthorized copying, distribution, or use
// of this file, via any medium, is strictly prohibited without prior
// written permission from 8848 Digital LLP.

frappe.ui.form.on("SharePoint File", {
	/**
	 * Status indicator, "Open in SharePoint" and "Retry Upload" buttons.
	 *
	 * @param {object} frm - Current form.
	 * @returns {void}
	 */
	refresh(frm) {
		const colors = { Queued: "orange", Uploaded: "green", Failed: "red" };
		frm.page.set_indicator(__(frm.doc.status), colors[frm.doc.status]);

		if (frm.doc.web_url) {
			frm.add_custom_button(__("Open in SharePoint"), () =>
				window.open(frm.doc.web_url, "_blank")
			);
		}
		if (frm.doc.status === "Failed") {
			frm.add_custom_button(__("Retry Upload"), () =>
				sharepoint_integration.call({
					method: "sharepoint_integration.sharepoint_core.api.v1.file.retry_upload",
					args: { sharepoint_file: frm.doc.name },
					callback: () => frm.reload_doc(),
				})
			);
		}
	},
});
