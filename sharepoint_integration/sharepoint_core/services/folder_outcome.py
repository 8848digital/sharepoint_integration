# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Follow-up steps after a folder is created."""

import frappe

from sharepoint_integration.sharepoint_core.services.access_sync import enqueue_access_sync


def link_reference(record) -> None:
	"""
	Write the folder link onto the reference document without re-triggering its
	hooks (the old implementation called doc.save(), which re-ran every hook).

	Parameters:
	        record (Document, required): Created SharePoint Folder.

	Returns:
	        None
	"""
	meta = frappe.get_meta(record.reference_doctype)
	values = {"sharepoint_folder": record.name, "sharepoint_link": record.web_url}
	values = {field: value for field, value in values.items() if meta.has_field(field)}
	if values:
		frappe.db.set_value(
			record.reference_doctype, record.reference_name, values, update_modified=False
		)


def sync_pending_access(record) -> None:
	"""
	Members added before the folder existed are still Pending; sync them now.

	Parameters:
	        record (Document, required): Created SharePoint Folder.

	Returns:
	        None
	"""
	access = frappe.db.get_value(
		"SharePoint Access",
		{"reference_doctype": record.reference_doctype, "reference_name": record.reference_name},
	)
	if access:
		enqueue_access_sync(access)
