# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Read-side helpers for the folder/access API endpoints (permission-checked)."""

import frappe

from sharepoint_integration.sharepoint_core.services.access_sync import enqueue_access_sync


def get_folder_details(doctype: str, docname: str) -> dict:
	"""
	Folder status and link of a document the user can read.

	Parameters:
	        doctype (str, required): Reference DocType.
	        docname (str, required): Reference document name.

	Returns:
	        dict: {"folder", "status", "web_url", "folder_name", "error"} (empty values if none).
	"""
	frappe.get_doc(doctype, docname).check_permission("read")
	details = frappe.db.get_value(
		"SharePoint Folder",
		{"reference_doctype": doctype, "reference_name": docname},
		["name as folder", "status", "web_url", "folder_name", "error"],
		as_dict=True,
	)
	return details or {
		"folder": None,
		"status": None,
		"web_url": None,
		"folder_name": None,
		"error": None,
	}


def queue_access_sync_now(access: str) -> None:
	"""
	Queue a membership sync for a SharePoint Access the user can write.

	Parameters:
	        access (str, required): SharePoint Access name.

	Returns:
	        None
	"""
	frappe.get_doc("SharePoint Access", access).check_permission("write")
	enqueue_access_sync(access)
