# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Validation for SharePoint Folder."""

import frappe
from frappe import _


def validate_folder(doc) -> None:
	"""
	SharePoint Folder records are written only by the integration (rule match,
	button, or job) and there is exactly one per reference document.

	Parameters:
	        doc (Document, required): SharePoint Folder.

	Returns:
	        None
	"""
	if not doc.flags.ignore_permissions:
		frappe.throw(
			_("SharePoint Folders are created by the integration; use a Folder Rule or the form button")
		)

	if not frappe.db.exists(doc.reference_doctype, doc.reference_name):
		frappe.throw(_("{0} {1} does not exist").format(doc.reference_doctype, doc.reference_name))

	duplicate = frappe.db.exists(
		"SharePoint Folder",
		{
			"reference_doctype": doc.reference_doctype,
			"reference_name": doc.reference_name,
			"name": ["!=", doc.name],
		},
	)
	if duplicate:
		frappe.throw(
			_("{0} {1} already has SharePoint Folder {2}").format(
				doc.reference_doctype, doc.reference_name, duplicate
			)
		)

	if not doc.folder_name:
		frappe.throw(_("Folder Name is required"))
