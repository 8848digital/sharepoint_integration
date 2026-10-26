# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Validation for SharePoint File."""

import frappe
from frappe import _


def validate_file(doc) -> None:
	"""
	SharePoint File records are written only by the integration, one per attachment.

	Parameters:
	        doc (Document, required): SharePoint File.

	Returns:
	        None
	"""
	if not doc.flags.ignore_permissions:
		frappe.throw(_("SharePoint Files are created by the integration when a file is attached"))

	duplicate = frappe.db.exists("SharePoint File", {"file": doc.file, "name": ["!=", doc.name]})
	if duplicate:
		frappe.throw(_("File {0} already has SharePoint File {1}").format(doc.file, duplicate))
