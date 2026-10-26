# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""doc_events for File (owned by Frappe)."""

from sharepoint_integration.sharepoint_core.services.file_request import queue_file_upload


def after_insert(doc, method=None):
	"""
	Queue the upload of a new attachment when its document's folder rule
	uploads attachments.

	Parameters:
	        doc (Document, required): The new File.
	        method (str, optional): Hook name passed by Frappe.

	Returns:
	        None
	"""
	queue_file_upload(doc)
