# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

from frappe.model.document import Document

from sharepoint_integration.sharepoint_core.doctype.sharepoint_folder.sharepoint_folder_utils import (
	validate_folder,
)


class SharePointFolder(Document):
	"""One SharePoint folder per reference document: its status, ids, link and errors."""

	def validate(self):
		"""
		Allow only system-created records, one per reference document.

		Returns:
		        None
		"""
		validate_folder(self)
