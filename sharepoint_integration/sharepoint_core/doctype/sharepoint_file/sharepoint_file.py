# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

from frappe.model.document import Document

from sharepoint_integration.sharepoint_core.doctype.sharepoint_file.sharepoint_file_utils import (
	validate_file,
)


class SharePointFile(Document):
	"""One attachment uploaded (or waiting to be uploaded) to a document's SharePoint folder."""

	def validate(self):
		"""
		Allow only system-created records, one per attachment.

		Returns:
		        None
		"""
		validate_file(self)
