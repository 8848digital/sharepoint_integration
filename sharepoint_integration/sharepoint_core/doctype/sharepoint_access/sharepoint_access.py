# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

from frappe.model.document import Document

from sharepoint_integration.sharepoint_core.doctype.sharepoint_access.sharepoint_access_utils import (
	queue_access_sync,
	validate_access,
	validate_access_deletion,
)


class SharePointAccess(Document):
	"""Who can open a document's SharePoint folder, and for which dates."""

	def validate(self):
		"""
		Validate reference, members and dates; mark changed members Pending.

		Returns:
		        None
		"""
		validate_access(self)

	def on_update(self):
		"""
		Queue the group membership sync in the background.

		Returns:
		        None
		"""
		queue_access_sync(self)

	def on_trash(self):
		"""
		Block deletion while any member still has access.

		Returns:
		        None
		"""
		validate_access_deletion(self)
