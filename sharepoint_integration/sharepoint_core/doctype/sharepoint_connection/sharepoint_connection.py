# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

import frappe
from frappe.model.document import Document

from sharepoint_integration.sharepoint_core.doctype.sharepoint_connection.sharepoint_connection_utils import (
	validate_connection,
)


class SharePointConnection(Document):
	"""One Microsoft tenant / Entra app registration; rules pick which one they use."""

	def validate(self):
		"""
		Validate base URL, tenant/client ids and secret.

		Returns:
		        None
		"""
		validate_connection(self)

	def on_update(self):
		"""
		Drop this connection's cached token so new credentials apply immediately.

		Returns:
		        None
		"""
		frappe.cache.delete_keys("sharepoint_integration:token:")
