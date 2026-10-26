# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

import frappe
from frappe.model.document import Document

from sharepoint_integration.sharepoint_core.doctype.sharepoint_settings.sharepoint_settings_utils import (
	validate_settings,
)


class SharePointSettings(Document):
	"""Global SharePoint / Microsoft Graph connection settings."""

	def validate(self):
		"""
		Validate URL, ID formats and required credentials.

		Returns:
		        None
		"""
		validate_settings(self)

	def on_update(self):
		"""
		Drop cached tokens so changed credentials take effect immediately.

		Returns:
		        None
		"""
		frappe.cache.delete_keys("sharepoint_integration:token:")
