# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

from frappe.model.document import Document

from sharepoint_integration.sharepoint_core.doctype.sharepoint_folder_rule.sharepoint_folder_rule_utils import (
	validate_rule,
)
from sharepoint_integration.sharepoint_core.services.custom_fields import ensure_reference_fields
from sharepoint_integration.sharepoint_core.services.rule_matcher import clear_rule_cache


class SharePointFolderRule(Document):
	"""Where and how to create SharePoint folders for one DocType (optionally one variant of it)."""

	def validate(self):
		"""
		Validate reference DocType, trigger, location, subfolders, owners and template.

		Returns:
		        None
		"""
		validate_rule(self)

	def on_update(self):
		"""
		Add SharePoint fields to the reference DocType and refresh the rule cache.

		Returns:
		        None
		"""
		ensure_reference_fields(self.reference_doctype)
		clear_rule_cache(self.reference_doctype)

		before = self.get_doc_before_save()
		if before and before.reference_doctype != self.reference_doctype:
			clear_rule_cache(before.reference_doctype)

	def on_trash(self):
		"""
		Refresh the rule cache after deletion.

		Returns:
		        None
		"""
		clear_rule_cache(self.reference_doctype)
