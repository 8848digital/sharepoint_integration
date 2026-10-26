# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Data the SharePoint form buttons need, sent once with the desk boot."""

import frappe

from sharepoint_integration.sharepoint_core.services.rule_matcher import get_rule_doctypes


def boot_session(bootinfo) -> None:
	"""
	Tell the desk whether SharePoint is enabled and which DocTypes have an
	enabled Folder Rule, so forms show the SharePoint button only then.

	Parameters:
	        bootinfo (frappe._dict, required): Boot data sent to the browser.

	Returns:
	        None
	"""
	enabled = bool(frappe.db.get_single_value("SharePoint Settings", "enabled"))
	bootinfo.sharepoint = {
		"enabled": enabled,
		"doctypes": sorted(get_rule_doctypes()) if enabled else [],
	}
