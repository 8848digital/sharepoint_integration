# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""
doc_events "*" targets: every document save on the site passes through here,
so each handler exits after one cached set lookup unless the DocType has a rule.
"""

from sharepoint_integration.sharepoint_core.services.folder_request import request_folder
from sharepoint_integration.sharepoint_core.services.rule_matcher import get_rule_doctypes


def after_insert(doc, method=None):
	"""
	Queue folder creation for "On Create" / "On Save" rules.

	Parameters:
	        doc (Document, required): The inserted document.
	        method (str, optional): Hook name passed by Frappe.

	Returns:
	        None
	"""
	if doc.doctype in get_rule_doctypes():
		request_folder(doc, "insert")


def on_update(doc, method=None):
	"""
	Queue folder creation for "On Save" / "On Workflow State" rules.
	Skipped during insert, which after_insert already handled.

	Parameters:
	        doc (Document, required): The saved document.
	        method (str, optional): Hook name passed by Frappe.

	Returns:
	        None
	"""
	if not doc.flags.in_insert and doc.doctype in get_rule_doctypes():
		request_folder(doc, "update")
