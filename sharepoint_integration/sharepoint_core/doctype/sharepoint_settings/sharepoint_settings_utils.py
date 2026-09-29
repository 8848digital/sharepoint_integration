# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Validation for SharePoint Settings."""

import frappe
from frappe import _


def validate_settings(doc) -> None:
	"""
	Validate global options: default connection, expiry days, failure recipients.

	Parameters:
	        doc (Document, required): SharePoint Settings.

	Returns:
	        None
	"""
	if doc.default_connection and not frappe.db.get_value(
		"SharePoint Connection", doc.default_connection, "enabled"
	):
		frappe.throw(_("Default Connection {0} is disabled").format(doc.default_connection))
	if doc.enabled and not frappe.db.exists("SharePoint Connection", {"enabled": 1}):
		frappe.throw(
			_("Create at least one enabled SharePoint Connection before enabling the integration")
		)

	if (doc.auto_expire_after_days or 0) < 0:
		frappe.throw(_("Auto Expire After (Days) cannot be negative"))
	if doc.notify_on_failure and not doc.failure_recipients:
		frappe.throw(_("Add at least one recipient for failure emails"))
	if doc.enabled and doc.sandbox_mode:
		frappe.msgprint(
			_("Sandbox Mode is on: no calls are made to Microsoft. Turn it off before going live."),
			indicator="orange",
			alert=True,
		)
