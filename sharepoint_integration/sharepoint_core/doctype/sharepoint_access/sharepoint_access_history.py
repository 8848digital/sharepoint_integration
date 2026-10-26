# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Protects SharePoint Access history: rows that had access are revoked, never rewritten."""

import frappe
from frappe import _
from frappe.utils import getdate

HAD_ACCESS_STATUSES = ("Granted", "Expired", "Revoked")
CLOSED_STATUSES = ("Expired", "Revoked")
SYSTEM_FIELDS = ("access_status", "access_given", "granted_on", "removed_on", "last_error")


def track_member_changes(doc) -> None:
	"""
	New rows start Pending. Existing rows keep their system fields from the
	database (never from the browser), may not be deleted or switched to another
	user once they had access, and go back to Pending when an expired row's
	Expiry Date is extended.

	Parameters:
	        doc (Document, required): SharePoint Access.

	Returns:
	        None
	"""
	before = doc.get_doc_before_save()
	previous = {row.name: row for row in before.members} if before else {}
	current = {row.name for row in doc.members}

	deleted = [
		row.user
		for row in previous.values()
		if row.name not in current and row.access_status in HAD_ACCESS_STATUSES
	]
	if deleted:
		frappe.throw(
			_("Use Revoke Access instead of deleting rows that had access: {0}").format(", ".join(deleted))
		)

	for row in doc.members:
		old = previous.get(row.name)
		if not old:
			row.update(
				{
					"access_status": "Pending",
					"access_given": 0,
					"granted_on": None,
					"removed_on": None,
					"last_error": None,
				}
			)
			continue
		if old.user != row.user and old.access_status in HAD_ACCESS_STATUSES:
			frappe.throw(
				_("Row {0}: revoke {1} and add a new row instead of changing the user").format(
					row.idx, old.user
				)
			)

		row.update({field: old.get(field) for field in SYSTEM_FIELDS})
		if old.user != row.user or __expiry_extended(row):
			row.update({"access_status": "Pending", "removed_on": None, "last_error": None})


def __expiry_extended(row) -> bool:
	"""
	True when an Expired row's Expiry Date was moved to today or later.

	Parameters:
	        row (Document, required): SharePoint Access Member row.

	Returns:
	        bool: Whether access should be granted again.
	"""
	return bool(
		row.access_status == "Expired" and row.expiry_date and getdate(row.expiry_date) >= getdate()
	)
