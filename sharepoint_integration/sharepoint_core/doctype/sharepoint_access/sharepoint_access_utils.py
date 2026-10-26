# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Validation and change tracking for SharePoint Access."""

import frappe
from frappe import _
from frappe.utils import getdate

from sharepoint_integration.sharepoint_core.doctype.sharepoint_access.sharepoint_access_history import (
	CLOSED_STATUSES,
	track_member_changes,
)
from sharepoint_integration.sharepoint_core.services.access_sync import enqueue_access_sync


def validate_access(doc) -> None:
	"""
	Validate the reference, members and dates, and protect the access history
	(rows that had access can only be revoked, never deleted or re-pointed).

	Parameters:
	        doc (Document, required): SharePoint Access.

	Returns:
	        None
	"""
	__validate_reference(doc)
	__validate_members(doc)
	__validate_dates(doc)
	track_member_changes(doc)
	doc.sharepoint_folder = frappe.db.get_value(
		"SharePoint Folder",
		{"reference_doctype": doc.reference_doctype, "reference_name": doc.reference_name},
	)


def queue_access_sync(doc) -> None:
	"""
	After save, sync the group if any member is Pending.

	Parameters:
	        doc (Document, required): SharePoint Access.

	Returns:
	        None
	"""
	if any(row.access_status == "Pending" for row in doc.members):
		enqueue_access_sync(doc.name)


def validate_access_deletion(doc) -> None:
	"""
	Block deleting a SharePoint Access while anyone still has access through it.

	Parameters:
	        doc (Document, required): SharePoint Access.

	Returns:
	        None
	"""
	granted = [row.user for row in doc.members if row.access_status == "Granted"]
	if granted:
		frappe.throw(_("Revoke access first for: {0}").format(", ".join(granted)))


def __validate_reference(doc) -> None:
	"""
	The reference document must exist, have a rule, and have only one SharePoint Access.

	Parameters:
	        doc (Document, required): SharePoint Access.

	Returns:
	        None
	"""
	if not frappe.db.exists(doc.reference_doctype, doc.reference_name):
		frappe.throw(_("{0} {1} does not exist").format(doc.reference_doctype, doc.reference_name))
	if not frappe.db.exists(
		"SharePoint Folder Rule", {"reference_doctype": doc.reference_doctype, "enabled": 1}
	):
		frappe.throw(_("No enabled SharePoint Folder Rule exists for {0}").format(doc.reference_doctype))

	duplicate = frappe.db.exists(
		"SharePoint Access",
		{
			"reference_doctype": doc.reference_doctype,
			"reference_name": doc.reference_name,
			"name": ["!=", doc.name],
		},
	)
	if duplicate:
		frappe.throw(
			_("{0} {1} already has SharePoint Access {2}").format(
				doc.reference_doctype, doc.reference_name, duplicate
			)
		)
	if not doc.is_new() and (
		doc.has_value_changed("reference_doctype") or doc.has_value_changed("reference_name")
	):
		frappe.throw(_("Reference cannot be changed; create a new SharePoint Access instead"))


def __validate_members(doc) -> None:
	"""
	Each user at most once among open rows (closed rows are history, so a user
	can be re-added after a revoke), enabled, and with an email Microsoft can resolve.

	Parameters:
	        doc (Document, required): SharePoint Access.

	Returns:
	        None
	"""
	users = [row.user for row in doc.members]
	open_users = [row.user for row in doc.members if row.access_status not in CLOSED_STATUSES]
	duplicates = {user for user in open_users if open_users.count(user) > 1}
	if duplicates:
		frappe.throw(_("Users listed more than once: {0}").format(", ".join(sorted(duplicates))))

	details = {
		user.name: user
		for user in frappe.get_all(
			"User", filters={"name": ["in", users or [""]]}, fields=["name", "enabled", "email"]
		)
	}
	for row in doc.members:
		user = details.get(row.user)
		if row.access_status not in CLOSED_STATUSES and not (user and user.enabled):
			frappe.throw(_("Row {0}: User {1} is disabled").format(row.idx, row.user))
		if not (user and user.email):
			frappe.throw(_("Row {0}: User {1} has no email").format(row.idx, row.user))


def __validate_dates(doc) -> None:
	"""
	From <= To, From <= Expiry, and all dates inside the reference document's
	start/end dates when the rule defines those fields.

	Parameters:
	        doc (Document, required): SharePoint Access.

	Returns:
	        None
	"""
	start, end = __reference_bounds(doc)
	for row in doc.members:
		if row.from_date and row.to_date and getdate(row.from_date) > getdate(row.to_date):
			frappe.throw(_("Row {0}: From Date cannot be after To Date").format(row.idx))
		if row.from_date and row.expiry_date and getdate(row.expiry_date) < getdate(row.from_date):
			frappe.throw(_("Row {0}: Expiry Date cannot be before From Date").format(row.idx))

		dates = [getdate(value) for value in (row.from_date, row.to_date) if value]
		if start and any(value < start for value in dates):
			frappe.throw(
				_("Row {0}: dates must be on or after {1}").format(row.idx, frappe.format(start, "Date"))
			)
		if end and any(value > end for value in dates):
			frappe.throw(
				_("Row {0}: dates must be on or before {1}").format(row.idx, frappe.format(end, "Date"))
			)


def __reference_bounds(doc) -> tuple:
	"""
	Start/end dates of the reference document, from the rule's date fields.

	Parameters:
	        doc (Document, required): SharePoint Access.

	Returns:
	        tuple: (start date | None, end date | None)
	"""
	rule = frappe.db.get_value(
		"SharePoint Folder Rule",
		{"reference_doctype": doc.reference_doctype, "enabled": 1},
		["start_date_field", "end_date_field"],
		as_dict=True,
		order_by="priority desc",
	)
	fields = (
		[field for field in (rule.start_date_field, rule.end_date_field) if field] if rule else []
	)
	values = (
		frappe.db.get_value(doc.reference_doctype, doc.reference_name, fields, as_dict=True)
		if fields
		else {}
	)
	start = values.get(rule.start_date_field) if rule and rule.start_date_field else None
	end = values.get(rule.end_date_field) if rule and rule.end_date_field else None
	return (getdate(start) if start else None, getdate(end) if end else None)
