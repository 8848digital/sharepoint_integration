# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Scheduled clean-up: expire access after Expiry Date and retry pending grants."""

from collections import defaultdict

import frappe
from frappe.utils import add_days, getdate

from sharepoint_integration.sharepoint_core.services.access_sync import (
	AccessSync,
	enqueue_access_sync,
)

MEMBER_DOCTYPE = "SharePoint Access Member"
RETRY_BATCH_SIZE = 50


def expire_access() -> None:
	"""
	Daily: fill Expiry Date from To Date + grace days (if configured), then remove
	members whose Expiry Date has passed. Rows are kept (status Expired) as an
	audit trail of who had access and when.

	Parameters:
	        None

	Returns:
	        None
	"""
	if not frappe.db.get_single_value("SharePoint Settings", "enabled"):
		return

	__apply_auto_expiry()
	expired = frappe.get_all(
		MEMBER_DOCTYPE,
		# "is set" is required: Frappe compares an empty date as 0001-01-01, so
		# a bare "<" filter would expire every row without an Expiry Date.
		filters=[
			["access_status", "=", "Granted"],
			["parenttype", "=", "SharePoint Access"],
			["expiry_date", "is", "set"],
			["expiry_date", "<", getdate()],
		],
		fields=["name", "parent", "user", "access_status"],
	)
	rows_by_access = defaultdict(list)
	for row in expired:
		rows_by_access[row.parent].append(row)

	for access, rows in rows_by_access.items():
		AccessSync(access).revoke(rows, status="Expired")


def retry_pending_access() -> None:
	"""
	Hourly: re-queue syncs for accesses that still have Pending/Failed members
	(e.g. saved before the folder existed, or a transient Graph error).

	Parameters:
	        None

	Returns:
	        None
	"""
	if not frappe.db.get_single_value("SharePoint Settings", "enabled"):
		return

	accesses = frappe.get_all(
		MEMBER_DOCTYPE,
		filters={"access_status": ["in", ["Pending", "Failed"]], "parenttype": "SharePoint Access"},
		pluck="parent",
		distinct=True,
		limit=RETRY_BATCH_SIZE,
	)
	for access in accesses:
		enqueue_access_sync(access)


def __apply_auto_expiry() -> None:
	"""
	Set Expiry Date = To Date + N days for granted rows whose To Date passed.

	Parameters:
	        None

	Returns:
	        None
	"""
	grace_days = frappe.db.get_single_value("SharePoint Settings", "auto_expire_after_days") or 0
	if grace_days <= 0:
		return

	rows = frappe.get_all(
		MEMBER_DOCTYPE,
		filters=[
			["access_status", "=", "Granted"],
			["expiry_date", "is", "not set"],
			["to_date", "is", "set"],
			["to_date", "<", getdate()],
		],
		fields=["name", "to_date"],
	)
	for row in rows:
		frappe.db.set_value(
			MEMBER_DOCTYPE,
			row.name,
			"expiry_date",
			add_days(row.to_date, grace_days),
			update_modified=False,
		)
