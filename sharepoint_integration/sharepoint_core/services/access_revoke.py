# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Revoke selected members' access now, keeping their rows as an audit trail."""

import frappe
from frappe import _

from sharepoint_integration.sharepoint_core.services.access_sync import AccessSync

MEMBER_DOCTYPE = "SharePoint Access Member"
OPEN_STATUSES = ("Pending", "Granted", "Failed")
INLINE_LIMIT = 20  # more members than this are revoked in a background job


def request_revoke(access: str, members: list[str], remarks: str | None = None) -> dict:
	"""
	Validate and revoke the selected members. Small batches run immediately so
	the user sees the result; large ones are queued to avoid request timeouts.

	Parameters:
	        access (str, required): SharePoint Access name.
	        members (list[str], required): SharePoint Access Member row names.
	        remarks (str, optional): Reason, stored on each row.

	Returns:
	        dict: {"queued": bool, "errors": list[str]}
	"""
	doc = frappe.get_doc("SharePoint Access", access)
	doc.check_permission("write")

	rows = {row.name: row for row in doc.members}
	unknown = [name for name in members if name not in rows]
	if not members or unknown:
		frappe.throw(_("Select member rows of this SharePoint Access to revoke"))
	closed = [rows[name].user for name in members if rows[name].access_status not in OPEN_STATUSES]
	if closed:
		frappe.throw(_("Already expired or revoked: {0}").format(", ".join(closed)))

	if remarks:
		for name in members:
			frappe.db.set_value(MEMBER_DOCTYPE, name, "remarks", remarks, update_modified=False)

	if len(members) > INLINE_LIMIT:
		frappe.enqueue(
			"sharepoint_integration.sharepoint_core.tasks.revoke_access",
			queue="long",
			timeout=3600,
			enqueue_after_commit=True,
			access=access,
			members=members,
		)
		return {"queued": True, "errors": []}
	return {"queued": False, "errors": revoke_members(access, members)}


def revoke_members(access: str, members: list[str]) -> list[str]:
	"""
	Remove the members from the group and mark their rows Revoked.

	Parameters:
	        access (str, required): SharePoint Access name.
	        members (list[str], required): SharePoint Access Member row names.

	Returns:
	        list[str]: "<user>: <error>" for members that could not be removed.
	"""
	rows = frappe.get_all(
		MEMBER_DOCTYPE,
		filters={"parent": access, "name": ["in", members], "access_status": ["in", OPEN_STATUSES]},
		fields=["name", "user", "access_status"],
	)
	return AccessSync(access).revoke(rows, status="Revoked")
