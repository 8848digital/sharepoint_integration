# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Keep a document's security group membership in line with its SharePoint Access."""

import frappe
from frappe.utils import now_datetime

from sharepoint_integration.sharepoint_core.graph.client import GraphClient
from sharepoint_integration.sharepoint_core.graph.directory import DirectoryService
from sharepoint_integration.sharepoint_core.services.notifications import notify_failure

MEMBER_DOCTYPE = "SharePoint Access Member"
SYNCABLE_STATUSES = ("Pending", "Failed")


def enqueue_access_sync(access: str) -> None:
	"""
	Queue a membership sync after the current transaction commits.
	Adding members to a group is done in the background because Graph/SharePoint
	can take minutes to apply and large teams would time out a web request.

	Parameters:
	        access (str, required): SharePoint Access name.

	Returns:
	        None
	"""
	frappe.enqueue(
		"sharepoint_integration.sharepoint_core.tasks.sync_access",
		queue="long",
		timeout=3600,
		enqueue_after_commit=True,
		access=access,
		job_id=f"sharepoint-access-{access}",
		deduplicate=True,
	)


class AccessSync:
	"""Adds Pending/Failed members to the group and removes dropped users."""

	def __init__(self, access: str):
		"""
		Load the SharePoint Access and the folder it grants access to.

		Parameters:
		        access (str, required): SharePoint Access name.

		Returns:
		        None
		"""
		self.access = frappe.get_doc("SharePoint Access", access)
		self.folder = frappe.db.get_value(
			"SharePoint Folder",
			{
				"reference_doctype": self.access.reference_doctype,
				"reference_name": self.access.reference_name,
				"status": "Created",
			},
			["name", "connection", "group_id"],
			as_dict=True,
		)
		self.directory = None

	def run(self) -> None:
		"""
		Sync membership. Without a created folder + group, members stay Pending and
		are picked up when the folder job finishes (or by the hourly retry).

		Parameters:
		        None

		Returns:
		        None
		"""
		if not (self.folder and self.folder.group_id):
			return

		self.directory = DirectoryService(GraphClient.from_connection(self.folder.connection))

		rows = [row for row in self.access.members if row.access_status in SYNCABLE_STATUSES]
		failures = [error for error in (self.__grant(row) for row in rows) if error]
		if failures:
			self.access.error = "\n".join(failures)
			notify_failure(self.access, subject="SharePoint access could not be granted")

	def revoke(self, rows: list[dict], status: str = "Expired") -> list[str]:
		"""
		Remove members from the group and close their rows (kept for audit).
		Rows that never reached the group (Pending/Failed) are closed directly.

		Parameters:
		        rows (list[dict], required): Member rows with `name`, `user`, `access_status`.
		        status (str, optional): Final status, "Expired" (scheduler) or "Revoked" (user action).

		Returns:
		        list[str]: "<user>: <error>" for members that could not be removed.
		"""
		has_group = bool(self.folder and self.folder.group_id)
		if has_group:
			self.directory = DirectoryService(GraphClient.from_connection(self.folder.connection))

		errors = []
		for row in rows:
			if row["access_status"] == "Granted" and not has_group:
				errors.append(f"{row['user']}: folder/group not found")
				continue
			error = self.__remove(row["user"]) if row["access_status"] == "Granted" else None
			values = (
				{"last_error": error}
				if error
				else {"access_status": status, "access_given": 0, "removed_on": now_datetime()}
			)
			frappe.db.set_value(MEMBER_DOCTYPE, row["name"], values, update_modified=False)
			if error:
				errors.append(f"{row['user']}: {error}")
		return errors

	def __grant(self, row) -> str | None:
		"""
		Add one member to the group and record the result on the row.

		Parameters:
		        row (Document, required): SharePoint Access Member row.

		Returns:
		        str | None: "<user>: <error>" when the grant failed.
		"""
		error_message = None
		try:
			self.directory.add_member(self.folder.group_id, _user_email(row.user))
			values = {
				"access_status": "Granted",
				"access_given": 1,
				"granted_on": now_datetime(),
				"last_error": None,
			}
		except Exception as error:
			frappe.log_error("SharePoint | Grant access failed", frappe.get_traceback())
			error_message = f"{row.user}: {error}"
			values = {"access_status": "Failed", "last_error": str(error)[:1000]}
		frappe.db.set_value(MEMBER_DOCTYPE, row.name, values, update_modified=False)
		return error_message

	def __remove(self, user: str) -> str | None:
		"""
		Remove one user from the group, logging (not raising) on failure.

		Parameters:
		        user (str, required): User name.

		Returns:
		        str | None: Error message, or None when the user is no longer a member.
		"""
		try:
			self.directory.remove_member(self.folder.group_id, _user_email(user))
			return None
		except Exception as error:
			frappe.log_error("SharePoint | Remove access failed", frappe.get_traceback())
			return str(error)[:1000]


def _user_email(user: str) -> str:
	"""
	Email of a User (single underscore: called from inside AccessSync).

	Parameters:
	        user (str, required): User name.

	Returns:
	        str: The user's email (falls back to the user name).
	"""
	return frappe.db.get_value("User", user, "email") or user
