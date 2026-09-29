# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Revoking access keeps history; rows that had access cannot be deleted or re-pointed."""

import frappe
from frappe.tests.utils import FrappeTestCase

from sharepoint_integration.sharepoint_core import tasks
from sharepoint_integration.sharepoint_core.graph.sandbox import STATE_KEY
from sharepoint_integration.sharepoint_core.services.access_revoke import request_revoke
from sharepoint_integration.tests.sharepoint_fixtures import (
	USERS,
	make_todo,
	member_statuses,
	setup_sandbox,
)


class TestSharePointRevoke(FrappeTestCase):
	"""Revoke Access behaviour and the access-history guards."""

	@classmethod
	def setUpClass(cls):
		"""
		Sandbox connections, users and ToDo rules.

		Returns:
		        None
		"""
		super().setUpClass()
		setup_sandbox()

	def setUp(self):
		"""
		A provisioned ToDo folder with both users granted.

		Returns:
		        None
		"""
		todo = make_todo("High")
		folder = frappe.db.get_value("SharePoint Folder", {"reference_name": todo.name})
		tasks.provision_folder(folder)
		self.group_id = frappe.db.get_value("SharePoint Folder", folder, "group_id")
		self.access = frappe.get_doc(
			{
				"doctype": "SharePoint Access",
				"reference_doctype": "ToDo",
				"reference_name": todo.name,
				"members": [{"user": USERS[0]}, {"user": USERS[1]}],
			}
		).insert()
		tasks.sync_access(self.access.name)
		self.access.reload()

	def test_revoke_removes_from_group_and_keeps_row(self):
		"""
		Revoke takes the user out of the group now and keeps the row as history.

		Returns:
		        None
		"""
		row = self.access.members[1]
		result = request_revoke(self.access.name, [row.name], "Left the project")

		self.assertEqual(result, {"queued": False, "errors": []})
		self.assertEqual(member_statuses(self.access), ["Granted", "Revoked"])
		self.assertEqual(self.access.members[1].remarks, "Left the project")
		self.assertTrue(self.access.members[1].removed_on)
		self.assertEqual([row.access_given for row in self.access.members], [1, 0])
		self.assertNotIn(f"user-{USERS[1]}", self.__group_members())
		self.assertIn(f"user-{USERS[0]}", self.__group_members())

	def test_granted_row_cannot_be_deleted(self):
		"""
		Deleting a granted row would erase history, so it must go through Revoke.

		Returns:
		        None
		"""
		self.access.members = self.access.members[:1]
		self.assertRaises(frappe.ValidationError, self.access.save)

	def test_granted_row_user_cannot_be_switched(self):
		"""
		Re-pointing a granted row to another user is refused.

		Returns:
		        None
		"""
		self.access.members[0].user = "Administrator"
		self.assertRaises(frappe.ValidationError, self.access.save)

	def test_user_can_be_re_added_after_revoke(self):
		"""
		After a revoke the old row stays and a new row grants access again.

		Returns:
		        None
		"""
		request_revoke(self.access.name, [self.access.members[1].name], "Paused")
		self.access.reload()
		self.access.append("members", {"user": USERS[1]})
		self.access.save()
		tasks.sync_access(self.access.name)

		self.assertEqual(member_statuses(self.access), ["Granted", "Revoked", "Granted"])
		self.assertIn(f"user-{USERS[1]}", self.__group_members())

	def test_status_cannot_be_forged_from_client(self):
		"""
		System fields come from the database, not from the submitted form.

		Returns:
		        None
		"""
		self.access.members[0].access_status = "Pending"
		self.access.members[0].granted_on = None
		self.access.save()
		self.assertEqual(member_statuses(self.access)[0], "Granted")
		self.assertTrue(self.access.members[0].granted_on)

	def test_access_with_granted_members_cannot_be_deleted(self):
		"""
		Deleting the whole SharePoint Access is blocked while anyone has access.

		Returns:
		        None
		"""
		self.assertRaises(frappe.ValidationError, self.access.delete)

	def __group_members(self) -> list[str]:
		"""
		Members of the ToDo's security group in the sandbox.

		Returns:
		        list[str]: Sandbox user ids.
		"""
		groups = frappe.cache.get_value(STATE_KEY)["groups"].values()
		return next(group for group in groups if group["id"] == self.group_id)["members"]
