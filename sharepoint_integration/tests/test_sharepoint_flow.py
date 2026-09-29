# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""
End-to-end tests of the SharePoint flow against the Graph sandbox, using the
core "ToDo" DocType to prove nothing is tied to ERPNext/specific DocTypes.
"""

import frappe
from frappe.tests.utils import FrappeTestCase
from frappe.utils import add_days, today

from sharepoint_integration.sharepoint_core import tasks
from sharepoint_integration.sharepoint_core.graph.sandbox import STATE_KEY
from sharepoint_integration.sharepoint_core.services.folder_request import sanitize_folder_name
from sharepoint_integration.tests.sharepoint_fixtures import (
	CONNECTIONS,
	USERS,
	make_todo,
	member_statuses,
	setup_sandbox,
)


class TestSharePointFlow(FrappeTestCase):
	"""Rule match -> folder provisioning -> access grant -> expiry, in sandbox mode."""

	@classmethod
	def setUpClass(cls):
		"""
		Create two connections (two tenants), enable sandbox settings, create
		users and one ToDo rule per connection.

		Returns:
		        None
		"""
		super().setUpClass()
		setup_sandbox()

	def test_rules_route_to_their_own_tenant(self):
		"""
		Two rules on two connections: each folder is created in its rule's tenant
		and remembers that connection.

		Returns:
		        None
		"""
		folders = {}
		for priority in ("High", "Medium"):
			todo = make_todo(priority)
			name = frappe.db.get_value("SharePoint Folder", {"reference_name": todo.name})
			tasks.provision_folder(name)
			folders[priority] = frappe.db.get_value(
				"SharePoint Folder", name, ["connection", "web_url"], as_dict=True
			)

		self.assertEqual(folders["High"].connection, "Test Tenant A")
		self.assertTrue(folders["High"].web_url.startswith(CONNECTIONS["Test Tenant A"]))
		self.assertEqual(folders["Medium"].connection, "Test Tenant B")
		self.assertTrue(folders["Medium"].web_url.startswith(CONNECTIONS["Test Tenant B"]))

	def test_folder_tree_and_permissions(self):
		"""
		A matching ToDo gets folder + subfolders; the group reaches only "Shared".

		Returns:
		        None
		"""
		todo = make_todo("High")
		folder = frappe.get_doc("SharePoint Folder", {"reference_name": todo.name})
		tasks.provision_folder(folder.name)
		folder.reload()

		self.assertEqual(folder.status, "Created", folder.error)
		self.assertEqual(folder.subfolders.split("\n"), ["Shared", "Private"])
		self.assertEqual(frappe.db.get_value("ToDo", todo.name, "sharepoint_folder"), folder.name)

		items = frappe.cache.get_value(STATE_KEY)["items"]
		shared = self.__principals(items[f"{folder.folder_path}/Shared"])
		private = self.__principals(items[f"{folder.folder_path}/Private"])
		self.assertIn(folder.group_id, shared)
		self.assertNotIn(folder.group_id, private)
		self.assertIn(USERS[0], private)

	def test_non_matching_condition_creates_nothing(self):
		"""
		The rule's condition filters documents out.

		Returns:
		        None
		"""
		todo = make_todo("Low")
		self.assertFalse(frappe.db.exists("SharePoint Folder", {"reference_name": todo.name}))

	def test_access_grant_and_expiry(self):
		"""
		Members are granted, and only rows with a past Expiry Date expire
		(regression: empty dates must not count as expired).

		Returns:
		        None
		"""
		todo = make_todo("High")
		tasks.provision_folder(frappe.db.get_value("SharePoint Folder", {"reference_name": todo.name}))
		access = frappe.get_doc(
			{
				"doctype": "SharePoint Access",
				"reference_doctype": "ToDo",
				"reference_name": todo.name,
				"members": [{"user": USERS[0]}, {"user": USERS[1], "from_date": add_days(today(), -5)}],
			}
		).insert()
		tasks.sync_access(access.name)
		self.assertEqual(member_statuses(access), ["Granted", "Granted"])

		frappe.db.set_value(
			"SharePoint Access Member", access.members[1].name, "expiry_date", add_days(today(), -1)
		)
		tasks.daily()
		self.assertEqual(member_statuses(access), ["Granted", "Expired"])

	def test_sanitize_folder_name(self):
		"""
		Invalid SharePoint characters are dropped; Unicode is kept.

		Returns:
		        None
		"""
		self.assertEqual(sanitize_folder_name('  ACME: "Q3" RFP/2026. '), "ACME Q3 RFP2026")
		self.assertEqual(sanitize_folder_name("مشروع #1"), "مشروع 1")

	def __principals(self, item: dict) -> set[str]:
		"""
		Emails/ids that hold a permission on a sandbox item.

		Parameters:
		        item (dict, required): Sandbox driveItem.

		Returns:
		        set[str]: Principal keys.
		"""
		return {
			value
			for perm in item["permissions"]
			for value in perm["grantedToV2"][next(iter(perm["grantedToV2"]))].values()
		}
