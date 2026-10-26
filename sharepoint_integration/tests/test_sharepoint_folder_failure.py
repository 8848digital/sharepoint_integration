# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""
A new security group needs time to reach SharePoint: the group invite waits
for it, and a failed folder records a plain-language reason.
"""

from unittest.mock import MagicMock, patch

import frappe
from frappe.tests.utils import FrappeTestCase

from sharepoint_integration.sharepoint_core import tasks
from sharepoint_integration.sharepoint_core.graph.common import GraphError
from sharepoint_integration.sharepoint_core.graph.drive import GROUP_WAITS, DriveService
from sharepoint_integration.tests.sharepoint_fixtures import make_todo, setup_sandbox

UNRESOLVED = GraphError(400, "One or more users could not be resolved")
INVITE = "sharepoint_integration.sharepoint_core.graph.drive.DriveService.invite"
SLEEP = "sharepoint_integration.sharepoint_core.graph.drive.time.sleep"


class TestSharePointFolderFailure(FrappeTestCase):
	"""Group invite waiting and failure reasons on SharePoint Folder."""

	@classmethod
	def setUpClass(cls):
		"""
		Sandbox connections, settings, users and ToDo rules.

		Returns:
		        None
		"""
		super().setUpClass()
		setup_sandbox()

	def test_group_invite_waits_until_sharepoint_knows_the_group(self):
		"""
		"Could not be resolved" twice, then success: no error, two waits.

		Returns:
		        None
		"""
		drive = DriveService(MagicMock(), "drive")
		with patch(INVITE, side_effect=[UNRESOLVED, UNRESOLVED, [{"id": "p1"}]]) as invite:
			with patch(SLEEP) as sleep:
				self.assertEqual(drive.invite_group("item", "group", "write"), [{"id": "p1"}])

		self.assertEqual(invite.call_count, 3)
		self.assertEqual(sleep.call_count, 2)

	def test_group_invite_gives_up_after_the_last_wait(self):
		"""
		Still unknown after every wait: the error is raised.

		Returns:
		        None
		"""
		drive = DriveService(MagicMock(), "drive")
		with patch(INVITE, side_effect=UNRESOLVED) as invite, patch(SLEEP):
			with self.assertRaises(GraphError):
				drive.invite_group("item", "group", "write")

		self.assertEqual(invite.call_count, len(GROUP_WAITS) + 1)

	def test_group_invite_does_not_wait_on_other_errors(self):
		"""
		Any other error (e.g. 403) is raised at once.

		Returns:
		        None
		"""
		drive = DriveService(MagicMock(), "drive")
		with patch(INVITE, side_effect=GraphError(403, "Access denied")) as invite:
			with patch(SLEEP) as sleep, self.assertRaises(GraphError):
				drive.invite_group("item", "group", "write")

		self.assertEqual(invite.call_count, 1)
		sleep.assert_not_called()

	def test_failed_group_access_records_the_reason(self):
		"""
		The group never reaches SharePoint: the folder fails with a reason
		that says so and tells the user to retry.

		Returns:
		        None
		"""
		folder = self.__provision(invite_group=UNRESOLVED)

		self.assertEqual(folder.status, "Failed")
		self.assertIn("Giving the security group access failed", folder.error)
		self.assertIn("does not see the new security group yet", folder.error)
		self.assertIn("Microsoft said: Graph 400", folder.error)

	def test_unknown_owner_records_the_reason(self):
		"""
		An owner who is not a Microsoft 365 user: the reason names the owners.

		Returns:
		        None
		"""
		folder = self.__provision(invite=UNRESOLVED)

		self.assertEqual(folder.status, "Failed")
		self.assertIn("Giving the owners access failed", folder.error)
		self.assertIn("is not a user in the Microsoft 365 tenant", folder.error)

	def test_retry_after_failure_creates_the_folder(self):
		"""
		Retry once the group is known: the folder is created and the error cleared.

		Returns:
		        None
		"""
		folder = self.__provision(invite_group=UNRESOLVED)
		tasks.provision_folder(folder.name)
		folder.reload()

		self.assertEqual(folder.status, "Created", folder.error)
		self.assertFalse(folder.error)
		self.assertEqual(folder.attempts, 2)

	def __provision(self, **failing) -> "frappe.model.document.Document":
		"""
		Provision a new matching ToDo's folder with one DriveService call failing.

		Parameters:
		        failing (dict, required): DriveService method name -> exception it raises.

		Returns:
		        Document: The reloaded SharePoint Folder.
		"""
		todo = make_todo("High")
		name = frappe.db.get_value("SharePoint Folder", {"reference_name": todo.name})
		((method, error),) = failing.items()
		with patch.object(DriveService, method, side_effect=error):
			tasks.provision_folder(name)
		return frappe.get_doc("SharePoint Folder", name)
