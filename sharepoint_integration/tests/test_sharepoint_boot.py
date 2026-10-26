# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""The SharePoint form button follows SharePoint Settings and the Folder Rules."""

import frappe
from frappe.tests.utils import FrappeTestCase

from sharepoint_integration.boot_session import boot_session
from sharepoint_integration.sharepoint_core.services.folder_request import request_folder_manually
from sharepoint_integration.tests.sharepoint_fixtures import make_todo, setup_sandbox


class TestSharePointBoot(FrappeTestCase):
	"""Boot data for the form buttons and the disabled-integration guard."""

	@classmethod
	def setUpClass(cls):
		"""
		Sandbox connections and ToDo rules.

		Returns:
		        None
		"""
		super().setUpClass()
		setup_sandbox()

	def test_enabled_lists_rule_doctypes(self):
		"""
		With SharePoint enabled, DocTypes with an enabled rule are sent to the desk.

		Returns:
		        None
		"""
		bootinfo = frappe._dict()
		boot_session(bootinfo)
		self.assertTrue(bootinfo.sharepoint["enabled"])
		self.assertIn("ToDo", bootinfo.sharepoint["doctypes"])

	def test_disabled_hides_buttons_and_refuses_folders(self):
		"""
		With SharePoint disabled, no DocType gets the button and a manual
		"Create Folder" call explains why instead of claiming it was queued.

		Returns:
		        None
		"""
		todo = make_todo("High")
		frappe.db.set_single_value("SharePoint Settings", "enabled", 0)
		self.addCleanup(frappe.db.set_single_value, "SharePoint Settings", "enabled", 1)

		bootinfo = frappe._dict()
		boot_session(bootinfo)
		self.assertEqual(bootinfo.sharepoint, {"enabled": False, "doctypes": []})
		with self.assertRaises(frappe.ValidationError):
			request_folder_manually("ToDo", todo.name)
