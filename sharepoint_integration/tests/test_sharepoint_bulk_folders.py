# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

""""Create Folders for Existing Records" on a Folder Rule, in the sandbox."""

import frappe
from frappe.tests.utils import FrappeTestCase

from sharepoint_integration.sharepoint_core.services.bulk_folders import create_existing_folders
from sharepoint_integration.tests.sharepoint_files_helpers import (
	attach,
	collect_jobs,
	run_jobs,
	todo_with_folder,
)
from sharepoint_integration.tests.sharepoint_fixtures import RULE, make_todo, setup_sandbox


class TestSharePointBulkFolders(FrappeTestCase):
	"""Folders and attachment uploads for records that existed before the rule."""

	@classmethod
	def setUpClass(cls):
		"""
		Sandbox connections and rules; the tenant A rule uploads attachments.

		Returns:
		        None
		"""
		super().setUpClass()
		setup_sandbox()
		frappe.db.set_value("SharePoint Folder Rule", RULE, "upload_attachments", 1)

	def setUp(self):
		"""
		Run background jobs inline, on demand.

		Returns:
		        None
		"""
		self.jobs = collect_jobs(self)

	def test_folders_only_for_matching_records(self):
		"""
		Records created while the integration was off get folders when they
		match the condition; others are skipped.

		Returns:
		        None
		"""
		frappe.db.set_single_value("SharePoint Settings", "enabled", 0)
		matching, other = make_todo("High"), make_todo("Low")
		frappe.db.set_single_value("SharePoint Settings", "enabled", 1)

		counts = create_existing_folders(RULE)
		run_jobs(self.jobs)

		self.assertGreaterEqual(counts["folders"], 1)
		self.assertGreaterEqual(counts["skipped"], 1)
		self.assertEqual(
			frappe.db.get_value("SharePoint Folder", {"reference_name": matching.name}, "status"), "Created"
		)
		self.assertFalse(frappe.db.exists("SharePoint Folder", {"reference_name": other.name}))

	def test_second_run_creates_nothing_new(self):
		"""
		Running again does not duplicate folders.

		Returns:
		        None
		"""
		create_existing_folders(RULE)
		run_jobs(self.jobs)
		before = frappe.db.count("SharePoint Folder")

		self.assertEqual(create_existing_folders(RULE)["folders"], 0)
		self.assertEqual(frappe.db.count("SharePoint Folder"), before)

	def test_uploads_attachments_of_existing_folders(self):
		"""
		A document whose folder already exists gets its old attachments uploaded.

		Returns:
		        None
		"""
		frappe.db.set_value("SharePoint Folder Rule", RULE, "upload_attachments", 0)
		todo = todo_with_folder(self.jobs)
		file = attach(todo, "old.txt", b"attached before uploads were on")
		frappe.db.set_value("SharePoint Folder Rule", RULE, "upload_attachments", 1)
		self.assertFalse(frappe.db.exists("SharePoint File", {"file": file.name}))

		create_existing_folders(RULE)
		run_jobs(self.jobs)
		self.assertEqual(
			frappe.db.get_value("SharePoint File", {"file": file.name}, "status"), "Uploaded"
		)
