# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Attachments uploaded to a document's SharePoint folder, in the sandbox."""

import os
from unittest.mock import patch

import frappe
from frappe.tests.utils import FrappeTestCase

from sharepoint_integration.tests.sharepoint_files_helpers import (
	attach,
	collect_jobs,
	run_jobs,
	sandbox_item,
	todo_with_folder,
)
from sharepoint_integration.tests.sharepoint_fixtures import RULE, make_todo, setup_sandbox

UPLOAD = "sharepoint_integration.sharepoint_core.graph.upload"


class TestSharePointFiles(FrappeTestCase):
	"""Attachment upload, local copy removal and chunked upload."""

	@classmethod
	def setUpClass(cls):
		"""
		Sandbox connections and rules; the tenant A rule uploads attachments.

		Returns:
		        None
		"""
		super().setUpClass()
		setup_sandbox()
		frappe.db.set_value(
			"SharePoint Folder Rule",
			RULE,
			{"upload_attachments": 1, "attachment_subfolder": "Attachments", "remove_local_copy": 0},
		)

	def setUp(self):
		"""
		Run background jobs inline, on demand.

		Returns:
		        None
		"""
		self.jobs = collect_jobs(self)

	def test_attachment_is_uploaded_into_the_subfolder(self):
		"""
		A file attached after the folder exists lands in <folder>/Attachments.

		Returns:
		        None
		"""
		todo = todo_with_folder(self.jobs)
		record = self.__upload(attach(todo, "report.txt", b"quarterly numbers"))

		self.assertEqual(record.status, "Uploaded")
		self.assertIn("/Forms/AllItems.aspx?id=", record.web_url)
		item = sandbox_item(record.item_id)
		self.assertTrue(item["path"].endswith("/Attachments/report.txt"))
		self.assertEqual(item["size"], len(b"quarterly numbers"))

	def test_file_attached_before_the_folder_waits_and_uploads(self):
		"""
		An attachment added while the folder is Queued uploads once it is created.

		Returns:
		        None
		"""
		todo = make_todo("High")
		file = attach(todo, "early.txt", b"before the folder")
		record = frappe.get_doc("SharePoint File", {"file": file.name})
		self.assertEqual(record.status, "Queued")

		run_jobs(self.jobs)
		record.reload()
		self.assertEqual(record.status, "Uploaded")

	def test_same_name_is_not_overwritten(self):
		"""
		Two attachments with the same name both stay in SharePoint.

		Returns:
		        None
		"""
		todo = todo_with_folder(self.jobs)
		first = self.__upload(attach(todo, "scan.txt", b"one"))
		second = self.__upload(attach(todo, "scan.txt", b"two!"))

		self.assertNotEqual(first.item_id, second.item_id)
		self.assertEqual(sandbox_item(first.item_id)["size"], len(b"one"))
		self.assertEqual(sandbox_item(second.item_id)["size"], len(b"two!"))

	def test_local_copy_removed_after_upload(self):
		"""
		With "Remove Local Copy" the attachment points to SharePoint and the disk file is gone.

		Returns:
		        None
		"""
		self.__remove_local_copies()
		todo = todo_with_folder(self.jobs)
		file = attach(todo, "big-scan.txt", b"remove me from the server")
		path = file.get_full_path()
		record = self.__upload(file)

		self.assertTrue(record.local_copy_removed)
		self.assertEqual(frappe.db.get_value("File", file.name, "file_url"), record.web_url)
		self.assertFalse(os.path.exists(path))

	def test_file_used_in_a_field_keeps_its_local_copy(self):
		"""
		A file referenced by a form field is uploaded but never removed locally.

		Returns:
		        None
		"""
		self.__remove_local_copies()
		todo = todo_with_folder(self.jobs)
		file = attach(todo, "photo.txt", b"used by a field", attached_to_field="description")
		record = self.__upload(file)

		self.assertEqual(record.status, "Uploaded")
		self.assertFalse(record.local_copy_removed)
		self.assertTrue(os.path.exists(file.get_full_path()))

	def test_large_file_uses_an_upload_session(self):
		"""
		Files above the simple-upload limit go up in chunks without a bearer token
		(the sandbox rejects one, like Microsoft).

		Returns:
		        None
		"""
		todo = todo_with_folder(self.jobs)
		content = b"0123456789" * 5
		with patch(f"{UPLOAD}.SIMPLE_UPLOAD_LIMIT", 10), patch(f"{UPLOAD}.CHUNK_SIZE", 16):
			record = self.__upload(attach(todo, "large.bin", content))

		self.assertEqual(record.status, "Uploaded", record.error)
		self.assertEqual(sandbox_item(record.item_id)["size"], len(content))

	def __remove_local_copies(self) -> None:
		"""
		Turn on "Remove Local Copy" for this test only.

		Returns:
		        None
		"""
		frappe.db.set_value("SharePoint Folder Rule", RULE, "remove_local_copy", 1)
		self.addCleanup(frappe.db.set_value, "SharePoint Folder Rule", RULE, "remove_local_copy", 0)

	def __upload(self, file):
		"""
		Run the queued jobs and return the file's SharePoint File record.

		Parameters:
		        file (Document, required): The File.

		Returns:
		        Document: SharePoint File.
		"""
		run_jobs(self.jobs)
		return frappe.get_doc("SharePoint File", {"file": file.name})
