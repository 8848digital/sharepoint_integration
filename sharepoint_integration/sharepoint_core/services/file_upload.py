# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Background job body: upload one attachment to its document's SharePoint folder."""

import os

import frappe
from frappe import _
from frappe.utils import now_datetime

from sharepoint_integration.sharepoint_core.graph.client import GraphClient
from sharepoint_integration.sharepoint_core.graph.drive import DriveService
from sharepoint_integration.sharepoint_core.graph.upload import preview_url, upload_file
from sharepoint_integration.sharepoint_core.services.folder_request import sanitize_folder_name


class FileUploader:
	"""Uploads a SharePoint File's attachment and, if the rule says so, removes the local copy."""

	def __init__(self, name: str):
		"""
		Load the SharePoint File record.

		Parameters:
		        name (str, required): SharePoint File name.

		Returns:
		        None
		"""
		self.record = frappe.get_doc("SharePoint File", name)

	def run(self) -> None:
		"""
		Upload the file and record the outcome. Waits (stays Queued) while the
		document's folder is not created yet.

		Parameters:
		        None

		Returns:
		        None
		"""
		if self.record.status == "Uploaded":
			return
		folder = frappe.get_doc("SharePoint Folder", self.record.folder)
		if folder.status != "Created":
			return

		attempts = (self.record.attempts or 0) + 1
		try:
			if not frappe.db.exists("File", self.record.file):
				frappe.throw(_("The file was deleted from ERPNext before it could be uploaded"))
			file_doc = frappe.get_doc("File", self.record.file)
			rule = frappe.get_doc("SharePoint Folder Rule", folder.rule)
			client = GraphClient.from_connection(folder.connection)

			parent_id = self.__target_folder(client, folder, rule)
			path = file_doc.get_full_path()
			item = upload_file(client, folder.drive_id, parent_id, file_doc.file_name, path)
			link = preview_url(client, folder.drive_id, item["webUrl"])
		except Exception as error:
			frappe.log_error("SharePoint | File upload failed", frappe.get_traceback())
			self.record.db_set(
				{"status": "Failed", "attempts": attempts, "error": _("Upload failed: {0}").format(str(error))}
			)
			return

		self.record.db_set(
			{
				"status": "Uploaded",
				"attempts": attempts,
				"item_id": item.get("id"),
				"web_url": link,
				"file_size": item.get("size"),
				"uploaded_on": now_datetime(),
				"error": None,
			}
		)
		if rule.remove_local_copy:
			self.record.db_set("local_copy_removed", int(remove_local_copy(file_doc, link)))

	def __target_folder(self, client: GraphClient, folder, rule) -> str:
		"""
		Folder id to upload into: the document folder or its attachment subfolder.

		Parameters:
		        client (GraphClient, required): Graph client.
		        folder (Document, required): SharePoint Folder.
		        rule (Document, required): SharePoint Folder Rule.

		Returns:
		        str: driveItem id of the target folder.
		"""
		subfolder = sanitize_folder_name(rule.attachment_subfolder or "")
		if not subfolder:
			return folder.folder_id
		return DriveService(client, folder.drive_id).ensure_folder(folder.folder_path, subfolder)["id"]


def remove_local_copy(file_doc, web_url: str | None) -> bool:
	"""
	Point the attachment to SharePoint and delete the file from this server.
	Kept when the file is used in a form field (e.g. an image field), since that
	field stores the local URL. The disk file is kept while another File record
	(Frappe de-duplicates identical uploads) still uses it.

	Parameters:
	        file_doc (Document, required): File.
	        web_url (str, optional): SharePoint link of the uploaded file.

	Returns:
	        bool: True when the attachment now points to SharePoint.
	"""
	if file_doc.attached_to_field or not web_url:
		return False

	old_url, path = file_doc.file_url, file_doc.get_full_path()
	frappe.db.set_value(
		"File",
		file_doc.name,
		{"file_url": web_url, "thumbnail_url": None, "content_hash": None},
		update_modified=False,
	)
	shared = frappe.db.exists("File", {"file_url": old_url, "name": ["!=", file_doc.name]})
	if not shared and os.path.exists(path):
		os.remove(path)
	return True
