# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Decide which attachments go to SharePoint and queue their upload jobs."""

import frappe
from frappe import _

FILE_FIELDS = [
	"name",
	"file_name",
	"file_url",
	"is_folder",
	"attached_to_doctype",
	"attached_to_name",
]
# Hourly retries stop after this many attempts; "Retry Upload" still works after that.
MAX_AUTO_ATTEMPTS = 3


def queue_file_upload(file_doc) -> None:
	"""
	A file was attached: queue its upload when the document's folder rule
	uploads attachments. If the folder is still being created, the record
	waits as Queued and is picked up once the folder exists.

	Parameters:
	        file_doc (Document, required): The new File.

	Returns:
	        None
	"""
	if not __is_uploadable(file_doc) or not frappe.db.get_single_value(
		"SharePoint Settings", "enabled"
	):
		return

	folder = frappe.db.get_value(
		"SharePoint Folder",
		{
			"reference_doctype": file_doc.attached_to_doctype,
			"reference_name": file_doc.attached_to_name,
		},
		["name", "status", "rule"],
		as_dict=True,
	)
	if not folder or not __rule_uploads(folder.rule):
		return
	if frappe.db.exists("SharePoint File", {"file": file_doc.name}):
		return

	record = __insert_record(file_doc, folder.name)
	if folder.status == "Created":
		enqueue_upload(record)


def queue_document_files(folder) -> int:
	"""
	Queue every attachment of a folder's document that is not uploaded yet.
	Used when the folder is created and by "Create Folders for Existing Records".

	Parameters:
	        folder (Document | dict, required): SharePoint Folder (name, rule, reference).

	Returns:
	        int: Number of uploads queued.
	"""
	if not __rule_uploads(folder.rule):
		return 0

	files = frappe.get_all(
		"File",
		filters={
			"attached_to_doctype": folder.reference_doctype,
			"attached_to_name": folder.reference_name,
		},
		fields=FILE_FIELDS,
	)
	queued = 0
	for file in files:
		if not __is_uploadable(file):
			continue
		record = frappe.db.get_value(
			"SharePoint File", {"file": file.name}, ["name", "status"], as_dict=True
		)
		if record and record.status == "Uploaded":
			continue
		enqueue_upload(record.name if record else __insert_record(file, folder.name))
		queued += 1
	return queued


def retry_upload(name: str) -> None:
	"""
	Re-queue a failed upload ("Retry Upload" button).

	Parameters:
	        name (str, required): SharePoint File name.

	Returns:
	        None
	"""
	record = frappe.get_doc("SharePoint File", name)
	frappe.get_doc(record.reference_doctype, record.reference_name).check_permission("write")
	if record.status == "Uploaded":
		frappe.throw(_("File is already uploaded"))
	record.db_set({"status": "Queued", "error": None})
	enqueue_upload(name)


def retry_waiting_uploads() -> None:
	"""
	Scheduler: re-queue Failed uploads (up to MAX_AUTO_ATTEMPTS) and Queued
	ones whose folder now exists.

	Returns:
	        None
	"""
	created = frappe.get_all("SharePoint Folder", filters={"status": "Created"}, pluck="name")
	if not created:
		return
	names = frappe.get_all(
		"SharePoint File",
		filters={
			"status": ["in", ["Queued", "Failed"]],
			"attempts": ["<", MAX_AUTO_ATTEMPTS],
			"folder": ["in", created],
		},
		pluck="name",
	)
	for name in names:
		enqueue_upload(name)


def enqueue_upload(name: str) -> None:
	"""
	Enqueue the upload job once the current transaction commits.

	Parameters:
	        name (str, required): SharePoint File name.

	Returns:
	        None
	"""
	frappe.enqueue(
		"sharepoint_integration.sharepoint_core.tasks.upload_file",
		queue="long",
		timeout=3600,
		job_id=f"sharepoint-file-{name}",
		deduplicate=True,
		enqueue_after_commit=True,
		sharepoint_file=name,
	)


def __is_uploadable(file) -> bool:
	"""
	Only real files attached to a document; links (http...) and folders are skipped.

	Parameters:
	        file (Document | dict, required): File.

	Returns:
	        bool: True when the file can be uploaded.
	"""
	url = file.get("file_url") or ""
	return bool(
		file.get("attached_to_doctype")
		and file.get("attached_to_name")
		and not file.get("is_folder")
		and url
		and not url.startswith(("http://", "https://"))
	)


def __rule_uploads(rule: str | None) -> bool:
	"""
	Whether the folder's rule has "Upload Attachments" ticked.

	Parameters:
	        rule (str, optional): SharePoint Folder Rule name.

	Returns:
	        bool: True when attachments go to SharePoint.
	"""
	return bool(rule and frappe.db.get_value("SharePoint Folder Rule", rule, "upload_attachments"))


def __insert_record(file, folder: str) -> str:
	"""
	Insert the Queued SharePoint File record.

	Parameters:
	        file (Document | dict, required): File.
	        folder (str, required): SharePoint Folder name.

	Returns:
	        str: New SharePoint File name.
	"""
	record = frappe.get_doc(
		{
			"doctype": "SharePoint File",
			"file": file.name,
			"file_name": file.file_name,
			"reference_doctype": file.attached_to_doctype,
			"reference_name": file.attached_to_name,
			"folder": folder,
			"status": "Queued",
		}
	)
	record.insert(ignore_permissions=True)
	return record.name
