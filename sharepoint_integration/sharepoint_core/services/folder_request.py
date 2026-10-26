# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Turn a document event (or a button click) into a queued SharePoint Folder job."""

import re

import frappe
from frappe import _

from sharepoint_integration.sharepoint_core.services.rule_matcher import find_matching_rule

INVALID_NAME_CHARS = re.compile(r'[\x00-\x1f"*:<>?/\\|#%]')
MAX_NAME_LENGTH = 200


def request_folder(doc, event: str) -> str | None:
	"""
	Queue folder creation for a document if a rule matches and no folder exists yet.
	Runs inside the document's save, so it only inserts a small record and
	enqueues; all Graph calls happen in the background job.

	Parameters:
	        doc (Document, required): The reference document.
	        event (str, required): "insert", "update" or "manual".

	Returns:
	        str | None: SharePoint Folder name when a job was queued.
	"""
	if not frappe.db.get_single_value("SharePoint Settings", "enabled"):
		return None

	existing = frappe.db.get_value(
		"SharePoint Folder",
		{"reference_doctype": doc.doctype, "reference_name": doc.name},
		["name", "status"],
		as_dict=True,
	)
	if existing and not (event == "manual" and existing.status == "Failed"):
		return existing.name if event == "manual" else None

	rule = find_matching_rule(doc, event)
	if not rule:
		if event == "manual":
			frappe.throw(_("No enabled SharePoint Folder Rule matches this {0}").format(_(doc.doctype)))
		return None

	if existing:
		folder_name = existing.name
		frappe.db.set_value("SharePoint Folder", folder_name, {"status": "Queued", "error": None})
	else:
		folder_name = __insert_folder(doc, rule)
	enqueue_provisioning(folder_name)
	return folder_name


def request_folder_for_rule(doc, rule: str) -> str:
	"""
	Queue a folder for an existing document with a given rule (bulk run).
	The caller has checked the rule's condition and that no folder exists.

	Parameters:
	        doc (Document, required): The reference document.
	        rule (str, required): SharePoint Folder Rule name.

	Returns:
	        str: New SharePoint Folder name.
	"""
	folder_name = __insert_folder(doc, rule)
	enqueue_provisioning(folder_name)
	return folder_name


def request_folder_manually(doctype: str, docname: str) -> dict:
	"""
	Create (or retry) the folder of a document from the "Create SharePoint Folder" button.

	Parameters:
	        doctype (str, required): Reference DocType.
	        docname (str, required): Reference document name.

	Returns:
	        dict: {"folder": <SharePoint Folder name>, "status": <status>}
	"""
	doc = frappe.get_doc(doctype, docname)
	doc.check_permission("write")
	if not frappe.db.get_single_value("SharePoint Settings", "enabled"):
		frappe.throw(_("SharePoint Integration is disabled in SharePoint Settings"))
	folder = request_folder(doc, "manual")
	return {"folder": folder, "status": frappe.db.get_value("SharePoint Folder", folder, "status")}


def retry_folder(folder: str) -> None:
	"""
	Re-queue a failed SharePoint Folder.

	Parameters:
	        folder (str, required): SharePoint Folder name.

	Returns:
	        None
	"""
	record = frappe.get_doc("SharePoint Folder", folder)
	frappe.get_doc(record.reference_doctype, record.reference_name).check_permission("write")
	if record.status == "Created":
		frappe.throw(_("Folder is already created"))
	record.db_set({"status": "Queued", "error": None})
	enqueue_provisioning(folder)


def enqueue_provisioning(folder: str) -> None:
	"""
	Enqueue the provisioning job once the current transaction commits.

	Parameters:
	        folder (str, required): SharePoint Folder name.

	Returns:
	        None
	"""
	frappe.enqueue(
		"sharepoint_integration.sharepoint_core.tasks.provision_folder",
		queue="long",
		timeout=1800,
		job_id=f"sharepoint-folder-{folder}",
		deduplicate=True,
		enqueue_after_commit=True,
		folder=folder,
	)


def sanitize_folder_name(name: str) -> str:
	"""
	Make a string safe as a SharePoint folder name. Unicode (e.g. Arabic) is kept;
	only characters SharePoint rejects are removed.
	e.g. '  ACME: "Q3" RFP/2026. ' -> 'ACME Q3 RFP2026'

	Parameters:
	        name (str, required): Raw folder name.

	Returns:
	        str: Sanitized name ("" if nothing usable remains).
	"""
	cleaned = INVALID_NAME_CHARS.sub("", name or "")
	cleaned = re.sub(r"\s+", " ", cleaned).strip(" .")
	return cleaned[:MAX_NAME_LENGTH].rstrip(" .")


def __insert_folder(doc, rule: str) -> str:
	"""
	Insert the Queued SharePoint Folder record with the rendered folder name.

	Parameters:
	        doc (Document, required): The reference document.
	        rule (str, required): Matching SharePoint Folder Rule name.

	Returns:
	        str: New SharePoint Folder name.
	"""
	template, connection = frappe.db.get_value(
		"SharePoint Folder Rule", rule, ["folder_name_template", "connection"]
	)
	# The template comes only from SharePoint Folder Rule, which only System Managers
	# can edit, and Frappe renders it in its sandboxed Jinja environment.
	# nosemgrep: frappe-semgrep-rules.rules.security.frappe-ssti
	folder_name = sanitize_folder_name(frappe.render_template(template, {"doc": doc})) or doc.name

	record = frappe.get_doc(
		{
			"doctype": "SharePoint Folder",
			"reference_doctype": doc.doctype,
			"reference_name": doc.name,
			"rule": rule,
			"connection": connection,
			"status": "Queued",
			"folder_name": folder_name,
		}
	)
	record.insert(ignore_permissions=True)
	return record.name
