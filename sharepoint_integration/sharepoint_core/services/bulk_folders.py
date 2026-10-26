# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""
"Create Folders for Existing Records": folders (and attachment uploads) for
documents that existed before a Folder Rule was set up, e.g. at go-live.
"""

import frappe
from frappe import _

from sharepoint_integration.sharepoint_core.services.file_request import queue_document_files
from sharepoint_integration.sharepoint_core.services.folder_request import request_folder_for_rule
from sharepoint_integration.sharepoint_core.services.rule_matcher import condition_holds

BATCH_SIZE = 500
FOLDER_FIELDS = ["name", "status", "rule", "reference_doctype", "reference_name"]


def count_existing(rule_name: str) -> dict:
	"""
	How many records the bulk run would look at, for the confirmation dialog.

	Parameters:
	        rule_name (str, required): SharePoint Folder Rule name.

	Returns:
	        dict: {"records": <matching the trigger>, "with_folder": <already have a folder>}
	"""
	rule = __load_rule(rule_name)
	return {
		"records": frappe.db.count(rule.reference_doctype, __base_filters(rule)),
		"with_folder": frappe.db.count(
			"SharePoint Folder", {"reference_doctype": rule.reference_doctype}
		),
	}


def queue_existing(rule_name: str) -> None:
	"""
	Start the bulk run in the background; the user gets a message when it ends.

	Parameters:
	        rule_name (str, required): SharePoint Folder Rule name.

	Returns:
	        None
	"""
	__load_rule(rule_name)
	frappe.enqueue(
		"sharepoint_integration.sharepoint_core.tasks.create_existing_folders",
		queue="long",
		timeout=4 * 3600,
		job_id=f"sharepoint-bulk-{rule_name}",
		deduplicate=True,
		rule=rule_name,
		user=frappe.session.user,
	)


def create_existing_folders(rule_name: str, user: str | None = None) -> dict:
	"""
	Queue a folder for every record that matches the rule and has none; for
	records whose folder already exists, queue their attachments (when the
	rule uploads attachments). Commits every BATCH_SIZE records.

	Parameters:
	        rule_name (str, required): SharePoint Folder Rule name.
	        user (str, optional): User to notify when done.

	Returns:
	        dict: {"folders": <queued>, "files": <uploads queued>, "skipped": <condition not met>}
	"""
	rule = frappe.get_doc("SharePoint Folder Rule", rule_name)
	folders = {
		row.reference_name: row
		for row in frappe.get_all(
			"SharePoint Folder", filters={"reference_doctype": rule.reference_doctype}, fields=FOLDER_FIELDS
		)
	}
	counts = {"folders": 0, "files": 0, "skipped": 0}

	for position, name in enumerate(__record_names(rule), start=1):
		__process_record(rule, name, folders.get(name), counts)
		# A go-live run can cover thousands of records: commit per batch so queued
		# folder jobs start (enqueue_after_commit) and one failure keeps earlier batches.
		if position % BATCH_SIZE == 0:
			frappe.db.commit()  # nosemgrep

	frappe.db.commit()  # nosemgrep
	if user:
		__notify(user, rule_name, counts)
	return counts


def __process_record(rule, name: str, folder, counts: dict) -> None:
	"""
	Queue the record's folder, or its attachments when the folder already exists.

	Parameters:
	        rule (Document, required): SharePoint Folder Rule.
	        name (str, required): Reference document name.
	        folder (dict, optional): Existing SharePoint Folder row.
	        counts (dict, required): Running totals, updated in place.

	Returns:
	        None
	"""
	if folder:
		if folder.status == "Created":
			counts["files"] += queue_document_files(folder)
		return

	doc = frappe.get_doc(rule.reference_doctype, name)
	if not condition_holds(rule.name, rule.condition, doc):
		counts["skipped"] += 1
		return
	request_folder_for_rule(doc, rule.name)
	counts["folders"] += 1


def __record_names(rule):
	"""
	Yield matching record names page by page (oldest first), so huge DocTypes
	are never loaded at once.

	Parameters:
	        rule (Document, required): SharePoint Folder Rule.

	Returns:
	        Iterator[str]: Record names.
	"""
	start = 0
	while True:
		page = frappe.get_all(
			rule.reference_doctype,
			filters=__base_filters(rule),
			pluck="name",
			order_by="creation asc",
			limit_start=start,
			limit_page_length=BATCH_SIZE,
		)
		if not page:
			return
		yield from page
		start += BATCH_SIZE


def __base_filters(rule) -> dict:
	"""
	Records the rule's trigger would cover: not cancelled, and in the rule's
	workflow state for "On Workflow State" rules.

	Parameters:
	        rule (Document, required): SharePoint Folder Rule.

	Returns:
	        dict: get_all filters.
	"""
	filters = {}
	if frappe.get_meta(rule.reference_doctype).is_submittable:
		filters["docstatus"] = ["<", 2]
	if rule.trigger_event == "On Workflow State":
		filters["workflow_state"] = rule.workflow_state
	return filters


def __load_rule(rule_name: str):
	"""
	Load a rule the user may edit, refusing a disabled rule or integration.

	Parameters:
	        rule_name (str, required): SharePoint Folder Rule name.

	Returns:
	        Document: SharePoint Folder Rule.
	"""
	rule = frappe.get_doc("SharePoint Folder Rule", rule_name)
	rule.check_permission("write")
	if not frappe.db.get_single_value("SharePoint Settings", "enabled"):
		frappe.throw(_("SharePoint Integration is disabled in SharePoint Settings"))
	if not rule.enabled:
		frappe.throw(_("Enable the rule first"))
	return rule


def __notify(user: str, rule_name: str, counts: dict) -> None:
	"""
	Show the result to the user who started the run.

	Parameters:
	        user (str, required): User.
	        rule_name (str, required): SharePoint Folder Rule name.
	        counts (dict, required): Totals.

	Returns:
	        None
	"""
	message = _(
		"Rule {0}: {1} folder(s) queued, {2} attachment upload(s) queued, "
		"{3} record(s) skipped because the condition did not match."
	).format(rule_name, counts["folders"], counts["files"], counts["skipped"])
	frappe.publish_realtime(
		"msgprint", {"message": message, "title": _("SharePoint"), "indicator": "green"}, user=user
	)
