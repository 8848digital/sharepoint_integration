# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Validation for SharePoint Folder Rule."""

import frappe
from frappe import _
from frappe.utils.jinja import validate_template

from sharepoint_integration.sharepoint_core.doctype.sharepoint_folder_rule.sharepoint_folder_rule_fields import (
	validate_date_fields,
	validate_owner_fields,
	validate_template_file,
)
from sharepoint_integration.sharepoint_core.services.folder_request import sanitize_folder_name

OWN_DOCTYPES = (
	"SharePoint Folder",
	"SharePoint Access",
	"SharePoint Folder Rule",
	"SharePoint Settings",
)
LOCATION_FIELDS = ("connection", "site_name", "library_name")


def validate_rule(doc) -> None:
	"""
	Validate every section of a SharePoint Folder Rule.

	Parameters:
	        doc (Document, required): SharePoint Folder Rule.

	Returns:
	        None
	"""
	__validate_connection(doc)
	__validate_reference_doctype(doc.reference_doctype)
	__validate_condition(doc.condition)
	__validate_workflow(doc)
	__validate_location(doc)
	__validate_subfolders(doc)
	validate_owner_fields(doc)
	validate_date_fields(doc)
	validate_template_file(doc)
	__warn_on_overlap(doc)


def __validate_connection(doc) -> None:
	"""
	Default the connection from Settings and require it to be enabled.

	Parameters:
	        doc (Document, required): SharePoint Folder Rule.

	Returns:
	        None
	"""
	doc.connection = doc.connection or frappe.db.get_single_value(
		"SharePoint Settings", "default_connection"
	)
	if not doc.connection:
		frappe.throw(
			_("Select a SharePoint Connection (or set a Default Connection in SharePoint Settings)")
		)
	if doc.enabled and not frappe.db.get_value("SharePoint Connection", doc.connection, "enabled"):
		frappe.throw(_("SharePoint Connection {0} is disabled").format(doc.connection))


def __validate_reference_doctype(doctype: str) -> None:
	"""
	Only normal, stored, non-child DocTypes can own folders.

	Parameters:
	        doctype (str, required): Reference DocType.

	Returns:
	        None
	"""
	meta = frappe.get_meta(doctype)
	if meta.istable or meta.issingle or meta.is_virtual or doctype in OWN_DOCTYPES:
		frappe.throw(
			_("{0} cannot have SharePoint folders (child, single, virtual or internal DocType)").format(
				doctype
			)
		)


def __validate_condition(condition: str | None) -> None:
	"""
	The condition must be a single valid Python expression.

	Parameters:
	        condition (str, optional): Python expression using `doc`.

	Returns:
	        None
	"""
	if not (condition or "").strip():
		return
	try:
		compile(condition, "<condition>", "eval")
	except SyntaxError as error:
		frappe.throw(_("Condition is not a valid Python expression: {0}").format(error.msg))


def __validate_workflow(doc) -> None:
	"""
	Workflow triggers need a state that exists in the DocType's active workflow.

	Parameters:
	        doc (Document, required): SharePoint Folder Rule.

	Returns:
	        None
	"""
	if doc.trigger_event != "On Workflow State":
		doc.workflow_state = doc.from_workflow_state = None
		return
	if doc.workflow_state == doc.from_workflow_state:
		frappe.throw(_("From Workflow State must differ from Workflow State"))

	workflow = frappe.db.get_value(
		"Workflow", {"document_type": doc.reference_doctype, "is_active": 1}
	)
	if not workflow:
		frappe.throw(
			_("{0} has no active Workflow; choose another trigger").format(doc.reference_doctype)
		)

	states = set(
		frappe.get_all("Workflow Document State", filters={"parent": workflow}, pluck="state")
	)
	unknown = [
		state for state in (doc.workflow_state, doc.from_workflow_state) if state and state not in states
	]
	if unknown:
		frappe.throw(_("Workflow {0} has no state(s): {1}").format(workflow, ", ".join(unknown)))


def __validate_location(doc) -> None:
	"""
	Validate folder-name template and path; reset cached ids if the site changed.

	Parameters:
	        doc (Document, required): SharePoint Folder Rule.

	Returns:
	        None
	"""
	validate_template(doc.folder_name_template)
	if doc.site_name and ("/" in doc.site_name.strip("/") or " " in doc.site_name):
		frappe.throw(
			_("Site Name is the URL segment after /sites/ (e.g. Proposals), without slashes or spaces")
		)

	segments = [segment for segment in (doc.base_path or "").split("/") if segment.strip()]
	if any(sanitize_folder_name(segment) != segment.strip() for segment in segments):
		frappe.throw(_('Base Path contains characters SharePoint does not allow: " * : < > ? \\ | # %'))
	doc.base_path = "/".join(segment.strip() for segment in segments)

	if not doc.is_new() and any(doc.has_value_changed(field) for field in LOCATION_FIELDS):
		doc.site_id = doc.drive_id = None


def __validate_subfolders(doc) -> None:
	"""
	Subfolder names must be valid for SharePoint and unique (case-insensitive).

	Parameters:
	        doc (Document, required): SharePoint Folder Rule.

	Returns:
	        None
	"""
	seen = set()
	for row in doc.sub_folders:
		row.folder_name = sanitize_folder_name(row.folder_name)
		if not row.folder_name:
			frappe.throw(
				_("Row {0}: Subfolder Name is empty after removing invalid characters").format(row.idx)
			)
		if row.folder_name.lower() in seen:
			frappe.throw(_("Row {0}: Subfolder {1} is listed twice").format(row.idx, row.folder_name))
		seen.add(row.folder_name.lower())


def __warn_on_overlap(doc) -> None:
	"""
	Warn when another enabled rule for the DocType has the same trigger and priority
	(then only creation order decides which one wins).

	Parameters:
	        doc (Document, required): SharePoint Folder Rule.

	Returns:
	        None
	"""
	if not doc.enabled:
		return
	clash = frappe.db.exists(
		"SharePoint Folder Rule",
		{
			"name": ["!=", doc.name],
			"enabled": 1,
			"reference_doctype": doc.reference_doctype,
			"trigger_event": doc.trigger_event,
			"priority": doc.priority,
			"condition": doc.condition or ["is", "not set"],
		},
	)
	if clash:
		frappe.throw(
			_("Rule {0} already has the same DocType, trigger, condition and priority").format(clash)
		)
