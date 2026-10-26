# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Custom fields added to every DocType that has a SharePoint Folder Rule."""

import frappe
from frappe.custom.doctype.custom_field.custom_field import create_custom_fields

LINK_FIELDS = [
	{
		"fieldname": "sharepoint_folder",
		"fieldtype": "Link",
		"label": "SharePoint Folder",
		"options": "SharePoint Folder",
		"read_only": 1,
		"no_copy": 1,
		"insert_after": "sharepoint_section",
	},
	{
		"fieldname": "sharepoint_link",
		"fieldtype": "Data",
		"label": "SharePoint Link",
		"options": "URL",
		"read_only": 1,
		"no_copy": 1,
		"insert_after": "sharepoint_folder",
	},
]


def ensure_reference_fields(doctype: str) -> None:
	"""
	Add a "SharePoint" tab (tabbed forms, e.g. Employee) or section (plain forms)
	with the folder link and URL, placed after the DocType's last field
	(idempotent; existing fields are updated in place).

	Parameters:
	        doctype (str, required): Reference DocType of a rule.

	Returns:
	        None
	"""
	fields = [
		field
		for field in frappe.get_meta(doctype).fields
		if not field.fieldname.startswith("sharepoint_")
	]
	has_tabs = any(field.fieldtype == "Tab Break" for field in fields)
	container = {
		"fieldname": "sharepoint_section",
		"fieldtype": "Tab Break" if has_tabs else "Section Break",
		"label": "SharePoint",
		"collapsible": 0 if has_tabs else 1,
		"insert_after": fields[-1].fieldname if fields else None,
	}
	# Frappe refuses to change a field's type in place; the container holds no
	# data, so drop and recreate it (e.g. after a DocType gains tabs).
	existing = frappe.db.get_value(
		"Custom Field", {"dt": doctype, "fieldname": "sharepoint_section"}, ["name", "fieldtype"]
	)
	if existing and existing[1] != container["fieldtype"]:
		frappe.delete_doc("Custom Field", existing[0], ignore_permissions=True)

	create_custom_fields({doctype: [container, *LINK_FIELDS]}, update=True)
