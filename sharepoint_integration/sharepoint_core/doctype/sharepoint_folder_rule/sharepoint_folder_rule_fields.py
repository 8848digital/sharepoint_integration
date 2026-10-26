# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Validation of a Folder Rule's owner, date and template fields."""

import frappe
from frappe import _

from sharepoint_integration.sharepoint_core.services.folder_request import sanitize_folder_name

USER_FIELD_TYPES = ("Link", "Data", "Dynamic Link")
DATE_FIELD_TYPES = ("Date", "Datetime")


def validate_owner_fields(doc) -> None:
	"""
	Owner fields must exist on the reference DocType and hold a User.

	Parameters:
	        doc (Document, required): SharePoint Folder Rule.

	Returns:
	        None
	"""
	meta = frappe.get_meta(doc.reference_doctype)
	seen = set()
	for row in doc.owner_fields:
		field = meta.get_field(row.fieldname)
		if not field or field.fieldtype not in USER_FIELD_TYPES:
			frappe.throw(
				_("Row {0}: {1} is not a Link/Data field of {2}").format(
					row.idx, row.fieldname, doc.reference_doctype
				)
			)
		if field.fieldtype == "Link" and field.options != "User":
			frappe.throw(
				_("Row {0}: {1} must link to User, not {2}").format(row.idx, row.fieldname, field.options)
			)
		if row.fieldname in seen:
			frappe.throw(_("Row {0}: Owner field {1} is listed twice").format(row.idx, row.fieldname))
		seen.add(row.fieldname)

	if not doc.owner_fields and not doc.create_security_group:
		frappe.msgprint(
			_("No owner fields and no security group: only site owners will reach these folders"),
			alert=True,
		)


def validate_date_fields(doc) -> None:
	"""
	Optional start/end fields must be Date/Datetime fields of the reference DocType.

	Parameters:
	        doc (Document, required): SharePoint Folder Rule.

	Returns:
	        None
	"""
	meta = frappe.get_meta(doc.reference_doctype)
	for label, fieldname in (
		(_("Start Date Field"), doc.start_date_field),
		(_("End Date Field"), doc.end_date_field),
	):
		field = meta.get_field(fieldname) if fieldname else None
		if fieldname and (not field or field.fieldtype not in DATE_FIELD_TYPES):
			frappe.throw(
				_("{0}: {1} is not a Date field of {2}").format(label, fieldname, doc.reference_doctype)
			)
	if doc.start_date_field and doc.start_date_field == doc.end_date_field:
		frappe.throw(_("Start Date Field and End Date Field must be different"))


def validate_template_file(doc) -> None:
	"""
	Template copy needs a file id and a valid file name with an extension.

	Parameters:
	        doc (Document, required): SharePoint Folder Rule.

	Returns:
	        None
	"""
	if not doc.attach_template:
		return
	if not doc.template_item_id:
		frappe.throw(_("Template File ID is required to copy a template"))
	name = sanitize_folder_name(doc.template_file_name or "")
	if not name or name != (doc.template_file_name or "").strip() or "." not in name:
		frappe.throw(
			_("File Name In Folder must be a valid file name with an extension, e.g. Proposal.pptx")
		)
