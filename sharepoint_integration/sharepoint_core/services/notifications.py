# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Failure emails for SharePoint folder/access jobs."""

import frappe
from frappe import _


def notify_failure(record, subject: str | None = None) -> None:
	"""
	Email the Settings' failure recipients about a failed folder or access job.

	Parameters:
	        record (Document, required): SharePoint Folder or SharePoint Access that failed.
	        subject (str, optional): Email subject; defaults to a folder-failure subject.

	Returns:
	        None
	"""
	settings = frappe.get_cached_doc("SharePoint Settings")
	if not settings.notify_on_failure:
		return

	recipients = [row.user for row in settings.failure_recipients]
	if not recipients:
		return

	message = _("SharePoint job failed for {0} {1}.<br><br>{2}").format(
		record.reference_doctype,
		record.reference_name,
		frappe.utils.escape_html(record.get("error") or ""),
	)
	frappe.sendmail(
		recipients=recipients,
		subject=subject or _("SharePoint folder creation failed"),
		message=message,
		reference_doctype=record.doctype,
		reference_name=record.name,
	)
