# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Plain-language reason saved on a SharePoint Folder when provisioning fails."""

import frappe
from frappe import _

from sharepoint_integration.sharepoint_core.graph.common import GraphError, is_unresolved_recipient
from sharepoint_integration.sharepoint_core.services.notifications import notify_failure

STEP_FOLDER = "Creating the folder"
STEP_GROUP = "Creating the security group"
STEP_OWNER_ACCESS = "Giving the owners access"
STEP_GROUP_ACCESS = "Giving the security group access"
STEP_SUBFOLDERS = "Creating the subfolders"
STEP_RESTRICT = "Removing access copied from the parent folder"

STATUS_HINTS = {
	401: "The Client ID or Client Secret of the SharePoint Connection is wrong or expired.",
	403: "The Azure app is missing a Microsoft Graph permission or admin consent "
	"(Sites.ReadWrite.All, Group.ReadWrite.All, User.Read.All).",
	404: "The site, document library or base path was not found. Check the Folder Rule.",
	429: "Microsoft is busy (throttling). Click Retry in a few minutes.",
	503: "Microsoft is temporarily unavailable. Click Retry in a few minutes.",
}


def mark_failed(record, error: Exception, step: str, owners: list[str]) -> None:
	"""
	Save the failure with a plain-language reason, log the traceback and
	email the configured recipients.

	Parameters:
	        record (Document, required): SharePoint Folder being provisioned.
	        error (Exception, required): The exception raised during provisioning.
	        step (str, required): One of the STEP_* constants.
	        owners (list[str], required): Owner emails of the folder.

	Returns:
	        None
	"""
	frappe.log_error("SharePoint | Folder provisioning failed", frappe.get_traceback())
	record.db_set(
		{
			"status": "Failed",
			"error": failure_reason(error, step, owners)[:5000],
			"attempts": record.attempts,
		}
	)
	notify_failure(record)


def failure_reason(error: Exception, step: str, owners: list[str] | None = None) -> str:
	"""
	Explain a failed provisioning run: which step failed, the likely cause and
	what to do, followed by Microsoft's own message.
	e.g. "Giving the owners access failed. One of the owners (a@x.com) is not a
	user in the Microsoft 365 tenant. ... Microsoft said: Graph 400: ..."

	Parameters:
	        error (Exception, required): Error raised during provisioning.
	        step (str, required): One of the STEP_* constants.
	        owners (list[str], optional): Owner emails of the folder.

	Returns:
	        str: Reason to store in the SharePoint Folder's Error field.
	"""
	hint = __hint(error, step, owners or [])
	reason = _("{0} failed.").format(_(step))
	if hint:
		reason = f"{reason} {hint}"
	label = _("Microsoft said") if isinstance(error, GraphError) else _("Details")
	return f"{reason}\n\n{label}: {error}"


def __hint(error: Exception, step: str, owners: list[str]) -> str:
	"""
	Likely cause and fix for the error at this step ("" when unknown).

	Parameters:
	        error (Exception, required): Error raised during provisioning.
	        step (str, required): One of the STEP_* constants.
	        owners (list[str], required): Owner emails of the folder.

	Returns:
	        str: Translated hint.
	"""
	if is_unresolved_recipient(error) and step == STEP_GROUP_ACCESS:
		return _(
			"SharePoint does not see the new security group yet; Microsoft can take a few "
			"minutes to copy a new group to SharePoint. The folder and group are already "
			"created, so click Retry in a few minutes."
		)
	if is_unresolved_recipient(error):
		return _(
			"One of the owners ({0}) is not a user in the Microsoft 365 tenant, so SharePoint "
			"cannot give them access. Use ERPNext users whose email exists in Microsoft 365."
		).format(", ".join(owners))
	if isinstance(error, GraphError) and error.status_code in STATUS_HINTS:
		return _(STATUS_HINTS[error.status_code])
	return ""
