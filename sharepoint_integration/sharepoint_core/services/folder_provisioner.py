# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Background job body: create a document's SharePoint folder tree and grant access."""

import frappe

from sharepoint_integration.sharepoint_core.graph.client import GraphClient
from sharepoint_integration.sharepoint_core.graph.common import GraphError
from sharepoint_integration.sharepoint_core.graph.directory import DirectoryService
from sharepoint_integration.sharepoint_core.graph.drive import DriveService
from sharepoint_integration.sharepoint_core.services.connection import resolve_rule_location
from sharepoint_integration.sharepoint_core.services.folder_outcome import (
	link_reference,
	sync_pending_access,
)
from sharepoint_integration.sharepoint_core.services.folder_request import sanitize_folder_name
from sharepoint_integration.sharepoint_core.services.notifications import notify_failure


class FolderProvisioner:
	"""
	Creates <base path>/<folder>/<subfolders>, optional security group, owner and
	group grants, and the template copy. Every step is idempotent, so a failed
	run can simply be retried.
	"""

	def __init__(self, folder: str):
		"""
		Load the SharePoint Folder record and everything it depends on.

		Parameters:
		        folder (str, required): SharePoint Folder name.

		Returns:
		        None
		"""
		self.record = frappe.get_doc("SharePoint Folder", folder)
		self.rule = frappe.get_doc("SharePoint Folder Rule", self.record.rule)
		self.settings = frappe.get_cached_doc("SharePoint Settings")
		self.reference = frappe.get_doc(self.record.reference_doctype, self.record.reference_name)
		self.drive = None

	def run(self) -> None:
		"""
		Provision the folder and record the outcome on the SharePoint Folder.

		Parameters:
		        None

		Returns:
		        None
		"""
		if self.record.status == "Created":
			return

		self.record.attempts = (self.record.attempts or 0) + 1
		# Not created yet, so follow the rule's current connection (its cached
		# drive id belongs to that tenant). Created folders keep theirs.
		self.record.connection = self.rule.connection
		try:
			client = GraphClient.from_connection(self.record.connection)
			drive_id = resolve_rule_location(self.rule, client)
			self.drive = DriveService(client, drive_id)

			parent = self.drive.ensure_folder(self.rule.base_path, self.record.folder_name)
			group = self.__ensure_group(DirectoryService(client))
			owners = self.__owner_emails()

			self.__grant(parent["id"], owners, group, self.rule.grant_parent_folder_access)
			subfolders = self.__create_subfolders(owners, group)
			template_error = self.__copy_template(parent["id"], drive_id)
			self.__mark_created(parent, drive_id, subfolders, group, template_error)
		except Exception as error:
			self.__mark_failed(error)

	def __ensure_group(self, directory: DirectoryService) -> dict | None:
		"""
		Get or create the document's security group when the rule asks for one.

		Parameters:
		        directory (DirectoryService, required): Directory service.

		Returns:
		        dict | None: {"id", "displayName"} of the group.
		"""
		if not self.rule.create_security_group:
			return None
		doctype, name = self.record.reference_doctype, self.record.reference_name
		return directory.ensure_security_group(
			f"SharePoint - {doctype} {name}",
			f"Access to SharePoint folder '{self.record.folder_name}' of {doctype} {name}",
		)

	def __owner_emails(self) -> list[str]:
		"""
		Emails of enabled users held in the rule's owner fields.

		Parameters:
		        None

		Returns:
		        list[str]: Owner emails (deduplicated).
		"""
		users = {self.reference.get(row.fieldname) for row in self.rule.owner_fields} - {None, ""}
		if not users:
			return []
		return frappe.get_all("User", filters={"name": ["in", list(users)], "enabled": 1}, pluck="email")

	def __create_subfolders(self, owners: list[str], group: dict | None) -> list[str]:
		"""
		Create each configured subfolder and grant access on it.

		Parameters:
		        owners (list[str], required): Owner emails.
		        group (dict, optional): Security group.

		Returns:
		        list[str]: Names of the subfolders created.
		"""
		parent_path = self.__folder_path()
		created = []
		for row in self.rule.sub_folders:
			name = sanitize_folder_name(row.folder_name)
			item = self.drive.ensure_folder(parent_path, name)
			self.__grant(item["id"], owners, group, row.grant_group_access)
			created.append(name)
		return created

	def __grant(
		self, item_id: str, owners: list[str], group: dict | None, include_group: bool
	) -> None:
		"""
		Give owners write access and (optionally) the group its role, then
		strip grants copied from the parent when "Restrict Folder Access" is on.

		Parameters:
		        item_id (str, required): driveItem id.
		        owners (list[str], required): Owner emails.
		        group (dict, optional): Security group.
		        include_group (bool, required): Whether the group gets access here.

		Returns:
		        None
		"""
		self.drive.invite(item_id, [{"email": email} for email in owners], "write")
		allowed = {email.lower() for email in owners}

		if group and include_group:
			self.drive.invite(item_id, [{"objectId": group["id"]}], self.settings.group_access_role)
			allowed.add(group["id"].lower())

		if self.settings.restrict_folder_access and allowed:
			self.drive.restrict_to(item_id, allowed, bool(self.settings.keep_site_owners))

	def __copy_template(self, folder_id: str, drive_id: str) -> str | None:
		"""
		Copy the rule's template file into the folder. A failed copy does not fail
		the folder; the error is kept on the record instead.

		Parameters:
		        folder_id (str, required): Target folder driveItem id.
		        drive_id (str, required): Folder's drive id.

		Returns:
		        str | None: Error message if the copy failed.
		"""
		if not self.rule.attach_template:
			return None
		try:
			self.drive.copy_item(
				self.rule.template_drive_id or drive_id,
				self.rule.template_item_id,
				folder_id,
				self.rule.template_file_name,
			)
		except GraphError as error:
			frappe.log_error("SharePoint | Template copy failed", frappe.get_traceback())
			return f"Template copy failed: {error}"
		return None

	def __mark_created(
		self, parent: dict, drive_id: str, subfolders: list[str], group, template_error
	) -> None:
		"""
		Save success details and link the folder back onto the reference document.

		Parameters:
		        parent (dict, required): Folder driveItem.
		        drive_id (str, required): Drive id.
		        subfolders (list[str], required): Subfolder names created.
		        group (dict, optional): Security group.
		        template_error (str, optional): Template copy error.

		Returns:
		        None
		"""
		self.record.update(
			{
				"status": "Created",
				"folder_id": parent["id"],
				"web_url": parent.get("webUrl"),
				"folder_path": self.__folder_path(),
				"drive_id": drive_id,
				"subfolders": "\n".join(subfolders),
				"group_id": group["id"] if group else None,
				"group_name": group.get("displayName") if group else None,
				"template_copied": int(bool(self.rule.attach_template and not template_error)),
				"error": template_error,
			}
		)
		self.record.save(ignore_permissions=True)
		link_reference(self.record)
		sync_pending_access(self.record)

	def __mark_failed(self, error: Exception) -> None:
		"""
		Save the failure, log the traceback and email the configured recipients.

		Parameters:
		        error (Exception, required): The exception raised during provisioning.

		Returns:
		        None
		"""
		frappe.log_error("SharePoint | Folder provisioning failed", frappe.get_traceback())
		self.record.db_set(
			{"status": "Failed", "error": str(error)[:5000], "attempts": self.record.attempts}
		)
		notify_failure(self.record)

	def __folder_path(self) -> str:
		"""
		Library-relative path of the main folder.

		Parameters:
		        None

		Returns:
		        str: e.g. "General/Proposals/ACME-OPP-0001"
		"""
		return "/".join(
			part.strip("/") for part in [self.rule.base_path, self.record.folder_name] if part
		)
