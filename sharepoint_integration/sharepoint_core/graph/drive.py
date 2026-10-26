# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Folder, file-copy and permission operations on one SharePoint document library."""

import time

from sharepoint_integration.sharepoint_core.graph.client import GraphClient
from sharepoint_integration.sharepoint_core.graph.common import (
	GraphError,
	encode_path,
	is_unresolved_recipient,
)

WRITE_ROLES = ("write", "read")
# Seconds to wait between tries while a new group reaches SharePoint (about 2 minutes in total).
GROUP_WAITS = (10, 20, 30, 60)


class DriveService:
	"""Operations on a single SharePoint document library (a Graph "drive")."""

	def __init__(self, client: GraphClient, drive_id: str):
		"""
		Bind the service to one drive.

		Parameters:
		        client (GraphClient, required): Authenticated Graph client.
		        drive_id (str, required): Graph drive id of the document library.

		Returns:
		        None
		"""
		self.client = client
		self.drive_id = drive_id

	def ensure_folder(self, parent_path: str, name: str) -> dict:
		"""
		Create a folder, or return the existing one with the same name.
		Idempotent: re-running a failed job never creates "XYZ 1" duplicates.

		Parameters:
		        parent_path (str, required): Parent path relative to the library root ("" = root).
		        name (str, required): Folder name (already sanitized).

		Returns:
		        dict: Graph driveItem of the folder (id, name, webUrl, ...).
		"""
		parent = encode_path(parent_path or "")
		children = (
			f"/drives/{self.drive_id}/root:/{parent}:/children"
			if parent
			else f"/drives/{self.drive_id}/root/children"
		)
		body = {"name": name, "folder": {}, "@microsoft.graph.conflictBehavior": "fail"}

		try:
			return self.client.post(children, body).json()
		except GraphError as error:
			if error.status_code != 409:
				raise
		return self.get_item_by_path(f"{parent_path}/{name}")

	def get_item_by_path(self, path: str) -> dict:
		"""
		Fetch a driveItem by its library-relative path.

		Parameters:
		        path (str, required): e.g. "General/Proposals/XYZ".

		Returns:
		        dict: Graph driveItem.
		"""
		return self.client.get(f"/drives/{self.drive_id}/root:/{encode_path(path)}")

	def copy_item(self, source_drive_id: str, item_id: str, target_folder_id: str, name: str) -> str:
		"""
		Server-side copy of a template file into a folder (no download/upload,
		so large templates cost nothing). Graph runs the copy asynchronously.

		Parameters:
		        source_drive_id (str, required): Drive holding the template.
		        item_id (str, required): driveItem id of the template file.
		        target_folder_id (str, required): driveItem id of the destination folder.
		        name (str, required): File name to create.

		Returns:
		        str: Monitor URL of the async copy operation ("" if Graph returned none).
		"""
		body = {"parentReference": {"driveId": self.drive_id, "id": target_folder_id}, "name": name}
		response = self.client.post(f"/drives/{source_drive_id}/items/{item_id}/copy", body)
		return response.headers.get("Location", "")

	def invite(self, item_id: str, recipients: list[dict], role: str) -> list[dict]:
		"""
		Grant access on an item to many users/groups in one call. SharePoint
		breaks inheritance on the item as part of this, copying the parent's grants.

		Parameters:
		        item_id (str, required): driveItem id.
		        recipients (list[dict], required): [{"email": ...}] or [{"objectId": <group id>}].
		        role (str, required): "write" or "read".

		Returns:
		        list[dict]: Permissions created by Graph.
		"""
		if not recipients:
			return []
		body = {
			"recipients": recipients,
			"roles": [role if role in WRITE_ROLES else "read"],
			"requireSignIn": True,
			"sendInvitation": False,
		}
		return (
			self.client.post(f"/drives/{self.drive_id}/items/{item_id}/invite", body)
			.json()
			.get("value", [])
		)

	def invite_group(self, item_id: str, group_id: str, role: str) -> list[dict]:
		"""
		Grant a group access, waiting while SharePoint does not know the group yet.
		A group created in Entra takes from seconds to minutes to reach SharePoint,
		so an invite right after creating it fails with "could not be resolved".

		Parameters:
		        item_id (str, required): driveItem id.
		        group_id (str, required): Group object id.
		        role (str, required): "write" or "read".

		Returns:
		        list[dict]: Permissions created by Graph.
		"""
		for wait in GROUP_WAITS:
			try:
				return self.invite(item_id, [{"objectId": group_id}], role)
			except GraphError as error:
				if not is_unresolved_recipient(error):
					raise
			time.sleep(wait)
		return self.invite(item_id, [{"objectId": group_id}], role)

	def restrict_to(
		self, item_id: str, allowed_principals: set[str], keep_site_owners: bool = True
	) -> int:
		"""
		Remove every direct permission on the item that is not granted to an
		allowed principal. Run after `invite`, which has already broken inheritance.

		Parameters:
		        item_id (str, required): driveItem id.
		        allowed_principals (set[str], required): Lower-cased emails / object ids to keep.
		        keep_site_owners (bool, optional): Keep "owner" grants (site owners group).

		Returns:
		        int: Number of permissions removed.
		"""
		permissions = self.client.get(f"/drives/{self.drive_id}/items/{item_id}/permissions").get(
			"value", []
		)
		removable = [
			permission
			for permission in permissions
			if not permission.get("inheritedFrom")
			and not (keep_site_owners and "owner" in permission.get("roles", []))
			and not (_principal_keys(permission) & allowed_principals)
		]
		for permission in removable:
			self.client.delete(f"/drives/{self.drive_id}/items/{item_id}/permissions/{permission['id']}")
		return len(removable)


def _principal_keys(permission: dict) -> set[str]:
	"""
	Collect the lower-cased ids/emails a permission is granted to.
	Single underscore (not "__") because DriveService calls it from inside the
	class body, where a "__" name would be mangled.

	Parameters:
	        permission (dict, required): Graph permission resource.

	Returns:
	        set[str]: e.g. {"a@contoso.com", "<group-object-id>"}
	"""
	identities = [
		permission.get("grantedToV2") or {},
		*(permission.get("grantedToIdentitiesV2") or []),
	]
	principals = [
		value for identity in identities for value in identity.values() if isinstance(value, dict)
	]
	return {
		str(principal.get(key)).lower()
		for principal in principals
		for key in ("id", "email", "loginName")
		if principal.get(key)
	}
