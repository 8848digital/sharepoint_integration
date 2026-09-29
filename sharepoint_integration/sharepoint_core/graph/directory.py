# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Entra ID (Azure AD) lookups: sites, drives, users and security groups."""

import re
from urllib.parse import urlparse

from sharepoint_integration.sharepoint_core.graph.client import GraphClient
from sharepoint_integration.sharepoint_core.graph.common import GraphError

ALREADY_MEMBER_HINT = "already exist"


class DirectoryService:
	"""Resolves SharePoint sites/drives and manages per-document security groups."""

	def __init__(self, client: GraphClient):
		"""
		Bind the service to a Graph client.

		Parameters:
		        client (GraphClient, required): Authenticated Graph client.

		Returns:
		        None
		"""
		self.client = client
		self._user_ids: dict[str, str | None] = {}

	def resolve_site_id(self, base_site: str, site_name: str | None) -> str:
		"""
		Turn "https://contoso.sharepoint.com" + "Proposals" into a Graph site id.

		Parameters:
		        base_site (str, required): Tenant SharePoint URL.
		        site_name (str, optional): Site under /sites/; blank = root site.

		Returns:
		        str: Graph site id.
		"""
		hostname = urlparse(base_site).hostname or base_site
		path = f"/sites/{hostname}:/sites/{site_name.strip('/')}" if site_name else f"/sites/{hostname}"
		return self.client.get(path)["id"]

	def resolve_drive_id(self, site_id: str, library_name: str | None = None) -> str:
		"""
		Find the document library (drive) id of a site.

		Parameters:
		        site_id (str, required): Graph site id.
		        library_name (str, optional): Library display name; blank = default library.

		Returns:
		        str: Graph drive id.
		"""
		if not library_name:
			return self.client.get(f"/sites/{site_id}/drive")["id"]

		drives = self.client.get(f"/sites/{site_id}/drives").get("value", [])
		match = next((drive for drive in drives if drive.get("name") == library_name), None)
		if not match:
			raise GraphError(404, f"Document library '{library_name}' not found on site")
		return match["id"]

	def ensure_security_group(self, display_name: str, description: str) -> dict:
		"""
		Return the security group for a document, creating it on first use.
		Looked up by mailNickname so retries never create duplicate groups.

		Parameters:
		        display_name (str, required): Group display name.
		        description (str, required): Group description.

		Returns:
		        dict: {"id", "displayName"} of the group.
		"""
		nickname = _mail_nickname(display_name)
		existing = self.client.get("/groups", params={"$filter": f"mailNickname eq '{nickname}'"}).get(
			"value", []
		)
		if existing:
			return existing[0]

		body = {
			"displayName": display_name[:256],
			"description": description[:1024],
			"mailEnabled": False,
			"mailNickname": nickname,
			"securityEnabled": True,
			"groupTypes": [],
		}
		return self.client.post("/groups", body).json()

	def add_member(self, group_id: str, email: str) -> None:
		"""
		Add a user to a group; adding an existing member is treated as success.

		Parameters:
		        group_id (str, required): Group object id.
		        email (str, required): User principal name / email.

		Returns:
		        None
		"""
		user_id = self.get_user_id(email)
		body = {"@odata.id": f"https://graph.microsoft.com/v1.0/directoryObjects/{user_id}"}
		try:
			self.client.post(f"/groups/{group_id}/members/$ref", body)
		except GraphError as error:
			if not (error.status_code == 400 and ALREADY_MEMBER_HINT in str(error)):
				raise

	def remove_member(self, group_id: str, email: str) -> None:
		"""
		Remove a user from a group; a user who is already gone is treated as success.

		Parameters:
		        group_id (str, required): Group object id.
		        email (str, required): User principal name / email.

		Returns:
		        None
		"""
		try:
			user_id = self.get_user_id(email)
			self.client.delete(f"/groups/{group_id}/members/{user_id}/$ref")
		except GraphError as error:
			if error.status_code != 404:
				raise

	def get_user_id(self, email: str) -> str:
		"""
		Resolve an email/UPN to an Entra object id (memoized per job).

		Parameters:
		        email (str, required): User principal name / email.

		Returns:
		        str: User object id.
		"""
		key = email.lower()
		if key not in self._user_ids:
			self._user_ids[key] = self.client.get(f"/users/{email}", params={"$select": "id"})["id"]
		return self._user_ids[key]


def _mail_nickname(display_name: str) -> str:
	"""
	Build a valid, deterministic mailNickname (ASCII letters/digits, max 64).

	Parameters:
	        display_name (str, required): Group display name.

	Returns:
	        str: e.g. "SharePoint - PRJ-0001 Alpha" -> "sp-prj-0001alpha"
	"""
	slug = re.sub(r"[^a-z0-9-]", "", display_name.lower().replace("sharepoint", "sp"))
	return (slug or "sp-group")[:64]
