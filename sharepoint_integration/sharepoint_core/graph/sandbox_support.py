# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Fake response object and state helpers for the Graph sandbox."""

import frappe


class SandboxResponse:
	"""Just enough of `requests.Response` for GraphClient."""

	def __init__(self, status_code: int, payload: dict | None = None, headers: dict | None = None):
		"""
		Build a fake response.

		Parameters:
		        status_code (int, required): HTTP status.
		        payload (dict, optional): JSON body.
		        headers (dict, optional): Response headers.

		Returns:
		        None
		"""
		self.status_code = status_code
		self.payload = payload or {}
		self.headers = headers or {}
		self.text = frappe.as_json(self.payload)

	def json(self) -> dict:
		"""
		Return the JSON body.

		Parameters:
		        None

		Returns:
		        dict: Payload.
		"""
		return self.payload


def _item_by_id(state: dict, item_id: str) -> dict:
	"""
	Find a sandbox item by id (single underscore: called from the class body).

	Parameters:
	        state (dict, required): Sandbox state.
	        item_id (str, required): Item id.

	Returns:
	        dict: The item.
	"""
	return next(item for item in state["items"].values() if item["id"] == item_id)


def _group_by_id(state: dict, group_id: str) -> dict:
	"""
	Find a sandbox group by id (single underscore: called from the class body).

	Parameters:
	        state (dict, required): Sandbox state.
	        group_id (str, required): Group id.

	Returns:
	        dict: The group.
	"""
	return next(group for group in state["groups"].values() if group["id"] == group_id)


def _permission(recipient: dict, roles: list[str]) -> dict:
	"""
	Build a fake Graph permission for an invite recipient.

	Parameters:
	        recipient (dict, required): {"email": ...} or {"objectId": ...}.
	        roles (list[str], required): Granted roles.

	Returns:
	        dict: Graph-shaped permission.
	"""
	principal = (
		{"user": {"email": recipient["email"]}}
		if "email" in recipient
		else {"group": {"id": recipient["objectId"]}}
	)
	return {"id": frappe.generate_hash(length=10), "roles": roles, "grantedToV2": principal}
