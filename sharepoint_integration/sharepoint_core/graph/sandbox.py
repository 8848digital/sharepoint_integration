# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""
In-memory stand-in for Microsoft Graph, used when SharePoint Settings has
"Sandbox Mode" on. Lets QA/demo the full flow (folders, template copy, groups,
access) on a local site with no Microsoft tenant. State lives in Redis so it
survives across background jobs. Never enable on production.
"""

import re
from urllib.parse import unquote, urlparse

import frappe

from sharepoint_integration.sharepoint_core.graph.sandbox_support import (
	SandboxResponse,
	_group_by_id,
	_item_by_id,
	_permission,
)

STATE_KEY = "sharepoint_integration:sandbox_state"


class SandboxTransport:
	"""Routes Graph URLs to small handlers operating on a Redis-backed dict."""

	def __init__(self, base_site: str | None):
		"""
		Create the transport.

		Parameters:
		        base_site (str, optional): Tenant URL used to build fake webUrls.

		Returns:
		        None
		"""
		self.base_site = (base_site or "https://sandbox.sharepoint.com").rstrip("/")
		self.routes = [
			("POST", r"/oauth2/v2\.0/token$", self.__token),
			("GET", r"/sites/[^/]+:/sites/(?P<site>[^/]+)$", self.__site),
			("GET", r"/sites/(?P<site>[^/:]+)$", self.__site),
			("GET", r"/sites/(?P<site_id>[^/]+)/drive$", self.__drive),
			("GET", r"/sites/(?P<site_id>[^/]+)/drives$", self.__drives),
			("POST", r"/drives/(?P<drive>[^/]+)/root(:/(?P<path>.+):)?/children$", self.__create_folder),
			("GET", r"/drives/(?P<drive>[^/]+)/root:/(?P<path>.+)$", self.__get_by_path),
			("POST", r"/drives/[^/]+/items/(?P<item>[^/]+)/copy$", self.__copy),
			("POST", r"/drives/[^/]+/items/(?P<item>[^/]+)/invite$", self.__invite),
			("GET", r"/drives/[^/]+/items/(?P<item>[^/]+)/permissions$", self.__permissions),
			(
				"DELETE",
				r"/drives/[^/]+/items/(?P<item>[^/]+)/permissions/(?P<perm>[^/]+)$",
				self.__delete_perm,
			),
			("GET", r"/groups$", self.__find_group),
			("POST", r"/groups$", self.__create_group),
			("POST", r"/groups/(?P<group>[^/]+)/members/\$ref$", self.__add_member),
			("DELETE", r"/groups/(?P<group>[^/]+)/members/(?P<user>[^/]+)/\$ref$", self.__remove_member),
			("GET", r"/users/(?P<email>[^/]+)$", self.__user),
		]

	def request(self, method: str, url: str, **kwargs) -> SandboxResponse:
		"""
		Dispatch a request to the matching fake handler.

		Parameters:
		        method (str, required): HTTP verb.
		        url (str, required): Full URL.
		        **kwargs: `json`, `params`, `data` as passed to `requests.request`.

		Returns:
		        SandboxResponse: Fake response (404 if no route matches).
		"""
		path = unquote(urlparse(url).path).replace("/v1.0", "", 1)
		for verb, pattern, handler in self.routes:
			match = re.search(pattern, path)
			if verb == method and match:
				state = frappe.cache.get_value(STATE_KEY) or {"items": {}, "groups": {}}
				try:
					response = handler(state, match.groupdict(), kwargs)
				except StopIteration:
					# Unknown item/group id (e.g. sandbox state reset): answer like Graph.
					return SandboxResponse(404, {"error": {"message": "itemNotFound"}})
				frappe.cache.set_value(STATE_KEY, state)
				return response
		return SandboxResponse(404, {"error": {"message": f"Sandbox has no route for {method} {path}"}})

	def __token(self, state, args, kwargs):
		"""Issue a fake app-only token."""
		return SandboxResponse(200, {"access_token": "sandbox-token", "expires_in": 3600})

	def __site(self, state, args, kwargs):
		"""Resolve any site name to a deterministic fake site id."""
		return SandboxResponse(200, {"id": f"sandbox-site-{args['site']}", "webUrl": self.base_site})

	def __drive(self, state, args, kwargs):
		"""Return the default "Documents" library of a fake site."""
		return SandboxResponse(200, {"id": f"sandbox-drive-{args['site_id']}", "name": "Documents"})

	def __drives(self, state, args, kwargs):
		"""List a fake site's libraries."""
		return SandboxResponse(200, {"value": [self.__drive(state, args, kwargs).payload]})

	def __create_folder(self, state, args, kwargs):
		"""Create a folder; 409 if it exists (mirrors conflictBehavior=fail)."""
		body = kwargs.get("json") or {}
		path = "/".join(part for part in [args.get("path"), body.get("name")] if part)
		if path in state["items"]:
			return SandboxResponse(409, {"error": {"message": "nameAlreadyExists"}})
		item = {
			"id": f"item-{len(state['items']) + 1}",
			"name": body.get("name"),
			"path": path,
			"permissions": [],
		}
		item["webUrl"] = f"{self.base_site}/Shared%20Documents/{path}"
		state["items"][path] = item
		return SandboxResponse(201, item)

	def __get_by_path(self, state, args, kwargs):
		"""Fetch a folder by path."""
		item = state["items"].get(args["path"])
		return (
			SandboxResponse(200, item)
			if item
			else SandboxResponse(404, {"error": {"message": "itemNotFound"}})
		)

	def __copy(self, state, args, kwargs):
		"""Accept an async copy request."""
		return SandboxResponse(
			202, {}, {"Location": f"{self.base_site}/_sandbox/monitor/{args['item']}"}
		)

	def __invite(self, state, args, kwargs):
		"""Record one permission per recipient."""
		item = _item_by_id(state, args["item"])
		body = kwargs.get("json") or {}
		created = [
			_permission(recipient, body.get("roles", ["read"])) for recipient in body.get("recipients", [])
		]
		item["permissions"].extend(created)
		return SandboxResponse(200, {"value": created})

	def __permissions(self, state, args, kwargs):
		"""List an item's permissions (plus a fake site-owners grant)."""
		owners = {
			"id": "site-owners",
			"roles": ["owner"],
			"grantedToV2": {"siteGroup": {"id": "owners"}},
		}
		return SandboxResponse(
			200, {"value": [owners, *_item_by_id(state, args["item"])["permissions"]]}
		)

	def __delete_perm(self, state, args, kwargs):
		"""Delete a permission from an item."""
		item = _item_by_id(state, args["item"])
		item["permissions"] = [perm for perm in item["permissions"] if perm["id"] != args["perm"]]
		return SandboxResponse(204)

	def __find_group(self, state, args, kwargs):
		"""Filter groups by mailNickname."""
		nickname = (kwargs.get("params") or {}).get("$filter", "").split("'")[1:2]
		group = state["groups"].get(nickname[0]) if nickname else None
		return SandboxResponse(200, {"value": [group] if group else []})

	def __create_group(self, state, args, kwargs):
		"""Create a security group."""
		body = kwargs.get("json") or {}
		group = {
			"id": f"group-{body['mailNickname']}",
			"displayName": body["displayName"],
			"members": [],
		}
		state["groups"][body["mailNickname"]] = group
		return SandboxResponse(201, group)

	def __add_member(self, state, args, kwargs):
		"""Add a member; 400 if already present (mirrors Graph)."""
		group = _group_by_id(state, args["group"])
		user_id = (kwargs.get("json") or {})["@odata.id"].rsplit("/", 1)[-1]
		if user_id in group["members"]:
			return SandboxResponse(
				400, {"error": {"message": "One or more added object references already exist"}}
			)
		group["members"].append(user_id)
		return SandboxResponse(204)

	def __remove_member(self, state, args, kwargs):
		"""Remove a member; 404 if not present."""
		group = _group_by_id(state, args["group"])
		if args["user"] not in group["members"]:
			return SandboxResponse(404, {"error": {"message": "Member not found"}})
		group["members"].remove(args["user"])
		return SandboxResponse(204)

	def __user(self, state, args, kwargs):
		"""Resolve an email to a fake object id."""
		return SandboxResponse(200, {"id": f"user-{args['email'].lower()}"})
