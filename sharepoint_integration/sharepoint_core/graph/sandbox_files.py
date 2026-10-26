# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Sandbox handlers for file uploads (simple PUT and chunked upload sessions)."""

import re

import frappe

from sharepoint_integration.sharepoint_core.graph.sandbox_support import (
	SandboxResponse,
	_item_by_id,
)

CONTENT_RANGE = re.compile(r"bytes (\d+)-(\d+)/(\d+)")


def upload(base_site: str, state: dict, args: dict, kwargs: dict) -> SandboxResponse:
	"""
	PUT .../items/{parent}:/{name}:/content - store the file (size only).

	Parameters:
	        base_site (str, required): Fake tenant URL for webUrl.
	        state (dict, required): Sandbox state.
	        args (dict, required): Route groups (parent, name).
	        kwargs (dict, required): Request kwargs (`data`).

	Returns:
	        SandboxResponse: 201 with the new driveItem.
	"""
	item = __add_file(base_site, state, args["parent"], args["name"], len(kwargs.get("data") or b""))
	return SandboxResponse(201, item)


def upload_session(base_site: str, state: dict, args: dict, kwargs: dict) -> SandboxResponse:
	"""
	POST .../items/{parent}:/{name}:/createUploadSession - open a session.

	Parameters:
	        base_site (str, required): Fake tenant URL for the upload URL.
	        state (dict, required): Sandbox state.
	        args (dict, required): Route groups (parent, name).
	        kwargs (dict, required): Request kwargs.

	Returns:
	        SandboxResponse: 200 with uploadUrl.
	"""
	_item_by_id(state, args["parent"])
	session = frappe.generate_hash(length=12)
	state.setdefault("sessions", {})[session] = {"parent": args["parent"], "name": args["name"]}
	return SandboxResponse(200, {"uploadUrl": f"{base_site}/_sandbox/upload/{session}"})


def upload_chunk(base_site: str, state: dict, args: dict, kwargs: dict) -> SandboxResponse:
	"""
	PUT {uploadUrl} - accept one chunk; the last one creates the file.
	Rejects a bearer token, as Microsoft does on pre-authenticated URLs.

	Parameters:
	        base_site (str, required): Fake tenant URL for webUrl.
	        state (dict, required): Sandbox state.
	        args (dict, required): Route groups (session).
	        kwargs (dict, required): Request kwargs (`headers`).

	Returns:
	        SandboxResponse: 202 with the next range, or 201 with the driveItem.
	"""
	headers = kwargs.get("headers") or {}
	if headers.get("Authorization"):
		return SandboxResponse(401, {"error": {"message": "Unexpected Authorization header"}})

	session = state.get("sessions", {}).get(args["session"])
	match = CONTENT_RANGE.fullmatch(headers.get("Content-Range", ""))
	if not session or not match:
		return SandboxResponse(404, {"error": {"message": "itemNotFound"}})

	end, total = int(match.group(2)), int(match.group(3))
	if end + 1 < total:
		return SandboxResponse(202, {"nextExpectedRanges": [f"{end + 1}-"]})

	del state["sessions"][args["session"]]
	return SandboxResponse(
		201, __add_file(base_site, state, session["parent"], session["name"], total)
	)


def __add_file(base_site: str, state: dict, parent_id: str, name: str, size: int) -> dict:
	"""
	Add a file under a folder, renaming like Graph on a name clash
	(e.g. "report.pdf" -> "report 1.pdf").

	Parameters:
	        base_site (str, required): Fake tenant URL.
	        state (dict, required): Sandbox state.
	        parent_id (str, required): Folder item id.
	        name (str, required): Requested file name.
	        size (int, required): File size in bytes.

	Returns:
	        dict: The stored driveItem.
	"""
	parent_path = _item_by_id(state, parent_id)["path"]
	stem, dot, extension = name.rpartition(".")
	stem, extension = (stem, f".{extension}") if dot else (name, "")

	final_name, copy_number = name, 0
	while f"{parent_path}/{final_name}" in state["items"]:
		copy_number += 1
		final_name = f"{stem} {copy_number}{extension}"

	path = f"{parent_path}/{final_name}"
	item = {
		"id": f"item-{len(state['items']) + 1}",
		"name": final_name,
		"path": path,
		"size": size,
		"file": {},
		"permissions": [],
		"webUrl": f"{base_site}/Shared%20Documents/{path}",
	}
	state["items"][path] = item
	return item
