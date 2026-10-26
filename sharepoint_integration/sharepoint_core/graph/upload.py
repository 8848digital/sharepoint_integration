# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Upload a local file into a SharePoint folder (simple PUT or chunked upload session)."""

import os
from urllib.parse import quote, unquote, urlencode, urlparse

from sharepoint_integration.sharepoint_core.graph.client import GraphClient
from sharepoint_integration.sharepoint_core.graph.common import encode_path

# Graph accepts a single PUT only up to 4 MB; bigger files need an upload session.
SIMPLE_UPLOAD_LIMIT = 4 * 1024 * 1024
# Session chunks must be a multiple of 320 KiB; 32 x 320 KiB = 10 MiB.
CHUNK_SIZE = 32 * 320 * 1024
# Two attachments with the same name never overwrite each other ("report 1.pdf").
CONFLICT = {"@microsoft.graph.conflictBehavior": "rename"}


def upload_file(
	client: GraphClient, drive_id: str, parent_id: str, file_name: str, file_path: str
) -> dict:
	"""
	Upload a file from disk into a folder, choosing the upload method by size.

	Parameters:
	        client (GraphClient, required): Authenticated Graph client.
	        drive_id (str, required): Drive (document library) id.
	        parent_id (str, required): driveItem id of the target folder.
	        file_name (str, required): Name to give the file in SharePoint.
	        file_path (str, required): Full path of the file on this server.

	Returns:
	        dict: Graph driveItem of the uploaded file (id, name, webUrl, size).
	"""
	size = os.path.getsize(file_path)
	target = f"/drives/{drive_id}/items/{parent_id}:/{encode_path(file_name)}:"

	if size <= SIMPLE_UPLOAD_LIMIT:
		# file_path is File.get_full_path() (inside the site's files folder), not user input.
		with open(file_path, "rb") as handle:  # nosemgrep
			return client.request("PUT", f"{target}/content", data=handle.read(), params=CONFLICT).json()

	session = client.post(f"{target}/createUploadSession", {"item": CONFLICT}).json()
	return __upload_chunks(client, session["uploadUrl"], file_path, size)


def preview_url(client: GraphClient, drive_id: str, file_url: str) -> str:
	"""
	Link that opens a file in SharePoint's viewer inside its library, instead of
	the direct file address (which makes the browser download images, zips, etc.).
	e.g. .../sites/X/Shared%20Documents/A/b.png ->
	     .../sites/X/Shared%20Documents/Forms/AllItems.aspx?id=/sites/X/Shared Documents/A/b.png&parent=...

	Parameters:
	        client (GraphClient, required): Graph client.
	        drive_id (str, required): Drive (document library) id.
	        file_url (str, required): The file's webUrl from Graph.

	Returns:
	        str: Preview link.
	"""
	library = client.get(f"/drives/{drive_id}", params={"$select": "webUrl"})["webUrl"].rstrip("/")
	file_path = unquote(urlparse(file_url).path)
	query = urlencode({"id": file_path, "parent": file_path.rsplit("/", 1)[0]}, quote_via=quote)
	return f"{library}/Forms/AllItems.aspx?{query}"


def __upload_chunks(client: GraphClient, upload_url: str, file_path: str, size: int) -> dict:
	"""
	Send the file in 10 MiB pieces to an upload session; the last piece returns the item.

	Parameters:
	        client (GraphClient, required): Graph client.
	        upload_url (str, required): Pre-authenticated session URL.
	        file_path (str, required): Full path of the file.
	        size (int, required): File size in bytes.

	Returns:
	        dict: Graph driveItem of the uploaded file.
	"""
	start = 0
	# file_path is File.get_full_path() (inside the site's files folder), not user input.
	with open(file_path, "rb") as handle:  # nosemgrep
		while start < size:
			chunk = handle.read(CHUNK_SIZE)
			end = start + len(chunk) - 1
			response = client.request(
				"PUT",
				upload_url,
				data=chunk,
				headers={"Content-Range": f"bytes {start}-{end}/{size}"},
				auth=False,
			)
			start = end + 1
	return response.json()
