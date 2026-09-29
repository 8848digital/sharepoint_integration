# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Shared Graph error type and path helper."""

from urllib.parse import quote


class GraphError(Exception):
	"""Raised when Microsoft Graph returns a non-success response."""

	def __init__(self, status_code: int, message: str, payload: dict | None = None):
		"""
		Build the error with the HTTP status and Graph's error message.

		Parameters:
		        status_code (int, required): HTTP status returned by Graph.
		        message (str, required): Human readable error message.
		        payload (dict, optional): Parsed error body, if any.

		Returns:
		        None
		"""
		super().__init__(f"Graph {status_code}: {message}")
		self.status_code = status_code
		self.payload = payload or {}


def encode_path(path: str) -> str:
	"""
	URL-encode a drive-relative path while keeping "/" separators.
	e.g. "General/Proposals/A&B #1" -> "General/Proposals/A%26B%20%231"

	Parameters:
	        path (str, required): Folder path relative to the drive root.

	Returns:
	        str: Encoded path without leading/trailing slashes.
	"""
	return quote(path.strip("/"), safe="/")


def error_message(response) -> str:
	"""
	Extract Graph's `error.message` (or the OAuth error description) from a failed response.

	Parameters:
	        response (Response, required): Failed HTTP response.

	Returns:
	        str: Error message.
	"""
	payload = safe_json(response)
	error = payload.get("error")
	if isinstance(error, dict):
		return error.get("message") or error.get("code") or response.text
	return payload.get("error_description") or response.text[:500]


def safe_json(response) -> dict:
	"""
	Parse a response body as JSON without raising.

	Parameters:
	        response (Response, required): HTTP response.

	Returns:
	        dict: Parsed body, or {} when the body is empty/not JSON.
	"""
	try:
		return response.json() or {}
	except ValueError:
		return {}
