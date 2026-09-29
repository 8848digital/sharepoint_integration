# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Microsoft Graph HTTP client: app-only token (cached in Redis) + retrying requests."""

import time

import requests

import frappe
from frappe import _

from sharepoint_integration.sharepoint_core.graph.common import (
	GraphError,
	error_message,
	safe_json,
)

GRAPH_ROOT = "https://graph.microsoft.com/v1.0"
LOGIN_URL = "https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"
RETRY_STATUSES = (429, 503, 504)
MAX_ATTEMPTS = 3
TIMEOUT_SECONDS = 30
TOKEN_SAFETY_SECONDS = 300


class GraphClient:
	"""
	Minimal Microsoft Graph client used by every SharePoint service.
	One token per (tenant, client_id) is cached in Redis, so a job that makes
	many calls authenticates once instead of once per call.
	"""

	def __init__(self, credentials: dict, transport=None):
		"""
		Create a client for one Entra app registration.

		Parameters:
		        credentials (dict, required): {"tenant_id", "client_id", "client_secret"}.
		        transport (object, optional): Object exposing `request(method, url, **kw)`;
		                defaults to the `requests` module. The sandbox passes a fake here.

		Returns:
		        None
		"""
		self.credentials = credentials
		self.transport = transport or requests

	@classmethod
	def from_connection(cls, connection: str) -> "GraphClient":
		"""
		Build a client for one SharePoint Connection (tenant / app registration).

		Parameters:
		        connection (str, required): SharePoint Connection name.

		Returns:
		        GraphClient: Ready-to-use client (sandboxed if Sandbox Mode is on).
		"""
		settings = frappe.get_cached_doc("SharePoint Settings")
		if not settings.enabled:
			frappe.throw(_("SharePoint Integration is disabled in SharePoint Settings"))
		if not connection:
			frappe.throw(_("No SharePoint Connection given"))

		conn = frappe.get_cached_doc("SharePoint Connection", connection)
		if not conn.enabled:
			frappe.throw(_("SharePoint Connection {0} is disabled").format(connection))

		credentials = {
			"tenant_id": conn.tenant_id,
			"client_id": conn.client_id,
			"client_secret": conn.get_password("client_secret", raise_exception=False),
		}
		if not settings.sandbox_mode:
			return cls(credentials)

		from sharepoint_integration.sharepoint_core.graph.sandbox import SandboxTransport

		credentials = {key: value or "sandbox" for key, value in credentials.items()}
		return cls(credentials, SandboxTransport(conn.base_site))

	def get(self, path: str, params: dict | None = None) -> dict:
		"""
		GET a Graph resource and return its JSON body.

		Parameters:
		        path (str, required): Path relative to the Graph v1.0 root, e.g. "/sites/root".
		        params (dict, optional): Query string parameters.

		Returns:
		        dict: Parsed JSON response.
		"""
		return self.request("GET", path, params=params).json()

	def post(self, path: str, body: dict | None = None):
		"""
		POST JSON to a Graph resource.

		Parameters:
		        path (str, required): Path relative to the Graph v1.0 root.
		        body (dict, optional): JSON payload.

		Returns:
		        Response: Raw response (callers need headers for async copy).
		"""
		return self.request("POST", path, body=body)

	def delete(self, path: str):
		"""
		DELETE a Graph resource.

		Parameters:
		        path (str, required): Path relative to the Graph v1.0 root.

		Returns:
		        Response: Raw response.
		"""
		return self.request("DELETE", path)

	def request(self, method: str, path: str, body: dict | None = None, params: dict | None = None):
		"""
		Send an authenticated request, retrying on throttling (429/503/504).

		Parameters:
		        method (str, required): HTTP verb.
		        path (str, required): Path relative to the Graph v1.0 root.
		        body (dict, optional): JSON payload.
		        params (dict, optional): Query string parameters.

		Returns:
		        Response: The successful response.

		Raises:
		        GraphError: On any non-2xx response after retries.
		"""
		url = f"{GRAPH_ROOT}{path}"
		for attempt in range(1, MAX_ATTEMPTS + 1):
			response = self.transport.request(
				method, url, headers=self.__headers(), json=body, params=params, timeout=TIMEOUT_SECONDS
			)
			if response.status_code not in RETRY_STATUSES or attempt == MAX_ATTEMPTS:
				break
			time.sleep(int(response.headers.get("Retry-After") or 2 * attempt))

		if response.status_code >= 400:
			raise GraphError(response.status_code, error_message(response), safe_json(response))
		return response

	def get_token(self) -> str:
		"""
		Return a cached app-only token, fetching a new one when it is about to expire.

		Parameters:
		        None

		Returns:
		        str: Bearer access token for https://graph.microsoft.com.
		"""
		cache_key = (
			f"sharepoint_integration:token:{self.credentials['tenant_id']}:{self.credentials['client_id']}"
		)
		token = frappe.cache.get_value(cache_key)
		if token:
			return token

		token, expires_in = self.__fetch_token()
		frappe.cache.set_value(
			cache_key, token, expires_in_sec=max(expires_in - TOKEN_SAFETY_SECONDS, 60)
		)
		return token

	def __fetch_token(self) -> tuple[str, int]:
		"""
		Run the OAuth2 client-credentials flow against Microsoft Entra ID.

		Parameters:
		        None

		Returns:
		        tuple[str, int]: (access_token, expires_in seconds)
		"""
		missing = [key for key, value in self.credentials.items() if not value]
		if missing:
			frappe.throw(_("SharePoint credentials missing: {0}").format(", ".join(missing)))

		response = self.transport.request(
			"POST",
			LOGIN_URL.format(tenant=self.credentials["tenant_id"]),
			data={
				"grant_type": "client_credentials",
				"client_id": self.credentials["client_id"],
				"client_secret": self.credentials["client_secret"],
				"scope": "https://graph.microsoft.com/.default",
			},
			timeout=TIMEOUT_SECONDS,
		)
		if response.status_code >= 400:
			raise GraphError(response.status_code, error_message(response), safe_json(response))

		payload = response.json()
		return payload["access_token"], int(payload.get("expires_in", 3600))

	def __headers(self) -> dict:
		"""
		Build the standard JSON + bearer headers.

		Parameters:
		        None

		Returns:
		        dict: Request headers.
		"""
		return {
			"Authorization": f"Bearer {self.get_token()}",
			"Accept": "application/json",
			"Content-Type": "application/json",
		}
