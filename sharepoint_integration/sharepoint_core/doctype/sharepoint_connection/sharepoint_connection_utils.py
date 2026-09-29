# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Validation for SharePoint Connection (one Microsoft tenant / app registration)."""

import re
from urllib.parse import urlparse

import frappe
from frappe import _

GUID_PATTERN = re.compile(
	r"^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$", re.IGNORECASE
)
TENANT_DOMAIN_PATTERN = re.compile(r"^[a-z0-9-]+(\.[a-z0-9-]+)+$", re.IGNORECASE)


def validate_connection(doc) -> None:
	"""
	Validate URL/ID formats; the secret is required unless Sandbox Mode is on.

	Parameters:
	        doc (Document, required): SharePoint Connection.

	Returns:
	        None
	"""
	doc.base_site = normalize_base_site(doc.base_site)
	validate_guid(doc.client_id, _("Client ID"))
	__validate_tenant(doc.tenant_id)

	sandbox = frappe.db.get_single_value("SharePoint Settings", "sandbox_mode")
	if doc.enabled and not sandbox and not doc.client_secret:
		frappe.throw(_("Client Secret is required (only Sandbox Mode works without it)"))

	if not doc.enabled and frappe.db.exists(
		"SharePoint Folder Rule", {"connection": doc.name, "enabled": 1}
	):
		frappe.throw(_("Disable the Folder Rules using this connection before disabling it"))


def normalize_base_site(base_site: str | None) -> str | None:
	"""
	Validate and normalize the SharePoint base URL to "https://<host>".

	Parameters:
	        base_site (str, optional): URL entered by the user.

	Returns:
	        str | None: Normalized URL, e.g. "https://contoso.sharepoint.com".
	"""
	if not base_site:
		return base_site

	parsed = urlparse(base_site.strip())
	if parsed.scheme != "https" or not parsed.hostname:
		frappe.throw(_("SharePoint Base URL must look like https://contoso.sharepoint.com"))
	if parsed.path.strip("/"):
		frappe.throw(
			_("SharePoint Base URL must be the tenant root only; set the site on each Folder Rule")
		)
	return f"https://{parsed.hostname}"


def validate_guid(value: str | None, label: str) -> None:
	"""
	Throw unless the value is empty or a GUID (Entra application/client ids are GUIDs).

	Parameters:
	        value (str, optional): Value to check.
	        label (str, required): Field label for the error message.

	Returns:
	        None
	"""
	if value and not GUID_PATTERN.match(value.strip()):
		frappe.throw(_("{0} must be a GUID, e.g. 00000000-0000-0000-0000-000000000000").format(label))


def __validate_tenant(tenant_id: str | None) -> None:
	"""
	A tenant can be given as its GUID or as a verified domain (contoso.onmicrosoft.com).

	Parameters:
	        tenant_id (str, optional): Tenant id.

	Returns:
	        None
	"""
	if tenant_id and not (GUID_PATTERN.match(tenant_id) or TENANT_DOMAIN_PATTERN.match(tenant_id)):
		frappe.throw(_("Tenant ID must be a GUID or a domain like contoso.onmicrosoft.com"))
