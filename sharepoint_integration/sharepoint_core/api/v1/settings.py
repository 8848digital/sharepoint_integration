# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

import frappe

from sharepoint_integration.sharepoint_core.services.connection import (
	refresh_rule_location,
	test_connection,
)
from sharepoint_integration.utils.api_handlers.response_formatter import api_response


@frappe.whitelist(methods=["POST"])
def check_connection(connection: str):
	"""
	        Authenticate with a SharePoint Connection's credentials and read its tenant
	        root site. System Manager only.

	        **Endpoint:** `/api/method/sharepoint_integration.sharepoint_core.api.v1.settings.check_connection`
	        **HTTP Method:** POST
	        **Parameters:**
	                - connection (str, required): SharePoint Connection name
	        **Response:**
	```json
	                {
	                        "status": true,
	                        "status_code": 200,
	                        "message": "Connected to SharePoint",
	                        "data": {"connection": "Contoso", "site_id": "contoso.sharepoint.com,...", "web_url": "https://contoso.sharepoint.com", "sandbox": false},
	                        "errors": null
	                }
	```
	"""
	return api_response(message="Connected to SharePoint", data=test_connection(connection))


@frappe.whitelist(methods=["POST"])
def resolve_rule_site(rule: str):
	"""
	        Resolve and store the site id and document library (drive) id of a Folder Rule.
	        System Manager only.

	        **Endpoint:** `/api/method/sharepoint_integration.sharepoint_core.api.v1.settings.resolve_rule_site`
	        **HTTP Method:** POST
	        **Parameters:**
	                - rule (str, required): SharePoint Folder Rule name
	        **Response:**
	```json
	                {
	                        "status": true,
	                        "status_code": 200,
	                        "message": "Site resolved",
	                        "data": {"site_id": "...", "drive_id": "b!..."},
	                        "errors": null
	                }
	```
	"""
	return api_response(message="Site resolved", data=refresh_rule_location(rule))
