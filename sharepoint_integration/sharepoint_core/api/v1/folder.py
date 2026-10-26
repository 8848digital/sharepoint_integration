# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

import frappe

from sharepoint_integration.sharepoint_core.services.folder_query import get_folder_details
from sharepoint_integration.sharepoint_core.services.folder_request import (
	request_folder_manually,
	retry_folder,
)
from sharepoint_integration.utils.api_handlers.response_formatter import api_response


@frappe.whitelist(methods=["POST"])
def create_folder(doctype: str, docname: str):
	"""
	        Create (or retry) the SharePoint folder of a document using its matching
	        Folder Rule. Creation runs in the background; poll `get_folder` for status.

	        **Endpoint:** `/api/method/sharepoint_integration.sharepoint_core.api.v1.folder.create_folder`
	        **HTTP Method:** POST
	        **Parameters:**
	                - doctype (str, required): Reference DocType, e.g. "Opportunity"
	                - docname (str, required): Reference document name
	        **Response:**
	```json
	                {
	                        "status": true,
	                        "status_code": 200,
	                        "message": "SharePoint folder queued",
	                        "data": {"folder": "SPF-00001", "status": "Queued"},
	                        "errors": null
	                }
	```
	"""
	return api_response(
		message="SharePoint folder queued", data=request_folder_manually(doctype, docname)
	)


@frappe.whitelist(methods=["GET"])
def get_folder(doctype: str, docname: str):
	"""
	        Return the SharePoint folder status and link of a document (used by the
	        "Open in SharePoint" button and by search pages).

	        **Endpoint:** `/api/method/sharepoint_integration.sharepoint_core.api.v1.folder.get_folder`
	        **HTTP Method:** GET
	        **Parameters:**
	                - doctype (str, required): Reference DocType
	                - docname (str, required): Reference document name
	        **Response:**
	```json
	                {
	                        "status": true,
	                        "status_code": 200,
	                        "message": "Success",
	                        "data": {"folder": "SPF-00001", "status": "Created", "web_url": "https://...", "folder_name": "ACME-CRM-OPP-0001"},
	                        "errors": null
	                }
	```
	"""
	return api_response(data=get_folder_details(doctype, docname))


@frappe.whitelist(methods=["POST"])
def retry(folder: str):
	"""
	        Re-queue a failed SharePoint Folder.

	        **Endpoint:** `/api/method/sharepoint_integration.sharepoint_core.api.v1.folder.retry`
	        **HTTP Method:** POST
	        **Parameters:**
	                - folder (str, required): SharePoint Folder name
	        **Response:**
	```json
	                {"status": true, "status_code": 200, "message": "SharePoint folder queued", "data": null, "errors": null}
	```
	"""
	retry_folder(folder)
	return api_response(message="SharePoint folder queued")
