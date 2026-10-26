# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

import frappe

from sharepoint_integration.sharepoint_core.services.access_revoke import request_revoke
from sharepoint_integration.sharepoint_core.services.folder_query import queue_access_sync_now
from sharepoint_integration.utils.api_handlers.response_formatter import api_response


@frappe.whitelist(methods=["POST"])
def sync_now(access: str):
	"""
	        Queue a group-membership sync for a SharePoint Access (Pending/Failed members).

	        **Endpoint:** `/api/method/sharepoint_integration.sharepoint_core.api.v1.access.sync_now`
	        **HTTP Method:** POST
	        **Parameters:**
	                - access (str, required): SharePoint Access name
	        **Response:**
	```json
	                {"status": true, "status_code": 200, "message": "Access sync queued", "data": null, "errors": null}
	```
	"""
	queue_access_sync_now(access)
	return api_response(message="Access sync queued")


@frappe.whitelist(methods=["POST"])
def revoke(access: str, members: str | list, remarks: str | None = None):
	"""
	        Remove selected members' access to the folder now. Their rows are kept as
	        "Revoked" with Removed On and Remarks (audit trail). More than 20 members
	        are revoked in a background job.

	        **Endpoint:** `/api/method/sharepoint_integration.sharepoint_core.api.v1.access.revoke`
	        **HTTP Method:** POST
	        **Parameters:**
	                - access (str, required): SharePoint Access name
	                - members (list[str] | JSON str, required): SharePoint Access Member row names
	                - remarks (str, optional): Reason for revoking, saved on each row
	        **Response:**
	```json
	                {
	                        "status": true,
	                        "status_code": 200,
	                        "message": "Access revoked",
	                        "data": {"queued": false, "errors": []},
	                        "errors": null
	                }
	```
	"""
	result = request_revoke(access, frappe.parse_json(members), remarks)
	return api_response(
		message="Access revoke queued" if result["queued"] else "Access revoked", data=result
	)
