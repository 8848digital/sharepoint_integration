# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

import frappe

from sharepoint_integration.sharepoint_core.services.file_request import (
	retry_upload as queue_retry,
)
from sharepoint_integration.utils.api_handlers.response_formatter import api_response


@frappe.whitelist(methods=["POST"])
def retry_upload(sharepoint_file: str):
	"""
	        Re-queue a failed attachment upload.

	        **Endpoint:** `/api/method/sharepoint_integration.sharepoint_core.api.v1.file.retry_upload`
	        **HTTP Method:** POST
	        **Parameters:**
	                - sharepoint_file (str, required): SharePoint File name, e.g. "SPFILE-00001"
	        **Response:**
	```json
	                {"status": true, "status_code": 200, "message": "Upload queued", "data": null, "errors": null}
	```
	"""
	queue_retry(sharepoint_file)
	return api_response(message="Upload queued")
