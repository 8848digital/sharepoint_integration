# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

app_name = "sharepoint_integration"
app_title = "SharePoint Integration"
app_publisher = "8848 Digital LLP"
app_description = "SharePoint integration for Frappe/ERPNext"
app_email = "mahak@8848digital.com"
app_license = "Proprietary"

custom_fixtures = [{"dt": "Custom Field", "filters": {"module": "SharePoint Core"}}]

commands = ["sharepoint_integration.commands.export_fixtures.export_fixtures"]

after_request = [
	"sharepoint_integration.utils.api_handlers.response_formatter.format_frappe_response_to_custom"
]

app_include_js = "sharepoint_integration.bundle.js"

boot_session = "sharepoint_integration.boot_session.boot_session"

# Every save passes through these; they exit after a cached lookup unless the
# DocType has a SharePoint Folder Rule.
doc_events = {
	"*": {
		"after_insert": "sharepoint_integration.sharepoint_core.document_events.after_insert",
		"on_update": "sharepoint_integration.sharepoint_core.document_events.on_update",
	},
	"File": {
		"after_insert": "sharepoint_integration.sharepoint_core.customization.file.file.after_insert",
	},
}

scheduler_events = {
	"daily": ["sharepoint_integration.sharepoint_core.tasks.daily"],
	"hourly": ["sharepoint_integration.sharepoint_core.tasks.hourly"],
}
