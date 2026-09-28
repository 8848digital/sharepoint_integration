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

# Uncomment and point at the conventional locations as features are added
# (see CLAUDE.md "Key Conventions"):
#
# after_install = "sharepoint_integration.install.after_install"
#
# doc_events = {
# 	"<DocType>": {
# 		"on_update": "sharepoint_integration.sharepoint_core.customization.<doctype>.<doctype>.on_update"
# 	}
# }
#
# scheduler_events = {
# 	"hourly": ["sharepoint_integration.sharepoint_core.tasks.hourly"],
# }
#
# permission_query_conditions = {
# 	"<DocType>": "sharepoint_integration.sharepoint_core.permissions.get_permission_query_conditions",
# }
