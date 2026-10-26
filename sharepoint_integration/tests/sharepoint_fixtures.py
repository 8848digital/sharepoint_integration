# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Shared sandbox setup for SharePoint tests (two connections, users, ToDo rules)."""

import frappe

from sharepoint_integration.sharepoint_core.graph.sandbox import STATE_KEY

RULE = "Test ToDo Rule"
RULE_B = "Test ToDo Rule (Tenant B)"
CONNECTIONS = {
	"Test Tenant A": "https://tenant-a.sharepoint.com",
	"Test Tenant B": "https://tenant-b.sharepoint.com",
}
USERS = ("sp.owner@example.com", "sp.member@example.com")


def setup_sandbox() -> None:
	"""
	Create two connections (two tenants), enable sandbox settings, create users
	and one ToDo rule per connection. Uses the core ToDo DocType on purpose:
	nothing in the app is tied to ERPNext or specific DocTypes.

	Returns:
	        None
	"""
	frappe.cache.delete_value(STATE_KEY)
	for name, base_site in CONNECTIONS.items():
		if not frappe.db.exists("SharePoint Connection", name):
			frappe.get_doc(
				{
					"doctype": "SharePoint Connection",
					"connection_name": name,
					"base_site": base_site,
					"tenant_id": "11111111-1111-1111-1111-111111111111",
					"client_id": "22222222-2222-2222-2222-222222222222",
					"client_secret": "test-secret",
				}
			).insert()
	settings = frappe.get_single("SharePoint Settings")
	settings.update({"enabled": 1, "sandbox_mode": 1, "default_connection": "Test Tenant A"})
	settings.save()
	for email in USERS:
		if not frappe.db.exists("User", email):
			frappe.get_doc({"doctype": "User", "email": email, "first_name": email.split("@")[0]}).insert()
	for rule in (RULE, RULE_B):
		if frappe.db.exists("SharePoint Folder Rule", rule):
			frappe.delete_doc("SharePoint Folder Rule", rule, force=1)
	frappe.get_doc(
		{
			"doctype": "SharePoint Folder Rule",
			"rule_name": RULE,
			"connection": "Test Tenant A",
			"reference_doctype": "ToDo",
			"condition": 'doc.priority == "High"',
			"priority": 99,  # win over any ToDo rule already on the site
			"base_path": "Tasks",
			"folder_name_template": "{{ doc.description }}-{{ doc.name }}",
			"owner_fields": [{"fieldname": "allocated_to"}],
			"sub_folders": [
				{"folder_name": "Shared", "grant_group_access": 1},
				{"folder_name": "Private", "grant_group_access": 0},
			],
		}
	).insert()
	frappe.get_doc(
		{
			"doctype": "SharePoint Folder Rule",
			"rule_name": RULE_B,
			"connection": "Test Tenant B",
			"reference_doctype": "ToDo",
			"condition": 'doc.priority == "Medium"',
			"priority": 99,
			"folder_name_template": "{{ doc.name }}",
		}
	).insert()


def make_todo(priority: str):
	"""
	Insert a ToDo allocated to the owner user.

	Parameters:
	        priority (str, required): "High" -> tenant A rule, "Medium" -> tenant B rule, else no rule.

	Returns:
	        Document: The ToDo.
	"""
	return frappe.get_doc(
		{
			"doctype": "ToDo",
			"description": "Quarterly report",
			"priority": priority,
			"allocated_to": USERS[0],
		}
	).insert()


def member_statuses(access) -> list[str]:
	"""
	Current access statuses of the members, in table order.

	Parameters:
	        access (Document, required): SharePoint Access.

	Returns:
	        list[str]: Statuses.
	"""
	access.reload()
	return [row.access_status for row in access.members]
