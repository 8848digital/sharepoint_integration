# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Connectivity checks and site/library resolution for rules."""

import frappe

from sharepoint_integration.sharepoint_core.graph.client import GraphClient
from sharepoint_integration.sharepoint_core.graph.directory import DirectoryService


def test_connection(connection: str) -> dict:
	"""
	Authenticate with a connection's credentials and read its tenant root site.

	Parameters:
	        connection (str, required): SharePoint Connection name.

	Returns:
	        dict: {"connection", "site_id", "web_url", "sandbox"}
	"""
	frappe.only_for("System Manager")
	base_site = frappe.db.get_value("SharePoint Connection", connection, "base_site")
	site_id = DirectoryService(GraphClient.from_connection(connection)).resolve_site_id(
		base_site, None
	)
	return {
		"connection": connection,
		"site_id": site_id,
		"web_url": base_site,
		"sandbox": bool(frappe.db.get_single_value("SharePoint Settings", "sandbox_mode")),
	}


def refresh_rule_location(rule_name: str) -> dict:
	"""
	Re-resolve and store a rule's site id and drive id (the "Resolve Site" button).

	Parameters:
	        rule_name (str, required): SharePoint Folder Rule name.

	Returns:
	        dict: {"site_id", "drive_id"}
	"""
	frappe.only_for("System Manager")
	rule = frappe.get_doc("SharePoint Folder Rule", rule_name)
	rule.db_set({"site_id": None, "drive_id": None})
	drive_id = resolve_rule_location(rule, GraphClient.from_connection(rule.connection))
	return {"site_id": rule.site_id, "drive_id": drive_id}


def resolve_rule_location(rule, client: GraphClient) -> str:
	"""
	Return the rule's drive id, resolving and caching site/drive ids on the
	rule the first time (saves two Graph calls on every later folder).

	Parameters:
	        rule (Document, required): SharePoint Folder Rule.
	        client (GraphClient, required): Client of the rule's connection.

	Returns:
	        str: Graph drive id.
	"""
	if rule.drive_id:
		return rule.drive_id

	base_site = frappe.db.get_value("SharePoint Connection", rule.connection, "base_site")
	directory = DirectoryService(client)
	site_id = directory.resolve_site_id(base_site, rule.site_name)
	drive_id = directory.resolve_drive_id(site_id, rule.library_name)
	rule.db_set({"site_id": site_id, "drive_id": drive_id})
	return drive_id
