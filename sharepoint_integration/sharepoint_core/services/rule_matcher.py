# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Decide which SharePoint Folder Rule (if any) applies to a document event."""

import frappe

DOCTYPES_CACHE_KEY = "sharepoint_integration:rule_doctypes"
RULES_CACHE_KEY = "sharepoint_integration:rules:{doctype}"
RULE_FIELDS = ["name", "trigger_event", "workflow_state", "from_workflow_state", "condition"]


def get_rule_doctypes() -> set[str]:
	"""
	DocTypes that have at least one enabled rule. Cached, because the "*"
	doc_events hook calls this on every save of every document on the site.

	Parameters:
	        None

	Returns:
	        set[str]: Reference DocType names.
	"""
	# During install/migrate DocTypes are synced before this app's tables exist.
	if frappe.flags.in_install or frappe.flags.in_migrate:
		return set()
	return set(frappe.cache.get_value(DOCTYPES_CACHE_KEY, generator=__load_rule_doctypes) or [])


def find_matching_rule(doc, event: str) -> str | None:
	"""
	Return the highest-priority enabled rule whose trigger and condition match.

	Parameters:
	        doc (Document, required): The reference document.
	        event (str, required): "insert", "update" or "manual".

	Returns:
	        str | None: SharePoint Folder Rule name.
	"""
	rules = frappe.cache.get_value(
		RULES_CACHE_KEY.format(doctype=doc.doctype), generator=lambda: __load_rules(doc.doctype)
	)
	for rule in rules or []:
		if __trigger_matches(rule, doc, event) and __condition_matches(rule, doc):
			return rule["name"]
	return None


def clear_rule_cache(doctype: str | None = None) -> None:
	"""
	Drop cached rule lookups after a rule is saved or deleted.

	Parameters:
	        doctype (str, optional): Reference DocType whose rule list changed.

	Returns:
	        None
	"""
	frappe.cache.delete_value(DOCTYPES_CACHE_KEY)
	if doctype:
		frappe.cache.delete_value(RULES_CACHE_KEY.format(doctype=doctype))


def __load_rule_doctypes() -> list[str]:
	"""
	Query the reference DocTypes of enabled rules.

	Parameters:
	        None

	Returns:
	        list[str]: Distinct DocType names.
	"""
	try:
		return frappe.get_all(
			"SharePoint Folder Rule", filters={"enabled": 1}, pluck="reference_doctype", distinct=True
		)
	except frappe.db.TableMissingError:
		return []


def __load_rules(doctype: str) -> list[dict]:
	"""
	Query enabled rules of one DocType, best first.

	Parameters:
	        doctype (str, required): Reference DocType.

	Returns:
	        list[dict]: Rule rows with trigger/condition fields.
	"""
	return frappe.get_all(
		"SharePoint Folder Rule",
		filters={"enabled": 1, "reference_doctype": doctype},
		fields=RULE_FIELDS,
		order_by="priority desc, creation asc",
	)


def __trigger_matches(rule: dict, doc, event: str) -> bool:
	"""
	Check the rule's trigger against the event.
	A manual request matches any rule; "Manual" rules match nothing else.

	Parameters:
	        rule (dict, required): Cached rule row.
	        doc (Document, required): The reference document.
	        event (str, required): "insert", "update" or "manual".

	Returns:
	        bool: True if the rule should fire.
	"""
	trigger = rule["trigger_event"]
	if event == "manual":
		return True
	if trigger == "On Create":
		return event == "insert"
	if trigger == "On Save":
		return True
	if trigger == "On Workflow State":
		return __workflow_reached(rule, doc)
	return False


def __workflow_reached(rule: dict, doc) -> bool:
	"""
	True when this save moved the document into the rule's workflow state
	(optionally from a specific previous state).

	Parameters:
	        rule (dict, required): Cached rule row.
	        doc (Document, required): The reference document.

	Returns:
	        bool: Whether the transition matches.
	"""
	if doc.get("workflow_state") != rule["workflow_state"]:
		return False

	before = doc.get_doc_before_save()
	previous_state = before.get("workflow_state") if before else None
	if previous_state == doc.get("workflow_state"):
		return False
	return not rule["from_workflow_state"] or previous_state == rule["from_workflow_state"]


def __condition_matches(rule: dict, doc) -> bool:
	"""
	Evaluate the rule's optional Python condition with `doc` in scope.

	Parameters:
	        rule (dict, required): Cached rule row.
	        doc (Document, required): The reference document.

	Returns:
	        bool: True if there is no condition or it evaluates truthy.
	"""
	if not (rule["condition"] or "").strip():
		return True
	try:
		return bool(frappe.safe_eval(rule["condition"], None, {"doc": frappe._dict(doc.as_dict())}))
	except Exception:
		frappe.log_error(f"SharePoint | Condition error in rule {rule['name']}", frappe.get_traceback())
		return False
