# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Shared helpers for SharePoint attachment and bulk-run tests: inline job runner and files."""

from unittest.mock import patch

import frappe

from sharepoint_integration.sharepoint_core.graph.sandbox import STATE_KEY
from sharepoint_integration.tests.sharepoint_fixtures import make_todo

JOB_OPTIONS = ("queue", "timeout", "job_id", "deduplicate", "enqueue_after_commit")


def collect_jobs(testcase) -> list:
	"""
	Collect enqueued jobs instead of sending them to the worker, and keep
	commits (the bulk run commits per batch) from saving test data.

	Parameters:
	        testcase (TestCase, required): Test registering the cleanups.

	Returns:
	        list: Collected (method, kwargs) pairs, filled as jobs are enqueued.
	"""
	jobs = []
	for patcher in (
		patch.object(frappe.db, "commit", lambda *args, **kwargs: None),
		patch("frappe.enqueue", side_effect=lambda method, **kwargs: jobs.append((method, kwargs))),
	):
		patcher.start()
		testcase.addCleanup(patcher.stop)
	return jobs


def run_jobs(jobs: list) -> None:
	"""
	Run collected jobs in order, including jobs they enqueue.

	Parameters:
	        jobs (list, required): List from collect_jobs.

	Returns:
	        None
	"""
	while jobs:
		method, kwargs = jobs.pop(0)
		for option in JOB_OPTIONS:
			kwargs.pop(option, None)
		frappe.get_attr(method)(**kwargs)


def todo_with_folder(jobs: list):
	"""
	A ToDo matching the tenant A rule, with its folder already created.

	Parameters:
	        jobs (list, required): List from collect_jobs.

	Returns:
	        Document: The ToDo.
	"""
	todo = make_todo("High")
	run_jobs(jobs)
	return todo


def attach(todo, file_name: str, content: bytes, **extra):
	"""
	Attach a private file to a ToDo.

	Parameters:
	        todo (Document, required): The ToDo.
	        file_name (str, required): File name.
	        content (bytes, required): File content.
	        **extra: More File fields.

	Returns:
	        Document: The File.
	"""
	return frappe.get_doc(
		{
			"doctype": "File",
			"file_name": file_name,
			"content": content,
			"is_private": 1,
			"attached_to_doctype": "ToDo",
			"attached_to_name": todo.name,
			**extra,
		}
	).insert()


def sandbox_item(item_id: str) -> dict:
	"""
	A sandbox driveItem by id.

	Parameters:
	        item_id (str, required): Item id.

	Returns:
	        dict: The item.
	"""
	items = frappe.cache.get_value(STATE_KEY)["items"].values()
	return next(item for item in items if item["id"] == item_id)
