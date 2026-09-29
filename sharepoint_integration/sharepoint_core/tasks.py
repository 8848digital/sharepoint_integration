# Copyright (c) 2026 8848 Digital LLP. All rights reserved.
# Proprietary and confidential. Unauthorized copying, distribution, or use
# of this file, via any medium, is strictly prohibited without prior
# written permission from 8848 Digital LLP.

"""Background job and scheduler entry points (thin; logic lives in services/)."""

from sharepoint_integration.sharepoint_core.services.access_expiry import (
	expire_access,
	retry_pending_access,
)
from sharepoint_integration.sharepoint_core.services.access_revoke import revoke_members
from sharepoint_integration.sharepoint_core.services.access_sync import AccessSync
from sharepoint_integration.sharepoint_core.services.folder_provisioner import FolderProvisioner


def provision_folder(folder: str) -> None:
	"""
	Job: create the SharePoint folder tree for a SharePoint Folder record.

	Parameters:
	        folder (str, required): SharePoint Folder name.

	Returns:
	        None
	"""
	FolderProvisioner(folder).run()


def sync_access(access: str) -> None:
	"""
	Job: add a SharePoint Access's Pending/Failed members to its security group.

	Parameters:
	        access (str, required): SharePoint Access name.

	Returns:
	        None
	"""
	AccessSync(access).run()


def revoke_access(access: str, members: list[str]) -> None:
	"""
	Job: remove selected members from the group and mark them Revoked.

	Parameters:
	        access (str, required): SharePoint Access name.
	        members (list[str], required): SharePoint Access Member row names.

	Returns:
	        None
	"""
	revoke_members(access, members)


def daily() -> None:
	"""
	Scheduler (daily): expire access past its Expiry Date.

	Returns:
	        None
	"""
	expire_access()


def hourly() -> None:
	"""
	Scheduler (hourly): retry Pending/Failed access grants.

	Returns:
	        None
	"""
	retry_pending_access()
