<!--
Copyright (c) 2026 8848 Digital LLP. All rights reserved.
Proprietary and confidential. Unauthorized copying, distribution, or use
of this file, via any medium, is strictly prohibited without prior
written permission from 8848 Digital LLP.
-->

# SharePoint Core

## Purpose
Core module of the SharePoint Integration app: connection settings, folder
rules, folder provisioning through Microsoft Graph, and access management via
per-document security groups.

## DocTypes

| DocType | Purpose |
| ------- | ------- |
| SharePoint Connection | One tenant / Entra app registration: base URL, tenant id, client id, secret. |
| SharePoint Settings | Single: global switch, sandbox, default connection, access/expiry/notification options. |
| SharePoint Folder Rule | Rule per DocType/variant: connection, trigger, condition, site/library/path, name template, subfolders, owners, template file. |
| SharePoint Rule Subfolder | Child of Folder Rule: subfolder name and whether the group may open it. |
| SharePoint Rule Owner Field | Child of Folder Rule: user field on the reference DocType whose user gets write access. |
| SharePoint Folder | One per reference document: connection, status, ids, link, group, error, attempts. |
| SharePoint Access | Members with access to a document's folder (one per document). |
| SharePoint Access Member | Child of SharePoint Access: user, dates, access status. |
| SharePoint Notification Recipient | Child of Settings: users emailed on failures. |

## Customizations

| DocType | Owned by | What's customized |
| ------- | -------- | ------------------ |
| Every rule's Reference DocType | Any app | Custom fields `sharepoint_section`, `sharepoint_folder`, `sharepoint_link` (added when a rule is saved). |

## Code layout
- `graph/` — Microsoft Graph client (cached token, retries), drive/directory services, sandbox.
- `services/` — rule matching, folder request/provisioning, access sync/expiry, notifications.
- `document_events.py` — `doc_events["*"]` handlers (cached early exit for DocTypes without rules).
- `tasks.py` — background job and scheduler entry points.
- `api/v1/` — `folder`, `settings`, `access` endpoints.
