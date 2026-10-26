<!--
Copyright (c) 2026 8848 Digital LLP. All rights reserved.
Proprietary and confidential. Unauthorized copying, distribution, or use
of this file, via any medium, is strictly prohibited without prior
written permission from 8848 Digital LLP.
-->

# SharePoint Integration

## Overview
Custom Frappe/ERPNext app by 8848 Digital LLP that connects an ERPNext site to
Microsoft SharePoint, so business documents and their attachments can be stored
and accessed in SharePoint document libraries alongside the ERP records they
belong to.

It is generic: any DocType (standard, ERPNext or custom) can get SharePoint
folders by adding a Folder Rule for it — no code changes per client.

## Key DocTypes

| DocType | Owned by this app? | Purpose |
| ------- | ------------------ | ------- |
| SharePoint Connection | Yes | One Microsoft tenant / Entra app registration (base URL, tenant id, client id, secret). A client can have several. |
| SharePoint Settings | Yes | Global switch, sandbox mode, default connection, access and failure-email options. |
| SharePoint Folder Rule | Yes | For one DocType (optionally one variant, e.g. External projects): which connection, when to create a folder, where, its name, subfolders, owners and template file. |
| SharePoint Folder | Yes | One record per document's folder: status, link, ids, security group and errors. Retry from here. |
| SharePoint Access | Yes | Who can open a document's folder and for which dates; synced to the document's security group. |
| SharePoint File | Yes | One record per attachment sent to SharePoint: status, link, whether the local copy was removed, errors. Retry from here. |
| Any DocType with a rule | No (customized) | Gets a "SharePoint" section (folder + link) and "Open in SharePoint" / "Create Folder" buttons. The buttons show only while SharePoint Settings is enabled and the DocType has an enabled rule; "Create Folder" only for users who can edit the document. |

## Features
- Create a SharePoint folder automatically when a document is created, saved,
  or reaches a workflow state — or on demand with a button.
- Multiple SharePoint tenants / app registrations per client: each rule picks
  its connection, and each folder remembers the tenant it was created in.
- Different site/path/subfolders per DocType and per variant (rule conditions,
  e.g. External vs Internal projects).
- Folder names from a template (e.g. `{{ doc.party_name }}-{{ doc.name }}`),
  cleaned of characters SharePoint rejects.
- Subfolders created automatically, each with its own group-access setting.
- A template file (PPT/Word/any size) copied into every new folder server-side.
- One security group per document; people get access by being added to the
  group, so access applies consistently and can be removed cleanly.
- Owners (from user fields on the document) always get write access.
- Access dates with automatic expiry, and **Revoke Access** to remove selected
  people immediately with a reason. Expired/revoked rows are kept as an audit
  trail (rows that had access can't be deleted or switched to another user;
  re-adding a person creates a new row).
- Attachments: files attached to a document are uploaded into its SharePoint
  folder (optionally into a subfolder such as "Attachments"), including files
  attached before the folder existed. Large files are uploaded in pieces.
  Optionally the local copy is deleted after upload and the attachment opens
  the SharePoint file instead, saving server disk space.
- **Create Folders for Existing Records** on a Folder Rule: at go-live, create
  folders (and upload attachments) for documents that existed before the rule.
- Failure log per document, retry button, and optional failure emails.
- Sandbox mode to demo/QA the whole flow without a Microsoft tenant.

## Integrations
- **Microsoft SharePoint & Entra ID (via Microsoft Graph)** — see [SETUP.md](./SETUP.md#microsoft-sharepoint--entra-id-microsoft-graph)

## Installation

Requires Frappe v15 or v16.

    bench get-app sharepoint_integration https://github.com/8848digital/sharepoint_integration --branch integration-master
    bench --site <site_name> install-app sharepoint_integration

## App Structure
See [CLAUDE.md](./CLAUDE.md) for internal module/folder layout and coding conventions.

## Contributing
This app uses `pre-commit` for formatting and linting (black, isort, flake8,
prettier, eslint, max-lines check) and Conventional Commits:

    cd apps/sharepoint_integration
    pre-commit install

## Maintainers
8848 Digital LLP — mahak@8848digital.com

## License
Proprietary — Copyright (c) 2026 8848 Digital LLP. All rights reserved.
See [license.txt](license.txt) for details.
