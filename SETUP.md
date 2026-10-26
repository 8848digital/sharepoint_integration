<!--
Copyright (c) 2026 8848 Digital LLP. All rights reserved.
Proprietary and confidential. Unauthorized copying, distribution, or use
of this file, via any medium, is strictly prohibited without prior
written permission from 8848 Digital LLP.
-->

# SETUP.md

## Microsoft SharePoint & Entra ID (Microsoft Graph)

### Overview
Creates document folders in SharePoint, copies template files, creates one
security group per document and manages its members — all through Microsoft
Graph with an app-only (client credentials) token.

### Required Credentials
| Credential | Where it's stored | Required? |
| ---------- | ------------------ | --------- |
| SharePoint Base URL | SharePoint Connection → `base_site` | Yes |
| Tenant ID | SharePoint Connection → `tenant_id` | Yes |
| Client ID | SharePoint Connection → `client_id` | Yes |
| Client Secret | SharePoint Connection → `client_secret` (encrypted) | Yes (not in Sandbox Mode) |

Create **one SharePoint Connection per Microsoft tenant or app registration**.
A client with several tenants (or a separate app per site) simply has several
connections; each Folder Rule picks one, and each SharePoint Folder records the
connection it was created on so later access changes go to the right tenant.

**Entra app registration** (Azure portal → App registrations → New):
Application permissions (admin consent required):

| Permission | Why |
| ---------- | --- |
| `Sites.ReadWrite.All` (or `Sites.Selected` + per-site grant) | Create folders, copy templates, set folder permissions |
| `Group.ReadWrite.All` | Create per-document security groups, add/remove members |
| `User.Read.All` | Resolve member emails to user ids |

### Settings DocType Fields
| Field | Type | Mandatory | Notes |
| ----- | ---- | --------- | ----- |
**SharePoint Connection**

| Field | Type | Mandatory | Notes |
| ----- | ---- | --------- | ----- |
| Connection Name | Data | Yes | e.g. `Contoso Main`, `HD SharePoint`. |
| SharePoint Base URL | Data | Yes | Tenant root only, e.g. `https://contoso.sharepoint.com`. |
| Tenant ID | Data | Yes | GUID or `contoso.onmicrosoft.com`. |
| Client ID | Data | Yes | GUID of the app registration. |
| Client Secret | Password | Yes (live) | App secret value. |

**SharePoint Settings**

| Field | Type | Mandatory | Notes |
| ----- | ---- | --------- | ----- |
| Enabled | Check | — | Master switch; needs at least one enabled connection. |
| Sandbox Mode | Check | — | Simulates Graph locally. Demo/QA only — never on production. |
| Default Connection | Link | — | Pre-filled on new Folder Rules. |
| Group Access Role | Select | — | `write` (default) or `read` for the document group. |
| Restrict Folder Access | Check | — | Remove grants copied from the parent folder (the "break inheritance" behaviour). |
| Keep Site Owners | Check | — | Never remove the site owners' grant. |
| Auto Expire After (Days) | Int | — | Expiry Date = To Date + N when not set. 0 = off. |
| Email on Failure / Recipients | Check / Table MultiSelect | Recipients required if on | Folder or access failures are emailed. |

**SharePoint Folder Rule** — mandatory: Rule Name, Connection, Reference DocType, Trigger,
Folder Name Template; Workflow State (for "On Workflow State"); Template File
ID + File Name (when "Copy Template File" is on). Site Name / Document Library
blank = root site / default library. Use **Resolve Site** to verify.

Attachments (optional): **Upload Attachments** sends files attached to the
document to its folder; **Attachment Subfolder** (e.g. `Attachments`) puts them
in a subfolder; **Remove Local Copy** deletes the file from the server after
upload and points the attachment to SharePoint (people then need SharePoint
access to open it; files used in a form field, such as an image, are always
kept). No extra Graph permission is needed beyond the site permission above.

**Create Folders for Existing Records** (button on a saved, enabled rule) runs
in the background and messages you the totals when it ends.

Template File ID: open the file in SharePoint → Details → copy the item id, or
`GET https://graph.microsoft.com/v1.0/drives/{drive-id}/root:/{path-to-file}`
and take `id`. If the template lives in another library, set Template Drive ID.

### Site Config / Environment Variables
None. Background workers (`bench start` / supervisor) must be running: folder
creation, attachment uploads and access sync run as background jobs; expiry
runs daily, and access grants and failed uploads are retried hourly via the
scheduler.

### Webhooks (if applicable)
None.

### How to Test
1. Create a **SharePoint Connection** → Save → **Test Connection**. Then
   SharePoint Settings → tick Enabled (and Sandbox Mode for a dry run) → Save.
2. Create a SharePoint Folder Rule (e.g. Reference DocType `ToDo`, Trigger
   `On Create`) → Save → **Resolve Site**.
3. Create a matching document → open **SharePoint Folder** list → status
   becomes `Created`; the document shows **SharePoint → Open in SharePoint**.
4. Create a **SharePoint Access** for that document with a member → Save →
   member status becomes `Granted`.
5. Select the member row → **Revoke Access** → enter a reason → status becomes
   `Revoked` with Removed On and Remarks filled.
6. Tick **Upload Attachments** on the rule → attach a file to the document →
   **SharePoint File** list shows it `Uploaded` with a link to the file.
7. On the rule click **Create Folders for Existing Records** → confirm → a
   message shows how many folders and uploads were queued.
