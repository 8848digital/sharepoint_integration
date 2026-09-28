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

## Key DocTypes
_None yet — this section is filled in as DocTypes are added._

## Features
_Initial app setup only — features are listed here as they are built._

## Integrations
- **Microsoft SharePoint (via Microsoft Graph)** — configuration will be
  documented in `SETUP.md` once the integration is added.

## Installation

    bench get-app sharepoint_integration https://github.com/8848digital/sharepoint_integration --branch develop
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
