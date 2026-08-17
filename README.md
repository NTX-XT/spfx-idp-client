# Nintex Workflow SharePoint Webparts

This project contains SharePoint Framework (SPFx) webparts for Nintex Automation Cloud that let users work directly from SharePoint:

- view workflow tasks
- approve or reject tasks where supported
- open task forms and start forms from SharePoint
- authenticate using Nintex OAuth with PKCE

It is designed for SharePoint tenants that want a lightweight, browser-based Nintex experience without leaving their intranet pages.

## Documentation

- [User Guide](sharepoint/USER-GUIDE.md) — installation, configuration, and day-to-day usage
- [Project setup and build notes](#sharepoint-app-sharepoint)

## Download

The packaged SharePoint solution is published as a GitHub Release asset so end users can install it without building anything locally.

- [Download v2.0.6 package](https://github.com/NTX-XT/spfx-idp-client/releases/download/v2.0.6/nac-tasks.sppkg)

Upload the package to your SharePoint App Catalog to deploy the solution. If you need a custom build, you can still build locally from the project source.

## Webparts

### Nintex Workflow Tasks
View your assigned Nintex tasks with status filtering, rich HTML message rendering, and direct approve/reject buttons for tasks with Express Approval enabled. Tasks without Express Approval show a link to open their task form instead.

### Nintex Workflow Forms
Browse the start forms available to you and open them directly from SharePoint. Each form card is fully clickable and opens the Nintex form in a new tab.

---

## SPFx Version

![version](https://img.shields.io/badge/SPFx-1.23.0-green.svg)
![node](https://img.shields.io/badge/Node-22.x-green.svg)

---

## Prerequisites

### Node.js
This project requires **Node.js 22 LTS**. Use [nvm-windows](https://github.com/coreybutler/nvm-windows) to manage Node versions:
```powershell
nvm install 22
nvm use 22
```

### Nintex OAuth Client App
Register a Client App in your Nintex tenant:
1. Go to your Nintex tenant → **Settings → Tenant → Apps and tokens → App registrations**
2. Create a new app and enable **Allow public clients (Implicit Grant & PKCE)**
3. Add your SharePoint page URL as the **Redirect URI** — this must exactly match the page where the webpart is placed, e.g.:
   ```
   https://yourtenant.sharepoint.com/sites/YourSite/SitePages/Home.aspx
   ```
4. Enable the following scopes:
   - `nc:task:read`
   - `nc:task:write` (required for approving/rejecting tasks)
   - `nc:forms:read`
5. Copy the **Client ID** — you'll need it when configuring the webpart

---

## SharePoint app (`sharepoint/`)

The production SharePoint package lives under the `sharepoint/` folder. It uses OAuth 2.0 PKCE against `auth.nintexcloud.com` with a full-page redirect flow: users sign in to Nintex, return to the SharePoint page, and the webpart exchanges the authorization code for a token before loading tasks or forms.

### Building

```powershell
cd sharepoint

# Install dependencies (only needed once, or after package.json changes)
npm install --legacy-peer-deps

# Build and package for production
npm run build
```

The `.sppkg` file will be output to `sharepoint/sharepoint/solution/`.

> **Note:** `npm run build` uses `heft` instead of `gulp`. Do not use `gulp bundle` or `gulp package-solution` — these commands no longer exist in SPFx 1.23.

### Deploying to SharePoint

1. Go to your **SharePoint Admin Centre → More features → Apps → App Catalog**
2. Upload the `.sppkg` from `sharepoint/sharepoint/solution/`
3. When prompted, select **"Make this solution available to all sites"** (tenant-wide deployment) — this avoids needing to add it to each site manually
4. Navigate to the SharePoint page where you want to add the webpart
5. Edit the page and add **"Nintex Workflow Tasks"** and/or **"Nintex Workflow Forms"** from the webpart picker
6. Configure each webpart via its property pane (gear icon):
   - **Client ID** — from your Nintex OAuth app registration
   - **Redirect URI** — the full URL of the current page (must match what's registered in Nintex)
   - **Region** — select your Nintex tenant region (AU, US, EU, CA, UK)
7. Save and republish the page
8. Click **"Log In and Get Tasks"** / **"Log In and Get Forms"** to authenticate

---

## Versioning

When making changes, bump the version in two files **within the project folder you're building** before building:

- `package.json` → `"version"` field (e.g. `2.0.0` → `2.0.1`)
- `config/package-solution.json` → both `solution.version` and `features[0].version` (e.g. `2.0.0.0` → `2.0.1.0`)

SharePoint uses the solution version to detect updates and prompt users to upgrade.

---

## Authentication Notes

The app uses [`oidc-client-ts`](https://github.com/authts/oidc-client-ts) for OAuth 2.0 PKCE authentication against `auth.nintexcloud.com`. Key configuration:

- The `userStore` and `stateStore` are set to `localStorage` (not `sessionStorage`) to survive SharePoint's page redirect behaviour during the OAuth callback
- The Nintex client ID format (`{tenantId}_{appId}`) includes an underscore — `oidc-client-ts`'s built-in `getUser()` may not resolve this correctly, so `AuthService.getAccessToken()` includes a fallback that scans `localStorage` directly for a matching `oidc.user:` key
- OIDC discovery metadata is provided inline (in `Config.ts`) so the SDK does not need to fetch `.well-known/openid-configuration` at runtime

---

## Project Structure

The folder follows this layout:

```
sharepoint/
├── config/
│   ├── config.json           # Bundle definitions for each webpart
│   ├── package-solution.json # Solution version and deployment settings
│   ├── rig.json              # Points build tools to spfx-web-build-rig
│   ├── sass.json             # Heft sass plugin config
│   └── typescript.json       # TypeScript build settings
├── src/webparts/
│   ├── nacTaskActions/       # Nintex Workflow Tasks webpart
│   │   └── components/
│   │       ├── AuthService.ts       # OIDC auth with localStorage fallback
│   │       ├── Config.ts            # oidc-client-ts UserManager settings
│   │       ├── ITaskProps.ts        # Task and assignment TypeScript interfaces
│   │       ├── NacTaskActions.tsx   # Main component — task list, filters, approve/reject
│   │       └── NacTaskActions.module.scss
│   └── nacForms/             # Nintex Workflow Forms webpart
│       └── components/
│           ├── AuthService.ts
│           ├── Config.ts
│           ├── IFormProps.ts        # Form TypeScript interface
│           ├── NacForms.tsx         # Main component — form list with clickable cards
│           └── NacForms.module.scss
├── eslint.config.js          # Flat ESLint config (ESLint 9)
├── tsconfig.json             # Extends spfx-web-build-rig base config
└── package.json
```

---

## Nintex API Endpoints Used

| Feature | Endpoint |
|---|---|
| List tasks | `GET /workflows/v2/tasks?status={status}&uiRequest=true` |
| Complete task | `PATCH /workflows/v2/tasks/{taskId}/assignments/{assignmentId}` |
| List forms | `GET /workflows/v1/forms` |

All endpoints target your regional base URL (e.g. `https://au.nintex.io`).

The `uiRequest=true` parameter on the tasks endpoint returns additional UI-specific fields including `expressApproval` (set to `"include"` when the task supports direct approve/reject) and `validOutcomes` (the available outcome options for that task).

---

## History

For historical reference, an earlier version of this project is preserved in git under the tag `v1-combined-pre-split`.
