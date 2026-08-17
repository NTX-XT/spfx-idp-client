# Nintex Workflow Tasks & Forms — Help Guide

This guide covers everything you need to add, configure, and use the Nintex Workflow webparts on a SharePoint page.

---

## What's included

Two webparts:

- **Nintex Workflow Tasks** — shows your pending and completed workflow tasks, with Approve/Reject actions and optional comments.
- **Nintex Workflow Forms** — shows available workflow start forms as clickable cards.

---

## Adding the webpart to a page

1. Edit the SharePoint page you want the webpart on.
2. Click the **+** to add a new webpart.
3. Search for **"Nintex Workflow Tasks"** or **"Nintex Workflow Forms"** and select it.
4. Click the **pencil/edit icon** on the webpart to open its settings pane.

---

## Configuration (one-time setup per webpart)

You'll need the following from your Nintex administrator:

| Setting | What it is |
|---|---|
| **Client ID** | The Nintex client app ID registered for this integration |
| **Redirect URI** | Usually auto-filled to match the page URL — only change if instructed |
| **Region** | Your Nintex tenant's region: US, Europe, Asia Pacific, Canada, or United Kingdom |
| **Use Production** | Toggle on for live/production Nintex, off for test/sandbox |
| **Hide logo and greeting** | Optional checkbox — hides the Nintex logo and "Good morning, [Name]" header if you want a more compact view |

Once configured, click **Apply** / **Publish** the page.

---

## Using the Tasks webpart

**First time on the page:**
Click **Log In and Get Tasks**. You'll be redirected to Nintex to sign in, then brought back with your tasks loaded.

**After that:**
If you still have an active Nintex session (e.g. you're logged into Nintex in another tab), your tasks will load automatically — no login needed. Sessions refresh themselves in the background while the page is open.

**Filtering tasks:**
Use the filter pills at the top (All, Active, Complete, etc.) to narrow down the task list.

**Responding to a task:**
1. Click **Approve** or **Reject** (or whichever outcome buttons are shown) on a task.
2. A confirmation dialog appears — you can optionally type a comment in the **Comment (optional)** box.
3. Click **Confirm** to submit. The task will update and show a success message.

**Opening a task's full form:**
If a task has an associated form, click **Open form to respond** to view/complete it in a new tab.

---

## Using the Forms webpart

Browse the list of available workflow start forms. Click any form card to open it in a new tab and start that workflow.

---

## Troubleshooting

| Problem | What to try |
|---|---|
| Stuck on "Log In" every time | Your Nintex session may be expiring quickly, or you don't have an active Nintex session elsewhere in the browser. Try logging into Nintex directly in another tab first. |
| Tasks not loading after login | Refresh the page. If it persists, contact your SharePoint/Nintex administrator — the Client ID or Region may be misconfigured. |
| "Log In" redirects to the wrong account | Your browser may be signed into multiple Microsoft/Nintex accounts. Try a private/incognito window, or sign out of other accounts first. |
| Comment box doesn't appear | Make sure you're clicking directly on an outcome button (Approve/Reject) — the comment field only appears in the confirmation step, not on the task card itself. |

---

## Who to contact

For access issues (can't see expected tasks/forms), contact your **Nintex administrator**.
For webpart display or configuration issues, contact your **SharePoint site administrator**.
