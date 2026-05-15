# Bang Creations — Employee Portal

## Architecture

Vanilla JS, no build step. Scripts load in order via `<script src="...">` tags at the bottom of `index.html`.

| File | Responsibility |
|------|---------------|
| `index.html` | HTML shell only — screens, nav, panel markup. No logic. |
| `style.css` | All portal CSS. Uses CSS custom properties defined in `:root`. Also contains Handbook tab UI styles (hero, search, tag pills, result rows, viewer panel). |
| `handbook.css` | Shared stylesheet for Handbook page fragments. Self-contained — includes its own `:root` custom properties (mirroring `style.css`) and a `body` reset so pages render correctly both inside the portal viewer iframe and when opened directly in a browser. Fetched and injected by `handbook.js` at viewer open time; not loaded via `<link>` in `index.html`. |
| `handbook-template-page.html` | Starter template for Handbook Editors. HTML fragment (no `<html>/<head>/<body>`) using `handbook.css` class names. Copy, fill in, upload to SharePoint. |
| `config.js` | `CFG` object (IDs, URLs, list names, group IDs), `SCOPES`, all global `g*` state variables. Load first. |
| `main.js` | `showScreen`, `setLoading`, `showFatalError`, `switchTab`, `applyHrAccess`, `applyAccAccess`. `switchTab` also closes the Handbook viewer when navigating away from the Handbook tab. |
| `graph.js` | Microsoft Graph helpers: `gGet`, `gPost`, `gDel`, `getSiteId`, `getListId`, `fetchUsers`, shared utilities (`fmt`, `escHtml`, `nameByEmail`, etc.). |
| `hr-tasks.js` | Everything for the HR Tasks tab: `fetchTasks`, `fetchComps`, `renderMy`, `renderManage`, `addTask`, `delTask`, `toggleComp`, assignee UI helpers. |
| `expenses.js` | Everything for the Expenses tab: `fetchExpenses`, `fetchMileage`, `renderMyExpenses`, `renderAccMgrTable`, `addStdExpense`, `addMilExpense`, VAT helpers, debug tool. |
| `handbook.js` | Everything for the Handbook tab. See detail below. |
| `auth.js` | MSAL init, `getToken`, `doSignIn`, `doSignOut`, `checkHrAdmin`, `checkAccMgr`, `checkHandbookEditor`, `bootstrap`. Loaded last — kicks off the app. |
| `msal-browser.min.js` | Microsoft Authentication Library — must be present in the repo root (not tracked here). |

## Script load order

```
config.js → main.js → graph.js → hr-tasks.js → expenses.js → handbook.js → auth.js
```

## SharePoint

- **Site:** https://bangcreations.sharepoint.com/sites/BANGTEAM
- **Lists:** `HR_Tasks`, `expenses_data`
- **HR_Tasks columns (completion):** `completed` (Yes/No, default No), `CompletedOn` (DateTime) — completion is task-level, not per-assignee
- **expenses_data columns (mileage merge):** `expense_type` (text: `'standard'` or `'mileage'`), `distance` (text), `rate` (text) — mileage records share `expense_date` and `person` with standard expenses
- **Handbook documents:** HTML files in the `Documents` library under the folder `Handbook and HR`. Documents are organised into category sub-folders (e.g. `Handbook and HR/Ethics & Conduct/`). The portal recurses into all sub-folders to build the full document list.
- Standard expense receipts are uploaded as native SharePoint list item attachments via the SharePoint REST API (not Graph), using a SharePoint-scoped token (`https://bangcreations.sharepoint.com/.default`).

## Entra groups

| Group | Object ID |
|-------|-----------|
| HR Administrators | `cec65ee4-f9fc-4beb-9b68-2f08f4ec78e6` |
| Accounts Manager | `a6ec43fb-a114-4c72-81e5-064fd2080d58` |
| Handbook Editors | `2f5f0cf0-9295-4cc0-90e2-b32fe775d011` |

Access checks use `checkMemberObjects` (transitive membership) for all three groups. Owners who are not members are not granted access — membership is required in all cases.

## Key patterns

- All Graph calls go through `gGet` / `gPost` / `gDel` in `graph.js`.
- Global state lives in `config.js` (`gTasks`, `gExpenses`, `gUsers`, `gHbDocs`, etc.).
- Render functions are named `renderMy`, `renderManage`, `renderMyExpenses`, `renderAccMgrTable`, `renderHandbook`.
- Toasts: `showToast` (HR tab), `expShowToast` (Expenses tab), `hbShowToast` (Handbook tab).
- `bootstrap()` in `auth.js` orchestrates the full startup sequence.

## Handbook tab — handbook.js

### Data flow
1. `fetchHbCss()` — fetches `./handbook.css` once and caches it in `gHbCss`.
2. `fetchHbDrive()` — resolves the SharePoint Documents drive ID via `/sites/{siteId}/drives`.
3. `fetchHbDocs()` → `fetchHbFolder(segments)` — recursively lists all `.html` files in `Handbook and HR` and every sub-folder. Each driveItem is expanded with `listItem($expand=fields)` to read the `FileTypeTag` multi-choice column (returns either an array or a `;#`-delimited string depending on SharePoint config — both are handled).
4. `renderHandbook()` — renders tag pills and result rows.

### Viewer
- `hbOpenViewer(docId)` fetches a fresh `@microsoft.graph.downloadUrl` per open (short-lived pre-auth URL), retrieves the HTML as text, passes it through `hbWrapDoc()`, then loads the result as a blob URL in a sandboxed `<iframe>`.
- `hbWrapDoc(html)` detects whether the HTML is a fragment or a full document. Fragments are wrapped in a complete `<html>` shell with `gHbCss` injected as a `<style>` block. Full documents get the style block spliced before `</head>`. This avoids relative URL resolution failures inside blob URLs.
- The viewer slides in from the right. Esc key or backdrop click closes it.

### Permissions
- `checkHandbookEditor()` in `handbook.js` checks membership of the Handbook Editors group.
- Handbook Editors see an **Edit in SharePoint** button on each result row and inside the viewer header, linking to the file's `webUrl`.

### Deep links
- Opening a document writes `#handbook/{itemId}` to the URL.
- On bootstrap completion, `hbHandleDeepLink()` reads the hash, switches to the Handbook tab, and opens the document automatically.
- The share button copies `origin + pathname + #handbook/{itemId}` to the clipboard.

### State variables (in config.js)
| Variable | Purpose |
|----------|---------|
| `gHbDriveId` | Documents drive ID |
| `gHbDocs` | Array of `{ id, name, tags[], updated, webUrl }` |
| `gHbTags` | Sorted unique tag strings derived from all documents |
| `gHbActiveTags` | Currently selected tag filters |
| `gHbSearch` | Current search input value |
| `gHbViewing` | ID of the document open in the viewer, or `null` |
| `gHbBlobUrl` | Active blob URL (revoked and replaced on each open) |
| `gIsHandbookEditor` | Whether the signed-in user is a Handbook Editor |

`gHbCss` (the cached stylesheet content) is a module-level variable in `handbook.js`.

## Expenses tab — layout note

The Accounts Manager — All Expenses section (`#acc-admin-section`) sits at the **top** of the Expenses panel so it is immediately visible to Accounts Managers. It remains `display:none` for all other users and is shown by `applyAccAccess()` in `main.js`.
