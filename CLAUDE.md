# Bang Creations — Employee Portal

## Architecture

Vanilla JS, no build step. Scripts load in order via `<script src="...">` tags at the bottom of `index.html`.

| File | Responsibility |
|------|---------------|
| `index.html` | HTML shell only — screens, nav, panel markup. No logic. |
| `style.css` | All CSS. Uses CSS custom properties defined in `:root`. |
| `config.js` | `CFG` object (IDs, URLs, list names), `SCOPES`, all global `g*` state variables. Load first. |
| `main.js` | `showScreen`, `setLoading`, `showFatalError`, `switchTab`, `applyHrAccess`, `applyAccAccess`. |
| `graph.js` | Microsoft Graph helpers: `gGet`, `gPost`, `gDel`, `getSiteId`, `getListId`, `fetchUsers`, shared utilities (`fmt`, `escHtml`, `nameByEmail`, etc.). |
| `hr-tasks.js` | Everything for the HR Tasks tab: `fetchTasks`, `fetchComps`, `renderMy`, `renderManage`, `addTask`, `delTask`, `toggleComp`, assignee UI helpers. |
| `expenses.js` | Everything for the Expenses tab: `fetchExpenses`, `fetchMileage`, `renderMyExpenses`, `renderAccMgrTable`, `addStdExpense`, `addMilExpense`, VAT helpers, debug tool. |
| `auth.js` | MSAL init, `getToken`, `doSignIn`, `doSignOut`, `checkHrAdmin`, `checkAccMgr`, `bootstrap`. Loaded last — kicks off the app. |
| `msal-browser.min.js` | Microsoft Authentication Library — must be present in the repo root (not tracked here). |

## Script load order

```
config.js → main.js → graph.js → hr-tasks.js → expenses.js → auth.js
```

## SharePoint

- **Site:** https://bangcreations.sharepoint.com/sites/BANGTEAM
- **Lists:** `HR_Tasks`, `HR_Completions`, `expenses_data`, `mileage_data`
- Standard expense receipts are uploaded as native SharePoint list item attachments via the SharePoint REST API (not Graph), using a SharePoint-scoped token (`https://bangcreations.sharepoint.com/.default`).

## Entra groups

| Group | Object ID |
|-------|-----------|
| HR Administrators | `cec65ee4-f9fc-4beb-9b68-2f08f4ec78e6` |
| Accounts Manager | `a6ec43fb-a114-4c72-81e5-064fd2080d58` |

Access checks use `checkMemberObjects` (transitive membership) with a fallback to `/groups/{id}/owners` for group owners who are not members.

## Key patterns

- All Graph calls go through `gGet` / `gPost` / `gDel` in `graph.js`.
- Global state lives in `config.js` (`gTasks`, `gExpenses`, `gUsers`, etc.).
- Render functions are named `renderMy`, `renderManage`, `renderMyExpenses`, `renderAccMgrTable`.
- Toasts: `showToast` (HR tab), `expShowToast` (Expenses tab).
- `bootstrap()` in `auth.js` orchestrates the full startup sequence.

## Pending

- Remove the debug button (`#acc-debug-wrap`) in `index.html` and `runAccDebug()` in `expenses.js` once the Accounts Manager access issue is resolved.
