# Secure Spreadsheet Grid

Browser-first dashboard that renders a live Google Sheet as a native editable grid. The current deployment path talks to a Google Apps Script web app, not to the Google Sheets API directly.

## What it does

- Loads the sheet matrix with `GET` from the Apps Script web app
- Renders the data as a responsive HTML table
- Lets you edit any cell inline
- Autosaves each cell back with `POST { row, col, val }`
- Persists the signed-in Google session in `localStorage` until the token expires
- Keeps the UI ready for future restricted-mode auth, without depending on it today

## Current data flow

1. `public/index.html` defines `window.APP_CONFIG`.
2. `public/app.js` reads `appsScriptUrl` from that config.
3. On page load, the app sends a `GET` request to the Apps Script web app URL.
4. The script returns a JSON matrix from the active sheet.
5. The browser renders the matrix into an editable grid.
6. Editing a cell sends a `POST` request with `{ row, col, val }`.
7. The Apps Script web app writes the change back to the spreadsheet.

## Apps Script contract

The deployed script is expected to expose:

- `doGet(e)` returning a JSON 2D array
- `doPost(e)` accepting JSON like:

```json
{ "row": 1, "col": 1, "val": "Hello" }
```

The current script updates the active sheet of the spreadsheet attached to that deployment.

## Setup

1. Open [public/index.html](/C:/Users/Marketing/Documents/Codex/2026-06-03/pull-this-github-repo-and-start/public/index.html) and confirm `window.APP_CONFIG.appsScriptUrl` points to your deployed Apps Script web app.
2. Keep `googleClientId` and `authScope` set for Google sign-in.
3. Deploy the files under `public/` to your static host.
4. Make sure your Apps Script deployment is configured so the browser can call it successfully.

## GitHub Pages deploy

This repo now includes a GitHub Pages workflow at [.github/workflows/pages.yml](/C:/Users/Marketing/Documents/Codex/2026-06-03/pull-this-github-repo-and-start/.github/workflows/pages.yml).

To make the site public on GitHub Pages:

1. Push `main` to GitHub.
2. In the repository settings, enable GitHub Pages and choose the GitHub Actions source if it is not already active.
3. After the workflow finishes, open the Pages URL shown by GitHub.

The app uses relative asset paths, so it works from the repository subpath that GitHub Pages serves.

## Session persistence

- The browser stores the Google access token, expiry, and profile in `localStorage`.
- Refreshing the page restores the session until the token expires.
- Logout clears the stored session.
- The stored session is for UI convenience only; the spreadsheet data still comes from the live web app.

## Security model

- The spreadsheet remains the source of truth in Google Drive.
- A signed-in user can edit cells only if the Apps Script deployment and spreadsheet permissions allow it.
- The browser is not using an HttpOnly cookie session for the current static flow.
- The spreadsheet ID is not secret and is safe to expose in `window.APP_CONFIG`.

## Notes

- If browser calls to `script.google.com` are blocked by CORS in your hosting environment, the frontend will need a same-origin proxy.
- The legacy Node backend files are still in the repo, but the current static path does not depend on them.

## Tests

```bash
node --test
```
