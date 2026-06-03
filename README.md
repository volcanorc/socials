# Secure Google Spreadsheet Dashboard

Browser-only dashboard for a private Google Spreadsheet. The app authenticates with Google in the browser, persists the signed-in state in `localStorage` until the access token expires, and loads or mutates data directly through the Google Sheets API.

## What it does

- Preserves Google sign-in until explicit logout
- Uses the signed-in user's Google permissions to access the private spreadsheet
- Denies access automatically if the account cannot open the spreadsheet
- Lets you browse, search, filter, view, edit, add, and delete rows
- Keeps Google Sheets as the source of truth
- Does not store spreadsheet contents permanently in the app

## Setup

1. Edit `public/index.html` and update the `window.APP_CONFIG` values if needed:
   - `googleClientId`
   - `spreadsheetId`
   - `authScope`
2. Deploy the `public/` folder as a static site, or open it through your preferred static host.
3. If you want to preview locally with the bundled Node server, you can still use the existing launcher scripts:

```bash
powershell -ExecutionPolicy Bypass -File run-server.ps1
```

For the simplest Windows workflow, double-click `start-local.cmd`.

## Google OAuth notes

- Create a Google OAuth client for a web application.
- Add the origin where you host the static app to the OAuth client configuration as an authorized JavaScript origin.
- For local popup mode, make sure the static origin matches the OAuth client configuration exactly, for example `http://127.0.0.1:3000`.
- The app requests `openid`, `email`, `profile`, and Google Sheets access.
- The spreadsheet itself must remain private in Google Drive.
- The browser stores the signed-in Google access token in `localStorage` so refreshes can reopen the dashboard until the token expires.
- If Google shows `Access blocked` or `Error 403: access_denied` with a message about the app being tested, open the OAuth consent screen in Google Cloud Console and add the signed-in Google account to **Test users**.
- While the app stays in Testing, only test users can sign in. To let any Google account use it, move the app to **Production** and complete Google verification.

## GitHub hosting note

- A static GitHub Pages site can host the UI because the browser talks to Google directly now.
- The spreadsheet ID is not secret, so exposing it in client config is acceptable.

## Security model

- The browser talks directly to Google Sheets after sign-in
- Access is still governed by Google's existing file permissions
- A signed-in user can load and edit the spreadsheet only if Google Drive already grants that account access
- Signed-in but unauthorized users get a clear no-access state instead of a silent logout

## Tests

```bash
node --test
```
