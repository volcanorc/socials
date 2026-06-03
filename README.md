# Secure Google Spreadsheet Dashboard

Local-first dashboard for a private Google Spreadsheet. The app authenticates with Google, checks the signed-in account's existing permission on the spreadsheet, and only then loads or mutates data through the Google Sheets API.

## What it does

- Preserves Google sign-in until explicit logout
- Uses the signed-in user's Google permissions to access the private spreadsheet
- Denies access automatically if the account cannot open the spreadsheet
- Lets you browse, search, filter, view, edit, add, and delete rows
- Keeps Google Sheets as the source of truth
- Does not store spreadsheet contents permanently in the app

## Setup

1. Copy `.env.example` to `.env`
2. Fill in your Google OAuth values in `.env`:
   - `GOOGLE_CLIENT_ID`
   - `GOOGLE_CLIENT_SECRET`
   - `GOOGLE_REDIRECT_URI`
3. Keep `SPREADSHEET_ID` set to the provided spreadsheet ID unless you want to target a different private sheet later.
4. Start the app:

```bash
powershell -ExecutionPolicy Bypass -File run-server.ps1
```

For the simplest Windows workflow, double-click `start-local.cmd`.

The launcher will try a cached portable Node runtime first, then a system-wide Node install, and if neither exists it will download a portable Node runtime into `.node-cache` the first time you run it.

## Google OAuth notes

- Create a Google OAuth client for a web application.
- Add the origin from `GOOGLE_REDIRECT_URI` to the OAuth client configuration as an authorized JavaScript origin.
- For local popup mode, `GOOGLE_REDIRECT_URI` should match the app origin exactly, for example `http://127.0.0.1:3001`.
- The app requests `openid`, `email`, `profile`, and Google Sheets access.
- The spreadsheet itself must remain private in Google Drive.
- The server automatically reads local `.env` values at startup.
- Local startup is self-contained, so you do not need `node` on PATH if the launcher can fetch its portable runtime once.
- If Google shows `Access blocked` or `Error 403: access_denied` with a message about the app being tested, open the OAuth consent screen in Google Cloud Console and add the signed-in Google account to **Test users**.
- While the app stays in Testing, only test users can sign in. To let any Google account use it, move the app to **Production** and complete Google verification.

## GitHub hosting note

- The repository can live on GitHub, but the Google OAuth code exchange still needs a running backend.
- A static GitHub Pages site can host the UI, but it cannot safely replace the server-side OAuth callback and Sheets access layer.

## Security model

- The browser never talks directly to Google Sheets
- All access is server-side and tied to the currently signed-in Google account
- There is no separate app allowlist
- Access depends entirely on Google's existing file permissions

## Tests

```bash
node --test
```
