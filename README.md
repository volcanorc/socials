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
2. Fill in your Google OAuth values:
   - `GOOGLE_CLIENT_ID`
   - `GOOGLE_CLIENT_SECRET`
   - `GOOGLE_REDIRECT_URI`
3. Keep `SPREADSHEET_ID` set to the provided spreadsheet ID unless you want to target a different private sheet later.
4. Start the app:

```bash
node server.js
```

## Google OAuth notes

- Create a Google OAuth client for a web application.
- Add the origin from `GOOGLE_REDIRECT_URI` to the OAuth client configuration as an authorized JavaScript origin.
- For popup mode, `GOOGLE_REDIRECT_URI` should be the app origin, for example `http://localhost:3000`.
- The app requests `openid`, `email`, `profile`, and Google Sheets access.
- The spreadsheet itself must remain private in Google Drive.

## Security model

- The browser never talks directly to Google Sheets
- All access is server-side and tied to the currently signed-in Google account
- There is no separate app allowlist
- Access depends entirely on Google's existing file permissions

## Tests

```bash
node --test
```
