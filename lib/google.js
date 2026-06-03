const GOOGLE_AUTH_URL = 'https://accounts.google.com/o/oauth2/v2/auth';
const GOOGLE_TOKEN_URL = 'https://oauth2.googleapis.com/token';
const GOOGLE_USERINFO_URL = 'https://openidconnect.googleapis.com/v1/userinfo';
const GOOGLE_SHEETS_BASE = 'https://sheets.googleapis.com/v4/spreadsheets';

function toError(message, status = 500, details = null) {
  const error = new Error(message);
  error.status = status;
  error.details = details;
  return error;
}

async function readJson(response) {
  const text = await response.text();
  try {
    return text ? JSON.parse(text) : {};
  } catch {
    return { raw: text };
  }
}

export function getGoogleAuthUrl({
  clientId,
  redirectUri,
  state,
  scope,
}) {
  const url = new URL(GOOGLE_AUTH_URL);
  url.searchParams.set('client_id', clientId);
  url.searchParams.set('redirect_uri', redirectUri);
  url.searchParams.set('response_type', 'code');
  url.searchParams.set('scope', scope);
  url.searchParams.set('state', state);
  url.searchParams.set('access_type', 'offline');
  url.searchParams.set('prompt', 'consent');
  url.searchParams.set('include_granted_scopes', 'true');
  return url.toString();
}

export async function exchangeCodeForToken({
  clientId,
  clientSecret,
  redirectUri,
  code,
}) {
  const body = new URLSearchParams({
    client_id: clientId,
    client_secret: clientSecret,
    code,
    grant_type: 'authorization_code',
    redirect_uri: redirectUri,
  });

  const response = await fetch(GOOGLE_TOKEN_URL, {
    method: 'POST',
    headers: { 'content-type': 'application/x-www-form-urlencoded' },
    body,
  });

  const payload = await readJson(response);
  if (!response.ok) {
    throw toError(payload.error_description || payload.error || 'Google token exchange failed', response.status, payload);
  }
  return payload;
}

export async function refreshAccessToken({
  clientId,
  clientSecret,
  refreshToken,
}) {
  const body = new URLSearchParams({
    client_id: clientId,
    client_secret: clientSecret,
    grant_type: 'refresh_token',
    refresh_token: refreshToken,
  });

  const response = await fetch(GOOGLE_TOKEN_URL, {
    method: 'POST',
    headers: { 'content-type': 'application/x-www-form-urlencoded' },
    body,
  });

  const payload = await readJson(response);
  if (!response.ok) {
    throw toError(payload.error_description || payload.error || 'Google token refresh failed', response.status, payload);
  }
  return payload;
}

export async function fetchGoogleUserInfo(accessToken) {
  const response = await fetch(GOOGLE_USERINFO_URL, {
    headers: { authorization: `Bearer ${accessToken}` },
  });
  const payload = await readJson(response);
  if (!response.ok) {
    throw toError(payload.error_description || payload.error || 'Unable to load Google profile', response.status, payload);
  }
  return payload;
}

export async function fetchSpreadsheetMetadata({ spreadsheetId, accessToken }) {
  const response = await fetch(
    `${GOOGLE_SHEETS_BASE}/${encodeURIComponent(spreadsheetId)}?fields=spreadsheetId,properties.title,sheets.properties`,
    { headers: { authorization: `Bearer ${accessToken}` } },
  );
  const payload = await readJson(response);
  if (!response.ok) {
    throw toError(payload.error?.message || payload.error_description || 'Unable to access spreadsheet', response.status, payload);
  }
  return payload;
}

export async function fetchSheetValues({ spreadsheetId, sheetName, accessToken }) {
  const range = encodeURIComponent(`${sheetName}`);
  const response = await fetch(
    `${GOOGLE_SHEETS_BASE}/${encodeURIComponent(spreadsheetId)}/values/${range}?majorDimension=ROWS`,
    { headers: { authorization: `Bearer ${accessToken}` } },
  );
  const payload = await readJson(response);
  if (!response.ok) {
    throw toError(payload.error?.message || payload.error_description || 'Unable to read sheet values', response.status, payload);
  }
  return payload.values || [];
}

export async function appendSheetRow({
  spreadsheetId,
  sheetName,
  accessToken,
  values,
}) {
  const range = encodeURIComponent(`${sheetName}`);
  const response = await fetch(
    `${GOOGLE_SHEETS_BASE}/${encodeURIComponent(spreadsheetId)}/values/${range}:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`,
    {
      method: 'POST',
      headers: {
        authorization: `Bearer ${accessToken}`,
        'content-type': 'application/json',
      },
      body: JSON.stringify({
        majorDimension: 'ROWS',
        values: [values],
      }),
    },
  );
  const payload = await readJson(response);
  if (!response.ok) {
    throw toError(payload.error?.message || payload.error_description || 'Unable to append row', response.status, payload);
  }
  return payload;
}

export async function updateSheetRow({
  spreadsheetId,
  sheetName,
  accessToken,
  rowNumber,
  values,
  columnCount,
}) {
  const endColumn = columnNumberToName(columnCount);
  const range = encodeURIComponent(`${sheetName}!A${rowNumber}:${endColumn}${rowNumber}`);
  const response = await fetch(
    `${GOOGLE_SHEETS_BASE}/${encodeURIComponent(spreadsheetId)}/values/${range}?valueInputOption=USER_ENTERED`,
    {
      method: 'PUT',
      headers: {
        authorization: `Bearer ${accessToken}`,
        'content-type': 'application/json',
      },
      body: JSON.stringify({
        majorDimension: 'ROWS',
        values: [values],
      }),
    },
  );
  const payload = await readJson(response);
  if (!response.ok) {
    throw toError(payload.error?.message || payload.error_description || 'Unable to update row', response.status, payload);
  }
  return payload;
}

export async function deleteSheetRow({
  spreadsheetId,
  sheetId,
  accessToken,
  rowNumber,
}) {
  const response = await fetch(`${GOOGLE_SHEETS_BASE}/${encodeURIComponent(spreadsheetId)}:batchUpdate`, {
    method: 'POST',
    headers: {
      authorization: `Bearer ${accessToken}`,
      'content-type': 'application/json',
    },
    body: JSON.stringify({
      requests: [{
        deleteDimension: {
          range: {
            sheetId,
            dimension: 'ROWS',
            startIndex: rowNumber - 1,
            endIndex: rowNumber,
          },
        },
      }],
    }),
  });
  const payload = await readJson(response);
  if (!response.ok) {
    throw toError(payload.error?.message || payload.error_description || 'Unable to delete row', response.status, payload);
  }
  return payload;
}

export function columnNumberToName(number) {
  let n = Math.max(1, Number(number) || 1);
  let result = '';
  while (n > 0) {
    const remainder = (n - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result;
}

export function parseSheetTable(values) {
  const rawRows = Array.isArray(values) ? values : [];
  const headerRowIndex = rawRows.findIndex((row) => Array.isArray(row) && row.some((cell) => String(cell ?? '').trim() !== ''));
  if (headerRowIndex === -1) {
    return { headers: [], rows: [], headerRowNumber: 1 };
  }

  const headers = rawRows[headerRowIndex].map((cell, index) => {
    const label = String(cell ?? '').trim();
    return label || `Column ${index + 1}`;
  });

  const rows = rawRows.slice(headerRowIndex + 1).map((row, index) => {
    const normalized = {};
    headers.forEach((header, headerIndex) => {
      normalized[header] = row?.[headerIndex] ?? '';
    });
    return {
      id: `${headerRowIndex + 2 + index}`,
      rowNumber: headerRowIndex + 2 + index,
      values: normalized,
    };
  }).filter((entry) => Object.values(entry.values).some((value) => String(value ?? '').trim() !== ''));

  return { headers, rows, headerRowNumber: headerRowIndex + 1 };
}

export function buildRowValues(headers, rowValues) {
  return headers.map((header) => {
    const value = rowValues[header];
    return value == null ? '' : value;
  });
}
