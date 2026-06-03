import {
  appendSheetRow,
  deleteSheetRow,
  exchangeCodeForToken,
  fetchGoogleUserInfo,
  fetchSheetValues,
  fetchSpreadsheetMetadata,
  refreshAccessToken,
  updateSheetRow,
  parseSheetTable,
  buildRowValues,
} from './google.js';

export async function ensureAuthorizedSession({ session, env, setSessionCookie }) {
  if (!session) {
    const error = new Error('Not signed in');
    error.status = 401;
    throw error;
  }

  let activeSession = session;
  const now = Date.now();
  if (activeSession.expiresAt && now >= activeSession.expiresAt - 60_000) {
    if (!activeSession.refreshToken) {
      const error = new Error('Session expired');
      error.status = 401;
      throw error;
    }
    const refreshed = await refreshAccessToken({
      clientId: env.GOOGLE_CLIENT_ID,
      clientSecret: env.GOOGLE_CLIENT_SECRET,
      refreshToken: activeSession.refreshToken,
    });
    activeSession = {
      ...activeSession,
      accessToken: refreshed.access_token,
      expiresAt: now + (Number(refreshed.expires_in || 3600) * 1000),
      refreshToken: refreshed.refresh_token || activeSession.refreshToken,
    };
    if (setSessionCookie) setSessionCookie(activeSession);
  }

  return activeSession;
}

export async function requireSpreadsheetAccess({ session, env, setSessionCookie }) {
  const activeSession = await ensureAuthorizedSession({ session, env, setSessionCookie });
  const metadata = await fetchSpreadsheetMetadata({
    spreadsheetId: env.SPREADSHEET_ID,
    accessToken: activeSession.accessToken,
  });
  return { session: activeSession, metadata };
}

export async function loadSheetData({ session, env, setSessionCookie, sheetName }) {
  const { session: activeSession, metadata } = await requireSpreadsheetAccess({ session, env, setSessionCookie });
  const values = await fetchSheetValues({
    spreadsheetId: env.SPREADSHEET_ID,
    sheetName,
    accessToken: activeSession.accessToken,
  });
  const table = parseSheetTable(values);
  const sheet = metadata.sheets?.find((entry) => entry.properties?.title === sheetName);
  return {
    spreadsheetTitle: metadata.properties?.title || 'Private spreadsheet',
    sheetId: sheet?.properties?.sheetId ?? null,
    sheetName,
    headers: table.headers,
    rows: table.rows,
  };
}

export async function addSheetRow({ session, env, setSessionCookie, sheetName, row }) {
  const { session: activeSession, metadata } = await requireSpreadsheetAccess({ session, env, setSessionCookie });
  const values = await fetchSheetValues({
    spreadsheetId: env.SPREADSHEET_ID,
    sheetName,
    accessToken: activeSession.accessToken,
  });
  const table = parseSheetTable(values);
  if (!table.headers.length) {
    throw Object.assign(new Error('This sheet needs a header row before rows can be added.'), { status: 400 });
  }
  const orderedValues = buildRowValues(table.headers, row);
  await appendSheetRow({
    spreadsheetId: env.SPREADSHEET_ID,
    sheetName,
    accessToken: activeSession.accessToken,
    values: orderedValues,
  });
  return loadSheetData({ session: activeSession, env, setSessionCookie, sheetName });
}

export async function updateSheetEntry({ session, env, setSessionCookie, sheetName, rowNumber, row }) {
  const { session: activeSession } = await requireSpreadsheetAccess({ session, env, setSessionCookie });
  const values = await fetchSheetValues({
    spreadsheetId: env.SPREADSHEET_ID,
    sheetName,
    accessToken: activeSession.accessToken,
  });
  const table = parseSheetTable(values);
  const orderedValues = buildRowValues(table.headers, row);
  await updateSheetRow({
    spreadsheetId: env.SPREADSHEET_ID,
    sheetName,
    accessToken: activeSession.accessToken,
    rowNumber,
    values: orderedValues,
    columnCount: Math.max(table.headers.length, orderedValues.length, 1),
  });
  return loadSheetData({ session: activeSession, env, setSessionCookie, sheetName });
}

export async function removeSheetRow({ session, env, setSessionCookie, sheetName, rowNumber }) {
  const { session: activeSession, metadata } = await requireSpreadsheetAccess({ session, env, setSessionCookie });
  const sheet = metadata.sheets?.find((entry) => entry.properties?.title === sheetName);
  if (!sheet?.properties?.sheetId && sheet?.properties?.sheetId !== 0) {
    throw Object.assign(new Error('Sheet not found'), { status: 404 });
  }
  await deleteSheetRow({
    spreadsheetId: env.SPREADSHEET_ID,
    sheetId: sheet.properties.sheetId,
    accessToken: activeSession.accessToken,
    rowNumber,
  });
  return loadSheetData({ session: activeSession, env, setSessionCookie, sheetName });
}

export async function completeGoogleLogin({ env, code, redirectUri }) {
  const token = await exchangeCodeForToken({
    clientId: env.GOOGLE_CLIENT_ID,
    clientSecret: env.GOOGLE_CLIENT_SECRET,
    redirectUri: redirectUri || new URL(env.GOOGLE_REDIRECT_URI).origin,
    code,
  });
  const profile = await fetchGoogleUserInfo(token.access_token);
  return {
    accessToken: token.access_token,
    refreshToken: token.refresh_token || null,
    expiresAt: Date.now() + (Number(token.expires_in || 3600) * 1000),
    profile: {
      id: profile.sub,
      email: profile.email,
      name: profile.name || profile.email,
      picture: profile.picture || '',
    },
  };
}
