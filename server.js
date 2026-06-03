import { createServer } from 'node:http';
import { existsSync, readFileSync } from 'node:fs';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import crypto from 'node:crypto';
import { fileURLToPath } from 'node:url';

import {
  buildOAuthStateCookie,
  buildSessionSetCookie,
  clearOAuthStateCookie,
  clearSessionCookie,
  createRandomState,
  readOAuthState,
  readSessionCookie,
} from './lib/session.js';
import {
  completeGoogleLogin,
  loadSheetData,
  removeSheetRow,
  requireSpreadsheetAccess,
  updateSheetEntry,
} from './lib/google-access.js';
import { appendSheetRow, buildRowValues } from './lib/google.js';

function loadDotEnv() {
  const envPath = path.join(process.cwd(), '.env');
  if (!existsSync(envPath)) return;
  const content = readFileSync(envPath, 'utf8');
  for (const rawLine of content.split(/\r?\n/)) {
    const line = rawLine.trim();
    if (!line || line.startsWith('#')) continue;
    const equalsIndex = line.indexOf('=');
    if (equalsIndex === -1) continue;
    const key = line.slice(0, equalsIndex).trim();
    let value = line.slice(equalsIndex + 1).trim();
    if ((value.startsWith('"') && value.endsWith('"')) || (value.startsWith("'") && value.endsWith("'"))) {
      value = value.slice(1, -1);
    }
    if (key && !(key in process.env)) {
      process.env[key] = value;
    }
  }
}

loadDotEnv();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const publicDir = path.join(__dirname, 'public');

const env = {
  PORT: Number(process.env.PORT || 3000),
  SESSION_SECRET: process.env.SESSION_SECRET || 'change-me-in-production',
  GOOGLE_CLIENT_ID: process.env.GOOGLE_CLIENT_ID || '',
  GOOGLE_CLIENT_SECRET: process.env.GOOGLE_CLIENT_SECRET || '',
  GOOGLE_REDIRECT_URI: process.env.GOOGLE_REDIRECT_URI || `http://localhost:${Number(process.env.PORT || 3000)}`,
  COOKIE_SECURE: String(process.env.COOKIE_SECURE || '').toLowerCase() === 'true',
  SPREADSHEET_ID: process.env.SPREADSHEET_ID || '1O2bRxJpxJRBVgv676W5oT0apVkHP_lWYYaDcG1axeCU',
};

const runtime = {
  sessionSecret: env.SESSION_SECRET,
};

function isConfigured() {
  return Boolean(env.GOOGLE_CLIENT_ID && env.GOOGLE_CLIENT_SECRET && env.GOOGLE_REDIRECT_URI && env.SPREADSHEET_ID);
}

function setCookieHeaders(res, cookies) {
  if (!cookies.length) return;
  const existing = res.getHeader('Set-Cookie');
  const merged = Array.isArray(existing) ? [...existing, ...cookies] : existing ? [existing, ...cookies] : cookies;
  res.setHeader('Set-Cookie', merged);
}

function sendJson(res, status, payload, cookies = []) {
  setCookieHeaders(res, cookies);
  res.writeHead(status, {
    'content-type': 'application/json; charset=utf-8',
    'cache-control': 'no-store',
  });
  res.end(JSON.stringify(payload));
}

function sendText(res, status, text, contentType = 'text/plain; charset=utf-8', cookies = []) {
  setCookieHeaders(res, cookies);
  res.writeHead(status, {
    'content-type': contentType,
    'cache-control': 'no-store',
  });
  res.end(text);
}

async function serveStatic(res, fileName, contentType) {
  try {
    const file = await readFile(path.join(publicDir, fileName));
    res.writeHead(200, {
      'content-type': contentType,
      'cache-control': 'no-store',
    });
    res.end(file);
  } catch {
    sendText(res, 404, 'Not found');
  }
}

function buildOAuthConfig() {
  const state = createRandomState();
  return {
    clientId: env.GOOGLE_CLIENT_ID,
    scope: 'openid email profile https://www.googleapis.com/auth/spreadsheets',
    state,
  };
}

async function readBody(req) {
  const chunks = [];
  for await (const chunk of req) {
    chunks.push(chunk);
    if (Buffer.concat(chunks).length > 1_000_000) {
      throw Object.assign(new Error('Request body too large'), { status: 413 });
    }
  }
  if (!chunks.length) return null;
  const text = Buffer.concat(chunks).toString('utf8');
  try {
    return JSON.parse(text);
  } catch {
    throw Object.assign(new Error('Invalid JSON body'), { status: 400 });
  }
}

function getRequestedOrigin(req) {
  return req.headers.origin || `http://${req.headers.host}`;
}

function isCookieSecure(req) {
  if (env.COOKIE_SECURE) return true;
  const forwardedProto = req.headers['x-forwarded-proto'];
  return forwardedProto === 'https';
}

function getSessionFromRequest(req) {
  return readSessionCookie(runtime.sessionSecret, req.headers.cookie || '');
}

function setSessionCookieForResponse(req, res, session, cookies) {
  cookies.push(buildSessionSetCookie(runtime.sessionSecret, session, {
    secure: isCookieSecure(req),
  }));
}

async function resolveSession(req, res, { refresh = true } = {}) {
  const cookies = [];
  const stored = getSessionFromRequest(req);
  if (!stored) return { session: null, cookies };

  try {
    const active = refresh
      ? await requireSpreadsheetAccess({ session: stored, env, setSessionCookie: (session) => setSessionCookieForResponse(req, res, session, cookies) })
      : { session: stored };
    return { session: active.session, cookies };
  } catch (error) {
    cookies.push(clearSessionCookie({ secure: isCookieSecure(req) }));
    return { session: null, cookies, error };
  }
}

async function loadSpreadsheetOverview(session, cookies, req, res) {
  const { session: activeSession, metadata } = await requireSpreadsheetAccess({
    session,
    env,
    setSessionCookie: (nextSession) => setSessionCookieForResponse(req, res, nextSession, cookies),
  });
  const sheets = (metadata.sheets || []).map((sheet) => ({
    title: sheet.properties?.title || 'Untitled',
    sheetId: sheet.properties?.sheetId ?? null,
  }));
  return {
    session: activeSession,
    spreadsheetTitle: metadata.properties?.title || 'Private spreadsheet',
    sheets,
  };
}

function findSheetTitle(metadata, sheetName) {
  return (metadata.sheets || []).find((sheet) => sheet.properties?.title === sheetName) || null;
}

async function handleAuthConfig(req, res) {
  if (!isConfigured()) {
    return sendJson(res, 500, {
      error: 'Google OAuth is not configured. Set GOOGLE_CLIENT_ID, GOOGLE_CLIENT_SECRET, and GOOGLE_REDIRECT_URI.',
    });
  }
  const config = buildOAuthConfig();
  const cookies = [buildOAuthStateCookie(config.state, { secure: isCookieSecure(req) })];
  return sendJson(res, 200, config, cookies);
}

async function handleAuthCallback(req, res, url) {
  if (!isConfigured()) {
    return sendText(res, 500, 'Google OAuth is not configured.');
  }

  const cookies = [];
  const stateFromCookie = readOAuthState(req.headers.cookie || '');
  let code = url.searchParams.get('code');
  let state = url.searchParams.get('state');

  if (req.method === 'POST') {
    const requestedWith = req.headers['x-requested-with'];
    const origin = req.headers.origin;
    const expectedOrigin = new URL(getRequestedOrigin(req)).origin;
    if (requestedWith !== 'XMLHttpRequest' || (origin && new URL(origin).origin !== expectedOrigin)) {
      cookies.push(clearOAuthStateCookie({ secure: isCookieSecure(req) }));
      return sendJson(res, 403, { error: 'invalid_popup_request' }, cookies);
    }
    const body = await readBody(req);
    code = body?.code || code;
    state = body?.state || state;
  }

  const storedState = readOAuthState(req.headers.cookie || '');
  if (!state || !code || !storedState || state !== storedState || (stateFromCookie && storedState !== stateFromCookie)) {
    cookies.push(clearOAuthStateCookie({ secure: isCookieSecure(req) }));
    return sendJson(res, 400, { error: 'invalid_oauth_response' }, cookies);
  }

  try {
    const login = await completeGoogleLogin({
      env,
      code,
      redirectUri: new URL(getRequestedOrigin(req)).origin,
    });
    const permissionCheck = await requireSpreadsheetAccess({
      session: login,
      env,
      setSessionCookie: () => {},
    });

    cookies.push(clearOAuthStateCookie({ secure: isCookieSecure(req) }));
    cookies.push(buildSessionSetCookie(runtime.sessionSecret, permissionCheck.session, {
      secure: isCookieSecure(req),
    }));
    if (req.method === 'POST') {
      return sendJson(res, 200, {
        ok: true,
        profile: permissionCheck.session.profile,
      }, cookies);
    }

    setCookieHeaders(res, cookies);
    res.writeHead(302, { location: '/?signedIn=1' });
    res.end();
  } catch (error) {
    cookies.push(clearOAuthStateCookie({ secure: isCookieSecure(req) }));
    const denied = error?.status === 403 || error?.status === 404;
    if (req.method === 'POST') {
      return sendJson(res, denied ? 403 : 500, {
        error: denied ? 'permission_denied' : 'login_failed',
      }, cookies);
    }
    setCookieHeaders(res, cookies);
    res.writeHead(302, {
      location: denied ? '/?error=permission-denied' : '/?error=login-failed',
    });
    res.end();
  }
}

async function handleLogout(req, res) {
  const cookies = [
    clearSessionCookie({ secure: isCookieSecure(req) }),
    clearOAuthStateCookie({ secure: isCookieSecure(req) }),
  ];
  if (req.method === 'POST') {
    return sendJson(res, 200, { ok: true }, cookies);
  }
  setCookieHeaders(res, cookies);
  res.writeHead(302, { location: '/' });
  res.end();
}

async function handleSession(req, res) {
  const cookies = [];
  const { session } = await resolveSession(req, res, { refresh: true });
  if (!session) {
    return sendJson(res, 200, {
      authenticated: false,
      configured: isConfigured(),
    }, cookies);
  }
  return sendJson(res, 200, {
    authenticated: true,
    configured: isConfigured(),
    profile: session.profile,
    expiresAt: session.expiresAt,
  }, cookies);
}

async function handleSpreadsheet(req, res) {
  const cookies = [];
  const { session } = await resolveSession(req, res, { refresh: true });
  if (!session) return sendJson(res, 401, { error: 'not_authenticated' }, cookies);

  try {
    const overview = await loadSpreadsheetOverview(session, cookies, req, res);
    return sendJson(res, 200, overview, cookies);
  } catch (error) {
    const status = error?.status === 403 || error?.status === 404 ? 403 : error?.status || 500;
    return sendJson(res, status, {
      error: status === 403 ? 'permission_denied' : 'Unable to load spreadsheet metadata',
    }, cookies);
  }
}

async function handleSheetRead(req, res, url) {
  const sheetName = url.searchParams.get('sheet');
  const cookies = [];
  if (!sheetName) return sendJson(res, 400, { error: 'sheet is required' }, cookies);

  const { session } = await resolveSession(req, res, { refresh: true });
  if (!session) return sendJson(res, 401, { error: 'not_authenticated' }, cookies);

  try {
    const data = await loadSheetData({
      session,
      env,
      setSessionCookie: (nextSession) => setSessionCookieForResponse(req, res, nextSession, cookies),
      sheetName,
    });
    return sendJson(res, 200, data, cookies);
  } catch (error) {
    const status = error?.status === 403 || error?.status === 404 ? 403 : error?.status || 500;
    return sendJson(res, status, {
      error: status === 403 ? 'permission_denied' : error.message || 'Unable to load sheet',
    }, cookies);
  }
}

async function handleSheetCreate(req, res, url) {
  const sheetName = url.searchParams.get('sheet');
  const body = await readBody(req);
  const cookies = [];
  if (!sheetName) return sendJson(res, 400, { error: 'sheet is required' }, cookies);

  const { session } = await resolveSession(req, res, { refresh: true });
  if (!session) return sendJson(res, 401, { error: 'not_authenticated' }, cookies);

  try {
    const data = await loadSheetData({
      session,
      env,
      setSessionCookie: (nextSession) => setSessionCookieForResponse(req, res, nextSession, cookies),
      sheetName,
    });
    const ordered = buildRowValues(data.headers, body?.values || {});
    await appendSheetRow({
      spreadsheetId: env.SPREADSHEET_ID,
      sheetName,
      accessToken: session.accessToken,
      values: ordered,
    });
    const refreshed = await loadSheetData({
      session,
      env,
      setSessionCookie: (nextSession) => setSessionCookieForResponse(req, res, nextSession, cookies),
      sheetName,
    });
    return sendJson(res, 200, refreshed, cookies);
  } catch (error) {
    return sendJson(res, error?.status || 500, { error: error.message || 'Unable to create row' }, cookies);
  }
}

async function handleSheetUpdate(req, res, url) {
  const sheetName = url.searchParams.get('sheet');
  const rowNumber = Number(url.searchParams.get('rowNumber'));
  const body = await readBody(req);
  const cookies = [];
  if (!sheetName || !Number.isFinite(rowNumber)) {
    return sendJson(res, 400, { error: 'sheet and rowNumber are required' }, cookies);
  }

  const { session } = await resolveSession(req, res, { refresh: true });
  if (!session) return sendJson(res, 401, { error: 'not_authenticated' }, cookies);

  try {
    await updateSheetEntry({
      session,
      env,
      setSessionCookie: (nextSession) => setSessionCookieForResponse(req, res, nextSession, cookies),
      sheetName,
      rowNumber,
      row: body?.values || {},
    });
    const refreshed = await loadSheetData({
      session,
      env,
      setSessionCookie: (nextSession) => setSessionCookieForResponse(req, res, nextSession, cookies),
      sheetName,
    });
    return sendJson(res, 200, refreshed, cookies);
  } catch (error) {
    return sendJson(res, error?.status || 500, { error: error.message || 'Unable to update row' }, cookies);
  }
}

async function handleSheetDelete(req, res, url) {
  const sheetName = url.searchParams.get('sheet');
  const rowNumber = Number(url.searchParams.get('rowNumber'));
  const cookies = [];
  if (!sheetName || !Number.isFinite(rowNumber)) {
    return sendJson(res, 400, { error: 'sheet and rowNumber are required' }, cookies);
  }

  const { session } = await resolveSession(req, res, { refresh: true });
  if (!session) return sendJson(res, 401, { error: 'not_authenticated' }, cookies);

  try {
    await removeSheetRow({
      session,
      env,
      setSessionCookie: (nextSession) => setSessionCookieForResponse(req, res, nextSession, cookies),
      sheetName,
      rowNumber,
    });
    const refreshed = await loadSheetData({
      session,
      env,
      setSessionCookie: (nextSession) => setSessionCookieForResponse(req, res, nextSession, cookies),
      sheetName,
    });
    return sendJson(res, 200, refreshed, cookies);
  } catch (error) {
    return sendJson(res, error?.status || 500, { error: error.message || 'Unable to delete row' }, cookies);
  }
}

async function handleRequest(req, res) {
  const url = new URL(req.url, getRequestedOrigin(req));

  try {
    if (url.pathname === '/auth/callback') return handleAuthCallback(req, res, url);
    if (url.pathname === '/api/auth/config' && req.method === 'GET') return handleAuthConfig(req, res);
    if (url.pathname === '/auth/logout') return handleLogout(req, res);

    if (url.pathname === '/api/session') return handleSession(req, res);
    if (url.pathname === '/api/spreadsheet') return handleSpreadsheet(req, res);
    if (url.pathname === '/api/sheet' && req.method === 'GET') return handleSheetRead(req, res, url);
    if (url.pathname === '/api/sheet' && req.method === 'POST') return handleSheetCreate(req, res, url);
    if (url.pathname === '/api/sheet' && req.method === 'PUT') return handleSheetUpdate(req, res, url);
    if (url.pathname === '/api/sheet' && req.method === 'DELETE') return handleSheetDelete(req, res, url);

    if (url.pathname === '/health') return sendJson(res, 200, { ok: true, configured: isConfigured() });

    if (url.pathname === '/' || url.pathname === '/index.html') {
      return serveStatic(res, 'index.html', 'text/html; charset=utf-8');
    }
    if (url.pathname === '/styles.css') {
      return serveStatic(res, 'styles.css', 'text/css; charset=utf-8');
    }
    if (url.pathname === '/app.js') {
      return serveStatic(res, 'app.js', 'text/javascript; charset=utf-8');
    }

    return sendText(res, 404, 'Not found');
  } catch (error) {
    const status = error?.status || 500;
    return sendJson(res, status, {
      error: error?.message || 'Unexpected server error',
    });
  }
}

const server = createServer((req, res) => {
  handleRequest(req, res);
});

server.listen(env.PORT, () => {
  console.log(`Secure dashboard running on http://localhost:${env.PORT}`);
});
