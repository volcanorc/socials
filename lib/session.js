import crypto from 'node:crypto';

const SESSION_COOKIE = 'pd_session';
const OAUTH_STATE_COOKIE = 'pd_oauth_state';

function keyFromSecret(secret) {
  return crypto.createHash('sha256').update(String(secret)).digest();
}

function base64urlEncode(input) {
  return Buffer.from(input).toString('base64url');
}

function base64urlDecode(input) {
  return Buffer.from(input, 'base64url');
}

export function parseCookies(cookieHeader = '') {
  return cookieHeader.split(';').reduce((acc, part) => {
    const index = part.indexOf('=');
    if (index === -1) return acc;
    const name = part.slice(0, index).trim();
    const value = part.slice(index + 1).trim();
    if (name) acc[name] = decodeURIComponent(value);
    return acc;
  }, {});
}

export function serializeCookie(name, value, options = {}) {
  const parts = [`${name}=${encodeURIComponent(value)}`];
  if (options.maxAge != null) parts.push(`Max-Age=${Math.floor(options.maxAge)}`);
  if (options.domain) parts.push(`Domain=${options.domain}`);
  parts.push(`Path=${options.path || '/'}`);
  if (options.httpOnly !== false) parts.push('HttpOnly');
  if (options.sameSite) parts.push(`SameSite=${options.sameSite}`);
  if (options.secure) parts.push('Secure');
  return parts.join('; ');
}

function encryptObject(secret, value) {
  const key = keyFromSecret(secret);
  const iv = crypto.randomBytes(12);
  const cipher = crypto.createCipheriv('aes-256-gcm', key, iv);
  const plaintext = Buffer.from(JSON.stringify(value), 'utf8');
  const encrypted = Buffer.concat([cipher.update(plaintext), cipher.final()]);
  const tag = cipher.getAuthTag();
  return `${base64urlEncode(iv)}.${base64urlEncode(encrypted)}.${base64urlEncode(tag)}`;
}

function decryptObject(secret, packed) {
  const [ivPart, dataPart, tagPart] = String(packed).split('.');
  if (!ivPart || !dataPart || !tagPart) return null;
  const key = keyFromSecret(secret);
  const decipher = crypto.createDecipheriv('aes-256-gcm', key, base64urlDecode(ivPart));
  decipher.setAuthTag(base64urlDecode(tagPart));
  const decrypted = Buffer.concat([
    decipher.update(base64urlDecode(dataPart)),
    decipher.final(),
  ]);
  return JSON.parse(decrypted.toString('utf8'));
}

export function createSessionCookie(secret, session) {
  return encryptObject(secret, session);
}

export function readSessionCookie(secret, cookieHeader) {
  const cookies = parseCookies(cookieHeader);
  if (!cookies[SESSION_COOKIE]) return null;
  try {
    return decryptObject(secret, cookies[SESSION_COOKIE]);
  } catch {
    return null;
  }
}

export function buildSessionSetCookie(secret, session, { secure = false, maxAge = 60 * 60 * 24 * 30 } = {}) {
  const value = createSessionCookie(secret, session);
  return serializeCookie(SESSION_COOKIE, value, {
    httpOnly: true,
    sameSite: 'Lax',
    secure,
    maxAge,
  });
}

export function clearSessionCookie({ secure = false } = {}) {
  return serializeCookie(SESSION_COOKIE, '', {
    httpOnly: true,
    sameSite: 'Lax',
    secure,
    maxAge: 0,
  });
}

export function buildOAuthStateCookie(state, { secure = false, maxAge = 600 } = {}) {
  return serializeCookie(OAUTH_STATE_COOKIE, state, {
    httpOnly: true,
    sameSite: 'Lax',
    secure,
    maxAge,
  });
}

export function clearOAuthStateCookie({ secure = false } = {}) {
  return serializeCookie(OAUTH_STATE_COOKIE, '', {
    httpOnly: true,
    sameSite: 'Lax',
    secure,
    maxAge: 0,
  });
}

export function readOAuthState(cookieHeader) {
  return parseCookies(cookieHeader)[OAUTH_STATE_COOKIE] || null;
}

export function createRandomState() {
  return crypto.randomBytes(24).toString('base64url');
}
