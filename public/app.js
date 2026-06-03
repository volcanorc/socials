const APP_CONFIG = window.APP_CONFIG || {};
const APPS_SCRIPT_URL = APP_CONFIG.appsScriptUrl || '';
const GOOGLE_CLIENT_ID = APP_CONFIG.googleClientId || '';
const AUTH_SCOPE = APP_CONFIG.authScope || 'openid email profile https://www.googleapis.com/auth/spreadsheets';
const FORWARD_AUTH_TOKEN = Boolean(APP_CONFIG.forwardAuthToken);

const SESSION_EXPIRY_SKEW_MS = 60_000;
const SAVE_DEBOUNCE_MS = 600;

const state = {
  session: null,
  authReady: false,
  matrix: [],
  lastSavedMatrix: [],
  activeSheetName: '',
  loading: false,
  selectedCell: null,
  searchQuery: '',
  notice: '',
  noticeKind: 'info',
  saveTimers: new Map(),
  dirtyCells: new Set(),
  savingCells: new Set(),
  cellErrors: new Map(),
  lastLoadedAt: null,
  viewMode: 'auth',
  accessDeniedMessage: '',
};

const elements = {};
let sessionExpiryTimer = null;

window.userToken = '';

function $(id) {
  return document.getElementById(id);
}

function escapeHtml(value) {
  return String(value ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

function cloneMatrix(matrix) {
  return (Array.isArray(matrix) ? matrix : []).map((row) => (
    Array.isArray(row) ? row.map((cell) => String(cell ?? '')) : ['']
  ));
}

function normalizeMatrix(payload) {
  const raw = Array.isArray(payload?.data)
    ? payload.data
    : Array.isArray(payload?.values)
      ? payload.values
      : Array.isArray(payload)
        ? payload
      : null;

  if (!raw) {
    throw Object.assign(new Error('Apps Script response must be a JSON matrix.'), { status: 500 });
  }

  const rows = raw.map((row) => (Array.isArray(row) ? row.map((cell) => String(cell ?? '')) : ['']));
  const maxColumns = Math.max(1, ...rows.map((row) => row.length));
  return rows.map((row) => {
    const normalized = row.slice(0, maxColumns);
    while (normalized.length < maxColumns) normalized.push('');
    return normalized;
  });
}

function extractMatrixData(payload) {
  if (Array.isArray(payload?.data)) return payload.data;
  if (Array.isArray(payload?.values)) return payload.values;
  if (Array.isArray(payload)) return payload;
  return null;
}

function keyFor(row, col) {
  return `${row}:${col}`;
}

function columnLabel(index) {
  let n = index;
  let result = '';
  while (n > 0) {
    const remainder = (n - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result;
}

function isSessionExpired(session) {
  return !session?.credential || !Number.isFinite(Number(session?.expiresAt)) || Date.now() >= Number(session.expiresAt) - SESSION_EXPIRY_SKEW_MS;
}

function decodeBase64Url(value) {
  const normalized = String(value || '').replaceAll('-', '+').replaceAll('_', '/');
  const padded = normalized.padEnd(Math.ceil(normalized.length / 4) * 4, '=');
  const binary = window.atob(padded);
  const bytes = Uint8Array.from(binary, (char) => char.charCodeAt(0));
  return new TextDecoder().decode(bytes);
}

function parseCredentialPayload(credential) {
  const parts = String(credential || '').split('.');
  if (parts.length < 2) return null;

  try {
    return JSON.parse(decodeBase64Url(parts[1]));
  } catch {
    return null;
  }
}

function buildSessionFromCredential(credential) {
  const payload = parseCredentialPayload(credential);
  const expiresAt = Number(payload?.exp || 0) * 1000 || (Date.now() + 55 * 60 * 1000);

  return {
    authenticated: true,
    credential,
    expiresAt,
    profile: {
      id: String(payload?.sub || ''),
      email: String(payload?.email || ''),
      name: String(payload?.name || payload?.email || ''),
      picture: String(payload?.picture || ''),
    },
  };
}

function clearSessionExpiryTimer() {
  if (sessionExpiryTimer) {
    clearTimeout(sessionExpiryTimer);
    sessionExpiryTimer = null;
  }
}

function setAppShellVisibility(isAuthenticated) {
  if (elements.authPanel) {
    elements.authPanel.classList.toggle('hidden', isAuthenticated);
  }
  if (elements.dashboard) {
    elements.dashboard.classList.toggle('hidden', !isAuthenticated);
  }
}

function setWorkspaceVisibility(isVisible) {
  if (elements.workspace) {
    elements.workspace.classList.toggle('hidden', !isVisible);
  }
}

function clearSheetViewport() {
  if (elements.dataTable) {
    const thead = elements.dataTable.querySelector('thead');
    const tbody = elements.dataTable.querySelector('tbody');
    if (thead) thead.innerHTML = '';
    if (tbody) tbody.innerHTML = '';
  }

  if (elements.rowCount) {
    elements.rowCount.textContent = '0 rows visible';
  }

  if (elements.spreadsheetTitle) {
    elements.spreadsheetTitle.textContent = 'Sign in to view the spreadsheet';
  }

  if (elements.sheetMeta) {
    elements.sheetMeta.textContent = 'No spreadsheet loaded.';
  }

  if (elements.sheetTitle) {
    elements.sheetTitle.textContent = 'Native spreadsheet grid';
  }

  if (elements.sheetSummary) {
    elements.sheetSummary.textContent = 'Sign in with Google to load the live sheet.';
  }

  if (elements.searchInput) {
    elements.searchInput.value = '';
  }

  if (elements.detailStatus) {
    elements.detailStatus.innerHTML = `
      <div><strong>Editing mode:</strong> Sign in to load the spreadsheet.</div>
      <div><strong>Tip:</strong> The grid only appears after a valid Google session is restored or created.</div>
    `;
  }

  state.matrix = [];
  state.lastSavedMatrix = [];
  state.selectedCell = null;
  state.searchQuery = '';
  state.lastLoadedAt = null;
  clearSaveTimers();
  resetCellState();
}

function showAuthOnlyState(message = '', kind = 'info') {
  state.viewMode = 'auth';
  state.accessDeniedMessage = '';
  state.activeSheetName = '';
  setAppShellVisibility(false);
  setWorkspaceVisibility(false);
  clearSheetViewport();
  renderSessionBox();
  renderConnectionCard();
  renderHeaderMeta();
  renderDetailPanel();
  if (message) {
    setNotice(message, kind);
  }
}

function showLoadingState(message = '') {
  state.viewMode = 'loading';
  state.accessDeniedMessage = '';
  setAppShellVisibility(true);
  setWorkspaceVisibility(false);
  renderSessionBox();
  renderConnectionCard();
  renderHeaderMeta();
  renderDetailPanel();
  if (message) {
    setNotice(message, 'info');
  }
}

function showAccessDeniedState(message) {
  state.viewMode = 'denied';
  state.accessDeniedMessage = 'Welcome';
  state.matrix = [];
  state.lastSavedMatrix = [];
  state.selectedCell = null;
  state.searchQuery = '';
  state.lastLoadedAt = null;
  clearSaveTimers();
  resetCellState();
  setAppShellVisibility(true);
  setWorkspaceVisibility(false);
  renderSessionBox();
  renderConnectionCard();
  renderHeaderMeta();
  renderDetailPanel();
  setNotice('Welcome');
}

function showLoadErrorState(message) {
  state.viewMode = 'error';
  state.accessDeniedMessage = message || 'Unable to load spreadsheet.';
  state.matrix = [];
  state.lastSavedMatrix = [];
  state.selectedCell = null;
  state.searchQuery = '';
  state.lastLoadedAt = null;
  clearSaveTimers();
  resetCellState();
  setAppShellVisibility(true);
  setWorkspaceVisibility(false);
  renderSessionBox();
  renderConnectionCard();
  renderHeaderMeta();
  renderDetailPanel();
  setNotice(state.accessDeniedMessage, 'error');
}

function scheduleSessionExpiryTimer() {
  clearSessionExpiryTimer();
  if (!state.session?.expiresAt) return;

  const delay = Math.max(0, Number(state.session.expiresAt) - Date.now() - SESSION_EXPIRY_SKEW_MS);
  sessionExpiryTimer = setTimeout(() => {
    if (isSessionExpired(state.session)) {
      void handleExpiredSession('Your Google sign-in expired. Please sign in again.');
      return;
    }
    scheduleSessionExpiryTimer();
  }, delay);
}

function clearRuntimeSession(message = '', { kind = 'info' } = {}) {
  clearSessionExpiryTimer();
  state.session = null;
  window.userToken = '';
  state.activeSheetName = '';
  state.viewMode = 'auth';
  state.accessDeniedMessage = '';
  showAuthOnlyState(message, kind);
  setSignInState({ loading: false, enabled: state.authReady, label: 'Continue with Google' });
}

function handleExpiredSession(message = 'Your Google sign-in expired. Please sign in again.') {
  clearRuntimeSession(message, { kind: 'error' });
}

function createApiError(message, status = 500, payload = null) {
  const error = new Error(message);
  error.status = status;
  error.payload = payload;
  return error;
}

async function readJson(response) {
  const text = await response.text();
  if (!text) return {};
  try {
    return JSON.parse(text);
  } catch {
    return { raw: text };
  }
}

function buildRequestHeaders(extraHeaders = {}) {
  const headers = { ...(extraHeaders || {}) };

  if (FORWARD_AUTH_TOKEN && getCurrentCredential()) {
    headers.authorization = `Bearer ${getCurrentCredential()}`;
  }

  return headers;
}

function getCurrentCredential() {
  return state.session?.credential || window.userToken || '';
}

function getCurrentSheetName(fallback = '') {
  return state.activeSheetName || fallback || '';
}

function buildSpreadsheetUrl(sheetName = '') {
  if (!APPS_SCRIPT_URL) {
    throw createApiError('Apps Script URL is missing from window.APP_CONFIG.', 500);
  }

  const idToken = getCurrentCredential();
  if (!idToken) {
    throw createApiError('Please sign in with Google to continue.', 401);
  }

  const url = new URL(APPS_SCRIPT_URL);
  url.searchParams.set('idToken', idToken);
  url.searchParams.set('sheetName', sheetName || '');
  return url.toString();
}

async function requestSpreadsheet(method, body = null) {
  const request = {
    method,
    mode: 'cors',
    cache: 'no-store',
  };

  if (method === 'GET') {
    const response = await fetch(buildSpreadsheetUrl(getCurrentSheetName()), request);
    return readSpreadsheetResponse(response);
  }

  if (method === 'POST') {
    const payload = {
      ...(body || {}),
      idToken: getCurrentCredential(),
      sheetName: getCurrentSheetName(body?.sheetName),
    };
    request.headers = buildRequestHeaders({ 'content-type': 'text/plain;charset=UTF-8' });
    request.body = JSON.stringify(payload);
    const response = await fetch(APPS_SCRIPT_URL, request);
    return readSpreadsheetResponse(response);
  }

  throw createApiError(`Unsupported request method: ${method}`, 500);
}

async function readSpreadsheetResponse(response) {
  const payload = await readJson(response);
  if (!response.ok || payload?.status === 'error') {
    throw createApiError(
      payload?.message || payload?.error || payload?.status || `Request failed (${response.status})`,
      response.status,
      payload,
    );
  }
  return payload;
}

async function loadMatrixFromSource() {
  const payload = await requestSpreadsheet('GET');
  state.activeSheetName = String(payload?.currentSheet || state.activeSheetName || '');
  const matrixData = extractMatrixData(payload);
  if (!matrixData) {
    throw createApiError('Apps Script response must include a JSON data matrix.', 500, payload);
  }
  return normalizeMatrix(matrixData);
}

async function saveCellToSource(row, col, val) {
  return requestSpreadsheet('POST', {
    row,
    col,
    val,
    sheetName: getCurrentSheetName(),
  });
}

function setNotice(message, kind = 'info') {
  state.notice = message || '';
  state.noticeKind = kind;
  const notice = elements.notice;
  if (!notice) return;

  if (!message) {
    notice.classList.add('hidden');
    notice.classList.remove('error');
    notice.textContent = '';
    return;
  }

  notice.classList.remove('hidden');
  notice.classList.toggle('error', kind === 'error');
  notice.textContent = message;
}

function setLoading(isLoading, label = '') {
  state.loading = isLoading;
  if (elements.refreshButton) elements.refreshButton.disabled = isLoading;
  if (elements.newButton) elements.newButton.disabled = isLoading;
  if (elements.saveButton) elements.saveButton.disabled = isLoading;
  if (elements.deleteButton) elements.deleteButton.disabled = isLoading;
  if (elements.cancelButton) elements.cancelButton.disabled = isLoading;
  if (elements.sheetTitle && label) elements.sheetTitle.textContent = label;
  renderConnectionCard();
}

function clearSaveTimers() {
  for (const timer of state.saveTimers.values()) {
    clearTimeout(timer);
  }
  state.saveTimers.clear();
}

function resetCellState() {
  state.dirtyCells.clear();
  state.savingCells.clear();
  state.cellErrors.clear();
}

function resetAppStateAfterLoad(matrix) {
  state.matrix = cloneMatrix(matrix);
  state.lastSavedMatrix = cloneMatrix(matrix);
  state.lastLoadedAt = new Date();
  state.searchQuery = elements.searchInput?.value || '';
  resetCellState();
  clearSaveTimers();
}

function getMaxColumns() {
  return Math.max(1, ...state.matrix.map((row) => row.length));
}

function getVisibleRows() {
  const query = state.searchQuery.trim().toLowerCase();
  return state.matrix
    .map((row, rowIndex) => ({ row, rowIndex }))
    .filter(({ row }) => !query || row.some((cell) => String(cell ?? '').toLowerCase().includes(query)));
}

function renderSessionBox() {
  const box = elements.sessionBox;
  if (!box) return;
  box.innerHTML = '';

  if (!state.session?.authenticated) {
    box.innerHTML = `
      <div class="profile">
        <strong>Not signed in</strong>
        <span>Browser session is available for future restricted mode</span>
      </div>
    `;
    return;
  }

  const avatar = document.createElement('img');
  avatar.className = 'avatar';
  avatar.alt = '';
  avatar.src = state.session.profile.picture || `data:image/svg+xml;utf8,${encodeURIComponent(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 96 96"><rect width="96" height="96" rx="48" fill="#0f2037"/><text x="48" y="55" text-anchor="middle" font-size="36" fill="#6dd3ff" font-family="Poppins">${(state.session.profile.name || '?')[0]?.toUpperCase() || '?'}</text></svg>`)}`;

  const profile = document.createElement('div');
  profile.className = 'profile';
  profile.innerHTML = `
    <strong>${escapeHtml(state.session.profile.name || state.session.profile.email)}</strong>
    <span>${escapeHtml(state.session.profile.email)}</span>
  `;

  const logout = document.createElement('button');
  logout.className = 'secondary-button inline-button';
  logout.textContent = 'Logout';
  logout.addEventListener('click', logoutUser);

  box.append(avatar, profile, logout);
}

function renderConnectionCard() {
  if (!elements.sheetTabs) return;
  if (!state.session?.authenticated) {
    elements.sheetTabs.innerHTML = '';
    return;
  }
  const lines = [];
  lines.push(`<div class="panel-meta"><strong>Source:</strong> Apps Script web app</div>`);
  lines.push(`<div class="panel-meta"><strong>Endpoint:</strong> ${escapeHtml(APPS_SCRIPT_URL || 'missing')}</div>`);
  lines.push(`<div class="panel-meta"><strong>Status:</strong> ${state.loading ? 'Loading...' : state.lastLoadedAt ? `Loaded ${state.lastLoadedAt.toLocaleTimeString()}` : 'Ready'}</div>`);
  elements.sheetTabs.innerHTML = lines.join('');
}

function renderHeaderMeta() {
  if (!state.session?.authenticated) {
    if (elements.spreadsheetTitle) {
      elements.spreadsheetTitle.textContent = 'Sign in required';
    }
    if (elements.sheetMeta) {
      elements.sheetMeta.textContent = 'Your spreadsheet stays hidden until you sign in with Google.';
    }
    if (elements.sheetSummary) {
      elements.sheetSummary.textContent = 'Sign in to load the live sheet.';
    }
    return;
  }

  if (elements.spreadsheetTitle) {
    elements.spreadsheetTitle.textContent = state.activeSheetName || 'Live spreadsheet';
  }
  if (elements.sheetMeta) {
  if (state.viewMode === 'denied') {
    elements.sheetMeta.textContent = 'Access denied';
  } else if (state.viewMode === 'loading') {
    elements.sheetMeta.textContent = 'Checking access...';
  } else if (state.viewMode === 'error') {
    elements.sheetMeta.textContent = 'Unable to load spreadsheet';
  } else {
    elements.sheetMeta.textContent = state.matrix.length
      ? `${state.matrix.length} row${state.matrix.length === 1 ? '' : 's'} · ${getMaxColumns()} column${getMaxColumns() === 1 ? '' : 's'}`
      : 'No rows loaded yet.';
    }
  }
  if (elements.sheetTitle) {
    elements.sheetTitle.textContent = state.activeSheetName || 'Native spreadsheet grid';
  }
  if (elements.sheetSummary) {
  if (state.viewMode === 'denied') {
    elements.sheetSummary.textContent = state.accessDeniedMessage || 'Access denied.';
  } else if (state.viewMode === 'loading') {
    elements.sheetSummary.textContent = 'Loading sheet data from Apps Script...';
  } else if (state.viewMode === 'error') {
    elements.sheetSummary.textContent = state.accessDeniedMessage || 'Unable to load spreadsheet.';
  } else {
    elements.sheetSummary.textContent = APPS_SCRIPT_URL
      ? 'Cells autosave back to your Apps Script web app as you type.'
      : 'Set the Apps Script URL in the page config to load data.';
    }
  }
}

function renderDetailPanel() {
  if (!elements.detailStatus) return;
  if (!state.session?.authenticated) {
    elements.detailStatus.innerHTML = `
      <div><strong>Editing mode:</strong> Sign in to load the spreadsheet.</div>
      <div><strong>Tip:</strong> The grid stays hidden until a Google session is active.</div>
    `;
    return;
  }

  if (state.viewMode === 'loading') {
    elements.detailStatus.innerHTML = `
      <div><strong>Signed in:</strong> Verifying access with Apps Script.</div>
      <div><strong>Status:</strong> Loading spreadsheet data...</div>
    `;
    return;
  }

  if (state.viewMode === 'denied') {
    elements.detailStatus.innerHTML = `
      <div><strong>Welcome</strong></div>
    `;
    return;
  }

  if (state.viewMode === 'error') {
    elements.detailStatus.innerHTML = `
      <div><strong>Unable to load spreadsheet:</strong> ${escapeHtml(state.accessDeniedMessage || 'An unexpected error occurred.')}</div>
      <div><strong>Status:</strong> Please try loading again.</div>
    `;
    return;
  }

  if (state.selectedCell) {
    const { row, col } = state.selectedCell;
    const value = state.matrix[row - 1]?.[col - 1] ?? '';
    const saved = state.lastSavedMatrix[row - 1]?.[col - 1] ?? '';
    const error = state.cellErrors.get(keyFor(row, col));
    elements.detailStatus.innerHTML = `
      <div><strong>Selected cell:</strong> ${row}, ${columnLabel(col)}</div>
      <div><strong>Current value:</strong> ${escapeHtml(value || '(empty)')}</div>
      <div><strong>Saved value:</strong> ${escapeHtml(saved || '(empty)')}</div>
      <div><strong>Status:</strong> ${error ? escapeHtml(error) : 'Editing is live and autosaved.'}</div>
    `;
  } else {
    elements.detailStatus.innerHTML = `
      <div><strong>Editing mode:</strong> Autosave is on.</div>
      <div><strong>Tip:</strong> Click any cell, type, and the change will save after you pause.</div>
    `;
  }
}

function isDirtyCell(row, col) {
  return state.dirtyCells.has(keyFor(row, col));
}

function isSavingCell(row, col) {
  return state.savingCells.has(keyFor(row, col));
}

function getCellError(row, col) {
  return state.cellErrors.get(keyFor(row, col));
}

function updateCellClasses(input, row, col) {
  const key = keyFor(row, col);
  input.classList.toggle('cell-dirty', state.dirtyCells.has(key));
  input.classList.toggle('cell-saving', state.savingCells.has(key));
  input.classList.toggle('cell-error', state.cellErrors.has(key));
}

function renderGrid() {
  const table = elements.dataTable;
  if (!table) return;

  const thead = table.querySelector('thead');
  const tbody = table.querySelector('tbody');
  thead.innerHTML = '';
  tbody.innerHTML = '';

  const maxColumns = getMaxColumns();
  const headerRow = document.createElement('tr');

  const corner = document.createElement('th');
  corner.textContent = '#';
  headerRow.appendChild(corner);

  for (let col = 1; col <= maxColumns; col += 1) {
    const th = document.createElement('th');
    th.textContent = columnLabel(col);
    headerRow.appendChild(th);
  }
  thead.appendChild(headerRow);

  const rows = getVisibleRows();
  elements.rowCount.textContent = `${rows.length} row${rows.length === 1 ? '' : 's'} visible`;

  rows.forEach(({ row, rowIndex }) => {
    const tr = document.createElement('tr');

    const rowHead = document.createElement('th');
    rowHead.className = 'row-index';
    rowHead.textContent = String(rowIndex + 1);
    tr.appendChild(rowHead);

    for (let col = 1; col <= maxColumns; col += 1) {
      const value = row[col - 1] ?? '';
      const td = document.createElement('td');
      td.dataset.row = String(rowIndex + 1);
      td.dataset.col = String(col);

      const input = document.createElement('input');
      input.type = 'text';
      input.spellcheck = false;
      input.autocomplete = 'off';
      input.value = value;
      input.dataset.row = String(rowIndex + 1);
      input.dataset.col = String(col);
      input.className = 'sheet-cell-input';
      input.addEventListener('focus', () => {
        state.selectedCell = { row: rowIndex + 1, col };
        renderDetailPanel();
      });
      input.addEventListener('input', handleCellInput);
      input.addEventListener('change', handleCellChange);
      input.addEventListener('keydown', handleCellKeydown);

      updateCellClasses(input, rowIndex + 1, col);
      td.appendChild(input);
      tr.appendChild(td);
    }

    tbody.appendChild(tr);
  });

  renderHeaderMeta();
  renderDetailPanel();
}

function setCellError(row, col, message) {
  const key = keyFor(row, col);
  if (message) {
    state.cellErrors.set(key, message);
  } else {
    state.cellErrors.delete(key);
  }
}

function setCellDirty(row, col, isDirty) {
  const key = keyFor(row, col);
  if (isDirty) {
    state.dirtyCells.add(key);
  } else {
    state.dirtyCells.delete(key);
  }
}

function updateInMemoryCell(row, col, value) {
  while (state.matrix.length < row) state.matrix.push(Array.from({ length: getMaxColumns() }, () => ''));
  const targetRow = state.matrix[row - 1];
  while (targetRow.length < col) targetRow.push('');
  targetRow[col - 1] = value;
}

function getInputFromEvent(event) {
  const input = event.currentTarget;
  const row = Number(input.dataset.row);
  const col = Number(input.dataset.col);
  return { input, row, col };
}

function scheduleSave(row, col, input, immediate = false) {
  const key = keyFor(row, col);
  const existingTimer = state.saveTimers.get(key);
  if (existingTimer) clearTimeout(existingTimer);

  if (immediate) {
    void commitCellSave(row, col, input);
    return;
  }

  state.saveTimers.set(key, setTimeout(() => {
    state.saveTimers.delete(key);
    void commitCellSave(row, col, input);
  }, SAVE_DEBOUNCE_MS));
}

async function commitCellSave(row, col, input) {
  const key = keyFor(row, col);
  const value = input.value;
  setCellDirty(row, col, false);
  setCellError(row, col, '');
  state.savingCells.add(key);
  updateCellClasses(input, row, col);
  renderDetailPanel();

  try {
    await saveCellToSource(row, col, value);
    state.lastSavedMatrix[row - 1] = state.lastSavedMatrix[row - 1] || [];
    state.lastSavedMatrix[row - 1][col - 1] = value;
    input.classList.remove('cell-error');
    setNotice(`Saved ${columnLabel(col)}${row}.`);
  } catch (error) {
    setCellDirty(row, col, true);
    setCellError(row, col, error.message || 'Save failed');
    input.classList.add('cell-error');
    setNotice(`Unable to save ${columnLabel(col)}${row}: ${error.message || 'unknown error'}`, 'error');
  } finally {
    state.savingCells.delete(key);
    updateCellClasses(input, row, col);
    renderDetailPanel();
  }
}

function handleCellInput(event) {
  const { input, row, col } = getInputFromEvent(event);
  updateInMemoryCell(row, col, input.value);
  setCellDirty(row, col, input.value !== (state.lastSavedMatrix[row - 1]?.[col - 1] ?? ''));
  setCellError(row, col, '');
  updateCellClasses(input, row, col);
  renderDetailPanel();
  scheduleSave(row, col, input, false);
}

function handleCellChange(event) {
  const { input, row, col } = getInputFromEvent(event);
  scheduleSave(row, col, input, true);
}

function handleCellKeydown(event) {
  if (event.key !== 'Enter') return;
  event.preventDefault();
  const { input, row, col } = getInputFromEvent(event);
  scheduleSave(row, col, input, true);
  input.blur();
}

function hasPendingChanges() {
  return state.dirtyCells.size > 0 || state.saveTimers.size > 0 || state.savingCells.size > 0;
}

async function flushPendingSaves() {
  if (!state.session?.authenticated) {
    return;
  }

  const pendingKeys = [...state.saveTimers.keys()];
  for (const key of pendingKeys) {
    const timer = state.saveTimers.get(key);
    if (timer) clearTimeout(timer);
    state.saveTimers.delete(key);
  }

  const inputs = [...document.querySelectorAll('.sheet-cell-input')];
  const saves = [];

  for (const input of inputs) {
    const row = Number(input.dataset.row);
    const col = Number(input.dataset.col);
    const key = keyFor(row, col);
    if (state.dirtyCells.has(key) || state.cellErrors.has(key)) {
      saves.push(commitCellSave(row, col, input));
    }
  }

  await Promise.all(saves);
}

async function reloadMatrix() {
  if (!state.session?.authenticated) {
    setNotice('Sign in with Google to load the spreadsheet.', 'error');
    return;
  }

  if (hasPendingChanges() && !window.confirm('Reloading will discard pending local edits that have not saved yet. Continue?')) {
    return;
  }

  await loadSpreadsheetIntoGrid('Reloading spreadsheet...');
}

async function loadSpreadsheetIntoGrid(loadingMessage = 'Loading spreadsheet...') {
  if (!state.session?.authenticated) {
    showAuthOnlyState();
    return;
  }

  state.viewMode = 'loading';
  state.accessDeniedMessage = '';
  setAppShellVisibility(true);
  setWorkspaceVisibility(false);
  renderSessionBox();
  renderConnectionCard();
  renderHeaderMeta();
  renderDetailPanel();
  setLoading(true, loadingMessage);

  try {
    const matrix = await loadMatrixFromSource();
    resetAppStateAfterLoad(matrix);
    state.viewMode = 'ready';
    state.accessDeniedMessage = '';
    setWorkspaceVisibility(true);
    setNotice('Spreadsheet loaded.');
    renderGrid();
    renderConnectionCard();
    renderHeaderMeta();
    renderDetailPanel();
  } catch (error) {
    const message = error.message || 'Unable to load spreadsheet.';
    const normalizedMessage = String(message);
    if (normalizedMessage.toLowerCase().includes('access denied')) {
      showAccessDeniedState(normalizedMessage);
    } else if (normalizedMessage.toLowerCase().includes('missing authentication token') || normalizedMessage.toLowerCase().includes('invalid or expired login session')) {
      handleExpiredSession(normalizedMessage);
    } else if (normalizedMessage.toLowerCase().includes('unauthorized')) {
      showAccessDeniedState(normalizedMessage);
    } else {
      showLoadErrorState(normalizedMessage);
    }
  } finally {
    setLoading(false);
  }
}

function addBlankRow() {
  if (!state.session?.authenticated) {
    setNotice('Sign in with Google to add rows.', 'error');
    return;
  }

  const width = getMaxColumns();
  const blankRow = Array.from({ length: width }, () => '');
  state.matrix.push(blankRow);
  state.lastSavedMatrix.push(Array.from({ length: width }, () => ''));
  renderGrid();
  renderHeaderMeta();

  window.requestAnimationFrame(() => {
    const input = document.querySelector(`.sheet-cell-input[data-row="${state.matrix.length}"][data-col="1"]`);
    if (input) input.focus();
  });

  setNotice(`Added blank row ${state.matrix.length}. Type to autosave it.`);
}

function clearAllErrors() {
  if (!state.session?.authenticated) {
    return;
  }

  state.cellErrors.clear();

  for (const input of document.querySelectorAll('.sheet-cell-input')) {
    const row = Number(input.dataset.row);
    const col = Number(input.dataset.col);
    updateCellClasses(input, row, col);
  }

  renderDetailPanel();
  setNotice('Cleared cell error highlights.');
}

async function waitForGoogleIdentityLibrary() {
  if (window.google?.accounts?.id) return;

  await new Promise((resolve, reject) => {
    const started = Date.now();
    const tick = () => {
      if (window.google?.accounts?.id) return resolve();
      if (Date.now() - started > 10000) return reject(new Error('Google sign-in library did not load.'));
      setTimeout(tick, 50);
    };
    tick();
  });
}

function initGoogleIdentityClient() {
  if (!GOOGLE_CLIENT_ID || !window.google?.accounts?.id) return;

  window.google.accounts.id.initialize({
    client_id: GOOGLE_CLIENT_ID,
    callback: handleGoogleCredentialResponse,
    auto_select: false,
    cancel_on_tap_outside: false,
  });
  state.authReady = true;
  setSignInState({ enabled: true, label: 'Continue with Google' });
}

function setSignInState({ loading = false, enabled = false, label = 'Continue with Google' } = {}) {
  if (elements.signInButton) elements.signInButton.disabled = !enabled || loading;
  if (elements.signInLabel) elements.signInLabel.textContent = loading ? 'Preparing Google sign-in...' : label;
  if (elements.oauthState) {
    elements.oauthState.textContent = loading
      ? 'Preparing Google sign-in...'
      : enabled
        ? 'Ready to sign in'
        : 'Google sign-in unavailable';
  }
}

async function prepareSignIn() {
  setSignInState({ loading: true, enabled: false });
  try {
    if (!GOOGLE_CLIENT_ID) {
      state.authReady = false;
      setSignInState({ loading: false, enabled: false, label: 'Google sign-in unavailable' });
      setNotice('Google client ID is missing from app config.', 'error');
      return;
    }
    await waitForGoogleIdentityLibrary();
    initGoogleIdentityClient();
  } catch (error) {
    state.authReady = false;
    setSignInState({ loading: false, enabled: false, label: 'Google sign-in unavailable' });
    setNotice(error.message || 'Google sign-in is unavailable.', 'error');
  }
}

async function handleGoogleCredentialResponse(response) {
  const credential = response?.credential || '';
  if (!credential) {
    setNotice('Google sign-in did not return a credential token.', 'error');
    return;
  }

  state.session = buildSessionFromCredential(credential);
  window.userToken = credential;
  state.activeSheetName = '';
  scheduleSessionExpiryTimer();
  setSignInState({ loading: true, enabled: false, label: 'Signing in...' });
  showLoadingState('Signing in...');

  try {
    await loadSpreadsheetIntoGrid('Loading spreadsheet...');
    setNotice('Signed in');
  } catch (error) {
    const message = error.message || 'Unable to load spreadsheet.';
    if (String(message).toLowerCase().includes('access denied')) {
      showAccessDeniedState(message);
    } else if (String(message).toLowerCase().includes('missing authentication token') || String(message).toLowerCase().includes('invalid or expired login session')) {
      handleExpiredSession(message);
    } else {
      showLoadErrorState(message);
    }
  } finally {
    setLoading(false);
    setSignInState({ loading: false, enabled: state.authReady, label: 'Continue with Google' });
  }
}

function requestSignIn() {
  if (!state.authReady || !window.google?.accounts?.id) {
    setNotice('Google sign-in is still preparing. Please try again in a moment.', 'error');
    return;
  }

  window.google.accounts.id.prompt((notification) => {
    if (notification.isNotDisplayed()) {
      setNotice('Google sign-in could not be displayed. Check browser settings or pop-up blocking.', 'error');
    } else if (notification.isSkippedMoment()) {
      setNotice('Google sign-in was skipped. Please try again.', 'error');
    }
  });
}

async function logoutUser() {
  if (window.google?.accounts?.id?.disableAutoSelect) {
    try {
      window.google.accounts.id.disableAutoSelect();
    } catch {
      // Logout should still succeed if the GIS helper rejects.
    }
  }

  clearRuntimeSession('You have been logged out.', { kind: 'info' });
  setSignInState({ loading: false, enabled: state.authReady, label: 'Continue with Google' });
}

function bindElements() {
  elements.notice = $('notice');
  elements.authPanel = $('authPanel');
  elements.dashboard = $('dashboard');
  elements.workspace = document.querySelector('.workspace');
  elements.signInButton = $('signInButton');
  elements.signInLabel = $('signInButton')?.querySelector('.button-label');
  elements.oauthState = $('oauthState');
  elements.sessionBox = $('sessionBox');
  elements.spreadsheetTitle = $('spreadsheetTitle');
  elements.sheetMeta = $('sheetMeta');
  elements.sheetTabs = $('sheetTabs');
  elements.searchInput = $('searchInput');
  elements.refreshButton = $('refreshButton');
  elements.newButton = $('newButton');
  elements.sheetTitle = $('sheetTitle');
  elements.sheetSummary = $('sheetSummary');
  elements.rowCount = $('rowCount');
  elements.dataTable = $('dataTable');
  elements.detailStatus = $('detailStatus');
  elements.detailForm = $('detailForm');
  elements.saveButton = $('saveButton');
  elements.deleteButton = $('deleteButton');
  elements.cancelButton = $('cancelButton');
}

function bindEvents() {
  elements.signInButton.addEventListener('click', requestSignIn);
  if (elements.refreshButton) {
    elements.refreshButton.addEventListener('click', reloadMatrix);
  }
  elements.newButton.addEventListener('click', addBlankRow);
  elements.saveButton.addEventListener('click', () => flushPendingSaves());
  elements.deleteButton.addEventListener('click', clearAllErrors);
  elements.cancelButton.addEventListener('click', reloadMatrix);
  elements.searchInput.addEventListener('input', () => {
    state.searchQuery = elements.searchInput.value;
    renderGrid();
  });
}

async function bootstrap() {
  bindElements();
  bindEvents();

  setNotice('');
  showAuthOnlyState();

  const authSetup = prepareSignIn();
  await authSetup;

  const params = new URLSearchParams(window.location.search);
  if (params.get('error') === 'permission-denied') {
    setNotice('The Google account you signed in with does not have access to the private spreadsheet.', 'error');
  } else if (params.get('error') === 'login-failed') {
    setNotice('Google sign-in failed. Please try again.', 'error');
  }
}

window.addEventListener('error', (event) => {
  if (String(event.error?.message || '').includes('CORS')) {
    setNotice('The Apps Script web app request was blocked by CORS. If that happens, the browser must call the web app through a same-origin proxy.', 'error');
  }
});

bootstrap().catch((error) => setNotice(error.message || 'Unexpected startup error', 'error'));
