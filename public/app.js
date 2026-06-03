const APP_CONFIG = window.APP_CONFIG || {};
const GOOGLE_CLIENT_ID = APP_CONFIG.googleClientId || '';
const SPREADSHEET_ID = APP_CONFIG.spreadsheetId || '';
const AUTH_SCOPE = APP_CONFIG.authScope || 'openid email profile https://www.googleapis.com/auth/spreadsheets';
const AUTH_STORAGE_KEY = 'secure-google-dashboard.auth.v1';
const SESSION_EXPIRY_SKEW_MS = 60_000;

const state = {
  session: null,
  spreadsheetTitle: '',
  sheets: [],
  sheetName: '',
  sheetData: null,
  loading: false,
  search: '',
  filterField: '',
  filterValue: '',
  selectedRow: null,
  mode: 'view',
  accessState: 'signed-out',
  accessMessage: '',
  authClient: null,
  authReady: false,
  pendingAuthRequest: null,
};

let sessionExpiryTimer = null;

const elements = {};

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

function readStoredSession() {
  try {
    const raw = window.localStorage.getItem(AUTH_STORAGE_KEY);
    if (!raw) return null;

    const parsed = JSON.parse(raw);
    const expiresAt = Number(parsed?.expiresAt);
    if (!parsed?.accessToken || !Number.isFinite(expiresAt)) {
      window.localStorage.removeItem(AUTH_STORAGE_KEY);
      return null;
    }

    if (Date.now() >= expiresAt - SESSION_EXPIRY_SKEW_MS) {
      window.localStorage.removeItem(AUTH_STORAGE_KEY);
      return null;
    }

    return {
      authenticated: true,
      accessToken: parsed.accessToken,
      expiresAt,
      profile: {
        id: String(parsed.profile?.id || ''),
        email: String(parsed.profile?.email || ''),
        name: String(parsed.profile?.name || parsed.profile?.email || ''),
        picture: String(parsed.profile?.picture || ''),
      },
    };
  } catch {
    try {
      window.localStorage.removeItem(AUTH_STORAGE_KEY);
    } catch {
      // Ignore storage cleanup errors.
    }
    return null;
  }
}

function persistSession(session) {
  try {
    window.localStorage.setItem(AUTH_STORAGE_KEY, JSON.stringify({
      accessToken: session.accessToken,
      expiresAt: session.expiresAt,
      profile: session.profile,
    }));
  } catch {
    // Ignore storage quota or privacy-mode failures.
  }
}

function clearStoredSession() {
  try {
    window.localStorage.removeItem(AUTH_STORAGE_KEY);
  } catch {
    // Ignore storage cleanup errors.
  }
}

function isSessionExpired(session) {
  return !session?.accessToken || !Number.isFinite(Number(session?.expiresAt)) || Date.now() >= Number(session.expiresAt) - SESSION_EXPIRY_SKEW_MS;
}

function resetWorkspaceState() {
  state.spreadsheetTitle = '';
  state.sheets = [];
  state.sheetName = '';
  state.sheetData = null;
  state.search = '';
  state.filterField = '';
  state.filterValue = '';
  state.selectedRow = null;
  state.mode = 'view';
}

function renderWorkspaceState() {
  renderSessionBox();
  renderSheetTabs();
  renderSheetMeta();
  renderTable();
  renderDetailForm();
  syncActionButtons();
}

function setNoAccessState(message) {
  state.accessState = 'no-access';
  state.accessMessage = message;
  resetWorkspaceState();
  renderWorkspaceState();
}

function setAuthorizedState() {
  state.accessState = 'authorized';
  state.accessMessage = '';
}

const GOOGLE_SHEETS_BASE = 'https://sheets.googleapis.com/v4/spreadsheets';
const GOOGLE_USERINFO_URL = 'https://openidconnect.googleapis.com/v1/userinfo';

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

async function browserApiFetch(url, accessToken, options = {}) {
  const response = await fetch(url, {
    ...options,
    headers: {
      authorization: `Bearer ${accessToken}`,
      ...(options.body ? { 'content-type': 'application/json' } : {}),
      ...(options.headers || {}),
    },
  });

  const payload = await readJson(response);
  if (!response.ok) {
    throw createApiError(
      payload?.error?.message || payload?.error_description || payload?.error || `Request failed (${response.status})`,
      response.status,
      payload,
    );
  }
  return payload;
}

async function fetchGoogleUserInfo(accessToken) {
  return browserApiFetch(GOOGLE_USERINFO_URL, accessToken);
}

async function fetchSpreadsheetMetadata(accessToken) {
  return browserApiFetch(
    `${GOOGLE_SHEETS_BASE}/${encodeURIComponent(SPREADSHEET_ID)}?fields=spreadsheetId,properties.title,sheets.properties`,
    accessToken,
  );
}

async function fetchSheetValues(sheetName, accessToken) {
  const range = encodeURIComponent(sheetName);
  const payload = await browserApiFetch(
    `${GOOGLE_SHEETS_BASE}/${encodeURIComponent(SPREADSHEET_ID)}/values/${range}?majorDimension=ROWS`,
    accessToken,
  );
  return payload.values || [];
}

async function appendSheetRow({ sheetName, accessToken, values }) {
  const range = encodeURIComponent(sheetName);
  return browserApiFetch(
    `${GOOGLE_SHEETS_BASE}/${encodeURIComponent(SPREADSHEET_ID)}/values/${range}:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`,
    accessToken,
    {
      method: 'POST',
      body: JSON.stringify({
        majorDimension: 'ROWS',
        values: [values],
      }),
    },
  );
}

async function updateSheetRow({ sheetName, accessToken, rowNumber, values, columnCount }) {
  const endColumn = columnNumberToName(columnCount);
  const range = encodeURIComponent(`${sheetName}!A${rowNumber}:${endColumn}${rowNumber}`);
  return browserApiFetch(
    `${GOOGLE_SHEETS_BASE}/${encodeURIComponent(SPREADSHEET_ID)}/values/${range}?valueInputOption=USER_ENTERED`,
    accessToken,
    {
      method: 'PUT',
      body: JSON.stringify({
        majorDimension: 'ROWS',
        values: [values],
      }),
    },
  );
}

async function deleteSheetRow({ sheetId, accessToken, rowNumber }) {
  return browserApiFetch(
    `${GOOGLE_SHEETS_BASE}/${encodeURIComponent(SPREADSHEET_ID)}:batchUpdate`,
    accessToken,
    {
      method: 'POST',
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
    },
  );
}

function columnNumberToName(number) {
  let n = Math.max(1, Number(number) || 1);
  let result = '';
  while (n > 0) {
    const remainder = (n - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result;
}

function parseSheetTable(values) {
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

function buildRowValues(headers, rowValues) {
  return headers.map((header) => {
    const value = rowValues[header];
    return value == null ? '' : value;
  });
}

function setNotice(message, kind = 'info') {
  const notice = elements.notice;
  if (!message) {
    notice.classList.add('hidden');
    notice.textContent = '';
    notice.classList.remove('error');
    return;
  }
  notice.textContent = message;
  notice.classList.remove('hidden');
  notice.classList.toggle('error', kind === 'error');
}

function setLoading(isLoading) {
  state.loading = isLoading;
  elements.refreshButton.disabled = isLoading || !state.session;
  elements.newButton.disabled = isLoading || state.accessState !== 'authorized' || !state.sheetData?.headers?.length;
  syncActionButtons();
}

function syncActionButtons() {
  const hasEditableSheet = state.accessState === 'authorized' && Boolean(state.sheetData?.headers?.length);
  const canSave = state.mode === 'create' ? hasEditableSheet : Boolean(state.selectedRow) && hasEditableSheet;
  elements.saveButton.disabled = state.loading || !canSave;
  elements.deleteButton.disabled = state.loading || !hasEditableSheet || !state.selectedRow || state.mode === 'create';
  elements.cancelButton.disabled = state.loading || (!state.selectedRow && state.mode !== 'create');
}

function setSignInState({ loading = false, enabled = false, label = 'Continue with Google' } = {}) {
  elements.signInButton.disabled = !enabled || loading;
  const labelNode = elements.signInLabel;
  if (labelNode) labelNode.textContent = loading ? 'Preparing Google sign-in...' : label;
  if (elements.oauthState) {
    elements.oauthState.textContent = loading
      ? 'Preparing Google sign-in...'
      : enabled
        ? 'Ready to sign in'
        : 'Google sign-in unavailable';
  }
}

function isGoogleVerificationBlock(response) {
  const text = [response?.error, response?.error_description, response?.message]
    .filter(Boolean)
    .join(' ')
    .toLowerCase();
  return text.includes('access_denied')
    || text.includes('verification process')
    || text.includes('developer-approved testers')
    || text.includes('app is currently being tested')
    || text.includes('test users');
}

function getGoogleVerificationHelpMessage() {
  return 'Google blocked this sign-in because the OAuth app is still in Testing or this Google account is not listed as a Test user. In Google Cloud Console, open APIs & Services > OAuth consent screen > Test users, add this account, then try again. If you want everyone to sign in, publish the app and submit it for verification.';
}

function clearSessionExpiryTimer() {
  if (sessionExpiryTimer) {
    clearTimeout(sessionExpiryTimer);
    sessionExpiryTimer = null;
  }
}

function scheduleSessionExpiryTimer() {
  clearSessionExpiryTimer();
  if (!state.session?.expiresAt) return;

  const delay = Math.max(0, Number(state.session.expiresAt) - Date.now() - SESSION_EXPIRY_SKEW_MS);
  sessionExpiryTimer = setTimeout(() => {
    if (isSessionExpired(state.session)) {
      void handleExpiredSession('Your Google sign-in expired. Please sign in again.');
    } else {
      scheduleSessionExpiryTimer();
    }
  }, delay);
}

function getLiveSession() {
  if (!state.session) return null;
  if (isSessionExpired(state.session)) return null;
  return state.session;
}

async function handleExpiredSession(message = 'Your Google sign-in expired. Please sign in again.') {
  clearSessionExpiryTimer();
  await bootstrapSignedOutView(message);
  await prepareSignIn();
}

async function bootstrapSignedOutView(message = '') {
  clearSessionExpiryTimer();
  state.session = null;
  state.accessState = 'signed-out';
  state.accessMessage = '';
  state.authReady = false;
  state.authClient = null;
  state.pendingAuthRequest = null;
  clearStoredSession();
  resetWorkspaceState();
  showSignedOutShell();
  renderWorkspaceState();
  setSignInState({ loading: true, enabled: false, label: 'Preparing Google sign-in...' });
  if (message) {
    setNotice(message, 'error');
  } else {
    setNotice('');
  }
}

async function bootstrapSignedInView({ autoLoad = true } = {}) {
  showSignedInShell();
  if (!autoLoad) {
    renderWorkspaceState();
    setNotice('Session restored. Click Refresh to load the spreadsheet.', 'info');
    return;
  }

  const authorized = await refreshSheetLists();
  if (authorized && state.sheets.length && state.accessState === 'authorized') {
    await selectSheet(state.sheets[0].title);
  }
}

function renderSessionBox() {
  const box = elements.sessionBox;
  box.innerHTML = '';

  if (!state.session?.authenticated) {
    box.innerHTML = '<div class="profile"><strong>Not signed in</strong><span>Google login required</span></div>';
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

function renderSheetTabs() {
  const container = elements.sheetTabs;
  container.innerHTML = '';

  if (state.accessState === 'no-access') {
    const message = document.createElement('div');
    message.className = 'panel-meta';
    message.textContent = 'This signed-in account cannot open the restricted spreadsheet.';
    container.appendChild(message);
    return;
  }

  state.sheets.forEach((sheet) => {
    const button = document.createElement('button');
    button.className = `sheet-tab ${sheet.title === state.sheetName ? 'active' : ''}`;
    button.innerHTML = `<span>${escapeHtml(sheet.title)}</span><span>${sheet.sheetId ?? ''}</span>`;
    button.addEventListener('click', () => selectSheet(sheet.title));
    container.appendChild(button);
  });
}

function renderSheetMeta() {
  if (!state.session?.authenticated) {
    elements.spreadsheetTitle.textContent = 'Private spreadsheet';
    elements.sheetMeta.textContent = 'Choose a worksheet to view rows.';
    elements.sheetTitle.textContent = 'Select a worksheet';
    elements.sheetSummary.textContent = 'The dashboard is waiting for a worksheet selection.';
    return;
  }

  if (state.accessState === 'no-access') {
    elements.spreadsheetTitle.textContent = state.spreadsheetTitle || 'Restricted spreadsheet';
    elements.sheetMeta.textContent = 'Signed in, but this account does not have spreadsheet access.';
    elements.sheetTitle.textContent = 'No access to this spreadsheet';
    elements.sheetSummary.textContent = state.accessMessage || 'Google denied permission for this account.';
    return;
  }

  elements.spreadsheetTitle.textContent = state.spreadsheetTitle || 'Private spreadsheet';
  elements.sheetMeta.textContent = state.sheetName
    ? `${state.sheets.length} worksheet${state.sheets.length === 1 ? '' : 's'} available`
    : 'Choose a worksheet to view rows.';
  elements.sheetTitle.textContent = state.sheetName || 'Select a worksheet';
  elements.sheetSummary.textContent = state.sheetName
    ? 'This data loads only after Google confirms the signed-in account can access the private spreadsheet.'
    : 'The dashboard is waiting for a worksheet selection.';
}

function getVisibleRows() {
  const rows = state.sheetData?.rows || [];
  const query = state.search.trim().toLowerCase();
  return rows.filter((row) => {
    const values = Object.values(row.values || {}).map((value) => String(value ?? '').toLowerCase());
    const matchesSearch = !query || values.some((value) => value.includes(query));
    const matchesFilter = !state.filterField || !state.filterValue || String(row.values?.[state.filterField] ?? '') === state.filterValue;
    return matchesSearch && matchesFilter;
  });
}

function populateFilterControls() {
  const headers = state.sheetData?.headers || [];
  elements.filterField.innerHTML = '';
  elements.filterValue.innerHTML = '';

  const allOption = document.createElement('option');
  allOption.value = '';
  allOption.textContent = 'All fields';
  elements.filterField.appendChild(allOption);

  headers.forEach((header) => {
    const option = document.createElement('option');
    option.value = header;
    option.textContent = header;
    elements.filterField.appendChild(option);
  });

  const valueAll = document.createElement('option');
  valueAll.value = '';
  valueAll.textContent = 'All values';
  elements.filterValue.appendChild(valueAll);

  if (state.filterField) {
    const unique = new Set();
    for (const row of state.sheetData?.rows || []) {
      const value = String(row.values?.[state.filterField] ?? '');
      if (value.trim()) unique.add(value);
    }
    [...unique].sort().forEach((value) => {
      const option = document.createElement('option');
      option.value = value;
      option.textContent = value;
      elements.filterValue.appendChild(option);
    });
  }

  elements.filterField.value = state.filterField;
  elements.filterValue.value = state.filterValue;
}

function renderTable() {
  const headers = state.sheetData?.headers || [];
  const rows = getVisibleRows();
  elements.rowCount.textContent = `${rows.length} row${rows.length === 1 ? '' : 's'}`;

  const thead = elements.dataTable.querySelector('thead');
  const tbody = elements.dataTable.querySelector('tbody');
  thead.innerHTML = '';
  tbody.innerHTML = '';

  if (state.accessState === 'no-access') {
    thead.innerHTML = '<tr><th>Status</th></tr>';
    tbody.innerHTML = `<tr><td>${escapeHtml(state.accessMessage || 'This signed-in account cannot open the restricted spreadsheet.')}</td></tr>`;
    return;
  }

  if (!headers.length) {
    thead.innerHTML = '<tr><th>No columns</th></tr>';
    tbody.innerHTML = '<tr><td>The selected worksheet does not appear to have a header row yet.</td></tr>';
    return;
  }

  const headerRow = document.createElement('tr');
  const rowNumberHead = document.createElement('th');
  rowNumberHead.textContent = '#';
  headerRow.appendChild(rowNumberHead);

  headers.forEach((header) => {
    const th = document.createElement('th');
    th.textContent = header;
    headerRow.appendChild(th);
  });

  const actionsHead = document.createElement('th');
  actionsHead.textContent = 'Actions';
  headerRow.appendChild(actionsHead);
  thead.appendChild(headerRow);

  rows.forEach((row) => {
    const tr = document.createElement('tr');
    if (state.selectedRow?.rowNumber === row.rowNumber) tr.style.background = 'rgba(109, 211, 255, 0.1)';
    tr.addEventListener('click', () => selectRow(row));

    const numberCell = document.createElement('td');
    numberCell.textContent = row.rowNumber;
    tr.appendChild(numberCell);

    headers.forEach((header) => {
      const td = document.createElement('td');
      td.innerHTML = escapeHtml(row.values?.[header] ?? '');
      tr.appendChild(td);
    });

    const actions = document.createElement('td');
    const group = document.createElement('div');
    group.className = 'row-actions';

    const view = document.createElement('button');
    view.type = 'button';
    view.className = 'secondary-button inline-button';
    view.textContent = 'View';
    view.addEventListener('click', (event) => {
      event.stopPropagation();
      selectRow(row);
    });

    const edit = document.createElement('button');
    edit.type = 'button';
    edit.className = 'primary-button inline-button';
    edit.textContent = 'Edit';
    edit.addEventListener('click', (event) => {
      event.stopPropagation();
      editRow(row);
    });

    const del = document.createElement('button');
    del.type = 'button';
    del.className = 'danger-button inline-button';
    del.textContent = 'Delete';
    del.addEventListener('click', async (event) => {
      event.stopPropagation();
      if (confirm(`Delete row ${row.rowNumber} from ${state.sheetName}?`)) {
        await deleteRow(row);
      }
    });

    group.append(view, edit, del);
    actions.appendChild(group);
    tr.appendChild(actions);
    tbody.appendChild(tr);
  });
}

function renderDetailForm() {
  const form = elements.detailForm;
  const headers = state.sheetData?.headers || [];
  form.innerHTML = '';

  if (state.accessState === 'no-access') {
    elements.detailStatus.textContent = state.accessMessage || 'This signed-in account cannot open the restricted spreadsheet.';
    syncActionButtons();
    return;
  }

  if (!state.sheetName) {
    elements.detailStatus.textContent = 'Select a worksheet first.';
    syncActionButtons();
    return;
  }

  if (!headers.length) {
    elements.detailStatus.textContent = 'This worksheet has no headers yet.';
    syncActionButtons();
    return;
  }

  const record = state.mode === 'create'
    ? Object.fromEntries(headers.map((header) => [header, '']))
    : state.selectedRow?.values || Object.fromEntries(headers.map((header) => [header, '']));

  headers.forEach((header) => {
    const field = document.createElement('div');
    field.className = 'field';
    const label = document.createElement('label');
    label.textContent = header;
    label.setAttribute('for', `field-${header}`);
    const input = document.createElement('textarea');
    input.rows = 2;
    input.id = `field-${header}`;
    input.dataset.header = header;
    input.value = record[header] ?? '';
    field.append(label, input);
    form.appendChild(field);
  });

  elements.detailStatus.textContent = state.mode === 'create'
    ? 'Create a new row. Values will be written back to the active worksheet.'
    : state.selectedRow
      ? `Row ${state.selectedRow.rowNumber} from ${state.sheetName}`
      : 'Select a row to inspect it.';
  syncActionButtons();
}

function collectFormValues() {
  const values = {};
  elements.detailForm.querySelectorAll('[data-header]').forEach((input) => {
    values[input.dataset.header] = input.value;
  });
  return values;
}

function selectRow(row) {
  state.selectedRow = row;
  state.mode = 'view';
  renderTable();
  renderDetailForm();
}

function editRow(row) {
  state.selectedRow = row;
  state.mode = 'edit';
  renderTable();
  renderDetailForm();
}

function beginCreateRow() {
  state.selectedRow = null;
  state.mode = 'create';
  renderTable();
  renderDetailForm();
}

async function saveRow() {
  if (!state.sheetName || !state.sheetData?.headers?.length) return;
  const session = getLiveSession();
  if (!session) {
    if (state.session) {
      await handleExpiredSession();
    }
    return;
  }
  const values = collectFormValues();
  setLoading(true);
  try {
    if (state.mode === 'create') {
      await appendSheetRow({
        sheetName: state.sheetName,
        accessToken: session.accessToken,
        values: buildRowValues(state.sheetData.headers, values),
      });
    } else if (state.selectedRow) {
      await updateSheetRow({
        sheetName: state.sheetName,
        accessToken: session.accessToken,
        rowNumber: state.selectedRow.rowNumber,
        values: buildRowValues(state.sheetData.headers, values),
        columnCount: Math.max(state.sheetData.headers.length, 1),
      });
    }

    state.mode = 'view';
    state.selectedRow = null;
    const reloaded = await loadSheet(state.sheetName);
    if (!reloaded) return;
    const refreshed = await refreshSheetLists();
    if (!refreshed) return;
    setNotice('Changes were synced back to Google Sheets.');
  } catch (error) {
    if (error.status === 401) {
      await handleExpiredSession('Your Google sign-in expired. Please sign in again.');
      return;
    }
    if (error.status === 403) {
      setNotice('This signed-in account can read the spreadsheet, but it cannot edit rows.', 'error');
      return;
    }
    setNotice(error.message, 'error');
  } finally {
    setLoading(false);
  }
}

async function deleteRow(row = state.selectedRow) {
  if (!row) return;
  const session = getLiveSession();
  if (!session) {
    if (state.session) {
      await handleExpiredSession();
    }
    return;
  }
  setLoading(true);
  try {
    await deleteSheetRow({
      sheetId: state.sheetData?.sheetId,
      accessToken: session.accessToken,
      rowNumber: row.rowNumber,
    });
    state.selectedRow = null;
    state.mode = 'view';
    const reloaded = await loadSheet(state.sheetName);
    if (!reloaded) return;
    const refreshed = await refreshSheetLists();
    if (!refreshed) return;
    setNotice('Row deleted and synced back to Google Sheets.');
  } catch (error) {
    if (error.status === 401) {
      await handleExpiredSession('Your Google sign-in expired. Please sign in again.');
      return;
    }
    if (error.status === 403) {
      setNotice('This signed-in account can read the spreadsheet, but it cannot delete rows.', 'error');
      return;
    }
    setNotice(error.message, 'error');
  } finally {
    setLoading(false);
  }
}

async function selectSheet(sheetName, triggerLoad = true) {
  state.sheetName = sheetName;
  state.search = '';
  state.filterField = '';
  state.filterValue = '';
  state.selectedRow = null;
  state.mode = 'view';
  elements.searchInput.value = '';
  elements.filterField.value = '';
  elements.filterValue.value = '';
  renderSheetTabs();
  renderSheetMeta();
  renderDetailForm();
  populateFilterControls();
  if (triggerLoad) {
    await loadSheet(sheetName);
  } else {
    renderTable();
  }
}

async function loadSheet(sheetName) {
  setLoading(true);
  try {
    const session = getLiveSession();
    if (!session) {
      if (state.session) {
        await handleExpiredSession();
      }
      return false;
    }

    const values = await fetchSheetValues(sheetName, session.accessToken);
    const data = parseSheetTable(values);
    setAuthorizedState();
    state.sheetData = {
      ...data,
      sheetName,
      spreadsheetTitle: state.spreadsheetTitle || 'Private spreadsheet',
      sheetId: state.sheets.find((sheet) => sheet.title === sheetName)?.sheetId ?? null,
    };
    populateFilterControls();
    renderSheetMeta();
    renderTable();
    renderDetailForm();
    return true;
  } catch (error) {
    if (error.status === 401) {
      await handleExpiredSession('Your Google sign-in expired. Please sign in again.');
      return false;
    }
    if (error.status === 403) {
      const message = 'This signed-in Google account does not have permission to open the restricted spreadsheet.';
      setNoAccessState(message);
      setNotice(message, 'error');
      return false;
    }
    setNotice(error.message, 'error');
    return false;
  } finally {
    setLoading(false);
  }
}

async function refreshSheetLists(autoLoad = false) {
  const session = getLiveSession();
  if (!session) {
    if (state.session) {
      await handleExpiredSession('Your Google sign-in expired. Please sign in again.');
    }
    return false;
  }

  try {
    const metadata = await fetchSpreadsheetMetadata(session.accessToken);
    state.spreadsheetTitle = metadata.properties?.title || state.spreadsheetTitle;
    state.sheets = (metadata.sheets || []).map((sheet) => ({
      title: sheet.properties?.title || 'Untitled',
      sheetId: sheet.properties?.sheetId ?? null,
    }));
    setAuthorizedState();
    renderWorkspaceState();
    if (!state.sheetName && state.sheets.length) {
      state.sheetName = state.sheets[0].title;
    }
    if (autoLoad && state.sheetName) {
      return await loadSheet(state.sheetName);
    }
    return true;
  } catch (error) {
    if (error.status === 401) {
      await handleExpiredSession('Your Google sign-in expired. Please sign in again.');
      return false;
    }
    if (error.status === 403) {
      const message = 'This signed-in Google account does not have permission to open the restricted spreadsheet.';
      setNoAccessState(message);
      setNotice(message, 'error');
      return false;
    }
    setNotice(error.message, 'error');
    return false;
  }
}

async function waitForGoogleIdentityLibrary() {
  if (window.google?.accounts?.oauth2) return;

  await new Promise((resolve, reject) => {
    const start = Date.now();
    const tick = () => {
      if (window.google?.accounts?.oauth2) return resolve();
      if (Date.now() - start > 10000) return reject(new Error('Google sign-in library did not load.'));
      setTimeout(tick, 50);
    };
    tick();
  });
}

function initGoogleTokenClient() {
  if (!GOOGLE_CLIENT_ID || !window.google?.accounts?.oauth2) return;

  state.authClient = window.google.accounts.oauth2.initTokenClient({
    client_id: GOOGLE_CLIENT_ID,
    scope: AUTH_SCOPE,
    include_granted_scopes: true,
    callback: handleGoogleTokenResponse,
  });
  state.authReady = true;
  setSignInState({ enabled: true, label: 'Continue with Google' });
}

function requestGoogleAccessToken() {
  if (!state.authReady || !state.authClient) {
    setNotice('Google sign-in is still preparing. Please try again in a moment.', 'error');
    return;
  }

  return new Promise((resolve, reject) => {
    state.pendingAuthRequest = { resolve, reject };
    try {
      state.authClient.requestAccessToken();
    } catch (error) {
      state.pendingAuthRequest = null;
      reject(error);
    }
  });
}

async function prepareSignIn() {
  setSignInState({ loading: true, enabled: false });
  try {
    if (!GOOGLE_CLIENT_ID || !SPREADSHEET_ID) {
      state.authReady = false;
      setSignInState({ loading: false, enabled: false, label: 'Google sign-in unavailable' });
      setNotice('App config is missing the Google client ID or spreadsheet ID.', 'error');
      return;
    }
    await waitForGoogleIdentityLibrary();
    initGoogleTokenClient();
  } catch (error) {
    state.authReady = false;
    setSignInState({ loading: false, enabled: false, label: 'Google sign-in unavailable' });
    setNotice(error.message || 'Google sign-in is unavailable.', 'error');
  }
}

async function handleGoogleTokenResponse(response) {
  if (response.error) {
    if (state.pendingAuthRequest?.reject) {
      state.pendingAuthRequest.reject(response);
    }
    state.pendingAuthRequest = null;

    if (isGoogleVerificationBlock(response)) {
      setNotice(getGoogleVerificationHelpMessage(), 'error');
      if (elements.oauthState) {
        elements.oauthState.textContent = 'Google blocked this sign-in. Add the account to Test users in Google Cloud Console.';
      }
      return;
    }
    setNotice(response.error_description || response.error || 'Google sign-in was cancelled.', 'error');
    return;
  }

  const accessToken = response.access_token;
  if (!accessToken) {
    if (state.pendingAuthRequest?.reject) {
      state.pendingAuthRequest.reject(new Error('Google sign-in did not return an access token.'));
    }
    state.pendingAuthRequest = null;
    setNotice('Google sign-in did not return an access token.', 'error');
    return;
  }

  setLoading(true);
  setSignInState({ loading: true, enabled: false, label: 'Signing in...' });

  try {
    const profilePayload = await fetchGoogleUserInfo(accessToken);
    state.session = {
      authenticated: true,
      accessToken,
      expiresAt: Date.now() + (Number(response.expires_in || 3600) * 1000),
      profile: {
        id: profilePayload.sub,
        email: profilePayload.email,
        name: profilePayload.name || profilePayload.email,
        picture: profilePayload.picture || '',
      },
    };
    persistSession(state.session);
    scheduleSessionExpiryTimer();
    setAuthorizedState();

    if (state.pendingAuthRequest?.resolve) {
      state.pendingAuthRequest.resolve(response);
    }
    state.pendingAuthRequest = null;

    setNotice('Signed in. Loading your private spreadsheet...');
    await bootstrapSignedInView();
  } catch (error) {
    if (state.pendingAuthRequest?.reject) {
      state.pendingAuthRequest.reject(error);
    }
    state.pendingAuthRequest = null;
    if (error.status === 401) {
      clearStoredSession();
    }
    setNotice(error.message || 'Google sign-in failed.', 'error');
  } finally {
    setLoading(false);
  }
}

function handleGooglePopupError(error) {
  const type = error?.type || 'unknown';
  if (type === 'access_denied') {
    setNotice(getGoogleVerificationHelpMessage(), 'error');
    if (elements.oauthState) {
      elements.oauthState.textContent = 'Google blocked this sign-in. Add the account to Test users in Google Cloud Console.';
    }
    return;
  }
  if (type === 'popup_closed') {
    setNotice('Google sign-in was closed before completion.');
    return;
  }
  if (type === 'popup_failed_to_open') {
    setNotice('The Google popup could not open. Check popup blockers.', 'error');
    return;
  }
  setNotice('Google sign-in could not start.', 'error');
}

function showSignedInShell() {
  elements.authPanel.classList.add('hidden');
  elements.dashboard.classList.remove('hidden');
  renderSessionBox();
}

function showSignedOutShell() {
  elements.dashboard.classList.add('hidden');
  elements.authPanel.classList.remove('hidden');
  renderSessionBox();
}

async function requestSignIn() {
  if (!state.authReady || !state.authClient) {
    setNotice('Google sign-in is still preparing. Please try again in a moment.', 'error');
    return;
  }
  await requestGoogleAccessToken();
}

async function logoutUser() {
  if (state.session?.accessToken && window.google?.accounts?.oauth2?.revoke) {
    try {
      await new Promise((resolve) => {
        window.google.accounts.oauth2.revoke(state.session.accessToken, () => resolve());
      });
    } catch {
      // Logout should still succeed if revoke fails.
    }
  }
  clearStoredSession();
  clearSessionExpiryTimer();
  await bootstrapSignedOutView();
  await prepareSignIn();
  setNotice('You have been logged out.');
}

function bindElements() {
  elements.sessionBox = $('sessionBox');
  elements.notice = $('notice');
  elements.authPanel = $('authPanel');
  elements.dashboard = $('dashboard');
  elements.signInButton = $('signInButton');
  elements.signInLabel = $('signInButton').querySelector('.button-label');
  elements.oauthState = $('oauthState');
  elements.spreadsheetTitle = $('spreadsheetTitle');
  elements.sheetMeta = $('sheetMeta');
  elements.sheetTabs = $('sheetTabs');
  elements.sheetTitle = $('sheetTitle');
  elements.sheetSummary = $('sheetSummary');
  elements.rowCount = $('rowCount');
  elements.dataTable = $('dataTable');
  elements.searchInput = $('searchInput');
  elements.filterField = $('filterField');
  elements.filterValue = $('filterValue');
  elements.refreshButton = $('refreshButton');
  elements.newButton = $('newButton');
  elements.detailStatus = $('detailStatus');
  elements.detailForm = $('detailForm');
  elements.saveButton = $('saveButton');
  elements.deleteButton = $('deleteButton');
  elements.cancelButton = $('cancelButton');
}

async function bootstrap() {
  bindElements();

  elements.signInButton.addEventListener('click', requestSignIn);
  elements.refreshButton.addEventListener('click', async () => {
    if (state.session) {
      await refreshSheetLists(true);
    }
  });
  elements.newButton.addEventListener('click', beginCreateRow);
  elements.saveButton.addEventListener('click', saveRow);
  elements.deleteButton.addEventListener('click', async () => {
    if (state.selectedRow && confirm(`Delete row ${state.selectedRow.rowNumber}?`)) {
      await deleteRow(state.selectedRow);
    }
  });
  elements.cancelButton.addEventListener('click', () => {
    state.mode = 'view';
    state.selectedRow = null;
    renderTable();
    renderDetailForm();
  });
  elements.searchInput.addEventListener('input', () => {
    state.search = elements.searchInput.value;
    renderTable();
  });
  elements.filterField.addEventListener('change', () => {
    state.filterField = elements.filterField.value;
    state.filterValue = '';
    populateFilterControls();
    renderTable();
  });
  elements.filterValue.addEventListener('change', () => {
    state.filterValue = elements.filterValue.value;
    renderTable();
  });

  setSignInState({ loading: true, enabled: false, label: 'Preparing Google sign-in...' });
  showSignedOutShell();
  renderWorkspaceState();

  const authSetup = prepareSignIn();
  const restoredSession = readStoredSession();
  if (restoredSession) {
    state.session = restoredSession;
    setAuthorizedState();
    scheduleSessionExpiryTimer();
    showSignedInShell();
    renderWorkspaceState();
    setNotice('Restored your saved Google session.');
    await bootstrapSignedInView({ autoLoad: false });
  } else {
    await bootstrapSignedOutView();
  }

  const params = new URLSearchParams(window.location.search);
  if (params.get('error') === 'permission-denied') {
    setNotice('The Google account you signed in with does not have access to the private spreadsheet.', 'error');
  } else if (params.get('error') === 'login-failed') {
    setNotice('Google sign-in failed. Please try again.', 'error');
  } else if (params.get('error') === 'verification-blocked') {
    setNotice(getGoogleVerificationHelpMessage(), 'error');
  }

  await authSetup;
}

bootstrap().catch((error) => setNotice(error.message, 'error'));
