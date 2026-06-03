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
  authConfig: {
    clientId: '200733782466-fimsnk62ainlnrholpjgbmc9s20jghuh.apps.googleusercontent.com',
    scope: 'openid email profile https://www.googleapis.com/auth/spreadsheets',
  },
  authClient: null,
  authReady: false,
  pendingAuthRequest: null,
};

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
  elements.refreshButton.disabled = isLoading;
  elements.newButton.disabled = isLoading || !state.sheetData?.headers?.length;
  syncActionButtons();
}

function syncActionButtons() {
  const hasSheet = Boolean(state.sheetData?.headers?.length);
  elements.saveButton.disabled = state.loading || !hasSheet;
  elements.deleteButton.disabled = state.loading || !state.selectedRow || state.mode === 'create';
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

  state.sheets.forEach((sheet) => {
    const button = document.createElement('button');
    button.className = `sheet-tab ${sheet.title === state.sheetName ? 'active' : ''}`;
    button.innerHTML = `<span>${escapeHtml(sheet.title)}</span><span>${sheet.sheetId ?? ''}</span>`;
    button.addEventListener('click', () => selectSheet(sheet.title));
    container.appendChild(button);
  });
}

function renderSheetMeta() {
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
  const values = collectFormValues();
  setLoading(true);
  try {
    if (state.mode === 'create') {
      await appendSheetRow({
        sheetName: state.sheetName,
        accessToken: state.session?.accessToken,
        values: buildRowValues(state.sheetData.headers, values),
      });
    } else if (state.selectedRow) {
      await updateSheetRow({
        sheetName: state.sheetName,
        accessToken: state.session?.accessToken,
        rowNumber: state.selectedRow.rowNumber,
        values: buildRowValues(state.sheetData.headers, values),
        columnCount: Math.max(state.sheetData.headers.length, 1),
      });
    }

    state.mode = 'view';
    state.selectedRow = null;
    await loadSheet(state.sheetName);
    await refreshSheetLists();
    setNotice('Changes were synced back to Google Sheets.');
  } catch (error) {
    setNotice(error.message, 'error');
  } finally {
    setLoading(false);
  }
}

async function deleteRow(row = state.selectedRow) {
  if (!row) return;
  setLoading(true);
  try {
    await deleteSheetRow({
      sheetId: state.sheetData?.sheetId,
      accessToken: state.session?.accessToken,
      rowNumber: row.rowNumber,
    });
    state.selectedRow = null;
    state.mode = 'view';
    await loadSheet(state.sheetName);
    await refreshSheetLists();
    setNotice('Row deleted and synced back to Google Sheets.');
  } catch (error) {
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
    const values = await fetchSheetValues(sheetName, state.session?.accessToken);
    const data = parseSheetTable(values);
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
  } catch (error) {
    if (error.status === 403) {
      setNotice('Google denied access to this spreadsheet for the signed-in account.', 'error');
      state.sheetData = null;
      state.sheetName = '';
      renderSheetMeta();
      renderTable();
      renderDetailForm();
    } else {
      setNotice(error.message, 'error');
    }
  } finally {
    setLoading(false);
  }
}

async function refreshSheetLists(autoLoad = false) {
  const metadata = await fetchSpreadsheetMetadata(state.session?.accessToken);
  state.spreadsheetTitle = metadata.properties?.title || state.spreadsheetTitle;
  state.sheets = (metadata.sheets || []).map((sheet) => ({
    title: sheet.properties?.title || 'Untitled',
    sheetId: sheet.properties?.sheetId ?? null,
  }));
  renderSessionBox();
  renderSheetTabs();
  renderSheetMeta();
  if (!state.sheetName && state.sheets.length) {
    state.sheetName = state.sheets[0].title;
  }
  if (autoLoad && state.sheetName) {
    await loadSheet(state.sheetName);
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
  if (!state.authConfig || !window.google?.accounts?.oauth2) return;

  state.authClient = window.google.accounts.oauth2.initTokenClient({
    client_id: state.authConfig.clientId,
    scope: state.authConfig.scope,
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

async function bootstrapSignedInView() {
  showSignedInShell();
  await refreshSheetLists();
  if (state.sheets.length) {
    await selectSheet(state.sheets[0].title);
  }
}

async function bootstrapSignedOutView() {
  state.session = null;
  state.sheetData = null;
  state.sheets = [];
  state.sheetName = '';
  state.selectedRow = null;
  state.mode = 'view';
  showSignedOutShell();
  setSignInState({ loading: false, enabled: false, label: 'Continue with Google' });
  await prepareSignIn();
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
  state.session = null;
  state.sheetData = null;
  state.sheets = [];
  state.sheetName = '';
  state.selectedRow = null;
  state.mode = 'view';
  state.search = '';
  state.filterField = '';
  state.filterValue = '';
  await bootstrapSignedOutView();
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
    if (state.sheetName) await loadSheet(state.sheetName);
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

  const params = new URLSearchParams(window.location.search);
  if (params.get('error') === 'permission-denied') {
    setNotice('The Google account you signed in with does not have access to the private spreadsheet.', 'error');
  } else if (params.get('error') === 'login-failed') {
    setNotice('Google sign-in failed. Please try again.', 'error');
  } else if (params.get('error') === 'verification-blocked') {
    setNotice(getGoogleVerificationHelpMessage(), 'error');
  }

  await prepareSignIn();
}

bootstrap().catch((error) => setNotice(error.message, 'error'));
