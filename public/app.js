const state = {
  session: null,
  configured: true,
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
  authConfig: null,
  authClient: null,
  authReady: false,
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

async function request(path, options = {}) {
  const response = await fetch(path, {
    credentials: 'include',
    ...options,
    headers: {
      ...(options.body ? { 'content-type': 'application/json' } : {}),
      ...(options.headers || {}),
    },
  });

  const contentType = response.headers.get('content-type') || '';
  const payload = contentType.includes('application/json')
    ? await response.json()
    : await response.text();

  if (!response.ok) {
    const error = new Error(payload?.error || payload || `Request failed (${response.status})`);
    error.status = response.status;
    error.payload = payload;
    throw error;
  }

  return payload;
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
        : state.authConfig
          ? 'Google sign-in unavailable'
          : 'Google OAuth not configured';
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
      const result = await request(`/api/sheet?sheet=${encodeURIComponent(state.sheetName)}`, {
        method: 'POST',
        body: JSON.stringify({ values }),
      });
      state.sheetData = result;
    } else if (state.selectedRow) {
      const result = await request(`/api/sheet?sheet=${encodeURIComponent(state.sheetName)}&rowNumber=${encodeURIComponent(state.selectedRow.rowNumber)}`, {
        method: 'PUT',
        body: JSON.stringify({ values }),
      });
      state.sheetData = result;
    }

    state.mode = 'view';
    state.selectedRow = null;
    await refreshSheetLists(true);
    await selectSheet(state.sheetName, false);
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
    const result = await request(`/api/sheet?sheet=${encodeURIComponent(state.sheetName)}&rowNumber=${encodeURIComponent(row.rowNumber)}`, {
      method: 'DELETE',
    });
    state.sheetData = result;
    state.selectedRow = null;
    state.mode = 'view';
    await refreshSheetLists(true);
    await selectSheet(state.sheetName, false);
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
    const data = await request(`/api/sheet?sheet=${encodeURIComponent(sheetName)}`);
    state.sheetData = data;
    state.spreadsheetTitle = data.spreadsheetTitle || state.spreadsheetTitle;
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
  const data = await request('/api/spreadsheet');
  state.configured = data.configured !== false;
  state.spreadsheetTitle = data.spreadsheetTitle || state.spreadsheetTitle;
  state.sheets = data.sheets || [];
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

async function loadAuthConfig() {
  const config = await request('/api/auth/config');
  state.authConfig = config;
  return config;
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

function initGoogleCodeClient() {
  if (!state.authConfig || !window.google?.accounts?.oauth2) return;

  state.authClient = window.google.accounts.oauth2.initCodeClient({
    client_id: state.authConfig.clientId,
    scope: state.authConfig.scope,
    ux_mode: 'popup',
    state: state.authConfig.state,
    include_granted_scopes: true,
    select_account: true,
    callback: handleGoogleCodeResponse,
    error_callback: handleGooglePopupError,
  });
  state.authReady = true;
  setSignInState({ enabled: true, label: 'Continue with Google' });
}

async function prepareSignIn() {
  setSignInState({ loading: true, enabled: false });
  try {
    await loadAuthConfig();
    await waitForGoogleIdentityLibrary();
    initGoogleCodeClient();
  } catch (error) {
    state.authReady = false;
    setSignInState({ loading: false, enabled: false, label: 'Google sign-in unavailable' });
    if (elements.oauthState) {
      elements.oauthState.textContent = 'Google OAuth not configured';
    }
    if (error.message && error.message !== 'Google OAuth is not configured. Set GOOGLE_CLIENT_ID, GOOGLE_CLIENT_SECRET, and GOOGLE_REDIRECT_URI.') {
      setNotice(error.message, 'error');
    }
  }
}

async function handleGoogleCodeResponse(response) {
  if (response.error) {
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

  setLoading(true);
  setSignInState({ loading: true, enabled: false });

  try {
    const result = await request('/auth/callback', {
      method: 'POST',
      headers: {
        'X-Requested-With': 'XMLHttpRequest',
      },
      body: JSON.stringify({
        code: response.code,
        state: response.state || state.authConfig?.state || '',
      }),
    });

    state.session = {
      authenticated: true,
      profile: result.profile,
    };

    setNotice('Signed in. Loading your private spreadsheet...');
    await bootstrapSignedInView();
  } catch (error) {
    const isDenied = error.status === 403 || error.payload?.error === 'permission_denied';
    setNotice(isDenied
      ? 'This Google account does not have access to the private spreadsheet.'
      : error.message, 'error');
    await prepareSignIn();
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

async function bootstrapSignedInView() {
  const session = await request('/api/session');
  state.session = session;
  renderSessionBox();
  elements.authPanel.classList.add('hidden');
  elements.dashboard.classList.remove('hidden');

  await refreshSheetLists();
  if (state.sheets.length) {
    await selectSheet(state.sheets[0].title);
  }
}

async function bootstrapSignedOutView() {
  state.session = { authenticated: false };
  renderSessionBox();
  elements.dashboard.classList.add('hidden');
  elements.authPanel.classList.remove('hidden');
  setSignInState({ loading: false, enabled: false, label: 'Continue with Google' });
  await prepareSignIn();
}

async function requestSignIn() {
  if (!state.authReady || !state.authClient) {
    setNotice('Google sign-in is still preparing. Please try again in a moment.', 'error');
    return;
  }
  state.authClient.requestCode();
}

async function logoutUser() {
  await fetch('/auth/logout', { method: 'POST', credentials: 'include' });
  state.authClient = null;
  state.authConfig = null;
  state.authReady = false;
  state.sheetData = null;
  state.sheets = [];
  state.sheetName = '';
  state.selectedRow = null;
  state.mode = 'view';
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

  const session = await request('/api/session');
  state.session = session;
  renderSessionBox();

  const params = new URLSearchParams(window.location.search);
  if (params.get('error') === 'permission-denied') {
    setNotice('The Google account you signed in with does not have access to the private spreadsheet.', 'error');
  } else if (params.get('error') === 'login-failed') {
    setNotice('Google sign-in failed. Please try again.', 'error');
  } else if (params.get('error') === 'verification-blocked') {
    setNotice(getGoogleVerificationHelpMessage(), 'error');
  }

  if (!session.authenticated) {
    elements.dashboard.classList.add('hidden');
    elements.authPanel.classList.remove('hidden');
    if (!session.configured) {
      setSignInState({ loading: false, enabled: false, label: 'Google OAuth not configured' });
      return;
    }
    await prepareSignIn();
    return;
  }

  await bootstrapSignedInView();
}

bootstrap().catch((error) => setNotice(error.message, 'error'));
