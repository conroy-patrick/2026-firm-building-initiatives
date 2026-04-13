// ═══════════════════════════════════════════════════════════════════
//  app.js – Beghou 2026 Firm Building Initiatives Dashboard
//  Microsoft Graph / SharePoint integration + interactive dashboard
// ═══════════════════════════════════════════════════════════════════

// ── Global State ──────────────────────────────────────────────────
let msalInstance     = null;
let currentAccount   = null;
let graphToken       = null;
let siteId           = null;
let initiativeListId = null;
let updatesListId    = null;
let initiatives      = [];
let updates          = [];   // full audit log across all initiatives
let selectedInitiative = null;
let isAuthenticated  = false;
let activeFilter     = 'all';

// ── Category metadata ─────────────────────────────────────────────
const CATEGORIES = [
  { key: 'Revenue & Operations', cssKey: 'revenue', color: '#E60F65', lightBg: '#FFF0F5' },
  { key: 'Product & Tech',       cssKey: 'tech',    color: '#7C3AED', lightBg: '#F5F3FF' },
  { key: 'Systems & People',     cssKey: 'people',  color: '#0EA5E9', lightBg: '#F0F9FF' }
];

const STATUS_META = {
  'On Track':    { cls: 'on-track',    barCls: 'bar-on-track',    badgeCls: 'badge-on-track',    icon: '✓', hex: '#22C55E', bg: '#DCFCE7', fg: '#16A34A' },
  'At Risk':     { cls: 'at-risk',     barCls: 'bar-at-risk',     badgeCls: 'badge-at-risk',     icon: '!', hex: '#F59E0B', bg: '#FEF3C7', fg: '#D97706' },
  'Behind':      { cls: 'behind',      barCls: 'bar-behind',      badgeCls: 'badge-behind',      icon: '✕', hex: '#EF4444', bg: '#FEE2E2', fg: '#DC2626' },
  'Not Started': { cls: 'not-started', barCls: 'bar-not-started', badgeCls: 'badge-not-started', icon: '○', hex: '#94A3B8', bg: '#F1F5F9', fg: '#64748B' },
  'Complete':    { cls: 'complete',    barCls: 'bar-complete',    badgeCls: 'badge-complete',    icon: '★', hex: '#3B82F6', bg: '#DBEAFE', fg: '#2563EB' }
};

function getCat(key)    { return CATEGORIES.find(c => c.key === key) || CATEGORIES[0]; }
function getStatus(key) { return STATUS_META[key] || STATUS_META['Not Started']; }

// ── MSAL Authentication ───────────────────────────────────────────
function initMSAL() {
  if (!CONFIG.clientId || CONFIG.clientId === 'YOUR_AZURE_AD_CLIENT_ID') {
    console.info('Auth not configured. Running in demo mode.');
    loadDemoMode();
    return;
  }

  const msalConfig = {
    auth: {
      clientId:    CONFIG.clientId,
      authority:   `https://login.microsoftonline.com/${CONFIG.tenantId}`,
      redirectUri: CONFIG.redirectUri
    },
    cache: { cacheLocation: 'sessionStorage', storeAuthStateInCookie: false }
  };

  try {
    msalInstance = new msal.PublicClientApplication(msalConfig);
    msalInstance.handleRedirectPromise()
      .then(onAuthResponse)
      .catch(e => { console.warn('MSAL redirect error:', e); loadDemoMode(); });
  } catch (e) {
    console.warn('MSAL init failed:', e);
    loadDemoMode();
  }
}

async function onAuthResponse(response) {
  if (response) {
    currentAccount = response.account;
    await completeSignIn();
  } else {
    const accounts = msalInstance?.getAllAccounts() || [];
    if (accounts.length > 0) {
      currentAccount = accounts[0];
      await completeSignIn();
    } else {
      loadDemoMode();
    }
  }
}

async function handleAuth() {
  if (isAuthenticated) { signOut(); return; }

  if (!CONFIG.clientId || CONFIG.clientId === 'YOUR_AZURE_AD_CLIENT_ID') {
    showToast('Auth not configured — see README.md for setup steps.', 'error');
    return;
  }

  try {
    const result = await msalInstance.loginPopup({
      scopes: ['User.Read', 'Sites.ReadWrite.All']
    });
    currentAccount = result.account;
    await completeSignIn();
  } catch (e) {
    console.error('Login error:', e);
    showToast('Sign-in failed. Please try again.', 'error');
  }
}

function signOut() {
  msalInstance?.logoutPopup({ account: currentAccount });
  currentAccount   = null;
  graphToken       = null;
  isAuthenticated  = false;
  loadDemoMode();
}

async function getToken() {
  if (!msalInstance || !currentAccount) return null;
  try {
    const r = await msalInstance.acquireTokenSilent({
      scopes: ['Sites.ReadWrite.All'],
      account: currentAccount
    });
    return r.accessToken;
  } catch {
    try {
      const r = await msalInstance.acquireTokenPopup({
        scopes: ['Sites.ReadWrite.All'],
        account: currentAccount
      });
      return r.accessToken;
    } catch (e) {
      console.error('Token error:', e);
      return null;
    }
  }
}

async function completeSignIn() {
  isAuthenticated = true;
  updateAuthUI();
  document.getElementById('authBanner').style.display = 'none';

  graphToken = await getToken();
  if (!graphToken) { loadDemoMode(); showToast('Could not acquire token. Showing demo data.', 'error'); return; }

  await resolveSharePointIds();
  await loadLiveData();
}

// ── Microsoft Graph / SharePoint ──────────────────────────────────
async function gFetch(path, opts = {}) {
  const token = graphToken || await getToken();
  if (!token) throw new Error('Not authenticated');

  const res = await fetch(CONFIG.graphEndpoint + path, {
    ...opts,
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
      ...(opts.headers || {})
    }
  });

  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err?.error?.message || `HTTP ${res.status}: ${res.statusText}`);
  }

  return res.status === 204 ? null : res.json();
}

async function resolveSharePointIds() {
  try {
    const site = await gFetch(`/sites/${CONFIG.siteHostname}:${CONFIG.sitePath}`);
    siteId = site.id;

    const listsRes = await gFetch(`/sites/${siteId}/lists?$select=id,name,displayName`);
    const lists    = listsRes?.value || [];

    const findList = name => lists.find(l =>
      l.name?.toLowerCase() === name.toLowerCase() ||
      l.displayName?.toLowerCase() === name.toLowerCase()
    );

    initiativeListId = findList(CONFIG.initiativeListName)?.id || null;
    updatesListId    = findList(CONFIG.updatesListName)?.id    || null;

    if (!initiativeListId) {
      console.warn(`List "${CONFIG.initiativeListName}" not found.`,
        'Available lists:', lists.map(l => l.displayName).join(', '));
    }
  } catch (e) {
    console.error('SharePoint resolve error:', e);
    showToast(`SharePoint connection error: ${e.message}`, 'error');
  }
}

async function loadLiveData() {
  if (!siteId || !initiativeListId) { loadDemoMode(); return; }

  try {
    setLoadingState(true);

    // ── Fetch initiatives ──────────────────────────────────────
    const fields = [
      'Title','Category','Description','AccountablePartner','Builders',
      'ExecSponsor','Status','PercentComplete','Q1Goals','Q2Goals','Q3Goals',
      'LatestComment','LastUpdatedBy','Modified'
    ].join(',');

    const itemsRes = await gFetch(
      `/sites/${siteId}/lists/${initiativeListId}/items` +
      `?$expand=fields($select=${fields})&$top=50&$orderby=fields/Category`
    );

    const spItems = itemsRes?.value || [];

    if (spItems.length === 0) {
      // List exists but empty — use seed data + offer to seed
      initiatives = JSON.parse(JSON.stringify(SEED_INITIATIVES));
      showToast('SharePoint list appears empty. Showing demo data. Seed the list from SharePoint.', 'error');
    } else {
      initiatives = spItems.map((item, idx) => ({
        id:                 item.id || idx + 1,
        spId:               item.id,
        title:              item.fields?.Title || '',
        category:           item.fields?.Category || 'Uncategorized',
        description:        item.fields?.Description || '',
        accountablePartner: item.fields?.AccountablePartner || '',
        builders:           item.fields?.Builders || '',
        execSponsor:        item.fields?.ExecSponsor || '',
        status:             item.fields?.Status || 'Not Started',
        percentComplete:    parseInt(item.fields?.PercentComplete) || 0,
        q1Goals:            item.fields?.Q1Goals || '',
        q2Goals:            item.fields?.Q2Goals || '',
        q3Goals:            item.fields?.Q3Goals || '',
        latestComment:      item.fields?.LatestComment || '',
        lastUpdatedBy:      item.fields?.LastUpdatedBy || '',
        modified:           item.fields?.Modified || null
      }));
    }

    // ── Fetch update log ───────────────────────────────────────
    if (updatesListId) {
      const logFields = 'Title,UpdatedBy,Status,PercentComplete,Comment,Created';
      const logRes = await gFetch(
        `/sites/${siteId}/lists/${updatesListId}/items` +
        `?$expand=fields($select=${logFields})&$top=200&$orderby=fields/Created desc`
      );
      updates = (logRes?.value || []).map(item => ({
        initiativeTitle: item.fields?.Title || '',
        updatedBy:       item.fields?.UpdatedBy || '',
        status:          item.fields?.Status || '',
        percentComplete: item.fields?.PercentComplete || 0,
        comment:         item.fields?.Comment || '',
        created:         item.fields?.Created || null
      }));
    }

    renderDashboard();
    setLoadingState(false);
    setLastUpdated();

  } catch (e) {
    console.error('Load error:', e);
    showToast(`Error loading data: ${e.message}`, 'error');
    loadDemoMode();
  }
}

async function saveUpdate(initiative, name, status, pct, comment) {
  // ── Local update (applies in both live & demo mode) ────────────
  initiative.status          = status;
  initiative.percentComplete = pct;
  if (comment) initiative.latestComment = comment;
  initiative.lastUpdatedBy   = name;
  initiative.modified        = new Date().toISOString();

  const newEntry = {
    initiativeTitle: initiative.title,
    updatedBy:       name,
    status,
    percentComplete: pct,
    comment,
    created:         new Date().toISOString()
  };
  updates.unshift(newEntry);

  // ── Live SharePoint writes ─────────────────────────────────────
  if (!isAuthenticated || !siteId) return; // demo mode stops here

  try {
    if (initiative.spId && initiativeListId) {
      await gFetch(
        `/sites/${siteId}/lists/${initiativeListId}/items/${initiative.spId}/fields`,
        {
          method: 'PATCH',
          body: JSON.stringify({
            Status:          status,
            PercentComplete: pct,
            ...(comment ? { LatestComment: comment } : {}),
            LastUpdatedBy:   name
          })
        }
      );
    }

    if (updatesListId) {
      await gFetch(`/sites/${siteId}/lists/${updatesListId}/items`, {
        method: 'POST',
        body: JSON.stringify({
          fields: {
            Title:           initiative.title,
            UpdatedBy:       name,
            Status:          status,
            PercentComplete: pct,
            Comment:         comment
          }
        })
      });
    }
  } catch (e) {
    console.error('Save error:', e);
    throw e;
  }
}

// ── Demo Mode ─────────────────────────────────────────────────────
function loadDemoMode() {
  initiatives = JSON.parse(JSON.stringify(SEED_INITIATIVES));
  updates     = JSON.parse(JSON.stringify(SEED_UPDATES));
  renderDashboard();
  setLastUpdated('Demo data — sign in to load live SharePoint data');
}

// ── Rendering ─────────────────────────────────────────────────────
function renderDashboard() {
  renderSummaryCards();
  renderChart();
  renderCards(activeFilter);
  renderOverallStats();
}

function renderSummaryCards() {
  const grid = document.getElementById('summaryGrid');

  const html = CATEGORIES.map(cat => {
    const items    = initiatives.filter(i => i.category === cat.key);
    const avg      = items.length
      ? Math.round(items.reduce((s, i) => s + i.percentComplete, 0) / items.length) : 0;
    const onTrack  = items.filter(i => i.status === 'On Track' || i.status === 'Complete').length;
    const atRisk   = items.filter(i => i.status === 'At Risk').length;
    const behind   = items.filter(i => i.status === 'Behind').length;
    const expected = CONFIG.expectedProgress;

    const delta    = avg - expected;
    const deltaStr = delta > 0 ? `+${delta}% ahead` : delta < 0 ? `${delta}% behind` : 'on pace';
    const deltaClr = delta >= 0 ? '#16A34A' : delta > -10 ? '#D97706' : '#DC2626';

    return `
      <div class="summary-card" style="border-top: 3px solid ${cat.color}">
        <div class="summary-card-header">
          <div class="cat-dot" style="background:${cat.color}"></div>
          <div class="summary-card-label">${cat.key}</div>
        </div>
        <div class="summary-stats-row">
          <div>
            <div class="summary-pct">${avg}<span class="summary-pct-unit">%</span></div>
            <div class="summary-pct-sub">avg complete</div>
          </div>
          <div class="summary-delta" style="color:${deltaClr}">${deltaStr}<br><span style="font-size:11px; font-weight:400; opacity:.8">vs expected ${expected}%</span></div>
        </div>
        <div class="summary-progress-track">
          <div class="summary-progress-fill" style="width:${avg}%; background:${cat.color}"></div>
          <div class="summary-expected-mark" style="left:${expected}%" title="Expected: ${expected}%"></div>
        </div>
        <div class="summary-footer">
          <span style="color:var(--gray-500)">${items.length} initiative${items.length !== 1 ? 's' : ''}</span>
          <span>
            <span style="color:#16A34A; font-weight:600">${onTrack} on track</span>
            ${atRisk  ? ` <span style="color:#D97706">· ${atRisk} at risk</span>` : ''}
            ${behind  ? ` <span style="color:#DC2626">· ${behind} behind</span>` : ''}
          </span>
        </div>
      </div>`;
  }).join('');

  grid.innerHTML = html;
}

function renderChart() {
  const container = document.getElementById('progressChart');
  const expected  = CONFIG.expectedProgress;

  // ── Axis header ────────────────────────────────────────────────
  let html = `
    <div class="chart-axis-row">
      <div class="chart-name-col"></div>
      <div class="chart-bar-col" style="position:relative; height:22px;">
        ${[0, 25, 50, 75, 100].map(v => `
          <div style="position:absolute; left:${v}%; transform:translateX(-50%); font-size:10px; color:var(--gray-400); font-weight:600;">
            ${v === 0 ? '0%' : v === 100 ? '100%' : 'Q' + (v / 25)}
          </div>`).join('')}
        <div style="position:absolute; top:14px; left:0; right:0; height:1px; background:var(--gray-200);"></div>
        ${[25, 50, 75].map(v => `
          <div style="position:absolute; top:14px; left:${v}%; width:1px; height:6px; background:var(--gray-300);"></div>`).join('')}
        <!-- Expected line in header -->
        <div style="position:absolute; top:0; left:${expected}%; transform:translateX(-50%); 
                    font-size:10px; color:var(--pink); font-weight:700; white-space:nowrap;">
          ▼ Today (${expected}%)
        </div>
      </div>
      <div class="chart-status-col"></div>
      <div class="chart-pct-col"></div>
    </div>`;

  // ── Initiative rows by category ────────────────────────────────
  CATEGORIES.forEach(cat => {
    const items = initiatives.filter(i => i.category === cat.key);
    if (!items.length) return;

    html += `
      <div class="chart-category-header">
        <div class="chart-category-label" style="color:${cat.color}">${cat.key}</div>
        <div class="chart-category-rule" style="background:${cat.color}20"></div>
      </div>`;

    items.forEach(init => {
      const sm  = getStatus(init.status);
      const pct = Math.max(0, Math.min(100, init.percentComplete));
      const barW = pct > 0 ? Math.max(pct, 2) : 0;

      html += `
        <div class="chart-row" onclick="openModal('${init.id}')" title="Click to view details or submit an update">
          <div class="chart-name-col">
            <div class="chart-name-text" title="${init.title}">${init.title}</div>
            <div class="chart-name-owner">${init.accountablePartner || 'TBD'}</div>
          </div>
          <div class="chart-bar-col" style="position:relative;">
            <!-- Quarter grid lines -->
            ${[25, 50, 75].map(v => `
              <div style="position:absolute; top:0; left:${v}%; width:1px; height:100%;
                          background:var(--gray-200); z-index:0; pointer-events:none;"></div>`).join('')}
            <!-- Expected marker -->
            <div class="chart-expected-marker" style="left:${expected}%"></div>
            <!-- Progress bar background -->
            <div class="chart-bar-bg">
              <!-- Filled bar -->
              <div class="chart-bar-fill ${sm.barCls}" style="width:${barW}%">
                ${pct >= 12 ? `<span class="chart-bar-label">${pct}%</span>` : ''}
              </div>
            </div>
          </div>
          <div class="chart-status-col">
            <span class="chart-status-badge ${sm.badgeCls}">${init.status}</span>
          </div>
          <div class="chart-pct-col">${pct}%</div>
        </div>`;
    });
  });

  container.innerHTML = html;
}

function renderCards(filter = 'all') {
  activeFilter = filter;
  const grid = document.getElementById('cardsGrid');

  const filtered = filter === 'all'
    ? initiatives
    : initiatives.filter(i => i.category === filter);

  if (!filtered.length) {
    grid.innerHTML = `<div style="grid-column:1/-1; text-align:center; padding:60px; color:var(--gray-400); font-size:14px;">No initiatives found.</div>`;
    return;
  }

  grid.innerHTML = filtered.map(init => {
    const cat  = getCat(init.category);
    const sm   = getStatus(init.status);
    const pct  = init.percentComplete;
    const hasComment = !!init.latestComment;

    return `
      <div class="initiative-card" onclick="openModal('${init.id}')">
        <div class="card-accent-bar" style="background:${cat.color}"></div>
        <div class="card-body">
          <div class="card-top-row">
            <div class="card-title">${init.title}</div>
            <span class="status-badge ${sm.badgeCls}">${init.status}</span>
          </div>

          <div class="card-meta-row">
            <div class="card-meta-item">
              <span class="meta-label">Partner</span>
              <span class="meta-val">${init.accountablePartner || 'TBD'}</span>
            </div>
            <div class="card-meta-divider">·</div>
            <div class="card-meta-item">
              <span class="meta-label">Builders</span>
              <span class="meta-val">${init.builders || 'TBD'}</span>
            </div>
            <div class="card-meta-divider">·</div>
            <div class="card-meta-item">
              <span class="meta-label">Sponsor</span>
              <span class="meta-val">${init.execSponsor || 'TBD'}</span>
            </div>
          </div>

          <div class="card-description">${init.description}</div>

          <div class="card-progress-row">
            <div class="card-progress-track">
              <div class="card-progress-expected" style="left:${CONFIG.expectedProgress}%" title="Expected: ${CONFIG.expectedProgress}%"></div>
              <div class="card-progress-fill ${sm.barCls}" style="width:${pct}%"></div>
            </div>
            <div class="card-progress-pct">${pct}%</div>
          </div>

          <div class="card-update-box ${hasComment ? '' : 'card-update-empty'}">
            ${hasComment
              ? `<span class="update-icon">💬</span> ${init.latestComment}
                 ${init.lastUpdatedBy
                   ? `<span class="update-byline">— ${init.lastUpdatedBy}${init.modified ? ' · ' + fmtDate(init.modified) : ''}</span>`
                   : ''}`
              : `<span style="font-style:italic; color:var(--gray-400)">No updates yet — be the first to submit one.</span>`}
          </div>

          <div class="card-footer-row">
            <span class="cat-pill" style="background:${cat.lightBg}; color:${cat.color}">${init.category}</span>
            <button class="btn-card-update" onclick="event.stopPropagation(); openModal('${init.id}')">
              + Submit Update
            </button>
          </div>
        </div>
      </div>`;
  }).join('');
}

// ── Modal ─────────────────────────────────────────────────────────
function openModal(initiativeId) {
  selectedInitiative = initiatives.find(i => String(i.id) === String(initiativeId));
  if (!selectedInitiative) return;

  const init = selectedInitiative;
  const cat  = getCat(init.category);
  const sm   = getStatus(init.status);

  // Header
  document.getElementById('mCatBadge').textContent  = init.category;
  document.getElementById('mCatBadge').style.color  = cat.color;
  document.getElementById('mTitle').textContent      = init.title;
  document.getElementById('mMeta').innerHTML = `
    <div class="modal-meta-item"><strong>Partner:</strong> ${init.accountablePartner || 'TBD'}</div>
    <div class="modal-meta-item"><strong>Builders:</strong> ${init.builders || 'TBD'}</div>
    <div class="modal-meta-item"><strong>Exec Sponsor:</strong> ${init.execSponsor || 'TBD'}</div>
  `;

  // Description
  document.getElementById('mDesc').textContent = init.description;

  // Status circle
  const circle = document.getElementById('mStatusCircle');
  circle.textContent      = sm.icon;
  circle.style.background = sm.bg;
  circle.style.color      = sm.fg;
  document.getElementById('mStatusVal').textContent = init.status;
  document.getElementById('mPctVal').textContent    = `${init.percentComplete}%`;

  // Progress bar in modal
  const mBar = document.getElementById('mProgressFill');
  mBar.className   = `card-progress-fill ${sm.barCls}`;
  mBar.style.width = `${init.percentComplete}%`;
  document.getElementById('mExpected').style.left = `${CONFIG.expectedProgress}%`;

  // Timeline
  document.getElementById('mTimeline').innerHTML = [
    { q: 'Q1', text: init.q1Goals },
    { q: 'Q2', text: init.q2Goals },
    { q: 'Q3', text: init.q3Goals }
  ].map(({ q, text }) => `
    <div class="timeline-item">
      <div class="timeline-q" style="color:${cat.color}">${q}</div>
      <div class="timeline-text">${text || 'Not yet defined'}</div>
    </div>`).join('');

  // Pre-fill form
  document.getElementById('fStatus').value  = init.status;
  document.getElementById('fPct').value     = init.percentComplete;
  document.getElementById('fPctVal').textContent = init.percentComplete;
  document.getElementById('fComment').value = '';
  document.getElementById('fName').value    = '';
  document.getElementById('submitBtn').disabled = false;
  document.getElementById('submitBtn').textContent = 'Submit Update';

  // History
  renderHistory(init.title);

  // Open
  document.getElementById('mOverlay').classList.add('active');
  document.getElementById('mPanel').classList.add('active');
  document.body.style.overflow = 'hidden';
}

function closeModal() {
  document.getElementById('mOverlay').classList.remove('active');
  document.getElementById('mPanel').classList.remove('active');
  document.body.style.overflow = '';
  selectedInitiative = null;
}

function renderHistory(title) {
  const tbody  = document.getElementById('historyBody');
  const relevant = updates.filter(u => u.initiativeTitle === title);

  if (!relevant.length) {
    tbody.innerHTML = `<tr><td colspan="5" class="history-empty">No updates submitted yet for this initiative.</td></tr>`;
    return;
  }

  tbody.innerHTML = relevant.map(u => {
    const sm = getStatus(u.status);
    return `
      <tr>
        <td style="white-space:nowrap; color:var(--gray-400); font-size:11px">${fmtDate(u.created)}</td>
        <td style="font-weight:500; color:var(--navy)">${u.updatedBy}</td>
        <td><span class="status-badge ${sm.badgeCls}" style="font-size:10px">${u.status}</span></td>
        <td style="font-weight:700; color:var(--navy); text-align:center">${u.percentComplete}%</td>
        <td style="font-size:12px; color:var(--gray-600); max-width:180px; word-break:break-word">${u.comment || '—'}</td>
      </tr>`;
  }).join('');
}

async function submitUpdate() {
  const name    = document.getElementById('fName').value.trim();
  const status  = document.getElementById('fStatus').value;
  const pct     = parseInt(document.getElementById('fPct').value);
  const comment = document.getElementById('fComment').value.trim();

  if (!name)   { showToast('Please enter your name before submitting.', 'error');   return; }
  if (!status) { showToast('Please select a status before submitting.', 'error'); return; }

  const btn = document.getElementById('submitBtn');
  btn.disabled     = true;
  btn.innerHTML    = '<span class="spinner"></span> Submitting…';

  try {
    await saveUpdate(selectedInitiative, name, status, pct, comment);

    // Refresh modal status display
    const sm = getStatus(status);
    document.getElementById('mStatusCircle').textContent      = sm.icon;
    document.getElementById('mStatusCircle').style.background = sm.bg;
    document.getElementById('mStatusCircle').style.color      = sm.fg;
    document.getElementById('mStatusVal').textContent          = status;
    document.getElementById('mPctVal').textContent             = `${pct}%`;

    const mBar  = document.getElementById('mProgressFill');
    mBar.className   = `card-progress-fill ${sm.barCls}`;
    mBar.style.width = `${pct}%`;

    renderHistory(selectedInitiative.title);
    renderDashboard(); // refresh chart + cards

    showToast('✓ Update submitted successfully!', 'success');

    document.getElementById('fComment').value = '';
    document.getElementById('fName').value    = '';
  } catch (e) {
    showToast(`Save failed: ${e.message}`, 'error');
  } finally {
    btn.disabled     = false;
    btn.textContent  = 'Submit Update';
  }
}

// ── Filters ───────────────────────────────────────────────────────
function filterCards(filter, btn) {
  document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  renderCards(filter);
}

// ── Utility ───────────────────────────────────────────────────────
function fmtDate(str) {
  if (!str) return '—';
  try {
    return new Date(str).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
  } catch { return str; }
}

function setLastUpdated(text) {
  document.getElementById('lastUpdated').textContent = text ||
    `Updated ${new Date().toLocaleString('en-US', { month:'short', day:'numeric', hour:'2-digit', minute:'2-digit' })}`;
}

function setLoadingState(on) {
  document.getElementById('progressChart').style.opacity = on ? '0.4' : '1';
}

function showToast(msg, type = 'success') {
  const t = document.getElementById('toast');
  t.textContent   = msg;
  t.className     = `toast show ${type}`;
  setTimeout(() => { t.className = 'toast'; }, 4500);
}

function updateAuthUI() {
  const section = document.getElementById('authSection');
  if (!isAuthenticated || !currentAccount) {
    section.innerHTML = `<button class="btn-auth" onclick="handleAuth()">Sign In with Microsoft</button>`;
    return;
  }
  const name     = currentAccount.name || currentAccount.username || 'User';
  const initials = name.split(' ').map(n => n[0] || '').join('').toUpperCase().slice(0, 2) || 'U';
  section.innerHTML = `
    <div class="user-badge">
      <div class="user-avatar">${initials}</div>
      <span class="user-name">${name.split(' ')[0]}</span>
    </div>
    <button class="btn-auth btn-auth-ghost" onclick="handleAuth()">Sign Out</button>`;
}

// ── Overall stats bar ─────────────────────────────────────────────
function renderOverallStats() {
  const total    = initiatives.length;
  const avgPct   = total ? Math.round(initiatives.reduce((s, i) => s + i.percentComplete, 0) / total) : 0;
  const onTrack  = initiatives.filter(i => i.status === 'On Track').length;
  const atRisk   = initiatives.filter(i => i.status === 'At Risk').length;
  const behind   = initiatives.filter(i => i.status === 'Behind').length;
  const complete = initiatives.filter(i => i.status === 'Complete').length;

  document.getElementById('statTotal').textContent    = total;
  document.getElementById('statAvg').textContent      = `${avgPct}%`;
  document.getElementById('statOnTrack').textContent  = onTrack;
  document.getElementById('statAtRisk').textContent   = atRisk;
  document.getElementById('statBehind').textContent   = behind;
  document.getElementById('statComplete').textContent = complete;
}

// ── Bootstrap ─────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  // Slider live display
  document.getElementById('fPct').addEventListener('input', function () {
    document.getElementById('fPctVal').textContent = this.value;
  });

  initMSAL();

  // Auto-refresh when live
  setInterval(() => {
    if (isAuthenticated && siteId) loadLiveData();
  }, CONFIG.refreshIntervalMs);
});
