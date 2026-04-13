// ═══════════════════════════════════════════════════════════════════
//  app.js – Beghou 2026 Firm Building Initiatives Dashboard
//  Unit of tracking: Workstreams, grouped by Initiative
// ═══════════════════════════════════════════════════════════════════

// ── Global State ──────────────────────────────────────────────────
let msalInstance      = null;
let currentAccount    = null;
let graphToken        = null;
let siteId            = null;
let workstreamListId  = null;
let updatesListId     = null;
let workstreams       = [];
let updates           = [];
let selectedWorkstream = null;
let isAuthenticated   = false;
let activeFilter      = 'all';

// ── Category metadata ─────────────────────────────────────────────
const CATEGORIES = [
  { key: 'Revenue & Operations', cssKey: 'revenue', color: '#E60F65', lightBg: '#FFF0F5' },
  { key: 'Product & Tech',       cssKey: 'tech',    color: '#7C3AED', lightBg: '#F5F3FF' },
  { key: 'Systems & People',     cssKey: 'people',  color: '#0EA5E9', lightBg: '#F0F9FF' }
];

const STATUS_META = {
  'On Track':    { cls:'on-track',    barCls:'bar-on-track',    badgeCls:'badge-on-track',    icon:'✓', hex:'#22C55E', bg:'#DCFCE7', fg:'#16A34A' },
  'At Risk':     { cls:'at-risk',     barCls:'bar-at-risk',     badgeCls:'badge-at-risk',     icon:'!', hex:'#F59E0B', bg:'#FEF3C7', fg:'#D97706' },
  'Behind':      { cls:'behind',      barCls:'bar-behind',      badgeCls:'badge-behind',      icon:'✕', hex:'#EF4444', bg:'#FEE2E2', fg:'#DC2626' },
  'Not Started': { cls:'not-started', barCls:'bar-not-started', badgeCls:'badge-not-started', icon:'○', hex:'#94A3B8', bg:'#F1F5F9', fg:'#64748B' },
  'Complete':    { cls:'complete',    barCls:'bar-complete',    badgeCls:'badge-complete',    icon:'★', hex:'#3B82F6', bg:'#DBEAFE', fg:'#2563EB' }
};

function getCat(key)    { return CATEGORIES.find(c => c.key === key) || CATEGORIES[0]; }
function getStatus(key) { return STATUS_META[key] || STATUS_META['Not Started']; }

function getInitiativeWorkstreams(initiativeTitle) {
  return workstreams.filter(w => w.initiative === initiativeTitle);
}

function calcInitiativeAvg(initiativeTitle) {
  const ws = getInitiativeWorkstreams(initiativeTitle);
  if (!ws.length) return 0;
  return Math.round(ws.reduce((s, w) => s + w.percentComplete, 0) / ws.length);
}

function calcInitiativeStatus(initiativeTitle) {
  const ws = getInitiativeWorkstreams(initiativeTitle);
  if (ws.some(w => w.status === 'Behind'))      return 'Behind';
  if (ws.some(w => w.status === 'At Risk'))     return 'At Risk';
  if (ws.every(w => w.status === 'Complete'))   return 'Complete';
  if (ws.every(w => w.status === 'Not Started'))return 'Not Started';
  return 'On Track';
}

// ── MSAL Auth ─────────────────────────────────────────────────────
function initMSAL() {
  if (!CONFIG.clientId || CONFIG.clientId === 'YOUR_AZURE_AD_CLIENT_ID') {
    loadDemoMode(); return;
  }
  const msalConfig = {
    auth: { clientId: CONFIG.clientId, authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`, redirectUri: CONFIG.redirectUri },
    cache: { cacheLocation: 'sessionStorage', storeAuthStateInCookie: false }
  };
  try {
    msalInstance = new msal.PublicClientApplication(msalConfig);
    msalInstance.handleRedirectPromise().then(onAuthResponse).catch(() => loadDemoMode());
  } catch { loadDemoMode(); }
}

async function onAuthResponse(response) {
  if (response) { currentAccount = response.account; await completeSignIn(); }
  else {
    const accounts = msalInstance?.getAllAccounts() || [];
    if (accounts.length > 0) { currentAccount = accounts[0]; await completeSignIn(); }
    else loadDemoMode();
  }
}

async function handleAuth() {
  if (isAuthenticated) { signOut(); return; }
  if (!CONFIG.clientId || CONFIG.clientId === 'YOUR_AZURE_AD_CLIENT_ID') {
    showToast('Auth not configured.', 'error'); return;
  }
  try {
    const result = await msalInstance.loginPopup({ scopes: ['User.Read', 'Sites.ReadWrite.All'] });
    currentAccount = result.account;
    await completeSignIn();
  } catch (e) { showToast('Sign-in failed. Please try again.', 'error'); }
}

function signOut() {
  msalInstance?.logoutPopup({ account: currentAccount });
  currentAccount = null; graphToken = null; isAuthenticated = false;
  loadDemoMode();
}

async function getToken() {
  if (!msalInstance || !currentAccount) return null;
  try {
    const r = await msalInstance.acquireTokenSilent({ scopes: ['Sites.ReadWrite.All'], account: currentAccount });
    return r.accessToken;
  } catch {
    try {
      const r = await msalInstance.acquireTokenPopup({ scopes: ['Sites.ReadWrite.All'], account: currentAccount });
      return r.accessToken;
    } catch { return null; }
  }
}

async function completeSignIn() {
  isAuthenticated = true;
  updateAuthUI();
  document.getElementById('authBanner').style.display = 'none';
  graphToken = await getToken();
  if (!graphToken) { loadDemoMode(); showToast('Could not acquire token.', 'error'); return; }
  await resolveSharePointIds();
  await loadLiveData();
}

// ── Graph / SharePoint ────────────────────────────────────────────
async function gFetch(path, opts = {}) {
  const token = graphToken || await getToken();
  if (!token) throw new Error('Not authenticated');
  const res = await fetch(CONFIG.graphEndpoint + path, {
    ...opts,
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json', ...(opts.headers || {}) }
  });
  if (!res.ok) { const err = await res.json().catch(() => ({})); throw new Error(err?.error?.message || `HTTP ${res.status}`); }
  return res.status === 204 ? null : res.json();
}

async function resolveSharePointIds() {
  try {
    const site = await gFetch(`/sites/${CONFIG.siteHostname}:${CONFIG.sitePath}`);
    siteId = site.id;
    const listsRes = await gFetch(`/sites/${siteId}/lists?$select=id,name,displayName`);
    const lists = listsRes?.value || [];
    const findList = name => lists.find(l => l.name?.toLowerCase() === name.toLowerCase() || l.displayName?.toLowerCase() === name.toLowerCase());
    workstreamListId = findList(CONFIG.workstreamListName)?.id || null;
    updatesListId    = findList(CONFIG.updatesListName)?.id    || null;
    if (!workstreamListId) console.warn(`List "${CONFIG.workstreamListName}" not found. Available:`, lists.map(l => l.displayName).join(', '));
  } catch (e) {
    showToast(`SharePoint error: ${e.message}`, 'error');
  }
}

async function loadLiveData() {
  if (!siteId || !workstreamListId) { loadDemoMode(); return; }
  try {
    setLoadingState(true);
    const fields = 'Title,Initiative,Category,Owner,Status,PercentComplete,Q1Goals,Q2Goals,Q3Goals,LatestComment,LastUpdatedBy,Modified';
    const res = await gFetch(`/sites/${siteId}/lists/${workstreamListId}/items?$expand=fields($select=${fields})&$top=200`);
    const items = res?.value || [];
    if (!items.length) {
      workstreams = JSON.parse(JSON.stringify(SEED_WORKSTREAMS));
      showToast('SharePoint list empty — showing demo data.', 'error');
    } else {
      workstreams = items.map((item, idx) => ({
        id: item.id || idx + 1,
        spId: item.id,
        title: item.fields?.Title || '',
        initiative: item.fields?.Initiative || '',
        category: item.fields?.Category || '',
        owner: item.fields?.Owner || '',
        status: item.fields?.Status || 'Not Started',
        percentComplete: parseInt(item.fields?.PercentComplete) || 0,
        q1Goals: item.fields?.Q1Goals || '',
        q2Goals: item.fields?.Q2Goals || '',
        q3Goals: item.fields?.Q3Goals || '',
        latestComment: item.fields?.LatestComment || '',
        lastUpdatedBy: item.fields?.LastUpdatedBy || '',
        modified: item.fields?.Modified || null
      }));
    }
    if (updatesListId) {
      const logRes = await gFetch(`/sites/${siteId}/lists/${updatesListId}/items?$expand=fields($select=Title,UpdatedBy,Status,PercentComplete,Comment,Created)&$top=500&$orderby=fields/Created desc`);
      updates = (logRes?.value || []).map(item => ({
        workstreamTitle: item.fields?.Title || '',
        updatedBy: item.fields?.UpdatedBy || '',
        status: item.fields?.Status || '',
        percentComplete: item.fields?.PercentComplete || 0,
        comment: item.fields?.Comment || '',
        created: item.fields?.Created || null
      }));
    }
    renderDashboard();
    setLoadingState(false);
    setLastUpdated();
  } catch (e) {
    showToast(`Error loading data: ${e.message}`, 'error');
    loadDemoMode();
  }
}

async function saveUpdate(ws, name, status, pct, comment) {
  // local update
  ws.status = status;
  ws.percentComplete = pct;
  if (comment) ws.latestComment = comment;
  ws.lastUpdatedBy = name;
  ws.modified = new Date().toISOString();
  updates.unshift({ workstreamTitle: ws.title, updatedBy: name, status, percentComplete: pct, comment, created: new Date().toISOString() });

  if (!isAuthenticated || !siteId) return;
  try {
    if (ws.spId && workstreamListId) {
      await gFetch(`/sites/${siteId}/lists/${workstreamListId}/items/${ws.spId}/fields`, {
        method: 'PATCH',
        body: JSON.stringify({ Status: status, PercentComplete: pct, ...(comment ? { LatestComment: comment } : {}), LastUpdatedBy: name })
      });
    }
    if (updatesListId) {
      await gFetch(`/sites/${siteId}/lists/${updatesListId}/items`, {
        method: 'POST',
        body: JSON.stringify({ fields: { Title: ws.title, UpdatedBy: name, Status: status, PercentComplete: pct, Comment: comment } })
      });
    }
  } catch (e) { throw e; }
}

// ── Demo Mode ─────────────────────────────────────────────────────
function loadDemoMode() {
  workstreams = JSON.parse(JSON.stringify(SEED_WORKSTREAMS));
  updates     = JSON.parse(JSON.stringify(SEED_UPDATES));
  renderDashboard();
  setLastUpdated('Demo data — sign in to load live SharePoint data');
}

// ── Rendering ─────────────────────────────────────────────────────
function renderDashboard() {
  renderSummaryCards();
  renderChart();
  renderGroups(activeFilter);
  renderOverallStats();
}

function renderOverallStats() {
  const total    = workstreams.length;
  const avgPct   = total ? Math.round(workstreams.reduce((s, w) => s + w.percentComplete, 0) / total) : 0;
  const onTrack  = workstreams.filter(w => w.status === 'On Track').length;
  const atRisk   = workstreams.filter(w => w.status === 'At Risk').length;
  const behind   = workstreams.filter(w => w.status === 'Behind').length;
  const complete = workstreams.filter(w => w.status === 'Complete').length;
  document.getElementById('statTotal').textContent    = total;
  document.getElementById('statAvg').textContent      = `${avgPct}%`;
  document.getElementById('statOnTrack').textContent  = onTrack;
  document.getElementById('statAtRisk').textContent   = atRisk;
  document.getElementById('statBehind').textContent   = behind;
  document.getElementById('statComplete').textContent = complete;
}

function renderSummaryCards() {
  const grid = document.getElementById('summaryGrid');
  grid.innerHTML = CATEGORIES.map(cat => {
    const ws       = workstreams.filter(w => w.category === cat.key);
    const avg      = ws.length ? Math.round(ws.reduce((s, w) => s + w.percentComplete, 0) / ws.length) : 0;
    const onTrack  = ws.filter(w => w.status === 'On Track' || w.status === 'Complete').length;
    const atRisk   = ws.filter(w => w.status === 'At Risk').length;
    const behind   = ws.filter(w => w.status === 'Behind').length;
    const expected = CONFIG.expectedProgress;
    const delta    = avg - expected;
    const deltaStr = delta > 0 ? `+${delta}% ahead` : delta < 0 ? `${delta}% behind` : 'on pace';
    const deltaClr = delta >= 0 ? '#16A34A' : delta > -10 ? '#D97706' : '#DC2626';
    const inits    = INITIATIVE_META.filter(i => i.category === cat.key).length;
    return `
      <div class="summary-card" style="border-top:3px solid ${cat.color}">
        <div class="summary-card-header">
          <div class="cat-dot" style="background:${cat.color}"></div>
          <div class="summary-card-label">${cat.key}</div>
        </div>
        <div class="summary-stats-row">
          <div>
            <div class="summary-pct">${avg}<span class="summary-pct-unit">%</span></div>
            <div class="summary-pct-sub">${ws.length} workstreams · ${inits} initiatives</div>
          </div>
          <div class="summary-delta" style="color:${deltaClr}">${deltaStr}<br><span style="font-size:11px;font-weight:400;opacity:.8">vs expected ${expected}%</span></div>
        </div>
        <div class="summary-progress-track">
          <div class="summary-progress-fill" style="width:${avg}%;background:${cat.color}"></div>
          <div class="summary-expected-mark" style="left:${expected}%"></div>
        </div>
        <div class="summary-footer">
          <span style="color:var(--gray-500)">${inits} initiatives</span>
          <span>
            <span style="color:#16A34A;font-weight:600">${onTrack} on track</span>
            ${atRisk  ? ` <span style="color:#D97706">· ${atRisk} at risk</span>` : ''}
            ${behind  ? ` <span style="color:#DC2626">· ${behind} behind</span>` : ''}
          </span>
        </div>
      </div>`;
  }).join('');
}

function renderChart() {
  const container = document.getElementById('progressChart');
  const expected  = CONFIG.expectedProgress;

  let html = `
    <div class="chart-axis-row">
      <div class="chart-name-col"></div>
      <div class="chart-bar-col" style="position:relative;height:22px;">
        ${[0,25,50,75,100].map(v => `<div style="position:absolute;left:${v}%;transform:translateX(-50%);font-size:10px;color:var(--gray-400);font-weight:600;">${v===0?'0%':v===100?'100%':'Q'+(v/25)}</div>`).join('')}
        <div style="position:absolute;top:14px;left:0;right:0;height:1px;background:var(--gray-200);"></div>
        <div style="position:absolute;top:0;left:${expected}%;transform:translateX(-50%);font-size:10px;color:var(--pink);font-weight:700;white-space:nowrap;">▼ Today (${expected}%)</div>
      </div>
      <div class="chart-status-col"></div>
      <div class="chart-pct-col"></div>
    </div>`;

  CATEGORIES.forEach(cat => {
    const catInits = INITIATIVE_META.filter(i => i.category === cat.key);
    if (!catInits.length) return;

    html += `<div class="chart-category-header">
      <div class="chart-category-label" style="color:${cat.color}">${cat.key}</div>
      <div class="chart-category-rule" style="background:${cat.color}20"></div>
    </div>`;

    catInits.forEach(im => {
      const ws   = getInitiativeWorkstreams(im.title);
      const avg  = calcInitiativeAvg(im.title);
      const stat = calcInitiativeStatus(im.title);
      const sm   = getStatus(stat);
      const barW = avg > 0 ? Math.max(avg, 2) : 0;
      const onT  = ws.filter(w => w.status === 'On Track' || w.status === 'Complete').length;
      const risk = ws.filter(w => w.status === 'At Risk').length;
      const bhnd = ws.filter(w => w.status === 'Behind').length;

      html += `
        <div class="chart-row" onclick="scrollToGroup('${escapeAttr(im.title)}')" title="Click to jump to this initiative">
          <div class="chart-name-col">
            <div class="chart-name-text" title="${im.title}">${im.title}</div>
            <div class="chart-name-owner">${im.partner} · ${ws.length} workstreams</div>
          </div>
          <div class="chart-bar-col" style="position:relative;">
            ${[25,50,75].map(v=>`<div style="position:absolute;top:0;left:${v}%;width:1px;height:100%;background:var(--gray-200);z-index:0;pointer-events:none;"></div>`).join('')}
            <div class="chart-expected-marker" style="left:${expected}%"></div>
            <div class="chart-bar-bg">
              <div class="chart-bar-fill ${sm.barCls}" style="width:${barW}%">
                ${avg >= 12 ? `<span class="chart-bar-label">${avg}%</span>` : ''}
              </div>
            </div>
          </div>
          <div class="chart-status-col">
            <span class="chart-status-badge ${sm.badgeCls}">${stat}</span>
          </div>
          <div class="chart-pct-col">${avg}%</div>
        </div>`;
    });
  });

  container.innerHTML = html;
}

function scrollToGroup(initiativeTitle) {
  const el = document.querySelector(`[data-initiative="${CSS.escape(initiativeTitle)}"]`);
  if (el) el.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function escapeAttr(s) { return s.replace(/"/g, '&quot;'); }

function renderGroups(filter) {
  activeFilter = filter;
  const grid = document.getElementById('cardsGrid');

  const filteredInits = INITIATIVE_META.filter(im =>
    filter === 'all' || im.category === filter
  );

  if (!filteredInits.length) {
    grid.innerHTML = `<div style="text-align:center;padding:60px;color:var(--gray-400);">No initiatives found.</div>`;
    return;
  }

  grid.innerHTML = filteredInits.map(im => {
    const ws       = getInitiativeWorkstreams(im.title);
    const cat      = getCat(im.category);
    const avg      = calcInitiativeAvg(im.title);
    const stat     = calcInitiativeStatus(im.title);
    const sm       = getStatus(stat);
    const onT      = ws.filter(w => w.status === 'On Track' || w.status === 'Complete').length;
    const risk     = ws.filter(w => w.status === 'At Risk').length;
    const bhnd     = ws.filter(w => w.status === 'Behind').length;
    const notStart = ws.filter(w => w.status === 'Not Started').length;
    const expected = CONFIG.expectedProgress;

    return `
      <div class="initiative-group" data-initiative="${escapeAttr(im.title)}">
        <div class="ig-header" style="border-left:4px solid ${cat.color}">
          <div class="ig-header-top">
            <div class="ig-title">${im.title}</div>
            <span class="status-badge ${sm.badgeCls}">${stat}</span>
          </div>
          <div class="ig-meta">
            <span class="ig-meta-item"><span class="meta-label">Partner</span> ${im.partner}</span>
            <span class="ig-meta-sep">·</span>
            <span class="ig-meta-item"><span class="meta-label">Sponsor</span> ${im.execSponsor}</span>
            <span class="ig-meta-sep">·</span>
            <span class="ig-meta-item"><span class="meta-label">Category</span> <span style="color:${cat.color};font-weight:600">${im.category}</span></span>
          </div>
          <div class="ig-progress-row">
            <div class="ig-progress-track">
              <div class="ig-progress-fill ${sm.barCls}" style="width:${avg}%"></div>
              <div class="ig-expected-mark" style="left:${expected}%"></div>
            </div>
            <div class="ig-pct">${avg}%</div>
            <div class="ig-status-pills">
              ${onT   ? `<span class="ws-pill ws-pill-green">${onT} on track</span>` : ''}
              ${risk  ? `<span class="ws-pill ws-pill-amber">${risk} at risk</span>` : ''}
              ${bhnd  ? `<span class="ws-pill ws-pill-red">${bhnd} behind</span>` : ''}
              ${notStart ? `<span class="ws-pill ws-pill-gray">${notStart} not started</span>` : ''}
            </div>
          </div>
        </div>

        <div class="ws-table">
          <div class="ws-table-head">
            <div class="ws-col-name">Workstream</div>
            <div class="ws-col-owner">Owner</div>
            <div class="ws-col-status">Status</div>
            <div class="ws-col-bar">Progress</div>
            <div class="ws-col-action"></div>
          </div>
          ${ws.map(w => {
            const wsm = getStatus(w.status);
            const pct = w.percentComplete;
            const hasComment = !!w.latestComment;
            return `
              <div class="ws-row" onclick="openModal(${w.id})">
                <div class="ws-col-name">
                  <div class="ws-name">${w.title}</div>
                  ${hasComment ? `<div class="ws-comment-preview">💬 ${w.latestComment.slice(0,80)}${w.latestComment.length>80?'…':''}</div>` : ''}
                </div>
                <div class="ws-col-owner">
                  <div class="ws-owner-avatar">${w.owner.split(' ').map(n=>n[0]).join('').slice(0,2)}</div>
                  <span class="ws-owner-name">${w.owner}</span>
                </div>
                <div class="ws-col-status">
                  <span class="status-badge ${wsm.badgeCls}">${w.status}</span>
                </div>
                <div class="ws-col-bar">
                  <div class="ws-progress-track">
                    <div class="ws-progress-fill ${wsm.barCls}" style="width:${Math.max(pct>0?2:0,pct)}%"></div>
                    <div class="ws-expected-mark" style="left:${CONFIG.expectedProgress}%"></div>
                  </div>
                  <span class="ws-pct-label">${pct}%</span>
                </div>
                <div class="ws-col-action">
                  <button class="btn-ws-update" onclick="event.stopPropagation();openModal(${w.id})">Update</button>
                </div>
              </div>`;
          }).join('')}
        </div>
      </div>`;
  }).join('');
}

// ── Filter ────────────────────────────────────────────────────────
function filterCards(filter, btn) {
  document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  renderGroups(filter);
}

// ── Modal ─────────────────────────────────────────────────────────
function openModal(workstreamId) {
  selectedWorkstream = workstreams.find(w => String(w.id) === String(workstreamId));
  if (!selectedWorkstream) return;

  const w   = selectedWorkstream;
  const im  = INITIATIVE_META.find(i => i.title === w.initiative) || {};
  const cat = getCat(w.category);
  const sm  = getStatus(w.status);

  document.getElementById('mCatBadge').textContent  = w.initiative;
  document.getElementById('mCatBadge').style.color  = cat.color;
  document.getElementById('mTitle').textContent      = w.title;
  document.getElementById('mMeta').innerHTML = `
    <div class="modal-meta-item"><strong>Owner:</strong> ${w.owner}</div>
    <div class="modal-meta-item"><strong>Initiative Partner:</strong> ${im.partner || 'TBD'}</div>
    <div class="modal-meta-item"><strong>Exec Sponsor:</strong> ${w.execSponsor || im.execSponsor || 'TBD'}</div>
  `;

  document.getElementById('mDesc').textContent = w.latestComment || 'No updates submitted yet.';

  const circle = document.getElementById('mStatusCircle');
  circle.textContent = sm.icon;
  circle.style.background = sm.bg;
  circle.style.color = sm.fg;
  document.getElementById('mStatusVal').textContent = w.status;
  document.getElementById('mPctVal').textContent    = `${w.percentComplete}%`;

  const mBar = document.getElementById('mProgressFill');
  mBar.className   = `card-progress-fill ${sm.barCls}`;
  mBar.style.width = `${w.percentComplete}%`;
  document.getElementById('mExpected').style.left = `${CONFIG.expectedProgress}%`;

  document.getElementById('mTimeline').innerHTML = [
    { q:'Q1', text: w.q1Goals },
    { q:'Q2', text: w.q2Goals },
    { q:'Q3', text: w.q3Goals }
  ].map(({ q, text }) => `
    <div class="timeline-item">
      <div class="timeline-q" style="color:${cat.color}">${q}</div>
      <div class="timeline-text">${text || 'Not yet defined'}</div>
    </div>`).join('');

  document.getElementById('fStatus').value         = w.status;
  document.getElementById('fPct').value            = w.percentComplete;
  document.getElementById('fPctVal').textContent   = w.percentComplete;
  document.getElementById('fComment').value        = '';
  document.getElementById('fName').value           = '';
  document.getElementById('submitBtn').disabled    = false;
  document.getElementById('submitBtn').textContent = 'Submit Update';

  renderHistory(w.title);

  document.getElementById('mOverlay').classList.add('active');
  document.getElementById('mPanel').classList.add('active');
  document.body.style.overflow = 'hidden';
}

function closeModal() {
  document.getElementById('mOverlay').classList.remove('active');
  document.getElementById('mPanel').classList.remove('active');
  document.body.style.overflow = '';
  selectedWorkstream = null;
}

function renderHistory(title) {
  const tbody    = document.getElementById('historyBody');
  const relevant = updates.filter(u => u.workstreamTitle === title);
  if (!relevant.length) {
    tbody.innerHTML = `<tr><td colspan="5" class="history-empty">No updates submitted yet.</td></tr>`;
    return;
  }
  tbody.innerHTML = relevant.map(u => {
    const sm = getStatus(u.status);
    return `<tr>
      <td style="white-space:nowrap;color:var(--gray-400);font-size:11px">${fmtDate(u.created)}</td>
      <td style="font-weight:500;color:var(--navy)">${u.updatedBy}</td>
      <td><span class="status-badge ${sm.badgeCls}" style="font-size:10px">${u.status}</span></td>
      <td style="font-weight:700;color:var(--navy);text-align:center">${u.percentComplete}%</td>
      <td style="font-size:12px;color:var(--gray-600);max-width:180px;word-break:break-word">${u.comment || '—'}</td>
    </tr>`;
  }).join('');
}

async function submitUpdate() {
  const name    = document.getElementById('fName').value.trim();
  const status  = document.getElementById('fStatus').value;
  const pct     = parseInt(document.getElementById('fPct').value);
  const comment = document.getElementById('fComment').value.trim();

  if (!name)   { showToast('Please enter your name.', 'error');   return; }
  if (!status) { showToast('Please select a status.', 'error'); return; }

  const btn = document.getElementById('submitBtn');
  btn.disabled  = true;
  btn.innerHTML = '<span class="spinner"></span> Submitting…';

  try {
    await saveUpdate(selectedWorkstream, name, status, pct, comment);

    const sm = getStatus(status);
    document.getElementById('mStatusCircle').textContent      = sm.icon;
    document.getElementById('mStatusCircle').style.background = sm.bg;
    document.getElementById('mStatusCircle').style.color      = sm.fg;
    document.getElementById('mStatusVal').textContent         = status;
    document.getElementById('mPctVal').textContent            = `${pct}%`;
    const mBar = document.getElementById('mProgressFill');
    mBar.className   = `card-progress-fill ${sm.barCls}`;
    mBar.style.width = `${pct}%`;

    renderHistory(selectedWorkstream.title);
    renderDashboard();
    showToast('✓ Update submitted!', 'success');
    document.getElementById('fComment').value = '';
    document.getElementById('fName').value    = '';
  } catch (e) {
    showToast(`Save failed: ${e.message}`, 'error');
  } finally {
    btn.disabled     = false;
    btn.textContent  = 'Submit Update';
  }
}

// ── Utility ───────────────────────────────────────────────────────
function fmtDate(str) {
  if (!str) return '—';
  try { return new Date(str).toLocaleDateString('en-US', { month:'short', day:'numeric', year:'numeric' }); }
  catch { return str; }
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
  t.textContent = msg;
  t.className   = `toast show ${type}`;
  setTimeout(() => { t.className = 'toast'; }, 4500);
}

function updateAuthUI() {
  const section = document.getElementById('authSection');
  if (!isAuthenticated || !currentAccount) {
    section.innerHTML = `<button class="btn-auth" onclick="handleAuth()">Sign In with Microsoft</button>`;
    return;
  }
  const name     = currentAccount.name || currentAccount.username || 'User';
  const initials = name.split(' ').map(n => n[0] || '').join('').toUpperCase().slice(0,2) || 'U';
  section.innerHTML = `
    <div class="user-badge">
      <div class="user-avatar">${initials}</div>
      <span class="user-name">${name.split(' ')[0]}</span>
    </div>
    <button class="btn-auth btn-auth-ghost" onclick="handleAuth()">Sign Out</button>`;
}

// ── Bootstrap ─────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  document.getElementById('fPct').addEventListener('input', function () {
    document.getElementById('fPctVal').textContent = this.value;
  });
  initMSAL();
  setInterval(() => { if (isAuthenticated && siteId) loadLiveData(); }, CONFIG.refreshIntervalMs);
});
