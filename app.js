// ═══════════════════════════════════════════════════════
//  app.js – 2026 Firm Building Initiatives Dashboard
// ═══════════════════════════════════════════════════════

// ── Global state ─────────────────────────────────────────────────
let msalInstance   = null;
let currentAccount = null;
let graphToken     = null;
let siteId         = null;
let wsListId       = null;
let updatesListId  = null;
let workstreams    = [];
let updates        = [];
let selectedWs     = null;
let isAuthenticated = false;
let activeFilter   = 'all';

// ── Category metadata ─────────────────────────────────────────────
const CATEGORIES = [
  { key:'Revenue & Operations', color:'#E60F65', lightBg:'#FFF0F5' },
  { key:'Product & Tech',       color:'#7C3AED', lightBg:'#F5F3FF' },
  { key:'Systems & People',     color:'#0EA5E9', lightBg:'#F0F9FF' }
];

const STATUS_META = {
  'On Track':    { barCls:'bar-on-track',    badgeCls:'badge-on-track',    icon:'\u2713', bg:'#DCFCE7', fg:'#16A34A' },
  'At Risk':     { barCls:'bar-at-risk',     badgeCls:'badge-at-risk',     icon:'!',       bg:'#FEF3C7', fg:'#D97706' },
  'Behind':      { barCls:'bar-behind',      badgeCls:'badge-behind',      icon:'\u2715', bg:'#FEE2E2', fg:'#DC2626' },
  'Not Started': { barCls:'bar-not-started', badgeCls:'badge-not-started', icon:'\u25CB', bg:'#F1F5F9', fg:'#64748B' },
  'Complete':    { barCls:'bar-complete',    badgeCls:'badge-complete',    icon:'\u2605', bg:'#DBEAFE', fg:'#2563EB' }
};

const STATUS_PRIORITY = { 'Behind':0, 'At Risk':1, 'Not Started':2, 'On Track':3, 'Complete':4 };

function getCat(key)    { return CATEGORIES.find(c => c.key === key) || CATEGORIES[0]; }
function getSM(key)     { return STATUS_META[key] || STATUS_META['Not Started']; }

// ── Initiative rollup helpers ─────────────────────────────────────
function getInitWs(id)    { return workstreams.filter(w => w.initiativeId === id); }
function calcInitPct(id)  {
  const ws = getInitWs(id);
  return ws.length ? Math.round(ws.reduce((s,w)=>s+w.percentComplete,0)/ws.length) : 0;
}
function calcInitStatus(id) {
  const ws = getInitWs(id);
  if (!ws.length) return 'Not Started';
  return ws.reduce((worst,w) =>
    STATUS_PRIORITY[w.status] < STATUS_PRIORITY[worst] ? w.status : worst, 'Complete');
}

// ── MSAL Auth ─────────────────────────────────────────────────────
function initMSAL() {
  if (!CONFIG.clientId || CONFIG.clientId === 'YOUR_AZURE_AD_CLIENT_ID') {
    loadDemoMode(); return;
  }
  const cfg = {
    auth: { clientId:CONFIG.clientId, authority:`https://login.microsoftonline.com/${CONFIG.tenantId}`, redirectUri:CONFIG.redirectUri },
    cache: { cacheLocation:'sessionStorage', storeAuthStateInCookie:false }
  };
  try {
    msalInstance = new msal.PublicClientApplication(cfg);
    msalInstance.handleRedirectPromise().then(onAuthResponse).catch(() => loadDemoMode());
  } catch { loadDemoMode(); }
}

async function onAuthResponse(resp) {
  if (resp) { currentAccount = resp.account; await completeSignIn(); }
  else {
    const accs = msalInstance?.getAllAccounts() || [];
    if (accs.length) { currentAccount = accs[0]; await completeSignIn(); }
    else loadDemoMode();
  }
}

async function handleAuth() {
  if (isAuthenticated) { signOut(); return; }
  if (!CONFIG.clientId || CONFIG.clientId === 'YOUR_AZURE_AD_CLIENT_ID') {
    showToast('Auth not configured.', 'error'); return;
  }
  try {
    const r = await msalInstance.loginPopup({ scopes:['User.Read','Sites.ReadWrite.All'] });
    currentAccount = r.account; await completeSignIn();
  } catch { showToast('Sign-in failed. Please try again.', 'error'); }
}

function signOut() {
  msalInstance?.logoutPopup({ account:currentAccount });
  currentAccount = null; graphToken = null; isAuthenticated = false;
  loadDemoMode();
}

async function getToken() {
  if (!msalInstance || !currentAccount) return null;
  try {
    const r = await msalInstance.acquireTokenSilent({ scopes:['Sites.ReadWrite.All'], account:currentAccount });
    return r.accessToken;
  } catch {
    try {
      const r = await msalInstance.acquireTokenPopup({ scopes:['Sites.ReadWrite.All'], account:currentAccount });
      return r.accessToken;
    } catch { return null; }
  }
}

async function completeSignIn() {
  isAuthenticated = true; updateAuthUI();
  document.getElementById('authBanner').style.display = 'none';
  graphToken = await getToken();
  if (!graphToken) { loadDemoMode(); showToast('Could not get token — showing demo data.', 'error'); return; }
  await resolveSharePointIds();
  await loadLiveData();
}

// ── Microsoft Graph ───────────────────────────────────────────────
async function gFetch(path, opts={}) {
  const token = graphToken || await getToken();
  if (!token) throw new Error('Not authenticated');
  const res = await fetch(CONFIG.graphEndpoint + path, {
    ...opts,
    headers: { Authorization:`Bearer ${token}`, 'Content-Type':'application/json', ...(opts.headers||{}) }
  });
  if (!res.ok) { const e = await res.json().catch(()=>({})); throw new Error(e?.error?.message || `HTTP ${res.status}`); }
  return res.status === 204 ? null : res.json();
}

async function resolveSharePointIds() {
  try {
    const site  = await gFetch(`/sites/${CONFIG.siteHostname}:${CONFIG.sitePath}`);
    siteId = site.id;
    const lists = (await gFetch(`/sites/${siteId}/lists?$select=id,displayName`))?.value || [];
    const find  = name => lists.find(l => l.displayName?.toLowerCase() === name.toLowerCase())?.id || null;
    wsListId      = find(CONFIG.wsListName);
    updatesListId = find(CONFIG.updatesListName);
    if (!wsListId) console.warn('Workstream list not found. Available:', lists.map(l=>l.displayName).join(', '));
  } catch(e) { showToast(`SharePoint error: ${e.message}`, 'error'); }
}

async function loadLiveData() {
  if (!siteId || !wsListId) { loadDemoMode(); return; }
  try {
    setLoadingState(true);
    const fields = 'Title,Initiative,InitiativeId,Category,Owner,Status,PercentComplete,Description,LatestComment,LastUpdatedBy,Modified';
    const res    = await gFetch(`/sites/${siteId}/lists/${wsListId}/items?$expand=fields($select=${fields})&$top=100`);
    const items  = res?.value || [];
    if (items.length === 0) {
      workstreams = JSON.parse(JSON.stringify(SEED_WORKSTREAMS));
      showToast('List is empty — showing demo data. Add workstream rows in SharePoint to go live.', 'error');
    } else {
      workstreams = items.map((item,idx) => ({
        id:              item.id || 'sp-'+idx,
        spId:            item.id,
        initiativeId:    parseInt(item.fields?.InitiativeId) || 0,
        initiativeTitle: item.fields?.Initiative     || '',
        category:        item.fields?.Category        || '',
        title:           item.fields?.Title           || '',
        owner:           item.fields?.Owner           || '',
        status:          item.fields?.Status          || 'Not Started',
        percentComplete: parseInt(item.fields?.PercentComplete) || 0,
        description:     item.fields?.Description     || '',
        latestComment:   item.fields?.LatestComment   || '',
        lastUpdatedBy:   item.fields?.LastUpdatedBy   || '',
        modified:        item.fields?.Modified        || null
      }));
    }
    if (updatesListId) {
      const logRes = await gFetch(`/sites/${siteId}/lists/${updatesListId}/items?$expand=fields($select=Title,UpdatedBy,Status,PercentComplete,Comment,Created)&$top=200&$orderby=fields/Created desc`);
      updates = (logRes?.value || []).map(i => ({
        wsTitle:         i.fields?.Title           || '',
        updatedBy:       i.fields?.UpdatedBy       || '',
        status:          i.fields?.Status          || '',
        percentComplete: i.fields?.PercentComplete || 0,
        comment:         i.fields?.Comment         || '',
        created:         i.fields?.Created         || null
      }));
    }
    renderDashboard(); setLoadingState(false); setLastUpdated();
  } catch(e) {
    showToast(`Load error: ${e.message}`, 'error');
    loadDemoMode();
  }
}

async function saveUpdate(ws, name, status, pct, comment) {
  ws.status = status; ws.percentComplete = pct;
  if (comment) ws.latestComment = comment;
  ws.lastUpdatedBy = name; ws.modified = new Date().toISOString();
  updates.unshift({ wsTitle:ws.title, updatedBy:name, status, percentComplete:pct, comment, created:new Date().toISOString() });
  if (!isAuthenticated || !siteId) return;
  if (ws.spId && wsListId) {
    await gFetch(`/sites/${siteId}/lists/${wsListId}/items/${ws.spId}/fields`, {
      method:'PATCH',
      body:JSON.stringify({ Status:status, PercentComplete:pct, ...(comment?{LatestComment:comment}:{}), LastUpdatedBy:name })
    });
  }
  if (updatesListId) {
    await gFetch(`/sites/${siteId}/lists/${updatesListId}/items`, {
      method:'POST',
      body:JSON.stringify({ fields:{ Title:ws.title, UpdatedBy:name, Status:status, PercentComplete:pct, Comment:comment } })
    });
  }
}

// ── Demo mode ─────────────────────────────────────────────────────
function loadDemoMode() {
  workstreams = JSON.parse(JSON.stringify(SEED_WORKSTREAMS));
  updates     = [];
  renderDashboard();
  setLastUpdated('Demo data — sign in with Microsoft to load live data');
}

// ── Rendering ─────────────────────────────────────────────────────
function renderDashboard() {
  renderOverallStats();
  renderSummaryCards();
  renderChart();
  renderCards(activeFilter);
}

function renderOverallStats() {
  const total    = workstreams.length;
  const avgPct   = total ? Math.round(workstreams.reduce((s,w)=>s+w.percentComplete,0)/total) : 0;
  const onTrack  = workstreams.filter(w=>w.status==='On Track').length;
  const atRisk   = workstreams.filter(w=>w.status==='At Risk').length;
  const behind   = workstreams.filter(w=>w.status==='Behind').length;
  const complete = workstreams.filter(w=>w.status==='Complete').length;
  document.getElementById('statTotal').textContent    = total;
  document.getElementById('statAvg').textContent      = avgPct + '%';
  document.getElementById('statOnTrack').textContent  = onTrack;
  document.getElementById('statAtRisk').textContent   = atRisk;
  document.getElementById('statBehind').textContent   = behind;
  document.getElementById('statComplete').textContent = complete;
}

function renderSummaryCards() {
  const expected = CONFIG.expectedProgress;
  document.getElementById('summaryGrid').innerHTML = CATEGORIES.map(cat => {
    const items   = workstreams.filter(w=>w.category===cat.key);
    const avg     = items.length ? Math.round(items.reduce((s,w)=>s+w.percentComplete,0)/items.length) : 0;
    const onTrack = items.filter(w=>['On Track','Complete'].includes(w.status)).length;
    const atRisk  = items.filter(w=>w.status==='At Risk').length;
    const behind  = items.filter(w=>w.status==='Behind').length;
    const delta   = avg - expected;
    const dStr    = delta>0 ? '+'+delta+'% ahead' : delta<0 ? delta+'% behind' : 'on pace';
    const dClr    = delta>=0 ? '#16A34A' : delta>-10 ? '#D97706' : '#DC2626';
    const nInits  = SEED_INITIATIVES.filter(i=>i.category===cat.key).length;
    return `<div class="summary-card" style="border-top:3px solid ${cat.color}">
      <div class="summary-card-header"><div class="cat-dot" style="background:${cat.color}"></div><div class="summary-card-label">${cat.key}</div></div>
      <div class="summary-stats-row">
        <div><div class="summary-pct">${avg}<span class="summary-pct-unit">%</span></div><div class="summary-pct-sub">${items.length} workstreams &middot; ${nInits} initiatives</div></div>
        <div class="summary-delta" style="color:${dClr}">${dStr}<br><span style="font-size:11px;font-weight:400;opacity:.8">vs expected ${expected}%</span></div>
      </div>
      <div class="summary-progress-track">
        <div class="summary-progress-fill" style="width:${avg}%;background:${cat.color}"></div>
        <div class="summary-expected-mark" style="left:${expected}%"></div>
      </div>
      <div class="summary-footer">
        <span style="color:var(--gray-500)">${nInits} initiatives</span>
        <span>
          <span style="color:#16A34A;font-weight:600">${onTrack} on track</span>
          ${atRisk ? '<span style="color:#D97706"> &middot; '+atRisk+' at risk</span>' : ''}
          ${behind ? '<span style="color:#DC2626"> &middot; '+behind+' behind</span>' : ''}
        </span>
      </div>
    </div>`;
  }).join('');
}

function renderChart() {
  const expected = CONFIG.expectedProgress;
  let html = `<div class="chart-axis-row">
    <div class="chart-name-col"></div>
    <div class="chart-bar-col" style="position:relative;height:22px;">
      ${[0,25,50,75,100].map(v=>`<div style="position:absolute;left:${v}%;transform:translateX(-50%);font-size:10px;color:var(--gray-400);font-weight:600;">${v===0?'0%':v===100?'100%':'Q'+(v/25)}</div>`).join('')}
      <div style="position:absolute;top:14px;left:0;right:0;height:1px;background:var(--gray-200);"></div>
      <div style="position:absolute;top:0;left:${expected}%;transform:translateX(-50%);font-size:10px;color:var(--pink);font-weight:700;white-space:nowrap;">&#9660; Today (${expected}%)</div>
    </div>
    <div class="chart-status-col"></div>
    <div class="chart-pct-col"></div>
  </div>`;

  CATEGORIES.forEach(cat => {
    const catInits = SEED_INITIATIVES.filter(i=>i.category===cat.key);
    if (!catInits.length) return;
    html += `<div class="chart-category-header">
      <div class="chart-category-label" style="color:${cat.color}">${cat.key}</div>
      <div class="chart-category-rule" style="background:${cat.color}20"></div>
    </div>`;
    catInits.forEach(init => {
      const pct    = calcInitPct(init.id);
      const status = calcInitStatus(init.id);
      const sm     = getSM(status);
      const barW   = pct>0 ? Math.max(pct,2) : 0;
      const wsCount = getInitWs(init.id).length;
      html += `<div class="chart-row" onclick="scrollToInit(${init.id})" title="Jump to workstreams">
        <div class="chart-name-col">
          <div class="chart-name-text" title="${init.title}">${init.title}</div>
          <div class="chart-name-owner">${init.accountablePartner} &middot; ${wsCount} workstreams &middot; avg ${pct}%</div>
        </div>
        <div class="chart-bar-col" style="position:relative;">
          ${[25,50,75].map(v=>`<div style="position:absolute;top:0;left:${v}%;width:1px;height:100%;background:var(--gray-200);z-index:0;"></div>`).join('')}
          <div class="chart-expected-marker" style="left:${expected}%"></div>
          <div class="chart-bar-bg">
            <div class="chart-bar-fill ${sm.barCls}" style="width:${barW}%">${pct>=12?'<span class="chart-bar-label">'+pct+'%</span>':''}</div>
          </div>
        </div>
        <div class="chart-status-col"><span class="chart-status-badge ${sm.badgeCls}">${status}</span></div>
        <div class="chart-pct-col">${pct}%</div>
      </div>`;
    });
  });
  document.getElementById('progressChart').innerHTML = html;
}

function scrollToInit(initId) {
  const el = document.getElementById('init-'+initId);
  if (el) el.scrollIntoView({ behavior:'smooth', block:'start' });
}

function renderCards(filter='all') {
  activeFilter = filter;
  const filteredInits = filter==='all' ? SEED_INITIATIVES : SEED_INITIATIVES.filter(i=>i.category===filter);
  const grid = document.getElementById('cardsGrid');
  if (!filteredInits.length) { grid.innerHTML='<div style="text-align:center;padding:60px;color:var(--gray-400)">No initiatives found.</div>'; return; }

  grid.innerHTML = filteredInits.map(init => {
    const cat    = getCat(init.category);
    const pct    = calcInitPct(init.id);
    const status = calcInitStatus(init.id);
    const sm     = getSM(status);
    const ws     = getInitWs(init.id);
    const onTrack = ws.filter(w=>['On Track','Complete'].includes(w.status)).length;
    const atRisk  = ws.filter(w=>w.status==='At Risk').length;
    const behind  = ws.filter(w=>w.status==='Behind').length;

    const wsRows = ws.map(w => {
      const wsm  = getSM(w.status);
      const wsPct = Math.max(0, Math.min(100, w.percentComplete));
      const initials = w.owner.split(' ').map(n=>n[0]||'').join('').toUpperCase().slice(0,2);
      return `<div class="ws-row" onclick="openModal('${w.id}')">
        <div class="ws-col-name">
          <div class="ws-name">${w.title}</div>
          ${w.latestComment ? '<div class="ws-comment-preview">'+w.latestComment.substring(0,70)+(w.latestComment.length>70?'&hellip;':'')+'</div>' : ''}
        </div>
        <div class="ws-col-owner">
          <div class="ws-owner-avatar">${initials}</div>
          <div class="ws-owner-name">${w.owner}</div>
        </div>
        <div class="ws-col-bar">
          <div class="ws-progress-track">
            <div class="ws-expected-mark" style="left:${CONFIG.expectedProgress}%"></div>
            <div class="ws-progress-fill ${wsm.barCls}" style="width:${wsPct>0?Math.max(wsPct,3):0}%"></div>
          </div>
          <span class="ws-pct-label">${wsPct}%</span>
        </div>
        <div><span class="status-badge ${wsm.badgeCls}" style="font-size:10px">${w.status}</span></div>
        <div class="ws-col-action"><button class="btn-ws-update" onclick="event.stopPropagation();openModal('${w.id}')">Update</button></div>
      </div>`;
    }).join('');

    return `<div class="initiative-group" id="init-${init.id}">
      <div class="ig-header">
        <div class="ig-header-top">
          <div>
            <div style="margin-bottom:6px"><span class="status-badge" style="background:${cat.lightBg};color:${cat.color};font-size:10px">${init.category}</span></div>
            <div class="ig-title">${init.title}</div>
          </div>
          <span class="status-badge ${sm.badgeCls}" style="flex-shrink:0;margin-top:4px">${status}</span>
        </div>
        <div class="ig-meta">
          <div class="ig-meta-item"><span class="meta-label">Partner</span>&nbsp;<strong>${init.accountablePartner}</strong></div>
          <span class="ig-meta-sep">&middot;</span>
          <div class="ig-meta-item"><span class="meta-label">Sponsor</span>&nbsp;<strong>${init.execSponsor}</strong></div>
          <span class="ig-meta-sep">&middot;</span>
          <div class="ig-meta-item" style="color:var(--gray-400)">${ws.length} workstreams</div>
        </div>
        <div class="ig-progress-row">
          <div class="ig-progress-track">
            <div class="ig-expected-mark" style="left:${CONFIG.expectedProgress}%"></div>
            <div class="ig-progress-fill ${sm.barCls}" style="width:${pct}%"></div>
          </div>
          <div class="ig-pct">${pct}%</div>
          <div class="ig-status-pills">
            ${onTrack ? '<span class="ws-pill ws-pill-green">'+onTrack+' on track</span>' : ''}
            ${atRisk  ? '<span class="ws-pill ws-pill-amber">'+atRisk+' at risk</span>'   : ''}
            ${behind  ? '<span class="ws-pill ws-pill-red">'+behind+' behind</span>'       : ''}
          </div>
        </div>
      </div>
      <div class="ws-table">
        <div class="ws-table-head">
          <div>Workstream</div><div>Owner</div><div>Progress</div><div>Status</div><div></div>
        </div>
        ${wsRows}
      </div>
    </div>`;
  }).join('');
}

// ── Modal ─────────────────────────────────────────────────────────
function openModal(wsId) {
  selectedWs = workstreams.find(w => String(w.id) === String(wsId));
  if (!selectedWs) return;
  const w  = selectedWs;
  const sm = getSM(w.status);
  const init = SEED_INITIATIVES.find(i=>i.id===w.initiativeId) || {};
  const cat  = getCat(w.category);

  document.getElementById('mCatBadge').textContent  = w.initiativeTitle || init.title || '';
  document.getElementById('mCatBadge').style.color  = cat.color;
  document.getElementById('mTitle').textContent      = w.title;
  document.getElementById('mMeta').innerHTML = `
    <div class="modal-meta-item"><strong>Owner:</strong> ${w.owner||'TBD'}</div>
    <div class="modal-meta-item"><strong>Category:</strong> ${w.category}</div>
    <div class="modal-meta-item"><strong>Initiative Partner:</strong> ${init.accountablePartner||'TBD'}</div>`;

  document.getElementById('mDesc').textContent = w.latestComment || 'No updates submitted yet.';

  const circle = document.getElementById('mStatusCircle');
  circle.textContent = sm.icon; circle.style.background = sm.bg; circle.style.color = sm.fg;
  document.getElementById('mStatusVal').textContent = w.status;
  document.getElementById('mPctVal').textContent    = w.percentComplete + '%';

  const fill = document.getElementById('mProgressFill');
  fill.className   = 'ig-progress-fill ' + sm.barCls;
  fill.style.width = w.percentComplete + '%';
  document.getElementById('mExpected').style.left = CONFIG.expectedProgress + '%';

  // Quarterly milestones from initiative
  const tl = document.getElementById('mTimeline');
  if (tl) {
    tl.innerHTML = [
      { q:'Q1', text:init.q1Goals }, { q:'Q2', text:init.q2Goals }, { q:'Q3', text:init.q3Goals }
    ].map(({q,text})=>`<div class="timeline-item"><div class="timeline-q" style="color:${cat.color}">${q}</div><div class="timeline-text">${text||'Not yet defined'}</div></div>`).join('');
  }

  document.getElementById('fStatus').value       = w.status;
  document.getElementById('fPct').value          = w.percentComplete;
  document.getElementById('fPctVal').textContent = w.percentComplete;
  document.getElementById('fComment').value      = '';
  document.getElementById('fName').value         = '';
  const btn = document.getElementById('submitBtn');
  btn.disabled = false; btn.textContent = 'Submit Update';

  renderHistory(w.title);
  document.getElementById('mOverlay').classList.add('active');
  document.getElementById('mPanel').classList.add('active');
  document.body.style.overflow = 'hidden';
}

function closeModal() {
  document.getElementById('mOverlay').classList.remove('active');
  document.getElementById('mPanel').classList.remove('active');
  document.body.style.overflow = '';
  selectedWs = null;
}

function renderHistory(wsTitle) {
  const tbody    = document.getElementById('historyBody');
  const relevant = updates.filter(u=>u.wsTitle===wsTitle);
  if (!relevant.length) {
    tbody.innerHTML = '<tr><td colspan="5" class="history-empty">No updates submitted yet.</td></tr>';
    return;
  }
  tbody.innerHTML = relevant.map(u => {
    const sm = getSM(u.status);
    return `<tr>
      <td style="white-space:nowrap;color:var(--gray-400);font-size:11px">${fmtDate(u.created)}</td>
      <td style="font-weight:500;color:var(--navy)">${u.updatedBy}</td>
      <td><span class="status-badge ${sm.badgeCls}" style="font-size:10px">${u.status}</span></td>
      <td style="font-weight:700;color:var(--navy);text-align:center">${u.percentComplete}%</td>
      <td style="font-size:12px;color:var(--gray-600);max-width:180px;word-break:break-word">${u.comment||'&mdash;'}</td>
    </tr>`;
  }).join('');
}

async function submitUpdate() {
  const name    = document.getElementById('fName').value.trim();
  const status  = document.getElementById('fStatus').value;
  const pct     = parseInt(document.getElementById('fPct').value);
  const comment = document.getElementById('fComment').value.trim();
  if (!name)   { showToast('Please enter your name.', 'error'); return; }
  if (!status) { showToast('Please select a status.', 'error'); return; }

  const btn = document.getElementById('submitBtn');
  btn.disabled = true; btn.innerHTML = '<span class="spinner"></span> Submitting&hellip;';
  try {
    await saveUpdate(selectedWs, name, status, pct, comment);
    const sm = getSM(status);
    document.getElementById('mStatusCircle').textContent = sm.icon;
    document.getElementById('mStatusCircle').style.background = sm.bg;
    document.getElementById('mStatusCircle').style.color = sm.fg;
    document.getElementById('mStatusVal').textContent = status;
    document.getElementById('mPctVal').textContent    = pct + '%';
    const fill = document.getElementById('mProgressFill');
    fill.className = 'ig-progress-fill ' + sm.barCls;
    fill.style.width = pct + '%';
    document.getElementById('mDesc').textContent = comment || selectedWs.latestComment || '';
    renderHistory(selectedWs.title);
    renderDashboard();
    showToast('\u2713 Update submitted!', 'success');
    document.getElementById('fComment').value = '';
    document.getElementById('fName').value    = '';
  } catch(e) {
    showToast('Save failed: ' + e.message, 'error');
  } finally {
    btn.disabled = false; btn.textContent = 'Submit Update';
  }
}

// ── Filters ───────────────────────────────────────────────────────
function filterCards(filter, btn) {
  document.querySelectorAll('.filter-btn').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');
  renderCards(filter);
}

// ── Utilities ─────────────────────────────────────────────────────
function fmtDate(str) {
  if (!str) return '&mdash;';
  try { return new Date(str).toLocaleDateString('en-US',{month:'short',day:'numeric',year:'numeric'}); }
  catch { return str; }
}
function setLastUpdated(text) {
  document.getElementById('lastUpdated').textContent = text ||
    'Updated ' + new Date().toLocaleString('en-US',{month:'short',day:'numeric',hour:'2-digit',minute:'2-digit'});
}
function setLoadingState(on) {
  document.getElementById('progressChart').style.opacity = on ? '0.4' : '1';
}
function showToast(msg, type='success') {
  const t = document.getElementById('toast');
  t.textContent = msg; t.className = 'toast show ' + type;
  setTimeout(() => t.className = 'toast', 4500);
}
function updateAuthUI() {
  const s = document.getElementById('authSection');
  if (!isAuthenticated || !currentAccount) {
    s.innerHTML = '<button class="btn-auth" onclick="handleAuth()">Sign In with Microsoft</button>';
    return;
  }
  const name = currentAccount.name || currentAccount.username || 'User';
  const initials = name.split(' ').map(n=>n[0]||'').join('').toUpperCase().slice(0,2) || 'U';
  s.innerHTML = `<div class="user-badge"><div class="user-avatar">${initials}</div><span class="user-name">${name.split(' ')[0]}</span></div><button class="btn-auth btn-auth-ghost" onclick="handleAuth()">Sign Out</button>`;
}

// ── Bootstrap ─────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  document.getElementById('fPct').addEventListener('input', function() {
    document.getElementById('fPctVal').textContent = this.value;
  });
  initMSAL();
  setInterval(() => { if (isAuthenticated && siteId) loadLiveData(); }, CONFIG.refreshIntervalMs);
});
