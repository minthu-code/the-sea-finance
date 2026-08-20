// ============================================================================
// THE SEA FINANCE — Application Logic
// ============================================================================

var $ = function(id){ return document.getElementById(id); };
var state = {
  exhibitions: [],
  currentExhibition: null,
  dashboard: null,
  artworks: [],
  sales: [],
  receivables: [],
  pendingExpenses: [],
  confirmedExpenses: [],
  budgets: [],
  artistRoi: [],
  payables: [],
  accountHeads: [],
  cashChart: null,
  donutChart: null,
  seaHistory: [],
  passwordRequired: false,
};

function fmt(n){
  n = Number(n || 0);
  var sign = n < 0 ? '-' : '';
  return sign + '\u0e3f' + Math.abs(n).toLocaleString('en-US', {minimumFractionDigits: 0, maximumFractionDigits: 0});
}
function fmt2(n){
  n = Number(n || 0);
  var sign = n < 0 ? '-' : '';
  return sign + '\u0e3f' + Math.abs(n).toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2});
}
function esc(s){
  var d = document.createElement('div');
  d.textContent = (s === null || s === undefined) ? '' : String(s);
  return d.innerHTML;
}
function getVar(v){ return getComputedStyle(document.documentElement).getPropertyValue(v).trim(); }

// ── FETCH HELPER ───────────────────────────────────────────────────────────
async function api(path, opts){
  opts = opts || {};
  opts.credentials = 'include';
  if (opts.body && typeof opts.body !== 'string') {
    opts.headers = Object.assign({'Content-Type': 'application/json'}, opts.headers || {});
    opts.body = JSON.stringify(opts.body);
  }
  var res = await fetch(path, opts);
  if (res.status === 401) {
    showLogin();
    throw new Error('Unauthorized');
  }
  var data = null;
  try { data = await res.json(); } catch(e) { data = null; }
  if (!res.ok) {
    throw new Error((data && data.detail) ? data.detail : 'Request failed');
  }
  return data;
}

// ── LOGIN / SESSION ─────────────────────────────────────────────────────────
function showLogin(){
  $('appMain').style.display = 'none';
  $('bottomNav').style.display = 'none';
  $('fabWrap').style.display = 'none';
  $('loginOverlay').classList.add('open');
}
function hideLogin(){
  $('loginOverlay').classList.remove('open');
  $('appMain').style.display = 'flex';
  $('bottomNav').style.display = 'flex';
  $('fabWrap').style.display = 'flex';
}

async function checkSession(){
  try {
    var res = await fetch('/api/session', {credentials: 'include'});
    var data = await res.json();
    state.passwordRequired = !!data.password_required;
    $('signOutRow').style.display = state.passwordRequired ? 'flex' : 'none';
    $('signOutDivider').style.display = state.passwordRequired ? 'block' : 'none';
    if (data.password_required && !data.authenticated) {
      showLogin();
      return false;
    }
    hideLogin();
    return true;
  } catch(e) {
    hideLogin();
    return true;
  }
}

$('loginSubmit').addEventListener('click', doLogin);
$('loginPassword').addEventListener('keydown', function(e){ if (e.key === 'Enter') doLogin(); });
async function doLogin(){
  var pw = $('loginPassword').value;
  var fd = new URLSearchParams();
  fd.append('password', pw);
  try {
    var res = await fetch('/api/login', {method: 'POST', body: fd, credentials: 'include'});
    if (!res.ok) throw new Error('bad');
    $('loginError').classList.remove('show');
    hideLogin();
    boot();
  } catch(e) {
    $('loginError').classList.add('show');
  }
}
$('signOutRow').addEventListener('click', function(e){
  e.preventDefault();
  document.cookie = 'auth_token=; Max-Age=0; path=/';
  location.reload();
});

// ── THEME ────────────────────────────────────────────────────────────────
function toggleTheme(){
  var theme = document.documentElement.getAttribute('data-theme');
  var next = theme === 'dark' ? 'light' : 'dark';
  document.documentElement.setAttribute('data-theme', next);
  localStorage.setItem('sea-theme', next);
  updateKnobIcon(next);
  if (state.cashChart) { state.cashChart.destroy(); state.cashChart = null; }
  if (state.donutChart) { state.donutChart.destroy(); state.donutChart = null; }
  renderHome();
}
function updateKnobIcon(theme){
  var knob = $('knobIcon');
  knob.innerHTML = theme === 'dark'
    ? '<svg viewBox="0 0 24 24" width="11" height="11" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"><path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"/></svg>'
    : '<svg viewBox="0 0 24 24" width="11" height="11" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"><circle cx="12" cy="12" r="4"/></svg>';
}
updateKnobIcon(document.documentElement.getAttribute('data-theme'));
$('themeMenuRow').addEventListener('click', function(e){ e.stopPropagation(); toggleTheme(); });

// ── AVATAR MENU ──────────────────────────────────────────────────────────
$('avatarBtn').addEventListener('click', function(e){
  e.stopPropagation();
  $('avatarMenu').classList.toggle('open');
});
document.addEventListener('click', function(e){
  var m = $('avatarMenu');
  if (m.classList.contains('open') && !m.contains(e.target) && e.target !== $('avatarBtn')) {
    m.classList.remove('open');
  }
});
$('menuExhibitions').addEventListener('click', function(){ $('avatarMenu').classList.remove('open'); openExhibitionManager(); });
$('menuReadiness').addEventListener('click', function(){ $('avatarMenu').classList.remove('open'); openReadinessModal(); });
$('menuExport').addEventListener('click', function(){ $('avatarMenu').classList.remove('open'); exportReport(); });
$('exportBtn').addEventListener('click', exportReport);
function exportReport(){
  if (!state.currentExhibition) return;
  window.location.href = '/api/export?exhibition_code=' + encodeURIComponent(state.currentExhibition);
}
$('exhPill').addEventListener('click', openExhibitionManager);

// ── PAGE NAVIGATION ─────────────────────────────────────────────────────
var PAGE_LOADERS = {
  'page-home': renderHome,
  'page-inv': renderInventory,
  'page-sales': renderSales,
  'page-exp': renderExpenses,
  'page-artists': renderArtists,
};
var currentPage = 'page-home';
function navigate(pageId){
  document.querySelectorAll('.page').forEach(function(p){ p.classList.remove('active'); });
  $(pageId).classList.add('active');
  document.querySelectorAll('.bnav-item').forEach(function(n){ n.classList.remove('active'); });
  var active = document.querySelector('.bnav-item[data-page="' + pageId + '"]');
  if (active) active.classList.add('active');
  currentPage = pageId;
  if (PAGE_LOADERS[pageId]) PAGE_LOADERS[pageId]();
}
document.querySelectorAll('.bnav-item').forEach(function(link){
  link.addEventListener('click', function(e){
    e.preventDefault();
    navigate(this.getAttribute('data-page'));
  });
});

// ── EXHIBITION SWITCHING ────────────────────────────────────────────────
async function loadExhibitions(){
  var data = await api('/api/exhibitions');
  state.exhibitions = data.exhibitions;
  if (!state.currentExhibition && state.exhibitions.length) {
    var active = state.exhibitions.find(function(e){ return e.status !== 'completed'; });
    state.currentExhibition = (active || state.exhibitions[0]).code;
  }
  updateExhPill();
}
function updateExhPill(){
  var ex = state.exhibitions.find(function(e){ return e.code === state.currentExhibition; });
  $('exhPillLabel').textContent = ex ? ex.name : (state.currentExhibition || 'No exhibition');
  $('exhDot').className = 'exh-dot' + (ex && ex.status === 'completed' ? ' completed' : '');
  $('exhPeriodLabel').textContent = ex ? periodLabel(ex) : '';
  var order = sortedExhibitions();
  var idx = order.findIndex(function(e){ return e.code === state.currentExhibition; });
  $('exhPrevBtn').disabled = idx <= 0;
  $('exhNextBtn').disabled = idx === -1 || idx >= order.length - 1;
}
function sortedExhibitions(){
  return state.exhibitions.slice().sort(function(a, b){
    var da = a.start_date || '0000-00-00', db = b.start_date || '0000-00-00';
    return da < db ? -1 : (da > db ? 1 : (a.code < b.code ? -1 : 1));
  });
}
function periodLabel(ex){
  if (!ex.start_date && !ex.end_date) return 'No dates set';
  var start = ex.start_date, end = ex.end_date;
  var days = null;
  if (start && end) {
    var d1 = new Date(start), d2 = new Date(end);
    days = Math.max(1, Math.round((d2 - d1) / 86400000) + 1);
  }
  var range = (start || '?') + ' \u2192 ' + (end || 'ongoing');
  return range + (days ? ' \u00b7 ' + days + ' day' + (days === 1 ? '' : 's') : '');
}
async function switchExhibition(code){
  state.currentExhibition = code;
  updateExhPill();
  closeModal();
  await refreshAll();
}
$('exhPrevBtn').addEventListener('click', function(){ stepExhibition(-1); });
$('exhNextBtn').addEventListener('click', function(){ stepExhibition(1); });
function stepExhibition(delta){
  var order = sortedExhibitions();
  var idx = order.findIndex(function(e){ return e.code === state.currentExhibition; });
  var next = idx + delta;
  if (next < 0 || next >= order.length) return;
  switchExhibition(order[next].code);
}

// ── MASTER REFRESH ──────────────────────────────────────────────────────
async function refreshAll(){
  if (!state.currentExhibition) return;
  var code = state.currentExhibition;
  var [dash, artworks, sales, receivables, pending, confirmed, budgets, roi, payables] = await Promise.all([
    api('/api/dashboard?exhibition_code=' + encodeURIComponent(code)),
    api('/api/artworks?exhibition_code=' + encodeURIComponent(code)),
    api('/api/sales?exhibition_code=' + encodeURIComponent(code)),
    api('/api/receivables?exhibition_code=' + encodeURIComponent(code)),
    api('/api/expenses/pending?exhibition_code=' + encodeURIComponent(code)),
    api('/api/expenses/confirmed?exhibition_code=' + encodeURIComponent(code)),
    api('/api/budgets?exhibition_code=' + encodeURIComponent(code)),
    api('/api/artist_roi?exhibition_code=' + encodeURIComponent(code)),
    api('/api/payables?exhibition_code=' + encodeURIComponent(code)),
  ]);
  state.dashboard = dash;
  state.artworks = artworks.artworks;
  state.sales = sales.sales;
  state.receivables = receivables.receivables;
  state.pendingExpenses = pending.pending;
  state.confirmedExpenses = confirmed.confirmed;
  state.budgets = budgets.budgets;
  state.artistRoi = roi.artist_roi;
  state.payables = payables.payables;

  var pendingCount = state.pendingExpenses.length;
  var badge = $('expBadge');
  if (pendingCount > 0) { badge.style.display = 'grid'; badge.textContent = pendingCount > 9 ? '9+' : pendingCount; }
  else badge.style.display = 'none';

  if (!state.accountHeads.length) {
    var heads = await api('/api/account_heads');
    state.accountHeads = heads.account_heads;
  }

  if (PAGE_LOADERS[currentPage]) PAGE_LOADERS[currentPage]();
}

// ============================================================================
// PAGE: HOME
// ============================================================================
function renderHome(){
  var el = $('page-home');
  if (!state.dashboard) { el.innerHTML = '<div class="empty-state">Loading…</div>'; return; }
  var d = state.dashboard;
  var ready = d.readiness.ready;
  var blockCount = d.readiness.blocking_count;

  el.innerHTML = ''
    + '<div class="readiness-banner ' + (ready ? 'ready' : 'warn') + '" id="readinessBanner">'
    +   '<div class="rb-icon">' + (ready
          ? '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>'
          : '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>') + '</div>'
    +   '<div>'
    +     '<div class="rb-title">' + (ready ? 'Ready for review' : blockCount + ' item' + (blockCount === 1 ? '' : 's') + ' need attention') + '</div>'
    +     '<div class="rb-sub">' + (ready ? 'No blocking issues found for this exhibition.' : esc(d.readiness.checks.find(function(c){return c.indexOf('No blocking issues') === -1;}) || 'Tap to review.')) + '</div>'
    +   '</div>'
    +   '<div class="rb-chevron"><svg viewBox="0 0 24 24"><polyline points="9 18 15 12 9 6"/></svg></div>'
    + '</div>'
    + '<div class="stats-row">'
    +   statCard('Gallery Revenue', fmt(d.report.revenue), true, moneyIcon(), 'Net margin ' + d.report.net_margin_pct.toFixed(1) + '%')
    +   statCard('Net Profit / Loss', fmt(d.report.net_profit), false, trendIcon(), d.report.net_profit >= 0 ? 'Profitable' : 'Currently a loss', d.report.net_profit >= 0 ? 'up' : 'down')
    +   statCard('Sell-through Rate', d.metrics.sell_through_rate.toFixed(1) + '%', false, artIcon(), d.metrics.sold_count + ' of ' + d.metrics.total_count + ' works sold')
    +   statCard('Receivables', fmt(d.metrics.receivables), false, receivablesIcon(), d.metrics.pending_receipts + ' receipt(s) pending approval')
    + '</div>'
    + '<div class="mid-row">'
    +   '<div class="chart-card">'
    +     '<div class="card-header"><div><div class="card-title">Cash Position</div><div class="card-sub">Running balance from collected sales &amp; confirmed expenses</div></div></div>'
    +     '<div class="chart-wrap"><canvas id="cashChart"></canvas></div>'
    +   '</div>'
    +   '<div class="donut-card">'
    +     '<div class="card-header"><div><div class="card-title">Expense Breakdown</div><div class="card-sub">By account head</div></div></div>'
    +     (d.expense_breakdown.length
          ? '<div class="donut-wrap"><canvas id="donutChart"></canvas><div class="donut-center"><div class="donut-pct" id="donutTotal">' + fmt(sumBy(d.expense_breakdown, 'amount')) + '</div><div class="donut-pct-sub">confirmed</div></div></div>'
            + '<div class="donut-legend" id="donutLegend"></div>'
          : '<div class="empty-state">No confirmed expenses yet.</div>')
    +   '</div>'
    + '</div>'
    + (d.budget_alerts.length ? budgetAlertsBlock(d.budget_alerts) : '')
    + activityLedgerBlock();

  $('readinessBanner').addEventListener('click', openReadinessModal);

  renderCashChart(d.forecast.projection_labels, d.forecast.projection_values);
  if (d.expense_breakdown.length) renderDonut(d.expense_breakdown);
}

var ledgerFilter = 'all';
function setLedgerFilter(f){ ledgerFilter = f; renderHome(); }
function activityLedgerBlock(){
  var items = [];
  (state.sales || []).forEach(function(s){
    items.push({
      type: 'sale', date: s.sale_date, title: s.title + ' \u2014 sold', sub: (s.buyer_name || 'Unknown buyer') + ' \u00b7 ' + (s.artist || ''),
      amount: Number(s.actual_price_thb || 0), positive: true,
      badge: Number(s.balance_due_thb||0) > 0.01 ? '<span class="status-badge pending">Balance due</span>' : '<span class="status-badge success">Collected</span>',
    });
  });
  (state.confirmedExpenses || []).forEach(function(c){
    items.push({
      type: 'expense', date: (c.created_at || '').slice(0,10), title: c.description || c.account_head, sub: c.account_head,
      amount: Number(c.amount_thb || 0), positive: false,
      badge: '<span class="status-badge failed">Expense</span>',
    });
  });
  items.sort(function(a,b){ return a.date < b.date ? 1 : (a.date > b.date ? -1 : 0); });
  var filtered = ledgerFilter === 'all' ? items : items.filter(function(i){ return i.type === ledgerFilter; });

  var rows = filtered.length ? filtered.slice(0, 60).map(function(i){
    return '<tr><td><div style="font-weight:600;">' + esc(i.title) + '</div><div class="cell-sub">' + esc(i.sub) + ' \u00b7 ' + esc(i.date) + '</div></td>'
      + '<td>' + i.badge + '</td>'
      + '<td style="text-align:right;" class="' + (i.positive ? 'amount-pos' : 'amount-neg') + '">' + (i.positive ? '+' : '\u2212') + fmt(i.amount) + '</td></tr>';
  }).join('') : '<tr><td colspan="3"><div class="empty-state">No activity recorded for this exhibition yet.</div></td></tr>';

  return '<div class="table-card">'
    + '<div class="card-header"><div><div class="card-title">Activity Ledger</div><div class="card-sub">Every sale and expense for this exhibition, in one feed</div></div>'
    + '<div class="table-filter">'
    +   '<button class="filter-btn' + (ledgerFilter==='all'?' active':'') + '" onclick="setLedgerFilter(\'all\')">All</button>'
    +   '<button class="filter-btn' + (ledgerFilter==='sale'?' active':'') + '" onclick="setLedgerFilter(\'sale\')">Sales</button>'
    +   '<button class="filter-btn' + (ledgerFilter==='expense'?' active':'') + '" onclick="setLedgerFilter(\'expense\')">Expenses</button>'
    + '</div></div>'
    + '<div class="txn-wrap"><table class="data-table"><thead><tr><th>Description</th><th>Status</th><th style="text-align:right;">Amount</th></tr></thead><tbody>' + rows + '</tbody></table></div>'
    + '</div>';
}
function budgetAlertsBlock(alerts){
  var rows = alerts.map(function(a){
    return '<div style="display:flex; justify-content:space-between; align-items:center; padding:10px 0; border-bottom:1px solid var(--divider);">'
      + '<div><div style="font-size:13px; font-weight:600;">' + esc(a.account_head) + '</div>'
      + '<div class="cell-sub">Over by ' + fmt(Math.abs(a.variance_thb)) + '</div></div>'
      + '<span class="status-badge failed">Over budget</span></div>';
  }).join('');
  return '<div class="table-card"><div class="card-header"><div><div class="card-title">Budget Alerts</div><div class="card-sub">' + alerts.length + ' account head(s) over budget</div></div></div><div>' + rows + '</div></div>';
}

function statCard(label, value, primary, icon, sub, badgeType){
  return '<div class="stat-card' + (primary ? ' primary' : '') + '">'
    + '<div class="stat-icon-wrap">' + icon + '</div>'
    + '<div class="stat-label">' + label + '</div>'
    + '<div class="stat-value">' + value + '</div>'
    + (sub ? '<div class="stat-sub">' + esc(sub) + '</div>' : '')
    + (badgeType ? '<span class="stat-badge ' + badgeType + '">' + (badgeType === 'up' ? '▲' : '▼') + '</span>' : '')
    + '</div>';
}
function moneyIcon(){ return '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="12" y1="1" x2="12" y2="23"/><path d="M17 5H9.5a3.5 3.5 0 000 7h5a3.5 3.5 0 010 7H6"/></svg>'; }
function trendIcon(){ return '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/></svg>'; }
function artIcon(){ return '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><circle cx="9" cy="9" r="2"/><path d="M21 15l-5-5L5 21"/></svg>'; }
function receivablesIcon(){ return '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="1" y="4" width="22" height="16" rx="2"/><line x1="1" y1="10" x2="23" y2="10"/></svg>'; }
function sumBy(arr, key){ return arr.reduce(function(s, x){ return s + Number(x[key] || 0); }, 0); }

function renderCashChart(labels, values){
  var ctx = document.getElementById('cashChart');
  if (!ctx) return;
  if (!labels.length) { ctx.parentElement.innerHTML = '<div class="empty-state">No cash flow activity yet.</div>'; return; }
  if (state.cashChart) state.cashChart.destroy();
  state.cashChart = new Chart(ctx.getContext('2d'), {
    type: 'line',
    data: { labels: labels, datasets: [{
      label: 'Cash Position', data: values,
      borderColor: getVar('--brand'), borderWidth: 3,
      backgroundColor: hexToRgba(getVar('--brand'), 0.12),
      fill: true, tension: 0.35, pointRadius: 0, pointHoverRadius: 4,
    }]},
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false }, tooltip: { callbacks: { label: function(c){ return fmt(c.parsed.y); } } } },
      scales: {
        y: { grid: { color: getVar('--chart-grid') }, ticks: { color: getVar('--text3'), font: { size: 10 }, callback: function(v){ return fmt(v); } } },
        x: { grid: { display: false }, ticks: { color: getVar('--text3'), font: { size: 10 }, maxTicksLimit: 6 } },
      },
    },
  });
}
function hexToRgba(hex, alpha){
  hex = hex.trim();
  if (hex[0] !== '#') return hex;
  var r = parseInt(hex.slice(1,3),16), g = parseInt(hex.slice(3,5),16), b = parseInt(hex.slice(5,7),16);
  return 'rgba(' + r + ',' + g + ',' + b + ',' + alpha + ')';
}
var DONUT_COLORS = ['#3763a8','#6f97d6','#34c759','#ff9500','#ff3b30','#af52de','#5ac8fa','#ffcc00','#a2845e','#8e8e93'];
function renderDonut(breakdown){
  var ctx = document.getElementById('donutChart');
  if (!ctx) return;
  if (state.donutChart) state.donutChart.destroy();
  var top = breakdown.slice(0, 6);
  var rest = breakdown.slice(6);
  var labels = top.map(function(x){ return x.account_head; });
  var data = top.map(function(x){ return x.amount; });
  if (rest.length) { labels.push('Other'); data.push(sumBy(rest, 'amount')); }
  state.donutChart = new Chart(ctx.getContext('2d'), {
    type: 'doughnut',
    data: { labels: labels, datasets: [{ data: data, backgroundColor: DONUT_COLORS, borderWidth: 0 }] },
    options: { responsive: true, maintainAspectRatio: false, cutout: '72%', plugins: { legend: { display: false }, tooltip: { callbacks: { label: function(c){ return c.label + ': ' + fmt(c.parsed); } } } } },
  });
  var total = sumBy(breakdown, 'amount');
  $('donutLegend').innerHTML = labels.map(function(l, i){
    var val = data[i];
    var pct = total ? (val / total * 100).toFixed(0) : 0;
    return '<div class="legend-item"><span class="legend-dot" style="background:' + DONUT_COLORS[i % DONUT_COLORS.length] + '"></span><span class="legend-label">' + esc(l) + '</span><span class="legend-val">' + pct + '%</span></div>';
  }).join('');
}

// ============================================================================
// PAGE: INVENTORY
// ============================================================================
function renderInventory(){
  var el = $('page-inv');
  var arts = state.artworks || [];
  var total = arts.length, sold = arts.filter(function(a){return a.status==='sold';}).length;
  var available = total - sold;
  var unsoldValue = arts.filter(function(a){return a.status!=='sold';}).reduce(function(s,a){return s+Number(a.asking_price_thb||0);},0);

  var rows = arts.length ? arts.map(function(a){
    var soldBadge = a.status === 'sold' ? '<span class="status-badge success">Sold</span>' : '<span class="status-badge pending">Available</span>';
    var action = a.status === 'sold' ? '' : '<button class="mini-btn go" onclick="openSaleModal(' + a.id + ')">Record Sale</button>';
    return '<tr><td><div style="font-weight:600;">' + esc(a.title) + '</div><div class="cell-sub">#' + a.id + '</div></td>'
      + '<td>' + esc(a.artist) + '</td>'
      + '<td>' + fmt(a.asking_price_thb) + '</td>'
      + '<td>' + soldBadge + '</td>'
      + '<td style="text-align:right;">' + action + '</td></tr>';
  }).join('') : '';

  el.innerHTML = ''
    + '<div class="stats-row">'
    +   statCard('Registered', String(total), false, artIcon())
    +   statCard('Sold', String(sold), false, trendIcon())
    +   statCard('Available', String(available), false, moneyIcon())
    +   statCard('Unsold Asking Value', fmt(unsoldValue), false, receivablesIcon())
    + '</div>'
    + '<div class="table-card">'
    +   '<div class="card-header"><div><div class="card-title">Artworks</div><div class="card-sub">' + total + ' registered for this exhibition</div></div>'
    +   '<div class="card-actions"><button class="chip-btn" onclick="openBulkImportModal()">Bulk Import</button><button class="chip-btn primary" onclick="openArtworkModal()">+ Add Artwork</button></div></div>'
    +   '<div class="txn-wrap"><table class="data-table"><thead><tr><th>Title</th><th>Artist</th><th>Asking Price</th><th>Status</th><th></th></tr></thead>'
    +   '<tbody>' + (rows || '<tr><td colspan="5"><div class="empty-state">No artworks registered yet. Add one to get started.</div></td></tr>') + '</tbody></table></div>'
    + '</div>';
}

// ============================================================================
// PAGE: SALES & RECEIVABLES
// ============================================================================
function renderSales(){
  var el = $('page-sales');
  var sales = state.sales || [];
  var grossSales = sumBy(sales, 'actual_price_thb');
  var collected = sumBy(sales, 'amount_collected_thb');
  var receivables = sumBy(sales, 'balance_due_thb');
  var avg = sales.length ? grossSales / sales.length : 0;

  var rows = sales.length ? sales.slice().reverse().map(function(s){
    var balance = Number(s.balance_due_thb || 0);
    var badge = balance <= 0.01 ? '<span class="status-badge success">Collected</span>' : (Number(s.amount_collected_thb||0) > 0 ? '<span class="status-badge pending">Partial</span>' : '<span class="status-badge failed">Uncollected</span>');
    var action = balance > 0.01 ? '<button class="mini-btn go" onclick="openCollectModal(' + s.id + ',' + balance + ')">Collect</button>' : '';
    return '<tr><td><div style="font-weight:600;">' + esc(s.title) + '</div><div class="cell-sub">' + esc(s.artist) + ' \u00b7 ' + esc(s.sale_date) + '</div></td>'
      + '<td>' + esc(s.buyer_name || '\u2014') + '</td>'
      + '<td>' + fmt(s.actual_price_thb) + '</td>'
      + '<td>' + fmt(s.amount_collected_thb) + '</td>'
      + '<td>' + (balance > 0.01 ? '<span style="color:var(--down); font-weight:600;">' + fmt(balance) + '</span>' : fmt(0)) + '</td>'
      + '<td>' + badge + '</td>'
      + '<td style="text-align:right;">' + action + '</td></tr>';
  }).join('') : '';

  el.innerHTML = ''
    + '<div class="stats-row">'
    +   statCard('Gross Sales', fmt(grossSales), true, moneyIcon())
    +   statCard('Cash Collected', fmt(collected), false, trendIcon(), null, 'up')
    +   statCard('Receivables', fmt(receivables), false, receivablesIcon(), receivables > 0 ? 'Outstanding balances' : 'All collected', receivables > 0 ? 'down' : 'up')
    +   statCard('Avg Sale Price', fmt(avg), false, artIcon())
    + '</div>'
    + '<div class="table-card">'
    +   '<div class="card-header"><div><div class="card-title">Sales Ledger</div><div class="card-sub">' + sales.length + ' recorded sale(s)</div></div></div>'
    +   '<div class="txn-wrap"><table class="data-table"><thead><tr><th>Artwork</th><th>Buyer</th><th>Sale Price</th><th>Collected</th><th>Balance</th><th>Status</th><th></th></tr></thead>'
    +   '<tbody>' + (rows || '<tr><td colspan="7"><div class="empty-state">No sales recorded yet.</div></td></tr>') + '</tbody></table></div>'
    + '</div>';
}

// ============================================================================
// PAGE: EXPENSES & BUDGET
// ============================================================================
function renderExpenses(){
  var el = $('page-exp');
  var pending = state.pendingExpenses || [];
  var confirmed = state.confirmedExpenses || [];
  var budgets = (state.budgets || []).filter(function(b){ return b.budget_thb > 0 || b.actual_thb > 0; });

  var pendingCards = pending.length ? pending.map(function(p){
    return '<div style="border:1px solid var(--divider); border-radius:14px; padding:14px; margin-bottom:10px;">'
      + '<div style="display:flex; justify-content:space-between; gap:10px;"><div style="font-weight:600; font-size:13.5px;">' + esc(p.description || 'Receipt') + '</div><div style="font-weight:700;">' + fmt(p.suggested_amount_thb) + '</div></div>'
      + '<div class="cell-sub" style="margin-bottom:10px;">Suggested: ' + esc(p.suggested_account_head) + '</div>'
      + '<div style="display:flex; gap:8px; flex-wrap:wrap;">'
      +   '<button class="mini-btn go" onclick="confirmPendingExpense(' + p.id + ')">Confirm</button>'
      +   '<button class="mini-btn" onclick="openReassignModal(' + p.id + ',' + p.suggested_amount_thb + ')">Change Account / Amount</button>'
      +   '<button class="mini-btn danger" onclick="ignorePendingExpense(' + p.id + ')">Ignore</button>'
      + '</div></div>';
  }).join('') : '<div class="empty-state">No pending receipts.</div>';

  var confirmedRows = confirmed.length ? confirmed.slice().reverse().slice(0, 40).map(function(c){
    var sub = [c.account_head, c.category, c.recipient].filter(Boolean).join(' \u00b7 ');
    return '<tr><td><div style="font-weight:600;">' + esc(c.description || c.account_head) + '</div><div class="cell-sub">' + esc(sub) + '</div></td>'
      + '<td>' + esc((c.created_at||'').slice(0,10)) + '</td>'
      + '<td style="text-align:right;">' + fmt(c.amount_thb) + '</td></tr>';
  }).join('') : '<tr><td colspan="3"><div class="empty-state">No confirmed expenses yet.</div></td></tr>';

  var budgetRows = budgets.length ? budgets.map(function(b){
    var pct = b.budget_thb > 0 ? Math.min(100, (b.actual_thb / b.budget_thb) * 100) : (b.actual_thb > 0 ? 100 : 0);
    var over = b.budget_thb > 0 && b.variance_thb < 0;
    return '<div class="progress-item">'
      + '<div class="progress-header"><span class="progress-label">' + esc(b.account_head) + '</span><span class="progress-val">' + fmt(b.actual_thb) + ' / ' + (b.budget_thb > 0 ? fmt(b.budget_thb) : 'no budget') + '</span></div>'
      + '<div class="progress-track"><div class="progress-fill' + (over ? ' over' : '') + '" style="width:' + pct + '%"></div></div>'
      + '<div style="display:flex; justify-content:space-between; margin-top:4px;"><span class="cell-sub">' + (over ? 'Over by ' + fmt(Math.abs(b.variance_thb)) : (b.budget_thb > 0 ? 'Remaining ' + fmt(b.variance_thb) : '')) + '</span>'
      + '<span class="cell-sub" style="cursor:pointer; color:var(--brand);" onclick="openSetBudgetModal(\'' + esc(b.account_head).replace(/'/g,"\\'") + '\',' + b.budget_thb + ')">Set budget</span></div>'
      + '</div>';
  }).join('') : '<div class="empty-state">No expenses or budgets recorded yet.</div>';

  el.innerHTML = ''
    + '<div class="table-card">'
    +   '<div class="card-header"><div><div class="card-title">Pending Receipts</div><div class="card-sub">' + pending.length + ' awaiting approval</div></div></div>'
    +   '<div style="margin-top:14px;">' + pendingCards + '</div>'
    + '</div>'
    + '<div class="table-card">'
    +   '<div class="card-header"><div><div class="card-title">Budget vs Actual</div><div class="card-sub">By account head</div></div><button class="chip-btn" onclick="openSetBudgetModal(\'\',0)">+ Set Budget</button></div>'
    +   '<div class="progress-list">' + budgetRows + '</div>'
    + '</div>'
    + '<div class="table-card">'
    +   '<div class="card-header"><div><div class="card-title">Confirmed Expense Ledger</div><div class="card-sub">' + confirmed.length + ' confirmed</div></div></div>'
    +   '<div class="txn-wrap"><table class="data-table"><thead><tr><th>Description</th><th>Date</th><th style="text-align:right;">Amount</th></tr></thead><tbody>' + confirmedRows + '</tbody></table></div>'
    + '</div>';
}

// ============================================================================
// PAGE: ARTISTS
// ============================================================================
function renderArtists(){
  var el = $('page-artists');
  var roi = state.artistRoi || [];
  var payables = state.payables || [];

  var roiRows = roi.length ? roi.map(function(r){
    var net = Number(r.net_contribution || 0);
    return '<tr><td style="font-weight:600;">' + esc(r.artist) + '</td>'
      + '<td>' + r.sold_artworks + ' / ' + r.total_artworks + '</td>'
      + '<td>' + fmt(r.gallery_share) + '</td>'
      + '<td>' + fmt(r.direct_costs) + '</td>'
      + '<td class="' + (net >= 0 ? 'amount-pos' : 'amount-neg') + '">' + fmt(net) + '</td>'
      + '<td>' + r.sell_through_rate.toFixed(0) + '%</td></tr>';
  }).join('') : '<tr><td colspan="6"><div class="empty-state">No artists registered yet.</div></td></tr>';

  var payRows = payables.length ? payables.map(function(p){
    var out = Number(p.outstanding_thb || 0);
    var badge = out <= 0.01 ? '<span class="status-badge success">Paid</span>' : '<span class="status-badge pending">Pending</span>';
    var action = out > 0.01 ? '<button class="mini-btn go" onclick="openPayArtistModal(' + p.id + ',' + out + ')">Pay</button>' : '';
    return '<tr><td style="font-weight:600;">' + esc(p.artist) + '</td>'
      + '<td>' + fmt(p.gross_sale_thb) + '</td>'
      + '<td>' + fmt(p.artist_payable_thb) + '</td>'
      + '<td>' + fmt(p.paid_thb) + '</td>'
      + '<td>' + fmt(out) + '</td>'
      + '<td>' + badge + '</td>'
      + '<td style="text-align:right;">' + action + '</td></tr>';
  }).join('') : '<tr><td colspan="7"><div class="empty-state">No artist payables yet.</div></td></tr>';

  el.innerHTML = ''
    + '<div class="table-card">'
    +   '<div class="card-header"><div><div class="card-title">Artist ROI Matrix</div><div class="card-sub">Gallery share minus direct costs tagged to each artist</div></div></div>'
    +   '<div class="txn-wrap"><table class="data-table"><thead><tr><th>Artist</th><th>Works</th><th>Gallery Share</th><th>Direct Costs</th><th>Net Contribution</th><th>Sell-through</th></tr></thead><tbody>' + roiRows + '</tbody></table></div>'
    + '</div>'
    + '<div class="table-card">'
    +   '<div class="card-header"><div><div class="card-title">Artist Payables</div><div class="card-sub">Settlement status per sale</div></div></div>'
    +   '<div class="txn-wrap"><table class="data-table"><thead><tr><th>Artist</th><th>Gross Sale</th><th>Payable</th><th>Paid</th><th>Outstanding</th><th>Status</th><th></th></tr></thead><tbody>' + payRows + '</tbody></table></div>'
    + '</div>';
}

// ============================================================================
// MODAL SYSTEM
// ============================================================================
function openModal(html){
  $('modalContent').innerHTML = html;
  $('modalOverlay').classList.add('open');
}
function closeModal(){
  $('modalOverlay').classList.remove('open');
  $('modalContent').innerHTML = '';
}
$('modalOverlay').addEventListener('click', function(e){ if (e.target === $('modalOverlay')) closeModal(); });

function modalShell(title, iconSvg, iconColor, bodyHtml){
  return '<div class="qa-modal-header"><div class="qa-modal-title-wrap"><div class="qa-modal-icon" style="background:' + iconColor + '22; color:' + iconColor + ';">' + iconSvg + '</div><div class="qa-modal-title">' + title + '</div></div>'
    + '<button class="qa-modal-close" onclick="closeModal()"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg></button></div>'
    + '<div class="qa-modal-body">' + bodyHtml + '</div>';
}
function errorBox(id){ return '<div class="qa-error" id="' + id + '"></div>'; }
function showErr(id, msg){ var e = $(id); e.textContent = msg; e.classList.add('show'); }

// ── Add Artwork ──
function openArtworkModal(){
  openModal(modalShell('Add Artwork', artIcon(), getVar('--brand'), ''
    + '<div class="qa-field"><span class="qa-label">Title</span><input class="qa-input" id="fArtTitle" placeholder="e.g. Golden Pagoda I"></div>'
    + '<div class="qa-field"><span class="qa-label">Artist</span><input class="qa-input" id="fArtArtist" placeholder="Artist name"></div>'
    + '<div class="qa-field"><span class="qa-label">Asking Price (THB)</span><input class="qa-input" type="number" id="fArtPrice" placeholder="0"></div>'
    + errorBox('artErr')
    + '<button class="qa-submit-btn" onclick="submitArtwork()">Add Artwork</button>'));
}
async function submitArtwork(){
  try {
    await api('/api/artworks', {method:'POST', body:{exhibition_code: state.currentExhibition, title: $('fArtTitle').value, artist: $('fArtArtist').value, price: $('fArtPrice').value}});
    closeModal(); await refreshAll();
  } catch(e) { showErr('artErr', e.message); }
}

// ── Bulk Import ──
function openBulkImportModal(){
  openModal(modalShell('Bulk Import Artworks', artIcon(), getVar('--brand'), ''
    + '<div class="qa-field"><span class="qa-label">📁 Upload Excel / CSV File (.xlsx, .csv)</span>'
    + '<input type="file" class="qa-input" id="fBulkFile" accept=".xlsx,.xls,.csv"></div>'
    + '<div style="text-align:center; font-size:11px; color:var(--text3); margin:8px 0;">— OR PASTE TEXT —</div>'
    + '<div class="qa-hint">One artwork per line: Title, Artist, Price</div>'
    + '<textarea class="qa-textarea" id="fBulkData" placeholder="Sunsets, Mike, 50000&#10;Ocean Breeze, Mike, 75000"></textarea>'
    + errorBox('bulkErr')
    + '<button class="qa-submit-btn" onclick="submitBulkArtworks()">Import Artworks</button>'));
}
async function submitBulkArtworks(){
  var fileInput = $('fBulkFile');
  if (fileInput && fileInput.files && fileInput.files[0]) {
    var fd = new FormData();
    fd.append('exhibition_code', state.currentExhibition);
    fd.append('file', fileInput.files[0]);
    try {
      var res = await fetch('/api/artworks/bulk-file', {method:'POST', body:fd, credentials:'include'});
      var data = await res.json();
      if (!res.ok) throw new Error(data.detail || 'Upload failed');
      closeModal(); await refreshAll();
    } catch(e) { showErr('bulkErr', e.message); }
  } else {
    try {
      var data = await api('/api/artworks/bulk', {method:'POST', body:{exhibition_code: state.currentExhibition, raw_data: $('fBulkData').value}});
      closeModal(); await refreshAll();
    } catch(e) { showErr('bulkErr', e.message); }
  }
}

// ── Record Sale ──
function openSaleModal(preselectId){
  var available = (state.artworks || []).filter(function(a){ return a.status !== 'sold'; });
  var options = available.map(function(a){
    return '<option value="' + a.id + '" data-price="' + a.asking_price_thb + '"' + (preselectId && a.id === preselectId ? ' selected' : '') + '>' + esc(a.title) + ' \u2014 ' + esc(a.artist) + ' (' + fmt(a.asking_price_thb) + ')</option>';
  }).join('');
  var pre = available.find(function(a){ return a.id === preselectId; });
  openModal(modalShell('Record Sale', trendIcon(), '#34c759', ''
    + '<div class="qa-field"><span class="qa-label">Artwork</span><select class="qa-select" id="fSaleArtwork" onchange="onSaleArtworkChange()">' + (options || '<option disabled selected>No available artworks</option>') + '</select></div>'
    + '<div class="qa-row2">'
    +   '<div class="qa-field"><span class="qa-label">Sale Price (THB)</span><input class="qa-input" type="number" id="fSalePrice" value="' + (pre ? pre.asking_price_thb : '') + '"></div>'
    +   '<div class="qa-field"><span class="qa-label">Collected Now (THB)</span><input class="qa-input" type="number" id="fSaleCollected" placeholder="Full amount if blank"></div>'
    + '</div>'
    + '<div class="qa-field"><span class="qa-label">Buyer Name</span><input class="qa-input" id="fSaleBuyer" placeholder="Optional"></div>'
    + '<div class="qa-field"><span class="qa-label">Payment Method</span><input class="qa-input" id="fSaleMethod" placeholder="e.g. Bank transfer, cash"></div>'
    + errorBox('saleErr')
    + '<button class="qa-submit-btn" onclick="submitSale()">Confirm Sale</button>'));
}
function onSaleArtworkChange(){
  var sel = $('fSaleArtwork');
  var opt = sel.options[sel.selectedIndex];
  if (opt) $('fSalePrice').value = opt.getAttribute('data-price');
}
async function submitSale(){
  var artworkId = $('fSaleArtwork').value;
  if (!artworkId) { showErr('saleErr', 'No artwork selected.'); return; }
  try {
    await api('/api/sales', {method:'POST', body:{
      artwork_id: artworkId, price: $('fSalePrice').value,
      collected: $('fSaleCollected').value || null,
      buyer: $('fSaleBuyer').value, payment_method: $('fSaleMethod').value,
    }});
    closeModal(); await refreshAll();
  } catch(e) { showErr('saleErr', e.message); }
}

// ── Collect Payment ──
function openCollectModal(saleId, balance){
  openModal(modalShell('Collect Payment', receivablesIcon(), getVar('--brand'), ''
    + '<div class="qa-note">Balance due: ' + fmt(balance) + '</div>'
    + '<div class="qa-field"><span class="qa-label">Amount Collected (THB)</span><input class="qa-input" type="number" id="fCollectAmt" value="' + balance + '"></div>'
    + errorBox('collectErr')
    + '<button class="qa-submit-btn" onclick="submitCollect(' + saleId + ')">Record Collection</button>'));
}
async function submitCollect(saleId){
  try {
    await api('/api/sales/collect', {method:'POST', body:{sale_id: saleId, amount: $('fCollectAmt').value}});
    closeModal(); await refreshAll();
  } catch(e) { showErr('collectErr', e.message); }
}

// ── Log Expense / Receipt Scan ──
function openExpenseModal(){
  var opts = state.accountHeads.map(function(h){ return '<option value="' + esc(h) + '">' + esc(h) + '</option>'; }).join('');
  openModal(modalShell('Log Expense / Scan Receipt', moneyIcon(), '#ff3b30', ''
    + '<div class="qa-field"><span class="qa-label">📷 Scan Receipt Image (Auto-detects Amount &amp; Category)</span>'
    + '<input type="file" class="qa-input" id="fExpScanFile" accept="image/*" onchange="onScanReceiptFileSelected()"></div>'
    + '<div style="text-align:center; font-size:11px; color:var(--text3); margin:8px 0;">— OR ENTER MANUALLY —</div>'
    + '<div class="qa-field"><span class="qa-label">Amount (THB)</span><input class="qa-input" type="number" id="fExpAmt" placeholder="0"></div>'
    + '<div class="qa-field"><span class="qa-label">Description</span><input class="qa-input" id="fExpDesc" placeholder="e.g. Venue deposit"></div>'
    + '<div class="qa-row2">'
    +   '<div class="qa-field"><span class="qa-label">Paid To (Recipient)</span><input class="qa-input" id="fExpRecipient" placeholder="e.g. CMU Art Center"></div>'
    +   '<div class="qa-field"><span class="qa-label">Category</span><input class="qa-input" id="fExpCategory" placeholder="e.g. Rental Fees"></div>'
    + '</div>'
    + '<div class="qa-field"><span class="qa-label">Account Head</span><select class="qa-select" id="fExpHead"><option value="">Auto-detect from description</option>' + opts + '</select></div>'
    + '<div class="qa-field"><span class="qa-label">Tag to Artist (optional)</span><input class="qa-input" id="fExpArtist" placeholder="Leave blank for gallery-wide"></div>'
    + errorBox('expErr')
    + '<button class="qa-submit-btn" onclick="submitExpense()">Log &amp; Confirm Expense</button>'));
}
async function onScanReceiptFileSelected(){
  var fileInput = $('fExpScanFile');
  if (!fileInput || !fileInput.files || !fileInput.files[0]) return;
  var fd = new FormData();
  fd.append('exhibition_code', state.currentExhibition);
  fd.append('file', fileInput.files[0]);
  fd.append('notes', $('fExpDesc').value || fileInput.files[0].name);
  try {
    var res = await fetch('/api/ai/scan-receipt', {method:'POST', body:fd, credentials:'include'});
    var data = await res.json();
    if (!res.ok) throw new Error(data.detail || 'Scan failed');
    closeModal(); await refreshAll();
  } catch(e) { showErr('expErr', e.message); }
}
async function submitExpense(){
  try {
    await api('/api/expenses', {method:'POST', body:{
      exhibition_code: state.currentExhibition, amount: $('fExpAmt').value, description: $('fExpDesc').value,
      account_head: $('fExpHead').value || null, artist_tag: $('fExpArtist').value || null,
      recipient: $('fExpRecipient').value || null, category: $('fExpCategory').value || null,
    }});
    closeModal(); await refreshAll();
  } catch(e) { showErr('expErr', e.message); }
}

// ── Pending expense actions ──
async function confirmPendingExpense(id){
  try { await api('/api/expenses/confirm', {method:'POST', body:{pending_id:id}}); await refreshAll(); }
  catch(e) { alert(e.message); }
}
async function ignorePendingExpense(id){
  try { await api('/api/expenses/ignore', {method:'POST', body:{pending_id:id}}); await refreshAll(); }
  catch(e) { alert(e.message); }
}
function openReassignModal(pendingId, currentAmount){
  var opts = state.accountHeads.map(function(h){ return '<option value="' + esc(h) + '">' + esc(h) + '</option>'; }).join('');
  openModal(modalShell('Update Receipt', moneyIcon(), '#ff9500', ''
    + '<div class="qa-field"><span class="qa-label">Account Head</span><select class="qa-select" id="fReassignHead">' + opts + '</select></div>'
    + '<div class="qa-field"><span class="qa-label">Amount (THB)</span><input class="qa-input" type="number" id="fReassignAmt" value="' + currentAmount + '"></div>'
    + errorBox('reassignErr')
    + '<button class="qa-submit-btn" onclick="submitReassign(' + pendingId + ')">Save &amp; Confirm</button>'));
}
async function submitReassign(pendingId){
  try {
    await api('/api/expenses/confirm', {method:'POST', body:{pending_id: pendingId, account_head: $('fReassignHead').value, amount: $('fReassignAmt').value}});
    closeModal(); await refreshAll();
  } catch(e) { showErr('reassignErr', e.message); }
}

// ── Set Budget ──
function openSetBudgetModal(accountHead, currentBudget){
  var opts = state.accountHeads.map(function(h){ return '<option value="' + esc(h) + '"' + (h === accountHead ? ' selected' : '') + '>' + esc(h) + '</option>'; }).join('');
  openModal(modalShell('Set Budget', receivablesIcon(), getVar('--brand'), ''
    + '<div class="qa-field"><span class="qa-label">Account Head</span><select class="qa-select" id="fBudgetHead">' + opts + '</select></div>'
    + '<div class="qa-field"><span class="qa-label">Budget (THB)</span><input class="qa-input" type="number" id="fBudgetAmt" value="' + (currentBudget || '') + '"></div>'
    + errorBox('budgetErr')
    + '<button class="qa-submit-btn" onclick="submitBudget()">Save Budget</button>'));
}
async function submitBudget(){
  try {
    await api('/api/budgets', {method:'POST', body:{exhibition_code: state.currentExhibition, account_head: $('fBudgetHead').value, budget: $('fBudgetAmt').value}});
    closeModal(); await refreshAll();
  } catch(e) { showErr('budgetErr', e.message); }
}

// ── Pay Artist ──
function openPayArtistModal(payableId, outstanding){
  openModal(modalShell('Pay Artist', artIcon(), getVar('--brand'), ''
    + '<div class="qa-note">Outstanding: ' + fmt(outstanding) + '</div>'
    + '<div class="qa-field"><span class="qa-label">Payment Amount (THB)</span><input class="qa-input" type="number" id="fPayAmt" value="' + outstanding + '"></div>'
    + errorBox('payErr')
    + '<button class="qa-submit-btn" onclick="submitPayArtist(' + payableId + ')">Record Payment</button>'));
}
async function submitPayArtist(payableId){
  try {
    await api('/api/payables/pay', {method:'POST', body:{payable_id: payableId, amount: $('fPayAmt').value}});
    closeModal(); await refreshAll();
  } catch(e) { showErr('payErr', e.message); }
}

// ── New Exhibition ──
function openNewExhibitionModal(){
  openModal(modalShell('New Exhibition', artIcon(), getVar('--brand'), ''
    + '<div class="qa-row2">'
    +   '<div class="qa-field"><span class="qa-label">Code</span><input class="qa-input" id="fExhCode" placeholder="e.g. AUTUMN2026"></div>'
    +   '<div class="qa-field"><span class="qa-label">Name</span><input class="qa-input" id="fExhName" placeholder="Exhibition name"></div>'
    + '</div>'
    + '<div class="qa-field"><span class="qa-label">Location</span><input class="qa-input" id="fExhLoc" placeholder="Optional"></div>'
    + '<div class="qa-row2">'
    +   '<div class="qa-field"><span class="qa-label">Start Date</span><input class="qa-input" type="date" id="fExhStart"></div>'
    +   '<div class="qa-field"><span class="qa-label">End Date</span><input class="qa-input" type="date" id="fExhEnd"></div>'
    + '</div>'
    + '<div class="qa-hint">New exhibitions start with a default 50/50 gallery/artist split — adjust it from Exhibitions &amp; Splits.</div>'
    + errorBox('newExhErr')
    + '<button class="qa-submit-btn" onclick="submitNewExhibition()">Create Exhibition</button>'));
}
async function submitNewExhibition(){
  try {
    var data = await api('/api/exhibitions', {method:'POST', body:{
      code: $('fExhCode').value, name: $('fExhName').value, location: $('fExhLoc').value,
      start_date: $('fExhStart').value || null, end_date: $('fExhEnd').value || null,
    }});
    await loadExhibitions();
    await switchExhibition(data.exhibition.code);
  } catch(e) { showErr('newExhErr', e.message); }
}

// ── Exhibition Manager (switch / splits / close) ──
async function openExhibitionManager(){
  $('avatarMenu').classList.remove('open');
  var list = state.exhibitions.map(function(e){
    var active = e.code === state.currentExhibition;
    return '<div class="mini-btn' + (active ? ' go' : '') + '" style="margin:0 6px 8px 0; display:inline-block;" onclick="switchExhibition(\'' + e.code + '\')">' + esc(e.name) + (e.status === 'completed' ? ' \u2713' : '') + '</div>';
  }).join('');
  var splits = await api('/api/splits?exhibition_code=' + encodeURIComponent(state.currentExhibition));
  var splitRows = splits.splits.length ? splits.splits.map(function(s){
    return '<div style="display:flex; justify-content:space-between; padding:6px 0; font-size:13px;"><span>' + esc(s.party_type) + ' \u2014 ' + esc(s.party_name) + '</span><span style="font-weight:600;">' + s.percent + '%</span></div>';
  }).join('') : '<div class="cell-sub">No split set.</div>';

  openModal(modalShell('Exhibitions &amp; Splits', artIcon(), getVar('--brand'), ''
    + '<div class="qa-label">Switch Exhibition</div>'
    + '<div style="margin-bottom:8px;">' + list + '</div>'
    + '<button class="chip-btn" style="width:100%; justify-content:center;" onclick="openNewExhibitionModal()">+ New Exhibition</button>'
    + '<div class="qa-modal-title" style="margin-top:8px; font-size:13px;">Commission Split — ' + esc(state.currentExhibition) + '</div>'
    + splitRows
    + '<div class="qa-row2">'
    +   '<div class="qa-field"><span class="qa-label">Gallery %</span><input class="qa-input" id="fSplitGallery" placeholder="50"></div>'
    +   '<div class="qa-field"><span class="qa-label">Artist %</span><input class="qa-input" id="fSplitArtist" placeholder="50"></div>'
    + '</div>'
    + '<div class="qa-row2">'
    +   '<div class="qa-field"><span class="qa-label">Collaborator Name (optional)</span><input class="qa-input" id="fSplitCollabName" placeholder="e.g. Curator"></div>'
    +   '<div class="qa-field"><span class="qa-label">Collaborator %</span><input class="qa-input" id="fSplitCollabPct" placeholder="0"></div>'
    + '</div>'
    + '<div class="qa-hint">Percentages must total 100%.</div>'
    + errorBox('splitErr')
    + '<button class="qa-submit-btn" onclick="submitSplits()">Save Split</button>'
    + '<div class="avatar-menu-divider"></div>'
    + '<button class="qa-submit-btn danger" onclick="submitCloseExhibition()">Close This Exhibition</button>'
  ));
}
async function submitSplits(){
  var entries = [];
  var g = parseFloat($('fSplitGallery').value), a = parseFloat($('fSplitArtist').value);
  var cn = $('fSplitCollabName').value.trim(), cp = parseFloat($('fSplitCollabPct').value);
  if (g) entries.push({party_type:'gallery', party_name:'Gallery', percent:g});
  if (a) entries.push({party_type:'artist', party_name:'Artist', percent:a});
  if (cn && cp) entries.push({party_type:'collaborator', party_name:cn, percent:cp});
  try {
    await api('/api/splits', {method:'POST', body:{exhibition_code: state.currentExhibition, entries: entries}});
    closeModal(); await refreshAll();
  } catch(e) { showErr('splitErr', e.message); }
}
async function submitCloseExhibition(){
  if (!confirm('Close ' + state.currentExhibition + '? This marks it completed.')) return;
  try {
    await api('/api/exhibitions/close', {method:'POST', body:{exhibition_code: state.currentExhibition}});
    closeModal(); await loadExhibitions(); await refreshAll();
  } catch(e) { alert(e.message); }
}

// ── Readiness Modal ──
async function openReadinessModal(){
  $('avatarMenu').classList.remove('open');
  var data = await api('/api/readiness?exhibition_code=' + encodeURIComponent(state.currentExhibition));
  var rows = data.checks.map(function(c){
    var ok = c.indexOf('No blocking issues') !== -1;
    return '<div style="display:flex; gap:10px; padding:8px 0; border-bottom:1px solid var(--divider); align-items:flex-start;">'
      + '<span style="flex-shrink:0; margin-top:2px;">' + (ok ? '\u2705' : '\u26a0\ufe0f') + '</span><span style="font-size:13px;">' + esc(c) + '</span></div>';
  }).join('');
  openModal(modalShell(data.ready ? 'Ready for Review' : 'Needs Attention', trendIcon(), data.ready ? '#34c759' : '#ff3b30', rows));
}

// ============================================================================
// FAB
// ============================================================================
$('fabAddBtn').addEventListener('click', function(){
  $('fabWrap').classList.toggle('open');
  $('fabBackdrop').classList.toggle('open');
});
$('fabBackdrop').addEventListener('click', function(){
  $('fabWrap').classList.remove('open');
  $('fabBackdrop').classList.remove('open');
});
function closeFab(){ $('fabWrap').classList.remove('open'); $('fabBackdrop').classList.remove('open'); }
$('qaOpenArtwork').addEventListener('click', function(){ closeFab(); openArtworkModal(); });
$('qaOpenSale').addEventListener('click', function(){ closeFab(); openSaleModal(); });
$('qaOpenExpense').addEventListener('click', function(){ closeFab(); openExpenseModal(); });

// ============================================================================
// SEA AI PANEL
// ============================================================================
$('seaBtn').addEventListener('click', function(){ $('seaPanel').classList.add('open'); });
$('seaClose').addEventListener('click', function(){ $('seaPanel').classList.remove('open'); });
$('seaPanel').addEventListener('click', function(e){ if (e.target === $('seaPanel')) $('seaPanel').classList.remove('open'); });
$('seaSend').addEventListener('click', sendSeaMessage);
$('seaInput').addEventListener('keydown', function(e){ if (e.key === 'Enter') sendSeaMessage(); });
window.sendSeaQuery = function(text){
  $('seaInput').value = text;
  sendSeaMessage();
};

function renderMarkdown(md){
  if (!md) return '';
  var html = esc(md);
  html = html.replace(/^### (.*$)/gim, '<h3 style="margin:8px 0 4px; font-weight:700; font-size:14px; color:var(--text);">$1</h3>');
  html = html.replace(/^#### (.*$)/gim, '<h4 style="margin:6px 0 3px; font-weight:600; font-size:13px; color:var(--text);">$1</h4>');
  html = html.replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>');
  html = html.replace(/`([^`]+)`/g, '<code style="background:var(--badge-bg); padding:1px 5px; border-radius:4px; font-size:11px;">$1</code>');
  html = html.replace(/^\- (.*$)/gim, '• $1<br/>');
  html = html.replace(/^1\. (.*$)/gim, '1. $1<br/>');
  html = html.replace(/^2\. (.*$)/gim, '2. $1<br/>');
  html = html.replace(/\n\n/g, '<br/><br/>');
  return html;
}

async function sendSeaMessage(){
  var input = $('seaInput');
  var text = input.value.trim();
  if (!text) return;
  input.value = '';
  var msgs = $('seaMessages');
  msgs.insertAdjacentHTML('beforeend', '<div class="sea-msg user"><div class="sea-msg-bubble">' + esc(text) + '</div></div>');
  msgs.scrollTop = msgs.scrollHeight;
  try {
    var data = await api('/api/ai/analyze', {method:'POST', body:{exhibition_code: state.currentExhibition, query: text, history: state.seaHistory}});
    state.seaHistory.push({role:'user', content:text}, {role:'assistant', content:data.analysis});
    msgs.insertAdjacentHTML('beforeend', '<div class="sea-msg sea"><div class="sea-msg-bubble">' + renderMarkdown(data.analysis) + '</div></div>');
    msgs.scrollTop = msgs.scrollHeight;
  } catch(e) {
    msgs.insertAdjacentHTML('beforeend', '<div class="sea-msg sea"><div class="sea-msg-bubble">Sorry, something went wrong: ' + esc(e.message) + '</div></div>');
  }
}

// ============================================================================
// BOOT
// ============================================================================
async function boot(){
  var ok = await checkSession();
  if (!ok) return;
  await loadExhibitions();
  await refreshAll();
  navigate('page-home');
}
boot();
