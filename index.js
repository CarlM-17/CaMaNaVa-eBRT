const express = require('express');
const { google } = require('googleapis');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json());

// ─── Google Sheets Config ───────────────────────────────────────────────────
const SHEET_ID = process.env.SHEET_ID || '1PhRyLx2viByS1J-dOWWPwSmAZwNwzzaM6ATmbpQIR8w';
const SHEET_NAME = 'DailySales';
const DATA_RANGE = `${SHEET_NAME}!A23:R`;

function getAuthClient() {
  let credentials;
  if (process.env.GOOGLE_SERVICE_ACCOUNT_JSON) {
    credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
  } else {
    throw new Error('GOOGLE_SERVICE_ACCOUNT_JSON env variable is not set');
  }

  return new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
  });
}

// ─── Parse helpers ──────────────────────────────────────────────────────────
function parseNum(s) {
  if (!s || String(s).trim() === '') return 0;
  return parseFloat(String(s).replace(/,/g, '').replace('%', '')) || 0;
}

function parseDate(s) {
  if (!s) return null;
  const parts = String(s).trim().split('/');
  if (parts.length === 3) {
    const [m, d, y] = parts;
    return `${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`;
  }
  return null;
}

// ─── API: GET /api/sales ─────────────────────────────────────────────────────
app.get('/api/sales', async (req, res) => {
  try {
    const auth = getAuthClient();
    const sheets = google.sheets({ version: 'v4', auth });

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: DATA_RANGE,
    });

    const rows = response.data.values || [];

    let data = rows
      .map((r) => ({
        date: parseDate(r[0]),
        month: (r[1] || '').trim(),
        day: (r[2] || '').trim(),
        holiday: (r[4] || '').trim(),
        area: (r[5] || '').trim(),
        storeId: (r[6] || '').trim(),
        storeName: (r[7] || '').trim(),
        sales: parseNum(r[8]),
        salesLY: parseNum(r[9]),
        diffVal: parseNum(r[10]),
        diffPct: parseNum(r[11]),
        trx: parseNum(r[12]),
        trxLY: parseNum(r[13]),
        basketSize: parseNum(r[14]),
        basketSizeLY: parseNum(r[15]),
        dateRemark: (r[16] || '').trim(),
        justification: (r[17] || '').trim(),
      }))
      .filter((r) => r.date && r.storeName);

    const { date, area, store } = req.query;
    if (date) data = data.filter((r) => r.date === date);
    if (area && area !== 'ALL') data = data.filter((r) => r.area === area);
    if (store && store !== 'ALL') data = data.filter((r) => r.storeName === store);

    const storeMap = {};
    data.forEach((r) => {
      const key = `${r.storeId}_${r.storeName}`;
      if (!storeMap[key]) {
        storeMap[key] = {
          storeId: r.storeId,
          storeName: r.storeName,
          area: r.area,
          sales: 0,
          salesLY: 0,
          trx: 0,
          trxLY: 0,
          justifications: [],
        };
      }
      storeMap[key].sales += r.sales;
      storeMap[key].salesLY += r.salesLY;
      storeMap[key].trx += r.trx;
      storeMap[key].trxLY += r.trxLY;
      if (r.justification) storeMap[key].justifications.push(r.justification);
    });

    const result = Object.values(storeMap)
      .map((r) => {
        const diffVal = r.sales - r.salesLY;
        const diffPct = r.salesLY !== 0 ? (diffVal / r.salesLY) * 100 : 0;
        return {
          storeId: r.storeId,
          storeName: r.storeName,
          area: r.area,
          sales: r.sales,
          salesLY: r.salesLY,
          trx: r.trx,
          trxLY: r.trxLY,
          diffVal,
          diffPct: parseFloat(diffPct.toFixed(2)),
          justification: [...new Set(r.justifications)].join(' | '),
        };
      })
      .sort((a, b) => b.sales - a.sales);

    res.json({ success: true, count: result.length, rows: result });
  } catch (err) {
    console.error('Sheets API error:', err.message);
    res.status(500).json({ success: false, error: err.message });
  }
});

// ─── API: GET /api/filters ───────────────────────────────────────────────────
app.get('/api/filters', async (req, res) => {
  try {
    const auth = getAuthClient();
    const sheets = google.sheets({ version: 'v4', auth });

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: DATA_RANGE,
    });

    const rows = response.data.values || [];
    const { area: filterArea } = req.query;
    const areas = new Set();
    const stores = new Set();

    rows.forEach((r) => {
      const area = (r[5] || '').trim();
      const store = (r[7] || '').trim();
      if (area) areas.add(area);
      if (store) {
        if (!filterArea || filterArea === 'ALL' || area === filterArea) {
          stores.add(store);
        }
      }
    });

    res.json({
      success: true,
      areas: [...areas].filter(Boolean).sort(),
      stores: [...stores].filter(Boolean).sort(),
    });
  } catch (err) {
    console.error('Filters error:', err.message);
    res.status(500).json({ success: false, error: err.message });
  }
});

// ─── Health check ────────────────────────────────────────────────────────────
app.get('/health', (req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

// ─── Serve the frontend (embedded HTML) ──────────────────────────────────────
const HTML = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>CaMaNaVa eBRT — Daily Sales Report</title>
<link rel="preconnect" href="https://fonts.googleapis.com"/>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=DM+Sans:ital,wght@0,300;0,400;0,500;0,600;1,400&family=DM+Mono:wght@400;500&display=swap"/>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css"/>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js"></script>
<style>
*{box-sizing:border-box;margin:0;padding:0}
:root{
  --bg:#0f1117;--bg2:#161b27;--bg3:#1e2535;--bg4:#252d3e;
  --border:#2a3347;--border2:#3a4560;
  --text:#e8ecf4;--text2:#9ba8c0;--text3:#6b7894;
  --green:#22c55e;--green2:#166534;--green3:#052e16;
  --red:#ef4444;--red2:#991b1b;--red3:#2d0a0a;
  --blue:#3b82f6;--amber:#f59e0b;
  --accent:#4f8ef7;--accent2:#6ba3ff;
  --radius:10px;--radius2:14px;
  --purple:#a78bfa;
  font-family:'DM Sans',sans-serif;
}
body{background:var(--bg);color:var(--text);min-height:100vh;overflow-x:hidden}

/* ── Header ── */
.header{background:var(--bg2);border-bottom:1px solid var(--border);padding:14px 28px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:100}
.logo{display:flex;align-items:center;gap:12px}
.logo-icon{width:38px;height:38px;background:linear-gradient(135deg,#4f8ef7,#22c55e);border-radius:9px;display:flex;align-items:center;justify-content:center;font-size:16px;font-weight:700;color:#fff;flex-shrink:0}
.logo-text{font-size:17px;font-weight:600;letter-spacing:-0.3px;line-height:1.2}
.logo-sub{font-size:10.5px;color:var(--text3);letter-spacing:0.6px;text-transform:uppercase;margin-top:1px}
.header-right{display:flex;align-items:center;gap:10px}
.sync-btn{background:var(--bg3);border:1px solid var(--border2);color:var(--text2);padding:7px 14px;border-radius:8px;cursor:pointer;font-size:12px;font-family:inherit;display:flex;align-items:center;gap:6px;transition:all .2s}
.sync-btn:hover{background:var(--bg4);color:var(--text);border-color:var(--accent)}
.sync-btn.loading i{animation:spin 0.8s linear infinite}
@keyframes spin{to{transform:rotate(360deg)}}
.badge{padding:5px 12px;border-radius:20px;font-size:11px;font-weight:500;display:flex;align-items:center;gap:5px}
.badge.live{background:var(--green3);border:1px solid var(--green2);color:var(--green)}
.badge.error{background:var(--red3);border:1px solid var(--red2);color:var(--red)}
.badge.loading-badge{background:var(--bg3);border:1px solid var(--border2);color:var(--text3)}
.pulse{width:6px;height:6px;border-radius:50%;background:currentColor;animation:pulse 1.5s ease-in-out infinite}
@keyframes pulse{0%,100%{opacity:1}50%{opacity:0.3}}

/* ── Main layout ── */
.main{padding:22px 28px;max-width:1440px;margin:0 auto}

/* ── Controls ── */
.controls{display:flex;align-items:center;gap:12px;flex-wrap:wrap;margin-bottom:20px;background:var(--bg2);border:1px solid var(--border);border-radius:var(--radius2);padding:14px 18px}
.ctrl-group{display:flex;align-items:center;gap:8px}
.ctrl-label{font-size:11px;color:var(--text3);font-weight:500;text-transform:uppercase;letter-spacing:0.5px;white-space:nowrap}
select,input[type=date]{background:var(--bg3);border:1px solid var(--border2);color:var(--text);padding:7px 12px;border-radius:8px;font-size:12.5px;font-family:inherit;cursor:pointer;outline:none;transition:border-color .2s;-webkit-appearance:none}
select{padding-right:28px;background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='10' height='6' viewBox='0 0 10 6'%3E%3Cpath d='M0 0l5 6 5-6z' fill='%236b7894'/%3E%3C/svg%3E");background-repeat:no-repeat;background-position:right 10px center}
select:hover,input[type=date]:hover,select:focus,input[type=date]:focus{border-color:var(--accent)}
input[type=date]::-webkit-calendar-picker-indicator{filter:invert(0.5) sepia(1) saturate(5) hue-rotate(175deg);cursor:pointer}
.divider{width:1px;height:24px;background:var(--border);margin:0 2px}
.records-count{margin-left:auto;font-size:11.5px;color:var(--text3)}
.records-count span{color:var(--accent);font-weight:600}

/* ── KPI cards ── */
.kpi-grid{display:grid;grid-template-columns:repeat(5,1fr);gap:12px;margin-bottom:20px}
@media(max-width:1100px){.kpi-grid{grid-template-columns:repeat(3,1fr)}}
@media(max-width:700px){.kpi-grid{grid-template-columns:repeat(2,1fr)}}
.kpi{background:var(--bg2);border:1px solid var(--border);border-radius:var(--radius2);padding:16px 18px;transition:border-color .2s,transform .15s}
.kpi:hover{border-color:var(--border2);transform:translateY(-1px)}
.kpi-label{font-size:11px;color:var(--text3);text-transform:uppercase;letter-spacing:0.5px;margin-bottom:8px;display:flex;align-items:center;gap:6px}
.kpi-icon{width:22px;height:22px;border-radius:6px;display:flex;align-items:center;justify-content:center;font-size:10px;flex-shrink:0}
.kpi-value{font-size:21px;font-weight:600;letter-spacing:-0.5px;font-family:'DM Mono',monospace;line-height:1.1}
.kpi-sub{font-size:11px;margin-top:5px;display:flex;align-items:center;gap:4px;color:var(--text3)}
.up{color:var(--green)!important}
.down{color:var(--red)!important}

/* ── Charts ── */
.charts-grid{display:grid;grid-template-columns:1.65fr 1fr;gap:16px;margin-bottom:20px}
@media(max-width:960px){.charts-grid{grid-template-columns:1fr}}
.chart-card{background:var(--bg2);border:1px solid var(--border);border-radius:var(--radius2);padding:18px 20px}
.chart-title{font-size:12.5px;font-weight:500;color:var(--text2);margin-bottom:14px;display:flex;align-items:center;gap:8px}
.chart-title i{color:var(--accent);font-size:13px}
.chart-legend{display:flex;flex-wrap:wrap;gap:12px;margin-bottom:12px}
.legend-item{display:flex;align-items:center;gap:5px;font-size:11px;color:var(--text3)}
.legend-dot{width:9px;height:9px;border-radius:2px;flex-shrink:0}

/* ── Table ── */
.table-card{background:var(--bg2);border:1px solid var(--border);border-radius:var(--radius2);overflow:hidden}
.table-header{padding:14px 20px;border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px}
.table-title{font-size:12.5px;font-weight:500;color:var(--text2);display:flex;align-items:center;gap:8px}
.table-title i{color:var(--accent)}
.table-date{font-size:11.5px;color:var(--text3)}
.table-wrap{overflow-x:auto}
table{width:100%;border-collapse:collapse;font-size:12.5px}
th{background:var(--bg3);padding:10px 14px;text-align:left;font-size:10.5px;font-weight:500;color:var(--text3);text-transform:uppercase;letter-spacing:0.5px;white-space:nowrap;border-bottom:1px solid var(--border);position:sticky;top:0}
td{padding:10px 14px;border-bottom:1px solid var(--border);vertical-align:middle;white-space:nowrap}
tr:last-child td{border-bottom:none}
tbody tr:hover td{background:rgba(79,142,247,0.04)}
.store-name{font-weight:500;color:var(--text);font-size:13px}
.store-id{font-size:10.5px;color:var(--text3);font-family:'DM Mono',monospace;margin-top:1px}
.num{font-family:'DM Mono',monospace;font-size:12px}
.area-tag{display:inline-flex;align-items:center;gap:5px;font-size:11px;color:var(--text2);white-space:nowrap}
.area-dot{width:7px;height:7px;border-radius:50%;flex-shrink:0}
.pill{display:inline-flex;align-items:center;gap:3px;padding:3px 9px;border-radius:20px;font-size:11px;font-weight:500;font-family:'DM Mono',monospace}
.pill.up{background:var(--green3);border:1px solid var(--green2);color:var(--green)}
.pill.down{background:var(--red3);border:1px solid var(--red2);color:var(--red)}
.pill.flat{background:var(--bg4);border:1px solid var(--border2);color:var(--text3)}
.just-cell{max-width:220px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;color:var(--text3);font-size:11.5px;cursor:help}
.summary-row td{background:var(--bg3)!important;font-weight:600;border-top:2px solid var(--border2)}

/* ── Empty / loading states ── */
.empty-cell{text-align:center;padding:56px 20px;color:var(--text3)}
.empty-cell i{font-size:32px;margin-bottom:12px;display:block;opacity:0.4}
.empty-cell p{font-size:13px;margin-bottom:6px;color:var(--text2)}
.empty-cell small{font-size:11px;color:var(--text3)}

/* ── Loading overlay ── */
.loading-overlay{position:fixed;inset:0;background:rgba(10,12,18,0.85);display:flex;align-items:center;justify-content:center;z-index:999;backdrop-filter:blur(6px)}
.loading-box{text-align:center}
.spinner{width:42px;height:42px;border:3px solid var(--border2);border-top-color:var(--accent);border-radius:50%;animation:spin 0.75s linear infinite;margin:0 auto 18px}
.loading-title{font-size:15px;font-weight:500;color:var(--text);margin-bottom:4px}
.loading-msg{font-size:12px;color:var(--text3)}

/* ── Scrollbar ── */
::-webkit-scrollbar{width:6px;height:6px}
::-webkit-scrollbar-track{background:var(--bg2)}
::-webkit-scrollbar-thumb{background:var(--border2);border-radius:3px}
::-webkit-scrollbar-thumb:hover{background:var(--text3)}
</style>
</head>
<body>

<!-- Loading overlay -->
<div id="loadingOverlay" class="loading-overlay">
  <div class="loading-box">
    <div class="spinner"></div>
    <div class="loading-title">CaMaNaVa eBRT</div>
    <div class="loading-msg" id="loadingMsg">Connecting to Google Sheets...</div>
  </div>
</div>

<!-- Header -->
<header class="header">
  <div class="logo">
    <div class="logo-icon">C</div>
    <div>
      <div class="logo-text">CaMaNaVa eBRT</div>
      <div class="logo-sub">Daily Sales Report</div>
    </div>
  </div>
  <div class="header-right">
    <div id="statusBadge" class="badge loading-badge"><span class="pulse"></span> Loading</div>
    <button class="sync-btn" id="syncBtn" onclick="loadFilters(true)">
      <i class="fa fa-rotate" id="syncIcon"></i> Refresh
    </button>
  </div>
</header>

<!-- Main -->
<main class="main">

  <!-- Controls -->
  <div class="controls">
    <div class="ctrl-group">
      <span class="ctrl-label"><i class="fa fa-calendar-day"></i> Date</span>
      <input type="date" id="dateFilter" onchange="applyFilters()"/>
    </div>
    <div class="divider"></div>
    <div class="ctrl-group">
      <span class="ctrl-label"><i class="fa fa-layer-group"></i> Area</span>
      <select id="areaFilter" onchange="onAreaChange()">
        <option value="ALL">All Areas</option>
      </select>
    </div>
    <div class="divider"></div>
    <div class="ctrl-group">
      <span class="ctrl-label"><i class="fa fa-store"></i> Store</span>
      <select id="storeFilter" onchange="applyFilters()">
        <option value="ALL">All Stores</option>
      </select>
    </div>
    <div class="records-count" id="recordsCount">—</div>
  </div>

  <!-- KPI Cards -->
  <div class="kpi-grid">
    <div class="kpi">
      <div class="kpi-label"><div class="kpi-icon" style="background:rgba(79,142,247,0.15);color:var(--accent)"><i class="fa fa-peso-sign"></i></div>Total Sales</div>
      <div class="kpi-value" id="kpi-sales" style="color:var(--accent2)">—</div>
      <div class="kpi-sub">Current period</div>
    </div>
    <div class="kpi">
      <div class="kpi-label"><div class="kpi-icon" style="background:rgba(167,139,250,0.15);color:var(--purple)"><i class="fa fa-clock-rotate-left"></i></div>Sales LY</div>
      <div class="kpi-value" id="kpi-ly" style="color:var(--purple)">—</div>
      <div class="kpi-sub">Same period last year</div>
    </div>
    <div class="kpi">
      <div class="kpi-label"><div class="kpi-icon" id="kpi-pct-icon" style="background:rgba(34,197,94,0.12);color:var(--green)"><i class="fa fa-percent"></i></div>Diff %</div>
      <div class="kpi-value" id="kpi-pct">—</div>
      <div class="kpi-sub" id="kpi-pct-sub">vs last year</div>
    </div>
    <div class="kpi">
      <div class="kpi-label"><div class="kpi-icon" id="kpi-diff-icon" style="background:rgba(34,197,94,0.12);color:var(--green)"><i class="fa fa-arrow-trend-up"></i></div>Diff Amount</div>
      <div class="kpi-value" id="kpi-diff">—</div>
      <div class="kpi-sub" id="kpi-diff-sub">variance</div>
    </div>
    <div class="kpi">
      <div class="kpi-label"><div class="kpi-icon" style="background:rgba(245,158,11,0.12);color:var(--amber)"><i class="fa fa-store"></i></div>Active Stores</div>
      <div class="kpi-value" id="kpi-stores" style="color:var(--amber)">—</div>
      <div class="kpi-sub">with sales data</div>
    </div>
  </div>

  <!-- Charts -->
  <div class="charts-grid">
    <div class="chart-card">
      <div class="chart-title"><i class="fa fa-chart-bar"></i> Sales by Store — Current vs Last Year</div>
      <div class="chart-legend" id="barLegend">
        <span class="legend-item"><span class="legend-dot" style="background:#4f8ef7"></span>Sales Current</span>
        <span class="legend-item"><span class="legend-dot" style="background:#2a3347;border:1px solid #3a4560"></span>Sales LY</span>
      </div>
      <div style="position:relative;height:260px">
        <canvas id="barChart" role="img" aria-label="Bar chart: Sales by store current vs last year">Sales by store comparison chart</canvas>
      </div>
    </div>
    <div class="chart-card">
      <div class="chart-title"><i class="fa fa-chart-pie"></i> Sales Share by Area</div>
      <div style="position:relative;height:296px">
        <canvas id="pieChart" role="img" aria-label="Doughnut chart: Sales distribution by area">Area sales distribution</canvas>
      </div>
    </div>
  </div>

  <!-- Table -->
  <div class="table-card">
    <div class="table-header">
      <div class="table-title"><i class="fa fa-table-list"></i> Store Sales Detail</div>
      <div class="table-date" id="tableDate">—</div>
    </div>
    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th>Store</th>
            <th>Area</th>
            <th style="text-align:right">Sales</th>
            <th style="text-align:right">Sales LY</th>
            <th style="text-align:center">Diff %</th>
            <th style="text-align:right">Diff Amount</th>
            <th>Justification</th>
          </tr>
        </thead>
        <tbody id="tableBody">
          <tr><td colspan="7" class="empty-cell"><i class="fa fa-spinner fa-spin"></i><p>Loading data...</p></td></tr>
        </tbody>
      </table>
    </div>
  </div>

</main>

<script>
// ─── Config ────────────────────────────────────────────────────────────────
const AREA_COLORS = {
  'Valenzuela':      '#4f8ef7',
  'South Caloocan':  '#22c55e',
  'Malabon-Navotas': '#f59e0b',
  'North Caloocan':  '#a78bfa'
};
const DEFAULT_COLOR = '#9ba8c0';

let allFilters = { areas: [], stores: [] };
let barChartInst = null;
let pieChartInst = null;

// ─── Helpers ───────────────────────────────────────────────────────────────
function today() {
  return new Date().toLocaleDateString('en-CA');
}

function fmt(n) {
  if (n === null || n === undefined || isNaN(n)) return '—';
  const abs = Math.abs(n);
  const sign = n < 0 ? '-' : '';
  if (abs >= 1000000) return sign + '₱' + (abs / 1000000).toFixed(2) + 'M';
  if (abs >= 1000)    return sign + '₱' + (abs / 1000).toFixed(1) + 'K';
  return sign + '₱' + abs.toFixed(2);
}

function fmtFull(n) {
  if (!n && n !== 0) return '—';
  return '₱' + Math.abs(n).toLocaleString('en-PH', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// ─── Load filters (areas + stores) ────────────────────────────────────────
async function loadFilters(refresh = false) {
  setSyncing(true);
  try {
    const res = await fetch('/api/filters');
    const json = await res.json();
    if (!json.success) throw new Error(json.error);

    allFilters = json;

    const areaSel = document.getElementById('areaFilter');
    areaSel.innerHTML = '<option value="ALL">All Areas</option>';
    json.areas.forEach(a => {
      const o = document.createElement('option');
      o.value = o.textContent = a;
      areaSel.appendChild(o);
    });

    populateStores(json.stores);
    document.getElementById('dateFilter').value = today();
    await applyFilters();
  } catch(e) {
    setStatus('error', '● Error');
    console.error(e);
    showTableError('Failed to load. Check server and sheet permissions.');
  } finally {
    setSyncing(false);
    hideOverlay();
  }
}

function populateStores(stores) {
  const sel = document.getElementById('storeFilter');
  sel.innerHTML = '<option value="ALL">All Stores</option>';
  stores.forEach(s => {
    const o = document.createElement('option');
    o.value = o.textContent = s;
    sel.appendChild(o);
  });
}

async function onAreaChange() {
  const area = document.getElementById('areaFilter').value;
  if (area === 'ALL') {
    populateStores(allFilters.stores);
  } else {
    // Filter stores by selected area — use current data if available
    // Re-fetch filter stores for this area
    try {
      const res = await fetch(\`/api/filters?area=\${encodeURIComponent(area)}\`);
      const json = await res.json();
      if (json.success) populateStores(json.stores);
    } catch(e) { /* fallback: keep current */ }
  }
  document.getElementById('storeFilter').value = 'ALL';
  await applyFilters();
}

// ─── Apply filters + fetch data ────────────────────────────────────────────
async function applyFilters() {
  const date  = document.getElementById('dateFilter').value;
  const area  = document.getElementById('areaFilter').value;
  const store = document.getElementById('storeFilter').value;

  const params = new URLSearchParams();
  if (date)             params.set('date', date);
  if (area  !== 'ALL')  params.set('area', area);
  if (store !== 'ALL')  params.set('store', store);

  try {
    const res  = await fetch('/api/sales?' + params);
    const json = await res.json();
    if (!json.success) throw new Error(json.error);

    const rows = json.rows || [];
    document.getElementById('recordsCount').innerHTML =
      \`<span>\${rows.length}</span> store\${rows.length !== 1 ? 's' : ''} · <span>\${json.count}</span> row\${json.count !== 1 ? 's' : ''}\`;

    const dateLabel = date
      ? new Date(date + 'T00:00:00').toLocaleDateString('en-PH', { weekday:'long', year:'numeric', month:'long', day:'numeric' })
      : 'All dates';
    document.getElementById('tableDate').textContent = dateLabel;

    renderKPIs(rows);
    renderTable(rows);
    renderCharts(rows);
    setStatus('live', '● Live');
  } catch(e) {
    console.error(e);
    setStatus('error', '● Error');
    showTableError('Error loading sales data: ' + e.message);
  }
}

// ─── KPI Render ────────────────────────────────────────────────────────────
function renderKPIs(rows) {
  const totSales = rows.reduce((s, r) => s + (r.sales || 0), 0);
  const totLY    = rows.reduce((s, r) => s + (r.salesLY || 0), 0);
  const totDiff  = totSales - totLY;
  const totPct   = totLY !== 0 ? (totDiff / totLY * 100) : 0;
  const active   = rows.filter(r => r.sales > 0).length;

  document.getElementById('kpi-sales').textContent = fmt(totSales);
  document.getElementById('kpi-ly').textContent    = fmt(totLY);

  const pctEl = document.getElementById('kpi-pct');
  pctEl.textContent = (totPct >= 0 ? '+' : '') + totPct.toFixed(2) + '%';
  pctEl.style.color = totPct >= 0 ? 'var(--green)' : 'var(--red)';
  document.getElementById('kpi-pct-sub').innerHTML =
    \`<i class="fa fa-arrow-\${totPct >= 0 ? 'up' : 'down'}" style="color:\${totPct>=0?'var(--green)':'var(--red)'}"></i> vs last year\`;

  const diffEl = document.getElementById('kpi-diff');
  diffEl.textContent = (totDiff >= 0 ? '+' : '') + fmt(totDiff);
  diffEl.style.color = totDiff >= 0 ? 'var(--green)' : 'var(--red)';
  document.getElementById('kpi-diff-sub').innerHTML =
    \`<i class="fa fa-\${totDiff >= 0 ? 'circle-check' : 'circle-xmark'}" style="color:\${totDiff>=0?'var(--green)':'var(--red)'}"></i> variance\`;

  document.getElementById('kpi-stores').textContent = active;
}

// ─── Table Render ──────────────────────────────────────────────────────────
function renderTable(rows) {
  const tbody = document.getElementById('tableBody');
  if (!rows.length) {
    tbody.innerHTML = \`<tr><td colspan="7" class="empty-cell"><i class="fa fa-magnifying-glass"></i><p>No data for selected filters</p><small>Try a different date or area</small></td></tr>\`;
    return;
  }

  const totSales = rows.reduce((s, r) => s + r.sales, 0);
  const totLY    = rows.reduce((s, r) => s + r.salesLY, 0);
  const totDiff  = totSales - totLY;
  const totPct   = totLY ? (totDiff / totLY * 100) : 0;

  let html = rows.map(r => {
    const pctCls  = r.diffPct > 0.05 ? 'up' : r.diffPct < -0.05 ? 'down' : 'flat';
    const color   = AREA_COLORS[r.area] || DEFAULT_COLOR;
    const diffColor = r.diffVal >= 0 ? 'var(--green)' : 'var(--red)';
    const just    = (r.justification || '—').replace(/"/g, '&quot;');
    const arrow   = r.diffPct > 0.05 ? '▲' : r.diffPct < -0.05 ? '▼' : '—';
    const pctStr  = r.diffPct !== 0 ? \`\${arrow} \${Math.abs(r.diffPct).toFixed(2)}%\` : '—';

    return \`<tr>
      <td>
        <div class="store-name">\${r.storeName || '—'}</div>
        <div class="store-id">#\${r.storeId}</div>
      </td>
      <td>
        <span class="area-tag">
          <span class="area-dot" style="background:\${color}"></span>
          \${r.area || '—'}
        </span>
      </td>
      <td style="text-align:right"><span class="num">\${fmtFull(r.sales)}</span></td>
      <td style="text-align:right"><span class="num" style="color:var(--text3)">\${fmtFull(r.salesLY)}</span></td>
      <td style="text-align:center"><span class="pill \${pctCls}">\${pctStr}</span></td>
      <td style="text-align:right"><span class="num" style="color:\${diffColor}">\${r.diffVal >= 0 ? '+' : ''}\${fmtFull(r.diffVal)}</span></td>
      <td><div class="just-cell" title="\${just}">\${just}</div></td>
    </tr>\`;
  }).join('');

  // Summary row
  const totPctCls = totPct > 0.05 ? 'up' : totPct < -0.05 ? 'down' : 'flat';
  const totArrow  = totPct > 0.05 ? '▲' : totPct < -0.05 ? '▼' : '—';
  html += \`<tr class="summary-row">
    <td colspan="2" style="font-size:12px;letter-spacing:0.5px;text-transform:uppercase">TOTAL · \${rows.length} Stores</td>
    <td style="text-align:right"><span class="num">\${fmtFull(totSales)}</span></td>
    <td style="text-align:right"><span class="num" style="color:var(--text3)">\${fmtFull(totLY)}</span></td>
    <td style="text-align:center"><span class="pill \${totPctCls}">\${totArrow} \${Math.abs(totPct).toFixed(2)}%</span></td>
    <td style="text-align:right"><span class="num" style="color:\${totDiff>=0?'var(--green)':'var(--red)'}">\${totDiff>=0?'+':''}\${fmtFull(totDiff)}</span></td>
    <td></td>
  </tr>\`;

  tbody.innerHTML = html;
}

// ─── Charts Render ─────────────────────────────────────────────────────────
function renderCharts(rows) {
  const top = rows.slice(0, 14);

  // Bar chart
  const labels   = top.map(r => r.storeName);
  const salesArr = top.map(r => r.sales);
  const lyArr    = top.map(r => r.salesLY);
  const bgColors = top.map(r => AREA_COLORS[r.area] || DEFAULT_COLOR);

  if (barChartInst) barChartInst.destroy();
  barChartInst = new Chart(document.getElementById('barChart'), {
    type: 'bar',
    data: {
      labels,
      datasets: [
        { label: 'Sales', data: salesArr, backgroundColor: bgColors, borderRadius: 4, borderSkipped: false },
        { label: 'Sales LY', data: lyArr, backgroundColor: '#22293a', borderColor: '#3a4560', borderWidth: 1, borderRadius: 4, borderSkipped: false }
      ]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          backgroundColor: '#1e2535', borderColor: '#3a4560', borderWidth: 1,
          titleFont: { family: "'DM Sans', sans-serif", size: 12 },
          bodyFont:  { family: "'DM Mono', monospace", size: 11 },
          callbacks: {
            label: ctx => \` \${ctx.dataset.label}: ₱\${ctx.raw.toLocaleString('en-PH', { minimumFractionDigits: 2 })}\`
          }
        }
      },
      scales: {
        x: {
          grid: { color: '#1e2535', drawBorder: false },
          ticks: { color: '#6b7894', font: { size: 9.5 }, maxRotation: 40 }
        },
        y: {
          grid: { color: '#1a2030' },
          ticks: {
            color: '#6b7894', font: { size: 10 },
            callback: v => v >= 1e6 ? '₱' + (v/1e6).toFixed(1)+'M' : v >= 1e3 ? '₱' + (v/1e3).toFixed(0)+'K' : v
          }
        }
      }
    }
  });

  // Doughnut chart
  const areaMap = {};
  rows.forEach(r => { areaMap[r.area] = (areaMap[r.area] || 0) + r.sales; });
  const areaLabels = Object.keys(areaMap);
  const areaVals   = Object.values(areaMap);
  const pieColors  = areaLabels.map(a => AREA_COLORS[a] || DEFAULT_COLOR);

  if (pieChartInst) pieChartInst.destroy();
  pieChartInst = new Chart(document.getElementById('pieChart'), {
    type: 'doughnut',
    data: {
      labels: areaLabels,
      datasets: [{ data: areaVals, backgroundColor: pieColors, borderColor: '#161b27', borderWidth: 3, hoverOffset: 8 }]
    },
    options: {
      responsive: true, maintainAspectRatio: false, cutout: '60%',
      plugins: {
        legend: {
          position: 'bottom',
          labels: {
            color: '#9ba8c0', font: { size: 11, family: "'DM Sans'" }, padding: 14, boxWidth: 10, boxHeight: 10,
            generateLabels: chart => chart.data.labels.map((label, i) => ({
              text: \`\${label}  \${fmt(areaVals[i])}\`,
              fillStyle: pieColors[i], strokeStyle: 'transparent',
              lineWidth: 0, pointStyle: 'rect'
            }))
          }
        },
        tooltip: {
          backgroundColor: '#1e2535', borderColor: '#3a4560', borderWidth: 1,
          callbacks: {
            label: ctx => {
              const total = ctx.dataset.data.reduce((a, b) => a + b, 0);
              const pct   = total ? ((ctx.raw / total) * 100).toFixed(1) : 0;
              return \` ₱\${ctx.raw.toLocaleString('en-PH', { minimumFractionDigits: 2 })}  (\${pct}%)\`;
            }
          }
        }
      }
    }
  });
}

// ─── UI helpers ────────────────────────────────────────────────────────────
function setStatus(type, text) {
  const b = document.getElementById('statusBadge');
  b.className = 'badge ' + (type === 'live' ? 'live' : type === 'error' ? 'error' : 'loading-badge');
  b.innerHTML = type === 'live'
    ? \`<span class="pulse"></span> Live\`
    : type === 'error'
    ? \`<i class="fa fa-circle-exclamation"></i> Error\`
    : \`<span class="pulse"></span> Loading\`;
}

function setSyncing(on) {
  const btn = document.getElementById('syncBtn');
  const ico = document.getElementById('syncIcon');
  btn.disabled = on;
  if (on) ico.style.animation = 'spin 0.8s linear infinite';
  else    ico.style.animation = '';
}

function hideOverlay() {
  document.getElementById('loadingOverlay').style.display = 'none';
}

function showTableError(msg) {
  document.getElementById('tableBody').innerHTML =
    \`<tr><td colspan="7" class="empty-cell"><i class="fa fa-triangle-exclamation" style="color:var(--amber)"></i><p>\${msg}</p></td></tr>\`;
}

// ─── Boot ──────────────────────────────────────────────────────────────────
loadFilters();
</script>
</body>
</html>
`;

app.get('/', (req, res) => {
  res.setHeader('Content-Type', 'text/html; charset=utf-8');
  res.send(HTML);
});

app.get('*', (req, res) => {
  res.setHeader('Content-Type', 'text/html; charset=utf-8');
  res.send(HTML);
});

app.listen(PORT, () => {
  console.log(`CaMaNaVa eBRT running on port ${PORT}`);
});
