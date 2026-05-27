const express = require('express');
const { google } = require('googleapis');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json());

const SHEET_ID = process.env.SHEET_ID || '1PhRyLx2viByS1J-dOWWPwSmAZwNwzzaM6ATmbpQIR8w';
const SHEET_NAME = 'DailySales';
const DATA_RANGE = `${SHEET_NAME}!A23:R`;
const STORE_LIST_RANGE = 'ListOfStores!A:E';

function getAuthClient() {
  if (!process.env.GOOGLE_SERVICE_ACCOUNT_JSON) {
    throw new Error('GOOGLE_SERVICE_ACCOUNT_JSON env variable is not set');
  }
  const credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
  return new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
  });
}

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

function monthKey(dateStr) {
  // dateStr is YYYY-MM-DD
  if (!dateStr) return null;
  return dateStr.substring(0, 7); // YYYY-MM
}

function monthLabel(key) {
  // key = YYYY-MM
  if (!key) return '';
  const [y, m] = key.split('-');
  const monthNames = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  return `${monthNames[parseInt(m)-1]} ${y}`;
}

// ─── Caches ─────────────────────────────────────────────────────────────────
let storeListCache = null;
let storeListCacheTime = 0;
let salesCache = null;
let salesCacheTime = 0;
const CACHE_TTL = 5 * 60 * 1000;

async function getMasterStoreList(sheets) {
  const now = Date.now();
  if (storeListCache && (now - storeListCacheTime) < CACHE_TTL) return storeListCache;
  const response = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: STORE_LIST_RANGE });
  const rows = response.data.values || [];
  storeListCache = rows
    .map(r => ({
      region:    (r[0] || '').trim(),
      area:      (r[1] || '').trim(),
      storeId:   (r[2] || '').trim(),
      storeName: (r[3] || '').trim(),
      remarks:   (r[4] || '').trim(),
    }))
    .filter(s => s.storeId && /^\d+$/.test(s.storeId) && s.storeName);
  storeListCacheTime = now;
  return storeListCache;
}

async function getSalesData(sheets) {
  const now = Date.now();
  if (salesCache && (now - salesCacheTime) < CACHE_TTL) return salesCache;
  const response = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: DATA_RANGE });
  const rows = response.data.values || [];
  salesCache = rows
    .map((r) => ({
      date: parseDate(r[0]),
      month: (r[1] || '').trim(),
      day: (r[2] || '').trim(),
      dayYA: (r[3] || '').trim(),
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
  salesCacheTime = now;
  return salesCache;
}

// ─── API: GET /api/sales ─────────────────────────────────────────────────────
app.get('/api/sales', async (req, res) => {
  try {
    const auth = getAuthClient();
    const sheets = google.sheets({ version: 'v4', auth });

    const [data, masterStores] = await Promise.all([getSalesData(sheets), getMasterStoreList(sheets)]);

    const { date, area, store } = req.query;
    let filtered = data;
    if (date) filtered = filtered.filter((r) => r.date === date);
    if (area && area !== 'ALL') filtered = filtered.filter((r) => r.area === area);
    if (store && store !== 'ALL') filtered = filtered.filter((r) => r.storeName === store);

    const storeMap = {};
    filtered.forEach((r) => {
      const key = `${r.storeId}_${r.storeName}`;
      if (!storeMap[key]) {
        storeMap[key] = { storeId: r.storeId, storeName: r.storeName, area: r.area, sales: 0, salesLY: 0, trx: 0, trxLY: 0, justifications: [] };
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
          storeId: r.storeId, storeName: r.storeName, area: r.area,
          sales: r.sales, salesLY: r.salesLY, trx: r.trx, trxLY: r.trxLY,
          diffVal, diffPct: parseFloat(diffPct.toFixed(2)),
          justification: [...new Set(r.justifications)].join(' | '),
        };
      })
      .sort((a, b) => b.sales - a.sales);

    const reportedStoreIds = new Set(result.map(r => r.storeId));
    let missing = [];
    if (date) {
      let pool = masterStores;
      if (area && area !== 'ALL') pool = pool.filter(s => s.area === area);
      if (store && store !== 'ALL') pool = pool.filter(s => s.storeName === store);
      missing = pool
        .filter(s => !reportedStoreIds.has(s.storeId))
        .map(s => ({ storeId: s.storeId, storeName: s.storeName, area: s.area, region: s.region, remarks: s.remarks }))
        .sort((a, b) => a.storeName.localeCompare(b.storeName));
    }

    res.json({ success: true, count: result.length, rows: result, missing, totalMasterStores: masterStores.length });
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
    const masterStores = await getMasterStoreList(sheets);
    const { area: filterArea } = req.query;
    const areas = new Set();
    const stores = new Set();
    masterStores.forEach((s) => {
      if (s.area) areas.add(s.area);
      if (s.storeName) {
        if (!filterArea || filterArea === 'ALL' || s.area === filterArea) stores.add(s.storeName);
      }
    });
    res.json({ success: true, areas: [...areas].filter(Boolean).sort(), stores: [...stores].filter(Boolean).sort() });
  } catch (err) {
    console.error('Filters error:', err.message);
    res.status(500).json({ success: false, error: err.message });
  }
});

// ─── API: GET /api/months ───────────────────────────────────────────────────
app.get('/api/months', async (req, res) => {
  try {
    const auth = getAuthClient();
    const sheets = google.sheets({ version: 'v4', auth });
    const data = await getSalesData(sheets);
    const monthSet = new Set();
    data.forEach(r => { const k = monthKey(r.date); if (k) monthSet.add(k); });
    const months = [...monthSet].sort().map(k => ({ value: k, label: monthLabel(k) }));
    res.json({ success: true, months });
  } catch (err) {
    console.error('Months error:', err.message);
    res.status(500).json({ success: false, error: err.message });
  }
});

// ─── API: GET /api/monthly ──────────────────────────────────────────────────
app.get('/api/monthly', async (req, res) => {
  try {
    const auth = getAuthClient();
    const sheets = google.sheets({ version: 'v4', auth });
    const data = await getSalesData(sheets);

    const { month, area, store, sign } = req.query;
    if (!month) return res.json({ success: true, summary: [], detail: [], trend: [] });

    let filtered = data.filter(r => monthKey(r.date) === month);
    if (area && area !== 'ALL') filtered = filtered.filter(r => r.area === area);
    if (store && store !== 'ALL') filtered = filtered.filter(r => r.storeName === store);

    // ── Summary by store ──
    const storeMap = {};
    filtered.forEach(r => {
      const key = `${r.storeId}_${r.storeName}`;
      if (!storeMap[key]) {
        storeMap[key] = { storeId: r.storeId, storeName: r.storeName, area: r.area, sales: 0, salesLY: 0 };
      }
      storeMap[key].sales += r.sales;
      storeMap[key].salesLY += r.salesLY;
    });
    const summary = Object.values(storeMap).map(r => {
      const diffVal = r.sales - r.salesLY;
      const diffPct = r.salesLY !== 0 ? (diffVal / r.salesLY) * 100 : 0;
      return { ...r, diffVal, diffPct: parseFloat(diffPct.toFixed(2)) };
    }).sort((a, b) => b.sales - a.sales);

    // ── Detail (one row per record) ──
    // Apply sign filter for detail only
    let detailRecords = filtered;
    if (sign === 'POS') detailRecords = detailRecords.filter(r => (r.sales - r.salesLY) > 0);
    if (sign === 'NEG') detailRecords = detailRecords.filter(r => (r.sales - r.salesLY) < 0);

    const detail = detailRecords.map(r => {
      const diffVal = r.sales - r.salesLY;
      const diffPct = r.salesLY !== 0 ? (diffVal / r.salesLY) * 100 : 0;
      return {
        date: r.date,
        day: r.day,
        dayYA: r.dayYA,
        storeId: r.storeId,
        storeName: r.storeName,
        area: r.area,
        sales: r.sales,
        salesLY: r.salesLY,
        diffVal,
        diffPct: parseFloat(diffPct.toFixed(2)),
        justification: r.justification || '',
      };
    }).sort((a, b) => {
      // Sort by date asc, then store asc
      if (a.date !== b.date) return a.date.localeCompare(b.date);
      return (a.storeName || '').localeCompare(b.storeName || '');
    });

    // ── Daily trend (no sign filter applied here, total only) ──
    const trendMap = {};
    filtered.forEach(r => {
      if (!trendMap[r.date]) trendMap[r.date] = { date: r.date, sales: 0, salesLY: 0 };
      trendMap[r.date].sales += r.sales;
      trendMap[r.date].salesLY += r.salesLY;
    });
    const trend = Object.values(trendMap).sort((a, b) => a.date.localeCompare(b.date));

    res.json({ success: true, summary, detail, trend });
  } catch (err) {
    console.error('Monthly error:', err.message);
    res.status(500).json({ success: false, error: err.message });
  }
});

app.get('/health', (req, res) => res.json({ status: 'ok', timestamp: new Date().toISOString() }));

const HTML = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>CaMaNaVa eBRT — Daily Sales Report</title>
<link rel="preconnect" href="https://fonts.googleapis.com"/>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;600;700&family=Inter:wght@300;400;500;600&family=JetBrains+Mono:wght@400;500;600&display=swap"/>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css"/>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
*{box-sizing:border-box;margin:0;padding:0}
:root{
  /* Refined dark palette with depth */
  --bg-base:#0a0e1a;
  --bg-deep:#070a14;
  --bg-glass:rgba(20,26,42,0.65);
  --bg-glass2:rgba(28,35,58,0.7);
  --bg-card:rgba(24,30,48,0.55);
  --bg-elevated:rgba(30,38,60,0.85);

  --border-glow:rgba(99,148,255,0.18);
  --border-soft:rgba(148,163,200,0.12);
  --border-strong:rgba(148,163,200,0.22);

  --text-1:#f0f3fb;
  --text-2:#a8b3d1;
  --text-3:#6b7693;
  --text-dim:#4a536b;

  /* Vibrant accents */
  --indigo:#6366f1;
  --indigo2:#818cf8;
  --indigo-glow:rgba(99,102,241,0.35);

  --cyan:#06b6d4;
  --cyan2:#22d3ee;
  --cyan-glow:rgba(34,211,238,0.35);

  --emerald:#10b981;
  --emerald2:#34d399;
  --emerald-glow:rgba(16,185,129,0.35);

  --rose:#f43f5e;
  --rose2:#fb7185;
  --rose-glow:rgba(244,63,94,0.3);

  --amber:#f59e0b;
  --amber2:#fbbf24;
  --amber-glow:rgba(245,158,11,0.3);

  --violet:#a855f7;
  --violet2:#c084fc;
  --violet-glow:rgba(168,85,247,0.3);

  --pink:#ec4899;
  --pink2:#f472b6;
}

html,body{background:var(--bg-base);color:var(--text-1);min-height:100vh;overflow-x:hidden;font-family:'Inter',sans-serif;font-weight:400;letter-spacing:-0.005em}

/* Animated gradient background */
body::before{
  content:'';position:fixed;inset:0;z-index:-2;
  background:
    radial-gradient(at 15% 10%, rgba(99,102,241,0.18) 0px, transparent 50%),
    radial-gradient(at 85% 15%, rgba(168,85,247,0.15) 0px, transparent 50%),
    radial-gradient(at 80% 85%, rgba(6,182,212,0.13) 0px, transparent 50%),
    radial-gradient(at 20% 90%, rgba(236,72,153,0.10) 0px, transparent 50%),
    var(--bg-base);
  animation:bgShift 24s ease-in-out infinite alternate;
}
@keyframes bgShift{
  0%{filter:hue-rotate(0deg)}
  100%{filter:hue-rotate(20deg)}
}

/* Subtle grid overlay */
body::after{
  content:'';position:fixed;inset:0;z-index:-1;pointer-events:none;
  background-image:
    linear-gradient(rgba(99,148,255,0.025) 1px, transparent 1px),
    linear-gradient(90deg, rgba(99,148,255,0.025) 1px, transparent 1px);
  background-size:48px 48px;
  mask-image:radial-gradient(ellipse at center, black 30%, transparent 80%);
}

/* ── Header ── */
.header{
  background:linear-gradient(180deg, rgba(20,26,42,0.85) 0%, rgba(20,26,42,0.65) 100%);
  backdrop-filter:blur(20px) saturate(180%);
  -webkit-backdrop-filter:blur(20px) saturate(180%);
  border-bottom:1px solid var(--border-soft);
  padding:16px 32px;
  display:flex;align-items:center;justify-content:space-between;
  position:sticky;top:0;z-index:100;
}
.logo{display:flex;align-items:center;gap:14px}
.logo-icon{
  width:44px;height:44px;
  background:linear-gradient(135deg,#6366f1 0%,#06b6d4 50%,#10b981 100%);
  border-radius:12px;
  display:flex;align-items:center;justify-content:center;
  font-size:18px;font-weight:700;color:#fff;
  flex-shrink:0;
  font-family:'Space Grotesk',sans-serif;
  box-shadow:0 8px 24px -8px rgba(99,102,241,0.6), inset 0 1px 0 rgba(255,255,255,0.2);
  position:relative;
}
.logo-icon::after{
  content:'';position:absolute;inset:-2px;border-radius:14px;
  background:linear-gradient(135deg,#6366f1,#06b6d4,#10b981);
  z-index:-1;filter:blur(12px);opacity:0.4;
}
.logo-text{font-family:'Space Grotesk',sans-serif;font-size:19px;font-weight:600;letter-spacing:-0.5px;line-height:1.1;background:linear-gradient(135deg,#fff 0%,#a8b3d1 100%);-webkit-background-clip:text;background-clip:text;-webkit-text-fill-color:transparent}
.logo-sub{font-size:10px;color:var(--text-3);letter-spacing:1.5px;text-transform:uppercase;margin-top:3px;font-weight:500}

.header-right{display:flex;align-items:center;gap:12px}
.sync-btn{
  background:var(--bg-glass);backdrop-filter:blur(10px);
  border:1px solid var(--border-strong);
  color:var(--text-1);
  padding:9px 16px;border-radius:10px;cursor:pointer;
  font-size:12.5px;font-family:'Inter',sans-serif;font-weight:500;
  display:flex;align-items:center;gap:7px;transition:all .25s cubic-bezier(0.4,0,0.2,1);
  letter-spacing:0.01em;
}
.sync-btn:hover{
  background:var(--bg-elevated);border-color:var(--indigo);color:#fff;
  transform:translateY(-1px);box-shadow:0 6px 16px -4px var(--indigo-glow);
}
.sync-btn:active{transform:translateY(0)}
.sync-btn.loading i{animation:spin 0.8s linear infinite}
@keyframes spin{to{transform:rotate(360deg)}}

.badge{
  padding:6px 13px;border-radius:20px;font-size:11px;font-weight:600;
  display:flex;align-items:center;gap:6px;letter-spacing:0.03em;
  backdrop-filter:blur(10px);
}
.badge.live{background:rgba(16,185,129,0.12);border:1px solid rgba(16,185,129,0.35);color:var(--emerald2);box-shadow:0 0 20px -4px var(--emerald-glow)}
.badge.error{background:rgba(244,63,94,0.12);border:1px solid rgba(244,63,94,0.35);color:var(--rose2)}
.badge.loading-badge{background:var(--bg-glass);border:1px solid var(--border-strong);color:var(--text-2)}
.pulse{width:7px;height:7px;border-radius:50%;background:currentColor;animation:pulse 1.5s ease-in-out infinite;box-shadow:0 0 8px currentColor}
@keyframes pulse{0%,100%{opacity:1;transform:scale(1)}50%{opacity:0.4;transform:scale(0.85)}}

/* ── Main ── */
.main{padding:28px 32px 40px;max-width:1480px;margin:0 auto}

/* ── Controls ── */
.controls{
  display:flex;align-items:center;gap:14px;flex-wrap:wrap;
  margin-bottom:24px;
  background:var(--bg-glass);
  backdrop-filter:blur(20px) saturate(180%);
  -webkit-backdrop-filter:blur(20px) saturate(180%);
  border:1px solid var(--border-soft);
  border-radius:16px;padding:16px 22px;
  box-shadow:0 8px 32px -12px rgba(0,0,0,0.4);
}
.ctrl-group{display:flex;align-items:center;gap:10px}
.ctrl-label{
  font-size:10.5px;color:var(--text-3);
  font-weight:600;text-transform:uppercase;letter-spacing:0.1em;
  white-space:nowrap;display:flex;align-items:center;gap:5px;
}
.ctrl-label i{font-size:11px;color:var(--indigo2)}
select,input[type=date]{
  background:rgba(15,20,35,0.7);
  border:1px solid var(--border-strong);
  color:var(--text-1);
  padding:9px 14px;border-radius:10px;
  font-size:13px;font-family:'Inter',sans-serif;font-weight:500;
  cursor:pointer;outline:none;transition:all .2s;
  -webkit-appearance:none;letter-spacing:-0.01em;
}
select{padding-right:34px;background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='10' height='6' viewBox='0 0 10 6'%3E%3Cpath d='M0 0l5 6 5-6z' fill='%236366f1'/%3E%3C/svg%3E");background-repeat:no-repeat;background-position:right 12px center}
select:hover,input[type=date]:hover{border-color:var(--indigo);background:rgba(99,102,241,0.05)}
select:focus,input[type=date]:focus{border-color:var(--indigo);box-shadow:0 0 0 3px rgba(99,102,241,0.15)}
input[type=date]::-webkit-calendar-picker-indicator{filter:invert(0.6) sepia(1) saturate(5) hue-rotate(210deg);cursor:pointer;opacity:0.8}
.divider{width:1px;height:24px;background:linear-gradient(180deg, transparent, var(--border-strong), transparent);margin:0 4px}
.records-count{margin-left:auto;font-size:12px;color:var(--text-2);font-weight:500}
.records-count span{color:var(--indigo2);font-weight:600;font-family:'JetBrains Mono',monospace}

/* ── KPI Cards ── */
.kpi-grid{display:grid;grid-template-columns:repeat(5,1fr);gap:14px;margin-bottom:24px}
@media(max-width:1100px){.kpi-grid{grid-template-columns:repeat(3,1fr)}}
@media(max-width:700px){.kpi-grid{grid-template-columns:repeat(2,1fr)}}

.kpi{
  position:relative;
  background:var(--bg-glass);
  backdrop-filter:blur(20px) saturate(180%);
  -webkit-backdrop-filter:blur(20px) saturate(180%);
  border:1px solid var(--border-soft);
  border-radius:18px;padding:20px 22px;
  transition:all .3s cubic-bezier(0.4,0,0.2,1);
  overflow:hidden;
}
.kpi::before{
  content:'';position:absolute;top:0;left:0;right:0;height:1px;
  background:linear-gradient(90deg, transparent, var(--accent-color, var(--indigo)), transparent);
  opacity:0.6;
}
.kpi::after{
  content:'';position:absolute;top:-40%;right:-30%;width:200px;height:200px;
  background:radial-gradient(circle, var(--accent-color, var(--indigo)) 0%, transparent 70%);
  opacity:0.08;pointer-events:none;
}
.kpi:hover{transform:translateY(-3px);border-color:var(--border-strong);box-shadow:0 16px 32px -16px rgba(0,0,0,0.6)}
.kpi:hover::after{opacity:0.15}
.kpi.k-sales{--accent-color:#6366f1}
.kpi.k-ly{--accent-color:#a855f7}
.kpi.k-pct{--accent-color:#10b981}
.kpi.k-diff{--accent-color:#06b6d4}
.kpi.k-stores{--accent-color:#f59e0b}

.kpi-label{
  font-size:10.5px;color:var(--text-3);
  text-transform:uppercase;letter-spacing:0.12em;
  margin-bottom:14px;display:flex;align-items:center;gap:9px;
  font-weight:600;
}
.kpi-icon{
  width:30px;height:30px;border-radius:9px;
  display:flex;align-items:center;justify-content:center;
  font-size:13px;flex-shrink:0;
  background:linear-gradient(135deg, var(--accent-color, var(--indigo)), color-mix(in srgb, var(--accent-color, var(--indigo)) 60%, transparent));
  color:#fff;
  box-shadow:0 4px 12px -2px var(--accent-color, var(--indigo-glow)), inset 0 1px 0 rgba(255,255,255,0.2);
}
.kpi-value{
  font-family:'Space Grotesk',sans-serif;
  font-size:26px;font-weight:600;
  letter-spacing:-0.02em;line-height:1.1;
}
.kpi-value.gradient-text{
  background:linear-gradient(135deg, var(--accent-color, var(--indigo)) 0%, color-mix(in srgb, var(--accent-color) 70%, white) 100%);
  -webkit-background-clip:text;background-clip:text;-webkit-text-fill-color:transparent;
}
.kpi-sub{font-size:11.5px;margin-top:8px;display:flex;align-items:center;gap:5px;color:var(--text-3);font-weight:500}

/* ── Charts ── */
.charts-grid{display:grid;grid-template-columns:1.7fr 1fr;gap:18px;margin-bottom:24px}
@media(max-width:960px){.charts-grid{grid-template-columns:1fr}}
.chart-card{
  background:var(--bg-glass);
  backdrop-filter:blur(20px) saturate(180%);
  -webkit-backdrop-filter:blur(20px) saturate(180%);
  border:1px solid var(--border-soft);
  border-radius:18px;padding:22px 24px;
  position:relative;overflow:hidden;
}
.chart-card::before{
  content:'';position:absolute;top:0;left:0;right:0;height:1px;
  background:linear-gradient(90deg, transparent, var(--indigo), var(--cyan), transparent);
  opacity:0.5;
}
.chart-title{
  font-family:'Space Grotesk',sans-serif;
  font-size:14px;font-weight:600;color:var(--text-1);
  margin-bottom:16px;display:flex;align-items:center;gap:10px;letter-spacing:-0.01em;
}
.chart-title i{
  color:var(--indigo2);font-size:14px;
  width:26px;height:26px;border-radius:8px;
  background:rgba(99,102,241,0.12);
  display:flex;align-items:center;justify-content:center;
}
.chart-legend{display:flex;flex-wrap:wrap;gap:14px;margin-bottom:14px}
.legend-item{display:flex;align-items:center;gap:6px;font-size:11.5px;color:var(--text-2);font-weight:500}
.legend-dot{width:10px;height:10px;border-radius:3px;flex-shrink:0;box-shadow:0 0 8px currentColor}

/* ── Table ── */
.table-card{
  background:var(--bg-glass);
  backdrop-filter:blur(20px) saturate(180%);
  -webkit-backdrop-filter:blur(20px) saturate(180%);
  border:1px solid var(--border-soft);
  border-radius:18px;overflow:hidden;
  position:relative;
}
.table-card::before{
  content:'';position:absolute;top:0;left:0;right:0;height:1px;
  background:linear-gradient(90deg, transparent, var(--violet), var(--pink), transparent);
  opacity:0.5;z-index:1;
}
.table-header{
  padding:18px 24px;border-bottom:1px solid var(--border-soft);
  display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:10px;
}
.table-title{
  font-family:'Space Grotesk',sans-serif;
  font-size:14px;font-weight:600;color:var(--text-1);
  display:flex;align-items:center;gap:10px;letter-spacing:-0.01em;
}
.table-title i{
  color:var(--violet2);font-size:14px;
  width:26px;height:26px;border-radius:8px;
  background:rgba(168,85,247,0.12);
  display:flex;align-items:center;justify-content:center;
}
.table-date{font-size:12px;color:var(--text-2);font-weight:500;font-family:'JetBrains Mono',monospace;padding:5px 12px;background:rgba(99,102,241,0.08);border:1px solid rgba(99,102,241,0.2);border-radius:8px}
.table-wrap{overflow-x:auto;max-height:600px;overflow-y:auto}
table{width:100%;border-collapse:collapse;font-size:13px}
th{
  background:rgba(15,20,35,0.85);
  backdrop-filter:blur(10px);
  padding:13px 16px;text-align:left;
  font-size:10px;font-weight:700;color:var(--text-3);
  text-transform:uppercase;letter-spacing:0.12em;
  white-space:nowrap;border-bottom:1px solid var(--border-soft);
  position:sticky;top:0;z-index:5;
}
td{
  padding:13px 16px;border-bottom:1px solid var(--border-soft);
  vertical-align:middle;white-space:nowrap;
  font-size:13px;
}
tr:last-child td{border-bottom:none}
tbody tr{transition:background .15s}
tbody tr:hover td{background:rgba(99,102,241,0.04)}

.store-cell{display:flex;align-items:center;gap:11px}
.store-avatar{
  width:36px;height:36px;border-radius:10px;
  display:flex;align-items:center;justify-content:center;
  font-family:'Space Grotesk',sans-serif;
  font-size:12px;font-weight:600;color:#fff;
  flex-shrink:0;letter-spacing:-0.5px;
  box-shadow:inset 0 1px 0 rgba(255,255,255,0.15);
}
.store-info{display:flex;flex-direction:column;gap:1px}
.store-name{font-weight:600;color:var(--text-1);font-size:13.5px;letter-spacing:-0.01em}
.store-id{font-size:10.5px;color:var(--text-3);font-family:'JetBrains Mono',monospace;font-weight:500}

.area-tag{
  display:inline-flex;align-items:center;gap:7px;
  font-size:12px;color:var(--text-2);white-space:nowrap;
  padding:5px 11px;border-radius:8px;
  background:rgba(99,148,255,0.06);border:1px solid rgba(99,148,255,0.12);
  font-weight:500;
}
.area-dot{width:7px;height:7px;border-radius:50%;flex-shrink:0;box-shadow:0 0 8px currentColor}

.num{font-family:'JetBrains Mono',monospace;font-size:12.5px;font-weight:500;letter-spacing:-0.01em}
.num-bold{font-weight:600}

.pill{
  display:inline-flex;align-items:center;gap:4px;
  padding:5px 11px;border-radius:8px;
  font-size:11.5px;font-weight:600;
  font-family:'JetBrains Mono',monospace;letter-spacing:-0.01em;
}
.pill.up{background:rgba(16,185,129,0.12);border:1px solid rgba(16,185,129,0.3);color:var(--emerald2);box-shadow:0 0 12px -4px var(--emerald-glow)}
.pill.down{background:rgba(244,63,94,0.12);border:1px solid rgba(244,63,94,0.3);color:var(--rose2);box-shadow:0 0 12px -4px var(--rose-glow)}
.pill.flat{background:rgba(148,163,200,0.08);border:1px solid var(--border-strong);color:var(--text-3)}

.just-cell{
  max-width:240px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;
  color:var(--text-2);font-size:12px;cursor:help;font-style:italic;font-weight:400;
}

.summary-row td{
  background:linear-gradient(180deg, rgba(99,102,241,0.06) 0%, rgba(99,102,241,0.02) 100%) !important;
  font-weight:700;border-top:2px solid rgba(99,102,241,0.3);
  font-family:'Space Grotesk',sans-serif;
}
.summary-row .num{font-size:13.5px;font-weight:700}

/* ── Tabs Navigation ── */
.tabs{
  display:flex;gap:4px;margin-bottom:24px;
  background:var(--bg-glass);
  backdrop-filter:blur(20px) saturate(180%);
  -webkit-backdrop-filter:blur(20px) saturate(180%);
  border:1px solid var(--border-soft);
  border-radius:14px;padding:5px;
  box-shadow:0 8px 32px -12px rgba(0,0,0,0.4);
  width:fit-content;
}
.tab-btn{
  background:transparent;border:none;color:var(--text-2);
  padding:10px 22px;border-radius:10px;cursor:pointer;
  font-size:13px;font-family:'Inter',sans-serif;font-weight:600;
  display:flex;align-items:center;gap:8px;
  transition:all .25s cubic-bezier(0.4,0,0.2,1);letter-spacing:-0.01em;
}
.tab-btn:hover{color:var(--text-1);background:rgba(99,102,241,0.06)}
.tab-btn.active{
  background:linear-gradient(135deg, rgba(99,102,241,0.18), rgba(6,182,212,0.18));
  color:#fff;
  box-shadow:0 4px 12px -2px var(--indigo-glow), inset 0 1px 0 rgba(255,255,255,0.1);
  border:1px solid rgba(99,102,241,0.35);
}
.tab-btn i{font-size:13px}
.tab-content{display:none}
.tab-content.active{display:block;animation:fadeInUp 0.4s cubic-bezier(0.4,0,0.2,1)}

/* ── Export button ── */
.export-btn{
  background:linear-gradient(135deg, var(--emerald), var(--emerald2));
  border:none;color:#fff;
  padding:9px 16px;border-radius:10px;cursor:pointer;
  font-size:12.5px;font-family:'Inter',sans-serif;font-weight:600;
  display:flex;align-items:center;gap:7px;transition:all .25s;
  letter-spacing:-0.01em;
  box-shadow:0 4px 12px -2px var(--emerald-glow), inset 0 1px 0 rgba(255,255,255,0.2);
}
.export-btn:hover{transform:translateY(-1px);box-shadow:0 8px 20px -4px var(--emerald-glow)}
.export-btn:active{transform:translateY(0)}
.export-btn:disabled{opacity:0.5;cursor:not-allowed;transform:none}

/* ── Justification full-text cell (no truncation) ── */
.just-full{
  white-space:normal;color:var(--text-2);font-size:12px;
  line-height:1.5;font-style:italic;max-width:380px;font-weight:400;
}

/* ── Sign filter toggle group ── */
.sign-toggle{
  display:flex;gap:4px;
  background:rgba(15,20,35,0.7);
  border:1px solid var(--border-strong);
  border-radius:10px;padding:3px;
}
.sign-toggle button{
  background:transparent;border:none;color:var(--text-3);
  padding:6px 12px;border-radius:7px;cursor:pointer;
  font-size:11.5px;font-family:'Inter',sans-serif;font-weight:600;
  transition:all .2s;letter-spacing:-0.01em;
  display:flex;align-items:center;gap:5px;
}
.sign-toggle button:hover{color:var(--text-1)}
.sign-toggle button.active.all{background:rgba(99,102,241,0.2);color:var(--indigo2);box-shadow:inset 0 1px 0 rgba(255,255,255,0.05)}
.sign-toggle button.active.pos{background:rgba(16,185,129,0.18);color:var(--emerald2);box-shadow:inset 0 1px 0 rgba(255,255,255,0.05)}
.sign-toggle button.active.neg{background:rgba(244,63,94,0.18);color:var(--rose2);box-shadow:inset 0 1px 0 rgba(255,255,255,0.05)}

/* ── Missing Stores card ── */
.missing-card{
  margin-top:24px;
  background:linear-gradient(180deg, rgba(245,158,11,0.06) 0%, rgba(20,26,42,0.65) 60%);
  backdrop-filter:blur(20px) saturate(180%);
  -webkit-backdrop-filter:blur(20px) saturate(180%);
  border:1px solid rgba(245,158,11,0.22);
  border-radius:18px;overflow:hidden;position:relative;
  animation:fadeInUp 0.5s cubic-bezier(0.4,0,0.2,1) 0.5s backwards;
}
.missing-card::before{
  content:'';position:absolute;top:0;left:0;right:0;height:1px;
  background:linear-gradient(90deg, transparent, var(--amber), var(--amber2), transparent);
  opacity:0.7;
}
.missing-card.empty-state{
  background:linear-gradient(180deg, rgba(16,185,129,0.05) 0%, rgba(20,26,42,0.65) 60%);
  border-color:rgba(16,185,129,0.22);
}
.missing-card.empty-state::before{
  background:linear-gradient(90deg, transparent, var(--emerald), var(--emerald2), transparent);
}
.missing-header{
  padding:18px 24px;border-bottom:1px solid var(--border-soft);
  display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:10px;
}
.missing-title{
  font-family:'Space Grotesk',sans-serif;
  font-size:14px;font-weight:600;color:var(--text-1);
  display:flex;align-items:center;gap:10px;letter-spacing:-0.01em;
}
.missing-title i{
  color:var(--amber2);font-size:13px;
  width:28px;height:28px;border-radius:8px;
  background:rgba(245,158,11,0.15);
  display:flex;align-items:center;justify-content:center;
}
.missing-card.empty-state .missing-title i{
  color:var(--emerald2);background:rgba(16,185,129,0.15);
}
.missing-count{
  display:inline-flex;align-items:center;justify-content:center;
  min-width:24px;height:22px;padding:0 8px;border-radius:11px;
  background:linear-gradient(135deg, var(--amber), var(--amber2));
  color:#fff;font-size:11px;font-weight:700;
  font-family:'JetBrains Mono',monospace;
  box-shadow:0 2px 8px -2px var(--amber-glow);
}
.missing-card.empty-state .missing-count{
  background:linear-gradient(135deg, var(--emerald), var(--emerald2));
  box-shadow:0 2px 8px -2px var(--emerald-glow);
}
.missing-sub{font-size:11.5px;color:var(--text-3);font-weight:500}
.missing-table th{
  background:rgba(245,158,11,0.05);
  color:var(--amber2);
}
.remark-pill{
  display:inline-flex;align-items:center;
  padding:4px 10px;border-radius:8px;
  font-size:11px;font-weight:600;font-family:'JetBrains Mono',monospace;
  letter-spacing:-0.01em;
}
.remark-pill.organic{background:rgba(16,185,129,0.12);border:1px solid rgba(16,185,129,0.3);color:var(--emerald2)}
.remark-pill.new{background:rgba(99,102,241,0.12);border:1px solid rgba(99,102,241,0.3);color:var(--indigo2)}
.remark-pill.other{background:rgba(148,163,200,0.08);border:1px solid var(--border-strong);color:var(--text-3)}
.all-reported{
  padding:50px 20px;text-align:center;color:var(--text-2);
}
.all-reported-icon{
  width:64px;height:64px;border-radius:18px;margin:0 auto 16px;
  background:linear-gradient(135deg, rgba(16,185,129,0.15), rgba(52,211,153,0.15));
  border:1px solid rgba(16,185,129,0.3);
  display:flex;align-items:center;justify-content:center;
  font-size:24px;color:var(--emerald2);
  box-shadow:0 0 24px -8px var(--emerald-glow);
}
.all-reported-title{font-family:'Space Grotesk',sans-serif;font-size:15px;font-weight:600;color:var(--emerald2);margin-bottom:6px}
.all-reported-sub{font-size:12px;color:var(--text-3);font-weight:500}

/* ── Empty / Error states ── */
.empty-cell{text-align:center;padding:64px 20px;color:var(--text-3)}
.empty-icon{
  width:64px;height:64px;border-radius:18px;margin:0 auto 16px;
  background:linear-gradient(135deg, rgba(99,102,241,0.1), rgba(168,85,247,0.1));
  border:1px solid var(--border-soft);
  display:flex;align-items:center;justify-content:center;
  font-size:24px;color:var(--text-3);
}
.empty-cell p{font-size:13.5px;margin-bottom:6px;color:var(--text-2);font-weight:500}
.empty-cell small{font-size:11.5px;color:var(--text-3)}

/* ── Loading overlay ── */
.loading-overlay{
  position:fixed;inset:0;
  background:radial-gradient(ellipse at center, rgba(10,14,26,0.92), rgba(7,10,20,0.98));
  backdrop-filter:blur(12px);
  display:flex;align-items:center;justify-content:center;z-index:999;
}
.loading-box{text-align:center}
.spinner{
  width:54px;height:54px;
  border:3px solid transparent;
  border-top-color:var(--indigo);
  border-right-color:var(--cyan);
  border-radius:50%;
  animation:spin 1s linear infinite;
  margin:0 auto 22px;
  filter:drop-shadow(0 0 12px var(--indigo-glow));
}
.loading-title{
  font-family:'Space Grotesk',sans-serif;
  font-size:18px;font-weight:600;color:var(--text-1);margin-bottom:6px;
  background:linear-gradient(135deg, #6366f1, #06b6d4);
  -webkit-background-clip:text;background-clip:text;-webkit-text-fill-color:transparent;
  letter-spacing:-0.01em;
}
.loading-msg{font-size:12.5px;color:var(--text-2);font-weight:500}

/* ── Scrollbar ── */
::-webkit-scrollbar{width:8px;height:8px}
::-webkit-scrollbar-track{background:rgba(15,20,35,0.4)}
::-webkit-scrollbar-thumb{background:linear-gradient(180deg, var(--indigo), #4f46e5);border-radius:4px}
::-webkit-scrollbar-thumb:hover{background:linear-gradient(180deg, var(--indigo2), var(--indigo))}

/* ── Animation ── */
@keyframes fadeInUp{
  from{opacity:0;transform:translateY(8px)}
  to{opacity:1;transform:translateY(0)}
}
.kpi,.chart-card,.table-card,.controls{animation:fadeInUp 0.5s cubic-bezier(0.4,0,0.2,1) backwards}
.controls{animation-delay:0.05s}
.kpi:nth-child(1){animation-delay:0.1s}
.kpi:nth-child(2){animation-delay:0.15s}
.kpi:nth-child(3){animation-delay:0.2s}
.kpi:nth-child(4){animation-delay:0.25s}
.kpi:nth-child(5){animation-delay:0.3s}
.chart-card:nth-child(1){animation-delay:0.35s}
.chart-card:nth-child(2){animation-delay:0.4s}
.table-card{animation-delay:0.45s}
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

  <!-- Tabs -->
  <div class="tabs">
    <button class="tab-btn active" data-tab="daily" onclick="switchTab('daily')">
      <i class="fa fa-calendar-day"></i> Daily Sales
    </button>
    <button class="tab-btn" data-tab="monthly" onclick="switchTab('monthly')">
      <i class="fa fa-calendar"></i> Monthly Sales
    </button>
  </div>

  <!-- ═══════════════════ DAILY TAB ═══════════════════ -->
  <div class="tab-content active" id="tab-daily">

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
    <div class="kpi k-sales">
      <div class="kpi-label"><div class="kpi-icon"><i class="fa fa-peso-sign"></i></div>Total Sales</div>
      <div class="kpi-value gradient-text" id="kpi-sales">—</div>
      <div class="kpi-sub">Current period</div>
    </div>
    <div class="kpi k-ly">
      <div class="kpi-label"><div class="kpi-icon"><i class="fa fa-clock-rotate-left"></i></div>Sales LY</div>
      <div class="kpi-value gradient-text" id="kpi-ly">—</div>
      <div class="kpi-sub">Same period last year</div>
    </div>
    <div class="kpi k-pct">
      <div class="kpi-label"><div class="kpi-icon" id="kpi-pct-icon"><i class="fa fa-percent"></i></div>Diff %</div>
      <div class="kpi-value" id="kpi-pct">—</div>
      <div class="kpi-sub" id="kpi-pct-sub">vs last year</div>
    </div>
    <div class="kpi k-diff">
      <div class="kpi-label"><div class="kpi-icon" id="kpi-diff-icon"><i class="fa fa-arrow-trend-up"></i></div>Diff Amount</div>
      <div class="kpi-value" id="kpi-diff">—</div>
      <div class="kpi-sub" id="kpi-diff-sub">variance</div>
    </div>
    <div class="kpi k-stores">
      <div class="kpi-label"><div class="kpi-icon"><i class="fa fa-store"></i></div>Active Stores</div>
      <div class="kpi-value gradient-text" id="kpi-stores">—</div>
      <div class="kpi-sub">with sales data</div>
    </div>
  </div>

  <!-- Charts -->
  <div class="charts-grid">
    <div class="chart-card">
      <div class="chart-title"><i class="fa fa-chart-bar"></i> Sales by Store — Current vs Last Year</div>
      <div class="chart-legend">
        <span class="legend-item"><span class="legend-dot" style="background:#6366f1;color:#6366f1"></span>Sales Current</span>
        <span class="legend-item"><span class="legend-dot" style="background:#3a4560;color:transparent;box-shadow:none"></span>Sales LY</span>
      </div>
      <div style="position:relative;height:280px">
        <canvas id="barChart" role="img" aria-label="Bar chart: Sales by store current vs last year">Sales by store chart</canvas>
      </div>
    </div>
    <div class="chart-card">
      <div class="chart-title"><i class="fa fa-chart-pie"></i> Sales Share by Area</div>
      <div style="position:relative;height:320px">
        <canvas id="pieChart" role="img" aria-label="Doughnut chart: Sales distribution by area">Area distribution chart</canvas>
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
          <tr><td colspan="7" class="empty-cell"><div class="empty-icon"><i class="fa fa-spinner fa-spin"></i></div><p>Loading data...</p></td></tr>
        </tbody>
      </table>
    </div>
  <!-- Stores Without Sales Entry -->
  <div class="missing-card" id="missingCard" style="display:none">
    <div class="missing-header">
      <div class="missing-title">
        <i class="fa fa-triangle-exclamation"></i>
        Stores Without Sales Entry
        <span class="missing-count" id="missingCount">0</span>
      </div>
      <div class="missing-sub">Based on Store ID vs ListOfStores</div>
    </div>
    <div class="table-wrap">
      <table class="missing-table">
        <thead>
          <tr>
            <th>Store ID</th>
            <th>Store Name</th>
            <th>Area</th>
            <th>Region</th>
            <th>Remarks</th>
          </tr>
        </thead>
        <tbody id="missingBody"></tbody>
      </table>
    </div>
  </div>

  </div><!-- /tab-daily -->

  <!-- ═══════════════════ MONTHLY TAB ═══════════════════ -->
  <div class="tab-content" id="tab-monthly">

    <!-- Controls -->
    <div class="controls">
      <div class="ctrl-group">
        <span class="ctrl-label"><i class="fa fa-calendar"></i> Month</span>
        <select id="mMonthFilter" onchange="applyMonthlyFilters()"></select>
      </div>
      <div class="divider"></div>
      <div class="ctrl-group">
        <span class="ctrl-label"><i class="fa fa-layer-group"></i> Area</span>
        <select id="mAreaFilter" onchange="onMAreaChange()">
          <option value="ALL">All Areas</option>
        </select>
      </div>
      <div class="records-count" id="mRecordsCount">—</div>
    </div>

    <!-- KPIs -->
    <div class="kpi-grid">
      <div class="kpi k-sales">
        <div class="kpi-label"><div class="kpi-icon"><i class="fa fa-peso-sign"></i></div>Month Sales</div>
        <div class="kpi-value gradient-text" id="m-kpi-sales">—</div>
        <div class="kpi-sub">Selected month</div>
      </div>
      <div class="kpi k-ly">
        <div class="kpi-label"><div class="kpi-icon"><i class="fa fa-clock-rotate-left"></i></div>Sales LY</div>
        <div class="kpi-value gradient-text" id="m-kpi-ly">—</div>
        <div class="kpi-sub">Same month last year</div>
      </div>
      <div class="kpi k-pct">
        <div class="kpi-label"><div class="kpi-icon" id="m-kpi-pct-icon"><i class="fa fa-percent"></i></div>Diff %</div>
        <div class="kpi-value" id="m-kpi-pct">—</div>
        <div class="kpi-sub" id="m-kpi-pct-sub">vs last year</div>
      </div>
      <div class="kpi k-diff">
        <div class="kpi-label"><div class="kpi-icon" id="m-kpi-diff-icon"><i class="fa fa-arrow-trend-up"></i></div>Diff Amount</div>
        <div class="kpi-value" id="m-kpi-diff">—</div>
        <div class="kpi-sub" id="m-kpi-diff-sub">variance</div>
      </div>
      <div class="kpi k-stores">
        <div class="kpi-label"><div class="kpi-icon"><i class="fa fa-store"></i></div>Active Stores</div>
        <div class="kpi-value gradient-text" id="m-kpi-stores">—</div>
        <div class="kpi-sub">with sales data</div>
      </div>
    </div>

    <!-- Charts -->
    <div class="charts-grid">
      <div class="chart-card">
        <div class="chart-title"><i class="fa fa-chart-bar"></i> Monthly Sales by Store — Current vs Last Year</div>
        <div class="chart-legend">
          <span class="legend-item"><span class="legend-dot" style="background:#6366f1;color:#6366f1"></span>Sales Current</span>
          <span class="legend-item"><span class="legend-dot" style="background:#3a4560;color:transparent;box-shadow:none"></span>Sales LY</span>
        </div>
        <div style="position:relative;height:280px">
          <canvas id="mBarChart"></canvas>
        </div>
      </div>
      <div class="chart-card">
        <div class="chart-title"><i class="fa fa-chart-line"></i> Daily Sales Trend</div>
        <div style="position:relative;height:320px">
          <canvas id="mLineChart"></canvas>
        </div>
      </div>
    </div>

    <!-- Store Summary Table -->
    <div class="table-card">
      <div class="table-header">
        <div class="table-title"><i class="fa fa-table-list"></i> Stores Summary</div>
        <div class="table-date" id="mTableDate">—</div>
      </div>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>Store ID</th>
              <th>Store Name</th>
              <th>Area</th>
              <th style="text-align:right">Sales</th>
              <th style="text-align:right">Sales LY</th>
              <th style="text-align:center">Diff %</th>
              <th style="text-align:right">Diff Amount</th>
            </tr>
          </thead>
          <tbody id="mSummaryBody">
            <tr><td colspan="7" class="empty-cell"><div class="empty-icon"><i class="fa fa-spinner fa-spin"></i></div><p>Loading...</p></td></tr>
          </tbody>
        </table>
      </div>
    </div>

    <!-- Detail Report -->
    <div class="table-card" style="margin-top:24px">
      <div class="table-header" style="gap:14px">
        <div class="table-title"><i class="fa fa-rectangle-list"></i> Detail Report</div>

        <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin-left:auto">

          <!-- Store filter for detail -->
          <div class="ctrl-group">
            <span class="ctrl-label" style="font-size:10px"><i class="fa fa-store"></i> Store</span>
            <select id="mStoreFilter" onchange="applyMonthlyFilters()">
              <option value="ALL">All Stores</option>
            </select>
          </div>

          <!-- Sign filter -->
          <div class="ctrl-group">
            <span class="ctrl-label" style="font-size:10px"><i class="fa fa-filter"></i> Variance</span>
            <div class="sign-toggle" id="signToggle">
              <button class="active all" data-sign="ALL" onclick="setSignFilter('ALL')">All</button>
              <button data-sign="POS" onclick="setSignFilter('POS')"><i class="fa fa-arrow-up"></i> Positive</button>
              <button data-sign="NEG" onclick="setSignFilter('NEG')"><i class="fa fa-arrow-down"></i> Negative</button>
            </div>
          </div>

          <!-- Export -->
          <button class="export-btn" id="exportBtn" onclick="exportDetailToExcel()">
            <i class="fa fa-file-excel"></i> Export Excel
          </button>
        </div>
      </div>
      <div class="table-wrap" style="max-height:700px">
        <table>
          <thead>
            <tr>
              <th>Date</th>
              <th>Day</th>
              <th>Day YA</th>
              <th>Store</th>
              <th style="text-align:right">Sales</th>
              <th style="text-align:right">Sales LY</th>
              <th style="text-align:center">Diff %</th>
              <th style="text-align:right">Diff Amount</th>
              <th>Justification</th>
            </tr>
          </thead>
          <tbody id="mDetailBody">
            <tr><td colspan="9" class="empty-cell"><div class="empty-icon"><i class="fa fa-spinner fa-spin"></i></div><p>Loading...</p></td></tr>
          </tbody>
        </table>
      </div>
    </div>

  </div><!-- /tab-monthly -->

</main>

<script>
// ─── Vibrant area palette ──────────────────────────────────────────────────
const AREA_COLORS = {
  'Valenzuela':      '#6366f1',  // indigo
  'South Caloocan':  '#10b981',  // emerald
  'Malabon-Navotas': '#f59e0b',  // amber
  'North Caloocan':  '#a855f7'   // violet
};
const AREA_GRADIENTS = {
  'Valenzuela':      ['#6366f1','#818cf8'],
  'South Caloocan':  ['#10b981','#34d399'],
  'Malabon-Navotas': ['#f59e0b','#fbbf24'],
  'North Caloocan':  ['#a855f7','#c084fc']
};
const DEFAULT_COLOR = '#94a3b8';

let allFilters = { areas: [], stores: [] };
let barChartInst = null;
let pieChartInst = null;

// Monthly tab state
let mBarChartInst = null;
let mLineChartInst = null;
let mSignFilter = 'ALL';
let mDetailRowsCache = []; // cached current detail rows for export
let mMonthFilters = { areas: [], stores: [], months: [] };

// ─── Tab switching ───────────────────────────────────────────────────────
function switchTab(tab) {
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.toggle('active', b.dataset.tab === tab));
  document.querySelectorAll('.tab-content').forEach(c => c.classList.toggle('active', c.id === \`tab-\${tab}\`));
  if (tab === 'monthly' && !mMonthFilters.months.length) {
    initMonthlyTab();
  }
}

function today() { return new Date().toLocaleDateString('en-CA'); }

function fmt(n) {
  if (n === null || n === undefined || isNaN(n)) return '—';
  const abs = Math.abs(n);
  const sign = n < 0 ? '-' : '';
  if (abs >= 1e6) return sign + '₱' + (abs/1e6).toFixed(2) + 'M';
  if (abs >= 1e3) return sign + '₱' + (abs/1e3).toFixed(1) + 'K';
  return sign + '₱' + abs.toFixed(2);
}

function fmtFull(n) {
  if (!n && n !== 0) return '—';
  return '₱' + Math.abs(n).toLocaleString('en-PH', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function initials(name) {
  return (name || '?').split(/\\s+/).filter(Boolean).slice(0,2).map(w => w[0]).join('').toUpperCase();
}

async function loadFilters() {
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
  if (area === 'ALL') populateStores(allFilters.stores);
  else {
    try {
      const res = await fetch(\`/api/filters?area=\${encodeURIComponent(area)}\`);
      const json = await res.json();
      if (json.success) populateStores(json.stores);
    } catch(e){}
  }
  document.getElementById('storeFilter').value = 'ALL';
  await applyFilters();
}

async function applyFilters() {
  const date = document.getElementById('dateFilter').value;
  const area = document.getElementById('areaFilter').value;
  const store = document.getElementById('storeFilter').value;
  const params = new URLSearchParams();
  if (date) params.set('date', date);
  if (area !== 'ALL') params.set('area', area);
  if (store !== 'ALL') params.set('store', store);
  try {
    const res = await fetch('/api/sales?' + params);
    const json = await res.json();
    if (!json.success) throw new Error(json.error);
    const rows = json.rows || [];
    document.getElementById('recordsCount').innerHTML = \`<span>\${rows.length}</span> store\${rows.length!==1?'s':''} · <span>\${json.count}</span> row\${json.count!==1?'s':''}\`;
    const dateLabel = date ? new Date(date + 'T00:00:00').toLocaleDateString('en-PH', { weekday:'long', year:'numeric', month:'long', day:'numeric' }) : 'All dates';
    document.getElementById('tableDate').textContent = dateLabel;
    renderKPIs(rows);
    renderTable(rows);
    renderCharts(rows);
    renderMissing(json.missing || [], !!date);
    setStatus('live', 'Live');
  } catch(e) {
    console.error(e);
    setStatus('error', 'Error');
    showTableError('Error loading sales data: ' + e.message);
  }
}

function renderKPIs(rows) {
  const totSales = rows.reduce((s,r)=>s+(r.sales||0),0);
  const totLY = rows.reduce((s,r)=>s+(r.salesLY||0),0);
  const totDiff = totSales - totLY;
  const totPct = totLY !== 0 ? (totDiff/totLY*100) : 0;
  const active = rows.filter(r=>r.sales>0).length;
  document.getElementById('kpi-sales').textContent = fmt(totSales);
  document.getElementById('kpi-ly').textContent = fmt(totLY);

  const pctEl = document.getElementById('kpi-pct');
  pctEl.textContent = (totPct>=0?'+':'') + totPct.toFixed(2) + '%';
  pctEl.style.background = \`linear-gradient(135deg, \${totPct>=0?'#10b981':'#f43f5e'} 0%, \${totPct>=0?'#34d399':'#fb7185'} 100%)\`;
  pctEl.style.webkitBackgroundClip = 'text';
  pctEl.style.backgroundClip = 'text';
  pctEl.style.webkitTextFillColor = 'transparent';
  document.getElementById('kpi-pct-sub').innerHTML = \`<i class="fa fa-arrow-\${totPct>=0?'up':'down'}" style="color:\${totPct>=0?'#34d399':'#fb7185'}"></i> vs last year\`;
  document.getElementById('kpi-pct-icon').style.background = \`linear-gradient(135deg, \${totPct>=0?'#10b981':'#f43f5e'}, \${totPct>=0?'#34d399':'#fb7185'})\`;

  const diffEl = document.getElementById('kpi-diff');
  diffEl.textContent = (totDiff>=0?'+':'') + fmt(totDiff);
  diffEl.style.background = \`linear-gradient(135deg, \${totDiff>=0?'#06b6d4':'#f43f5e'} 0%, \${totDiff>=0?'#22d3ee':'#fb7185'} 100%)\`;
  diffEl.style.webkitBackgroundClip = 'text';
  diffEl.style.backgroundClip = 'text';
  diffEl.style.webkitTextFillColor = 'transparent';
  document.getElementById('kpi-diff-sub').innerHTML = \`<i class="fa fa-\${totDiff>=0?'circle-check':'circle-xmark'}" style="color:\${totDiff>=0?'#22d3ee':'#fb7185'}"></i> variance\`;
  document.getElementById('kpi-diff-icon').style.background = \`linear-gradient(135deg, \${totDiff>=0?'#06b6d4':'#f43f5e'}, \${totDiff>=0?'#22d3ee':'#fb7185'})\`;

  document.getElementById('kpi-stores').textContent = active;
}

function renderTable(rows) {
  const tbody = document.getElementById('tableBody');
  if (!rows.length) {
    tbody.innerHTML = \`<tr><td colspan="7" class="empty-cell"><div class="empty-icon"><i class="fa fa-magnifying-glass"></i></div><p>No data for selected filters</p><small>Try a different date or area</small></td></tr>\`;
    return;
  }
  const totSales = rows.reduce((s,r)=>s+r.sales,0);
  const totLY = rows.reduce((s,r)=>s+r.salesLY,0);
  const totDiff = totSales - totLY;
  const totPct = totLY ? (totDiff/totLY*100) : 0;

  let html = rows.map(r => {
    const pctCls = r.diffPct > 0.05 ? 'up' : r.diffPct < -0.05 ? 'down' : 'flat';
    const color = AREA_COLORS[r.area] || DEFAULT_COLOR;
    const grad = AREA_GRADIENTS[r.area] || [DEFAULT_COLOR, DEFAULT_COLOR];
    const diffColor = r.diffVal >= 0 ? '#34d399' : '#fb7185';
    const just = (r.justification || '—').replace(/"/g, '&quot;');
    const arrow = r.diffPct > 0.05 ? '↑' : r.diffPct < -0.05 ? '↓' : '—';
    const pctStr = r.diffPct !== 0 ? \`\${arrow} \${Math.abs(r.diffPct).toFixed(2)}%\` : '—';
    return \`<tr>
      <td>
        <div class="store-cell">
          <div class="store-avatar" style="background:linear-gradient(135deg, \${grad[0]}, \${grad[1]})">\${initials(r.storeName)}</div>
          <div class="store-info">
            <div class="store-name">\${r.storeName || '—'}</div>
            <div class="store-id">#\${r.storeId}</div>
          </div>
        </div>
      </td>
      <td>
        <span class="area-tag">
          <span class="area-dot" style="background:\${color};color:\${color}"></span>
          \${r.area || '—'}
        </span>
      </td>
      <td style="text-align:right"><span class="num num-bold" style="color:var(--text-1)">\${fmtFull(r.sales)}</span></td>
      <td style="text-align:right"><span class="num" style="color:var(--text-3)">\${fmtFull(r.salesLY)}</span></td>
      <td style="text-align:center"><span class="pill \${pctCls}">\${pctStr}</span></td>
      <td style="text-align:right"><span class="num num-bold" style="color:\${diffColor}">\${r.diffVal >= 0 ? '+' : ''}\${fmtFull(r.diffVal)}</span></td>
      <td><div class="just-cell" title="\${just}">\${just}</div></td>
    </tr>\`;
  }).join('');

  const totPctCls = totPct > 0.05 ? 'up' : totPct < -0.05 ? 'down' : 'flat';
  const totArrow = totPct > 0.05 ? '↑' : totPct < -0.05 ? '↓' : '—';
  html += \`<tr class="summary-row">
    <td colspan="2" style="font-size:12px;letter-spacing:0.1em;text-transform:uppercase;color:var(--indigo2)">TOTAL · \${rows.length} STORES</td>
    <td style="text-align:right"><span class="num">\${fmtFull(totSales)}</span></td>
    <td style="text-align:right"><span class="num" style="color:var(--text-3)">\${fmtFull(totLY)}</span></td>
    <td style="text-align:center"><span class="pill \${totPctCls}">\${totArrow} \${Math.abs(totPct).toFixed(2)}%</span></td>
    <td style="text-align:right"><span class="num" style="color:\${totDiff>=0?'#34d399':'#fb7185'}">\${totDiff>=0?'+':''}\${fmtFull(totDiff)}</span></td>
    <td></td>
  </tr>\`;
  tbody.innerHTML = html;
}

function renderCharts(rows) {
  const top = rows.slice(0, 14);
  const labels = top.map(r => r.storeName);
  const salesArr = top.map(r => r.sales);
  const lyArr = top.map(r => r.salesLY);

  // Create gradients for bars
  if (barChartInst) barChartInst.destroy();
  const barCanvas = document.getElementById('barChart');
  const barCtx = barCanvas.getContext('2d');

  const gradients = top.map(r => {
    const [c1, c2] = AREA_GRADIENTS[r.area] || [DEFAULT_COLOR, DEFAULT_COLOR];
    const g = barCtx.createLinearGradient(0, 0, 0, 280);
    g.addColorStop(0, c2);
    g.addColorStop(1, c1);
    return g;
  });

  barChartInst = new Chart(barCanvas, {
    type: 'bar',
    data: {
      labels,
      datasets: [
        { label: 'Sales', data: salesArr, backgroundColor: gradients, borderRadius: 6, borderSkipped: false, barPercentage: 0.7 },
        { label: 'Sales LY', data: lyArr, backgroundColor: 'rgba(74, 83, 107, 0.4)', borderColor: 'rgba(148, 163, 200, 0.2)', borderWidth: 1, borderRadius: 6, borderSkipped: false, barPercentage: 0.7 }
      ]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          backgroundColor: 'rgba(15, 20, 35, 0.95)',
          borderColor: 'rgba(99, 102, 241, 0.4)',
          borderWidth: 1,
          padding: 12,
          titleFont: { family: "'Space Grotesk'", size: 13, weight: '600' },
          bodyFont: { family: "'JetBrains Mono'", size: 11.5 },
          titleColor: '#f0f3fb',
          bodyColor: '#a8b3d1',
          cornerRadius: 10,
          displayColors: true,
          boxPadding: 6,
          callbacks: {
            label: ctx => \` \${ctx.dataset.label}: ₱\${ctx.raw.toLocaleString('en-PH', { minimumFractionDigits: 2 })}\`
          }
        }
      },
      scales: {
        x: {
          grid: { color: 'rgba(148,163,200,0.04)', drawBorder: false },
          ticks: { color: '#a8b3d1', font: { size: 10, family: "'Inter'", weight: '500' }, maxRotation: 40 }
        },
        y: {
          grid: { color: 'rgba(148,163,200,0.06)' },
          ticks: {
            color: '#a8b3d1', font: { size: 10.5, family: "'JetBrains Mono'", weight: '500' },
            callback: v => v >= 1e6 ? '₱' + (v/1e6).toFixed(1)+'M' : v >= 1e3 ? '₱' + (v/1e3).toFixed(0)+'K' : v
          }
        }
      }
    }
  });

  // Pie chart
  const areaMap = {};
  rows.forEach(r => { areaMap[r.area] = (areaMap[r.area] || 0) + r.sales; });
  const areaLabels = Object.keys(areaMap);
  const areaVals = Object.values(areaMap);
  const pieColors = areaLabels.map(a => AREA_COLORS[a] || DEFAULT_COLOR);

  if (pieChartInst) pieChartInst.destroy();
  const pieCanvas = document.getElementById('pieChart');
  const pieCtx = pieCanvas.getContext('2d');
  const pieGradients = areaLabels.map(a => {
    const [c1, c2] = AREA_GRADIENTS[a] || [DEFAULT_COLOR, DEFAULT_COLOR];
    const g = pieCtx.createLinearGradient(0, 0, 320, 320);
    g.addColorStop(0, c2);
    g.addColorStop(1, c1);
    return g;
  });

  pieChartInst = new Chart(pieCanvas, {
    type: 'doughnut',
    data: { labels: areaLabels, datasets: [{ data: areaVals, backgroundColor: pieGradients, borderColor: 'rgba(10, 14, 26, 0.8)', borderWidth: 4, hoverOffset: 12, hoverBorderWidth: 2 }] },
    options: {
      responsive: true, maintainAspectRatio: false, cutout: '65%',
      plugins: {
        legend: {
          position: 'bottom',
          labels: {
            color: '#a8b3d1', font: { size: 11.5, family: "'Inter'", weight: '500' }, padding: 16, boxWidth: 11, boxHeight: 11,
            generateLabels: chart => chart.data.labels.map((label, i) => ({
              text: \`\${label}  \${fmt(areaVals[i])}\`,
              fillStyle: pieColors[i], strokeStyle: 'transparent',
              lineWidth: 0, pointStyle: 'rectRounded'
            }))
          }
        },
        tooltip: {
          backgroundColor: 'rgba(15, 20, 35, 0.95)',
          borderColor: 'rgba(99, 102, 241, 0.4)',
          borderWidth: 1, padding: 12, cornerRadius: 10,
          titleFont: { family: "'Space Grotesk'", weight: '600', size: 13 },
          bodyFont: { family: "'JetBrains Mono'", size: 11.5 },
          titleColor: '#f0f3fb', bodyColor: '#a8b3d1',
          callbacks: {
            label: ctx => {
              const total = ctx.dataset.data.reduce((a,b)=>a+b,0);
              const pct = total ? ((ctx.raw/total)*100).toFixed(1) : 0;
              return \` ₱\${ctx.raw.toLocaleString('en-PH',{minimumFractionDigits:2})}  (\${pct}%)\`;
            }
          }
        }
      }
    }
  });
}

function renderMissing(missing, hasDateFilter) {
  const card = document.getElementById('missingCard');
  const body = document.getElementById('missingBody');
  const countEl = document.getElementById('missingCount');

  // Only show this section when a specific date is selected
  if (!hasDateFilter) {
    card.style.display = 'none';
    return;
  }

  card.style.display = 'block';
  countEl.textContent = missing.length;

  if (!missing.length) {
    card.classList.add('empty-state');
    body.innerHTML = \`<tr><td colspan="5" style="padding:0">
      <div class="all-reported">
        <div class="all-reported-icon"><i class="fa fa-circle-check"></i></div>
        <div class="all-reported-title">All Stores Reported</div>
        <div class="all-reported-sub">Every store in the master list has a sales entry for this date</div>
      </div>
    </td></tr>\`;
    return;
  }

  card.classList.remove('empty-state');

  body.innerHTML = missing.map(s => {
    const color = AREA_COLORS[s.area] || DEFAULT_COLOR;
    const grad  = AREA_GRADIENTS[s.area] || [DEFAULT_COLOR, DEFAULT_COLOR];
    const remarkLower = (s.remarks || '').toLowerCase();
    const remarkCls = remarkLower === 'organic' ? 'organic' : remarkLower === 'new' ? 'new' : 'other';
    return \`<tr>
      <td><span class="num num-bold" style="color:var(--text-1)">#\${s.storeId}</span></td>
      <td>
        <div class="store-cell">
          <div class="store-avatar" style="background:linear-gradient(135deg, \${grad[0]}, \${grad[1]})">\${initials(s.storeName)}</div>
          <div class="store-info">
            <div class="store-name">\${s.storeName || '—'}</div>
          </div>
        </div>
      </td>
      <td>
        <span class="area-tag">
          <span class="area-dot" style="background:\${color};color:\${color}"></span>
          \${s.area || '—'}
        </span>
      </td>
      <td><span style="font-size:12px;color:var(--text-2);font-weight:500">\${s.region || '—'}</span></td>
      <td><span class="remark-pill \${remarkCls}">\${s.remarks || '—'}</span></td>
    </tr>\`;
  }).join('');
}

// ═══════════════════════ MONTHLY TAB ═══════════════════════════════════════
async function initMonthlyTab() {
  try {
    // Get available months from /api/months
    const res = await fetch('/api/months');
    const json = await res.json();
    if (!json.success) throw new Error(json.error);

    mMonthFilters.months = json.months || [];

    // Populate month dropdown
    const mSel = document.getElementById('mMonthFilter');
    mSel.innerHTML = '';
    mMonthFilters.months.forEach(m => {
      const o = document.createElement('option');
      o.value = m.value;
      o.textContent = m.label;
      mSel.appendChild(o);
    });

    // Default to the latest month
    if (mMonthFilters.months.length) {
      mSel.value = mMonthFilters.months[mMonthFilters.months.length - 1].value;
    }

    // Populate area dropdown (reuse master list from daily tab)
    const aSel = document.getElementById('mAreaFilter');
    aSel.innerHTML = '<option value="ALL">All Areas</option>';
    allFilters.areas.forEach(a => {
      const o = document.createElement('option');
      o.value = o.textContent = a;
      aSel.appendChild(o);
    });

    populateMStores(allFilters.stores);
    await applyMonthlyFilters();
  } catch(e) {
    console.error('Monthly init error:', e);
  }
}

function populateMStores(stores) {
  const sel = document.getElementById('mStoreFilter');
  sel.innerHTML = '<option value="ALL">All Stores</option>';
  stores.forEach(s => {
    const o = document.createElement('option');
    o.value = o.textContent = s;
    sel.appendChild(o);
  });
}

async function onMAreaChange() {
  const area = document.getElementById('mAreaFilter').value;
  if (area === 'ALL') populateMStores(allFilters.stores);
  else {
    try {
      const res = await fetch(\`/api/filters?area=\${encodeURIComponent(area)}\`);
      const json = await res.json();
      if (json.success) populateMStores(json.stores);
    } catch(e){}
  }
  document.getElementById('mStoreFilter').value = 'ALL';
  await applyMonthlyFilters();
}

function setSignFilter(sign) {
  mSignFilter = sign;
  document.querySelectorAll('#signToggle button').forEach(b => {
    b.classList.toggle('active', b.dataset.sign === sign);
  });
  applyMonthlyFilters();
}

async function applyMonthlyFilters() {
  const month = document.getElementById('mMonthFilter').value;
  const area  = document.getElementById('mAreaFilter').value;
  const store = document.getElementById('mStoreFilter').value;

  if (!month) return;

  const params = new URLSearchParams();
  params.set('month', month);
  if (area !== 'ALL') params.set('area', area);
  if (store !== 'ALL') params.set('store', store);
  if (mSignFilter !== 'ALL') params.set('sign', mSignFilter);

  try {
    const res  = await fetch('/api/monthly?' + params);
    const json = await res.json();
    if (!json.success) throw new Error(json.error);

    const summary = json.summary || [];
    const detail  = json.detail  || [];
    const trend   = json.trend   || [];

    mDetailRowsCache = detail;

    document.getElementById('mRecordsCount').innerHTML =
      \`<span>\${summary.length}</span> store\${summary.length!==1?'s':''} · <span>\${detail.length}</span> detail row\${detail.length!==1?'s':''}\`;

    const monthLabel = (mMonthFilters.months.find(m => m.value === month) || {}).label || month;
    document.getElementById('mTableDate').textContent = monthLabel;

    renderMKPIs(summary);
    renderMSummary(summary);
    renderMDetail(detail);
    renderMCharts(summary, trend);
  } catch(e) {
    console.error(e);
    document.getElementById('mSummaryBody').innerHTML =
      \`<tr><td colspan="7" class="empty-cell"><div class="empty-icon" style="background:linear-gradient(135deg,rgba(244,63,94,0.1),rgba(251,113,133,0.1));color:#fb7185"><i class="fa fa-triangle-exclamation"></i></div><p>\${e.message}</p></td></tr>\`;
  }
}

function renderMKPIs(rows) {
  const totSales = rows.reduce((s,r)=>s+(r.sales||0),0);
  const totLY = rows.reduce((s,r)=>s+(r.salesLY||0),0);
  const totDiff = totSales - totLY;
  const totPct = totLY !== 0 ? (totDiff/totLY*100) : 0;
  const active = rows.filter(r=>r.sales>0).length;

  document.getElementById('m-kpi-sales').textContent = fmt(totSales);
  document.getElementById('m-kpi-ly').textContent = fmt(totLY);

  const pctEl = document.getElementById('m-kpi-pct');
  pctEl.textContent = (totPct>=0?'+':'') + totPct.toFixed(2) + '%';
  pctEl.style.background = \`linear-gradient(135deg, \${totPct>=0?'#10b981':'#f43f5e'} 0%, \${totPct>=0?'#34d399':'#fb7185'} 100%)\`;
  pctEl.style.webkitBackgroundClip = 'text';
  pctEl.style.backgroundClip = 'text';
  pctEl.style.webkitTextFillColor = 'transparent';
  document.getElementById('m-kpi-pct-sub').innerHTML = \`<i class="fa fa-arrow-\${totPct>=0?'up':'down'}" style="color:\${totPct>=0?'#34d399':'#fb7185'}"></i> vs last year\`;
  document.getElementById('m-kpi-pct-icon').style.background = \`linear-gradient(135deg, \${totPct>=0?'#10b981':'#f43f5e'}, \${totPct>=0?'#34d399':'#fb7185'})\`;

  const diffEl = document.getElementById('m-kpi-diff');
  diffEl.textContent = (totDiff>=0?'+':'') + fmt(totDiff);
  diffEl.style.background = \`linear-gradient(135deg, \${totDiff>=0?'#06b6d4':'#f43f5e'} 0%, \${totDiff>=0?'#22d3ee':'#fb7185'} 100%)\`;
  diffEl.style.webkitBackgroundClip = 'text';
  diffEl.style.backgroundClip = 'text';
  diffEl.style.webkitTextFillColor = 'transparent';
  document.getElementById('m-kpi-diff-sub').innerHTML = \`<i class="fa fa-\${totDiff>=0?'circle-check':'circle-xmark'}" style="color:\${totDiff>=0?'#22d3ee':'#fb7185'}"></i> variance\`;
  document.getElementById('m-kpi-diff-icon').style.background = \`linear-gradient(135deg, \${totDiff>=0?'#06b6d4':'#f43f5e'}, \${totDiff>=0?'#22d3ee':'#fb7185'})\`;

  document.getElementById('m-kpi-stores').textContent = active;
}

function renderMSummary(rows) {
  const tbody = document.getElementById('mSummaryBody');
  if (!rows.length) {
    tbody.innerHTML = \`<tr><td colspan="7" class="empty-cell"><div class="empty-icon"><i class="fa fa-magnifying-glass"></i></div><p>No data for selected month</p></td></tr>\`;
    return;
  }

  const totSales = rows.reduce((s,r)=>s+r.sales,0);
  const totLY = rows.reduce((s,r)=>s+r.salesLY,0);
  const totDiff = totSales - totLY;
  const totPct = totLY ? (totDiff/totLY*100) : 0;

  let html = rows.map(r => {
    const pctCls = r.diffPct > 0.05 ? 'up' : r.diffPct < -0.05 ? 'down' : 'flat';
    const color = AREA_COLORS[r.area] || DEFAULT_COLOR;
    const grad = AREA_GRADIENTS[r.area] || [DEFAULT_COLOR, DEFAULT_COLOR];
    const diffColor = r.diffVal >= 0 ? '#34d399' : '#fb7185';
    const arrow = r.diffPct > 0.05 ? '↑' : r.diffPct < -0.05 ? '↓' : '—';
    const pctStr = r.diffPct !== 0 ? \`\${arrow} \${Math.abs(r.diffPct).toFixed(2)}%\` : '—';
    return \`<tr>
      <td><span class="num num-bold" style="color:var(--text-1)">#\${r.storeId}</span></td>
      <td>
        <div class="store-cell">
          <div class="store-avatar" style="background:linear-gradient(135deg, \${grad[0]}, \${grad[1]})">\${initials(r.storeName)}</div>
          <div class="store-info">
            <div class="store-name">\${r.storeName || '—'}</div>
          </div>
        </div>
      </td>
      <td><span class="area-tag"><span class="area-dot" style="background:\${color};color:\${color}"></span>\${r.area || '—'}</span></td>
      <td style="text-align:right"><span class="num num-bold" style="color:var(--text-1)">\${fmtFull(r.sales)}</span></td>
      <td style="text-align:right"><span class="num" style="color:var(--text-3)">\${fmtFull(r.salesLY)}</span></td>
      <td style="text-align:center"><span class="pill \${pctCls}">\${pctStr}</span></td>
      <td style="text-align:right"><span class="num num-bold" style="color:\${diffColor}">\${r.diffVal>=0?'+':''}\${fmtFull(r.diffVal)}</span></td>
    </tr>\`;
  }).join('');

  const totPctCls = totPct > 0.05 ? 'up' : totPct < -0.05 ? 'down' : 'flat';
  const totArrow = totPct > 0.05 ? '↑' : totPct < -0.05 ? '↓' : '—';
  html += \`<tr class="summary-row">
    <td colspan="3" style="font-size:12px;letter-spacing:0.1em;text-transform:uppercase;color:var(--indigo2)">TOTAL · \${rows.length} STORES</td>
    <td style="text-align:right"><span class="num">\${fmtFull(totSales)}</span></td>
    <td style="text-align:right"><span class="num" style="color:var(--text-3)">\${fmtFull(totLY)}</span></td>
    <td style="text-align:center"><span class="pill \${totPctCls}">\${totArrow} \${Math.abs(totPct).toFixed(2)}%</span></td>
    <td style="text-align:right"><span class="num" style="color:\${totDiff>=0?'#34d399':'#fb7185'}">\${totDiff>=0?'+':''}\${fmtFull(totDiff)}</span></td>
  </tr>\`;
  tbody.innerHTML = html;
}

function renderMDetail(rows) {
  const tbody = document.getElementById('mDetailBody');
  if (!rows.length) {
    tbody.innerHTML = \`<tr><td colspan="9" class="empty-cell"><div class="empty-icon"><i class="fa fa-magnifying-glass"></i></div><p>No detail records found</p><small>Try adjusting your filters</small></td></tr>\`;
    return;
  }

  const html = rows.map(r => {
    const pctCls = r.diffPct > 0.05 ? 'up' : r.diffPct < -0.05 ? 'down' : 'flat';
    const grad = AREA_GRADIENTS[r.area] || [DEFAULT_COLOR, DEFAULT_COLOR];
    const diffColor = r.diffVal >= 0 ? '#34d399' : '#fb7185';
    const arrow = r.diffPct > 0.05 ? '↑' : r.diffPct < -0.05 ? '↓' : '—';
    const pctStr = r.diffPct !== 0 ? \`\${arrow} \${Math.abs(r.diffPct).toFixed(2)}%\` : '—';
    const just = (r.justification || '—').replace(/</g,'&lt;').replace(/>/g,'&gt;');

    // Format date as M/D/YYYY
    const dateStr = r.date
      ? new Date(r.date+'T00:00:00').toLocaleDateString('en-US', {month:'numeric',day:'numeric',year:'numeric'})
      : '—';

    return \`<tr>
      <td><span class="num" style="color:var(--text-1);font-weight:600">\${dateStr}</span></td>
      <td><span style="font-size:12px;color:var(--text-2);font-weight:500">\${r.day || '—'}</span></td>
      <td><span style="font-size:12px;color:var(--text-3);font-weight:500">\${r.dayYA || '—'}</span></td>
      <td>
        <div class="store-cell">
          <div class="store-avatar" style="background:linear-gradient(135deg, \${grad[0]}, \${grad[1]});width:30px;height:30px;font-size:10.5px">\${initials(r.storeName)}</div>
          <div class="store-info">
            <div class="store-name" style="font-size:12.5px">\${r.storeName || '—'}</div>
            <div class="store-id">#\${r.storeId}</div>
          </div>
        </div>
      </td>
      <td style="text-align:right"><span class="num num-bold" style="color:var(--text-1)">\${fmtFull(r.sales)}</span></td>
      <td style="text-align:right"><span class="num" style="color:var(--text-3)">\${fmtFull(r.salesLY)}</span></td>
      <td style="text-align:center"><span class="pill \${pctCls}">\${pctStr}</span></td>
      <td style="text-align:right"><span class="num num-bold" style="color:\${diffColor}">\${r.diffVal>=0?'+':''}\${fmtFull(r.diffVal)}</span></td>
      <td><div class="just-full">\${just}</div></td>
    </tr>\`;
  }).join('');

  tbody.innerHTML = html;
}

function renderMCharts(summary, trend) {
  // Bar chart - top stores
  const top = summary.slice(0, 14);
  const labels = top.map(r => r.storeName);
  const salesArr = top.map(r => r.sales);
  const lyArr = top.map(r => r.salesLY);

  if (mBarChartInst) mBarChartInst.destroy();
  const barCanvas = document.getElementById('mBarChart');
  const barCtx = barCanvas.getContext('2d');
  const gradients = top.map(r => {
    const [c1, c2] = AREA_GRADIENTS[r.area] || [DEFAULT_COLOR, DEFAULT_COLOR];
    const g = barCtx.createLinearGradient(0, 0, 0, 280);
    g.addColorStop(0, c2); g.addColorStop(1, c1);
    return g;
  });

  mBarChartInst = new Chart(barCanvas, {
    type: 'bar',
    data: {
      labels,
      datasets: [
        { label: 'Sales', data: salesArr, backgroundColor: gradients, borderRadius: 6, borderSkipped: false, barPercentage: 0.7 },
        { label: 'Sales LY', data: lyArr, backgroundColor: 'rgba(74,83,107,0.4)', borderColor: 'rgba(148,163,200,0.2)', borderWidth: 1, borderRadius: 6, borderSkipped: false, barPercentage: 0.7 }
      ]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          backgroundColor: 'rgba(15,20,35,0.95)', borderColor: 'rgba(99,102,241,0.4)', borderWidth: 1,
          padding: 12, cornerRadius: 10,
          titleFont: { family: "'Space Grotesk'", size: 13, weight: '600' },
          bodyFont:  { family: "'JetBrains Mono'", size: 11.5 },
          titleColor: '#f0f3fb', bodyColor: '#a8b3d1',
          callbacks: { label: ctx => \` \${ctx.dataset.label}: ₱\${ctx.raw.toLocaleString('en-PH',{minimumFractionDigits:2})}\` }
        }
      },
      scales: {
        x: { grid: { color: 'rgba(148,163,200,0.04)', drawBorder: false }, ticks: { color: '#a8b3d1', font: { size: 10, family: "'Inter'", weight: '500' }, maxRotation: 40 } },
        y: { grid: { color: 'rgba(148,163,200,0.06)' }, ticks: { color: '#a8b3d1', font: { size: 10.5, family: "'JetBrains Mono'", weight: '500' }, callback: v => v >= 1e6 ? '₱' + (v/1e6).toFixed(1)+'M' : v >= 1e3 ? '₱' + (v/1e3).toFixed(0)+'K' : v } }
      }
    }
  });

  // Line chart - daily trend
  if (mLineChartInst) mLineChartInst.destroy();
  const lineCanvas = document.getElementById('mLineChart');
  const lineCtx = lineCanvas.getContext('2d');

  const trendLabels = trend.map(t => {
    const d = new Date(t.date+'T00:00:00');
    return d.toLocaleDateString('en-US', { month:'short', day:'numeric' });
  });
  const trendSales = trend.map(t => t.sales);
  const trendLY = trend.map(t => t.salesLY);

  const grad1 = lineCtx.createLinearGradient(0, 0, 0, 320);
  grad1.addColorStop(0, 'rgba(99,102,241,0.4)');
  grad1.addColorStop(1, 'rgba(99,102,241,0.01)');
  const grad2 = lineCtx.createLinearGradient(0, 0, 0, 320);
  grad2.addColorStop(0, 'rgba(168,85,247,0.25)');
  grad2.addColorStop(1, 'rgba(168,85,247,0.01)');

  mLineChartInst = new Chart(lineCanvas, {
    type: 'line',
    data: {
      labels: trendLabels,
      datasets: [
        { label: 'Sales', data: trendSales, borderColor: '#6366f1', backgroundColor: grad1, fill: true, tension: 0.4, borderWidth: 2.5, pointRadius: 0, pointHoverRadius: 5, pointBackgroundColor: '#818cf8', pointBorderColor: '#fff', pointBorderWidth: 2 },
        { label: 'Sales LY', data: trendLY, borderColor: '#a855f7', backgroundColor: grad2, fill: true, tension: 0.4, borderWidth: 2, borderDash: [5,4], pointRadius: 0, pointHoverRadius: 5, pointBackgroundColor: '#c084fc', pointBorderColor: '#fff', pointBorderWidth: 2 }
      ]
    },
    options: {
      responsive: true, maintainAspectRatio: false, interaction: { intersect: false, mode: 'index' },
      plugins: {
        legend: { position: 'bottom', labels: { color: '#a8b3d1', font: { size: 11.5, family: "'Inter'", weight: '500' }, padding: 14, boxWidth: 11, boxHeight: 11, usePointStyle: true, pointStyle: 'rectRounded' } },
        tooltip: {
          backgroundColor: 'rgba(15,20,35,0.95)', borderColor: 'rgba(99,102,241,0.4)', borderWidth: 1,
          padding: 12, cornerRadius: 10,
          titleFont: { family: "'Space Grotesk'", weight: '600', size: 13 },
          bodyFont:  { family: "'JetBrains Mono'", size: 11.5 },
          titleColor: '#f0f3fb', bodyColor: '#a8b3d1',
          callbacks: { label: ctx => \` \${ctx.dataset.label}: ₱\${ctx.raw.toLocaleString('en-PH',{minimumFractionDigits:2})}\` }
        }
      },
      scales: {
        x: { grid: { color: 'rgba(148,163,200,0.04)', drawBorder: false }, ticks: { color: '#a8b3d1', font: { size: 10, family: "'Inter'", weight: '500' }, maxRotation: 0, autoSkip: true, maxTicksLimit: 12 } },
        y: { grid: { color: 'rgba(148,163,200,0.06)' }, ticks: { color: '#a8b3d1', font: { size: 10.5, family: "'JetBrains Mono'", weight: '500' }, callback: v => v >= 1e6 ? '₱' + (v/1e6).toFixed(1)+'M' : v >= 1e3 ? '₱' + (v/1e3).toFixed(0)+'K' : v } }
      }
    }
  });
}

// ─── Export Detail Report to Excel ──────────────────────────────────────────
function exportDetailToExcel() {
  if (!mDetailRowsCache.length) {
    alert('No data to export with the current filters.');
    return;
  }

  const month = document.getElementById('mMonthFilter').value;
  const area  = document.getElementById('mAreaFilter').value;
  const store = document.getElementById('mStoreFilter').value;
  const monthLabel = (mMonthFilters.months.find(m => m.value === month) || {}).label || month;

  // Prepare data rows
  const data = mDetailRowsCache.map(r => ({
    'Date': r.date ? new Date(r.date+'T00:00:00').toLocaleDateString('en-US') : '',
    'Day': r.day || '',
    'Day YA': r.dayYA || '',
    'Store ID': r.storeId || '',
    'Store Name': r.storeName || '',
    'Area': r.area || '',
    'Sales': r.sales || 0,
    'Sales LY': r.salesLY || 0,
    'Diff %': r.diffPct || 0,
    'Diff Amount': r.diffVal || 0,
    'Justification': r.justification || ''
  }));

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(data);

  // Set column widths
  ws['!cols'] = [
    { wch: 12 }, // Date
    { wch: 12 }, // Day
    { wch: 12 }, // Day YA
    { wch: 10 }, // Store ID
    { wch: 22 }, // Store Name
    { wch: 18 }, // Area
    { wch: 14 }, // Sales
    { wch: 14 }, // Sales LY
    { wch: 10 }, // Diff %
    { wch: 14 }, // Diff Amount
    { wch: 60 }, // Justification
  ];

  // Style header row (dark green like booking app)
  const headerStyle = {
    fill: { fgColor: { rgb: '166534' } },
    font: { color: { rgb: 'FFFFFF' }, bold: true },
    alignment: { horizontal: 'center', vertical: 'center' }
  };
  const range = XLSX.utils.decode_range(ws['!ref']);
  for (let C = range.s.c; C <= range.e.c; C++) {
    const addr = XLSX.utils.encode_cell({ r: 0, c: C });
    if (ws[addr]) ws[addr].s = headerStyle;
  }

  // Format number columns
  const numCols = { 6: '#,##0.00', 7: '#,##0.00', 8: '0.00"%"', 9: '#,##0.00' };
  for (let R = 1; R <= range.e.r; R++) {
    for (const [col, fmt] of Object.entries(numCols)) {
      const addr = XLSX.utils.encode_cell({ r: R, c: parseInt(col) });
      if (ws[addr]) ws[addr].z = fmt;
    }
  }

  XLSX.utils.book_append_sheet(wb, ws, 'Detail Report');

  // Filename
  const safeMonth = monthLabel.replace(/[^a-z0-9]/gi, '_');
  const tag = mSignFilter === 'POS' ? '_Positive' : mSignFilter === 'NEG' ? '_Negative' : '';
  const areaTag = area !== 'ALL' ? '_' + area.replace(/[^a-z0-9]/gi, '_') : '';
  const storeTag = store !== 'ALL' ? '_' + store.replace(/[^a-z0-9]/gi, '_') : '';
  const filename = \`CaMaNaVa_Monthly_\${safeMonth}\${areaTag}\${storeTag}\${tag}.xlsx\`;

  XLSX.writeFile(wb, filename);
}

function setStatus(type, text) {
  const b = document.getElementById('statusBadge');
  b.className = 'badge ' + (type==='live' ? 'live' : type==='error' ? 'error' : 'loading-badge');
  b.innerHTML = type==='live' ? \`<span class="pulse"></span> \${text}\` : type==='error' ? \`<i class="fa fa-circle-exclamation"></i> \${text}\` : \`<span class="pulse"></span> \${text}\`;
}

function setSyncing(on) {
  const btn = document.getElementById('syncBtn');
  const ico = document.getElementById('syncIcon');
  btn.disabled = on;
  if (on) ico.style.animation = 'spin 0.8s linear infinite';
  else ico.style.animation = '';
}

function hideOverlay() { document.getElementById('loadingOverlay').style.display = 'none'; }
function showTableError(msg) {
  document.getElementById('tableBody').innerHTML = \`<tr><td colspan="7" class="empty-cell"><div class="empty-icon" style="background:linear-gradient(135deg,rgba(244,63,94,0.1),rgba(251,113,133,0.1));color:#fb7185"><i class="fa fa-triangle-exclamation"></i></div><p>\${msg}</p></td></tr>\`;
}

loadFilters();
</script>
</body>
</html>
`;

app.get('/', (req, res) => { res.setHeader('Content-Type', 'text/html; charset=utf-8'); res.send(HTML); });
app.get('*', (req, res) => { res.setHeader('Content-Type', 'text/html; charset=utf-8'); res.send(HTML); });

app.listen(PORT, () => console.log(`CaMaNaVa eBRT running on port ${PORT}`));
