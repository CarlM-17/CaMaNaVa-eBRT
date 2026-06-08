const express = require('express');
const { google } = require('googleapis');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json());

const SHEET_ID = process.env.SHEET_ID || '1PhRyLx2viByS1J-dOWWPwSmAZwNwzzaM6ATmbpQIR8w';
const SHEET_NAME = 'DailySales';
const DATA_RANGE = `${SHEET_NAME}!A23:R`;
const STORE_LIST_RANGE = 'ListOfStores!A:E';
const CATEGORY_RANGE = 'CategorySales!A:I';

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
let categoryCache = null;
let categoryCacheTime = 0;
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

async function getCategoryData(sheets) {
  const now = Date.now();
  if (categoryCache && (now - categoryCacheTime) < CACHE_TTL) return categoryCache;
  const response = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: CATEGORY_RANGE });
  const rows = response.data.values || [];

  // Skip the first row (headers). Then map rows.
  // Columns: A=Month, B=Area, C=Store Code, D=Store Name, E=SDep Code, F=Sub-Department Name, G=Sales, H=SalesYA, I=Category
  const data = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i] || [];
    const month     = (r[0] || '').trim();
    const area      = (r[1] || '').trim();
    const storeCode = (r[2] || '').trim();
    const storeName = (r[3] || '').trim();
    const sdepCode  = (r[4] || '').trim();
    const subDepName= (r[5] || '').trim();
    const sales     = parseNum(r[6]);
    const salesLY   = parseNum(r[7]);
    const category  = (r[8] || '').trim();
    // skip blank rows
    if (!month && !storeCode && !category && !subDepName) continue;
    data.push({ month, area, storeCode, storeName, sdepCode, subDepName, sales, salesLY, category });
  }
  categoryCache = data;
  categoryCacheTime = now;
  return categoryCache;
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

    let filtered = data;
    if (month) filtered = filtered.filter(r => monthKey(r.date) === month);
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

// ─── API: GET /api/category-filters ─────────────────────────────────────────
app.get('/api/category-filters', async (req, res) => {
  try {
    const auth = getAuthClient();
    const sheets = google.sheets({ version: 'v4', auth });
    const data = await getCategoryData(sheets);
    const { area: filterArea } = req.query;

    // Preserve month order from sheet (it's typically in calendar order)
    const monthSet = [];
    const seenMonths = new Set();
    data.forEach(r => { if (r.month && !seenMonths.has(r.month)) { seenMonths.add(r.month); monthSet.push(r.month); } });

    const categories = new Set();
    const areas = new Set();
    const stores = new Set();
    data.forEach(r => {
      if (r.category) categories.add(r.category);
      if (r.area) areas.add(r.area);
      if (r.storeName) {
        if (!filterArea || filterArea === 'ALL' || r.area === filterArea) stores.add(r.storeName);
      }
    });
    res.json({
      success: true,
      months: monthSet,
      categories: [...categories].sort(),
      areas: [...areas].sort(),
      stores: [...stores].sort(),
    });
  } catch (err) {
    console.error('Category-filters error:', err.message);
    res.status(500).json({ success: false, error: err.message });
  }
});

// ─── API: GET /api/category ─────────────────────────────────────────────────
app.get('/api/category', async (req, res) => {
  try {
    const auth = getAuthClient();
    const sheets = google.sheets({ version: 'v4', auth });
    const data = await getCategoryData(sheets);

    const { month, category, area, store, sign } = req.query;
    let filtered = data;
    if (month && month !== 'ALL') filtered = filtered.filter(r => r.month === month);
    if (category && category !== 'ALL') filtered = filtered.filter(r => r.category === category);
    if (area && area !== 'ALL') filtered = filtered.filter(r => r.area === area);
    if (store && store !== 'ALL') filtered = filtered.filter(r => r.storeName === store);

    // ── Category Summary ───────────────────────────────────────────────────
    const catMap = {};
    filtered.forEach(r => {
      const cat = r.category || '(uncategorized)';
      if (!catMap[cat]) catMap[cat] = { category: cat, sales: 0, salesLY: 0, subDeps: new Set() };
      catMap[cat].sales   += r.sales;
      catMap[cat].salesLY += r.salesLY;
      if (r.subDepName) catMap[cat].subDeps.add(r.subDepName);
    });
    const totalSales = Object.values(catMap).reduce((s,c)=>s+c.sales,0);
    const summary = Object.values(catMap).map(c => {
      const diffVal = c.sales - c.salesLY;
      const diffPct = c.salesLY !== 0 ? (diffVal / c.salesLY) * 100 : 0;
      const shareCur = totalSales !== 0 ? (c.sales / totalSales) * 100 : 0;
      return {
        category: c.category,
        subDepCount: c.subDeps.size,
        sales: c.sales,
        salesLY: c.salesLY,
        diffVal,
        diffPct: parseFloat(diffPct.toFixed(2)),
        shareCur: parseFloat(shareCur.toFixed(2)),
      };
    }).sort((a, b) => b.sales - a.sales);

    // ── By Area ────────────────────────────────────────────────────────────
    const areaMap = {};
    filtered.forEach(r => {
      if (!r.area) return;
      if (!areaMap[r.area]) areaMap[r.area] = { area: r.area, sales: 0, salesLY: 0 };
      areaMap[r.area].sales   += r.sales;
      areaMap[r.area].salesLY += r.salesLY;
    });
    const byArea = Object.values(areaMap).map(a => {
      const diffVal = a.sales - a.salesLY;
      const diffPct = a.salesLY !== 0 ? (diffVal / a.salesLY) * 100 : 0;
      return { ...a, diffVal, diffPct: parseFloat(diffPct.toFixed(2)) };
    }).sort((a, b) => b.sales - a.sales);

    // ── By Store ───────────────────────────────────────────────────────────
    const storeMap = {};
    filtered.forEach(r => {
      if (!r.storeCode) return;
      const key = r.storeCode + '_' + r.storeName;
      if (!storeMap[key]) storeMap[key] = { storeCode: r.storeCode, storeName: r.storeName, area: r.area, sales: 0, salesLY: 0 };
      storeMap[key].sales   += r.sales;
      storeMap[key].salesLY += r.salesLY;
    });
    const byStore = Object.values(storeMap).map(s => {
      const diffVal = s.sales - s.salesLY;
      const diffPct = s.salesLY !== 0 ? (diffVal / s.salesLY) * 100 : 0;
      return { ...s, diffVal, diffPct: parseFloat(diffPct.toFixed(2)) };
    }).sort((a, b) => b.sales - a.sales);

    // ── SubDeps grouped by category (for breakdown dropdown) ───────────────
    const subDepsByCategory = {};
    filtered.forEach(r => {
      if (!r.category || !r.subDepName) return;
      if (!subDepsByCategory[r.category]) subDepsByCategory[r.category] = new Set();
      subDepsByCategory[r.category].add(r.subDepName);
    });
    const subDepsByCategoryOut = {};
    Object.entries(subDepsByCategory).forEach(([k,v]) => { subDepsByCategoryOut[k] = [...v].sort(); });

    // ── Sub-Department Detail ──────────────────────────────────────────────
    // Aggregate by (category, sdepCode, subDepName) across the filter scope
    const subMap = {};
    filtered.forEach(r => {
      const key = (r.category || '') + '||' + (r.sdepCode || '') + '||' + (r.subDepName || '');
      if (!subMap[key]) subMap[key] = { category: r.category, sdepCode: r.sdepCode, subDepName: r.subDepName, sales: 0, salesLY: 0 };
      subMap[key].sales   += r.sales;
      subMap[key].salesLY += r.salesLY;
    });
    let detail = Object.values(subMap).map(s => {
      const diffVal = s.sales - s.salesLY;
      const diffPct = s.salesLY !== 0 ? (diffVal / s.salesLY) * 100 : 0;
      return { ...s, diffVal, diffPct: parseFloat(diffPct.toFixed(2)) };
    });

    // Movers calculated BEFORE sign filter (so users see top/bottom regardless of toggle)
    const moversSrc = [...detail].sort((a, b) => b.diffVal - a.diffVal);
    const movers = {
      top: moversSrc.filter(r => r.diffVal > 0).slice(0, 10),
      bottom: moversSrc.filter(r => r.diffVal < 0).slice(-10).reverse(),
    };

    if (sign === 'POS') detail = detail.filter(r => r.diffVal > 0);
    if (sign === 'NEG') detail = detail.filter(r => r.diffVal < 0);
    detail.sort((a, b) => b.sales - a.sales);

    res.json({ success: true, summary, detail, byArea, byStore, subDepsByCategory: subDepsByCategoryOut, movers });
  } catch (err) {
    console.error('Category error:', err.message);
    res.status(500).json({ success: false, error: err.message });
  }
});

// ─── API: GET /api/category-breakdown ───────────────────────────────────────
// Returns byArea + byStore tables with extra Category & Sub-Department filters
app.get('/api/category-breakdown', async (req, res) => {
  try {
    const auth = getAuthClient();
    const sheets = google.sheets({ version: 'v4', auth });
    const data = await getCategoryData(sheets);

    const { month, category, area, store, breakdownCategory, breakdownSubDep } = req.query;
    let filtered = data;
    if (month && month !== 'ALL') filtered = filtered.filter(r => r.month === month);
    if (category && category !== 'ALL') filtered = filtered.filter(r => r.category === category);
    if (area && area !== 'ALL') filtered = filtered.filter(r => r.area === area);
    if (store && store !== 'ALL') filtered = filtered.filter(r => r.storeName === store);
    if (breakdownCategory && breakdownCategory !== 'ALL') filtered = filtered.filter(r => r.category === breakdownCategory);
    if (breakdownSubDep && breakdownSubDep !== 'ALL') filtered = filtered.filter(r => r.subDepName === breakdownSubDep);

    // By Area
    const areaMap = {};
    filtered.forEach(r => {
      if (!r.area) return;
      if (!areaMap[r.area]) areaMap[r.area] = { area: r.area, sales: 0, salesLY: 0 };
      areaMap[r.area].sales   += r.sales;
      areaMap[r.area].salesLY += r.salesLY;
    });
    const byArea = Object.values(areaMap).map(a => {
      const diffVal = a.sales - a.salesLY;
      const diffPct = a.salesLY !== 0 ? (diffVal / a.salesLY) * 100 : 0;
      return { ...a, diffVal, diffPct: parseFloat(diffPct.toFixed(2)) };
    }).sort((a, b) => b.sales - a.sales);

    // By Store
    const storeMap = {};
    filtered.forEach(r => {
      if (!r.storeCode) return;
      const key = r.storeCode + '_' + r.storeName;
      if (!storeMap[key]) storeMap[key] = { storeCode: r.storeCode, storeName: r.storeName, area: r.area, sales: 0, salesLY: 0 };
      storeMap[key].sales   += r.sales;
      storeMap[key].salesLY += r.salesLY;
    });
    const byStore = Object.values(storeMap).map(s => {
      const diffVal = s.sales - s.salesLY;
      const diffPct = s.salesLY !== 0 ? (diffVal / s.salesLY) * 100 : 0;
      return { ...s, diffVal, diffPct: parseFloat(diffPct.toFixed(2)) };
    }).sort((a, b) => b.sales - a.sales);

    res.json({ success: true, byArea, byStore });
  } catch (err) {
    console.error('Category-breakdown error:', err.message);
    res.status(500).json({ success: false, error: err.message });
  }
});

// ─── API: GET /api/averages ─────────────────────────────────────────────────
// Returns average daily sales per store and per area
// Excludes records with blank sales (= 0) and any record dated January 1
app.get('/api/averages', async (req, res) => {
  try {
    const auth = getAuthClient();
    const sheets = google.sheets({ version: 'v4', auth });
    const data = await getSalesData(sheets);

    const { area, store } = req.query;

    // Filter: exclude January 1 records and blank sales
    let filtered = data.filter(r => {
      if (!r.date || !r.sales || r.sales <= 0) return false;
      // r.date is YYYY-MM-DD; check month-day != 01-01
      const mmdd = r.date.substring(5);
      if (mmdd === '01-01') return false;
      return true;
    });

    if (area && area !== 'ALL') filtered = filtered.filter(r => r.area === area);
    if (store && store !== 'ALL') filtered = filtered.filter(r => r.storeName === store);

    // Per-store averages: sum / number of distinct days that store had data
    const storeStats = {};
    filtered.forEach(r => {
      const key = r.storeId + '_' + r.storeName;
      if (!storeStats[key]) {
        storeStats[key] = {
          storeId: r.storeId, storeName: r.storeName, area: r.area,
          total: 0, days: new Set(),
        };
      }
      storeStats[key].total += r.sales;
      storeStats[key].days.add(r.date);
    });
    const perStore = Object.values(storeStats).map(s => ({
      storeId: s.storeId, storeName: s.storeName, area: s.area,
      avg: s.days.size ? s.total / s.days.size : 0,
      total: s.total, dayCount: s.days.size,
    })).sort((a, b) => b.avg - a.avg);

    // Per-area averages: sum / number of distinct (date, store) combos
    const areaStats = {};
    filtered.forEach(r => {
      if (!areaStats[r.area]) {
        areaStats[r.area] = { area: r.area, total: 0, count: 0, days: new Set() };
      }
      areaStats[r.area].total += r.sales;
      areaStats[r.area].count += 1;
      areaStats[r.area].days.add(r.date);
    });
    const perArea = Object.values(areaStats).map(a => ({
      area: a.area,
      avgPerRecord: a.count ? a.total / a.count : 0,
      avgPerDay: a.days.size ? a.total / a.days.size : 0,
      total: a.total, recordCount: a.count, dayCount: a.days.size,
    })).sort((a, b) => b.avgPerDay - a.avgPerDay);

    res.json({ success: true, perStore, perArea, totalRecords: filtered.length });
  } catch (err) {
    console.error('Averages error:', err.message);
    res.status(500).json({ success: false, error: err.message });
  }
});

const HTML = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=5,viewport-fit=cover"/>
<meta name="theme-color" content="#0a0e1a"/>
<title>CaMaNaVa eBRT — Daily Sales Report</title>
<link rel="preconnect" href="https://fonts.googleapis.com"/>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;600;700&family=Inter:wght@300;400;500;600&family=JetBrains+Mono:wght@400;500;600&display=swap"/>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css"/>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js"></script>
<script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0/dist/chartjs-plugin-datalabels.min.js"></script>
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
th.sortable{
  cursor:pointer;user-select:none;
  transition:color .15s, background .15s;
  position:sticky;top:0;z-index:5;
}
th.sortable:hover{color:var(--indigo2);background:rgba(99,102,241,0.06)}
th.sortable .sort-icon{
  display:inline-block;margin-left:5px;font-size:9px;
  opacity:0.35;transition:opacity .15s;color:var(--text-3);
}
th.sortable:hover .sort-icon{opacity:0.6}
th.sortable.sort-asc .sort-icon,
th.sortable.sort-desc .sort-icon{opacity:1;color:var(--indigo2)}
th.sortable.sort-asc,
th.sortable.sort-desc{color:var(--indigo2)}
td{
  padding:13px 16px;border-bottom:1px solid var(--border-soft);
  vertical-align:middle;white-space:nowrap;
  font-size:13px;
}
td:has(.just-full){white-space:normal}
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
.tab-content{display:none !important}
.tab-content.active{display:block !important}
.tab-content.active > *{opacity:1 !important}

/* Style native select options for dark theme */
select option{background:#1a1f2e;color:#e8ecf4}

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

/* ── Mobile performance ── */
@media (max-width: 768px) {
  /* Kill backdrop-filter on mobile — biggest perf win */
  .header, .controls, .kpi, .chart-card, .table-card,
  .missing-card, .tabs, .sync-btn, .badge,
  .tab-content, .loading-overlay {
    backdrop-filter: none !important;
    -webkit-backdrop-filter: none !important;
  }
  /* Solid backgrounds instead of glass */
  .header{background:rgba(22,27,39,0.98) !important}
  .controls{background:rgba(24,30,48,0.98) !important}
  .kpi{background:rgba(24,30,48,0.98) !important}
  .chart-card{background:rgba(24,30,48,0.98) !important}
  .table-card{background:rgba(24,30,48,0.98) !important}
  .missing-card{background:rgba(24,30,48,0.98) !important}
  .tabs{background:rgba(22,27,39,0.98) !important}

  /* Disable animated background on mobile */
  body::before{animation:none !important}
  body::after{display:none !important}

  /* Smaller layout adjustments */
  .main{padding:14px 12px 30px}
  .header{padding:12px 16px}
  .controls{padding:12px 14px;gap:10px}
  .kpi-grid{grid-template-columns:repeat(2,1fr) !important;gap:10px}
  .kpi{padding:14px 14px}
  .kpi-value{font-size:20px}
  .charts-grid{grid-template-columns:1fr !important}
  .tabs{width:100%}
  .tab-btn{flex:1;justify-content:center;padding:9px 10px;font-size:12px}

  /* Table: allow horizontal scroll, compact padding */
  .table-wrap{overflow-x:auto;-webkit-overflow-scrolling:touch}
  td,th{padding:9px 10px;font-size:11.5px}

  /* Reduce logo text */
  .logo-text{font-size:15px}
  .logo-icon{width:36px;height:36px}

  /* Reduce card border-radius */
  .kpi,.chart-card,.table-card,.controls,.tabs,.missing-card{border-radius:12px}

  /* Kill expensive glow/shadow effects on mobile */
  .logo-icon::after{display:none !important}
  .kpi::after{display:none !important}
  .pulse{box-shadow:none !important}
  .pill{box-shadow:none !important}
  .area-dot{box-shadow:none !important}
  .legend-dot{box-shadow:none !important}
  .kpi:hover{transform:none !important}

  /* Detail report justification - constrain width on mobile */
  .just-full{max-width:200px}
}
::-webkit-scrollbar{width:8px;height:8px}
::-webkit-scrollbar-track{background:rgba(15,20,35,0.4)}
::-webkit-scrollbar-thumb{background:linear-gradient(180deg, var(--indigo), #4f46e5);border-radius:4px}
::-webkit-scrollbar-thumb:hover{background:linear-gradient(180deg, var(--indigo2), var(--indigo))}

/* ── Animation ── */
@keyframes fadeInUp{
  from{opacity:0;transform:translateY(8px)}
  to{opacity:1;transform:translateY(0)}
}
/* Animations disabled to prevent visibility issues on tab switch */
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
    <button class="tab-btn" data-tab="category" onclick="switchTab('category')">
      <i class="fa fa-tags"></i> Category Sales
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
            <th class="sortable" data-sort-key="storeName" data-sort-type="string" onclick="sortDaily('storeName','string')">Store <span class="sort-icon">⇅</span></th>
            <th class="sortable" data-sort-key="area" data-sort-type="string" onclick="sortDaily('area','string')">Area <span class="sort-icon">⇅</span></th>
            <th class="sortable" data-sort-key="sales" data-sort-type="num" style="text-align:right" onclick="sortDaily('sales','num')">Sales <span class="sort-icon">⇅</span></th>
            <th class="sortable" data-sort-key="salesLY" data-sort-type="num" style="text-align:right" onclick="sortDaily('salesLY','num')">Sales LY <span class="sort-icon">⇅</span></th>
            <th class="sortable" data-sort-key="diffPct" data-sort-type="num" style="text-align:center" onclick="sortDaily('diffPct','num')">Diff % <span class="sort-icon">⇅</span></th>
            <th class="sortable" data-sort-key="diffVal" data-sort-type="num" style="text-align:right" onclick="sortDaily('diffVal','num')">Diff Amount <span class="sort-icon">⇅</span></th>
            <th class="sortable" data-sort-key="justification" data-sort-type="string" onclick="sortDaily('justification','string')">Justification <span class="sort-icon">⇅</span></th>
          </tr>
        </thead>
        <tbody id="tableBody">
          <tr><td colspan="7" class="empty-cell"><div class="empty-icon"><i class="fa fa-spinner fa-spin"></i></div><p>Loading data...</p></td></tr>
        </tbody>
      </table>
    </div>
  </div><!-- /table-card -->

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

  <!-- Average Daily Sales Charts -->
  <div class="charts-grid" style="margin-top:24px">
    <div class="chart-card">
      <div class="chart-title">
        <i class="fa fa-chart-column"></i> Average Daily Sales by Store
        <span style="margin-left:auto;font-size:10.5px;color:var(--text-3);font-weight:500;text-transform:uppercase;letter-spacing:0.08em">Excl. blanks &amp; Jan 1</span>
      </div>
      <div style="overflow-y:auto;max-height:520px;padding-right:4px">
        <div id="avgStoreChartWrap" style="position:relative;height:600px">
          <canvas id="avgStoreChart" role="img" aria-label="Average daily sales per store"></canvas>
        </div>
      </div>
    </div>
    <div class="chart-card">
      <div class="chart-title">
        <i class="fa fa-chart-simple"></i> Average Daily Sales by Area
        <span style="margin-left:auto;font-size:10.5px;color:var(--text-3);font-weight:500;text-transform:uppercase;letter-spacing:0.08em">Excl. blanks &amp; Jan 1</span>
      </div>
      <div style="position:relative;height:320px">
        <canvas id="avgAreaChart" role="img" aria-label="Average daily sales per area"></canvas>
      </div>
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
        <div class="kpi-sub" id="m-kpi-stores-sub" style="display:flex;gap:10px;flex-wrap:wrap;align-items:center;margin-top:6px">
          <span style="display:inline-flex;align-items:center;gap:4px"><i class="fa fa-arrow-up" style="color:#34d399;font-size:10px"></i> <span id="m-kpi-up" style="color:#34d399;font-weight:600;font-family:'JetBrains Mono',monospace">0</span> growth</span>
          <span style="color:var(--text-dim)">·</span>
          <span style="display:inline-flex;align-items:center;gap:4px"><i class="fa fa-arrow-down" style="color:#fb7185;font-size:10px"></i> <span id="m-kpi-down" style="color:#fb7185;font-weight:600;font-family:'JetBrains Mono',monospace">0</span> decline</span>
        </div>
      </div>
    </div>

    <!-- Charts -->
    <div class="charts-grid">
      <div class="chart-card" style="padding:0;overflow:hidden">
        <div class="table-header" style="border-bottom:1px solid var(--border-soft)">
          <div class="table-title"><i class="fa fa-map-location-dot"></i> Sales per Area</div>
          <div class="table-date" id="mAreaTableInfo">—</div>
        </div>
        <div class="table-wrap" style="max-height:340px">
          <table>
            <thead>
              <tr>
                <th class="sortable" data-sort-key="area" data-sort-type="string" onclick="sortMArea('area','string')">Area <span class="sort-icon">⇅</span></th>
                <th class="sortable" data-sort-key="sales" data-sort-type="num" style="text-align:right" onclick="sortMArea('sales','num')">Sales <span class="sort-icon">⇅</span></th>
                <th class="sortable" data-sort-key="salesLY" data-sort-type="num" style="text-align:right" onclick="sortMArea('salesLY','num')">Sales LY <span class="sort-icon">⇅</span></th>
                <th class="sortable" data-sort-key="diffPct" data-sort-type="num" style="text-align:center" onclick="sortMArea('diffPct','num')">Diff % <span class="sort-icon">⇅</span></th>
                <th class="sortable" data-sort-key="diffVal" data-sort-type="num" style="text-align:right" onclick="sortMArea('diffVal','num')">Diff Amount <span class="sort-icon">⇅</span></th>
              </tr>
            </thead>
            <tbody id="mAreaTableBody">
              <tr><td colspan="5" class="empty-cell"><div class="empty-icon"><i class="fa fa-spinner fa-spin"></i></div><p>Loading...</p></td></tr>
            </tbody>
          </table>
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
              <th class="sortable" data-sort-key="storeId" data-sort-type="num" onclick="sortMSummary('storeId','num')">Store ID <span class="sort-icon">⇅</span></th>
              <th class="sortable" data-sort-key="storeName" data-sort-type="string" onclick="sortMSummary('storeName','string')">Store Name <span class="sort-icon">⇅</span></th>
              <th class="sortable" data-sort-key="area" data-sort-type="string" onclick="sortMSummary('area','string')">Area <span class="sort-icon">⇅</span></th>
              <th class="sortable" data-sort-key="sales" data-sort-type="num" style="text-align:right" onclick="sortMSummary('sales','num')">Sales <span class="sort-icon">⇅</span></th>
              <th class="sortable" data-sort-key="salesLY" data-sort-type="num" style="text-align:right" onclick="sortMSummary('salesLY','num')">Sales LY <span class="sort-icon">⇅</span></th>
              <th class="sortable" data-sort-key="diffPct" data-sort-type="num" style="text-align:center" onclick="sortMSummary('diffPct','num')">Diff % <span class="sort-icon">⇅</span></th>
              <th class="sortable" data-sort-key="diffVal" data-sort-type="num" style="text-align:right" onclick="sortMSummary('diffVal','num')">Diff Amount <span class="sort-icon">⇅</span></th>
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

  <!-- ═══════════════════ CATEGORY SALES TAB ═══════════════════ -->
  <div class="tab-content" id="tab-category">

    <!-- Controls -->
    <div class="controls">
      <div class="ctrl-group">
        <span class="ctrl-label"><i class="fa fa-calendar"></i> Month</span>
        <select id="cMonthFilter" onchange="applyCategoryFilters()">
          <option value="ALL">All Months</option>
        </select>
      </div>
      <div class="divider"></div>
      <div class="ctrl-group">
        <span class="ctrl-label"><i class="fa fa-tags"></i> Category</span>
        <select id="cCategoryFilter" onchange="applyCategoryFilters()">
          <option value="ALL">All Categories</option>
        </select>
      </div>
      <div class="divider"></div>
      <div class="ctrl-group">
        <span class="ctrl-label"><i class="fa fa-layer-group"></i> Area</span>
        <select id="cAreaFilter" onchange="onCAreaChange()">
          <option value="ALL">All Areas</option>
        </select>
      </div>
      <div class="divider"></div>
      <div class="ctrl-group">
        <span class="ctrl-label"><i class="fa fa-store"></i> Store</span>
        <select id="cStoreFilter" onchange="applyCategoryFilters()">
          <option value="ALL">All Stores</option>
        </select>
      </div>
      <div class="records-count" id="cRecordsCount">—</div>
    </div>

    <!-- KPI Cards -->
    <div class="kpi-grid">
      <div class="kpi k-sales">
        <div class="kpi-label"><div class="kpi-icon"><i class="fa fa-peso-sign"></i></div>Total Sales</div>
        <div class="kpi-value gradient-text" id="c-kpi-sales">—</div>
        <div class="kpi-sub">Filtered period</div>
      </div>
      <div class="kpi k-ly">
        <div class="kpi-label"><div class="kpi-icon"><i class="fa fa-clock-rotate-left"></i></div>Sales LY</div>
        <div class="kpi-value gradient-text" id="c-kpi-ly">—</div>
        <div class="kpi-sub">Same period last year</div>
      </div>
      <div class="kpi k-pct">
        <div class="kpi-label"><div class="kpi-icon" id="c-kpi-pct-icon"><i class="fa fa-percent"></i></div>Diff %</div>
        <div class="kpi-value" id="c-kpi-pct">—</div>
        <div class="kpi-sub" id="c-kpi-pct-sub">vs last year</div>
      </div>
      <div class="kpi k-diff">
        <div class="kpi-label"><div class="kpi-icon" id="c-kpi-diff-icon"><i class="fa fa-arrow-trend-up"></i></div>Diff Amount</div>
        <div class="kpi-value" id="c-kpi-diff">—</div>
        <div class="kpi-sub" id="c-kpi-diff-sub">variance</div>
      </div>
      <div class="kpi k-stores">
        <div class="kpi-label"><div class="kpi-icon"><i class="fa fa-list"></i></div>Categories</div>
        <div class="kpi-value gradient-text" id="c-kpi-cats">—</div>
        <div class="kpi-sub" id="c-kpi-cats-sub" style="display:flex;gap:10px;flex-wrap:wrap;align-items:center;margin-top:6px">
          <span style="display:inline-flex;align-items:center;gap:4px"><i class="fa fa-arrow-up" style="color:#34d399;font-size:10px"></i> <span id="c-kpi-up" style="color:#34d399;font-weight:600;font-family:'JetBrains Mono',monospace">0</span> growth</span>
          <span style="color:var(--text-dim)">·</span>
          <span style="display:inline-flex;align-items:center;gap:4px"><i class="fa fa-arrow-down" style="color:#fb7185;font-size:10px"></i> <span id="c-kpi-down" style="color:#fb7185;font-weight:600;font-family:'JetBrains Mono',monospace">0</span> decline</span>
        </div>
      </div>
    </div>

    <!-- Charts Row 1 -->
    <div class="charts-grid">
      <div class="chart-card">
        <div class="chart-title"><i class="fa fa-chart-bar"></i> Diff % by Category — vs Last Year</div>
        <div style="position:relative;height:320px">
          <canvas id="cCategoryChart"></canvas>
        </div>
      </div>
      <div class="chart-card">
        <div class="chart-title"><i class="fa fa-chart-pie"></i> Category SOB %</div>
        <div style="position:relative;height:340px">
          <canvas id="cCategoryPie"></canvas>
        </div>
      </div>
    </div>

    <!-- Top & Bottom Movers Chart (full width) -->
    <div class="chart-card" style="margin-top:0">
      <div class="chart-title">
        <i class="fa fa-ranking-star"></i> Top &amp; Bottom Sub-Departments by Growth
        <span style="margin-left:auto;font-size:10.5px;color:var(--text-3);font-weight:500;text-transform:uppercase;letter-spacing:0.08em">Diff Amount</span>
      </div>
      <div style="position:relative;height:420px">
        <canvas id="cMoversChart"></canvas>
      </div>
    </div>

    <!-- Category Summary Table -->
    <div class="table-card">
      <div class="table-header">
        <div class="table-title"><i class="fa fa-table-list"></i> Category Summary</div>
        <div class="table-date" id="cTableInfo">—</div>
      </div>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th class="sortable" data-sort-key="category" data-sort-type="string" onclick="sortCSummary('category','string')">Category <span class="sort-icon">⇅</span></th>
              <th class="sortable" data-sort-key="subDepCount" data-sort-type="num" style="text-align:center" onclick="sortCSummary('subDepCount','num')">Sub-Depts <span class="sort-icon">⇅</span></th>
              <th class="sortable" data-sort-key="sales" data-sort-type="num" style="text-align:right" onclick="sortCSummary('sales','num')">Sales <span class="sort-icon">⇅</span></th>
              <th class="sortable" data-sort-key="salesLY" data-sort-type="num" style="text-align:right" onclick="sortCSummary('salesLY','num')">Sales LY <span class="sort-icon">⇅</span></th>
              <th class="sortable" data-sort-key="diffPct" data-sort-type="num" style="text-align:center" onclick="sortCSummary('diffPct','num')">Diff % <span class="sort-icon">⇅</span></th>
              <th class="sortable" data-sort-key="diffVal" data-sort-type="num" style="text-align:right" onclick="sortCSummary('diffVal','num')">Diff Amount <span class="sort-icon">⇅</span></th>
              <th class="sortable" data-sort-key="shareCur" data-sort-type="num" style="text-align:center" onclick="sortCSummary('shareCur','num')">Share % <span class="sort-icon">⇅</span></th>
            </tr>
          </thead>
          <tbody id="cSummaryBody">
            <tr><td colspan="7" class="empty-cell"><div class="empty-icon"><i class="fa fa-spinner fa-spin"></i></div><p>Loading...</p></td></tr>
          </tbody>
        </table>
      </div>
    </div>

    <!-- Sub-Department Detail Table -->
    <div class="table-card" style="margin-top:24px">
      <div class="table-header" style="gap:14px">
        <div class="table-title"><i class="fa fa-list-ul"></i> Sub-Department Detail</div>
        <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin-left:auto">
          <div class="ctrl-group">
            <span class="ctrl-label" style="font-size:10px"><i class="fa fa-filter"></i> Variance</span>
            <div class="sign-toggle" id="cSignToggle">
              <button class="active all" data-sign="ALL" onclick="setCSignFilter('ALL')">All</button>
              <button data-sign="POS" onclick="setCSignFilter('POS')"><i class="fa fa-arrow-up"></i> Positive</button>
              <button data-sign="NEG" onclick="setCSignFilter('NEG')"><i class="fa fa-arrow-down"></i> Negative</button>
            </div>
          </div>
          <button class="export-btn" onclick="exportCategoryToExcel()">
            <i class="fa fa-file-excel"></i> Export Excel
          </button>
        </div>
      </div>
      <div class="table-wrap" style="max-height:600px">
        <table>
          <thead>
            <tr>
              <th class="sortable" data-sort-key="category" data-sort-type="string" onclick="sortCDetail('category','string')">Category <span class="sort-icon">⇅</span></th>
              <th class="sortable" data-sort-key="sdepCode" data-sort-type="string" onclick="sortCDetail('sdepCode','string')">SDep Code <span class="sort-icon">⇅</span></th>
              <th class="sortable" data-sort-key="subDepName" data-sort-type="string" onclick="sortCDetail('subDepName','string')">Sub-Department <span class="sort-icon">⇅</span></th>
              <th class="sortable" data-sort-key="sales" data-sort-type="num" style="text-align:right" onclick="sortCDetail('sales','num')">Sales <span class="sort-icon">⇅</span></th>
              <th class="sortable" data-sort-key="salesLY" data-sort-type="num" style="text-align:right" onclick="sortCDetail('salesLY','num')">Sales LY <span class="sort-icon">⇅</span></th>
              <th class="sortable" data-sort-key="diffPct" data-sort-type="num" style="text-align:center" onclick="sortCDetail('diffPct','num')">Diff % <span class="sort-icon">⇅</span></th>
              <th class="sortable" data-sort-key="diffVal" data-sort-type="num" style="text-align:right" onclick="sortCDetail('diffVal','num')">Diff Amount <span class="sort-icon">⇅</span></th>
            </tr>
          </thead>
          <tbody id="cDetailBody">
            <tr><td colspan="7" class="empty-cell"><div class="empty-icon"><i class="fa fa-spinner fa-spin"></i></div><p>Loading...</p></td></tr>
          </tbody>
        </table>
      </div>
    </div>

    <!-- Area & Store Sales Breakdown -->
    <div class="table-card" style="margin-top:24px">
      <div class="table-header" style="gap:14px;flex-wrap:wrap">
        <div class="table-title"><i class="fa fa-layer-group"></i> Area &amp; Store Breakdown</div>
        <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin-left:auto">
          <div class="ctrl-group">
            <span class="ctrl-label" style="font-size:10px"><i class="fa fa-tags"></i> Category</span>
            <select id="cBreakdownCategory" onchange="applyBreakdownFilters()">
              <option value="ALL">All Categories</option>
            </select>
          </div>
          <div class="ctrl-group">
            <span class="ctrl-label" style="font-size:10px"><i class="fa fa-list"></i> Sub-Dept</span>
            <select id="cBreakdownSubDep" onchange="applyBreakdownFilters()">
              <option value="ALL">All Sub-Depts</option>
            </select>
          </div>
        </div>
      </div>
    </div>

    <!-- Sales by Area Table -->
    <div class="table-card" style="margin-top:14px">
      <div class="table-header">
        <div class="table-title"><i class="fa fa-map-location-dot"></i> Sales by Area</div>
        <div class="table-date" id="cAreaTableInfo">—</div>
      </div>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th class="sortable" data-sort-key="area" data-sort-type="string" onclick="sortCArea('area','string')">Area <span class="sort-icon">⇅</span></th>
              <th class="sortable" data-sort-key="sales" data-sort-type="num" style="text-align:right" onclick="sortCArea('sales','num')">Sales <span class="sort-icon">⇅</span></th>
              <th class="sortable" data-sort-key="salesLY" data-sort-type="num" style="text-align:right" onclick="sortCArea('salesLY','num')">Sales YA <span class="sort-icon">⇅</span></th>
              <th class="sortable" data-sort-key="diffPct" data-sort-type="num" style="text-align:center" onclick="sortCArea('diffPct','num')">Diff % <span class="sort-icon">⇅</span></th>
              <th class="sortable" data-sort-key="diffVal" data-sort-type="num" style="text-align:right" onclick="sortCArea('diffVal','num')">Diff Amount <span class="sort-icon">⇅</span></th>
            </tr>
          </thead>
          <tbody id="cAreaTableBody">
            <tr><td colspan="5" class="empty-cell"><div class="empty-icon"><i class="fa fa-spinner fa-spin"></i></div><p>Loading...</p></td></tr>
          </tbody>
        </table>
      </div>
    </div>

    <!-- Sales per Store Table -->
    <div class="table-card" style="margin-top:14px">
      <div class="table-header">
        <div class="table-title"><i class="fa fa-store"></i> Sales per Store</div>
        <div class="table-date" id="cStoreTableInfo">—</div>
      </div>
      <div class="table-wrap" style="max-height:600px">
        <table>
          <thead>
            <tr>
              <th class="sortable" data-sort-key="storeCode" data-sort-type="string" onclick="sortCStore('storeCode','string')">Store ID <span class="sort-icon">⇅</span></th>
              <th class="sortable" data-sort-key="storeName" data-sort-type="string" onclick="sortCStore('storeName','string')">Store Name <span class="sort-icon">⇅</span></th>
              <th class="sortable" data-sort-key="area" data-sort-type="string" onclick="sortCStore('area','string')">Area <span class="sort-icon">⇅</span></th>
              <th class="sortable" data-sort-key="sales" data-sort-type="num" style="text-align:right" onclick="sortCStore('sales','num')">Sales <span class="sort-icon">⇅</span></th>
              <th class="sortable" data-sort-key="salesLY" data-sort-type="num" style="text-align:right" onclick="sortCStore('salesLY','num')">Sales YA <span class="sort-icon">⇅</span></th>
              <th class="sortable" data-sort-key="diffPct" data-sort-type="num" style="text-align:center" onclick="sortCStore('diffPct','num')">Diff % <span class="sort-icon">⇅</span></th>
              <th class="sortable" data-sort-key="diffVal" data-sort-type="num" style="text-align:right" onclick="sortCStore('diffVal','num')">Diff Amount <span class="sort-icon">⇅</span></th>
            </tr>
          </thead>
          <tbody id="cStoreTableBody">
            <tr><td colspan="7" class="empty-cell"><div class="empty-icon"><i class="fa fa-spinner fa-spin"></i></div><p>Loading...</p></td></tr>
          </tbody>
        </table>
      </div>
    </div>

  </div><!-- /tab-category -->

</main>

<script>
// ─── Mobile detection & Chart.js global tuning ──────────────────────────────
const IS_MOBILE = window.matchMedia('(max-width: 768px)').matches;
if (window.Chart) {
  Chart.defaults.animation = IS_MOBILE ? false : { duration: 600 };
  Chart.defaults.font.family = "'Inter', sans-serif";
  if (IS_MOBILE) {
    Chart.defaults.elements.point.radius = 0;
    Chart.defaults.datasets.bar.maxBarThickness = 22;
  }
  // Register datalabels but make it opt-in per chart
  if (window.ChartDataLabels) {
    Chart.register(ChartDataLabels);
    Chart.defaults.plugins.datalabels = { display: false }; // off by default
  }
}

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
let avgStoreChartInst = null;
let avgAreaChartInst = null;

// Category tab state
let cCategoryChartInst = null;
let cCategoryPieInst = null;
let cMoversChartInst = null;
let cAreaChartInst = null;
let cSignFilter = 'ALL';
let cSummaryRowsCache = [];
let cDetailRowsCache = [];
let cAreaRowsCache = [];
let cStoreRowsCache = [];
let cSummarySort = { key: 'sales', dir: 'desc', type: 'num' };
let cDetailSort  = { key: 'sales', dir: 'desc', type: 'num' };
let cAreaSort    = { key: 'sales', dir: 'desc', type: 'num' };
let cStoreSort   = { key: 'sales', dir: 'desc', type: 'num' };
let cFilters = { months: [], categories: [], areas: [], stores: [] };
let cBreakdownData = []; // raw filtered data for breakdown filtering
let cSubDepsForBreakdown = []; // list of sub-depts grouped by category

// Daily sort state
let dailyRowsCache = [];
let dailySort = { key: 'sales', dir: 'desc', type: 'num' };

// Monthly tab state
let mBarChartInst = null;
let mLineChartInst = null;
let mSignFilter = 'ALL';
let mDetailRowsCache = [];
let mSummaryRowsCache = [];
let mAreaRowsCache = [];
let mSummarySort = { key: 'sales', dir: 'desc', type: 'num' };
let mAreaSort = { key: 'sales', dir: 'desc', type: 'num' };
let mMonthFilters = { areas: [], stores: [], months: [] };

// ─── Tab switching ───────────────────────────────────────────────────────
function switchTab(tab) {
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.toggle('active', b.dataset.tab === tab));
  document.querySelectorAll('.tab-content').forEach(c => c.classList.toggle('active', c.id === \`tab-\${tab}\`));
  if (tab === 'monthly' && !mMonthFilters.months.length) {
    initMonthlyTab();
  }
  if (tab === 'category' && !cFilters.categories.length) {
    initCategoryTab();
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
    dailyRowsCache = rows;
    renderTable(sortRows(rows, dailySort));
    updateSortHeaders('mainTable', dailySort);
    renderCharts(rows);
    renderMissing(json.missing || [], !!date);
    renderAverages(); // independent of date filter
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
      <td><div class="just-full">\${just}</div></td>
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
  const top = rows.slice(0, IS_MOBILE ? 8 : 14);
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
    mSel.innerHTML = '<option value="ALL">All Months</option>';
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
  if (month !== 'ALL') params.set('month', month);
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

    const monthLabel = month === 'ALL' ? 'All Months' : ((mMonthFilters.months.find(m => m.value === month) || {}).label || month);
    document.getElementById('mTableDate').textContent = monthLabel;

    renderMKPIs(summary);
    mSummaryRowsCache = summary;
    renderMSummary(sortRows(summary, mSummarySort));
    updateSortHeaders('mSummaryBody', mSummarySort);
    renderMDetail(detail);
    renderMCharts(summary, trend);

    // Aggregate by area from summary
    const areaMap = {};
    summary.forEach(r => {
      if (!r.area) return;
      if (!areaMap[r.area]) areaMap[r.area] = { area: r.area, sales: 0, salesLY: 0 };
      areaMap[r.area].sales   += r.sales;
      areaMap[r.area].salesLY += r.salesLY;
    });
    const areaRows = Object.values(areaMap).map(a => {
      const diffVal = a.sales - a.salesLY;
      const diffPct = a.salesLY !== 0 ? (diffVal / a.salesLY) * 100 : 0;
      return { ...a, diffVal, diffPct: parseFloat(diffPct.toFixed(2)) };
    });
    mAreaRowsCache = areaRows;
    renderMAreaTable(sortRows(areaRows, mAreaSort));
    updateSortHeaders('mAreaTableBody', mAreaSort);
    document.getElementById('mAreaTableInfo').textContent = monthLabel;
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

  // Growth vs decline counts
  const growth  = rows.filter(r => (r.sales - r.salesLY) > 0).length;
  const decline = rows.filter(r => (r.sales - r.salesLY) < 0).length;
  document.getElementById('m-kpi-up').textContent = growth;
  document.getElementById('m-kpi-down').textContent = decline;
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
  // Bar chart removed - replaced with Sales per Area table (rendered separately)

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

// ─── Sorting helpers ────────────────────────────────────────────────────────
function sortRows(rows, sortState) {
  if (!sortState || !sortState.key) return rows;
  const { key, dir, type } = sortState;
  const mult = dir === 'asc' ? 1 : -1;
  return [...rows].sort((a, b) => {
    let av = a[key], bv = b[key];
    if (type === 'num') {
      av = parseFloat(av) || 0;
      bv = parseFloat(bv) || 0;
      return (av - bv) * mult;
    }
    av = (av || '').toString().toLowerCase();
    bv = (bv || '').toString().toLowerCase();
    if (av < bv) return -1 * mult;
    if (av > bv) return  1 * mult;
    return 0;
  });
}

function updateSortHeaders(tbodyId, sortState) {
  const tbody = document.getElementById(tbodyId);
  if (!tbody) return;
  const table = tbody.closest('table');
  if (!table) return;
  table.querySelectorAll('th.sortable').forEach(th => {
    th.classList.remove('sort-asc','sort-desc');
    const icon = th.querySelector('.sort-icon');
    if (th.dataset.sortKey === sortState.key) {
      th.classList.add(sortState.dir === 'asc' ? 'sort-asc' : 'sort-desc');
      if (icon) icon.textContent = sortState.dir === 'asc' ? '▲' : '▼';
    } else {
      if (icon) icon.textContent = '⇅';
    }
  });
}

function renderMAreaTable(rows) {
  const tbody = document.getElementById('mAreaTableBody');
  if (!rows.length) {
    tbody.innerHTML = \`<tr><td colspan="5" class="empty-cell"><div class="empty-icon"><i class="fa fa-magnifying-glass"></i></div><p>No area data</p></td></tr>\`;
    return;
  }
  const totSales = rows.reduce((s,r)=>s+r.sales,0);
  const totLY    = rows.reduce((s,r)=>s+r.salesLY,0);
  const totDiff  = totSales - totLY;
  const totPct   = totLY ? (totDiff/totLY*100) : 0;

  let html = rows.map(r => {
    const pctCls = r.diffPct > 0.05 ? 'up' : r.diffPct < -0.05 ? 'down' : 'flat';
    const color = AREA_COLORS[r.area] || DEFAULT_COLOR;
    const diffColor = r.diffVal >= 0 ? '#34d399' : '#fb7185';
    const arrow = r.diffPct > 0.05 ? '↑' : r.diffPct < -0.05 ? '↓' : '—';
    const pctStr = r.diffPct !== 0 ? \`\${arrow} \${Math.abs(r.diffPct).toFixed(2)}%\` : '—';
    return \`<tr>
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
    <td style="font-size:12px;letter-spacing:0.1em;text-transform:uppercase;color:var(--indigo2)">TOTAL · \${rows.length} AREAS</td>
    <td style="text-align:right"><span class="num">\${fmtFull(totSales)}</span></td>
    <td style="text-align:right"><span class="num" style="color:var(--text-3)">\${fmtFull(totLY)}</span></td>
    <td style="text-align:center"><span class="pill \${totPctCls}">\${totArrow} \${Math.abs(totPct).toFixed(2)}%</span></td>
    <td style="text-align:right"><span class="num" style="color:\${totDiff>=0?'#34d399':'#fb7185'}">\${totDiff>=0?'+':''}\${fmtFull(totDiff)}</span></td>
  </tr>\`;
  tbody.innerHTML = html;
}

function sortMArea(key, type) {
  if (mAreaSort.key === key) mAreaSort.dir = mAreaSort.dir === 'asc' ? 'desc' : 'asc';
  else mAreaSort = { key, dir: 'desc', type };
  mAreaSort.type = type;
  renderMAreaTable(sortRows(mAreaRowsCache, mAreaSort));
  updateSortHeaders('mAreaTableBody', mAreaSort);
}

function sortDaily(key, type) {
  if (dailySort.key === key) {
    dailySort.dir = dailySort.dir === 'asc' ? 'desc' : 'asc';
  } else {
    dailySort = { key, dir: 'desc', type };
  }
  dailySort.type = type;
  renderTable(sortRows(dailyRowsCache, dailySort));
  updateSortHeaders('tableBody', dailySort);
}

function sortMSummary(key, type) {
  if (mSummarySort.key === key) {
    mSummarySort.dir = mSummarySort.dir === 'asc' ? 'desc' : 'asc';
  } else {
    mSummarySort = { key, dir: 'desc', type };
  }
  mSummarySort.type = type;
  renderMSummary(sortRows(mSummaryRowsCache, mSummarySort));
  updateSortHeaders('mSummaryBody', mSummarySort);
}

// ─── Average Daily Sales charts ────────────────────────────────────────────
async function renderAverages() {
  // Respect Area and Store filters; ignore Date (averages span all data)
  const area  = document.getElementById('areaFilter').value;
  const store = document.getElementById('storeFilter').value;

  const params = new URLSearchParams();
  if (area  !== 'ALL') params.set('area', area);
  if (store !== 'ALL') params.set('store', store);

  try {
    const res  = await fetch('/api/averages?' + params);
    const json = await res.json();
    if (!json.success) throw new Error(json.error);

    drawAvgStoreChart(json.perStore || []);
    drawAvgAreaChart(json.perArea || []);
  } catch(e) {
    console.error('Averages fetch error:', e);
  }
}

function drawAvgStoreChart(perStore) {
  const canvas = document.getElementById('avgStoreChart');
  if (!canvas) return;
  if (avgStoreChartInst) avgStoreChartInst.destroy();

  // Show ALL stores
  const all = perStore;
  const labels = all.map(r => r.storeName);
  const data = all.map(r => r.avg);

  // Dynamic height: ~26px per bar, min 320
  const wrap = document.getElementById('avgStoreChartWrap');
  const dynHeight = Math.max(320, all.length * 26 + 40);
  if (wrap) wrap.style.height = dynHeight + 'px';

  const ctx = canvas.getContext('2d');
  const bgColors = all.map(r => {
    const [c1, c2] = AREA_GRADIENTS[r.area] || [DEFAULT_COLOR, DEFAULT_COLOR];
    const g = ctx.createLinearGradient(0, 0, canvas.width || 800, 0);
    g.addColorStop(0, c1); g.addColorStop(1, c2);
    return g;
  });

  avgStoreChartInst = new Chart(canvas, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: 'Avg Daily Sales',
        data,
        backgroundColor: bgColors,
        borderRadius: 5, borderSkipped: false, barPercentage: 0.78,
      }]
    },
    options: {
      indexAxis: 'y',
      responsive: true, maintainAspectRatio: false,
      layout: { padding: { right: 80 } },
      plugins: {
        legend: { display: false },
        datalabels: {
          display: true,
          anchor: 'end',
          align: 'end',
          color: '#e8ecf4',
          font: { family: "'JetBrains Mono'", size: 10, weight: '600' },
          formatter: v => v >= 1e6 ? '₱' + (v/1e6).toFixed(2) + 'M' : v >= 1e3 ? '₱' + (v/1e3).toFixed(0) + 'K' : '₱' + v.toFixed(0),
          padding: { left: 6 }
        },
        tooltip: {
          backgroundColor: 'rgba(15,20,35,0.95)', borderColor: 'rgba(99,102,241,0.4)', borderWidth: 1,
          padding: 12, cornerRadius: 10,
          titleFont: { family: "'Space Grotesk'", size: 13, weight: '600' },
          bodyFont:  { family: "'JetBrains Mono'", size: 11.5 },
          titleColor: '#f0f3fb', bodyColor: '#a8b3d1',
          callbacks: {
            label: (ctx) => {
              const r = all[ctx.dataIndex];
              return [
                ' Avg/Day: ₱' + ctx.raw.toLocaleString('en-PH', { minimumFractionDigits: 2 }),
                ' Total: ₱' + r.total.toLocaleString('en-PH', { minimumFractionDigits: 2 }),
                ' Days: ' + r.dayCount,
                ' Area: ' + r.area,
              ];
            }
          }
        }
      },
      scales: {
        x: {
          grid: { color: 'rgba(148,163,200,0.06)' },
          ticks: { color: '#a8b3d1', font: { size: 10, family: "'JetBrains Mono'", weight: '500' }, callback: v => v >= 1e6 ? '₱' + (v/1e6).toFixed(1)+'M' : v >= 1e3 ? '₱' + (v/1e3).toFixed(0)+'K' : v }
        },
        y: {
          grid: { color: 'rgba(148,163,200,0.04)', drawBorder: false },
          ticks: { color: '#a8b3d1', font: { size: 11, family: "'Inter'", weight: '500' }, autoSkip: false }
        }
      }
    }
  });
}

function drawAvgAreaChart(perArea) {
  const canvas = document.getElementById('avgAreaChart');
  if (!canvas) return;
  if (avgAreaChartInst) avgAreaChartInst.destroy();

  const labels = perArea.map(a => a.area);
  const data = perArea.map(a => a.avgPerDay);

  const ctx = canvas.getContext('2d');
  const gradients = perArea.map(a => {
    const [c1, c2] = AREA_GRADIENTS[a.area] || [DEFAULT_COLOR, DEFAULT_COLOR];
    const g = ctx.createLinearGradient(0, 0, 0, 320);
    g.addColorStop(0, c2); g.addColorStop(1, c1);
    return g;
  });

  avgAreaChartInst = new Chart(canvas, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: 'Avg Daily Sales',
        data,
        backgroundColor: gradients,
        borderRadius: 8, borderSkipped: false, barPercentage: 0.6,
      }]
    },
    options: {
      indexAxis: 'y',
      responsive: true, maintainAspectRatio: false,
      layout: { padding: { right: 90 } },
      plugins: {
        legend: { display: false },
        datalabels: {
          display: true,
          anchor: 'end',
          align: 'end',
          color: '#e8ecf4',
          font: { family: "'JetBrains Mono'", size: 12, weight: '700' },
          formatter: v => v >= 1e6 ? '₱' + (v/1e6).toFixed(2) + 'M' : v >= 1e3 ? '₱' + (v/1e3).toFixed(0) + 'K' : '₱' + v.toFixed(0),
          padding: { left: 8 }
        },
        tooltip: {
          backgroundColor: 'rgba(15,20,35,0.95)', borderColor: 'rgba(99,102,241,0.4)', borderWidth: 1,
          padding: 12, cornerRadius: 10,
          titleFont: { family: "'Space Grotesk'", size: 13, weight: '600' },
          bodyFont:  { family: "'JetBrains Mono'", size: 11.5 },
          titleColor: '#f0f3fb', bodyColor: '#a8b3d1',
          callbacks: {
            label: (ctx) => {
              const a = perArea[ctx.dataIndex];
              return [
                ' Avg/Day: ₱' + ctx.raw.toLocaleString('en-PH', { minimumFractionDigits: 2 }),
                ' Total: ₱' + a.total.toLocaleString('en-PH', { minimumFractionDigits: 2 }),
                ' Days: ' + a.dayCount,
                ' Records: ' + a.recordCount,
              ];
            }
          }
        }
      },
      scales: {
        x: {
          grid: { color: 'rgba(148,163,200,0.06)' },
          ticks: { color: '#a8b3d1', font: { size: 10.5, family: "'JetBrains Mono'", weight: '500' }, callback: v => v >= 1e6 ? '₱' + (v/1e6).toFixed(1)+'M' : v >= 1e3 ? '₱' + (v/1e3).toFixed(0)+'K' : v }
        },
        y: {
          grid: { color: 'rgba(148,163,200,0.04)', drawBorder: false },
          ticks: { color: '#a8b3d1', font: { size: 11, family: "'Inter'", weight: '600' } }
        }
      }
    }
  });
}

// ═══════════════════════ CATEGORY SALES TAB ═══════════════════════════════
const CATEGORY_PALETTE = ['#6366f1','#10b981','#f59e0b','#a855f7','#06b6d4','#ec4899','#f43f5e','#22d3ee','#fbbf24','#34d399','#818cf8','#c084fc','#fb7185','#84cc16','#0ea5e9','#d946ef'];

function colorForCategory(name, idx) {
  return CATEGORY_PALETTE[idx % CATEGORY_PALETTE.length];
}

async function initCategoryTab() {
  try {
    const res = await fetch('/api/category-filters');
    const json = await res.json();
    if (!json.success) throw new Error(json.error);

    cFilters = json;

    // Month dropdown
    const mSel = document.getElementById('cMonthFilter');
    mSel.innerHTML = '<option value="ALL">All Months</option>';
    json.months.forEach(m => {
      const o = document.createElement('option');
      o.value = m;
      o.textContent = m;
      mSel.appendChild(o);
    });

    // Category dropdown
    const catSel = document.getElementById('cCategoryFilter');
    catSel.innerHTML = '<option value="ALL">All Categories</option>';
    json.categories.forEach(c => {
      const o = document.createElement('option');
      o.value = o.textContent = c;
      catSel.appendChild(o);
    });

    // Breakdown category dropdown (same options)
    const bCatSel = document.getElementById('cBreakdownCategory');
    bCatSel.innerHTML = '<option value="ALL">All Categories</option>';
    json.categories.forEach(c => {
      const o = document.createElement('option');
      o.value = o.textContent = c;
      bCatSel.appendChild(o);
    });

    // Area dropdown
    const aSel = document.getElementById('cAreaFilter');
    aSel.innerHTML = '<option value="ALL">All Areas</option>';
    json.areas.forEach(a => {
      const o = document.createElement('option');
      o.value = o.textContent = a;
      aSel.appendChild(o);
    });

    populateCStores(json.stores);
    await applyCategoryFilters();
  } catch(e) {
    console.error('Category init error:', e);
    document.getElementById('cSummaryBody').innerHTML =
      \`<tr><td colspan="7" class="empty-cell"><div class="empty-icon" style="background:linear-gradient(135deg,rgba(244,63,94,0.1),rgba(251,113,133,0.1));color:#fb7185"><i class="fa fa-triangle-exclamation"></i></div><p>\${e.message}</p></td></tr>\`;
  }
}

function populateCStores(stores) {
  const sel = document.getElementById('cStoreFilter');
  sel.innerHTML = '<option value="ALL">All Stores</option>';
  stores.forEach(s => {
    const o = document.createElement('option');
    o.value = o.textContent = s;
    sel.appendChild(o);
  });
}

async function onCAreaChange() {
  const area = document.getElementById('cAreaFilter').value;
  if (area === 'ALL') {
    populateCStores(cFilters.stores);
  } else {
    try {
      const res = await fetch('/api/category-filters?area=' + encodeURIComponent(area));
      const json = await res.json();
      if (json.success) populateCStores(json.stores);
    } catch(e) {}
  }
  document.getElementById('cStoreFilter').value = 'ALL';
  await applyCategoryFilters();
}

function setCSignFilter(sign) {
  cSignFilter = sign;
  document.querySelectorAll('#cSignToggle button').forEach(b => {
    b.classList.toggle('active', b.dataset.sign === sign);
  });
  applyCategoryFilters();
}

async function applyCategoryFilters() {
  const month = document.getElementById('cMonthFilter').value;
  const category = document.getElementById('cCategoryFilter').value;
  const area = document.getElementById('cAreaFilter').value;
  const store = document.getElementById('cStoreFilter').value;

  const params = new URLSearchParams();
  if (month !== 'ALL') params.set('month', month);
  if (category !== 'ALL') params.set('category', category);
  if (area !== 'ALL') params.set('area', area);
  if (store !== 'ALL') params.set('store', store);
  if (cSignFilter !== 'ALL') params.set('sign', cSignFilter);

  try {
    const res = await fetch('/api/category?' + params);
    const json = await res.json();
    if (!json.success) throw new Error(json.error);

    const summary = json.summary || [];
    const detail = json.detail || [];
    const byArea = json.byArea || [];
    const byStore = json.byStore || [];
    const subDepsByCategory = json.subDepsByCategory || {};
    const movers = json.movers || { top: [], bottom: [] };

    cSummaryRowsCache = summary;
    cDetailRowsCache = detail;
    cAreaRowsCache = byArea;
    cStoreRowsCache = byStore;
    cSubDepsForBreakdown = subDepsByCategory;

    // Reset breakdown filters to ALL on main filter change
    if (document.getElementById('cBreakdownCategory').value !== 'ALL') {
      document.getElementById('cBreakdownCategory').value = 'ALL';
    }
    updateBreakdownSubDepDropdown();

    document.getElementById('cRecordsCount').innerHTML =
      \`<span>\${summary.length}</span> categor\${summary.length!==1?'ies':'y'} · <span>\${detail.length}</span> sub-dept\${detail.length!==1?'s':''}\`;

    const parts = [];
    if (month !== 'ALL') parts.push(month);
    if (category !== 'ALL') parts.push(category);
    if (area !== 'ALL') parts.push(area);
    if (store !== 'ALL') parts.push(store);
    document.getElementById('cTableInfo').textContent = parts.length ? parts.join(' · ') : 'All data';

    renderCKPIs(summary);
    renderCSummary(sortRows(summary, cSummarySort));
    updateSortHeaders('cSummaryBody', cSummarySort);
    renderCDetail(sortRows(detail, cDetailSort));
    updateSortHeaders('cDetailBody', cDetailSort);
    renderCCharts(summary, byArea, movers);
    renderCAreaTable(sortRows(byArea, cAreaSort));
    renderCStoreTable(sortRows(byStore, cStoreSort));
    updateSortHeaders('cAreaTableBody', cAreaSort);
    updateSortHeaders('cStoreTableBody', cStoreSort);
    document.getElementById('cAreaTableInfo').textContent = parts.length ? parts.join(' · ') : 'All data';
    document.getElementById('cStoreTableInfo').textContent = parts.length ? parts.join(' · ') : 'All data';
  } catch(e) {
    console.error(e);
    document.getElementById('cSummaryBody').innerHTML =
      \`<tr><td colspan="7" class="empty-cell"><div class="empty-icon" style="background:linear-gradient(135deg,rgba(244,63,94,0.1),rgba(251,113,133,0.1));color:#fb7185"><i class="fa fa-triangle-exclamation"></i></div><p>\${e.message}</p></td></tr>\`;
  }
}

function renderCKPIs(summary) {
  const totSales = summary.reduce((s,r)=>s+(r.sales||0),0);
  const totLY    = summary.reduce((s,r)=>s+(r.salesLY||0),0);
  const totDiff  = totSales - totLY;
  const totPct   = totLY !== 0 ? (totDiff/totLY*100) : 0;

  document.getElementById('c-kpi-sales').textContent = fmt(totSales);
  document.getElementById('c-kpi-ly').textContent = fmt(totLY);

  const pctEl = document.getElementById('c-kpi-pct');
  pctEl.textContent = (totPct>=0?'+':'') + totPct.toFixed(2) + '%';
  pctEl.style.background = \`linear-gradient(135deg, \${totPct>=0?'#10b981':'#f43f5e'} 0%, \${totPct>=0?'#34d399':'#fb7185'} 100%)\`;
  pctEl.style.webkitBackgroundClip = 'text';
  pctEl.style.backgroundClip = 'text';
  pctEl.style.webkitTextFillColor = 'transparent';
  document.getElementById('c-kpi-pct-sub').innerHTML = \`<i class="fa fa-arrow-\${totPct>=0?'up':'down'}" style="color:\${totPct>=0?'#34d399':'#fb7185'}"></i> vs last year\`;
  document.getElementById('c-kpi-pct-icon').style.background = \`linear-gradient(135deg, \${totPct>=0?'#10b981':'#f43f5e'}, \${totPct>=0?'#34d399':'#fb7185'})\`;

  const diffEl = document.getElementById('c-kpi-diff');
  diffEl.textContent = (totDiff>=0?'+':'') + fmt(totDiff);
  diffEl.style.background = \`linear-gradient(135deg, \${totDiff>=0?'#06b6d4':'#f43f5e'} 0%, \${totDiff>=0?'#22d3ee':'#fb7185'} 100%)\`;
  diffEl.style.webkitBackgroundClip = 'text';
  diffEl.style.backgroundClip = 'text';
  diffEl.style.webkitTextFillColor = 'transparent';
  document.getElementById('c-kpi-diff-sub').innerHTML = \`<i class="fa fa-\${totDiff>=0?'circle-check':'circle-xmark'}" style="color:\${totDiff>=0?'#22d3ee':'#fb7185'}"></i> variance\`;
  document.getElementById('c-kpi-diff-icon').style.background = \`linear-gradient(135deg, \${totDiff>=0?'#06b6d4':'#f43f5e'}, \${totDiff>=0?'#22d3ee':'#fb7185'})\`;

  document.getElementById('c-kpi-cats').textContent = summary.length;

  const up = summary.filter(r => (r.sales - r.salesLY) > 0).length;
  const dn = summary.filter(r => (r.sales - r.salesLY) < 0).length;
  document.getElementById('c-kpi-up').textContent = up;
  document.getElementById('c-kpi-down').textContent = dn;
}

function renderCSummary(rows) {
  const tbody = document.getElementById('cSummaryBody');
  if (!rows.length) {
    tbody.innerHTML = \`<tr><td colspan="7" class="empty-cell"><div class="empty-icon"><i class="fa fa-magnifying-glass"></i></div><p>No category data for selected filters</p></td></tr>\`;
    return;
  }

  const totSales = rows.reduce((s,r)=>s+r.sales,0);
  const totLY    = rows.reduce((s,r)=>s+r.salesLY,0);
  const totDiff  = totSales - totLY;
  const totPct   = totLY ? (totDiff/totLY*100) : 0;
  const totSubDeps = rows.reduce((s,r)=>s+r.subDepCount,0);

  let html = rows.map((r,i) => {
    const pctCls = r.diffPct > 0.05 ? 'up' : r.diffPct < -0.05 ? 'down' : 'flat';
    const color = colorForCategory(r.category, i);
    const diffColor = r.diffVal >= 0 ? '#34d399' : '#fb7185';
    const arrow = r.diffPct > 0.05 ? '↑' : r.diffPct < -0.05 ? '↓' : '—';
    const pctStr = r.diffPct !== 0 ? \`\${arrow} \${Math.abs(r.diffPct).toFixed(2)}%\` : '—';
    return \`<tr>
      <td>
        <div class="store-cell">
          <div class="store-avatar" style="background:linear-gradient(135deg, \${color}, \${color}cc)">\${(r.category || '?').substring(0,2).toUpperCase()}</div>
          <div class="store-info">
            <div class="store-name">\${r.category || '—'}</div>
          </div>
        </div>
      </td>
      <td style="text-align:center"><span class="num" style="color:var(--text-2)">\${r.subDepCount}</span></td>
      <td style="text-align:right"><span class="num num-bold" style="color:var(--text-1)">\${fmtFull(r.sales)}</span></td>
      <td style="text-align:right"><span class="num" style="color:var(--text-3)">\${fmtFull(r.salesLY)}</span></td>
      <td style="text-align:center"><span class="pill \${pctCls}">\${pctStr}</span></td>
      <td style="text-align:right"><span class="num num-bold" style="color:\${diffColor}">\${r.diffVal>=0?'+':''}\${fmtFull(r.diffVal)}</span></td>
      <td style="text-align:center"><span class="num" style="color:var(--indigo2);font-weight:600">\${(r.shareCur||0).toFixed(2)}%</span></td>
    </tr>\`;
  }).join('');

  const totPctCls = totPct > 0.05 ? 'up' : totPct < -0.05 ? 'down' : 'flat';
  const totArrow = totPct > 0.05 ? '↑' : totPct < -0.05 ? '↓' : '—';
  html += \`<tr class="summary-row">
    <td style="font-size:12px;letter-spacing:0.1em;text-transform:uppercase;color:var(--indigo2)">TOTAL · \${rows.length} CATEGORIES</td>
    <td style="text-align:center"><span class="num">\${totSubDeps}</span></td>
    <td style="text-align:right"><span class="num">\${fmtFull(totSales)}</span></td>
    <td style="text-align:right"><span class="num" style="color:var(--text-3)">\${fmtFull(totLY)}</span></td>
    <td style="text-align:center"><span class="pill \${totPctCls}">\${totArrow} \${Math.abs(totPct).toFixed(2)}%</span></td>
    <td style="text-align:right"><span class="num" style="color:\${totDiff>=0?'#34d399':'#fb7185'}">\${totDiff>=0?'+':''}\${fmtFull(totDiff)}</span></td>
    <td style="text-align:center"><span class="num" style="color:var(--text-3)">100.00%</span></td>
  </tr>\`;

  tbody.innerHTML = html;
}

function renderCDetail(rows) {
  const tbody = document.getElementById('cDetailBody');
  if (!rows.length) {
    tbody.innerHTML = \`<tr><td colspan="7" class="empty-cell"><div class="empty-icon"><i class="fa fa-magnifying-glass"></i></div><p>No sub-departments found</p><small>Try adjusting your filters</small></td></tr>\`;
    return;
  }

  // Build category->color mapping based on summary rows for consistent coloring
  const catColors = {};
  cSummaryRowsCache.forEach((r,i) => { catColors[r.category] = colorForCategory(r.category, i); });

  const html = rows.map(r => {
    const pctCls = r.diffPct > 0.05 ? 'up' : r.diffPct < -0.05 ? 'down' : 'flat';
    const color = catColors[r.category] || DEFAULT_COLOR;
    const diffColor = r.diffVal >= 0 ? '#34d399' : '#fb7185';
    const arrow = r.diffPct > 0.05 ? '↑' : r.diffPct < -0.05 ? '↓' : '—';
    const pctStr = r.diffPct !== 0 ? \`\${arrow} \${Math.abs(r.diffPct).toFixed(2)}%\` : '—';

    return \`<tr>
      <td><span class="area-tag"><span class="area-dot" style="background:\${color};color:\${color}"></span>\${r.category || '—'}</span></td>
      <td><span class="num" style="color:var(--text-2);font-weight:600">\${r.sdepCode || '—'}</span></td>
      <td><span style="font-size:12.5px;color:var(--text-1);font-weight:500">\${r.subDepName || '—'}</span></td>
      <td style="text-align:right"><span class="num num-bold" style="color:var(--text-1)">\${fmtFull(r.sales)}</span></td>
      <td style="text-align:right"><span class="num" style="color:var(--text-3)">\${fmtFull(r.salesLY)}</span></td>
      <td style="text-align:center"><span class="pill \${pctCls}">\${pctStr}</span></td>
      <td style="text-align:right"><span class="num num-bold" style="color:\${diffColor}">\${r.diffVal>=0?'+':''}\${fmtFull(r.diffVal)}</span></td>
    </tr>\`;
  }).join('');

  tbody.innerHTML = html;
}

function renderCCharts(summary, byArea, movers) {
  // ─── Diff % by Category (bar) ──────────────────────────────────────────
  if (cCategoryChartInst) cCategoryChartInst.destroy();
  const cat = document.getElementById('cCategoryChart');
  const top = summary.slice(0, IS_MOBILE ? 8 : 14);
  const diffPcts = top.map(r => r.diffPct);
  const barColors = diffPcts.map(p => p >= 0 ? 'rgba(16,185,129,0.85)' : 'rgba(244,63,94,0.85)');
  const borderColors = diffPcts.map(p => p >= 0 ? '#10b981' : '#f43f5e');

  cCategoryChartInst = new Chart(cat, {
    type: 'bar',
    data: {
      labels: top.map(r => r.category),
      datasets: [{
        label: 'Diff %',
        data: diffPcts,
        backgroundColor: barColors,
        borderColor: borderColors,
        borderWidth: 1,
        borderRadius: 6, borderSkipped: false, barPercentage: 0.72,
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      layout: { padding: { top: 24, bottom: 10 } },
      plugins: {
        legend: { display: false },
        datalabels: {
          display: true,
          anchor: ctx => ctx.dataset.data[ctx.dataIndex] >= 0 ? 'end' : 'start',
          align:  ctx => ctx.dataset.data[ctx.dataIndex] >= 0 ? 'top' : 'bottom',
          color: ctx => ctx.dataset.data[ctx.dataIndex] >= 0 ? '#34d399' : '#fb7185',
          font: { family: "'JetBrains Mono'", size: 10.5, weight: '700' },
          formatter: v => (v >= 0 ? '+' : '') + v.toFixed(2) + '%',
          padding: { top: 4, bottom: 4 }
        },
        tooltip: {
          backgroundColor:'rgba(15,20,35,0.95)', borderColor:'rgba(99,102,241,0.4)', borderWidth:1,
          padding:12, cornerRadius:10,
          titleFont:{family:"'Space Grotesk'",size:13,weight:'600'},
          bodyFont:{family:"'JetBrains Mono'",size:11.5},
          titleColor:'#f0f3fb', bodyColor:'#a8b3d1',
          callbacks: {
            label: ctx => {
              const r = top[ctx.dataIndex];
              return [
                ' Diff %: ' + (r.diffPct>=0?'+':'') + r.diffPct.toFixed(2) + '%',
                ' Sales: ₱' + r.sales.toLocaleString('en-PH',{minimumFractionDigits:2}),
                ' Sales LY: ₱' + r.salesLY.toLocaleString('en-PH',{minimumFractionDigits:2}),
                ' Diff Amount: ' + (r.diffVal>=0?'+':'') + '₱' + Math.abs(r.diffVal).toLocaleString('en-PH',{minimumFractionDigits:2}),
              ];
            }
          }
        }
      },
      scales: {
        x: { grid:{color:'rgba(148,163,200,0.04)',drawBorder:false}, ticks:{color:'#a8b3d1',font:{size:10,family:"'Inter'",weight:'500'},maxRotation:40} },
        y: { grid:{color:'rgba(148,163,200,0.06)'}, ticks:{color:'#a8b3d1',font:{size:10.5,family:"'JetBrains Mono'",weight:'500'}, callback: v => (v>=0?'+':'') + v.toFixed(0) + '%' } }
      }
    }
  });

  // ─── Category SOB % (doughnut) ─────────────────────────────────────────
  if (cCategoryPieInst) cCategoryPieInst.destroy();
  const pie = document.getElementById('cCategoryPie');
  const pieColors = summary.map((r,i) => colorForCategory(r.category, i));
  const totalSales = summary.reduce((s,r)=>s+r.sales,0);

  cCategoryPieInst = new Chart(pie, {
    type: 'doughnut',
    data: { labels: summary.map(r=>r.category), datasets: [{ data: summary.map(r=>r.sales), backgroundColor: pieColors, borderColor: 'rgba(10,14,26,0.8)', borderWidth: 3, hoverOffset: 10 }] },
    options: {
      responsive: true, maintainAspectRatio: false, cutout: '55%',
      layout: { padding: 20 },
      plugins: {
        legend: { display: false },
        datalabels: {
          display: ctx => {
            const total = ctx.dataset.data.reduce((a,b)=>a+b,0);
            const val = ctx.dataset.data[ctx.dataIndex];
            const pct = total ? (val/total*100) : 0;
            return pct >= 2; // hide labels for tiny slices to avoid clutter
          },
          color: '#fff',
          font: { family: "'Inter'", size: 11, weight: '700' },
          textAlign: 'center',
          textStrokeColor: 'rgba(0,0,0,0.55)',
          textStrokeWidth: 3,
          formatter: (val, ctx) => {
            const total = ctx.dataset.data.reduce((a,b)=>a+b,0);
            const pct = total ? (val/total*100) : 0;
            const label = ctx.chart.data.labels[ctx.dataIndex];
            return label + '\\n' + pct.toFixed(1) + '%';
          }
        },
        tooltip: {
          backgroundColor:'rgba(15,20,35,0.95)', borderColor:'rgba(99,102,241,0.4)', borderWidth:1, padding:12, cornerRadius:10,
          titleFont:{family:"'Space Grotesk'",weight:'600',size:13}, bodyFont:{family:"'JetBrains Mono'",size:11.5},
          titleColor:'#f0f3fb', bodyColor:'#a8b3d1',
          callbacks: {
            label: ctx => {
              const total = ctx.dataset.data.reduce((a,b)=>a+b,0);
              const pct = total ? ((ctx.raw/total)*100).toFixed(2) : 0;
              return \` ₱\${ctx.raw.toLocaleString('en-PH',{minimumFractionDigits:2})}  (\${pct}%)\`;
            }
          }
        }
      }
    }
  });

  // ─── Top & Bottom Sub-Departments (horizontal bar) ─────────────────────
  if (cMoversChartInst) cMoversChartInst.destroy();
  const movCanvas = document.getElementById('cMoversChart');
  const moversN = IS_MOBILE ? 5 : 8;
  const topMovers = (movers.top || []).slice(0, moversN);
  const botMovers = (movers.bottom || []).slice(0, moversN);
  const combined = [...topMovers, ...botMovers];
  if (combined.length) {
    cMoversChartInst = new Chart(movCanvas, {
      type: 'bar',
      data: {
        labels: combined.map(m => m.subDepName.length > 28 ? m.subDepName.substring(0,28)+'…' : m.subDepName),
        datasets: [{
          label: 'Diff Amount',
          data: combined.map(m => m.diffVal),
          backgroundColor: combined.map(m => m.diffVal >= 0 ? 'rgba(16,185,129,0.85)' : 'rgba(244,63,94,0.85)'),
          borderColor: combined.map(m => m.diffVal >= 0 ? '#10b981' : '#f43f5e'),
          borderWidth: 1, borderRadius: 4, borderSkipped: false, barPercentage: 0.78,
        }]
      },
      options: {
        indexAxis: 'y', responsive: true, maintainAspectRatio: false,
        layout: { padding: { right: IS_MOBILE ? 70 : 90, left: IS_MOBILE ? 4 : 90 } },
        plugins: {
          legend: { display:false },
          datalabels: {
            display: true,
            anchor: 'end',
            align: 'end',
            clamp: true,
            offset: 4,
            color: ctx => ctx.dataset.data[ctx.dataIndex] >= 0 ? '#34d399' : '#fb7185',
            backgroundColor: 'rgba(10,14,26,0.85)',
            borderRadius: 4,
            padding: { top: 2, bottom: 2, left: 5, right: 5 },
            font: { family: "'JetBrains Mono'", size: IS_MOBILE ? 9.5 : 10.5, weight: '700' },
            formatter: v => {
              const abs = Math.abs(v); const sign = v < 0 ? '-' : '+';
              return abs >= 1e6 ? sign+'₱'+(abs/1e6).toFixed(2)+'M' : abs >= 1e3 ? sign+'₱'+(abs/1e3).toFixed(0)+'K' : sign+'₱'+abs.toFixed(0);
            }
          },
          tooltip: {
            backgroundColor:'rgba(15,20,35,0.95)', borderColor:'rgba(99,102,241,0.4)', borderWidth:1, padding:12, cornerRadius:10,
            titleFont:{family:"'Space Grotesk'",size:13,weight:'600'}, bodyFont:{family:"'JetBrains Mono'",size:11.5},
            titleColor:'#f0f3fb', bodyColor:'#a8b3d1',
            callbacks: {
              label: ctx => {
                const m = combined[ctx.dataIndex];
                return [
                  ' Diff: ₱'+ctx.raw.toLocaleString('en-PH',{minimumFractionDigits:2}),
                  ' Sales: ₱'+m.sales.toLocaleString('en-PH',{minimumFractionDigits:2}),
                  ' Sales LY: ₱'+m.salesLY.toLocaleString('en-PH',{minimumFractionDigits:2}),
                  ' Category: '+m.category
                ];
              },
              title: ctx => combined[ctx[0].dataIndex].subDepName
            }
          }
        },
        scales: {
          x: { grid:{color:'rgba(148,163,200,0.06)'}, ticks:{color:'#a8b3d1',font:{size:10.5,family:"'JetBrains Mono'",weight:'500'}, callback: v => { const abs=Math.abs(v); const sign=v<0?'-':''; return abs >= 1e6 ? sign+'₱'+(abs/1e6).toFixed(1)+'M' : abs >= 1e3 ? sign+'₱'+(abs/1e3).toFixed(0)+'K' : v; } } },
          y: { grid:{color:'rgba(148,163,200,0.04)',drawBorder:false}, ticks:{color:'#a8b3d1',font:{size:IS_MOBILE?10:10.5,family:"'Inter'",weight:'500'}, autoSkip: false, padding: IS_MOBILE ? 4 : 8 } }
        }
      }
    });
  }

  // ─── Sales by Area chart removed - replaced with table ─────────────────
}

function updateBreakdownSubDepDropdown() {
  const cat = document.getElementById('cBreakdownCategory').value;
  const sel = document.getElementById('cBreakdownSubDep');
  sel.innerHTML = '<option value="ALL">All Sub-Depts</option>';
  let subDeps = [];
  if (cat === 'ALL') {
    // All sub-depts across all categories
    const seen = new Set();
    Object.values(cSubDepsForBreakdown).forEach(arr => arr.forEach(sd => {
      if (!seen.has(sd)) { seen.add(sd); subDeps.push(sd); }
    }));
  } else {
    subDeps = cSubDepsForBreakdown[cat] || [];
  }
  subDeps.sort().forEach(sd => {
    const o = document.createElement('option');
    o.value = o.textContent = sd;
    sel.appendChild(o);
  });
}

async function applyBreakdownFilters() {
  // Re-fetch with breakdown filters applied to area/store aggregation
  const month = document.getElementById('cMonthFilter').value;
  const category = document.getElementById('cCategoryFilter').value;
  const area = document.getElementById('cAreaFilter').value;
  const store = document.getElementById('cStoreFilter').value;
  const bCat = document.getElementById('cBreakdownCategory').value;
  const bSub = document.getElementById('cBreakdownSubDep').value;

  // When breakdown category changes, refresh its sub-dept dropdown
  updateBreakdownSubDepDropdown();

  const params = new URLSearchParams();
  if (month !== 'ALL') params.set('month', month);
  if (category !== 'ALL') params.set('category', category);
  if (area !== 'ALL') params.set('area', area);
  if (store !== 'ALL') params.set('store', store);
  if (bCat !== 'ALL') params.set('breakdownCategory', bCat);
  if (bSub !== 'ALL') params.set('breakdownSubDep', bSub);

  try {
    const res = await fetch('/api/category-breakdown?' + params);
    const json = await res.json();
    if (!json.success) throw new Error(json.error);

    cAreaRowsCache = json.byArea || [];
    cStoreRowsCache = json.byStore || [];

    renderCAreaTable(sortRows(cAreaRowsCache, cAreaSort));
    renderCStoreTable(sortRows(cStoreRowsCache, cStoreSort));
    updateSortHeaders('cAreaTableBody', cAreaSort);
    updateSortHeaders('cStoreTableBody', cStoreSort);

    const parts = [];
    if (month !== 'ALL') parts.push(month);
    if (bCat !== 'ALL') parts.push(bCat);
    if (bSub !== 'ALL') parts.push(bSub);
    if (area !== 'ALL') parts.push(area);
    if (store !== 'ALL') parts.push(store);
    document.getElementById('cAreaTableInfo').textContent = parts.length ? parts.join(' · ') : 'All data';
    document.getElementById('cStoreTableInfo').textContent = parts.length ? parts.join(' · ') : 'All data';
  } catch(e) {
    console.error('Breakdown error:', e);
  }
}

function renderCAreaTable(rows) {
  const tbody = document.getElementById('cAreaTableBody');
  if (!rows.length) {
    tbody.innerHTML = \`<tr><td colspan="5" class="empty-cell"><div class="empty-icon"><i class="fa fa-magnifying-glass"></i></div><p>No area data found</p></td></tr>\`;
    return;
  }
  const totSales = rows.reduce((s,r)=>s+r.sales,0);
  const totLY = rows.reduce((s,r)=>s+r.salesLY,0);
  const totDiff = totSales - totLY;
  const totPct = totLY ? (totDiff/totLY*100) : 0;

  let html = rows.map(r => {
    const pctCls = r.diffPct > 0.05 ? 'up' : r.diffPct < -0.05 ? 'down' : 'flat';
    const color = AREA_COLORS[r.area] || DEFAULT_COLOR;
    const diffColor = r.diffVal >= 0 ? '#34d399' : '#fb7185';
    const arrow = r.diffPct > 0.05 ? '↑' : r.diffPct < -0.05 ? '↓' : '—';
    const pctStr = r.diffPct !== 0 ? \`\${arrow} \${Math.abs(r.diffPct).toFixed(2)}%\` : '—';
    return \`<tr>
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
    <td style="font-size:12px;letter-spacing:0.1em;text-transform:uppercase;color:var(--indigo2)">TOTAL · \${rows.length} AREAS</td>
    <td style="text-align:right"><span class="num">\${fmtFull(totSales)}</span></td>
    <td style="text-align:right"><span class="num" style="color:var(--text-3)">\${fmtFull(totLY)}</span></td>
    <td style="text-align:center"><span class="pill \${totPctCls}">\${totArrow} \${Math.abs(totPct).toFixed(2)}%</span></td>
    <td style="text-align:right"><span class="num" style="color:\${totDiff>=0?'#34d399':'#fb7185'}">\${totDiff>=0?'+':''}\${fmtFull(totDiff)}</span></td>
  </tr>\`;
  tbody.innerHTML = html;
}

function renderCStoreTable(rows) {
  const tbody = document.getElementById('cStoreTableBody');
  if (!rows.length) {
    tbody.innerHTML = \`<tr><td colspan="7" class="empty-cell"><div class="empty-icon"><i class="fa fa-magnifying-glass"></i></div><p>No store data found</p></td></tr>\`;
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
      <td><span class="num num-bold" style="color:var(--text-1)">#\${r.storeCode}</span></td>
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

function sortCArea(key, type) {
  if (cAreaSort.key === key) cAreaSort.dir = cAreaSort.dir === 'asc' ? 'desc' : 'asc';
  else cAreaSort = { key, dir: 'desc', type };
  cAreaSort.type = type;
  renderCAreaTable(sortRows(cAreaRowsCache, cAreaSort));
  updateSortHeaders('cAreaTableBody', cAreaSort);
}

function sortCStore(key, type) {
  if (cStoreSort.key === key) cStoreSort.dir = cStoreSort.dir === 'asc' ? 'desc' : 'asc';
  else cStoreSort = { key, dir: 'desc', type };
  cStoreSort.type = type;
  renderCStoreTable(sortRows(cStoreRowsCache, cStoreSort));
  updateSortHeaders('cStoreTableBody', cStoreSort);
}

function sortCSummary(key, type) {
  if (cSummarySort.key === key) cSummarySort.dir = cSummarySort.dir === 'asc' ? 'desc' : 'asc';
  else cSummarySort = { key, dir: 'desc', type };
  cSummarySort.type = type;
  renderCSummary(sortRows(cSummaryRowsCache, cSummarySort));
  updateSortHeaders('cSummaryBody', cSummarySort);
}

function sortCDetail(key, type) {
  if (cDetailSort.key === key) cDetailSort.dir = cDetailSort.dir === 'asc' ? 'desc' : 'asc';
  else cDetailSort = { key, dir: 'desc', type };
  cDetailSort.type = type;
  renderCDetail(sortRows(cDetailRowsCache, cDetailSort));
  updateSortHeaders('cDetailBody', cDetailSort);
}

function exportCategoryToExcel() {
  if (!cDetailRowsCache.length) {
    alert('No data to export with the current filters.');
    return;
  }
  const month = document.getElementById('cMonthFilter').value;
  const category = document.getElementById('cCategoryFilter').value;
  const area = document.getElementById('cAreaFilter').value;
  const store = document.getElementById('cStoreFilter').value;

  const data = cDetailRowsCache.map(r => ({
    'Category': r.category || '',
    'SDep Code': r.sdepCode || '',
    'Sub-Department': r.subDepName || '',
    'Sales': r.sales || 0,
    'Sales LY': r.salesLY || 0,
    'Diff %': r.diffPct || 0,
    'Diff Amount': r.diffVal || 0,
  }));

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(data);
  ws['!cols'] = [{wch:20},{wch:14},{wch:36},{wch:16},{wch:16},{wch:10},{wch:16}];

  const headerStyle = { fill:{fgColor:{rgb:'166534'}}, font:{color:{rgb:'FFFFFF'},bold:true}, alignment:{horizontal:'center',vertical:'center'} };
  const range = XLSX.utils.decode_range(ws['!ref']);
  for (let C = range.s.c; C <= range.e.c; C++) {
    const addr = XLSX.utils.encode_cell({ r: 0, c: C });
    if (ws[addr]) ws[addr].s = headerStyle;
  }
  const numCols = { 3:'#,##0.00', 4:'#,##0.00', 5:'0.00"%"', 6:'#,##0.00' };
  for (let R = 1; R <= range.e.r; R++) {
    for (const [col, fmt] of Object.entries(numCols)) {
      const addr = XLSX.utils.encode_cell({ r: R, c: parseInt(col) });
      if (ws[addr]) ws[addr].z = fmt;
    }
  }
  XLSX.utils.book_append_sheet(wb, ws, 'Category Detail');

  const parts = ['CaMaNaVa_Category'];
  if (month !== 'ALL') parts.push(month.replace(/[^a-z0-9]/gi,'_'));
  if (category !== 'ALL') parts.push(category.replace(/[^a-z0-9]/gi,'_'));
  if (area !== 'ALL') parts.push(area.replace(/[^a-z0-9]/gi,'_'));
  if (store !== 'ALL') parts.push(store.replace(/[^a-z0-9]/gi,'_'));
  if (cSignFilter !== 'ALL') parts.push(cSignFilter === 'POS' ? 'Positive' : 'Negative');
  XLSX.writeFile(wb, parts.join('_') + '.xlsx');
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
