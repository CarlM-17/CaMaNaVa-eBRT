const express = require('express');
const { google } = require('googleapis');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// ─── Google Sheets Config ───────────────────────────────────────────────────
const SHEET_ID = process.env.SHEET_ID || '1PhRyLx2viByS1J-dOWWPwSmAZwNwzzaM6ATmbpQIR8w';
const SHEET_NAME = 'DailySales';
const DATA_RANGE = `${SHEET_NAME}!A23:R`;

function getAuthClient() {
  const credentials = process.env.GOOGLE_SERVICE_ACCOUNT_JSON
    ? JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON)
    : require('./service-account.json');

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
// Query params: date (YYYY-MM-DD), area, store
app.get('/api/sales', async (req, res) => {
  try {
    const auth = getAuthClient();
    const sheets = google.sheets({ version: 'v4', auth });

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: DATA_RANGE,
    });

    const rows = response.data.values || [];

    // Map each row to an object
    // Columns: A=Date, B=Month, C=Day, D=Day_YA, E=Holiday, F=Area,
    //          G=StoreID, H=StoreName, I=Sales, J=SalesLY, K=DiffVal,
    //          L=DiffPct, M=TRX, N=TRX_YA, O=Basket, P=Basket_YA,
    //          Q=DateRemark, R=Justification
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

    // Apply filters
    const { date, area, store } = req.query;
    if (date) data = data.filter((r) => r.date === date);
    if (area && area !== 'ALL') data = data.filter((r) => r.area === area);
    if (store && store !== 'ALL') data = data.filter((r) => r.storeName === store);

    // Aggregate by store
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
          dates: [],
        };
      }
      storeMap[key].sales += r.sales;
      storeMap[key].salesLY += r.salesLY;
      storeMap[key].trx += r.trx;
      storeMap[key].trxLY += r.trxLY;
      if (r.justification) storeMap[key].justifications.push(r.justification);
      if (r.date) storeMap[key].dates.push(r.date);
    });

    const result = Object.values(storeMap)
      .map((r) => {
        const diffVal = r.sales - r.salesLY;
        const diffPct = r.salesLY !== 0 ? (diffVal / r.salesLY) * 100 : 0;
        return {
          ...r,
          diffVal,
          diffPct: parseFloat(diffPct.toFixed(2)),
          justification: [...new Set(r.justifications)].join(' | '),
          justifications: undefined,
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
// Returns distinct areas and stores for populating dropdowns
app.get('/api/filters', async (req, res) => {
  try {
    const auth = getAuthClient();
    const sheets = google.sheets({ version: 'v4', auth });

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: DATA_RANGE,
    });

    const rows = response.data.values || [];
    const areas = new Set();
    const stores = new Set();

    rows.forEach((r) => {
      if (r[5]) areas.add(r[5].trim());
      if (r[7]) stores.add(r[7].trim());
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

// ─── Catch-all: serve frontend ───────────────────────────────────────────────
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.listen(PORT, () => {
  console.log(`CaMaNaVa eBRT running on port ${PORT}`);
});
