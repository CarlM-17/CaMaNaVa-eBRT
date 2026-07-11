const express = require('express');
const { google } = require('googleapis');
const crypto = require('crypto');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json());
app.use(express.urlencoded({ extended: false }));

const SHEET_ID = process.env.SHEET_ID || '1PhRyLx2viByS1J-dOWWPwSmAZwNwzzaM6ATmbpQIR8w';
const SHEET_NAME = 'DailySales';
const DATA_RANGE = `${SHEET_NAME}!A23:R`;
const STORE_LIST_RANGE = 'ListOfStores!A:E';
const CATEGORY_RANGE = 'CategorySales!A:J';
const STORE_NOTES_RANGE = 'StoreNotes!A:M';
const ISSUES_RANGE = 'StoreOpsIssuesAndConcerns!A5:R';
const USERS_RANGE = 'user!B2:E';
const INVENTORY_LOGS_RANGE = 'InventoryLogs!A6:Z';
const SESSION_COOKIE = 'camanava_session';
const SESSION_MAX_AGE_MS = 12 * 60 * 60 * 1000;
const SESSION_SECRET = process.env.SESSION_SECRET || process.env.GOOGLE_SERVICE_ACCOUNT_JSON || 'change-me-in-env';

function parseGoogleCredentials(raw) {
  const attempts = [
    raw,
    String(raw || '').replace(/^['"]|['"]$/g, ''),
  ];

  try {
    const decoded = Buffer.from(String(raw || ''), 'base64').toString('utf8');
    if (decoded && decoded.includes('{')) attempts.push(decoded);
  } catch (err) {}

  for (const value of attempts) {
    if (!value) continue;
    try {
      let credentials = JSON.parse(value);
      if (typeof credentials === 'string') credentials = JSON.parse(credentials);
      if (credentials && credentials.private_key) {
        credentials.private_key = String(credentials.private_key).replace(/\\n/g, '\n');
      }
      if (credentials && credentials.client_email && credentials.private_key) return credentials;
    } catch (err) {}
  }

  throw new Error('GOOGLE_SERVICE_ACCOUNT_JSON is invalid. Paste the full Google service account JSON as one Railway variable.');
}

function getAuthClient() {
  const rawCredentials = process.env.GOOGLE_SERVICE_ACCOUNT_JSON || process.env.GOOGLE_CREDENTIALS_JSON || process.env.GOOGLE_APPLICATION_CREDENTIALS_JSON;
  if (!rawCredentials) {
    throw new Error('GOOGLE_SERVICE_ACCOUNT_JSON env variable is not set');
  }
  const credentials = parseGoogleCredentials(rawCredentials);
  return new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
  });
}

function normalizeKey(value) {
  return String(value || '').trim().replace(/\s+/g, ' ').toLowerCase();
}

function compactKey(value) {
  return normalizeKey(value).replace(/[^a-z0-9]/g, '');
}

function parseAreaList(value) {
  return String(value || '')
    .split(/[,\n;|]+/)
    .map(v => v.trim())
    .filter(Boolean);
}

function parseCookies(req) {
  return String(req.headers.cookie || '').split(';').reduce((acc, part) => {
    const idx = part.indexOf('=');
    if (idx === -1) return acc;
    const key = part.slice(0, idx).trim();
    const value = part.slice(idx + 1).trim();
    if (key) acc[key] = decodeURIComponent(value);
    return acc;
  }, {});
}

function signSession(payload) {
  const body = Buffer.from(JSON.stringify(payload)).toString('base64url');
  const sig = crypto.createHmac('sha256', SESSION_SECRET).update(body).digest('base64url');
  return body + '.' + sig;
}

function verifySession(token) {
  try {
    if (!token || !token.includes('.')) return null;
    const [body, sig] = token.split('.');
    const expected = crypto.createHmac('sha256', SESSION_SECRET).update(body).digest('base64url');
    if (Buffer.byteLength(sig) !== Buffer.byteLength(expected)) return null;
    if (!crypto.timingSafeEqual(Buffer.from(sig), Buffer.from(expected))) return null;
    const payload = JSON.parse(Buffer.from(body, 'base64url').toString('utf8'));
    if (!payload.exp || Date.now() > payload.exp) return null;
    return payload;
  } catch (err) {
    return null;
  }
}

function getRequestUser(req) {
  const session = verifySession(parseCookies(req)[SESSION_COOKIE]);
  if (!session || !session.username) return null;
  return {
    username: session.username,
    level: normalizeKey(session.level || 'user') === 'admin' ? 'admin' : 'user',
    areas: Array.isArray(session.areas) ? session.areas : [],
  };
}

function setSessionCookie(res, user) {
  const token = signSession({
    username: user.username,
    level: user.level,
    areas: user.areas,
    exp: Date.now() + SESSION_MAX_AGE_MS,
  });
  const secure = process.env.NODE_ENV === 'production' ? '; Secure' : '';
  res.setHeader('Set-Cookie', `${SESSION_COOKIE}=${encodeURIComponent(token)}; HttpOnly; SameSite=Lax; Path=/; Max-Age=${Math.floor(SESSION_MAX_AGE_MS / 1000)}${secure}`);
}

function clearSessionCookie(res) {
  res.setHeader('Set-Cookie', `${SESSION_COOKIE}=; HttpOnly; SameSite=Lax; Path=/; Max-Age=0`);
}

function requireAuth(req, res, next) {
  const user = getRequestUser(req);
  if (!user) return res.status(401).json({ success: false, error: 'Authentication required' });
  req.user = user;
  next();
}

function userCanSeeAll(user) {
  return normalizeKey(user && user.level) === 'admin' || (user && user.areas || []).some(a => ['*', 'all'].includes(normalizeKey(a)));
}

function scopeRowsByArea(rows, user) {
  if (userCanSeeAll(user)) return rows;
  const allowed = new Set((user && user.areas || []).map(normalizeKey));
  return rows.filter(r => allowed.has(normalizeKey(r.area)));
}

function userAreaLabels(user) {
  return userCanSeeAll(user) ? ['ALL'] : (user.areas || []);
}

let usersCache = null, usersCacheTime = 0;
const USERS_CACHE_TTL = 60 * 1000;

async function getSheetUsers(sheets) {
  const now = Date.now();
  if (usersCache && (now - usersCacheTime) < USERS_CACHE_TTL) return usersCache;
  const response = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: USERS_RANGE });
  const rows = response.data.values || [];
  usersCache = rows.map(r => {
    const username = String(r[0] || '').trim();
    const password = String(r[1] || '').trim();
    const level = normalizeKey(r[2]) === 'admin' ? 'admin' : 'user';
    const areas = parseAreaList(r[3]);
    return { username, usernameKey: normalizeKey(username), password, level, areas };
  }).filter(u => u.username && u.password && (u.level === 'admin' || u.areas.length));
  usersCacheTime = now;
  return usersCache;
}

function parseNum(s) {
  if (!s || String(s).trim() === '') return 0;
  return parseFloat(String(s).replace(/,/g, '').replace('%', '')) || 0;
}

function parseDate(s) {
  if (!s) return null;
  const parts = String(s).trim().split('/');
  if (parts.length === 3) {
    const m = parts[0].padStart(2, '0');
    const d = parts[1].padStart(2, '0');
    const y = parts[2];
    return y + '-' + m + '-' + d;
  }
  return null;
}

function monthKey(dateStr) {
  if (!dateStr) return null;
  return dateStr.substring(0, 7);
}

function monthLabel(key) {
  if (!key) return '';
  const [y, m] = key.split('-');
  const monthNames = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  return monthNames[parseInt(m)-1] + ' ' + y;
}

function cleanMonthText(value) {
  return String(value || '')
    .normalize('NFKC')
    .replace(/Â/g, ' ')
    .replace(/\u00a0/g, ' ')
    .replace(/\uFFFD/g, '')
    .replace(/[^\S\r\n]+/g, ' ')
    .trim();
}

function categoryMonthKey(value) {
  const clean = cleanMonthText(value);
  if (!clean) return '';
  const lower = clean.toLowerCase();
  const direct = lower.match(/\b(20\d{2})[-/ ](0?[1-9]|1[0-2])\b/);
  if (direct) return direct[1] + '-' + String(parseInt(direct[2], 10)).padStart(2, '0');

  const monthMap = {
    jan: 1, january: 1, feb: 2, february: 2, mar: 3, march: 3, apr: 4, april: 4,
    may: 5, jun: 6, june: 6, jul: 7, july: 7, aug: 8, august: 8, sep: 9, sept: 9,
    september: 9, oct: 10, october: 10, nov: 11, november: 11, dec: 12, december: 12,
  };
  const yearMatch = lower.match(/\b(20\d{2})\b/);
  const monthMatch = lower.match(/\b(january|jan|february|feb|march|mar|april|apr|may|june|jun|july|jul|august|aug|september|sept|sep|october|oct|november|nov|december|dec)\b/);
  if (monthMatch) {
    const year = yearMatch ? yearMatch[1] : String(new Date().getFullYear());
    return year + '-' + String(monthMap[monthMatch[1]]).padStart(2, '0');
  }
  return normalizeKey(clean);
}

function categoryMonthLabel(value, key) {
  if (key && /^\d{4}-\d{2}$/.test(key)) return monthLabel(key);
  const clean = cleanMonthText(value);
  return clean.replace(/\w\S*/g, w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase());
}

let storeListCache = null, storeListCacheTime = 0;
let salesCache = null, salesCacheTime = 0;
let categoryCache = null, categoryCacheTime = 0;
let storeNotesCache = null, storeNotesCacheTime = 0;
let issuesCache = null, issuesCacheTime = 0;
let inventoryLogsCache = null, inventoryLogsCacheTime = 0;
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

function findHeaderIndex(headers, patterns, fallback) {
  const normalized = headers.map(h => normalizeKey(h));
  for (const pattern of patterns) {
    const idx = normalized.findIndex(h => pattern.test(h));
    if (idx >= 0) return idx;
  }
  return fallback;
}

async function getInventoryLogsData(sheets) {
  const now = Date.now();
  if (inventoryLogsCache && (now - inventoryLogsCacheTime) < CACHE_TTL) return inventoryLogsCache;
  const response = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: INVENTORY_LOGS_RANGE });
  const rows = response.data.values || [];
  if (!rows.length) return [];

  const headers = rows[0] || [];
  const dateIdx = 0;
  const areaIdx = findHeaderIndex(headers, [/^area$/, /area/], 1);
  const storeIdIdx = findHeaderIndex(headers, [/store.*id/, /store.*code/, /^id$/], 2);
  const storeNameIdx = findHeaderIndex(headers, [/store.*name/, /^store$/, /branch/], 3);

  inventoryLogsCache = rows.slice(1).map(r => {
    const dateRaw = (r[dateIdx] || '').trim();
    return {
      date: parseDate(dateRaw) || parseDateFlexible(dateRaw)?.toISOString().slice(0, 10) || null,
      area: (r[areaIdx] || '').trim(),
      storeId: (r[storeIdIdx] || '').trim(),
      storeName: (r[storeNameIdx] || '').trim(),
    };
  }).filter(r => r.date && (r.storeId || r.storeName));
  inventoryLogsCacheTime = now;
  return inventoryLogsCache;
}

function maxISODate(a, b) {
  if (!a) return b;
  if (!b) return a;
  return a > b ? a : b;
}

function storeAuditStartDate(store, defaultStart) {
  const name = normalizeKey(store && store.storeName);
  const storeId = normalizeKey(store && store.storeId);
  if (storeId === '867' || name.includes('congressional') || name.includes('cogressional')) return maxISODate(defaultStart, '2026-06-24');
  return defaultStart;
}

function isCongressionalStore(store) {
  const name = normalizeKey(store && store.storeName);
  const storeId = normalizeKey(store && store.storeId);
  return storeId === '867' || name.includes('congressional') || name.includes('cogressional');
}

function monthEndsBeforeStoreOpened(monthKey, store) {
  if (!isCongressionalStore(store) || !/^\d{4}-\d{2}$/.test(monthKey || '')) return false;
  const [year, month] = monthKey.split('-').map(n => parseInt(n, 10));
  const monthEnd = year + '-' + String(month).padStart(2, '0') + '-' + String(new Date(year, month, 0).getDate()).padStart(2, '0');
  return monthEnd < '2026-06-24';
}

function daysInMonthWindow(year, monthIndex, startDate, endDate) {
  const monthStart = year + '-' + String(monthIndex + 1).padStart(2, '0') + '-01';
  const monthEnd = year + '-' + String(monthIndex + 1).padStart(2, '0') + '-' + String(new Date(year, monthIndex + 1, 0).getDate()).padStart(2, '0');
  const start = maxISODate(monthStart, startDate);
  const end = monthEnd < endDate ? monthEnd : endDate;
  if (!start || !end || start > end) return [];
  const days = [];
  const d = new Date(start + 'T00:00:00');
  const stop = new Date(end + 'T00:00:00');
  while (d <= stop) {
    days.push(formatISODateLocal(d));
    d.setDate(d.getDate() + 1);
  }
  return days;
}

function buildStoreDayGapMonitor(logRows, masterStores, filters = {}, options = {}) {
  const now = new Date();
  const endDateObj = options.includeToday
    ? new Date(now.getFullYear(), now.getMonth(), now.getDate())
    : new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1);
  const endDate = formatISODateLocal(endDateObj);
  const year = endDateObj.getFullYear();
  const startDate = options.startDate || (year + '-01-01');
  const selectedMonth = filters.month || 'ALL';
  const area = filters.area || 'ALL';
  const store = filters.store || 'ALL';
  const label = options.label || 'data';

  let expectedStores = masterStores;
  if (area && area !== 'ALL') expectedStores = expectedStores.filter(s => s.area === area);
  if (store && store !== 'ALL') expectedStores = expectedStores.filter(s => s.storeName === store);

  const reported = new Set();
  logRows.forEach(r => {
    if (!r.date || String(r.date).slice(0, 4) !== String(year)) return;
    if (r.storeId) reported.add('id:' + normalizeKey(r.storeId) + '|' + r.date);
    if (r.storeName) {
      reported.add('name:' + normalizeKey(r.storeName) + '|' + r.date);
      reported.add('cname:' + compactKey(r.storeName) + '|' + r.date);
    }
  });

  const months = [];
  for (let m = 0; m <= endDateObj.getMonth(); m++) {
    const key = year + '-' + String(m + 1).padStart(2, '0');
    if (selectedMonth !== 'ALL' && selectedMonth !== key) continue;
    months.push({ key, label: monthLabel(key) });
  }

  const storeMonthRows = [];
  months.forEach(m => {
    expectedStores.forEach(s => {
      const storeStart = options.applyStoreOpenDate === false ? startDate : storeAuditStartDate(s, startDate);
      const days = daysInMonthWindow(year, parseInt(m.key.slice(5, 7), 10) - 1, storeStart, endDate);
      if (!days.length) return;
      const missingDates = days.filter(date =>
        !reported.has('id:' + normalizeKey(s.storeId) + '|' + date) &&
        !reported.has('name:' + normalizeKey(s.storeName) + '|' + date) &&
        !reported.has('cname:' + compactKey(s.storeName) + '|' + date)
      );
      const expectedDays = days.length;
      const gapDays = missingDates.length;
      const reportedDays = expectedDays - gapDays;
      const completionPct = expectedDays ? (reportedDays / expectedDays) * 100 : 0;
      storeMonthRows.push({
        month: m.label,
        monthKey: m.key,
        area: s.area,
        storeId: s.storeId,
        storeName: s.storeName,
        expectedDays,
        reportedDays,
        gapDays,
        completionPct: parseFloat(completionPct.toFixed(2)),
        missingDates,
        status: gapDays ? 'With Gap' : 'Complete',
        remarks: gapDays ? (gapDays + ' day' + (gapDays !== 1 ? 's' : '') + ' missing') : 'Complete ' + label,
      });
    });
  });

  const gapRows = storeMonthRows.filter(r => r.gapDays > 0).sort((a, b) => b.gapDays - a.gapDays || a.monthKey.localeCompare(b.monthKey) || (a.storeName || '').localeCompare(b.storeName || ''));
  const completeRows = storeMonthRows.filter(r => r.gapDays === 0).sort((a, b) => a.monthKey.localeCompare(b.monthKey) || (a.area || '').localeCompare(b.area || '') || (a.storeName || '').localeCompare(b.storeName || ''));
  const monthly = months.map(m => {
    const rows = storeMonthRows.filter(r => r.monthKey === m.key);
    const expectedStoreDays = rows.reduce((sum, r) => sum + r.expectedDays, 0);
    const gapDays = rows.reduce((sum, r) => sum + r.gapDays, 0);
    return {
      month: m.label,
      monthKey: m.key,
      expectedStoreDays,
      reportedStoreDays: expectedStoreDays - gapDays,
      gapDays,
      completeStores: rows.filter(r => r.gapDays === 0).length,
      storesWithGaps: rows.filter(r => r.gapDays > 0).length,
      completionPct: expectedStoreDays ? parseFloat((((expectedStoreDays - gapDays) / expectedStoreDays) * 100).toFixed(2)) : 0,
    };
  });
  const totalExpectedStoreDays = monthly.reduce((sum, r) => sum + r.expectedStoreDays, 0);
  const totalGapDays = monthly.reduce((sum, r) => sum + r.gapDays, 0);
  return {
    summary: {
      year,
      throughDate: endDate,
      storeMonths: storeMonthRows.length,
      gapStoreMonths: gapRows.length,
      completeStoreMonths: completeRows.length,
      totalExpectedStoreDays,
      totalReportedStoreDays: totalExpectedStoreDays - totalGapDays,
      totalGapDays,
      completionPct: totalExpectedStoreDays ? parseFloat((((totalExpectedStoreDays - totalGapDays) / totalExpectedStoreDays) * 100).toFixed(2)) : 0,
    },
    monthly,
    gapRows,
    completeRows,
  };
}
async function getCategoryData(sheets) {
  const now = Date.now();
  if (categoryCache && (now - categoryCacheTime) < CACHE_TTL) return categoryCache;
  const response = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: CATEGORY_RANGE });
  const rows = response.data.values || [];
  const data = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i] || [];
    const monthRaw  = cleanMonthText(r[0]);
    const monthKeyValue = categoryMonthKey(monthRaw);
    const month     = categoryMonthLabel(monthRaw, monthKeyValue);
    const area      = (r[1] || '').trim();
    const storeCode = (r[2] || '').trim();
    const storeName = (r[3] || '').trim();
    const sdepCode  = (r[4] || '').trim();
    const subDepName= (r[5] || '').trim();
    const sales     = parseNum(r[6]);
    const salesLY   = parseNum(r[7]);
    const category  = (r[8] || '').trim();
    const lastUpdate = (r[9] || '').trim();
    if (!monthRaw && !storeCode && !category && !subDepName) continue;
    if (normalizeKey(monthRaw) === 'month' || normalizeKey(storeCode) === 'store code' || normalizeKey(storeName) === 'store name') continue;
    data.push({ month, monthRaw, monthKey: monthKeyValue, area, storeCode, storeName, sdepCode, subDepName, sales, salesLY, category, lastUpdate });
  }
  categoryCache = data;
  categoryCacheTime = now;
  return categoryCache;
}

function lastUpdateSortValue(value) {
  const parsed = parseDateFlexible(value);
  if (parsed) return parsed.getTime();
  const d = new Date(value);
  return isNaN(d.getTime()) ? 0 : d.getTime();
}

function buildDailySalesDiffPctLookup(salesRows, filters = {}) {
  const selectedMonthKey = filters.monthKey || '';
  const area = filters.area || 'ALL';
  const store = filters.store || 'ALL';
  let rows = salesRows || [];
  if (selectedMonthKey) rows = rows.filter(r => monthKey(r.date) === selectedMonthKey);
  if (area && area !== 'ALL') rows = rows.filter(r => r.area === area);
  if (store && store !== 'ALL') rows = rows.filter(r => r.storeName === store);

  const map = new Map();
  rows.forEach(r => {
    if (!r.storeName && !r.storeId) return;
    const keys = [];
    if (r.storeId) keys.push('id:' + normalizeKey(r.storeId));
    if (r.storeName) keys.push('name:' + compactKey(r.storeName));
    const key = keys[0];
    const current = map.get(key) || { sales: 0, salesLY: 0, keys };
    current.sales += Number(r.sales || 0);
    current.salesLY += Number(r.salesLY || 0);
    keys.forEach(k => { if (!current.keys.includes(k)) current.keys.push(k); });
    map.set(key, current);
  });

  const lookup = new Map();
  map.forEach(v => {
    const diffValue = v.sales - v.salesLY;
    const diffPct = v.salesLY ? (diffValue / v.salesLY) * 100 : (v.sales ? 100 : null);
    v.keys.forEach(k => lookup.set(k, diffPct === null ? null : parseFloat(diffPct.toFixed(2))));
  });
  return lookup;
}

function buildCategoryLastUpdateRows(categoryRows, filters = {}, salesRows = []) {
  const selectedMonthKey = filters.monthKey || '';
  const area = filters.area || 'ALL';
  const store = filters.store || 'ALL';
  let rows = categoryRows;
  if (selectedMonthKey) rows = rows.filter(r => r.monthKey === selectedMonthKey);
  if (area && area !== 'ALL') rows = rows.filter(r => r.area === area);
  if (store && store !== 'ALL') rows = rows.filter(r => r.storeName === store);

  const dailyDiffLookup = buildDailySalesDiffPctLookup(salesRows, filters);
  const map = new Map();
  rows.forEach(r => {
    if (!r.storeName && !r.storeCode) return;
    const key = (r.storeCode ? 'id:' + normalizeKey(r.storeCode) : 'name:' + compactKey(r.storeName));
    const nameKey = r.storeName ? 'name:' + compactKey(r.storeName) : '';
    const current = map.get(key) || {
      storeName: r.storeName || r.storeCode || '',
      sales: 0,
      salesYA: 0,
      dailySalesDiffPct: dailyDiffLookup.has(key) ? dailyDiffLookup.get(key) : (nameKey && dailyDiffLookup.has(nameKey) ? dailyDiffLookup.get(nameKey) : null),
      lastUpdate: '',
      lastUpdateTs: 0,
    };
    current.sales += Number(r.sales || 0);
    current.salesYA += Number(r.salesLY || 0);
    const dsDiff = dailyDiffLookup.has(key) ? dailyDiffLookup.get(key) : (nameKey && dailyDiffLookup.has(nameKey) ? dailyDiffLookup.get(nameKey) : null);
    if (dsDiff !== null && dsDiff !== undefined) current.dailySalesDiffPct = dsDiff;
    const updateTs = lastUpdateSortValue(r.lastUpdate);
    if (updateTs >= (current.lastUpdateTs || 0)) {
      current.lastUpdate = r.lastUpdate || current.lastUpdate || '';
      current.lastUpdateTs = updateTs;
    }
    map.set(key, current);
  });
  return [...map.values()].sort((a, b) => (b.lastUpdateTs || 0) - (a.lastUpdateTs || 0) || (a.storeName || '').localeCompare(b.storeName || ''));
}
function buildCategoryMonitorRows(categoryRows, masterStores, filters = {}) {
  return buildCategoryMonitorStatus(categoryRows, masterStores, filters).gapRows;
}

function buildCategoryMonitorStatus(categoryRows, masterStores, filters = {}) {
  const selectedMonthKey = filters.monthKey || '';
  const area = filters.area || 'ALL';
  const store = filters.store || 'ALL';
  const monthMap = new Map();
  categoryRows.forEach(r => {
    if (r.monthKey && !monthMap.has(r.monthKey)) monthMap.set(r.monthKey, r.month || categoryMonthLabel(r.monthRaw, r.monthKey));
  });
  const now = new Date();
  const yesterday = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1);
  const monitorYear = yesterday.getFullYear();
  for (let m = 0; m <= yesterday.getMonth(); m++) {
    const key = monitorYear + '-' + String(m + 1).padStart(2, '0');
    if (!monthMap.has(key)) monthMap.set(key, monthLabel(key));
  }
  const months = selectedMonthKey
    ? [{ key: selectedMonthKey, label: monthMap.get(selectedMonthKey) || categoryMonthLabel(selectedMonthKey, selectedMonthKey) }]
    : [...monthMap.entries()].map(([key, label]) => ({ key, label }));

  let expectedStores = masterStores;
  if (area && area !== 'ALL') expectedStores = expectedStores.filter(s => s.area === area);
  if (store && store !== 'ALL') expectedStores = expectedStores.filter(s => s.storeName === store);

  const gapRows = [];
  const completeRows = [];
  months.forEach(m => {
    const actual = new Set();
    categoryRows
      .filter(r => r.monthKey === m.key)
      .forEach(r => {
        if (r.storeCode) actual.add('id:' + normalizeKey(r.storeCode));
        if (r.storeName) {
          actual.add('name:' + normalizeKey(r.storeName));
          actual.add('cname:' + compactKey(r.storeName));
        }
      });
    expectedStores.forEach(s => {
      if (monthEndsBeforeStoreOpened(m.key, s)) return;
      const hasData = actual.has('id:' + normalizeKey(s.storeId)) || actual.has('name:' + normalizeKey(s.storeName)) || actual.has('cname:' + compactKey(s.storeName));
      if (!hasData) {
        gapRows.push({
          month: m.label,
          monthKey: m.key,
          area: s.area,
          storeId: s.storeId,
          storeName: s.storeName,
          remarks: 'No Category Sales data found for this month',
        });
      } else {
        completeRows.push({
          month: m.label,
          monthKey: m.key,
          area: s.area,
          storeId: s.storeId,
          storeName: s.storeName,
          remarks: 'Complete Category Sales data found for this month',
        });
      }
    });
  });
  const sortMonitor = (a, b) => a.monthKey.localeCompare(b.monthKey) || (a.area || '').localeCompare(b.area || '') || (a.storeName || '').localeCompare(b.storeName || '');
  return {
    gapRows: gapRows.sort(sortMonitor),
    completeRows: completeRows.sort(sortMonitor),
  };
}

function formatISODateLocal(date) {
  const yyyy = date.getFullYear();
  const mm = String(date.getMonth() + 1).padStart(2, '0');
  const dd = String(date.getDate()).padStart(2, '0');
  return yyyy + '-' + mm + '-' + dd;
}

function daysInMonthUntil(year, monthIndex, endDate) {
  const lastDay = monthIndex === endDate.getMonth() && year === endDate.getFullYear()
    ? endDate.getDate()
    : new Date(year, monthIndex + 1, 0).getDate();
  const days = [];
  for (let d = 1; d <= lastDay; d++) days.push(year + '-' + String(monthIndex + 1).padStart(2, '0') + '-' + String(d).padStart(2, '0'));
  return days;
}

function buildDailySalesGapMonitor(salesRows, masterStores, filters = {}) {
  const now = new Date();
  const yesterday = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1);
  return buildStoreDayGapMonitor(salesRows, masterStores, filters, {
    startDate: yesterday.getFullYear() + '-01-01',
    includeToday: false,
    label: 'Daily Sales data',
  });
}

async function getStoreNotesData(sheets) {
  const now = Date.now();
  if (storeNotesCache && (now - storeNotesCacheTime) < CACHE_TTL) return storeNotesCache;

  // Fetch BOTH values and hyperlinks for column M (so we get the URL behind "Open Photo")
  const [valsResp, linksResp] = await Promise.all([
    sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: STORE_NOTES_RANGE }),
    sheets.spreadsheets.get({
      spreadsheetId: SHEET_ID,
      ranges: [STORE_NOTES_RANGE],
      fields: 'sheets.data.rowData.values(hyperlink,formattedValue,userEnteredValue,textFormatRuns(format/link))'
    })
  ]);

  const rows = valsResp.data.values || [];
  const richRows = ((((linksResp.data.sheets || [])[0] || {}).data || [])[0] || {}).rowData || [];

  const data = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i] || [];
    const id          = (r[0] || '').trim();
    const timestamp   = (r[1] || '').trim();
    const storeId     = (r[2] || '').trim();
    const storeName   = (r[3] || '').trim();
    const area        = (r[4] || '').trim();
    const notes       = (r[5] || '').trim();
    const photo       = (r[6] || '').trim();
    const photoDraft  = (r[7] || '').trim();
    const submittedBy = (r[8] || '').trim();
    const photoLinkJ  = (r[9] || '').trim();
    const status      = (r[10] || '').trim();
    const remarks     = (r[11] || '').trim();
    const photoLinkM  = (r[12] || '').trim();

    if (!timestamp && !storeId && !notes) continue;

    // Extract URL from column M (index 12) - it can be either a hyperlink, textFormatRun link, or HYPERLINK formula
    let photoUrl = null;
    const cellM = ((richRows[i] || {}).values || [])[12];
    if (cellM) {
      // Try direct hyperlink property first
      if (cellM.hyperlink) {
        photoUrl = cellM.hyperlink;
      }
      // Try text format runs (rich text hyperlinks)
      else if (cellM.textFormatRuns && cellM.textFormatRuns.length) {
        for (const run of cellM.textFormatRuns) {
          if (run.format && run.format.link && run.format.link.uri) {
            photoUrl = run.format.link.uri;
            break;
          }
        }
      }
      // Try HYPERLINK() formula
      else if (cellM.userEnteredValue && cellM.userEnteredValue.formulaValue) {
        const formula = cellM.userEnteredValue.formulaValue;
        const m = formula.match(/HYPERLINK\(\s*"([^"]+)"/i);
        if (m) photoUrl = m[1];
      }
    }

    // Fallback to column G if M didn't have a link but G has a URL
    if (!photoUrl && photo && /^https?:\/\//i.test(photo)) {
      photoUrl = photo;
    }

    // Parse timestamp to numeric epoch for sorting (0 if unparseable so they sort last on desc)
    let ts = 0;
    if (timestamp) {
      const d = new Date(timestamp);
      if (!isNaN(d.getTime())) ts = d.getTime();
    }

    data.push({
      id, timestamp, ts,
      storeId, storeName, area, notes,
      photo, photoDraft, submittedBy,
      photoLink: photoLinkJ, status, remarks,
      photoLinkM, photoUrl
    });
  }
  storeNotesCache = data;
  storeNotesCacheTime = now;
  return storeNotesCache;
}

function parseDateFlexible(s) {
  if (!s) return null;
  const t = String(s).trim();
  if (!t) return null;
  // Try standard parse first
  let d = new Date(t);
  if (!isNaN(d.getTime())) return d;
  // Try M/D/YYYY format
  const parts = t.split('/');
  if (parts.length === 3) {
    const m = parseInt(parts[0], 10);
    const day = parseInt(parts[1], 10);
    const y = parseInt(parts[2], 10);
    if (!isNaN(m) && !isNaN(day) && !isNaN(y)) {
      d = new Date(y, m - 1, day);
      if (!isNaN(d.getTime())) return d;
    }
  }
  return null;
}

function normalizeFilterText(value) {
  return String(value || '').trim().replace(/\s+/g, ' ').toLowerCase();
}

function addCanonicalFilterValue(map, value) {
  const label = String(value || '').trim().replace(/\s+/g, ' ');
  const key = normalizeFilterText(label);
  if (key && !map.has(key)) map.set(key, label);
}

function parseMultiFilterParam(value) {
  const rawValues = Array.isArray(value) ? value : [value];
  return new Set(rawValues
    .filter(v => v != null)
    .flatMap(v => String(v).split(','))
    .map(normalizeFilterText)
    .filter(v => v && v !== 'all'));
}

function prioritySortRank(value) {
  const s = normalizeFilterText(value);
  if (s.includes('critic')) return 1;
  if (s.includes('high') || s === '1' || s === 'p1') return 2;
  if (s.includes('med') || s === '2' || s === 'p2') return 3;
  if (s.includes('low') || s === '3' || s === 'p3') return 4;
  return 5;
}

async function getIssuesData(sheets) {
  const now = Date.now();
  if (issuesCache && (now - issuesCacheTime) < CACHE_TTL) return issuesCache;
  const response = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: ISSUES_RANGE });
  const rows = response.data.values || [];
  const data = [];
  const today = new Date(); today.setHours(0, 0, 0, 0);

  // Row 0 in the response = sheet row 5 (header). Skip it.
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i] || [];
    // Columns A..R: Area, Store ID, Store Name, Date, Reported By, Issue Category,
    // Issue Sub Category, Issue Description, Priority, Impact Level, Assign To,
    // Target Resolution Date, Status, Resolution Details, Date Resolved,
    // Days Open, Remarks/Notes, Last Update
    const area              = (r[0] || '').trim();
    const storeId           = (r[1] || '').trim();
    const storeName         = (r[2] || '').trim();
    const dateRaw           = (r[3] || '').trim();
    const reportedBy        = (r[4] || '').trim();
    const issueCategory     = (r[5] || '').trim();
    const issueSubCategory  = (r[6] || '').trim();
    const issueDescription  = (r[7] || '').trim();
    const priority          = (r[8] || '').trim();
    const impactLevel       = (r[9] || '').trim();
    const assignTo          = (r[10] || '').trim();
    const targetDateRaw     = (r[11] || '').trim();
    const status            = (r[12] || '').trim();
    const resolutionDetails = (r[13] || '').trim();
    const dateResolvedRaw   = (r[14] || '').trim();
    const daysOpenRaw       = (r[15] || '').trim();
    const remarks           = (r[16] || '').trim();
    const lastUpdate        = (r[17] || '').trim();

    // skip blank rows
    if (!area && !storeId && !storeName && !issueCategory && !issueDescription) continue;

    const dateObj          = parseDateFlexible(dateRaw);
    const targetDateObj    = parseDateFlexible(targetDateRaw);
    const dateResolvedObj  = parseDateFlexible(dateResolvedRaw);

    // Compute days open server-side if not present in sheet
    let daysOpen = parseInt(daysOpenRaw, 10);
    if (isNaN(daysOpen)) {
      const start = dateObj;
      const end = dateResolvedObj || today;
      if (start) {
        const ms = end.getTime() - start.getTime();
        daysOpen = Math.max(0, Math.floor(ms / 86400000));
      } else {
        daysOpen = 0;
      }
    }

    // Overdue if not resolved and target date has passed
    const sLower = status.toLowerCase();
    const isResolved = /resolv|closed|done|complete/.test(sLower);
    const overdue = !isResolved && targetDateObj && targetDateObj < today;

    data.push({
      area, storeId, storeName,
      date: dateRaw,
      dateTs: dateObj ? dateObj.getTime() : 0,
      reportedBy,
      issueCategory, issueSubCategory, issueDescription,
      priority, impactLevel, assignTo,
      targetDate: targetDateRaw,
      targetDateTs: targetDateObj ? targetDateObj.getTime() : 0,
      status,
      resolutionDetails,
      dateResolved: dateResolvedRaw,
      dateResolvedTs: dateResolvedObj ? dateResolvedObj.getTime() : 0,
      daysOpen,
      remarks,
      lastUpdate,
      isResolved,
      overdue,
    });
  }
  issuesCache = data;
  issuesCacheTime = now;
  return issuesCache;
}

// ─── API: GET /api/sales ─────────────────────────────────────────────────────
app.post('/login', async (req, res) => {
  try {
    const username = String(req.body.username || '').trim();
    const password = String(req.body.password || '').trim();
    const auth = getAuthClient();
    const sheets = google.sheets({ version: 'v4', auth });
    const users = await getSheetUsers(sheets);
    const user = users.find(u => u.usernameKey === normalizeKey(username));
    if (!user) return res.redirect('/?error=no_user');
    if (String(user.password) !== password) return res.redirect('/?error=bad_password');
    setSessionCookie(res, user);
    res.redirect('/');
  } catch (err) {
    console.error('Login error:', err.message);
    res.redirect('/?error=config');
  }
});

app.post('/logout', (req, res) => {
  clearSessionCookie(res);
  res.redirect('/');
});

app.use('/api', requireAuth);

app.get('/api/me', (req, res) => {
  res.json({
    success: true,
    user: {
      username: req.user.username,
      level: req.user.level,
      areas: userAreaLabels(req.user),
    },
  });
});

app.get('/api/sales', async (req, res) => {
  try {
    const auth = getAuthClient();
    const sheets = google.sheets({ version: 'v4', auth });
    let [data, masterStores] = await Promise.all([getSalesData(sheets), getMasterStoreList(sheets)]);
    data = scopeRowsByArea(data, req.user);
    masterStores = scopeRowsByArea(masterStores, req.user);
    const { date, area, store } = req.query;
    let filtered = data;
    if (date) filtered = filtered.filter((r) => r.date === date);
    if (area && area !== 'ALL') filtered = filtered.filter((r) => r.area === area);
    if (store && store !== 'ALL') filtered = filtered.filter((r) => r.storeName === store);

    const storeMap = {};
    filtered.forEach((r) => {
      const key = r.storeId + '_' + r.storeName;
      if (!storeMap[key]) storeMap[key] = { storeId: r.storeId, storeName: r.storeName, area: r.area, sales: 0, salesLY: 0, trx: 0, trxLY: 0, justifications: [] };
      storeMap[key].sales += r.sales;
      storeMap[key].salesLY += r.salesLY;
      storeMap[key].trx += r.trx;
      storeMap[key].trxLY += r.trxLY;
      if (r.justification) storeMap[key].justifications.push(r.justification);
    });

    const result = Object.values(storeMap).map((r) => {
      const diffVal = r.sales - r.salesLY;
      const diffPct = r.salesLY !== 0 ? (diffVal / r.salesLY) * 100 : 0;
      return {
        storeId: r.storeId, storeName: r.storeName, area: r.area,
        sales: r.sales, salesLY: r.salesLY, trx: r.trx, trxLY: r.trxLY,
        diffVal, diffPct: parseFloat(diffPct.toFixed(2)),
        justification: [...new Set(r.justifications)].join(' | '),
      };
    }).sort((a, b) => b.sales - a.sales);

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
    console.error('Sales API error:', err.message);
    res.status(500).json({ success: false, error: err.message });
  }
});

app.get('/api/filters', async (req, res) => {
  try {
    const auth = getAuthClient();
    const sheets = google.sheets({ version: 'v4', auth });
    let masterStores = await getMasterStoreList(sheets);
    masterStores = scopeRowsByArea(masterStores, req.user);
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

app.get('/api/months', async (req, res) => {
  try {
    const auth = getAuthClient();
    const sheets = google.sheets({ version: 'v4', auth });
    let data = await getSalesData(sheets);
    data = scopeRowsByArea(data, req.user);
    const monthSet = new Set();
    data.forEach(r => { const k = monthKey(r.date); if (k) monthSet.add(k); });
    const months = [...monthSet].sort().map(k => ({ value: k, label: monthLabel(k) }));
    res.json({ success: true, months });
  } catch (err) {
    console.error('Months error:', err.message);
    res.status(500).json({ success: false, error: err.message });
  }
});

app.get('/api/monthly', async (req, res) => {
  try {
    const auth = getAuthClient();
    const sheets = google.sheets({ version: 'v4', auth });
    let data = await getSalesData(sheets);
    data = scopeRowsByArea(data, req.user);
    const { month, area, store, sign } = req.query;
    const monthSet = parseMultiFilterParam(month);

    let filtered = data;
    if (monthSet.size) filtered = filtered.filter(r => monthSet.has(normalizeKey(monthKey(r.date))));
    if (area && area !== 'ALL') filtered = filtered.filter(r => r.area === area);
    if (store && store !== 'ALL') filtered = filtered.filter(r => r.storeName === store);

    const storeMap = {};
    filtered.forEach(r => {
      const key = r.storeId + '_' + r.storeName;
      if (!storeMap[key]) storeMap[key] = { storeId: r.storeId, storeName: r.storeName, area: r.area, sales: 0, salesLY: 0 };
      storeMap[key].sales += r.sales;
      storeMap[key].salesLY += r.salesLY;
    });
    const summary = Object.values(storeMap).map(r => {
      const diffVal = r.sales - r.salesLY;
      const diffPct = r.salesLY !== 0 ? (diffVal / r.salesLY) * 100 : 0;
      return { ...r, diffVal, diffPct: parseFloat(diffPct.toFixed(2)) };
    }).sort((a, b) => b.sales - a.sales);

    let detailRecords = filtered;
    if (sign === 'POS') detailRecords = detailRecords.filter(r => (r.sales - r.salesLY) > 0);
    if (sign === 'NEG') detailRecords = detailRecords.filter(r => (r.sales - r.salesLY) < 0);

    const detail = detailRecords.map(r => {
      const diffVal = r.sales - r.salesLY;
      const diffPct = r.salesLY !== 0 ? (diffVal / r.salesLY) * 100 : 0;
      return {
        date: r.date, day: r.day, dayYA: r.dayYA,
        storeId: r.storeId, storeName: r.storeName, area: r.area,
        sales: r.sales, salesLY: r.salesLY,
        diffVal, diffPct: parseFloat(diffPct.toFixed(2)),
        justification: r.justification || '',
      };
    }).sort((a, b) => {
      if (a.date !== b.date) return a.date.localeCompare(b.date);
      return (a.storeName || '').localeCompare(b.storeName || '');
    });

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

app.get('/api/category-filters', async (req, res) => {
  try {
    const auth = getAuthClient();
    const sheets = google.sheets({ version: 'v4', auth });
    let data = await getCategoryData(sheets);
    data = scopeRowsByArea(data, req.user);
    const { area: filterArea } = req.query;
    const monthMap = new Map();
    data.forEach(r => {
      if (r.monthKey && !monthMap.has(r.monthKey)) monthMap.set(r.monthKey, r.month || categoryMonthLabel(r.monthRaw, r.monthKey));
    });
    const now = new Date();
    const monitorDate = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const monitorYear = monitorDate.getFullYear();
    for (let m = 0; m <= monitorDate.getMonth(); m++) {
      const key = monitorYear + '-' + String(m + 1).padStart(2, '0');
      if (!monthMap.has(key)) monthMap.set(key, monthLabel(key));
    }
    const months = [...monthMap.entries()]
      .map(([value, label]) => ({ value, label }))
      .sort((a, b) => a.value.localeCompare(b.value));
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
    res.json({ success: true, months, categories: [...categories].sort(), areas: [...areas].sort(), stores: [...stores].sort() });
  } catch (err) {
    console.error('Category-filters error:', err.message);
    res.status(500).json({ success: false, error: err.message });
  }
});

app.get('/api/category', async (req, res) => {
  try {
    const auth = getAuthClient();
    const sheets = google.sheets({ version: 'v4', auth });
    let [data, masterStores] = await Promise.all([getCategoryData(sheets), getMasterStoreList(sheets)]);
    data = scopeRowsByArea(data, req.user);
    masterStores = scopeRowsByArea(masterStores, req.user);
    const { month, category, area, store, sign } = req.query;
    const selectedMonthKey = month && month !== 'ALL' ? categoryMonthKey(month) : '';
    let filtered = data;
    if (selectedMonthKey && month !== 'ALL') filtered = filtered.filter(r => r.monthKey === selectedMonthKey);
    if (category && category !== 'ALL') filtered = filtered.filter(r => r.category === category);
    if (area && area !== 'ALL') filtered = filtered.filter(r => r.area === area);
    if (store && store !== 'ALL') filtered = filtered.filter(r => r.storeName === store);

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
        category: c.category, subDepCount: c.subDeps.size,
        sales: c.sales, salesLY: c.salesLY,
        diffVal, diffPct: parseFloat(diffPct.toFixed(2)), shareCur: parseFloat(shareCur.toFixed(2)),
      };
    }).sort((a, b) => b.sales - a.sales);

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

    const subDepsByCategory = {};
    filtered.forEach(r => {
      if (!r.category || !r.subDepName) return;
      if (!subDepsByCategory[r.category]) subDepsByCategory[r.category] = new Set();
      subDepsByCategory[r.category].add(r.subDepName);
    });
    const subDepsByCategoryOut = {};
    Object.entries(subDepsByCategory).forEach(([k,v]) => { subDepsByCategoryOut[k] = [...v].sort(); });

    const subMap = {};
    filtered.forEach(r => {
      const key = (r.category || '') + '||' + (r.storeCode || '') + '||' + (r.subDepName || '');
      if (!subMap[key]) subMap[key] = {
        category: r.category, storeCode: r.storeCode, storeName: r.storeName,
        area: r.area, subDepName: r.subDepName, sdepCode: r.sdepCode,
        sales: 0, salesLY: 0,
      };
      subMap[key].sales   += r.sales;
      subMap[key].salesLY += r.salesLY;
    });
    let detail = Object.values(subMap).map(s => {
      const diffVal = s.sales - s.salesLY;
      const diffPct = s.salesLY !== 0 ? (diffVal / s.salesLY) * 100 : 0;
      return { ...s, diffVal, diffPct: parseFloat(diffPct.toFixed(2)) };
    });

    const moversAgg = {};
    detail.forEach(r => {
      const key = r.subDepName || '';
      if (!moversAgg[key]) moversAgg[key] = { subDepName: r.subDepName, category: r.category, sales: 0, salesLY: 0 };
      moversAgg[key].sales   += r.sales;
      moversAgg[key].salesLY += r.salesLY;
    });
    const moversList = Object.values(moversAgg).map(m => {
      const diffVal = m.sales - m.salesLY;
      return { ...m, diffVal };
    });
    const moversSrc = [...moversList].sort((a, b) => b.diffVal - a.diffVal);
    const movers = {
      top: moversSrc.filter(r => r.diffVal > 0).slice(0, 10),
      bottom: moversSrc.filter(r => r.diffVal < 0).slice(-10).reverse(),
    };

    if (sign === 'POS') detail = detail.filter(r => r.diffVal > 0);
    if (sign === 'NEG') detail = detail.filter(r => r.diffVal < 0);
    detail.sort((a, b) => b.sales - a.sales);

    const monitorRows = buildCategoryMonitorRows(data, masterStores, { monthKey: selectedMonthKey, area, store });

    res.json({ success: true, summary, detail, byArea, byStore, subDepsByCategory: subDepsByCategoryOut, movers, monitorRows });
  } catch (err) {
    console.error('Category error:', err.message);
    res.status(500).json({ success: false, error: err.message });
  }
});

app.get('/api/data-gaps', async (req, res) => {
  try {
    const auth = getAuthClient();
    const sheets = google.sheets({ version: 'v4', auth });
    let [data, masterStores, dailySalesRows] = await Promise.all([getCategoryData(sheets), getMasterStoreList(sheets), getSalesData(sheets)]);
    data = scopeRowsByArea(data, req.user);
    masterStores = scopeRowsByArea(masterStores, req.user);
    dailySalesRows = scopeRowsByArea(dailySalesRows, req.user);

    const { month, area, store } = req.query;
    const selectedMonthKey = month && month !== 'ALL' ? categoryMonthKey(month) : '';
    const status = buildCategoryMonitorStatus(data, masterStores, { monthKey: selectedMonthKey, area, store });
    const lastUpdateRows = buildCategoryLastUpdateRows(data, { monthKey: selectedMonthKey, area, store }, dailySalesRows);

    res.json({
      success: true,
      rows: status.gapRows,
      gapRows: status.gapRows,
      completeRows: status.completeRows,
      lastUpdateRows,
      count: status.gapRows.length,
      completeCount: status.completeRows.length,
    });
  } catch (err) {
    console.error('Data-gaps error:', err.message);
    res.status(500).json({ success: false, error: err.message });
  }
});

app.get('/api/daily-sales-gaps', async (req, res) => {
  try {
    const auth = getAuthClient();
    const sheets = google.sheets({ version: 'v4', auth });
    let [data, masterStores] = await Promise.all([getSalesData(sheets), getMasterStoreList(sheets)]);
    data = scopeRowsByArea(data, req.user);
    masterStores = scopeRowsByArea(masterStores, req.user);

    const { month, area, store } = req.query;
    const result = buildDailySalesGapMonitor(data, masterStores, { month, area, store });
    res.json({ success: true, ...result });
  } catch (err) {
    console.error('Daily-sales-gaps error:', err.message);
    res.status(500).json({ success: false, error: err.message });
  }
});

app.get('/api/inventory-log-gaps', async (req, res) => {
  try {
    const auth = getAuthClient();
    const sheets = google.sheets({ version: 'v4', auth });
    let [data, masterStores] = await Promise.all([getInventoryLogsData(sheets), getMasterStoreList(sheets)]);
    const storeLookup = new Map();
    masterStores.forEach(s => {
      if (s.storeId) storeLookup.set('id:' + normalizeKey(s.storeId), s);
      if (s.storeName) {
        storeLookup.set('name:' + normalizeKey(s.storeName), s);
        storeLookup.set('cname:' + compactKey(s.storeName), s);
      }
    });
    data = data.map(r => {
      const master = storeLookup.get('id:' + normalizeKey(r.storeId)) || storeLookup.get('name:' + normalizeKey(r.storeName)) || storeLookup.get('cname:' + compactKey(r.storeName));
      return master ? { ...r, area: r.area || master.area, storeId: r.storeId || master.storeId, storeName: r.storeName || master.storeName } : r;
    });
    data = scopeRowsByArea(data, req.user);
    masterStores = scopeRowsByArea(masterStores, req.user);

    const { month, area, store } = req.query;
    const now = new Date();
    const result = buildStoreDayGapMonitor(data, masterStores, { month, area, store }, {
      startDate: now.getFullYear() + '-02-02',
      includeToday: true,
      label: 'Inventory Logs data',
    });
    res.json({ success: true, ...result });
  } catch (err) {
    console.error('Inventory-log-gaps error:', err.message);
    res.status(500).json({ success: false, error: err.message });
  }
});
app.get('/api/data-gap-filters', async (req, res) => {
  try {
    const auth = getAuthClient();
    const sheets = google.sheets({ version: 'v4', auth });
    let [data, masterStores] = await Promise.all([getCategoryData(sheets), getMasterStoreList(sheets)]);
    data = scopeRowsByArea(data, req.user);
    masterStores = scopeRowsByArea(masterStores, req.user);

    const { area: filterArea } = req.query;
    const monthMap = new Map();
    data.forEach(r => {
      if (r.monthKey && !monthMap.has(r.monthKey)) monthMap.set(r.monthKey, r.month || categoryMonthLabel(r.monthRaw, r.monthKey));
    });
    const now = new Date();
    const monitorDate = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const monitorYear = monitorDate.getFullYear();
    for (let m = 0; m <= monitorDate.getMonth(); m++) {
      const key = monitorYear + '-' + String(m + 1).padStart(2, '0');
      if (!monthMap.has(key)) monthMap.set(key, monthLabel(key));
    }
    const months = [...monthMap.entries()]
      .map(([value, label]) => ({ value, label }))
      .sort((a, b) => a.value.localeCompare(b.value));

    const areas = new Set();
    const stores = new Set();
    masterStores.forEach(s => {
      if (s.area) areas.add(s.area);
      if (s.storeName && (!filterArea || filterArea === 'ALL' || s.area === filterArea)) stores.add(s.storeName);
    });

    res.json({ success: true, months, areas: [...areas].sort(), stores: [...stores].sort() });
  } catch (err) {
    console.error('Data-gap-filters error:', err.message);
    res.status(500).json({ success: false, error: err.message });
  }
});

app.get('/api/category-breakdown', async (req, res) => {
  try {
    const auth = getAuthClient();
    const sheets = google.sheets({ version: 'v4', auth });
    let data = await getCategoryData(sheets);
    data = scopeRowsByArea(data, req.user);
    const { month, category, area, store, breakdownCategory, breakdownSubDep } = req.query;
    const selectedMonthKey = month && month !== 'ALL' ? categoryMonthKey(month) : '';
    let filtered = data;
    if (selectedMonthKey && month !== 'ALL') filtered = filtered.filter(r => r.monthKey === selectedMonthKey);
    if (category && category !== 'ALL') filtered = filtered.filter(r => r.category === category);
    if (area && area !== 'ALL') filtered = filtered.filter(r => r.area === area);
    if (store && store !== 'ALL') filtered = filtered.filter(r => r.storeName === store);
    if (breakdownCategory && breakdownCategory !== 'ALL') filtered = filtered.filter(r => r.category === breakdownCategory);
    if (breakdownSubDep && breakdownSubDep !== 'ALL') filtered = filtered.filter(r => r.subDepName === breakdownSubDep);

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

app.get('/api/averages', async (req, res) => {
  try {
    const auth = getAuthClient();
    const sheets = google.sheets({ version: 'v4', auth });
    let data = await getSalesData(sheets);
    data = scopeRowsByArea(data, req.user);
    const { area, store } = req.query;
    let filtered = data.filter(r => {
      if (!r.date || !r.sales || r.sales <= 0) return false;
      const mmdd = r.date.substring(5);
      if (mmdd === '01-01') return false;
      return true;
    });
    if (area && area !== 'ALL') filtered = filtered.filter(r => r.area === area);
    if (store && store !== 'ALL') filtered = filtered.filter(r => r.storeName === store);

    const storeStats = {};
    filtered.forEach(r => {
      const key = r.storeId + '_' + r.storeName;
      if (!storeStats[key]) storeStats[key] = { storeId: r.storeId, storeName: r.storeName, area: r.area, total: 0, days: new Set() };
      storeStats[key].total += r.sales;
      storeStats[key].days.add(r.date);
    });
    const perStore = Object.values(storeStats).map(s => ({
      storeId: s.storeId, storeName: s.storeName, area: s.area,
      avg: s.days.size ? s.total / s.days.size : 0,
      total: s.total, dayCount: s.days.size,
    })).sort((a, b) => b.avg - a.avg);

    const areaStats = {};
    filtered.forEach(r => {
      if (!areaStats[r.area]) areaStats[r.area] = { area: r.area, total: 0, count: 0, days: new Set() };
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

// ─── API: GET /api/store-notes ──────────────────────────────────────────────
app.get('/api/store-notes', async (req, res) => {
  try {
    const auth = getAuthClient();
    const sheets = google.sheets({ version: 'v4', auth });
    let data = await getStoreNotesData(sheets);
    data = scopeRowsByArea(data, req.user);
    const { area, store, status, q } = req.query;
    let filtered = data;
    if (area && area !== 'ALL') filtered = filtered.filter(r => r.area === area);
    if (store && store !== 'ALL') filtered = filtered.filter(r => r.storeName === store);
    const statusSet = parseMultiFilterParam(status);
    if (statusSet.size) filtered = filtered.filter(r => statusSet.has(normalizeFilterText(r.status) || '__blank__'));
    if (q) {
      const needle = q.toLowerCase();
      filtered = filtered.filter(r =>
        (r.notes || '').toLowerCase().includes(needle) ||
        (r.storeName || '').toLowerCase().includes(needle) ||
        (r.remarks || '').toLowerCase().includes(needle)
      );
    }
    // Sort newest first by parsed timestamp (ts is pre-parsed numeric)
    filtered.sort((a, b) => (b.ts || 0) - (a.ts || 0));
    res.json({ success: true, count: filtered.length, rows: filtered });
  } catch (err) {
    console.error('Store-notes error:', err.message);
    res.status(500).json({ success: false, error: err.message });
  }
});

// ─── API: GET /api/store-notes-filters ──────────────────────────────────────
app.get('/api/store-notes-filters', async (req, res) => {
  try {
    const auth = getAuthClient();
    const sheets = google.sheets({ version: 'v4', auth });
    let data = await getStoreNotesData(sheets);
    data = scopeRowsByArea(data, req.user);
    const { area: filterArea } = req.query;
    const areas = new Set();
    const stores = new Set();
    const statuses = new Map();
    data.forEach(r => {
      if (r.area) areas.add(r.area);
      const statusKey = normalizeFilterText(r.status);
      if (statusKey && statusKey !== 'status') addCanonicalFilterValue(statuses, r.status);
      if (r.storeName) {
        if (!filterArea || filterArea === 'ALL' || r.area === filterArea) stores.add(r.storeName);
      }
    });
    res.json({
      success: true,
      areas: [...areas].sort(),
      stores: [...stores].sort(),
      statuses: [...statuses.values()].sort((a, b) => normalizeFilterText(a).localeCompare(normalizeFilterText(b))),
    });
  } catch (err) {
    console.error('Store-notes-filters error:', err.message);
    res.status(500).json({ success: false, error: err.message });
  }
});

// ─── API: GET /api/issues-filters ───────────────────────────────────────────
app.get('/api/issues-filters', async (req, res) => {
  try {
    const auth = getAuthClient();
    const sheets = google.sheets({ version: 'v4', auth });
    let data = await getIssuesData(sheets);
    data = scopeRowsByArea(data, req.user);
    const { area: filterArea } = req.query;
    const areas = new Set();
    const stores = new Set();
    const priorities = new Map();
    const statuses = new Map();
    const categories = new Set();
    data.forEach(r => {
      if (r.area) areas.add(r.area);
      addCanonicalFilterValue(priorities, r.priority);
      addCanonicalFilterValue(statuses, r.status);
      if (r.issueCategory) categories.add(r.issueCategory);
      if (r.storeName) {
        if (!filterArea || filterArea === 'ALL' || r.area === filterArea) stores.add(r.storeName);
      }
    });
    res.json({
      success: true,
      areas: [...areas].sort(),
      stores: [...stores].sort(),
      priorities: [...priorities.values()].sort((a, b) => prioritySortRank(a) - prioritySortRank(b) || a.localeCompare(b)),
      statuses: [...statuses.values()].sort((a, b) => normalizeFilterText(a).localeCompare(normalizeFilterText(b))),
      categories: [...categories].sort(),
    });
  } catch (err) {
    console.error('Issues-filters error:', err.message);
    res.status(500).json({ success: false, error: err.message });
  }
});

// ─── API: GET /api/issues ───────────────────────────────────────────────────
app.get('/api/issues', async (req, res) => {
  try {
    const auth = getAuthClient();
    const sheets = google.sheets({ version: 'v4', auth });
    let data = await getIssuesData(sheets);
    data = scopeRowsByArea(data, req.user);
    const { area, store, priority, status, category, q } = req.query;
    let filtered = data;
    if (area && area !== 'ALL') filtered = filtered.filter(r => r.area === area);
    if (store && store !== 'ALL') filtered = filtered.filter(r => r.storeName === store);
    const prioritySet = parseMultiFilterParam(priority);
    const statusSet = parseMultiFilterParam(status);
    if (prioritySet.size) filtered = filtered.filter(r => prioritySet.has(normalizeFilterText(r.priority)));
    if (statusSet.size) filtered = filtered.filter(r => statusSet.has(normalizeFilterText(r.status)));
    if (category && category !== 'ALL') filtered = filtered.filter(r => r.issueCategory === category);
    if (q) {
      const needle = q.toLowerCase();
      filtered = filtered.filter(r =>
        (r.issueDescription || '').toLowerCase().includes(needle) ||
        (r.storeName || '').toLowerCase().includes(needle) ||
        (r.issueCategory || '').toLowerCase().includes(needle) ||
        (r.issueSubCategory || '').toLowerCase().includes(needle) ||
        (r.assignTo || '').toLowerCase().includes(needle) ||
        (r.remarks || '').toLowerCase().includes(needle)
      );
    }
    // Sort by Date descending (newest first)
    filtered.sort((a, b) => (b.dateTs || 0) - (a.dateTs || 0));
    res.json({ success: true, count: filtered.length, rows: filtered });
  } catch (err) {
    console.error('Issues error:', err.message);
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
<script>
try {
  if (localStorage.getItem('camanava-theme') === 'light') document.documentElement.classList.add('theme-light');
} catch (err) {}
</script>
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

html.theme-light{
  color-scheme:light;
  --bg-base:#f4faf9;
  --bg-deep:#e8f4f2;
  --bg-glass:rgba(255,255,255,0.82);
  --bg-glass2:rgba(241,249,248,0.9);
  --bg-card:rgba(255,255,255,0.92);
  --bg-elevated:rgba(255,255,255,0.98);

  --border-glow:rgba(10,92,112,0.18);
  --border-soft:rgba(15,93,117,0.16);
  --border-strong:rgba(15,93,117,0.28);

  --text-1:#102b3a;
  --text-2:#3e6171;
  --text-3:#6e8896;
  --text-dim:#94a7ae;

  --indigo:#0f766e;
  --indigo2:#0891b2;
  --indigo-glow:rgba(8,145,178,0.25);
  --cyan:#0891b2;
  --cyan2:#06b6d4;
  --cyan-glow:rgba(8,145,178,0.22);
  --emerald:#047857;
  --emerald2:#10b981;
  --emerald-glow:rgba(16,185,129,0.22);
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

html.theme-light body::before{
  background:
    radial-gradient(at 12% 8%, rgba(8,145,178,0.16) 0px, transparent 48%),
    radial-gradient(at 88% 12%, rgba(16,185,129,0.12) 0px, transparent 48%),
    radial-gradient(at 72% 88%, rgba(15,118,110,0.12) 0px, transparent 52%),
    linear-gradient(180deg,#f8fcfb 0%,#eef8f6 100%);
}
html.theme-light body::after{
  background-image:
    linear-gradient(rgba(15,93,117,0.055) 1px, transparent 1px),
    linear-gradient(90deg, rgba(15,93,117,0.055) 1px, transparent 1px);
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
html.theme-light .header{
  background:linear-gradient(180deg, rgba(255,255,255,0.92) 0%, rgba(245,252,250,0.82) 100%);
  box-shadow:0 10px 28px -24px rgba(15,93,117,0.35);
}
html.theme-light .logo-text{background:linear-gradient(135deg,#0b2534 0%,#0f766e 100%);-webkit-background-clip:text;background-clip:text;-webkit-text-fill-color:transparent}

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
.theme-toggle{min-width:112px;justify-content:center}
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
.main{padding:24px 28px 40px;max-width:1800px;margin:0 auto;width:100%}

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
  position:relative;z-index:20;
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
select.multi-select{min-width:150px;min-height:92px;padding-right:14px;background-image:none}
select.multi-select option{padding:5px 8px;border-radius:6px;margin:1px 0}
.tick-dropdown{position:relative;min-width:170px;z-index:200}
.tick-trigger{
  width:100%;min-width:170px;
  background:rgba(15,20,35,0.7);
  border:1px solid var(--border-strong);
  color:var(--text-1);
  padding:9px 34px 9px 14px;border-radius:10px;
  font-size:13px;font-family:'Inter',sans-serif;font-weight:500;
  cursor:pointer;outline:none;transition:all .2s;
  text-align:left;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;
  position:relative;
}
.tick-trigger::after{
  content:'';position:absolute;right:13px;top:50%;transform:translateY(-50%);
  border-left:4px solid transparent;border-right:4px solid transparent;border-top:5px solid var(--indigo2);
}
.tick-trigger:hover,.tick-dropdown.open .tick-trigger{border-color:var(--indigo);background:rgba(99,102,241,0.05)}
.tick-menu{
  display:none;position:absolute;left:0;top:calc(100% + 6px);z-index:9999;
  width:220px;max-height:340px;overflow:auto;
  background:rgba(15,20,35,0.98);
  border:1px solid var(--border-strong);
  border-radius:12px;padding:10px;
  box-shadow:0 18px 40px -18px rgba(0,0,0,0.65);
}
.tick-dropdown.open .tick-menu{display:block}
.controls:has(.tick-dropdown.open){margin-bottom:260px}
.controls.dropdown-open{margin-bottom:260px}
.tick-actions{display:flex;gap:7px;margin-bottom:9px}
.tick-action{
  border:0;border-radius:7px;padding:6px 10px;
  background:rgba(16,185,129,0.18);color:var(--emerald2);
  font-size:11.5px;font-weight:700;cursor:pointer;
}
.tick-action.clear{background:rgba(148,163,200,0.12);color:var(--text-2)}
.tick-option{
  display:flex;align-items:center;gap:9px;
  padding:7px 8px;border-radius:7px;
  color:var(--text-1);font-size:12.5px;font-weight:700;
  letter-spacing:.02em;cursor:pointer;text-transform:uppercase;
}
.tick-option:hover{background:rgba(99,102,241,0.08)}
.tick-option input{width:14px;height:14px;accent-color:var(--emerald);flex:0 0 auto}
.check-filter{
  min-width:150px;max-height:116px;overflow:auto;
  background:rgba(15,20,35,0.7);border:1px solid var(--border-strong);
  border-radius:10px;padding:7px 9px;display:grid;gap:5px;
}
.check-filter label{display:flex;align-items:center;gap:7px;font-size:12px;color:var(--text-1);font-weight:500;white-space:nowrap;cursor:pointer}
.check-filter input{accent-color:var(--indigo);width:13px;height:13px;flex:0 0 auto}
.check-filter:hover{border-color:var(--indigo);background:rgba(99,102,241,0.05)}
select:hover,input[type=date]:hover{border-color:var(--indigo);background:rgba(99,102,241,0.05)}
select:focus,input[type=date]:focus{border-color:var(--indigo);box-shadow:0 0 0 3px rgba(99,102,241,0.15)}
input[type=date]::-webkit-calendar-picker-indicator{filter:invert(0.6) sepia(1) saturate(5) hue-rotate(210deg);cursor:pointer;opacity:0.8}
html.theme-light select,
html.theme-light input[type=date],
html.theme-light .tick-trigger,
html.theme-light .check-filter{
  background:rgba(255,255,255,0.94);
  color:var(--text-1);
}
html.theme-light select:hover,
html.theme-light input[type=date]:hover,
html.theme-light .tick-trigger:hover,
html.theme-light .tick-dropdown.open .tick-trigger,
html.theme-light .check-filter:hover{
  background:rgba(235,249,247,0.98);
}
html.theme-light .tick-menu{
  background:rgba(255,255,255,0.98);
  box-shadow:0 18px 40px -24px rgba(15,93,117,0.45);
}
html.theme-light input[type=date]::-webkit-calendar-picker-indicator{filter:none;opacity:.72}
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
table{width:100%;border-collapse:collapse;font-size:14px;table-layout:auto}
th{
  background:rgba(15,20,35,0.85);
  backdrop-filter:blur(10px);
  padding:14px 18px;text-align:left;
  font-size:11px;font-weight:700;color:var(--text-3);
  text-transform:uppercase;letter-spacing:0.1em;
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
  display:inline-block;margin-left:5px;font-size:10px;
  opacity:0.35;transition:opacity .15s;color:var(--text-3);
}
th.sortable:hover .sort-icon{opacity:0.6}
th.sortable.sort-asc .sort-icon,
th.sortable.sort-desc .sort-icon{opacity:1;color:var(--indigo2)}
th.sortable.sort-asc,
th.sortable.sort-desc{color:var(--indigo2)}
td{
  padding:14px 18px;border-bottom:1px solid var(--border-soft);
  vertical-align:middle;white-space:nowrap;
  font-size:14px;
}
td:has(.just-full),td:has(.notes-cell),td:has(.remarks-cell){white-space:normal}
tr:last-child td{border-bottom:none}
tbody tr{transition:background .15s}
tbody tr:hover td{background:rgba(99,102,241,0.04)}

.store-cell{display:flex;align-items:center;gap:12px}
.store-avatar{
  width:40px;height:40px;border-radius:11px;
  display:flex;align-items:center;justify-content:center;
  font-family:'Space Grotesk',sans-serif;
  font-size:13px;font-weight:600;color:#fff;
  flex-shrink:0;letter-spacing:-0.5px;
  box-shadow:inset 0 1px 0 rgba(255,255,255,0.15);
}
.store-info{display:flex;flex-direction:column;gap:2px}
.store-name{font-weight:600;color:var(--text-1);font-size:14.5px;letter-spacing:-0.01em}
.store-id{font-size:11.5px;color:var(--text-3);font-family:'JetBrains Mono',monospace;font-weight:500;margin-top:2px}

.area-tag{
  display:inline-flex;align-items:center;gap:8px;
  font-size:13px;color:var(--text-2);white-space:nowrap;
  padding:6px 12px;border-radius:8px;
  background:rgba(99,148,255,0.06);border:1px solid rgba(99,148,255,0.12);
  font-weight:500;
}
.area-dot{width:8px;height:8px;border-radius:50%;flex-shrink:0;box-shadow:0 0 8px currentColor}

.num{font-family:'JetBrains Mono',monospace;font-size:13.5px;font-weight:500;letter-spacing:-0.01em}
.num-bold{font-weight:600}

.pill{
  display:inline-flex;align-items:center;gap:4px;
  padding:5px 12px;border-radius:8px;
  font-size:12.5px;font-weight:600;
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
  white-space:normal;color:var(--text-2);font-size:13px;
  line-height:1.5;font-style:italic;max-width:none;font-weight:400;
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

/* ── Photo Lightbox ── */
.lightbox{
  position:fixed;inset:0;
  background:rgba(5,7,14,0.92);
  backdrop-filter:blur(14px);
  -webkit-backdrop-filter:blur(14px);
  display:none;align-items:center;justify-content:center;
  z-index:9999;
  padding:30px;
  animation:lbFadeIn .2s ease-out;
}
.lightbox.active{display:flex}
@keyframes lbFadeIn{from{opacity:0}to{opacity:1}}
.lightbox-content{
  position:relative;
  max-width:min(1200px, 95vw);
  max-height:90vh;
  display:flex;flex-direction:column;
  background:rgba(20,26,42,0.6);
  border:1px solid var(--border-strong);
  border-radius:18px;overflow:hidden;
  box-shadow:0 24px 80px -16px rgba(0,0,0,0.8);
  animation:lbZoomIn .25s cubic-bezier(0.34, 1.56, 0.64, 1);
}
@keyframes lbZoomIn{from{transform:scale(0.92);opacity:0}to{transform:scale(1);opacity:1}}
.lightbox-header{
  display:flex;align-items:center;justify-content:space-between;
  padding:14px 20px;
  border-bottom:1px solid var(--border-soft);
  background:rgba(15,20,35,0.8);
  gap:14px;flex-shrink:0;
}
.lightbox-title{
  font-family:'Space Grotesk',sans-serif;
  font-size:14px;font-weight:600;color:var(--text-1);
  display:flex;align-items:center;gap:10px;letter-spacing:-0.01em;
  min-width:0;
}
.lightbox-title i{color:var(--indigo2);font-size:14px;flex-shrink:0}
.lightbox-title-text{
  white-space:nowrap;overflow:hidden;text-overflow:ellipsis;
}
.lightbox-actions{display:flex;align-items:center;gap:8px;flex-shrink:0}
.lightbox-btn{
  background:var(--bg-glass);
  border:1px solid var(--border-strong);
  color:var(--text-2);
  width:36px;height:36px;border-radius:9px;
  display:flex;align-items:center;justify-content:center;
  cursor:pointer;font-size:14px;
  transition:all .2s;text-decoration:none;
}
.lightbox-btn:hover{
  background:var(--bg-elevated);color:var(--text-1);
  border-color:var(--indigo);
}
.lightbox-btn.close-btn:hover{border-color:var(--rose);color:var(--rose2)}
.lightbox-body{
  position:relative;
  flex:1;min-height:0;
  display:flex;align-items:center;justify-content:center;
  background:#0a0e1a;
  overflow:hidden;
}
.lightbox-img{
  max-width:100%;max-height:100%;
  object-fit:contain;
  display:block;
  user-select:none;
  -webkit-user-drag:none;
}
.lightbox-spinner{
  position:absolute;
  width:48px;height:48px;
  border:3px solid transparent;
  border-top-color:var(--indigo);
  border-right-color:var(--cyan);
  border-radius:50%;
  animation:spin 1s linear infinite;
  filter:drop-shadow(0 0 12px var(--indigo-glow));
}
.lightbox-error{
  padding:40px 30px;text-align:center;color:var(--text-2);
  display:none;flex-direction:column;align-items:center;gap:12px;
}
.lightbox-error i{font-size:36px;color:var(--amber);opacity:0.8}
.lightbox-error.show{display:flex}
.lightbox-error a{color:var(--indigo2);text-decoration:underline;font-size:12.5px}
.lightbox-meta{
  padding:12px 20px;
  background:rgba(15,20,35,0.8);
  border-top:1px solid var(--border-soft);
  font-size:12px;color:var(--text-3);
  display:flex;align-items:center;gap:14px;flex-wrap:wrap;flex-shrink:0;
}
.lightbox-meta-item{display:inline-flex;align-items:center;gap:5px}
.lightbox-meta-item i{color:var(--text-dim);font-size:11px}
@media(max-width:768px){
  .lightbox{padding:12px}
  .lightbox-content{max-width:100vw;max-height:95vh}
  .lightbox-header{padding:10px 14px}
  .lightbox-title{font-size:13px}
  .lightbox-btn{width:34px;height:34px}
  .lightbox-meta{font-size:11px;padding:10px 14px;gap:10px}
}

/* ── Store Notes specific ── */
.status-pill{
  display:inline-flex;align-items:center;gap:6px;
  padding:6px 12px;border-radius:8px;
  font-size:12px;font-weight:600;
  font-family:'Inter',sans-serif;letter-spacing:-0.01em;
  white-space:nowrap;
}
.status-pill.done,.status-pill.completed,.status-pill.closed,.status-pill.resolved{
  background:rgba(16,185,129,0.12);border:1px solid rgba(16,185,129,0.3);color:var(--emerald2);
}
.status-pill.pending,.status-pill.open,.status-pill.inprogress,.status-pill.ongoing{
  background:rgba(245,158,11,0.12);border:1px solid rgba(245,158,11,0.3);color:var(--amber2);
}
.status-pill.urgent,.status-pill.cancelled,.status-pill.failed,.status-pill.rejected{
  background:rgba(244,63,94,0.12);border:1px solid rgba(244,63,94,0.3);color:var(--rose2);
}
.status-pill.review,.status-pill.draft,.status-pill.new{
  background:rgba(99,102,241,0.12);border:1px solid rgba(99,102,241,0.3);color:var(--indigo2);
}
.status-pill.default{
  background:rgba(148,163,200,0.08);border:1px solid var(--border-strong);color:var(--text-2);
}
.photo-link-btn{
  display:inline-flex;align-items:center;gap:6px;
  padding:8px 14px;border-radius:9px;
  background:linear-gradient(135deg, var(--indigo), #4f46e5);
  color:#fff;font-size:12.5px;font-weight:600;
  text-decoration:none;
  border:none;cursor:pointer;
  transition:all .2s;letter-spacing:-0.01em;
  box-shadow:0 2px 8px -2px var(--indigo-glow);
  font-family:'Inter',sans-serif;
}
.photo-link-btn:hover{
  transform:translateY(-1px);
  box-shadow:0 4px 12px -2px var(--indigo-glow);
}
.photo-link-btn.disabled{
  background:rgba(148,163,200,0.08);
  color:var(--text-3);
  cursor:default;
  box-shadow:none;
  border:1px solid var(--border-strong);
}
.photo-link-btn.disabled:hover{transform:none}
.notes-cell{
  max-width:none;white-space:normal;word-wrap:break-word;
  color:var(--text-2);font-size:13.5px;line-height:1.5;font-weight:400;
}
.remarks-cell{
  max-width:none;white-space:normal;word-wrap:break-word;
  color:var(--text-3);font-size:12.5px;line-height:1.45;font-style:italic;
}
.timestamp-cell{
  white-space:nowrap;font-family:'JetBrains Mono',monospace;
  font-size:12.5px;color:var(--text-2);font-weight:500;line-height:1.5;
}

html.theme-light .controls,
html.theme-light .kpi,
html.theme-light .chart-card,
html.theme-light .table-card,
html.theme-light .tabs,
html.theme-light .missing-card{
  box-shadow:0 14px 32px -26px rgba(15,93,117,0.38);
}
html.theme-light th{
  background:rgba(238,248,246,0.96);
  color:#5c7480;
}
html.theme-light tbody tr:hover td{background:rgba(8,145,178,0.055)}
html.theme-light .tab-btn.active{
  background:linear-gradient(135deg, rgba(15,118,110,0.15), rgba(8,145,178,0.16));
  color:#0b3a42;
  border-color:rgba(8,145,178,0.35);
}
html.theme-light select option{background:#ffffff;color:#102b3a}
html.theme-light .sign-toggle{background:rgba(255,255,255,0.92)}
html.theme-light #nSearch,
html.theme-light #iSearch{
  background:rgba(255,255,255,0.94) !important;
  color:var(--text-1) !important;
}
html.theme-light .loading-overlay{background:rgba(244,250,249,0.88)}

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
  .main{padding:14px 10px 30px;max-width:100%}
  .header{padding:12px 16px;gap:10px;flex-wrap:wrap}
  .header-right{width:100%;overflow-x:auto;-webkit-overflow-scrolling:touch;padding-bottom:2px}
  .header-right::-webkit-scrollbar{display:none}
  #userBadge{max-width:220px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;flex:0 0 auto}
  .sync-btn{flex:0 0 auto}
  .controls{padding:12px 14px;gap:10px}
  .kpi-grid{grid-template-columns:repeat(2,1fr) !important;gap:10px}
  .kpi{padding:14px 14px}
  .kpi-value{font-size:20px}
  .charts-grid{grid-template-columns:1fr !important}
  .tabs{
    width:100%;max-width:100%;
    overflow-x:auto;overflow-y:hidden;
    flex-wrap:nowrap;
    -webkit-overflow-scrolling:touch;
    scrollbar-width:none;
  }
  .tabs::-webkit-scrollbar{display:none}
  .tab-btn{
    flex:0 0 auto;
    justify-content:center;
    padding:9px 12px;
    font-size:12px;
    white-space:nowrap;
    transition:none;
  }

  /* Table: allow horizontal scroll, compact padding */
  .table-wrap{overflow-x:auto;-webkit-overflow-scrolling:touch}
  table{font-size:12.5px}
  td,th{padding:10px 12px;font-size:12px}
  .store-name{font-size:13px}
  .store-id{font-size:10.5px}
  .num{font-size:12px}
  .area-tag{font-size:11.5px;padding:4px 9px}
  .pill{font-size:11px;padding:4px 9px}
  .status-pill{font-size:11px;padding:4px 9px}
  .photo-link-btn{font-size:11.5px;padding:6px 11px}
  .notes-cell{font-size:12.5px}
  .remarks-cell{font-size:11.5px}
  .timestamp-cell{font-size:11.5px}
  .just-full{font-size:12px;max-width:220px}

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
  .chart-card::before,.table-card::before,.kpi::before{opacity:.25}
  .sync-btn:hover,.tab-btn:hover,.photo-link-btn:hover,.export-btn:hover{transform:none !important;box-shadow:none !important}
}
@media (max-width: 768px) {
  html.theme-light .header{background:rgba(255,255,255,0.98) !important}
  html.theme-light .controls{background:rgba(255,255,255,0.98) !important}
  html.theme-light .kpi{background:rgba(255,255,255,0.98) !important}
  html.theme-light .chart-card{background:rgba(255,255,255,0.98) !important}
  html.theme-light .table-card{background:rgba(255,255,255,0.98) !important}
  html.theme-light .missing-card{background:rgba(255,255,255,0.98) !important}
  html.theme-light .tabs{background:rgba(255,255,255,0.98) !important}
  .theme-toggle{min-width:44px;padding:9px 11px}
  .theme-toggle span{display:none}
}
::-webkit-scrollbar{width:8px;height:8px}
::-webkit-scrollbar-track{background:rgba(15,20,35,0.4)}
::-webkit-scrollbar-thumb{background:linear-gradient(180deg, var(--indigo), #4f46e5);border-radius:4px}
::-webkit-scrollbar-thumb:hover{background:linear-gradient(180deg, var(--indigo2), var(--indigo))}
html.theme-light ::-webkit-scrollbar-track{background:rgba(225,241,238,0.7)}
html.theme-light ::-webkit-scrollbar-thumb{background:linear-gradient(180deg,#0891b2,#0f766e)}

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

<!-- Photo Lightbox -->
<div id="lightbox" class="lightbox" onclick="closeLightbox(event)" role="dialog" aria-modal="true" aria-labelledby="lbTitle">
  <div class="lightbox-content" onclick="event.stopPropagation()">
    <div class="lightbox-header">
      <div class="lightbox-title">
        <i class="fa fa-image"></i>
        <span class="lightbox-title-text" id="lbTitle">Photo</span>
      </div>
      <div class="lightbox-actions">
        <a id="lbOpenExternal" href="#" target="_blank" rel="noopener noreferrer" class="lightbox-btn" title="Open in new tab">
          <i class="fa fa-arrow-up-right-from-square"></i>
        </a>
        <button class="lightbox-btn close-btn" onclick="closeLightbox()" title="Close (Esc)" aria-label="Close">
          <i class="fa fa-xmark"></i>
        </button>
      </div>
    </div>
    <div class="lightbox-body">
      <div class="lightbox-spinner" id="lbSpinner"></div>
      <img id="lbImage" class="lightbox-img" alt="" style="display:none"/>
      <div class="lightbox-error" id="lbError">
        <i class="fa fa-triangle-exclamation"></i>
        <div>Couldn't load the preview.</div>
        <a id="lbErrorLink" href="#" target="_blank" rel="noopener noreferrer">Open in Google Drive instead →</a>
      </div>
    </div>
    <div class="lightbox-meta" id="lbMeta"></div>
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
    <div id="userBadge" class="badge loading-badge">Signed in</div>
    <form method="post" action="/logout" style="margin:0">
      <button class="sync-btn" type="submit">
        <i class="fa fa-right-from-bracket"></i> Logout
      </button>
    </form>
    <button class="sync-btn theme-toggle" id="themeToggle" type="button" onclick="toggleTheme()" aria-label="Switch to daylight mode">
      <i class="fa fa-sun" id="themeIcon"></i> <span id="themeText">Daylight</span>
    </button>
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
    <button class="tab-btn" data-tab="gaps" onclick="switchTab('gaps')">
      <i class="fa fa-clipboard-check"></i> Data Gaps Monitoring
    </button>
    <button class="tab-btn" data-tab="notes" onclick="switchTab('notes')">
      <i class="fa fa-note-sticky"></i> Store Notes
    </button>
    <button class="tab-btn" data-tab="issues" onclick="switchTab('issues')">
      <i class="fa fa-triangle-exclamation"></i> Store Issues &amp; Concerns
    </button>
  </div>

  <!--  DAILY TAB  -->
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
            <th class="sortable" data-sort-key="sales" data-sort-type="num" style="text-align:right" onclick="sortDaily('sales','num')">Sales <span class="sort-icon">⇅</span></th>
            <th class="sortable" data-sort-key="salesLY" data-sort-type="num" style="text-align:right" onclick="sortDaily('salesLY','num')">Sales LY <span class="sort-icon">⇅</span></th>
            <th class="sortable" data-sort-key="diffPct" data-sort-type="num" style="text-align:center" onclick="sortDaily('diffPct','num')">Diff % <span class="sort-icon">⇅</span></th>
            <th class="sortable" data-sort-key="diffVal" data-sort-type="num" style="text-align:right" onclick="sortDaily('diffVal','num')">Diff Amount <span class="sort-icon">⇅</span></th>
            <th class="sortable" data-sort-key="justification" data-sort-type="string" onclick="sortDaily('justification','string')">Justification <span class="sort-icon">⇅</span></th>
          </tr>
        </thead>
        <tbody id="tableBody">
          <tr><td colspan="6" class="empty-cell"><div class="empty-icon"><i class="fa fa-spinner fa-spin"></i></div><p>Loading data...</p></td></tr>
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

  <!--  MONTHLY TAB  -->
  <div class="tab-content" id="tab-monthly">

    <!-- Controls -->
    <div class="controls">
      <div class="ctrl-group">
        <span class="ctrl-label"><i class="fa fa-calendar"></i> Month</span>
        <div class="tick-dropdown" id="mMonthDropdown">
          <button class="tick-trigger" id="mMonthTrigger" type="button" onclick="toggleMMonthDropdown()">All Months</button>
          <div class="tick-menu" id="mMonthMenu"></div>
        </div>
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
              <th class="sortable" data-sort-key="sales" data-sort-type="num" style="text-align:right" onclick="sortMSummary('sales','num')">Sales <span class="sort-icon">⇅</span></th>
              <th class="sortable" data-sort-key="salesLY" data-sort-type="num" style="text-align:right" onclick="sortMSummary('salesLY','num')">Sales LY <span class="sort-icon">⇅</span></th>
              <th class="sortable" data-sort-key="diffPct" data-sort-type="num" style="text-align:center" onclick="sortMSummary('diffPct','num')">Diff % <span class="sort-icon">⇅</span></th>
              <th class="sortable" data-sort-key="diffVal" data-sort-type="num" style="text-align:right" onclick="sortMSummary('diffVal','num')">Diff Amount <span class="sort-icon">⇅</span></th>
            </tr>
          </thead>
          <tbody id="mSummaryBody">
            <tr><td colspan="6" class="empty-cell"><div class="empty-icon"><i class="fa fa-spinner fa-spin"></i></div><p>Loading...</p></td></tr>
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

  <!--  CATEGORY SALES TAB  -->
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
              <th class="sortable" data-sort-key="storeName" data-sort-type="string" onclick="sortCDetail('storeName','string')">Store Name <span class="sort-icon">⇅</span></th>
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

  <!--  DATA GAPS MONITORING TAB  -->
  <div class="tab-content" id="tab-gaps">

    <!-- Controls -->
    <div class="controls">
      <div class="ctrl-group">
        <span class="ctrl-label"><i class="fa fa-calendar"></i> Month</span>
        <select id="gMonthFilter" onchange="applyGapsFilters()">
          <option value="ALL">All Months</option>
        </select>
      </div>
      <div class="divider"></div>
      <div class="ctrl-group">
        <span class="ctrl-label"><i class="fa fa-layer-group"></i> Area</span>
        <select id="gAreaFilter" onchange="onGAreaChange()">
          <option value="ALL">All Areas</option>
        </select>
      </div>
      <div class="divider"></div>
      <div class="ctrl-group">
        <span class="ctrl-label"><i class="fa fa-store"></i> Store</span>
        <select id="gStoreFilter" onchange="applyGapsFilters()">
          <option value="ALL">All Stores</option>
        </select>
      </div>
      <div class="records-count" id="gRecordsCount">-</div>
    </div>

    <div class="charts-grid">
      <div class="table-card">
        <div class="table-header" style="gap:14px;flex-wrap:wrap">
          <div class="table-title"><i class="fa fa-triangle-exclamation"></i> Category Sales Gap</div>
          <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin-left:auto">
            <button class="export-btn" onclick="exportGapsToExcel('gap')">
              <i class="fa fa-file-excel"></i> Export Excel
            </button>
            <div class="table-date" id="gMonitorInfo">-</div>
          </div>
        </div>
        <div class="table-wrap" style="max-height:520px">
          <table>
            <thead>
              <tr>
                <th class="sortable" data-sort-key="monthKey" data-sort-type="string" onclick="sortGaps('monthKey','string')">Month <span class="sort-icon">&harr;</span></th>
                <th class="sortable" data-sort-key="area" data-sort-type="string" onclick="sortGaps('area','string')">Area <span class="sort-icon">&harr;</span></th>
                <th class="sortable" data-sort-key="storeId" data-sort-type="num" style="text-align:center" onclick="sortGaps('storeId','num')">Store ID <span class="sort-icon">&harr;</span></th>
                <th class="sortable" data-sort-key="storeName" data-sort-type="string" onclick="sortGaps('storeName','string')">Store Name <span class="sort-icon">&harr;</span></th>
                <th>Remarks</th>
              </tr>
            </thead>
            <tbody id="gMonitorBody">
              <tr><td colspan="5" class="empty-cell"><div class="empty-icon"><i class="fa fa-spinner fa-spin"></i></div><p>Loading...</p></td></tr>
            </tbody>
          </table>
        </div>
      </div>
<div class="table-card">
        <div class="table-header" style="gap:14px;flex-wrap:wrap">
          <div class="table-title"><i class="fa fa-circle-check"></i> Category Sales Complete</div>
          <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin-left:auto">
            <button class="export-btn" onclick="exportGapsToExcel('complete')">
              <i class="fa fa-file-excel"></i> Export Excel
            </button>
            <div class="table-date" id="gCompleteInfo">-</div>
          </div>
        </div>
        <div class="table-wrap" style="max-height:520px">
          <table>
            <thead>
              <tr>
                <th class="sortable" data-sort-key="area" data-sort-type="string" onclick="sortGapsComplete('area','string')">Area <span class="sort-icon">&harr;</span></th>
                <th class="sortable" data-sort-key="storeId" data-sort-type="num" style="text-align:center" onclick="sortGapsComplete('storeId','num')">Store ID <span class="sort-icon">&harr;</span></th>
                <th class="sortable" data-sort-key="storeName" data-sort-type="string" onclick="sortGapsComplete('storeName','string')">Store Name <span class="sort-icon">&harr;</span></th>
                <th>Remarks</th>
              </tr>
            </thead>
            <tbody id="gCompleteBody">
              <tr><td colspan="4" class="empty-cell"><div class="empty-icon"><i class="fa fa-spinner fa-spin"></i></div><p>Loading...</p></td></tr>
            </tbody>
          </table>
        </div>
      </div>
    </div>

    <div class="table-card" style="margin-top:24px;grid-column:1 / -1">
        <div class="table-header" style="gap:14px;flex-wrap:wrap">
          <div class="table-title"><i class="fa fa-clock-rotate-left"></i> Category Sales Last Update</div>
          <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin-left:auto">
            <button class="export-btn" onclick="exportCategoryLastUpdateToExcel()">
              <i class="fa fa-file-excel"></i> Export Excel
            </button>
            <div class="table-date" id="gLastUpdateInfo">-</div>
          </div>
        </div>
        <div class="table-wrap" style="max-height:520px">
          <table>
            <thead>
              <tr>
                <th>Store Name</th>
                <th style="text-align:right">Sales</th>
                <th style="text-align:right">SalesYA</th>
                <th style="text-align:right">Sales Diff %</th>
                <th style="text-align:right">Diff Value</th>
                <th style="text-align:right">Diff % (DS)</th>
                <th>Last Update</th>
              </tr>
            </thead>
            <tbody id="gLastUpdateBody">
              <tr><td colspan="7" class="empty-cell"><div class="empty-icon"><i class="fa fa-spinner fa-spin"></i></div><p>Loading...</p></td></tr>
            </tbody>
          </table>
        </div>
      </div>

    <div class="kpi-grid" style="margin-top:24px">
      <div class="kpi k-sales"><div class="kpi-label"><div class="kpi-icon"><i class="fa fa-triangle-exclamation"></i></div>Category Gap Rows</div><div class="kpi-value gradient-text" id="cgGapRows">-</div><div class="kpi-sub">store-month gaps</div></div>
      <div class="kpi k-ly"><div class="kpi-label"><div class="kpi-icon"><i class="fa fa-store"></i></div>Affected Stores</div><div class="kpi-value gradient-text" id="cgAffectedStores">-</div><div class="kpi-sub">stores with any gap</div></div>
      <div class="kpi k-diff"><div class="kpi-label"><div class="kpi-icon"><i class="fa fa-circle-check"></i></div>Complete Stores</div><div class="kpi-value" id="cgCompleteStores">-</div><div class="kpi-sub">complete for selected period</div></div>
      <div class="kpi k-pct"><div class="kpi-label"><div class="kpi-icon"><i class="fa fa-percent"></i></div>Category Completion</div><div class="kpi-value" id="cgCompletion">-</div><div class="kpi-sub">complete store-months / expected</div></div>
      <div class="kpi k-stores"><div class="kpi-label"><div class="kpi-icon"><i class="fa fa-calendar"></i></div>Gap Months</div><div class="kpi-value gradient-text" id="cgGapMonths">-</div><div class="kpi-sub">months with category gaps</div></div>
    </div>

    <div class="table-card" style="margin-top:24px">
      <div class="table-header">
        <div class="table-title"><i class="fa fa-chart-column"></i> Daily Sales Gap Metrics</div>
        <div class="table-date">Separate from Category Sales gaps</div>
      </div>
    </div>

    <div class="kpi-grid" style="margin-top:14px">
      <div class="kpi k-sales"><div class="kpi-label"><div class="kpi-icon"><i class="fa fa-calendar-check"></i></div>Expected Store-Days</div><div class="kpi-value gradient-text" id="dgExpected">-</div><div class="kpi-sub">January to yesterday</div></div>
      <div class="kpi k-ly"><div class="kpi-label"><div class="kpi-icon"><i class="fa fa-circle-check"></i></div>Reported Store-Days</div><div class="kpi-value gradient-text" id="dgReported">-</div><div class="kpi-sub">with Daily Sales data</div></div>
      <div class="kpi k-diff"><div class="kpi-label"><div class="kpi-icon"><i class="fa fa-triangle-exclamation"></i></div>Gap Days</div><div class="kpi-value" id="dgGapDays">-</div><div class="kpi-sub">missing store-days</div></div>
      <div class="kpi k-pct"><div class="kpi-label"><div class="kpi-icon"><i class="fa fa-percent"></i></div>Completion</div><div class="kpi-value" id="dgCompletion">-</div><div class="kpi-sub">reported / expected</div></div>
      <div class="kpi k-stores"><div class="kpi-label"><div class="kpi-icon"><i class="fa fa-store"></i></div>Store-Months</div><div class="kpi-value gradient-text" id="dgStoreMonths">-</div><div class="kpi-sub" id="dgStoreMonthsSub">gap vs complete</div></div>
    </div>

    <div class="chart-card" style="margin-top:24px">
      <div class="chart-title">
        <i class="fa fa-chart-column"></i> Daily Sales Gap
        <span style="margin-left:auto;font-size:10.5px;color:var(--text-3);font-weight:500;text-transform:uppercase;letter-spacing:0.08em" id="dgChartInfo">January to yesterday</span>
      </div>
      <div style="position:relative;height:340px">
        <canvas id="dailyGapChart"></canvas>
      </div>
    </div>

    <div class="charts-grid" style="margin-top:24px">
      <div class="table-card">
        <div class="table-header" style="gap:14px;flex-wrap:wrap">
          <div class="table-title"><i class="fa fa-triangle-exclamation"></i> Daily Sales Stores With Gap</div>
          <button class="export-btn" onclick="exportStoreDayGapTable('daily', 'gap')"><i class="fa fa-file-excel"></i> Export Excel</button>
          <div class="table-date" id="dgGapInfo">-</div>
        </div>
        <div class="table-wrap" style="max-height:520px">
          <table>
            <thead>
              <tr>
                <th>Area</th><th style="text-align:center">Store ID</th><th>Store Name</th><th style="text-align:center">Expected</th><th style="text-align:center">Reported</th><th style="text-align:center">Gap Days</th><th style="text-align:center">Completion</th>
              </tr>
            </thead>
            <tbody id="dgGapBody"><tr><td colspan="7" class="empty-cell"><div class="empty-icon"><i class="fa fa-spinner fa-spin"></i></div><p>Loading...</p></td></tr></tbody>
          </table>
        </div>
      </div>
      <div class="table-card">
        <div class="table-header" style="gap:14px;flex-wrap:wrap">
          <div class="table-title"><i class="fa fa-circle-check"></i> Daily Sales Complete Stores</div>
          <button class="export-btn" onclick="exportStoreDayGapTable('daily', 'complete')"><i class="fa fa-file-excel"></i> Export Excel</button>
          <div class="table-date" id="dgCompleteInfo">-</div>
        </div>
        <div class="table-wrap" style="max-height:520px">
          <table>
            <thead>
              <tr>
                <th>Area</th><th style="text-align:center">Store ID</th><th>Store Name</th><th style="text-align:center">Expected</th><th style="text-align:center">Reported</th><th style="text-align:center">Gap Days</th><th style="text-align:center">Completion</th>
              </tr>
            </thead>
            <tbody id="dgCompleteBody"><tr><td colspan="7" class="empty-cell"><div class="empty-icon"><i class="fa fa-spinner fa-spin"></i></div><p>Loading...</p></td></tr></tbody>
          </table>
        </div>
      </div>
    </div>

    <div class="kpi-grid" style="margin-top:24px">
      <div class="kpi k-sales"><div class="kpi-label"><div class="kpi-icon"><i class="fa fa-clipboard-list"></i></div>Inventory Expected</div><div class="kpi-value gradient-text" id="igExpected">-</div><div class="kpi-sub">February 2 to today</div></div>
      <div class="kpi k-ly"><div class="kpi-label"><div class="kpi-icon"><i class="fa fa-circle-check"></i></div>Inventory Reported</div><div class="kpi-value gradient-text" id="igReported">-</div><div class="kpi-sub">with InventoryLogs data</div></div>
      <div class="kpi k-diff"><div class="kpi-label"><div class="kpi-icon"><i class="fa fa-triangle-exclamation"></i></div>Inventory Gap Days</div><div class="kpi-value" id="igGapDays">-</div><div class="kpi-sub">missing store-days</div></div>
      <div class="kpi k-pct"><div class="kpi-label"><div class="kpi-icon"><i class="fa fa-percent"></i></div>Inventory Completion</div><div class="kpi-value" id="igCompletion">-</div><div class="kpi-sub">reported / expected</div></div>
      <div class="kpi k-stores"><div class="kpi-label"><div class="kpi-icon"><i class="fa fa-store"></i></div>Inventory Store-Months</div><div class="kpi-value gradient-text" id="igStoreMonths">-</div><div class="kpi-sub" id="igStoreMonthsSub">gap vs complete</div></div>
    </div>

    <div class="chart-card" style="margin-top:24px">
      <div class="chart-title">
        <i class="fa fa-chart-column"></i> Inventory Logs Gap
        <span style="margin-left:auto;font-size:10.5px;color:var(--text-3);font-weight:500;text-transform:uppercase;letter-spacing:0.08em" id="igChartInfo">February 2 to today</span>
      </div>
      <div style="position:relative;height:340px">
        <canvas id="inventoryGapChart"></canvas>
      </div>
    </div>

    <div class="charts-grid" style="margin-top:24px">
      <div class="table-card">
        <div class="table-header" style="gap:14px;flex-wrap:wrap">
          <div class="table-title"><i class="fa fa-triangle-exclamation"></i> Inventory Logs Stores With Gap</div>
          <button class="export-btn" onclick="exportStoreDayGapTable('inventory', 'gap')"><i class="fa fa-file-excel"></i> Export Excel</button>
          <div class="table-date" id="igGapInfo">-</div>
        </div>
        <div class="table-wrap" style="max-height:520px">
          <table>
            <thead>
              <tr>
                <th>Area</th><th style="text-align:center">Store ID</th><th>Store Name</th><th style="text-align:center">Expected</th><th style="text-align:center">Reported</th><th style="text-align:center">Gap Days</th><th style="text-align:center">Completion</th>
              </tr>
            </thead>
            <tbody id="igGapBody"><tr><td colspan="7" class="empty-cell"><div class="empty-icon"><i class="fa fa-spinner fa-spin"></i></div><p>Loading...</p></td></tr></tbody>
          </table>
        </div>
      </div>
      <div class="table-card">
        <div class="table-header" style="gap:14px;flex-wrap:wrap">
          <div class="table-title"><i class="fa fa-circle-check"></i> Inventory Logs Complete Stores</div>
          <button class="export-btn" onclick="exportStoreDayGapTable('inventory', 'complete')"><i class="fa fa-file-excel"></i> Export Excel</button>
          <div class="table-date" id="igCompleteInfo">-</div>
        </div>
        <div class="table-wrap" style="max-height:520px">
          <table>
            <thead>
              <tr>
                <th>Area</th><th style="text-align:center">Store ID</th><th>Store Name</th><th style="text-align:center">Expected</th><th style="text-align:center">Reported</th><th style="text-align:center">Gap Days</th><th style="text-align:center">Completion</th>
              </tr>
            </thead>
            <tbody id="igCompleteBody"><tr><td colspan="7" class="empty-cell"><div class="empty-icon"><i class="fa fa-spinner fa-spin"></i></div><p>Loading...</p></td></tr></tbody>
          </table>
        </div>
      </div>
    </div>
  </div><!-- /tab-gaps -->

  <!--  STORE NOTES TAB  -->
  <div class="tab-content" id="tab-notes">

    <!-- Controls -->
    <div class="controls">
      <div class="ctrl-group">
        <span class="ctrl-label"><i class="fa fa-layer-group"></i> Area</span>
        <select id="nAreaFilter" onchange="onNAreaChange()">
          <option value="ALL">All Areas</option>
        </select>
      </div>
      <div class="divider"></div>
      <div class="ctrl-group">
        <span class="ctrl-label"><i class="fa fa-store"></i> Store</span>
        <select id="nStoreFilter" onchange="applyNotesFilters()">
          <option value="ALL">All Stores</option>
        </select>
      </div>
      <div class="divider"></div>
      <div class="ctrl-group">
        <span class="ctrl-label"><i class="fa fa-circle-check"></i> Status</span>
        <div id="nStatusFilter" class="check-filter"></div>
      </div>
      <div class="divider"></div>
      <div class="ctrl-group" style="flex:1;min-width:180px">
        <span class="ctrl-label"><i class="fa fa-magnifying-glass"></i> Search</span>
        <input type="text" id="nSearch" placeholder="Search notes, store, remarks..." oninput="debouncedNotesSearch()" style="background:rgba(15,20,35,0.7);border:1px solid var(--border-strong);color:var(--text-1);padding:9px 14px;border-radius:10px;font-size:13px;font-family:'Inter',sans-serif;font-weight:500;outline:none;flex:1;min-width:160px"/>
      </div>
      <div class="records-count" id="nRecordsCount">—</div>
    </div>

    <!-- Notes Table -->
    <div class="table-card">
      <div class="table-header">
        <div class="table-title"><i class="fa fa-note-sticky"></i> Store Notes</div>
        <div class="table-date" id="nTableInfo">Latest entries</div>
      </div>
      <div class="table-wrap" style="max-height:720px">
        <table>
          <thead>
            <tr>
              <th class="sortable" data-sort-key="ts" data-sort-type="num" onclick="sortNotes('ts','num')">Time Stamp <span class="sort-icon">⇅</span></th>
              <th class="sortable" data-sort-key="storeName" data-sort-type="string" onclick="sortNotes('storeName','string')">Store Name <span class="sort-icon">⇅</span></th>
              <th>Notes</th>
              <th class="sortable" data-sort-key="status" data-sort-type="string" onclick="sortNotes('status','string')">Status <span class="sort-icon">⇅</span></th>
              <th>Remarks</th>
              <th style="text-align:center">Photo Link</th>
            </tr>
          </thead>
          <tbody id="nTableBody">
            <tr><td colspan="6" class="empty-cell"><div class="empty-icon"><i class="fa fa-spinner fa-spin"></i></div><p>Loading...</p></td></tr>
          </tbody>
        </table>
      </div>
    </div>

    <!-- Status Performance KPIs -->
    <div class="kpi-grid" style="margin-top:24px">
      <div class="kpi k-sales">
        <div class="kpi-label"><div class="kpi-icon"><i class="fa fa-list-check"></i></div>Total Notes</div>
        <div class="kpi-value gradient-text" id="nKpiTotal">—</div>
        <div class="kpi-sub">In current view</div>
      </div>
      <div class="kpi" style="--accent-color:#10b981">
        <div class="kpi-label"><div class="kpi-icon" style="background:linear-gradient(135deg,#10b981,#34d399)"><i class="fa fa-circle-check"></i></div>Done</div>
        <div class="kpi-value" id="nKpiDone" style="color:#34d399">—</div>
        <div class="kpi-sub" id="nKpiDoneSub">0%</div>
      </div>
      <div class="kpi" style="--accent-color:#f59e0b">
        <div class="kpi-label"><div class="kpi-icon" style="background:linear-gradient(135deg,#f59e0b,#fbbf24)"><i class="fa fa-clock"></i></div>Pending</div>
        <div class="kpi-value" id="nKpiPending" style="color:#fbbf24">—</div>
        <div class="kpi-sub" id="nKpiPendingSub">0%</div>
      </div>
      <div class="kpi" style="--accent-color:#6366f1">
        <div class="kpi-label"><div class="kpi-icon" style="background:linear-gradient(135deg,#6366f1,#818cf8)"><i class="fa fa-spinner"></i></div>Ongoing</div>
        <div class="kpi-value" id="nKpiOngoing" style="color:#818cf8">—</div>
        <div class="kpi-sub" id="nKpiOngoingSub">0%</div>
      </div>
      <div class="kpi" style="--accent-color:#94a3b8">
        <div class="kpi-label"><div class="kpi-icon" style="background:linear-gradient(135deg,#64748b,#94a3b8)"><i class="fa fa-circle-question"></i></div>No Status</div>
        <div class="kpi-value" id="nKpiBlank" style="color:#94a3b8">—</div>
        <div class="kpi-sub" id="nKpiBlankSub">0%</div>
      </div>
    </div>

    <!-- Performance Charts -->
    <div class="charts-grid">
      <div class="chart-card">
        <div class="chart-title">
          <i class="fa fa-ranking-star"></i> Area Status Performance
          <span style="margin-left:auto;font-size:10.5px;color:var(--text-3);font-weight:500;text-transform:uppercase;letter-spacing:0.08em">Best → Worst</span>
        </div>
        <div style="position:relative;height:300px">
          <canvas id="nAreaChart"></canvas>
        </div>
      </div>
      <div class="chart-card">
        <div class="chart-title">
          <i class="fa fa-chart-pie"></i> Overall Status Mix
        </div>
        <div style="position:relative;height:300px">
          <canvas id="nStatusPie"></canvas>
        </div>
      </div>
    </div>

    <!-- Store Performance Chart (full width, scrollable for many stores) -->
    <div class="chart-card" style="margin-top:16px">
      <div class="chart-title">
        <i class="fa fa-store"></i> Store Status Performance
        <span style="margin-left:auto;font-size:10.5px;color:var(--text-3);font-weight:500;text-transform:uppercase;letter-spacing:0.08em">Ranked Best → Worst</span>
      </div>
      <div style="overflow-y:auto;max-height:560px;padding-right:4px">
        <div id="nStoreChartWrap" style="position:relative;height:600px">
          <canvas id="nStoreChart"></canvas>
        </div>
      </div>
    </div>

    <!-- Status Detail Tables -->
    <div class="charts-grid" style="margin-top:16px">
      <div class="table-card">
        <div class="table-header">
          <div class="table-title"><i class="fa fa-map-location-dot"></i> Area Status Breakdown</div>
          <div class="table-date" id="nAreaTableInfo">—</div>
        </div>
        <div class="table-wrap" style="max-height:380px">
          <table>
            <thead>
              <tr>
                <th>Rank</th>
                <th>Area</th>
                <th style="text-align:center">Total</th>
                <th style="text-align:center" title="Done %">Done %</th>
                <th style="text-align:center" title="Pending %">Pending %</th>
                <th style="text-align:center" title="Ongoing %">Ongoing %</th>
                <th style="text-align:center" title="No Status %">Blank %</th>
              </tr>
            </thead>
            <tbody id="nAreaTableBody"></tbody>
          </table>
        </div>
      </div>
      <div class="table-card">
        <div class="table-header">
          <div class="table-title"><i class="fa fa-store"></i> Store Status Breakdown</div>
          <div class="table-date" id="nStoreTableInfo">—</div>
        </div>
        <div class="table-wrap" style="max-height:380px">
          <table>
            <thead>
              <tr>
                <th>Rank</th>
                <th>Store</th>
                <th style="text-align:center">Total</th>
                <th style="text-align:center">Done %</th>
                <th style="text-align:center">Pending %</th>
                <th style="text-align:center">Ongoing %</th>
                <th style="text-align:center">Blank %</th>
              </tr>
            </thead>
            <tbody id="nStoreTableBody"></tbody>
          </table>
        </div>
      </div>
    </div>

  </div><!-- /tab-notes -->

  <!--  ISSUES & CONCERNS TAB  -->
  <div class="tab-content" id="tab-issues">

    <!-- Controls -->
    <div class="controls">
      <div class="ctrl-group">
        <span class="ctrl-label"><i class="fa fa-layer-group"></i> Area</span>
        <select id="iAreaFilter" onchange="onIAreaChange()">
          <option value="ALL">All Areas</option>
        </select>
      </div>
      <div class="divider"></div>
      <div class="ctrl-group">
        <span class="ctrl-label"><i class="fa fa-store"></i> Store</span>
        <select id="iStoreFilter" onchange="applyIssuesFilters()">
          <option value="ALL">All Stores</option>
        </select>
      </div>
      <div class="divider"></div>
      <div class="ctrl-group">
        <span class="ctrl-label"><i class="fa fa-flag"></i> Priority</span>
        <div id="iPriorityFilter" class="check-filter"></div>
      </div>
      <div class="divider"></div>
      <div class="ctrl-group">
        <span class="ctrl-label"><i class="fa fa-circle-check"></i> Status</span>
        <div id="iStatusFilter" class="check-filter"></div>
      </div>
      <div class="divider"></div>
      <div class="ctrl-group" style="flex:1;min-width:160px">
        <span class="ctrl-label"><i class="fa fa-magnifying-glass"></i> Search</span>
        <input type="text" id="iSearch" placeholder="Search description, store, assignee..." oninput="debouncedIssuesSearch()" style="background:rgba(15,20,35,0.7);border:1px solid var(--border-strong);color:var(--text-1);padding:9px 14px;border-radius:10px;font-size:13px;font-family:'Inter',sans-serif;font-weight:500;outline:none;flex:1;min-width:160px"/>
      </div>
      <div class="records-count" id="iRecordsCount">—</div>
    </div>

    <!-- KPI Cards -->
    <div class="kpi-grid">
      <div class="kpi k-sales">
        <div class="kpi-label"><div class="kpi-icon"><i class="fa fa-list-check"></i></div>Total Issues</div>
        <div class="kpi-value gradient-text" id="i-kpi-total">—</div>
        <div class="kpi-sub">In current view</div>
      </div>
      <div class="kpi" style="--accent-color:#f43f5e">
        <div class="kpi-label"><div class="kpi-icon" style="background:linear-gradient(135deg,#f43f5e,#fb7185)"><i class="fa fa-circle-exclamation"></i></div>Open</div>
        <div class="kpi-value" id="i-kpi-open" style="color:#fb7185">—</div>
        <div class="kpi-sub" id="i-kpi-open-sub">—</div>
      </div>
      <div class="kpi" style="--accent-color:#10b981">
        <div class="kpi-label"><div class="kpi-icon" style="background:linear-gradient(135deg,#10b981,#34d399)"><i class="fa fa-circle-check"></i></div>Resolved</div>
        <div class="kpi-value" id="i-kpi-resolved" style="color:#34d399">—</div>
        <div class="kpi-sub" id="i-kpi-resolved-sub">—</div>
      </div>
      <div class="kpi" style="--accent-color:#f59e0b">
        <div class="kpi-label"><div class="kpi-icon" style="background:linear-gradient(135deg,#f59e0b,#fbbf24)"><i class="fa fa-clock"></i></div>Overdue</div>
        <div class="kpi-value" id="i-kpi-overdue" style="color:#fbbf24">—</div>
        <div class="kpi-sub" id="i-kpi-overdue-sub">Past target date</div>
      </div>
      <div class="kpi" style="--accent-color:#6366f1">
        <div class="kpi-label"><div class="kpi-icon" style="background:linear-gradient(135deg,#6366f1,#818cf8)"><i class="fa fa-stopwatch"></i></div>Avg Days Open</div>
        <div class="kpi-value gradient-text" id="i-kpi-avg-days">—</div>
        <div class="kpi-sub">Across all open issues</div>
      </div>
    </div>

    <!-- Charts Row 1 -->
    <div class="charts-grid">
      <div class="chart-card">
        <div class="chart-title"><i class="fa fa-chart-pie"></i> Issues by Priority</div>
        <div style="position:relative;height:300px">
          <canvas id="iPriorityChart"></canvas>
        </div>
      </div>
      <div class="chart-card">
        <div class="chart-title"><i class="fa fa-chart-pie"></i> Issues by Status</div>
        <div style="position:relative;height:300px">
          <canvas id="iStatusChart"></canvas>
        </div>
      </div>
    </div>

    <div class="table-card" style="margin:16px 0 24px">
      <div class="table-header">
        <div class="table-title"><i class="fa fa-triangle-exclamation"></i> High Priority Not Done Summary</div>
        <button class="export-btn" onclick="exportHighPriorityNotDoneToExcel()">
          <i class="fa fa-file-excel"></i> Export Excel
        </button>
        <div class="table-date" id="iHighNotDoneInfo">—</div>
      </div>
      <div class="table-wrap" style="max-height:360px">
        <table>
          <thead>
            <tr>
              <th>Date</th>
              <th>Store Name</th>
              <th>Issue Description</th>
              <th>Impact</th>
              <th>Status</th>
              <th>Remarks</th>
              <th>Last Update</th>
            </tr>
          </thead>
          <tbody id="iHighNotDoneBody"></tbody>
        </table>
      </div>
    </div>

    <!-- Charts Row 2 -->
    <div class="charts-grid">
      <div class="chart-card">
        <div class="chart-title"><i class="fa fa-tags"></i> Issues by Category</div>
        <div style="position:relative;height:340px">
          <canvas id="iCategoryChart"></canvas>
        </div>
      </div>
      <div class="chart-card">
        <div class="chart-title"><i class="fa fa-layer-group"></i> Issues by Area</div>
        <div style="position:relative;height:340px">
          <canvas id="iAreaChart"></canvas>
        </div>
      </div>
    </div>

    <!-- Top Affected Stores -->
    <div class="chart-card" style="margin-top:16px">
      <div class="chart-title">
        <i class="fa fa-ranking-star"></i> Top Affected Stores
        <span style="margin-left:auto;font-size:10.5px;color:var(--text-3);font-weight:500;text-transform:uppercase;letter-spacing:0.08em">By open issues</span>
      </div>
      <div style="overflow-y:auto;max-height:520px;padding-right:4px">
        <div id="iStoreChartWrap" style="position:relative;height:400px">
          <canvas id="iStoreChart"></canvas>
        </div>
      </div>
    </div>

    <!-- Issues Table -->
    <div class="table-card" style="margin-top:24px">
      <div class="table-header">
        <div class="table-title"><i class="fa fa-list-ul"></i> Issues &amp; Concerns Detail</div>
        <button class="export-btn" onclick="exportIssuesDetailToExcel()">
          <i class="fa fa-file-excel"></i> Export Excel
        </button>
        <div class="table-date" id="iTableInfo">—</div>
      </div>
      <div class="table-wrap" style="max-height:720px">
        <table>
          <thead>
            <tr>
              <th class="sortable" data-sort-key="area" data-sort-type="string" onclick="sortIssues('area','string')">Area <span class="sort-icon">⇅</span></th>
              <th class="sortable" data-sort-key="storeId" data-sort-type="string" onclick="sortIssues('storeId','string')">Store ID <span class="sort-icon">⇅</span></th>
              <th class="sortable" data-sort-key="storeName" data-sort-type="string" onclick="sortIssues('storeName','string')">Store Name <span class="sort-icon">⇅</span></th>
              <th class="sortable" data-sort-key="dateTs" data-sort-type="num" onclick="sortIssues('dateTs','num')">Date <span class="sort-icon">⇅</span></th>
              <th class="sortable" data-sort-key="reportedBy" data-sort-type="string" onclick="sortIssues('reportedBy','string')">Reported By <span class="sort-icon">⇅</span></th>
              <th class="sortable" data-sort-key="issueCategory" data-sort-type="string" onclick="sortIssues('issueCategory','string')">Issue Category <span class="sort-icon">⇅</span></th>
              <th class="sortable" data-sort-key="issueSubCategory" data-sort-type="string" onclick="sortIssues('issueSubCategory','string')">Sub Category <span class="sort-icon">⇅</span></th>
              <th>Issue Description</th>
              <th class="sortable" data-sort-key="priority" data-sort-type="string" onclick="sortIssues('priority','string')" style="text-align:center">Priority <span class="sort-icon">⇅</span></th>
              <th class="sortable" data-sort-key="impactLevel" data-sort-type="string" onclick="sortIssues('impactLevel','string')" style="text-align:center">Impact <span class="sort-icon">⇅</span></th>
              <th class="sortable" data-sort-key="assignTo" data-sort-type="string" onclick="sortIssues('assignTo','string')">Assign To <span class="sort-icon">⇅</span></th>
              <th class="sortable" data-sort-key="targetDateTs" data-sort-type="num" onclick="sortIssues('targetDateTs','num')">Target Date <span class="sort-icon">⇅</span></th>
              <th class="sortable" data-sort-key="status" data-sort-type="string" onclick="sortIssues('status','string')" style="text-align:center">Status <span class="sort-icon">⇅</span></th>
              <th>Resolution Details</th>
              <th class="sortable" data-sort-key="dateResolvedTs" data-sort-type="num" onclick="sortIssues('dateResolvedTs','num')">Date Resolved <span class="sort-icon">⇅</span></th>
              <th class="sortable" data-sort-key="daysOpen" data-sort-type="num" onclick="sortIssues('daysOpen','num')" style="text-align:center">Days Open <span class="sort-icon">⇅</span></th>
              <th>Remarks / Notes</th>
              <th>Last Update</th>
            </tr>
          </thead>
          <tbody id="iTableBody">
            <tr><td colspan="18" class="empty-cell"><div class="empty-icon"><i class="fa fa-spinner fa-spin"></i></div><p>Loading...</p></td></tr>
          </tbody>
        </table>
      </div>
    </div>

  </div><!-- /tab-issues -->

</main>

<script>
const nativeFetch = window.fetch.bind(window);
window.fetch = async (...args) => {
  const res = await nativeFetch(...args);
  if (res.status === 401) window.location.href = '/';
  return res;
};

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

// Data gaps monitoring state
let gState = { initialized: false, filters: { months: [], areas: [], stores: [] }, rows: [], completeRows: [], lastUpdateRows: [], daily: { summary: {}, monthly: [], gapRows: [], completeRows: [] }, inventory: { summary: {}, monthly: [], gapRows: [], completeRows: [] } };
let gSort = { key: 'monthKey', dir: 'desc', type: 'string' };
let gCompleteSort = { key: 'area', dir: 'asc', type: 'string' };
let dailyGapChartInst = null;
let inventoryGapChartInst = null;

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
  if (tab === 'gaps' && !gState.initialized) {
    initGapsTab();
  }
  if (tab === 'notes' && !nState.initialized) {
    initNotesTab();
  }
  if (tab === 'issues' && !iState.initialized) {
    initIssuesTab();
  }
  setTimeout(recolorExistingCharts, 50);
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

function isLightTheme() {
  return document.documentElement.classList.contains('theme-light');
}

function chartPalette() {
  return isLightTheme()
    ? {
        strong: '#102b3a',
        muted: '#3e6171',
        grid: 'rgba(15,93,117,0.12)',
        gridSoft: 'rgba(15,93,117,0.08)',
        tooltipBg: 'rgba(255,255,255,0.96)',
        tooltipBorder: 'rgba(8,145,178,0.35)',
        dataLabel: '#102b3a',
      }
    : {
        strong: '#f0f3fb',
        muted: '#a8b3d1',
        grid: 'rgba(148,163,200,0.06)',
        gridSoft: 'rgba(148,163,200,0.04)',
        tooltipBg: 'rgba(15,20,35,0.95)',
        tooltipBorder: 'rgba(99,102,241,0.4)',
        dataLabel: '#e8ecf4',
      };
}

function recolorChartOptions(obj, palette) {
  if (!obj || typeof obj !== 'object') return;
  Object.keys(obj).forEach(key => {
    const val = obj[key];
    if (key === 'titleColor') obj[key] = palette.strong;
    else if (key === 'bodyColor') obj[key] = palette.muted;
    else if (key === 'color') obj[key] = val === '#e8ecf4' ? palette.dataLabel : palette.muted;
    else if (key === 'backgroundColor' && typeof val === 'string' && val.includes('15,20,35')) obj[key] = palette.tooltipBg;
    else if (key === 'borderColor' && typeof val === 'string' && (val.includes('99,102,241') || val.includes('8,145,178'))) obj[key] = palette.tooltipBorder;
    else if (key === 'grid' && val && typeof val === 'object') {
      val.color = val.drawBorder === false ? palette.gridSoft : palette.grid;
      recolorChartOptions(val, palette);
    } else {
      recolorChartOptions(val, palette);
    }
  });
}

function recolorExistingCharts() {
  if (!window.Chart || !Chart.instances) return;
  const palette = chartPalette();
  Object.values(Chart.instances).forEach(chart => {
    if (!chart || !chart.options) return;
    recolorChartOptions(chart.options, palette);
    chart.update('none');
  });
}

function updateThemeToggle() {
  const light = isLightTheme();
  const icon = document.getElementById('themeIcon');
  const text = document.getElementById('themeText');
  const btn = document.getElementById('themeToggle');
  if (icon) icon.className = light ? 'fa fa-moon' : 'fa fa-sun';
  if (text) text.textContent = light ? 'Dark' : 'Daylight';
  if (btn) btn.setAttribute('aria-label', light ? 'Switch to dark mode' : 'Switch to daylight mode');
  const meta = document.querySelector('meta[name="theme-color"]');
  if (meta) meta.setAttribute('content', light ? '#f4faf9' : '#0a0e1a');
}

function toggleTheme() {
  const nextLight = !isLightTheme();
  document.documentElement.classList.toggle('theme-light', nextLight);
  try { localStorage.setItem('camanava-theme', nextLight ? 'light' : 'dark'); } catch (err) {}
  updateThemeToggle();
  recolorExistingCharts();
}

async function apiJson(url) {
  const res = await fetch(url);
  if (res.status === 401) {
    window.location.href = '/';
    throw new Error('Authentication required.');
  }
  const text = await res.text();
  let json = {};
  try {
    json = text ? JSON.parse(text) : {};
  } catch (err) {
    throw new Error('Server returned an unreadable response.');
  }
  if (!res.ok || !json.success) throw new Error(json.error || 'Request failed.');
  return json;
}

async function loadCurrentUser() {
  const json = await apiJson('/api/me');
  const user = json.user || {};
  const areas = (user.areas || []).join(', ');
  const badge = document.getElementById('userBadge');
  if (badge) {
    badge.textContent = (user.username || 'User') + ' · ' + (user.level || 'user') + (areas ? ' · ' + areas : '');
  }
}

async function loadFilters() {
  setSyncing(true);
  try {
    await loadCurrentUser();
    const json = await apiJson('/api/filters');
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
    recolorExistingCharts();
  } catch(e) {
    setStatus('error', ' Error');
    console.error(e);
    showTableError('Failed to load: ' + e.message);
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
      const json = await apiJson(\`/api/filters?area=\${encodeURIComponent(area)}\`);
      populateStores(json.stores);
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
    const json = await apiJson('/api/sales?' + params);
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
    tbody.innerHTML = \`<tr><td colspan="6" class="empty-cell"><div class="empty-icon"><i class="fa fa-magnifying-glass"></i></div><p>No data for selected filters</p><small>Try a different date or area</small></td></tr>\`;
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
    <td style="font-size:12px;letter-spacing:0.1em;text-transform:uppercase;color:var(--indigo2)">TOTAL - \${rows.length} STORES</td>
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
      <td><span style="font-size:13px;color:var(--text-2);font-weight:500">\${s.region || '—'}</span></td>
      <td><span class="remark-pill \${remarkCls}">\${s.remarks || '—'}</span></td>
    </tr>\`;
  }).join('');
}

//  MONTHLY TAB 
async function initMonthlyTab() {
  try {
    // Get available months from /api/months
    const res = await fetch('/api/months');
    const json = await res.json();
    if (!json.success) throw new Error(json.error);

    mMonthFilters.months = json.months || [];

    renderMMonthDropdown();

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

function getSelectedMMonths() {
  return Array.from(document.querySelectorAll('#mMonthMenu input[data-month]:checked')).map(input => input.value);
}

function monthlySelectionLabel(months) {
  if (!months.length) return 'All Months';
  if (months.length === 1) return (mMonthFilters.months.find(m => m.value === months[0]) || {}).label || months[0];
  const first = (mMonthFilters.months.find(m => m.value === months[0]) || {}).label || months[0];
  return first + ' +' + (months.length - 1) + ' more';
}

function updateMMonthTrigger() {
  const trigger = document.getElementById('mMonthTrigger');
  if (trigger) trigger.textContent = monthlySelectionLabel(getSelectedMMonths());
}

function renderMMonthDropdown() {
  const menu = document.getElementById('mMonthMenu');
  if (!menu) return;
  const latestMonth = mMonthFilters.months.length ? mMonthFilters.months[mMonthFilters.months.length - 1].value : '';
  menu.innerHTML =
    '<div class="tick-actions">' +
      '<button class="tick-action" type="button" onclick="selectAllMMonths()">Select All</button>' +
      '<button class="tick-action clear" type="button" onclick="clearMMonths()">Clear</button>' +
    '</div>' +
    mMonthFilters.months.map(m =>
      '<label class="tick-option"><input type="checkbox" data-month value="' + escHtml(m.value) + '"' + (m.value === latestMonth ? ' checked' : '') + '> ' + escHtml(m.label) + '</label>'
    ).join('');
  menu.querySelectorAll('input[data-month]').forEach(input => {
    input.addEventListener('change', onMMonthChange);
  });
  updateMMonthTrigger();
}

function toggleMMonthDropdown() {
  const dd = document.getElementById('mMonthDropdown');
  if (!dd) return;
  const open = dd.classList.toggle('open');
  const controls = dd.closest('.controls');
  if (controls) controls.classList.toggle('dropdown-open', open);
}

function selectAllMMonths() {
  document.querySelectorAll('#mMonthMenu input[data-month]').forEach(input => { input.checked = true; });
  onMMonthChange();
}

function clearMMonths() {
  document.querySelectorAll('#mMonthMenu input[data-month]').forEach(input => { input.checked = false; });
  onMMonthChange();
}

function onMMonthChange() {
  updateMMonthTrigger();
  applyMonthlyFilters();
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
  const months = getSelectedMMonths();
  const area  = document.getElementById('mAreaFilter').value;
  const store = document.getElementById('mStoreFilter').value;

  const params = new URLSearchParams();
  if (months.length) params.set('month', months.join(','));
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

    const monthLabel = monthlySelectionLabel(months);
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
      \`<tr><td colspan="6" class="empty-cell"><div class="empty-icon" style="background:linear-gradient(135deg,rgba(244,63,94,0.1),rgba(251,113,133,0.1));color:#fb7185"><i class="fa fa-triangle-exclamation"></i></div><p>\${e.message}</p></td></tr>\`;
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
    tbody.innerHTML = \`<tr><td colspan="6" class="empty-cell"><div class="empty-icon"><i class="fa fa-magnifying-glass"></i></div><p>No data for selected month</p></td></tr>\`;
    return;
  }

  const totSales = rows.reduce((s,r)=>s+r.sales,0);
  const totLY = rows.reduce((s,r)=>s+r.salesLY,0);
  const totDiff = totSales - totLY;
  const totPct = totLY ? (totDiff/totLY*100) : 0;

  let html = rows.map(r => {
    const pctCls = r.diffPct > 0.05 ? 'up' : r.diffPct < -0.05 ? 'down' : 'flat';
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
      <td style="text-align:right"><span class="num num-bold" style="color:var(--text-1)">\${fmtFull(r.sales)}</span></td>
      <td style="text-align:right"><span class="num" style="color:var(--text-3)">\${fmtFull(r.salesLY)}</span></td>
      <td style="text-align:center"><span class="pill \${pctCls}">\${pctStr}</span></td>
      <td style="text-align:right"><span class="num num-bold" style="color:\${diffColor}">\${r.diffVal>=0?'+':''}\${fmtFull(r.diffVal)}</span></td>
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
      <td><span style="font-size:13px;color:var(--text-2);font-weight:500">\${r.day || '—'}</span></td>
      <td><span style="font-size:13px;color:var(--text-3);font-weight:500">\${r.dayYA || '—'}</span></td>
      <td>
        <div class="store-cell">
          <div class="store-avatar" style="background:linear-gradient(135deg, \${grad[0]}, \${grad[1]});width:34px;height:34px;font-size:12px">\${initials(r.storeName)}</div>
          <div class="store-info">
            <div class="store-name" style="font-size:13.5px">\${r.storeName || '—'}</div>
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

  const months = getSelectedMMonths();
  const area  = document.getElementById('mAreaFilter').value;
  const store = document.getElementById('mStoreFilter').value;
  const monthLabel = monthlySelectionLabel(months);

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
  const safeMonth = (months.length > 1 ? months.length + '_Months' : monthLabel).replace(/[^a-z0-9]/gi, '_');
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

//  CATEGORY SALES TAB 
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
      o.value = m.value || m;
      o.textContent = m.label || m;
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

function selectedOptionText(selectId) {
  const sel = document.getElementById(selectId);
  if (!sel || !sel.selectedOptions || !sel.selectedOptions.length) return '';
  return sel.selectedOptions[0].textContent || sel.value;
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
    if (month !== 'ALL') parts.push(selectedOptionText('cMonthFilter'));
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
    const areaColor = AREA_COLORS[r.area] || DEFAULT_COLOR;
    const grad = AREA_GRADIENTS[r.area] || [DEFAULT_COLOR, DEFAULT_COLOR];
    const diffColor = r.diffVal >= 0 ? '#34d399' : '#fb7185';
    const arrow = r.diffPct > 0.05 ? '↑' : r.diffPct < -0.05 ? '↓' : '—';
    const pctStr = r.diffPct !== 0 ? \`\${arrow} \${Math.abs(r.diffPct).toFixed(2)}%\` : '—';

    return \`<tr>
      <td><span class="area-tag"><span class="area-dot" style="background:\${color};color:\${color}"></span>\${r.category || '—'}</span></td>
      <td>
        <div class="store-cell">
          <div class="store-avatar" style="background:linear-gradient(135deg, \${grad[0]}, \${grad[1]});width:34px;height:34px;font-size:12px">\${initials(r.storeName)}</div>
          <div class="store-info">
            <div class="store-name" style="font-size:13.5px">\${r.storeName || '—'}</div>
            <div class="store-id">\${r.area || ''}</div>
          </div>
        </div>
      </td>
      <td><span style="font-size:13.5px;color:var(--text-1);font-weight:500">\${r.subDepName || '—'}</span></td>
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
    if (month !== 'ALL') parts.push(selectedOptionText('cMonthFilter'));
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

async function initGapsTab() {
  try {
    const res = await fetch('/api/data-gap-filters');
    const json = await res.json();
    if (!json.success) throw new Error(json.error);

    gState.filters = json;

    const mSel = document.getElementById('gMonthFilter');
    mSel.innerHTML = '<option value="ALL">All Months</option>';
    json.months.forEach(m => {
      const o = document.createElement('option');
      o.value = m.value || m;
      o.textContent = m.label || m;
      mSel.appendChild(o);
    });

    const aSel = document.getElementById('gAreaFilter');
    aSel.innerHTML = '<option value="ALL">All Areas</option>';
    json.areas.forEach(a => {
      const o = document.createElement('option');
      o.value = o.textContent = a;
      aSel.appendChild(o);
    });

    populateGStores(json.stores);
    gState.initialized = true;
    await applyGapsFilters();
  } catch(e) {
    console.error('Data gaps init error:', e);
    const body = document.getElementById('gMonitorBody');
    if (body) body.innerHTML = '<tr><td colspan="5" class="empty-cell"><div class="empty-icon" style="background:linear-gradient(135deg,rgba(244,63,94,0.1),rgba(251,113,133,0.1));color:#fb7185"><i class="fa fa-triangle-exclamation"></i></div><p>' + escHtml(e.message) + '</p></td></tr>';
  }
}

function populateGStores(stores) {
  const sel = document.getElementById('gStoreFilter');
  sel.innerHTML = '<option value="ALL">All Stores</option>';
  stores.forEach(s => {
    const o = document.createElement('option');
    o.value = o.textContent = s;
    sel.appendChild(o);
  });
}

async function onGAreaChange() {
  const area = document.getElementById('gAreaFilter').value;
  if (area === 'ALL') {
    populateGStores(gState.filters.stores);
  } else {
    try {
      const res = await fetch('/api/data-gap-filters?area=' + encodeURIComponent(area));
      const json = await res.json();
      if (json.success) populateGStores(json.stores);
    } catch(e) {}
  }
  document.getElementById('gStoreFilter').value = 'ALL';
  await applyGapsFilters();
}

async function applyGapsFilters() {
  const month = document.getElementById('gMonthFilter').value;
  const area = document.getElementById('gAreaFilter').value;
  const store = document.getElementById('gStoreFilter').value;

  const params = new URLSearchParams();
  if (month !== 'ALL') params.set('month', month);
  if (area !== 'ALL') params.set('area', area);
  if (store !== 'ALL') params.set('store', store);

  try {
    const [categoryRes, dailyRes, inventoryRes] = await Promise.all([
      fetch('/api/data-gaps?' + params),
      fetch('/api/daily-sales-gaps?' + params),
      fetch('/api/inventory-log-gaps?' + params)
    ]);
    const categoryJson = await categoryRes.json();
    const dailyJson = await dailyRes.json();
    const inventoryJson = await inventoryRes.json();
    if (!categoryJson.success) throw new Error(categoryJson.error);
    if (!dailyJson.success) throw new Error(dailyJson.error);
    if (!inventoryJson.success) throw new Error(inventoryJson.error);

    gState.rows = categoryJson.gapRows || categoryJson.rows || [];
    gState.completeRows = categoryJson.completeRows || [];
    gState.lastUpdateRows = categoryJson.lastUpdateRows || [];
    gState.daily = {
      summary: dailyJson.summary || {},
      monthly: dailyJson.monthly || [],
      gapRows: dailyJson.gapRows || [],
      completeRows: dailyJson.completeRows || []
    };
    gState.inventory = {
      summary: inventoryJson.summary || {},
      monthly: inventoryJson.monthly || [],
      gapRows: inventoryJson.gapRows || [],
      completeRows: inventoryJson.completeRows || []
    };

    renderGapsTable(sortRows(gState.rows, gSort), 'gMonitorBody', 'gap');
    renderGapsTable(sortRows(categoryCompleteStores(), gCompleteSort), 'gCompleteBody', 'complete');
    renderCategoryLastUpdateTable(gState.lastUpdateRows || []);
    renderCategoryGapKpis();
    updateSortHeaders('gMonitorBody', gSort);
    updateSortHeaders('gCompleteBody', gCompleteSort);
    renderDailyGapSection();
    renderInventoryGapSection();

    const parts = [];
    if (month !== 'ALL') parts.push(selectedOptionText('gMonthFilter'));
    if (area !== 'ALL') parts.push(area);
    if (store !== 'ALL') parts.push(store);
    const filterText = parts.length ? parts.join(' - ') : 'All months';

    document.getElementById('gRecordsCount').innerHTML = '<span>' + gState.rows.length + '</span> category gap - <span>' + categoryCompleteStores().length + '</span> complete stores - <span>' + (gState.daily.summary.totalGapDays || 0) + '</span> daily gap days - <span>' + (gState.inventory.summary.totalGapDays || 0) + '</span> inventory gap days';
    document.getElementById('gMonitorInfo').textContent = gState.rows.length ? (gState.rows.length + ' store-month gap' + (gState.rows.length !== 1 ? 's' : '') + ' - ' + filterText) : 'No category gaps - ' + filterText;
    document.getElementById('gCompleteInfo').textContent = categoryCompleteStores().length + ' complete store' + (categoryCompleteStores().length !== 1 ? 's' : '') + ' - ' + filterText;
  } catch(e) {
    console.error(e);
    document.getElementById('gMonitorBody').innerHTML = '<tr><td colspan="5" class="empty-cell"><div class="empty-icon" style="background:linear-gradient(135deg,rgba(244,63,94,0.1),rgba(251,113,133,0.1));color:#fb7185"><i class="fa fa-triangle-exclamation"></i></div><p>' + escHtml(e.message) + '</p></td></tr>';
  }
}

function clientNormalizeKey(value) {
  return String(value || '').trim().replace(/\s+/g, ' ').toLowerCase();
}

function clientCompactKey(value) {
  return clientNormalizeKey(value).replace(/[^a-z0-9]/g, '');
}

function renderCategoryLastUpdateTable(rows) {
  const tbody = document.getElementById('gLastUpdateBody');
  const info = document.getElementById('gLastUpdateInfo');
  if (!tbody) return;
  if (info) info.textContent = (rows || []).length + ' unique store' + ((rows || []).length !== 1 ? 's' : '');
  if (!rows || !rows.length) {
    tbody.innerHTML = '<tr><td colspan="7" class="empty-cell"><div class="empty-icon"><i class="fa fa-magnifying-glass"></i></div><p>No Category Sales update data found</p></td></tr>';
    return;
  }
  tbody.innerHTML = rows.map(r => {
    const sales = Number(r.sales || 0);
    const salesYA = Number(r.salesYA || 0);
    const diffValue = sales - salesYA;
    const diffPct = salesYA ? (diffValue / salesYA) * 100 : (sales ? 100 : 0);
    const diffColor = diffValue >= 0 ? 'var(--good)' : 'var(--bad)';
    const dailyDiffPct = r.dailySalesDiffPct;
    const dailyDiffColor = Number(dailyDiffPct || 0) >= 0 ? 'var(--good)' : 'var(--bad)';
    const dailyDiffText = dailyDiffPct === null || dailyDiffPct === undefined || dailyDiffPct === '' ? '-' : Number(dailyDiffPct).toFixed(2) + '%';
    return '<tr>' +
      '<td><div class="store-name">' + escHtml(r.storeName || '-') + '</div></td>' +
      '<td style="text-align:right"><span class="num num-bold" style="color:var(--text-1)">' + fmtFull(sales) + '</span></td>' +
      '<td style="text-align:right"><span class="num" style="color:var(--text-3)">' + fmtFull(salesYA) + '</span></td>' +
      '<td style="text-align:right"><span class="num num-bold" style="color:' + diffColor + '">' + diffPct.toFixed(2) + '%</span></td>' +
      '<td style="text-align:right"><span class="num num-bold" style="color:' + diffColor + '">' + fmtFull(diffValue) + '</span></td>' +
      '<td style="text-align:right"><span class="num num-bold" style="color:' + dailyDiffColor + '">' + dailyDiffText + '</span></td>' +
      '<td><span class="timestamp-cell">' + (escHtml(r.lastUpdate) || '-') + '</span></td>' +
    '</tr>';
  }).join('');
}
function exportCategoryLastUpdateToExcel() {
  const rows = gState.lastUpdateRows || [];
  if (!rows.length) {
    alert('No data to export with the current filters.');
    return;
  }
  const data = rows.map(r => {
    const sales = Number(r.sales || 0);
    const salesYA = Number(r.salesYA || 0);
    const diffValue = sales - salesYA;
    const diffPct = salesYA ? (diffValue / salesYA) * 100 : (sales ? 100 : 0);
    return {
      'Store Name': r.storeName || '',
      'Sales': sales,
      'SalesYA': salesYA,
      'Sales Diff %': Number(diffPct.toFixed(2)),
      'Diff Value': diffValue,
      'Diff % (DS)': r.dailySalesDiffPct === null || r.dailySalesDiffPct === undefined ? '' : Number(r.dailySalesDiffPct),
      'Last Update': r.lastUpdate || ''
    };
  });
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(data);
  ws['!cols'] = [{ wch: 30 }, { wch: 16 }, { wch: 16 }, { wch: 16 }, { wch: 16 }, { wch: 16 }, { wch: 24 }];
  if (ws['!ref']) {
    const range = XLSX.utils.decode_range(ws['!ref']);
    const headerStyle = {
      fill: { fgColor: { rgb: '166534' } },
      font: { color: { rgb: 'FFFFFF' }, bold: true },
      alignment: { horizontal: 'center', vertical: 'center' }
    };
    for (let C = range.s.c; C <= range.e.c; C++) {
      const addr = XLSX.utils.encode_cell({ r: 0, c: C });
      if (ws[addr]) ws[addr].s = headerStyle;
    }
  }
  XLSX.utils.book_append_sheet(wb, ws, 'Category Last Update');
  XLSX.writeFile(wb, 'CaMaNaVa_Category_Last_Update_' + issueExportDateTag() + '.xlsx');
}
function renderCategoryGapKpis() {
  const gapRows = gState.rows || [];
  const completeRows = gState.completeRows || [];
  const completeStores = categoryCompleteStores();
  const affectedStores = new Set(gapRows.map(categoryStoreKey));
  const gapMonths = new Set(gapRows.map(r => r.monthKey || r.month).filter(Boolean));
  const totalStoreMonths = gapRows.length + completeRows.length;
  const completionPct = totalStoreMonths ? ((completeRows.length / totalStoreMonths) * 100) : 0;
  const setText = (id, value) => {
    const el = document.getElementById(id);
    if (el) el.textContent = value;
  };
  setText('cgGapRows', gapRows.length.toLocaleString('en-PH'));
  setText('cgAffectedStores', affectedStores.size.toLocaleString('en-PH'));
  setText('cgCompleteStores', completeStores.length.toLocaleString('en-PH'));
  setText('cgCompletion', completionPct.toFixed(2) + '%');
  setText('cgGapMonths', gapMonths.size.toLocaleString('en-PH'));
}
function categoryStoreKey(row) {
  return clientNormalizeKey(row.area) + '|' + clientNormalizeKey(row.storeId) + '|' + clientCompactKey(row.storeName);
}

function categoryCompleteStores() {
  const gapKeys = new Set((gState.rows || []).map(categoryStoreKey));
  const seen = new Set();
  return (gState.completeRows || [])
    .filter(r => !gapKeys.has(categoryStoreKey(r)))
    .filter(r => {
      const key = categoryStoreKey(r);
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    })
    .map(r => ({
      area: r.area,
      storeId: r.storeId,
      storeName: r.storeName,
      remarks: 'Complete Category Sales data for the selected period'
    }));
}
function renderGapsTable(rows, tbodyId, mode) {
  const tbody = document.getElementById(tbodyId);
  if (!tbody) return;
  const isComplete = mode === 'complete';
  const colspan = isComplete ? 4 : 5;
  if (!rows.length) {
    const msg = isComplete ? 'No stores with complete Category Sales data found' : 'All expected stores have Category Sales data';
    tbody.innerHTML = '<tr><td colspan="' + colspan + '" class="empty-cell"><div class="empty-icon"><i class="fa fa-circle-check"></i></div><p>' + msg + '</p><small>Based on the selected month, area, and store filters</small></td></tr>';
    return;
  }
  tbody.innerHTML = rows.map(r => {
    const color = AREA_COLORS[r.area] || DEFAULT_COLOR;
    if (isComplete) {
      return '<tr>' +
        '<td><span class="area-tag"><span class="area-dot" style="background:' + color + ';color:' + color + '"></span>' + (escHtml(r.area) || '-') + '</span></td>' +
        '<td style="text-align:center"><span class="num num-bold" style="color:var(--text-1)">#' + (escHtml(r.storeId) || '-') + '</span></td>' +
        '<td><div class="store-name">' + (escHtml(r.storeName) || '-') + '</div></td>' +
        '<td><div class="remarks-cell" style="text-align:left">' + (escHtml(r.remarks) || '-') + '</div></td>' +
      '</tr>';
    }
    return '<tr>' +
      '<td><span class="timestamp-cell">' + (escHtml(r.month) || '-') + '</span></td>' +
      '<td><span class="area-tag"><span class="area-dot" style="background:' + color + ';color:' + color + '"></span>' + (escHtml(r.area) || '-') + '</span></td>' +
      '<td style="text-align:center"><span class="num num-bold" style="color:var(--text-1)">#' + (escHtml(r.storeId) || '-') + '</span></td>' +
      '<td><div class="store-name">' + (escHtml(r.storeName) || '-') + '</div></td>' +
      '<td><div class="remarks-cell" style="text-align:left">' + (escHtml(r.remarks) || '-') + '</div></td>' +
    '</tr>';
  }).join('');
}

function renderDailyGapSection() {
  const summary = gState.daily.summary || {};
  document.getElementById('dgExpected').textContent = (summary.totalExpectedStoreDays || 0).toLocaleString('en-PH');
  document.getElementById('dgReported').textContent = (summary.totalReportedStoreDays || 0).toLocaleString('en-PH');
  document.getElementById('dgGapDays').textContent = (summary.totalGapDays || 0).toLocaleString('en-PH');
  document.getElementById('dgCompletion').textContent = (summary.completionPct || 0).toFixed(2) + '%';
  document.getElementById('dgStoreMonths').textContent = (summary.storeMonths || 0).toLocaleString('en-PH');
  document.getElementById('dgStoreMonthsSub').textContent = (summary.gapStoreMonths || 0) + ' with gap � ' + (summary.completeStoreMonths || 0) + ' complete';
  document.getElementById('dgChartInfo').textContent = summary.throughDate ? ('Through ' + summary.throughDate) : 'January to yesterday';
  document.getElementById('dgGapInfo').textContent = storeDaySummaryRows(gState.daily, 'gap').length + ' stores with gaps';
  document.getElementById('dgCompleteInfo').textContent = storeDaySummaryRows(gState.daily, 'complete').length + ' complete stores';
  renderDailyGapChart(gState.daily.monthly || []);
  renderDailyRows(storeDaySummaryRows(gState.daily, 'gap'), 'dgGapBody', true);
  renderDailyRows(storeDaySummaryRows(gState.daily, 'complete'), 'dgCompleteBody', false);
}

function renderInventoryGapSection() {
  const summary = gState.inventory.summary || {};
  document.getElementById('igExpected').textContent = (summary.totalExpectedStoreDays || 0).toLocaleString('en-PH');
  document.getElementById('igReported').textContent = (summary.totalReportedStoreDays || 0).toLocaleString('en-PH');
  document.getElementById('igGapDays').textContent = (summary.totalGapDays || 0).toLocaleString('en-PH');
  document.getElementById('igCompletion').textContent = (summary.completionPct || 0).toFixed(2) + '%';
  document.getElementById('igStoreMonths').textContent = (summary.storeMonths || 0).toLocaleString('en-PH');
  document.getElementById('igStoreMonthsSub').textContent = (summary.gapStoreMonths || 0) + ' with gap - ' + (summary.completeStoreMonths || 0) + ' complete';
  document.getElementById('igChartInfo').textContent = summary.throughDate ? ('Through ' + summary.throughDate) : 'February 2 to today';
  document.getElementById('igGapInfo').textContent = storeDaySummaryRows(gState.inventory, 'gap').length + ' stores with gaps';
  document.getElementById('igCompleteInfo').textContent = storeDaySummaryRows(gState.inventory, 'complete').length + ' complete stores';
  renderInventoryGapChart(gState.inventory.monthly || []);
  renderDailyRows(storeDaySummaryRows(gState.inventory, 'gap'), 'igGapBody', true);
  renderDailyRows(storeDaySummaryRows(gState.inventory, 'complete'), 'igCompleteBody', false);
}

function renderInventoryGapChart(monthly) {
  const canvas = document.getElementById('inventoryGapChart');
  if (!canvas) return;
  if (inventoryGapChartInst) inventoryGapChartInst.destroy();
  const palette = chartPalette();
  inventoryGapChartInst = new Chart(canvas, {
    type: 'bar',
    data: {
      labels: monthly.map(r => r.month),
      datasets: [
        { label: 'Gap Days', data: monthly.map(r => r.gapDays), backgroundColor: '#f59e0b', borderRadius: 6, yAxisID: 'y' },
        { label: 'Completion %', data: monthly.map(r => r.completionPct), type: 'line', borderColor: '#10b981', backgroundColor: '#10b981', tension: 0.35, pointRadius: 3, yAxisID: 'y1' }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { intersect: false, mode: 'index' },
      plugins: { legend: { labels: { color: palette.muted, font: { family: "'Inter'", size: 11 } } } },
      scales: {
        x: { grid: { color: palette.gridSoft }, ticks: { color: palette.muted, font: { family: "'Inter'", size: 10.5, weight: '600' } } },
        y: { beginAtZero: true, grid: { color: palette.grid }, ticks: { color: palette.muted, font: { family: "'JetBrains Mono'", size: 10.5 } }, title: { display: true, text: 'Gap Days', color: palette.muted } },
        y1: { beginAtZero: true, max: 100, position: 'right', grid: { drawOnChartArea: false }, ticks: { color: palette.muted, callback: v => v + '%' }, title: { display: true, text: 'Completion %', color: palette.muted } }
      }
    }
  });
}
function summarizeStoreDayRows(rows) {
  const map = new Map();
  (rows || []).forEach(r => {
    const key = clientNormalizeKey(r.area) + '|' + clientNormalizeKey(r.storeId) + '|' + clientCompactKey(r.storeName);
    if (!map.has(key)) {
      map.set(key, { area: r.area, storeId: r.storeId, storeName: r.storeName, expectedDays: 0, reportedDays: 0, gapDays: 0 });
    }
    const item = map.get(key);
    item.expectedDays += r.expectedDays || 0;
    item.reportedDays += r.reportedDays || 0;
    item.gapDays += r.gapDays || 0;
  });
  return [...map.values()].map(r => ({
    ...r,
    completionPct: r.expectedDays ? parseFloat(((r.reportedDays / r.expectedDays) * 100).toFixed(2)) : 0,
  })).sort((a, b) => b.gapDays - a.gapDays || (a.area || '').localeCompare(b.area || '') || (a.storeName || '').localeCompare(b.storeName || ''));
}
function storeDaySummaryRows(dataSource, mode) {
  const summary = summarizeStoreDayRows([...(dataSource.gapRows || []), ...(dataSource.completeRows || [])]);
  return summary.filter(r => mode === 'complete' ? r.gapDays === 0 : r.gapDays > 0);
}
function renderDailyRows(rows, tbodyId, isGap) {
  const tbody = document.getElementById(tbodyId);
  if (!tbody) return;
  const summaryRows = summarizeStoreDayRows(rows);
  const colspan = 7;
  if (!summaryRows.length) {
    tbody.innerHTML = '<tr><td colspan="' + colspan + '" class="empty-cell"><div class="empty-icon"><i class="fa fa-circle-check"></i></div><p>' + (isGap ? 'No gaps found' : 'No complete stores found') + '</p></td></tr>';
    return;
  }
  tbody.innerHTML = summaryRows.map(r => {
    const color = AREA_COLORS[r.area] || DEFAULT_COLOR;
    const pct = (r.completionPct || 0).toFixed(2) + '%';
    return '<tr>' +
      '<td><span class="area-tag"><span class="area-dot" style="background:' + color + ';color:' + color + '"></span>' + escHtml(r.area) + '</span></td>' +
      '<td style="text-align:center"><span class="num num-bold" style="color:var(--text-1)">#' + escHtml(r.storeId) + '</span></td>' +
      '<td><div class="store-name">' + escHtml(r.storeName) + '</div></td>' +
      '<td style="text-align:center"><span class="num">' + r.expectedDays + '</span></td>' +
      '<td style="text-align:center"><span class="num">' + r.reportedDays + '</span></td>' +
      '<td style="text-align:center"><span class="pill ' + (r.gapDays > 0 ? 'down' : 'up') + '">' + r.gapDays + '</span></td>' +
      '<td style="text-align:center"><span class="pill ' + (r.completionPct >= 99.999 ? 'up' : 'down') + '">' + pct + '</span></td>' +
    '</tr>';
  }).join('');
}

function renderDailyGapChart(monthly) {
  const canvas = document.getElementById('dailyGapChart');
  if (!canvas) return;
  if (dailyGapChartInst) dailyGapChartInst.destroy();
  const palette = chartPalette();
  dailyGapChartInst = new Chart(canvas, {
    type: 'bar',
    data: {
      labels: monthly.map(r => r.month),
      datasets: [
        { label: 'Gap Days', data: monthly.map(r => r.gapDays), backgroundColor: '#f43f5e', borderRadius: 6, yAxisID: 'y' },
        { label: 'Completion %', data: monthly.map(r => r.completionPct), type: 'line', borderColor: '#10b981', backgroundColor: '#10b981', tension: 0.35, pointRadius: 3, yAxisID: 'y1' }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { intersect: false, mode: 'index' },
      plugins: { legend: { labels: { color: palette.muted, font: { family: "'Inter'", size: 11 } } } },
      scales: {
        x: { grid: { color: palette.gridSoft }, ticks: { color: palette.muted, font: { family: "'Inter'", size: 10.5, weight: '600' } } },
        y: { beginAtZero: true, grid: { color: palette.grid }, ticks: { color: palette.muted, font: { family: "'JetBrains Mono'", size: 10.5 } }, title: { display: true, text: 'Gap Days', color: palette.muted } },
        y1: { beginAtZero: true, max: 100, position: 'right', grid: { drawOnChartArea: false }, ticks: { color: palette.muted, callback: v => v + '%' }, title: { display: true, text: 'Completion %', color: palette.muted } }
      }
    }
  });
}

function exportStoreDayGapTable(source, mode) {
  const dataSource = source === 'inventory' ? gState.inventory : gState.daily;
  const rows = storeDaySummaryRows(dataSource, mode);
  const sourceLabel = source === 'inventory' ? 'Inventory Logs' : 'Daily Sales';
  const modeLabel = mode === 'complete' ? 'Complete' : 'Gap';
  if (!rows.length) {
    alert('No data to export with the current filters.');
    return;
  }

  const summaryRows = rows;
  const data = summaryRows.map(r => ({
    'Area': r.area || '',
    'Store ID': r.storeId || '',
    'Store Name': r.storeName || '',
    'Expected': r.expectedDays || 0,
    'Reported': r.reportedDays || 0,
    'Gap Days': r.gapDays || 0,
    'Completion %': r.completionPct || 0
  }));

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(data);
  ws['!cols'] = [{ wch: 24 }, { wch: 12 }, { wch: 30 }, { wch: 14 }, { wch: 14 }, { wch: 12 }, { wch: 14 }];

  if (ws['!ref']) {
    const range = XLSX.utils.decode_range(ws['!ref']);
    const headerStyle = {
      fill: { fgColor: { rgb: '166534' } },
      font: { color: { rgb: 'FFFFFF' }, bold: true },
      alignment: { horizontal: 'center', vertical: 'center' }
    };
    for (let C = range.s.c; C <= range.e.c; C++) {
      const addr = XLSX.utils.encode_cell({ r: 0, c: C });
      if (ws[addr]) ws[addr].s = headerStyle;
    }
    for (let R = 1; R <= range.e.r; R++) {
      for (let C = range.s.c; C <= range.e.c; C++) {
        const addr = XLSX.utils.encode_cell({ r: R, c: C });
        if (!ws[addr]) continue;
        const centerCols = [1, 3, 4, 5, 6];
        ws[addr].s = { alignment: { horizontal: centerCols.includes(C) ? 'center' : 'left', vertical: 'center', wrapText: false } };
      }
    }
  }

  XLSX.utils.book_append_sheet(wb, ws, sourceLabel + ' ' + modeLabel);
  XLSX.writeFile(wb, 'CaMaNaVa_' + sourceLabel.replace(/\s+/g, '_') + '_' + modeLabel + '_' + issueExportDateTag() + '.xlsx');
}
function sortGaps(key, type) {
  if (gSort.key === key) gSort.dir = gSort.dir === 'asc' ? 'desc' : 'asc';
  else gSort = { key, dir: key === 'monthKey' ? 'desc' : 'asc', type };
  gSort.type = type;
  renderGapsTable(sortRows(gState.rows, gSort), 'gMonitorBody', 'gap');
  updateSortHeaders('gMonitorBody', gSort);
}

function sortGapsComplete(key, type) {
  if (gCompleteSort.key === key) gCompleteSort.dir = gCompleteSort.dir === 'asc' ? 'desc' : 'asc';
  else gCompleteSort = { key, dir: 'asc', type };
  gCompleteSort.type = type;
  renderGapsTable(sortRows(categoryCompleteStores(), gCompleteSort), 'gCompleteBody', 'complete');
    renderCategoryLastUpdateTable(gState.lastUpdateRows || []);
  updateSortHeaders('gCompleteBody', gCompleteSort);
}

function exportGapsToExcel(mode) {
  const isComplete = mode === 'complete';
  const rows = isComplete ? categoryCompleteStores() : gState.rows;
  const sortState = isComplete ? gCompleteSort : gSort;
  if (!rows.length) {
    alert('No data to export with the current filters.');
    return;
  }

  const data = sortRows(rows, sortState).map(r => isComplete ? ({
    'Area': r.area || '',
    'Store ID': r.storeId || '',
    'Store Name': r.storeName || '',
    'Remarks': r.remarks || ''
  }) : ({
    'Month': r.month || '',
    'Area': r.area || '',
    'Store ID': r.storeId || '',
    'Store Name': r.storeName || '',
    'Remarks': r.remarks || ''
  }));

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(data);
  ws['!cols'] = isComplete ? [{ wch: 24 }, { wch: 12 }, { wch: 30 }, { wch: 48 }] : [{ wch: 16 }, { wch: 24 }, { wch: 12 }, { wch: 30 }, { wch: 48 }];

  if (ws['!ref']) {
    const range = XLSX.utils.decode_range(ws['!ref']);
    const headerStyle = {
      fill: { fgColor: { rgb: '166534' } },
      font: { color: { rgb: 'FFFFFF' }, bold: true },
      alignment: { horizontal: 'center', vertical: 'center' }
    };
    for (let C = range.s.c; C <= range.e.c; C++) {
      const addr = XLSX.utils.encode_cell({ r: 0, c: C });
      if (ws[addr]) ws[addr].s = headerStyle;
    }
    for (let R = 1; R <= range.e.r; R++) {
      for (let C = range.s.c; C <= range.e.c; C++) {
        const addr = XLSX.utils.encode_cell({ r: R, c: C });
        if (!ws[addr]) continue;
        ws[addr].s = { alignment: { horizontal: C === 2 ? 'center' : 'left', vertical: 'center', wrapText: C === 4 } };
      }
    }
  }

  const sheetName = isComplete ? 'Category Complete' : 'Category Gaps';
  const fileTag = isComplete ? 'Category_Complete' : 'Category_Gaps';
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  XLSX.writeFile(wb, 'CaMaNaVa_' + fileTag + '_' + issueExportDateTag() + '.xlsx');
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
  const monthText = selectedOptionText('cMonthFilter');
  const category = document.getElementById('cCategoryFilter').value;
  const area = document.getElementById('cAreaFilter').value;
  const store = document.getElementById('cStoreFilter').value;

  const data = cDetailRowsCache.map(r => ({
    'Category': r.category || '',
    'Store Name': r.storeName || '',
    'Area': r.area || '',
    'Sub-Department': r.subDepName || '',
    'Sales': r.sales || 0,
    'Sales LY': r.salesLY || 0,
    'Diff %': r.diffPct || 0,
    'Diff Amount': r.diffVal || 0,
  }));

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(data);
  ws['!cols'] = [{wch:20},{wch:22},{wch:18},{wch:36},{wch:16},{wch:16},{wch:10},{wch:16}];

  const headerStyle = { fill:{fgColor:{rgb:'166534'}}, font:{color:{rgb:'FFFFFF'},bold:true}, alignment:{horizontal:'center',vertical:'center'} };
  const range = XLSX.utils.decode_range(ws['!ref']);
  for (let C = range.s.c; C <= range.e.c; C++) {
    const addr = XLSX.utils.encode_cell({ r: 0, c: C });
    if (ws[addr]) ws[addr].s = headerStyle;
  }
  const numCols = { 4:'#,##0.00', 5:'#,##0.00', 6:'0.00"%"', 7:'#,##0.00' };
  for (let R = 1; R <= range.e.r; R++) {
    for (const [col, fmt] of Object.entries(numCols)) {
      const addr = XLSX.utils.encode_cell({ r: R, c: parseInt(col) });
      if (ws[addr]) ws[addr].z = fmt;
    }
  }
  XLSX.utils.book_append_sheet(wb, ws, 'Category Detail');

  const parts = ['CaMaNaVa_Category'];
  if (month !== 'ALL') parts.push(monthText.replace(/[^a-z0-9]/gi,'_'));
  if (category !== 'ALL') parts.push(category.replace(/[^a-z0-9]/gi,'_'));
  if (area !== 'ALL') parts.push(area.replace(/[^a-z0-9]/gi,'_'));
  if (store !== 'ALL') parts.push(store.replace(/[^a-z0-9]/gi,'_'));
  if (cSignFilter !== 'ALL') parts.push(cSignFilter === 'POS' ? 'Positive' : 'Negative');
  XLSX.writeFile(wb, parts.join('_') + '.xlsx');
}

//  STORE NOTES TAB 
const nState = {
  initialized: false,
  rows: [],
  filters: { areas: [], stores: [], statuses: [] },
  sort: { key: 'ts', dir: 'desc', type: 'num' },
  areaChart: null,
  storeChart: null,
  statusPie: null,
};

async function initNotesTab() {
  try {
    const res = await fetch('/api/store-notes-filters');
    const json = await res.json();
    if (!json.success) throw new Error(json.error);
    nState.filters = json;

    const aSel = document.getElementById('nAreaFilter');
    aSel.innerHTML = '<option value="ALL">All Areas</option>';
    json.areas.forEach(a => {
      const o = document.createElement('option');
      o.value = o.textContent = a;
      aSel.appendChild(o);
    });

    populateNStores(json.stores);

    renderNotesStatusFilter(json.statuses);

    nState.initialized = true;
    await applyNotesFilters();
  } catch(e) {
    console.error('Notes init error:', e);
    document.getElementById('nTableBody').innerHTML =
      \`<tr><td colspan="6" class="empty-cell"><div class="empty-icon" style="background:linear-gradient(135deg,rgba(244,63,94,0.1),rgba(251,113,133,0.1));color:#fb7185"><i class="fa fa-triangle-exclamation"></i></div><p>\${e.message}</p></td></tr>\`;
  }
}

function populateNStores(stores) {
  const sel = document.getElementById('nStoreFilter');
  sel.innerHTML = '<option value="ALL">All Stores</option>';
  stores.forEach(s => {
    const o = document.createElement('option');
    o.value = o.textContent = s;
    sel.appendChild(o);
  });
}

function renderNotesStatusFilter(statuses) {
  const wrap = document.getElementById('nStatusFilter');
  if (!wrap) return;
  wrap.innerHTML = '';

  const noteStatusKey = s => String(s || '').trim().replace(/\s+/g, ' ').toLowerCase();

  const options = [
    { value: 'ALL', label: 'All Status', checked: true },
    { value: '__blank__', label: 'All Blank', checked: false },
    ...statuses
      .filter(s => noteStatusKey(s) && noteStatusKey(s) !== 'status')
      .map(s => ({ value: s, label: s, checked: false }))
  ];

  options.forEach(opt => {
    const label = document.createElement('label');
    label.innerHTML = '<input type="checkbox" value="' + escHtml(opt.value) + '"' + (opt.checked ? ' checked' : '') + '> ' + escHtml(opt.label);
    wrap.appendChild(label);
  });

  wrap.querySelectorAll('input').forEach(input => {
    input.addEventListener('change', e => onNotesStatusFilterChange(e));
  });
}

function onNotesStatusFilterChange(e) {
  const wrap = document.getElementById('nStatusFilter');
  if (!wrap) return;
  const all = wrap.querySelector('input[value="ALL"]');
  const items = Array.from(wrap.querySelectorAll('input:not([value="ALL"])'));
  if (e && e.target === all && all.checked) {
    items.forEach(input => { input.checked = false; });
  } else {
    const anyChecked = items.some(input => input.checked);
    if (all) all.checked = !anyChecked;
  }
  applyNotesFilters();
}

function getNotesStatusValues() {
  const wrap = document.getElementById('nStatusFilter');
  if (!wrap) return [];
  return Array.from(wrap.querySelectorAll('input:checked'))
    .map(input => input.value)
    .filter(value => value && value !== 'ALL');
}

async function onNAreaChange() {
  const area = document.getElementById('nAreaFilter').value;
  if (area === 'ALL') {
    populateNStores(nState.filters.stores);
  } else {
    try {
      const res = await fetch('/api/store-notes-filters?area=' + encodeURIComponent(area));
      const json = await res.json();
      if (json.success) populateNStores(json.stores);
    } catch(e) {}
  }
  document.getElementById('nStoreFilter').value = 'ALL';
  await applyNotesFilters();
}

let _notesSearchTimer = null;
function debouncedNotesSearch() {
  clearTimeout(_notesSearchTimer);
  _notesSearchTimer = setTimeout(applyNotesFilters, 300);
}

async function applyNotesFilters() {
  const area   = document.getElementById('nAreaFilter').value;
  const store  = document.getElementById('nStoreFilter').value;
  const statuses = getNotesStatusValues();
  const q      = document.getElementById('nSearch').value.trim();

  const params = new URLSearchParams();
  if (area !== 'ALL') params.set('area', area);
  if (store !== 'ALL') params.set('store', store);
  statuses.forEach(status => params.append('status', status));
  if (q) params.set('q', q);

  try {
    const res = await fetch('/api/store-notes?' + params);
    const json = await res.json();
    if (!json.success) throw new Error(json.error);
    nState.rows = json.rows || [];
    document.getElementById('nRecordsCount').innerHTML = \`<span>\${nState.rows.length}</span> note\${nState.rows.length!==1?'s':''}\`;
    renderNotesTable(sortRows(nState.rows, nState.sort));
    updateSortHeaders('nTableBody', nState.sort);
    renderNotesAnalytics(nState.rows);
  } catch(e) {
    console.error(e);
    document.getElementById('nTableBody').innerHTML =
      \`<tr><td colspan="6" class="empty-cell"><div class="empty-icon" style="background:linear-gradient(135deg,rgba(244,63,94,0.1),rgba(251,113,133,0.1));color:#fb7185"><i class="fa fa-triangle-exclamation"></i></div><p>\${e.message}</p></td></tr>\`;
  }
}

function statusClass(status) {
  if (!status) return 'default';
  const s = status.toLowerCase().replace(/[\\s_-]/g, '');
  if (['done','completed','closed','resolved','complete','finished'].includes(s)) return 'done';
  if (['pending','open','inprogress','ongoing','progress'].includes(s)) return 'pending';
  if (['urgent','cancelled','canceled','failed','rejected','critical'].includes(s)) return 'urgent';
  if (['review','draft','new','submitted'].includes(s)) return 'review';
  return 'default';
}

function statusIcon(status) {
  const s = (status || '').toLowerCase();
  if (s.includes('done') || s.includes('complete') || s.includes('closed') || s.includes('resolved')) return 'fa-circle-check';
  if (s.includes('pending') || s.includes('progress') || s.includes('ongoing') || s.includes('open')) return 'fa-clock';
  if (s.includes('urgent') || s.includes('cancel') || s.includes('fail') || s.includes('reject')) return 'fa-circle-exclamation';
  if (s.includes('review') || s.includes('draft') || s.includes('new') || s.includes('submit')) return 'fa-eye';
  return 'fa-circle';
}

function formatTimestamp(ts) {
  if (!ts) return '—';
  const d = new Date(ts);
  if (isNaN(d.getTime())) return ts; // return raw if can't parse
  const date = d.toLocaleDateString('en-PH', { month:'short', day:'numeric', year:'numeric' });
  const time = d.toLocaleTimeString('en-PH', { hour:'numeric', minute:'2-digit', hour12:true });
  return \`\${date}<br><span style="color:var(--text-3);font-size:10.5px">\${time}</span>\`;
}

function isValidUrl(s) {
  if (!s) return false;
  return /^https?:\\/\\//i.test(s.trim());
}

// ─── Photo Lightbox ────────────────────────────────────────────────────────
function driveDirectUrl(url) {
  if (!url) return url;
  try {
    let id = null;
    let m = url.match(/\\/file\\/d\\/([a-zA-Z0-9_-]{10,})/);
    if (m) id = m[1];
    if (!id) {
      m = url.match(/[?&]id=([a-zA-Z0-9_-]{10,})/);
      if (m) id = m[1];
    }
    if (id) {
      return 'https://drive.google.com/thumbnail?id=' + id + '&sz=w1600';
    }
  } catch(e) {}
  return url;
}

function openLightbox(url, storeName, timestamp) {
  if (!url) return;
  const lb = document.getElementById('lightbox');
  const img = document.getElementById('lbImage');
  const spinner = document.getElementById('lbSpinner');
  const errEl = document.getElementById('lbError');
  const errLink = document.getElementById('lbErrorLink');
  const titleEl = document.getElementById('lbTitle');
  const meta = document.getElementById('lbMeta');
  const openExt = document.getElementById('lbOpenExternal');

  img.style.display = 'none';
  img.src = '';
  errEl.classList.remove('show');
  spinner.style.display = 'block';

  titleEl.textContent = storeName || 'Photo';
  openExt.href = url;
  errLink.href = url;

  const metaParts = [];
  if (storeName) metaParts.push('<span class="lightbox-meta-item"><i class="fa fa-store"></i> ' + storeName.replace(/</g,'&lt;') + '</span>');
  if (timestamp) {
    const d = new Date(timestamp);
    const tsTxt = !isNaN(d.getTime())
      ? d.toLocaleDateString('en-PH',{month:'short',day:'numeric',year:'numeric',hour:'numeric',minute:'2-digit',hour12:true})
      : timestamp;
    metaParts.push('<span class="lightbox-meta-item"><i class="fa fa-clock"></i> ' + tsTxt + '</span>');
  }
  meta.innerHTML = metaParts.join('');

  const directUrl = driveDirectUrl(url);
  img.onload = () => {
    spinner.style.display = 'none';
    img.style.display = 'block';
  };
  img.onerror = () => {
    spinner.style.display = 'none';
    errEl.classList.add('show');
  };
  img.src = directUrl;

  lb.classList.add('active');
  document.body.style.overflow = 'hidden';
}

function closeLightbox(event) {
  if (event && event.target && event.currentTarget && event.target !== event.currentTarget) return;
  const lb = document.getElementById('lightbox');
  lb.classList.remove('active');
  document.body.style.overflow = '';
  const img = document.getElementById('lbImage');
  img.src = '';
}

document.addEventListener('keydown', e => {
  if (e.key === 'Escape') {
    const lb = document.getElementById('lightbox');
    if (lb && lb.classList.contains('active')) closeLightbox();
  }
});

function renderNotesTable(rows) {
  const tbody = document.getElementById('nTableBody');
  if (!rows.length) {
    tbody.innerHTML = \`<tr><td colspan="6" class="empty-cell"><div class="empty-icon"><i class="fa fa-magnifying-glass"></i></div><p>No notes found</p><small>Try adjusting your filters</small></td></tr>\`;
    return;
  }

  const html = rows.map(r => {
    const areaColor = AREA_COLORS[r.area] || DEFAULT_COLOR;
    const grad = AREA_GRADIENTS[r.area] || [DEFAULT_COLOR, DEFAULT_COLOR];
    const stCls = statusClass(r.status);
    const stIcon = statusIcon(r.status);

    // Prefer the resolved URL from column M (rich text hyperlink), fallback to column G
    const photoUrl = r.photoUrl || (r.photo && isValidUrl(r.photo) ? r.photo : null);
    const safeUrl = photoUrl ? photoUrl.replace(/"/g,'&quot;').replace(/'/g,'&#39;') : '';
    const safeStore = (r.storeName || '').replace(/"/g,'&quot;').replace(/'/g,'&#39;');
    const safeTs = (r.timestamp || '').replace(/"/g,'&quot;').replace(/'/g,'&#39;');
    const photoBtn = photoUrl
      ? \`<button type="button" class="photo-link-btn" onclick="openLightbox('\${safeUrl}','\${safeStore}','\${safeTs}')"><i class="fa fa-image"></i> Open Photo</button>\`
      : \`<span class="photo-link-btn disabled"><i class="fa fa-image-portrait"></i> —</span>\`;

    const notes = (r.notes || '—').replace(/</g,'&lt;').replace(/>/g,'&gt;');
    const remarks = (r.remarks || '—').replace(/</g,'&lt;').replace(/>/g,'&gt;');

    return \`<tr>
      <td><div class="timestamp-cell">\${formatTimestamp(r.timestamp)}</div></td>
      <td>
        <div class="store-cell">
          <div class="store-avatar" style="background:linear-gradient(135deg, \${grad[0]}, \${grad[1]});width:36px;height:36px;font-size:12.5px">\${initials(r.storeName)}</div>
          <div class="store-info">
            <div class="store-name" style="font-size:13.5px">\${r.storeName || '—'}</div>
          </div>
        </div>
      </td>
      <td><div class="notes-cell">\${notes}</div></td>
      <td><span class="status-pill \${stCls}"><i class="fa \${stIcon}"></i> \${r.status || '—'}</span></td>
      <td><div class="remarks-cell">\${remarks}</div></td>
      <td style="text-align:center">\${photoBtn}</td>
    </tr>\`;
  }).join('');

  tbody.innerHTML = html;
}

// ─── Status bucketing ────────────────────────────────────────────────────
function statusBucket(status) {
  const s = (status || '').toLowerCase().trim();
  if (!s) return 'blank';
  if (s.includes('done') || s.includes('complete') || s.includes('closed') || s.includes('resolved') || s.includes('finished')) return 'done';
  if (s.includes('pending') || s === 'open' || s.includes('new') || s.includes('draft')) return 'pending';
  if (s.includes('ongoing') || s.includes('progress') || s.includes('in-progress') || s.includes('inprogress') || s.includes('review')) return 'ongoing';
  // Anything else also goes to ongoing as a "work in flight" catch-all? Better: put unknowns under blank for clarity.
  // We'll treat unknown statuses as "other" → group into blank since user only asked for these 4.
  return 'blank';
}

const STATUS_COLORS = {
  done:    { fill: '#10b981', glow: '#34d399', label: 'Done' },
  pending: { fill: '#f59e0b', glow: '#fbbf24', label: 'Pending' },
  ongoing: { fill: '#6366f1', glow: '#818cf8', label: 'Ongoing' },
  blank:   { fill: '#64748b', glow: '#94a3b8', label: 'No Status' },
};

function calcStatusBreakdown(rows) {
  const counts = { done: 0, pending: 0, ongoing: 0, blank: 0, total: rows.length };
  rows.forEach(r => { counts[statusBucket(r.status)]++; });
  const pct = k => counts.total ? (counts[k] / counts.total * 100) : 0;
  return {
    counts,
    pctDone:    pct('done'),
    pctPending: pct('pending'),
    pctOngoing: pct('ongoing'),
    pctBlank:   pct('blank'),
  };
}

function renderNotesAnalytics(rows) {
  const overall = calcStatusBreakdown(rows);

  // ─── KPI cards ───
  document.getElementById('nKpiTotal').textContent   = overall.counts.total;
  document.getElementById('nKpiDone').textContent    = overall.counts.done;
  document.getElementById('nKpiDoneSub').textContent = overall.pctDone.toFixed(1) + '% of total';
  document.getElementById('nKpiPending').textContent = overall.counts.pending;
  document.getElementById('nKpiPendingSub').textContent = overall.pctPending.toFixed(1) + '% of total';
  document.getElementById('nKpiOngoing').textContent = overall.counts.ongoing;
  document.getElementById('nKpiOngoingSub').textContent = overall.pctOngoing.toFixed(1) + '% of total';
  document.getElementById('nKpiBlank').textContent   = overall.counts.blank;
  document.getElementById('nKpiBlankSub').textContent = overall.pctBlank.toFixed(1) + '% of total';

  // ─── Group by area & store ───
  const groupBy = (rows, keyFn) => {
    const map = {};
    rows.forEach(r => {
      const k = keyFn(r) || '—';
      if (!map[k]) map[k] = [];
      map[k].push(r);
    });
    return map;
  };

  const byArea = Object.entries(groupBy(rows, r => r.area)).map(([area, list]) => {
    const b = calcStatusBreakdown(list);
    return { area, total: list.length, ...b };
  });
  const byStore = Object.entries(groupBy(rows, r => r.storeName)).map(([storeName, list]) => {
    const b = calcStatusBreakdown(list);
    const area = list[0] ? list[0].area : '';
    const storeId = list[0] ? list[0].storeId : '';
    return { storeName, storeId, area, total: list.length, ...b };
  });

  // Rank by "best": highest Done % first; tiebreak by lower Blank %, then by total (more notes = stronger signal)
  const rank = (a, b) =>
    (b.pctDone - a.pctDone) ||
    (a.pctBlank - b.pctBlank) ||
    (b.total - a.total);
  byArea.sort(rank);
  byStore.sort(rank);

  renderAreaStatusChart(byArea);
  renderStatusPie(overall.counts);
  renderStoreStatusChart(byStore);
  renderAreaStatusTable(byArea);
  renderStoreStatusTable(byStore);

  document.getElementById('nAreaTableInfo').textContent = \`\${byArea.length} area\${byArea.length !== 1 ? 's' : ''}\`;
  document.getElementById('nStoreTableInfo').textContent = \`\${byStore.length} store\${byStore.length !== 1 ? 's' : ''}\`;
}

function renderAreaStatusChart(rows) {
  const canvas = document.getElementById('nAreaChart');
  if (!canvas) return;
  if (nState.areaChart) nState.areaChart.destroy();

  const labels = rows.map(r => r.area);
  const datasets = [
    { label: 'Done',      data: rows.map(r => r.pctDone),    backgroundColor: STATUS_COLORS.done.fill,    borderRadius: 4 },
    { label: 'Ongoing',   data: rows.map(r => r.pctOngoing), backgroundColor: STATUS_COLORS.ongoing.fill, borderRadius: 4 },
    { label: 'Pending',   data: rows.map(r => r.pctPending), backgroundColor: STATUS_COLORS.pending.fill, borderRadius: 4 },
    { label: 'No Status', data: rows.map(r => r.pctBlank),   backgroundColor: STATUS_COLORS.blank.fill,   borderRadius: 4 },
  ];

  nState.areaChart = new Chart(canvas, {
    type: 'bar',
    data: { labels, datasets },
    options: {
      indexAxis: 'y', responsive: true, maintainAspectRatio: false,
      layout: { padding: { right: 12 } },
      plugins: {
        legend: { position: 'bottom', labels: { color: '#a8b3d1', font: { size: 11, family: "'Inter'", weight: '500' }, padding: 12, boxWidth: 10, boxHeight: 10, usePointStyle: true, pointStyle: 'rectRounded' } },
        datalabels: {
          display: ctx => ctx.dataset.data[ctx.dataIndex] >= 8,
          color: '#fff',
          font: { family: "'Inter'", size: 10.5, weight: '700' },
          formatter: v => v.toFixed(0) + '%'
        },
        tooltip: {
          backgroundColor: 'rgba(15,20,35,0.95)', borderColor: 'rgba(99,102,241,0.4)', borderWidth: 1, padding: 12, cornerRadius: 10,
          titleFont: { family: "'Space Grotesk'", size: 13, weight: '600' }, bodyFont: { family: "'JetBrains Mono'", size: 11.5 },
          titleColor: '#f0f3fb', bodyColor: '#a8b3d1',
          callbacks: {
            title: ctx => rows[ctx[0].dataIndex].area + '  (Total: ' + rows[ctx[0].dataIndex].total + ')',
            label: ctx => \` \${ctx.dataset.label}: \${ctx.raw.toFixed(1)}%\`
          }
        }
      },
      scales: {
        x: { stacked: true, grid: { color: 'rgba(148,163,200,0.06)' }, ticks: { color: '#a8b3d1', font: { size: 10, family: "'JetBrains Mono'", weight: '500' }, callback: v => v + '%' }, min: 0, max: 100 },
        y: { stacked: true, grid: { color: 'rgba(148,163,200,0.04)', drawBorder: false }, ticks: { color: '#a8b3d1', font: { size: 11, family: "'Inter'", weight: '600' } } }
      }
    }
  });
}

function renderStatusPie(counts) {
  const canvas = document.getElementById('nStatusPie');
  if (!canvas) return;
  if (nState.statusPie) nState.statusPie.destroy();

  const keys = ['done','ongoing','pending','blank'];
  const data = keys.map(k => counts[k]);
  const colors = keys.map(k => STATUS_COLORS[k].fill);
  const labels = keys.map(k => STATUS_COLORS[k].label);

  nState.statusPie = new Chart(canvas, {
    type: 'doughnut',
    data: { labels, datasets: [{ data, backgroundColor: colors, borderColor: 'rgba(10,14,26,0.8)', borderWidth: 3, hoverOffset: 10 }] },
    options: {
      responsive: true, maintainAspectRatio: false, cutout: '60%',
      layout: { padding: 16 },
      plugins: {
        legend: { position: 'bottom', labels: { color: '#a8b3d1', font: { size: 11, family: "'Inter'", weight: '500' }, padding: 14, boxWidth: 10, boxHeight: 10, usePointStyle: true, pointStyle: 'rectRounded' } },
        datalabels: {
          display: ctx => {
            const total = ctx.dataset.data.reduce((a,b)=>a+b,0);
            const val = ctx.dataset.data[ctx.dataIndex];
            return total && (val/total*100) >= 4;
          },
          color: '#fff',
          font: { family: "'Inter'", size: 12, weight: '700' },
          textStrokeColor: 'rgba(0,0,0,0.55)', textStrokeWidth: 3,
          formatter: (val, ctx) => {
            const total = ctx.dataset.data.reduce((a,b)=>a+b,0);
            const pct = total ? (val/total*100) : 0;
            return pct.toFixed(1) + '%';
          }
        },
        tooltip: {
          backgroundColor: 'rgba(15,20,35,0.95)', borderColor: 'rgba(99,102,241,0.4)', borderWidth: 1, padding: 12, cornerRadius: 10,
          titleFont: { family: "'Space Grotesk'", weight: '600', size: 13 }, bodyFont: { family: "'JetBrains Mono'", size: 11.5 },
          titleColor: '#f0f3fb', bodyColor: '#a8b3d1',
          callbacks: {
            label: ctx => {
              const total = ctx.dataset.data.reduce((a,b)=>a+b,0);
              const pct = total ? (ctx.raw/total*100).toFixed(1) : 0;
              return ' ' + ctx.raw + ' notes  (' + pct + '%)';
            }
          }
        }
      }
    }
  });
}

function renderStoreStatusChart(rows) {
  const canvas = document.getElementById('nStoreChart');
  if (!canvas) return;
  if (nState.storeChart) nState.storeChart.destroy();

  // Dynamic height for scrollable view
  const wrap = document.getElementById('nStoreChartWrap');
  const dynHeight = Math.max(320, rows.length * 28 + 40);
  if (wrap) wrap.style.height = dynHeight + 'px';

  const labels = rows.map(r => r.storeName);
  const datasets = [
    { label: 'Done',      data: rows.map(r => r.pctDone),    backgroundColor: STATUS_COLORS.done.fill,    borderRadius: 3 },
    { label: 'Ongoing',   data: rows.map(r => r.pctOngoing), backgroundColor: STATUS_COLORS.ongoing.fill, borderRadius: 3 },
    { label: 'Pending',   data: rows.map(r => r.pctPending), backgroundColor: STATUS_COLORS.pending.fill, borderRadius: 3 },
    { label: 'No Status', data: rows.map(r => r.pctBlank),   backgroundColor: STATUS_COLORS.blank.fill,   borderRadius: 3 },
  ];

  nState.storeChart = new Chart(canvas, {
    type: 'bar',
    data: { labels, datasets },
    options: {
      indexAxis: 'y', responsive: true, maintainAspectRatio: false,
      layout: { padding: { right: 12 } },
      plugins: {
        legend: { position: 'top', labels: { color: '#a8b3d1', font: { size: 11, family: "'Inter'", weight: '500' }, padding: 12, boxWidth: 10, boxHeight: 10, usePointStyle: true, pointStyle: 'rectRounded' } },
        datalabels: {
          display: ctx => ctx.dataset.data[ctx.dataIndex] >= 12,
          color: '#fff',
          font: { family: "'Inter'", size: 10, weight: '700' },
          formatter: v => v.toFixed(0) + '%'
        },
        tooltip: {
          backgroundColor: 'rgba(15,20,35,0.95)', borderColor: 'rgba(99,102,241,0.4)', borderWidth: 1, padding: 12, cornerRadius: 10,
          titleFont: { family: "'Space Grotesk'", size: 13, weight: '600' }, bodyFont: { family: "'JetBrains Mono'", size: 11.5 },
          titleColor: '#f0f3fb', bodyColor: '#a8b3d1',
          callbacks: {
            title: ctx => {
              const r = rows[ctx[0].dataIndex];
              return r.storeName + '  (Total: ' + r.total + ', Area: ' + r.area + ')';
            },
            label: ctx => ' ' + ctx.dataset.label + ': ' + ctx.raw.toFixed(1) + '%'
          }
        }
      },
      scales: {
        x: { stacked: true, grid: { color: 'rgba(148,163,200,0.06)' }, ticks: { color: '#a8b3d1', font: { size: 10, family: "'JetBrains Mono'", weight: '500' }, callback: v => v + '%' }, min: 0, max: 100 },
        y: { stacked: true, grid: { color: 'rgba(148,163,200,0.04)', drawBorder: false }, ticks: { color: '#a8b3d1', font: { size: 10.5, family: "'Inter'", weight: '500' }, autoSkip: false } }
      }
    }
  });
}

function renderAreaStatusTable(rows) {
  const tbody = document.getElementById('nAreaTableBody');
  if (!rows.length) {
    tbody.innerHTML = '<tr><td colspan="7" class="empty-cell"><p>No data</p></td></tr>';
    return;
  }
  tbody.innerHTML = rows.map((r,i) => {
    const areaColor = AREA_COLORS[r.area] || DEFAULT_COLOR;
    const rankBadge = rankBadgeHtml(i);
    return \`<tr>
      <td>\${rankBadge}</td>
      <td><span class="area-tag"><span class="area-dot" style="background:\${areaColor};color:\${areaColor}"></span>\${r.area || '—'}</span></td>
      <td style="text-align:center"><span class="num num-bold" style="color:var(--text-1)">\${r.total}</span></td>
      <td style="text-align:center"><span class="status-pill done">\${r.pctDone.toFixed(1)}%</span></td>
      <td style="text-align:center"><span class="status-pill pending">\${r.pctPending.toFixed(1)}%</span></td>
      <td style="text-align:center"><span class="status-pill review">\${r.pctOngoing.toFixed(1)}%</span></td>
      <td style="text-align:center"><span class="status-pill default">\${r.pctBlank.toFixed(1)}%</span></td>
    </tr>\`;
  }).join('');
}

function renderStoreStatusTable(rows) {
  const tbody = document.getElementById('nStoreTableBody');
  if (!rows.length) {
    tbody.innerHTML = '<tr><td colspan="7" class="empty-cell"><p>No data</p></td></tr>';
    return;
  }
  tbody.innerHTML = rows.map((r,i) => {
    const grad = AREA_GRADIENTS[r.area] || [DEFAULT_COLOR, DEFAULT_COLOR];
    const rankBadge = rankBadgeHtml(i);
    return \`<tr>
      <td>\${rankBadge}</td>
      <td>
        <div class="store-cell">
          <div class="store-avatar" style="background:linear-gradient(135deg, \${grad[0]}, \${grad[1]});width:32px;height:32px;font-size:11px">\${initials(r.storeName)}</div>
          <div class="store-info">
            <div class="store-name" style="font-size:13px">\${r.storeName || '—'}</div>
            <div class="store-id">\${r.storeId ? '#'+r.storeId : ''}</div>
          </div>
        </div>
      </td>
      <td style="text-align:center"><span class="num num-bold" style="color:var(--text-1)">\${r.total}</span></td>
      <td style="text-align:center"><span class="status-pill done">\${r.pctDone.toFixed(1)}%</span></td>
      <td style="text-align:center"><span class="status-pill pending">\${r.pctPending.toFixed(1)}%</span></td>
      <td style="text-align:center"><span class="status-pill review">\${r.pctOngoing.toFixed(1)}%</span></td>
      <td style="text-align:center"><span class="status-pill default">\${r.pctBlank.toFixed(1)}%</span></td>
    </tr>\`;
  }).join('');
}

function rankBadgeHtml(idx) {
  const rank = idx + 1;
  if (rank === 1) return '<span style="display:inline-flex;align-items:center;justify-content:center;width:26px;height:26px;border-radius:8px;background:linear-gradient(135deg,#fbbf24,#f59e0b);color:#1a0e00;font-weight:800;font-size:12px;font-family:\\'Space Grotesk\\'">1</span>';
  if (rank === 2) return '<span style="display:inline-flex;align-items:center;justify-content:center;width:26px;height:26px;border-radius:8px;background:linear-gradient(135deg,#cbd5e1,#94a3b8);color:#0a0e1a;font-weight:800;font-size:12px;font-family:\\'Space Grotesk\\'">2</span>';
  if (rank === 3) return '<span style="display:inline-flex;align-items:center;justify-content:center;width:26px;height:26px;border-radius:8px;background:linear-gradient(135deg,#d97706,#92400e);color:#fff;font-weight:800;font-size:12px;font-family:\\'Space Grotesk\\'">3</span>';
  return '<span style="display:inline-flex;align-items:center;justify-content:center;width:26px;height:26px;border-radius:8px;background:rgba(99,102,241,0.1);border:1px solid rgba(99,102,241,0.25);color:var(--indigo2);font-weight:700;font-size:11.5px;font-family:\\'JetBrains Mono\\'">' + rank + '</span>';
}

function sortNotes(key, type) {
  if (nState.sort.key === key) nState.sort.dir = nState.sort.dir === 'asc' ? 'desc' : 'asc';
  else nState.sort = { key, dir: 'desc', type };
  nState.sort.type = type;
  renderNotesTable(sortRows(nState.rows, nState.sort));
  updateSortHeaders('nTableBody', nState.sort);
}

//  ISSUES & CONCERNS TAB 
const iState = {
  initialized: false,
  rows: [],
  filters: { areas: [], stores: [], priorities: [], statuses: [], categories: [] },
  sort: { key: 'dateTs', dir: 'desc', type: 'num' },
  priorityChart: null,
  statusChart: null,
  categoryChart: null,
  areaChart: null,
  storeChart: null,
};

const PRIORITY_COLORS = {
  critical: '#dc2626', high: '#f43f5e', medium: '#f59e0b', low: '#10b981',
  default: '#94a3b8',
};
const IMPACT_COLORS = {
  high: '#f43f5e', medium: '#f59e0b', low: '#10b981',
  default: '#94a3b8',
};

function priorityKey(p) {
  const s = (p || '').toLowerCase().trim();
  if (!s) return 'default';
  if (s.includes('critic')) return 'critical';
  if (s.includes('high') || s === '1' || s === 'p1') return 'high';
  if (s.includes('med') || s === '2' || s === 'p2') return 'medium';
  if (s.includes('low') || s === '3' || s === 'p3') return 'low';
  return 'default';
}

function priorityColor(p) {
  return PRIORITY_COLORS[priorityKey(p)] || PRIORITY_COLORS.default;
}

function impactKey(i) {
  const s = (i || '').toLowerCase().trim();
  if (!s) return 'default';
  if (s.includes('high')) return 'high';
  if (s.includes('med')) return 'medium';
  if (s.includes('low')) return 'low';
  return 'default';
}

function impactColor(i) {
  return IMPACT_COLORS[impactKey(i)] || IMPACT_COLORS.default;
}

function renderCheckboxFilter(id, allLabel, values) {
  const wrap = document.getElementById(id);
  if (!wrap) return;
  const name = id + 'Option';
  wrap.innerHTML = '';

  const all = document.createElement('label');
  all.innerHTML = '<input type="checkbox" name="' + escHtml(name) + '" value="ALL" checked> ' + escHtml(allLabel);
  wrap.appendChild(all);

  values.forEach(value => {
    const label = document.createElement('label');
    label.innerHTML = '<input type="checkbox" name="' + escHtml(name) + '" value="' + escHtml(value) + '"> ' + escHtml(value);
    wrap.appendChild(label);
  });

  wrap.querySelectorAll('input').forEach(input => {
    input.addEventListener('change', e => onCheckboxFilterChange(id, e));
  });
}

function onCheckboxFilterChange(id, e) {
  const wrap = document.getElementById(id);
  if (!wrap) return;
  const all = wrap.querySelector('input[value="ALL"]');
  const items = Array.from(wrap.querySelectorAll('input:not([value="ALL"])'));
  const changed = e && e.target;

  if (changed === all && all.checked) {
    items.forEach(input => { input.checked = false; });
  } else {
    const anyChecked = items.some(input => input.checked);
    if (all) all.checked = !anyChecked;
  }

  applyIssuesFilters();
}

function getCheckedFilterValues(id) {
  const wrap = document.getElementById(id);
  if (!wrap) return [];
  return Array.from(wrap.querySelectorAll('input:checked'))
    .map(input => input.value)
    .filter(value => value && value !== 'ALL');
}

function addMultiFilterParams(params, name, values) {
  values.forEach(v => params.append(name, v));
}

function describeMulti(values, allLabel) {
  if (!values.length) return null;
  return values.length === 1 ? values[0] : values.length + ' ' + allLabel;
}

function issueStatusClass(s) {
  const t = (s || '').toLowerCase().trim();
  if (!t) return 'default';
  if (/resolv|closed|done|complete/.test(t)) return 'done';
  if (/pending|open|new/.test(t)) return 'pending';
  if (/progress|ongoing|review/.test(t)) return 'review';
  if (/cancel|reject|urgent|critical/.test(t)) return 'urgent';
  return 'default';
}

async function initIssuesTab() {
  try {
    const res = await fetch('/api/issues-filters');
    const json = await res.json();
    if (!json.success) throw new Error(json.error);
    iState.filters = json;

    const aSel = document.getElementById('iAreaFilter');
    aSel.innerHTML = '<option value="ALL">All Areas</option>';
    json.areas.forEach(a => {
      const o = document.createElement('option');
      o.value = o.textContent = a;
      aSel.appendChild(o);
    });

    populateIStores(json.stores);

    // Sort priorities by severity (critical, high, medium, low, others)
    const priOrder = { critical: 1, high: 2, medium: 3, low: 4, default: 5 };
    const priorities = [...json.priorities].sort((a,b) => (priOrder[priorityKey(a)]||9) - (priOrder[priorityKey(b)]||9));
    renderCheckboxFilter('iPriorityFilter', 'All Priorities', priorities);
    renderCheckboxFilter('iStatusFilter', 'All Statuses', json.statuses);

    iState.initialized = true;
    await applyIssuesFilters();
  } catch(e) {
    console.error('Issues init error:', e);
    document.getElementById('iTableBody').innerHTML =
      \`<tr><td colspan="18" class="empty-cell"><div class="empty-icon" style="background:linear-gradient(135deg,rgba(244,63,94,0.1),rgba(251,113,133,0.1));color:#fb7185"><i class="fa fa-triangle-exclamation"></i></div><p>\${e.message}</p></td></tr>\`;
  }
}

function populateIStores(stores) {
  const sel = document.getElementById('iStoreFilter');
  sel.innerHTML = '<option value="ALL">All Stores</option>';
  stores.forEach(s => {
    const o = document.createElement('option');
    o.value = o.textContent = s;
    sel.appendChild(o);
  });
}

async function onIAreaChange() {
  const area = document.getElementById('iAreaFilter').value;
  if (area === 'ALL') {
    populateIStores(iState.filters.stores);
  } else {
    try {
      const res = await fetch('/api/issues-filters?area=' + encodeURIComponent(area));
      const json = await res.json();
      if (json.success) populateIStores(json.stores);
    } catch(e) {}
  }
  document.getElementById('iStoreFilter').value = 'ALL';
  await applyIssuesFilters();
}

let _issuesSearchTimer = null;
function debouncedIssuesSearch() {
  clearTimeout(_issuesSearchTimer);
  _issuesSearchTimer = setTimeout(applyIssuesFilters, 300);
}

async function applyIssuesFilters() {
  const area     = document.getElementById('iAreaFilter').value;
  const store    = document.getElementById('iStoreFilter').value;
  const priorities = getCheckedFilterValues('iPriorityFilter');
  const statuses   = getCheckedFilterValues('iStatusFilter');
  const q        = document.getElementById('iSearch').value.trim();

  const params = new URLSearchParams();
  if (area !== 'ALL') params.set('area', area);
  if (store !== 'ALL') params.set('store', store);
  addMultiFilterParams(params, 'priority', priorities);
  addMultiFilterParams(params, 'status', statuses);
  if (q) params.set('q', q);

  try {
    const res = await fetch('/api/issues?' + params);
    const json = await res.json();
    if (!json.success) throw new Error(json.error);
    iState.rows = json.rows || [];
    document.getElementById('iRecordsCount').innerHTML = \`<span>\${iState.rows.length}</span> issue\${iState.rows.length!==1?'s':''}\`;

    const parts = [];
    if (area !== 'ALL') parts.push(area);
    if (store !== 'ALL') parts.push(store);
    const priorityDesc = describeMulti(priorities, 'priorities');
    const statusDesc = describeMulti(statuses, 'statuses');
    if (priorityDesc) parts.push(priorityDesc);
    if (statusDesc) parts.push(statusDesc);
    document.getElementById('iTableInfo').textContent = parts.length ? parts.join(' · ') : 'All data';

    renderIssuesKPIs(iState.rows);
    renderHighPriorityNotDoneTable(iState.rows);
    renderIssuesTable(sortRows(iState.rows, iState.sort));
    updateSortHeaders('iTableBody', iState.sort);
    renderIssuesCharts(iState.rows);
  } catch(e) {
    console.error(e);
    document.getElementById('iTableBody').innerHTML =
      \`<tr><td colspan="18" class="empty-cell"><div class="empty-icon" style="background:linear-gradient(135deg,rgba(244,63,94,0.1),rgba(251,113,133,0.1));color:#fb7185"><i class="fa fa-triangle-exclamation"></i></div><p>\${e.message}</p></td></tr>\`;
  }
}

function renderIssuesKPIs(rows) {
  const total = rows.length;
  const resolved = rows.filter(r => r.isResolved).length;
  const open = total - resolved;
  const overdue = rows.filter(r => r.overdue).length;
  const openRows = rows.filter(r => !r.isResolved);
  const avgDaysOpen = openRows.length ? openRows.reduce((s,r)=>s+(r.daysOpen||0),0) / openRows.length : 0;

  document.getElementById('i-kpi-total').textContent    = total;
  document.getElementById('i-kpi-open').textContent     = open;
  document.getElementById('i-kpi-open-sub').textContent = total ? ((open/total*100).toFixed(1) + '% of total') : '0% of total';
  document.getElementById('i-kpi-resolved').textContent = resolved;
  document.getElementById('i-kpi-resolved-sub').textContent = total ? ((resolved/total*100).toFixed(1) + '% resolution rate') : '0% resolution rate';
  document.getElementById('i-kpi-overdue').textContent  = overdue;
  document.getElementById('i-kpi-avg-days').textContent = avgDaysOpen.toFixed(1);
}

function formatShortDate(ts) {
  if (!ts) return '—';
  const d = new Date(ts);
  if (isNaN(d.getTime())) return '—';
  return d.toLocaleDateString('en-PH', { month: 'short', day: 'numeric', year: 'numeric' });
}

function escHtml(s) {
  return (s == null ? '' : String(s)).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function renderHighPriorityNotDoneTable(rows) {
  const tbody = document.getElementById('iHighNotDoneBody');
  const info = document.getElementById('iHighNotDoneInfo');
  if (!tbody) return;

  const focused = highPriorityNotDoneRows(rows);

  if (info) info.textContent = focused.length + ' issue' + (focused.length !== 1 ? 's' : '');

  if (!focused.length) {
    tbody.innerHTML = '<tr><td colspan="7" class="empty-cell"><div class="empty-icon"><i class="fa fa-circle-check"></i></div><p>No high priority not done issues</p><small>Current filters have no matching records</small></td></tr>';
    return;
  }

  tbody.innerHTML = focused.map(r => {
    const iColor = impactColor(r.impactLevel);
    const stCls = issueStatusClass(r.status);
    return '<tr>' +
      '<td><span class="timestamp-cell">' + (escHtml(r.date) || '—') + '</span></td>' +
      '<td><div class="store-cell"><div class="store-info">' +
        '<div class="store-name" style="font-size:13.5px">' + (escHtml(r.storeName) || '—') + '</div>' +
        '<div class="store-id">' + escHtml(r.storeId ? '#'+r.storeId : r.area || '') + '</div>' +
      '</div></div></td>' +
      '<td><div class="notes-cell">' + (escHtml(r.issueDescription) || '—') + '</div></td>' +
      '<td style="text-align:left"><span class="status-pill" style="background:' + iColor + '22;border:1px solid ' + iColor + '66;color:' + iColor + '">' + (escHtml(r.impactLevel) || '—') + '</span></td>' +
      '<td style="text-align:center"><span class="status-pill ' + stCls + '">' + (escHtml(r.status) || '—') + '</span></td>' +
      '<td><div class="remarks-cell">' + (escHtml(r.remarks) || '—') + '</div></td>' +
      '<td><span class="timestamp-cell">' + (escHtml(r.lastUpdate) || '—') + '</span></td>' +
    '</tr>';
  }).join('');
}

function highPriorityNotDoneRows(rows) {
  return rows
    .filter(r => priorityKey(r.priority) === 'high' && !r.isResolved)
    .sort((a, b) => (b.dateTs || 0) - (a.dateTs || 0) || (a.storeName || '').localeCompare(b.storeName || ''));
}

function issueExportDateTag() {
  const d = new Date();
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  return yyyy + '-' + mm + '-' + dd;
}

function writeIssueExcelFile(rows, sheetName, filename, columns) {
  if (!rows.length) {
    alert('No data to export with the current filters.');
    return;
  }
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(rows);
  ws['!cols'] = columns.map(wch => ({ wch }));

  if (ws['!ref']) {
    const range = XLSX.utils.decode_range(ws['!ref']);
    const headerStyle = {
      fill: { fgColor: { rgb: '166534' } },
      font: { color: { rgb: 'FFFFFF' }, bold: true },
      alignment: { horizontal: 'center', vertical: 'center' }
    };
    for (let C = range.s.c; C <= range.e.c; C++) {
      const addr = XLSX.utils.encode_cell({ r: 0, c: C });
      if (ws[addr]) ws[addr].s = headerStyle;
    }
  }

  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  XLSX.writeFile(wb, filename);
}

function exportHighPriorityNotDoneToExcel() {
  const rows = highPriorityNotDoneRows(iState.rows).map(r => ({
    'Date': r.date || '',
    'Store Name': r.storeName || '',
    'Issue Description': r.issueDescription || '',
    'Impact': r.impactLevel || '',
    'Status': r.status || '',
    'Remarks': r.remarks || '',
    'Last Update': r.lastUpdate || ''
  }));
  writeIssueExcelFile(
    rows,
    'High Priority Not Done',
    'CaMaNaVa_High_Priority_Not_Done_' + issueExportDateTag() + '.xlsx',
    [14, 24, 60, 14, 18, 40, 22]
  );
}

function exportIssuesDetailToExcel() {
  const rows = sortRows(iState.rows, iState.sort).map(r => ({
    'Area': r.area || '',
    'Store ID': r.storeId || '',
    'Store Name': r.storeName || '',
    'Date': r.date || '',
    'Reported By': r.reportedBy || '',
    'Issue Category': r.issueCategory || '',
    'Sub Category': r.issueSubCategory || '',
    'Issue Description': r.issueDescription || '',
    'Priority': r.priority || '',
    'Impact': r.impactLevel || '',
    'Assign To': r.assignTo || '',
    'Target Date': r.targetDate || '',
    'Status': r.status || '',
    'Resolution Details': r.resolutionDetails || '',
    'Date Resolved': r.dateResolved || '',
    'Days Open': r.daysOpen || 0,
    'Remarks / Notes': r.remarks || '',
    'Last Update': r.lastUpdate || ''
  }));
  writeIssueExcelFile(
    rows,
    'Issues Detail',
    'CaMaNaVa_Issues_Detail_' + issueExportDateTag() + '.xlsx',
    [16, 10, 24, 14, 18, 22, 22, 60, 14, 14, 18, 16, 18, 44, 16, 10, 40, 22]
  );
}

function renderIssuesTable(rows) {
  const tbody = document.getElementById('iTableBody');
  if (!rows.length) {
    tbody.innerHTML = \`<tr><td colspan="18" class="empty-cell"><div class="empty-icon"><i class="fa fa-magnifying-glass"></i></div><p>No issues found</p><small>Try adjusting your filters</small></td></tr>\`;
    return;
  }

  const html = rows.map(r => {
    const areaColor = AREA_COLORS[r.area] || DEFAULT_COLOR;
    const grad = AREA_GRADIENTS[r.area] || [DEFAULT_COLOR, DEFAULT_COLOR];
    const pColor = priorityColor(r.priority);
    const iColor = impactColor(r.impactLevel);
    const stCls = issueStatusClass(r.status);

    const overdueBadge = r.overdue ? ' <span style="color:#fb7185;font-weight:700" title="Overdue"><i class="fa fa-triangle-exclamation"></i></span>' : '';

    return \`<tr>
      <td><span class="area-tag"><span class="area-dot" style="background:\${areaColor};color:\${areaColor}"></span>\${escHtml(r.area)||'—'}</span></td>
      <td><span class="num num-bold" style="color:var(--text-1)">\${escHtml(r.storeId)||'—'}</span></td>
      <td>
        <div class="store-cell">
          <div class="store-avatar" style="background:linear-gradient(135deg, \${grad[0]}, \${grad[1]});width:34px;height:34px;font-size:12px">\${initials(r.storeName)}</div>
          <div class="store-info"><div class="store-name" style="font-size:13.5px">\${escHtml(r.storeName)||'—'}</div></div>
        </div>
      </td>
      <td><span class="timestamp-cell">\${escHtml(r.date)||'—'}</span></td>
      <td><span style="font-size:13px;color:var(--text-2)">\${escHtml(r.reportedBy)||'—'}</span></td>
      <td><span style="font-size:13px;color:var(--text-1);font-weight:500">\${escHtml(r.issueCategory)||'—'}</span></td>
      <td><span style="font-size:12.5px;color:var(--text-2)">\${escHtml(r.issueSubCategory)||'—'}</span></td>
      <td><div class="notes-cell">\${escHtml(r.issueDescription)||'—'}</div></td>
      <td style="text-align:center"><span class="status-pill" style="background:\${pColor}22;border:1px solid \${pColor}66;color:\${pColor}"><i class="fa fa-flag"></i> \${escHtml(r.priority)||'—'}</span></td>
      <td style="text-align:center"><span class="status-pill" style="background:\${iColor}22;border:1px solid \${iColor}66;color:\${iColor}">\${escHtml(r.impactLevel)||'—'}</span></td>
      <td><span style="font-size:13px;color:var(--text-2)">\${escHtml(r.assignTo)||'—'}</span></td>
      <td><span class="timestamp-cell" style="\${r.overdue?'color:#fb7185;font-weight:600':''}">\${escHtml(r.targetDate)||'—'}\${overdueBadge}</span></td>
      <td style="text-align:center"><span class="status-pill \${stCls}">\${escHtml(r.status)||'—'}</span></td>
      <td><div class="notes-cell">\${escHtml(r.resolutionDetails)||'—'}</div></td>
      <td><span class="timestamp-cell">\${escHtml(r.dateResolved)||'—'}</span></td>
      <td style="text-align:center"><span class="num num-bold" style="color:\${r.daysOpen>14?'#fb7185':r.daysOpen>7?'#fbbf24':'var(--text-1)'}">\${r.daysOpen}</span></td>
      <td><div class="remarks-cell">\${escHtml(r.remarks)||'—'}</div></td>
      <td><span class="timestamp-cell">\${escHtml(r.lastUpdate)||'—'}</span></td>
    </tr>\`;
  }).join('');

  tbody.innerHTML = html;
}

function sortIssues(key, type) {
  if (iState.sort.key === key) iState.sort.dir = iState.sort.dir === 'asc' ? 'desc' : 'asc';
  else iState.sort = { key, dir: 'desc', type };
  iState.sort.type = type;
  renderIssuesTable(sortRows(iState.rows, iState.sort));
  updateSortHeaders('iTableBody', iState.sort);
}

// ─── Issues Charts ─────────────────────────────────────────────────────────
function renderIssuesCharts(rows) {
  // Priority doughnut
  const priorityMap = {};
  rows.forEach(r => {
    const k = r.priority || 'No Priority';
    priorityMap[k] = (priorityMap[k] || 0) + 1;
  });
  const priorityLabels = Object.keys(priorityMap);
  const priOrder = { critical: 1, high: 2, medium: 3, low: 4, default: 5 };
  priorityLabels.sort((a,b) => (priOrder[priorityKey(a)]||9) - (priOrder[priorityKey(b)]||9));
  const priorityValues = priorityLabels.map(k => priorityMap[k]);
  const priorityCols = priorityLabels.map(k => priorityColor(k));

  if (iState.priorityChart) iState.priorityChart.destroy();
  iState.priorityChart = makeDoughnut('iPriorityChart', priorityLabels, priorityValues, priorityCols);

  // Status doughnut
  const statusMap = {};
  rows.forEach(r => {
    const k = r.status || 'No Status';
    statusMap[k] = (statusMap[k] || 0) + 1;
  });
  const statusLabels = Object.keys(statusMap);
  const statusValues = statusLabels.map(k => statusMap[k]);
  const statusCols = statusLabels.map(k => {
    const t = k.toLowerCase();
    if (/resolv|closed|done|complete/.test(t)) return '#10b981';
    if (/progress|ongoing|review/.test(t)) return '#6366f1';
    if (/pending|open|new/.test(t)) return '#f59e0b';
    if (/cancel|reject/.test(t)) return '#f43f5e';
    return '#94a3b8';
  });
  if (iState.statusChart) iState.statusChart.destroy();
  iState.statusChart = makeDoughnut('iStatusChart', statusLabels, statusValues, statusCols);

  // Category horizontal bar (sorted by count desc)
  const catMap = {};
  rows.forEach(r => {
    const k = r.issueCategory || 'No Category';
    catMap[k] = (catMap[k] || 0) + 1;
  });
  const catEntries = Object.entries(catMap).sort((a,b) => b[1] - a[1]).slice(0, IS_MOBILE ? 8 : 14);
  if (iState.categoryChart) iState.categoryChart.destroy();
  iState.categoryChart = makeHBar('iCategoryChart', catEntries.map(e=>e[0]), catEntries.map(e=>e[1]), '#a855f7');

  // Area horizontal bar
  const areaMap = {};
  rows.forEach(r => {
    const k = r.area || 'No Area';
    areaMap[k] = (areaMap[k] || 0) + 1;
  });
  const areaEntries = Object.entries(areaMap).sort((a,b) => b[1] - a[1]);
  const areaCols = areaEntries.map(e => AREA_COLORS[e[0]] || DEFAULT_COLOR);
  if (iState.areaChart) iState.areaChart.destroy();
  iState.areaChart = makeHBar('iAreaChart', areaEntries.map(e=>e[0]), areaEntries.map(e=>e[1]), areaCols);

  // Top affected stores (by OPEN issues)
  const storeMap = {};
  rows.forEach(r => {
    if (r.isResolved) return; // count only open
    const k = r.storeName || 'No Store';
    if (!storeMap[k]) storeMap[k] = { name: k, area: r.area, count: 0 };
    storeMap[k].count++;
  });
  const storeEntries = Object.values(storeMap).sort((a,b) => b.count - a.count);
  const wrap = document.getElementById('iStoreChartWrap');
  const dynHeight = Math.max(320, storeEntries.length * 28 + 40);
  if (wrap) wrap.style.height = dynHeight + 'px';
  const storeCols = storeEntries.map(s => AREA_COLORS[s.area] || DEFAULT_COLOR);
  if (iState.storeChart) iState.storeChart.destroy();
  iState.storeChart = makeHBar('iStoreChart', storeEntries.map(s => s.name), storeEntries.map(s => s.count), storeCols);
}

function makeDoughnut(canvasId, labels, data, colors) {
  const canvas = document.getElementById(canvasId);
  if (!canvas) return null;
  return new Chart(canvas, {
    type: 'doughnut',
    data: { labels, datasets: [{ data, backgroundColor: colors, borderColor: 'rgba(10,14,26,0.8)', borderWidth: 3, hoverOffset: 10 }] },
    options: {
      responsive: true, maintainAspectRatio: false, cutout: '58%',
      layout: { padding: 16 },
      plugins: {
        legend: { position: 'bottom', labels: { color: '#a8b3d1', font: { size: 11, family: "'Inter'", weight: '500' }, padding: 12, boxWidth: 10, boxHeight: 10, usePointStyle: true, pointStyle: 'rectRounded',
          generateLabels: chart => chart.data.labels.map((label, i) => ({
            text: label + '  (' + chart.data.datasets[0].data[i] + ')',
            fillStyle: colors[i], strokeStyle: 'transparent', lineWidth: 0, pointStyle: 'rectRounded',
          })),
        } },
        datalabels: {
          display: ctx => {
            const total = ctx.dataset.data.reduce((a,b)=>a+b,0);
            return total && (ctx.dataset.data[ctx.dataIndex]/total*100) >= 4;
          },
          color: '#fff', font: { family: "'Inter'", size: 12, weight: '700' },
          textStrokeColor: 'rgba(0,0,0,0.55)', textStrokeWidth: 3,
          formatter: (val, ctx) => {
            const total = ctx.dataset.data.reduce((a,b)=>a+b,0);
            return total ? (val/total*100).toFixed(1) + '%' : '0%';
          },
        },
        tooltip: {
          backgroundColor: 'rgba(15,20,35,0.95)', borderColor: 'rgba(99,102,241,0.4)', borderWidth: 1, padding: 12, cornerRadius: 10,
          titleFont: { family: "'Space Grotesk'", weight: '600', size: 13 }, bodyFont: { family: "'JetBrains Mono'", size: 11.5 },
          titleColor: '#f0f3fb', bodyColor: '#a8b3d1',
          callbacks: { label: ctx => {
            const total = ctx.dataset.data.reduce((a,b)=>a+b,0);
            const pct = total ? (ctx.raw/total*100).toFixed(1) : 0;
            return ' ' + ctx.raw + ' issues  (' + pct + '%)';
          } }
        }
      }
    }
  });
}

function makeHBar(canvasId, labels, data, colors) {
  const canvas = document.getElementById(canvasId);
  if (!canvas) return null;
  const bg = Array.isArray(colors) ? colors : labels.map(_ => colors);
  return new Chart(canvas, {
    type: 'bar',
    data: { labels, datasets: [{ label: 'Issues', data, backgroundColor: bg, borderRadius: 5, borderSkipped: false, barPercentage: 0.78 }] },
    options: {
      indexAxis: 'y', responsive: true, maintainAspectRatio: false,
      layout: { padding: { right: 40 } },
      plugins: {
        legend: { display: false },
        datalabels: {
          display: true, anchor: 'end', align: 'end',
          color: '#e8ecf4', font: { family: "'JetBrains Mono'", size: 11, weight: '700' },
          padding: { left: 6 },
          formatter: v => v,
        },
        tooltip: {
          backgroundColor: 'rgba(15,20,35,0.95)', borderColor: 'rgba(99,102,241,0.4)', borderWidth: 1, padding: 12, cornerRadius: 10,
          titleFont: { family: "'Space Grotesk'", size: 13, weight: '600' }, bodyFont: { family: "'JetBrains Mono'", size: 11.5 },
          titleColor: '#f0f3fb', bodyColor: '#a8b3d1',
          callbacks: { label: ctx => ' ' + ctx.raw + ' issue' + (ctx.raw !== 1 ? 's' : '') }
        }
      },
      scales: {
        x: { grid: { color: 'rgba(148,163,200,0.06)' }, ticks: { color: '#a8b3d1', font: { size: 10.5, family: "'JetBrains Mono'", weight: '500' }, precision: 0 } },
        y: { grid: { color: 'rgba(148,163,200,0.04)', drawBorder: false }, ticks: { color: '#a8b3d1', font: { size: 11, family: "'Inter'", weight: '500' }, autoSkip: false } }
      }
    }
  });
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
  document.getElementById('tableBody').innerHTML = \`<tr><td colspan="6" class="empty-cell"><div class="empty-icon" style="background:linear-gradient(135deg,rgba(244,63,94,0.1),rgba(251,113,133,0.1));color:#fb7185"><i class="fa fa-triangle-exclamation"></i></div><p>\${msg}</p></td></tr>\`;
}

updateThemeToggle();
loadFilters();
document.addEventListener('click', e => {
  const dd = document.getElementById('mMonthDropdown');
  if (dd && !dd.contains(e.target)) {
    dd.classList.remove('open');
    const controls = dd.closest('.controls');
    if (controls) controls.classList.remove('dropdown-open');
  }
});
</script>
</body>
</html>
`;

const LOGIN_HTML = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=5"/>
<meta name="theme-color" content="#0a0e1a"/>
<title>CaMaNaVa eBRT Login</title>
<link rel="preconnect" href="https://fonts.googleapis.com"/>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600&display=swap"/>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css"/>
<style>
*{box-sizing:border-box;margin:0;padding:0}
:root{--bg:#0a0e1a;--panel:rgba(20,26,42,.86);--border:rgba(148,163,200,.18);--text:#f0f3fb;--muted:#a8b3d1;--dim:#6b7693;--accent:#10b981;--cyan:#06b6d4;--rose:#fb7185}
html,body{min-height:100vh;background:var(--bg);color:var(--text);font-family:Inter,system-ui,sans-serif}
body{display:grid;place-items:center;padding:22px;background:
  radial-gradient(at 16% 8%,rgba(6,182,212,.18),transparent 42%),
  radial-gradient(at 86% 82%,rgba(16,185,129,.15),transparent 45%),
  linear-gradient(135deg,#070a14,#0a0e1a 55%,#0b1d26)}
.card{width:min(430px,100%);background:var(--panel);border:1px solid var(--border);border-radius:18px;padding:30px;box-shadow:0 24px 80px rgba(0,0,0,.38);backdrop-filter:blur(18px)}
.brand{display:flex;align-items:center;gap:13px;margin-bottom:24px}
.logo{width:46px;height:46px;border-radius:13px;background:linear-gradient(135deg,#06b6d4,#10b981);display:grid;place-items:center;font-family:"Space Grotesk";font-weight:700;box-shadow:0 12px 26px -12px rgba(6,182,212,.8)}
h1{font-family:"Space Grotesk";font-size:23px;letter-spacing:-.02em;line-height:1.1}
.sub{font-size:12px;color:var(--dim);text-transform:uppercase;letter-spacing:.14em;margin-top:3px}
p{color:var(--muted);font-size:14px;line-height:1.45;margin-bottom:22px}
label{display:block;margin:15px 0 7px;color:var(--muted);font-size:11px;text-transform:uppercase;letter-spacing:.12em;font-weight:700}
input{width:100%;border:1px solid rgba(148,163,200,.22);background:rgba(10,14,26,.72);color:var(--text);border-radius:11px;padding:13px 14px;font-size:15px;outline:none}
input:focus{border-color:var(--cyan);box-shadow:0 0 0 3px rgba(6,182,212,.18)}
button{width:100%;margin-top:22px;border:0;border-radius:11px;padding:13px 14px;background:linear-gradient(135deg,var(--cyan),var(--accent));color:white;font-weight:800;font-size:14px;cursor:pointer}
.error{display:none;margin-top:15px;color:var(--rose);font-size:13px;line-height:1.45}
.hint{margin-top:16px;color:var(--dim);font-size:12px;line-height:1.45}
</style>
</head>
<body>
  <form class="card" method="post" action="/login">
    <div class="brand">
      <div class="logo">C</div>
      <div>
        <h1>CaMaNaVa eBRT</h1>
        <div class="sub">Daily Sales Report</div>
      </div>
    </div>
    <label for="username">Username</label>
    <input id="username" name="username" autocomplete="username" required autofocus/>
    <label for="password">Password</label>
    <input id="password" name="password" type="password" autocomplete="current-password" required/>
    <button type="submit"><i class="fa fa-right-to-bracket"></i> Log in</button>
    <div class="error" id="error">Invalid username or password.</div>
  </form>
  <script>
    const err = new URLSearchParams(location.search).get('error');
    if (err) {
      const messages = {
        no_user: 'Username was not found in the Google Sheet user tab.',
        bad_password: 'Password did not match this username.',
        config: 'Login could not read the Google Sheet user tab. Check sheet access and columns B:E.',
      };
      const el = document.getElementById('error');
      el.textContent = messages[err] || 'Invalid username or password.';
      el.style.display = 'block';
    }
  </script>
</body>
</html>`;

app.get('/', (req, res) => {
  res.setHeader('Content-Type', 'text/html; charset=utf-8');
  res.send(getRequestUser(req) ? HTML : LOGIN_HTML);
});
app.get('*', (req, res) => {
  res.setHeader('Content-Type', 'text/html; charset=utf-8');
  res.send(getRequestUser(req) ? HTML : LOGIN_HTML);
});

app.listen(PORT, () => console.log('CaMaNaVa eBRT running on port ' + PORT));
