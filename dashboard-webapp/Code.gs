const SPREADSHEET_ID = "1QGFrQHb7c6NAZTlhxMWClPox1TcQUrwVMeIlI2UMY8I"; // Replace with your Spreadsheet ID


// ─── doGet ────────────────────────────────────────────────────────────────────
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Login Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ─── Login ────────────────────────────────────────────────────────────────────
function doLogin(username, password) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hrAdminSheet = ss.getSheetByName('HRAdmin');
    const data = hrAdminSheet.getDataRange().getValues();

    if (data.length > 1) {
      const row = data[1];
      const sheetUsername = row[0];
      const sheetPassword = row[1];
      const adminName = row[2];

      if (username === sheetUsername && password === sheetPassword) {
        return {
          success: true,
          adminName: adminName,
          lastTimestamp: 0
        };
      }
    }

    return { success: false };
  } catch (e) {
    Logger.log('Login error: ' + e);
    return { success: false, error: e.toString() };
  }
}

// ─── Column indices (0-based, no header in DataCache) ────────────────────────
// A=0  Response
// B=1  Time (GMT)
// C=2  Governorate
// D=3  Center
// H=7  Area of Malfunction
// I=8  Malfunction
// L=11 Supplier
// T=19 Company Name
// U=20 TimeStamp Admin Sent
// W=22 Received/Reject
// Z=25 TimeStamp Company Received
// AD=29 TimeStamp Technician Received
// AE=30 Fixed/Not Fixed
// AH=33 TimeStamp of Not Fixed
// AK=36 TimeStamp of Fixed
// AP=41 Invoice File (pdf)

const COL = {
  RESPONSE:       0,
  TIME_GMT:       1,
  GOVERNORATE:    2,
  CENTER:         3,
  AREA_OF_MALFUNCTION: 7,
  MALFUNCTION:    8,
  SUPPLIER:       11,
  COMPANY:        19,
  TS_ADMIN_SENT:  20,
  STATUS:         22,
  TS_CO_RECEIVED: 25,
  TS_TECH_RECEIVED: 29,
  TECH_STATUS:    30,
  TS_NOT_FIXED:   33,
  TS_FIXED:       36,
  PDF:            41
};

// ─── Raw DataCache read, shared by getFilterOptions/getDashboardData ─────────
// Keyed on CacheTimestamp so repeated calls (e.g. getFilterOptions +
// getDashboardData back-to-back on login, or rapid filter re-applies) reuse
// the same read instead of each doing their own full 42-column sheet read —
// and it's automatically invalidated the moment any of the 6 write-capable
// apps actually changes the data, not on a blind timer.
const RAW_CACHE_KEY_PREFIX = 'dashboard_raw_datacache_';
const RAW_CACHE_TTL_SECONDS = 300;

function readRawDataCache() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const cacheSheet = ss.getSheetByName('DataCache');
  if (!cacheSheet) return null;

  const lastRow = cacheSheet.getLastRow();
  if (lastRow < 1) return [];

  const tsSheet = ss.getSheetByName('CacheTimestamp');
  const currentTs = tsSheet ? tsSheet.getRange('A1').getValue() : 0;
  const cacheKey = RAW_CACHE_KEY_PREFIX + currentTs;

  const scriptCache = CacheService.getScriptCache();
  const cached = scriptCache.get(cacheKey);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {
      // Corrupt cache entry — fall through to a live read below.
    }
  }

  // Normalize any real Date-typed cells into formatted strings right away.
  // Without this, a Date survives a live read fine, but JSON.stringify()
  // below (for caching) silently converts it to an ISO "...T...Z" string,
  // and JSON.parse() on a cache hit never converts it back — so cached
  // rows would show raw ISO timestamps instead of the formatted ones.
  const data = cacheSheet.getRange(1, 1, lastRow, 42).getValues().map(row =>
    row.map(cell => cell instanceof Date
      ? Utilities.formatDate(cell, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
      : cell
    )
  );

  try {
    scriptCache.put(cacheKey, JSON.stringify(data), RAW_CACHE_TTL_SECONDS);
  } catch (e) {
    // Payload too large for CacheService's 100KB/key limit — skip caching,
    // the live read above already succeeded so this is not fatal.
    Logger.log('readRawDataCache: skipping cache, payload too large: ' + e);
  }

  return data;
}

// ─── Get Filter Options ───────────────────────────────────────────────────────
function getFilterOptions() {
  try {
    const data = readRawDataCache();
    if (data === null) return { error: 'DataCache not found' };
    if (data.length < 1) return { success: true, governorates: [], centers: [], suppliers: [] };

    const govSet = new Set();
    const centerSet = new Set();
    const supplierSet = new Set();

    data.forEach(row => {
      if (row[COL.RESPONSE]) {
        if (row[COL.GOVERNORATE]) govSet.add(String(row[COL.GOVERNORATE]).trim());
        if (row[COL.CENTER])      centerSet.add(String(row[COL.CENTER]).trim());
        if (row[COL.SUPPLIER])    supplierSet.add(String(row[COL.SUPPLIER]).trim());
      }
    });

    return {
      success: true,
      governorates: [...govSet].sort(),
      centers: [...centerSet].sort(),
      suppliers: [...supplierSet].sort()
    };
  } catch (e) {
    return { error: e.toString() };
  }
}

// ─── Get Dashboard Data (with optional filters) ───────────────────────────────
function getDashboardData(filters, page) {
  try {
    const raw = readRawDataCache();
    if (raw === null) return { error: 'Sheet "DataCache" not found' };
    if (raw.length < 1) {
      return { success: true, totalReports: 0, centers: [], malfunctions: [], companies: [], fixedPerGovernorate: [], tableRows: [], serverPaginated: false };
    }

    // Apply filters
    const filterGov      = filters && filters.governorate ? filters.governorate.trim().toLowerCase() : '';
    const filterCenter   = filters && filters.center      ? filters.center.trim().toLowerCase()      : '';
    const filterSupplier = filters && filters.supplier    ? filters.supplier.trim().toLowerCase()    : '';

    const data = raw.filter(row => {
      if (!row[COL.RESPONSE]) return false;
      if (filterGov      && String(row[COL.GOVERNORATE] || '').trim().toLowerCase() !== filterGov)      return false;
      if (filterCenter   && String(row[COL.CENTER]      || '').trim().toLowerCase() !== filterCenter)   return false;
      if (filterSupplier && String(row[COL.SUPPLIER]    || '').trim().toLowerCase() !== filterSupplier) return false;
      return true;
    });

    const totalReports = data.length;

    // Centers distribution
    const centersMap = {};
    data.forEach(row => {
      if (row[COL.GOVERNORATE]) {
        const v = String(row[COL.GOVERNORATE]).trim();
        centersMap[v] = (centersMap[v] || 0) + 1;
      }
    });
    const centers = Object.keys(centersMap)
      .map(name => ({ name, count: centersMap[name] }))
      .sort((a, b) => b.count - a.count)
      .slice(0, 5);

    // Malfunctions
    const malfunctionsMap = {};
    data.forEach(row => {
      if (row[COL.MALFUNCTION]) {
        const v = String(row[COL.MALFUNCTION]).trim();
        malfunctionsMap[v] = (malfunctionsMap[v] || 0) + 1;
      }
    });
    const malfunctions = Object.keys(malfunctionsMap)
      .map(name => ({ name, count: malfunctionsMap[name] }))
      .sort((a, b) => b.count - a.count);

    // Companies (Received only)
    const companiesMap = {};
    data.forEach(row => {
      if (row[COL.COMPANY] && row[COL.STATUS]) {
        const status = String(row[COL.STATUS]).trim().toLowerCase();
        if (status === 'received') {
          const v = String(row[COL.COMPANY]).trim();
          companiesMap[v] = (companiesMap[v] || 0) + 1;
        }
      }
    });
    const companies = Object.keys(companiesMap)
      .map(name => ({ name, count: companiesMap[name] }))
      .sort((a, b) => b.count - a.count)
      .slice(0, 8);

    // Fixed per Governorate
    const fixedMap = {};
    data.forEach(row => {
      if (row[COL.GOVERNORATE] && row[COL.TECH_STATUS]) {
        const ts = String(row[COL.TECH_STATUS]).trim();
        if (ts === 'Fixed') {
          const v = String(row[COL.GOVERNORATE]).trim();
          fixedMap[v] = (fixedMap[v] || 0) + 1;
        }
      }
    });
    const fixedPerGovernorate = Object.keys(fixedMap)
      .map(name => ({ name, count: fixedMap[name] }))
      .sort((a, b) => b.count - a.count);

    // Table rows — ship every matching row for small result sets (keeps
    // instant client-side paging for the common, filtered case); once the
    // result set is large, only ship the requested page so we're not
    // serializing/transferring thousands of rows nobody's currently
    // looking at. Charts/totals above are always computed from the full
    // filtered set regardless of this — only the row-level table is paged.
    const fmt = (val) => {
      if (!val || val === '' || (typeof val === 'number' && isNaN(val))) return '';
      if (val instanceof Date) {
        return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
      }
      return String(val).trim();
    };

    const TABLE_PAGE_SIZE = 20;
    const SERVER_PAGINATION_THRESHOLD = 200;
    const serverPaginated = totalReports > SERVER_PAGINATION_THRESHOLD;

    let rowsForClient = data;
    if (serverPaginated) {
      const pageNum = Math.max(1, page || 1);
      const start = (pageNum - 1) * TABLE_PAGE_SIZE;
      rowsForClient = data.slice(start, start + TABLE_PAGE_SIZE);
    }

    const tableRows = rowsForClient.map(row => ({
      supplier:          fmt(row[COL.SUPPLIER]),
      timeGMT:           fmt(row[COL.TIME_GMT]),
      tsAdminSent:       fmt(row[COL.TS_ADMIN_SENT]),
      tsCoReceived:      fmt(row[COL.TS_CO_RECEIVED]),
      tsTechReceived:    fmt(row[COL.TS_TECH_RECEIVED]),
      tsNotFixed:        fmt(row[COL.TS_NOT_FIXED]),
      tsFixed:           fmt(row[COL.TS_FIXED]),
      center:            fmt(row[COL.CENTER]),
      areaOfMalfunction: fmt(row[COL.AREA_OF_MALFUNCTION]),
      pdf:               fmt(row[COL.PDF])
    }));

    return {
      success: true,
      totalReports,
      centers,
      malfunctions,
      companies,
      fixedPerGovernorate,
      tableRows,
      serverPaginated
    };
  } catch (e) {
    Logger.log('getDashboardData error: ' + e);
    return { error: 'Error fetching dashboard data: ' + e.toString() };
  }
}

// ─── Check for Cache Update ───────────────────────────────────────────────────
function checkForCacheUpdate(lastKnownTs) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const tsSheet = ss.getSheetByName('CacheTimestamp');

    if (!tsSheet) return { hasUpdate: false, currentTs: 0 };

    const currentTs = tsSheet.getRange('A1').getValue();
    const tsNum = typeof currentTs === 'number' ? currentTs : 0;

    return {
      hasUpdate: lastKnownTs > 0 && tsNum > lastKnownTs,
      currentTs: tsNum
    };
  } catch (e) {
    Logger.log('checkForCacheUpdate error: ' + e);
    return { hasUpdate: false, currentTs: 0 };
  }
}
