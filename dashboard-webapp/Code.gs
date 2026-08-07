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

const COL = {
  RESPONSE:       0,
  TIME_GMT:       1,
  GOVERNORATE:    2,
  CENTER:         3,
  MALFUNCTION:    8,
  SUPPLIER:       11,
  COMPANY:        19,
  TS_ADMIN_SENT:  20,
  STATUS:         22,
  TS_CO_RECEIVED: 25,
  TS_TECH_RECEIVED: 29,
  TECH_STATUS:    30,
  TS_NOT_FIXED:   33,
  TS_FIXED:       36
};

// ─── Get Filter Options ───────────────────────────────────────────────────────
function getFilterOptions() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const cacheSheet = ss.getSheetByName('DataCache');
    if (!cacheSheet) return { error: 'DataCache not found' };

    const lastRow = cacheSheet.getLastRow();
    if (lastRow < 1) return { success: true, governorates: [], centers: [], suppliers: [] };

    const data = cacheSheet.getRange(1, 1, lastRow, 42).getValues();

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
function getDashboardData(filters) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const cacheSheet = ss.getSheetByName('DataCache');

    if (!cacheSheet) return { error: 'Sheet "DataCache" not found' };

    const lastRow = cacheSheet.getLastRow();
    if (lastRow < 1) {
      return { success: true, totalReports: 0, centers: [], malfunctions: [], companies: [], fixedPerGovernorate: [], tableRows: [] };
    }

    const raw = cacheSheet.getRange(1, 1, lastRow, 42).getValues();

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

    // Table rows
    const fmt = (val) => {
      if (!val || val === '' || (typeof val === 'number' && isNaN(val))) return '';
      if (val instanceof Date) {
        return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
      }
      return String(val).trim();
    };

    const tableRows = data.map(row => ({
      supplier:          fmt(row[COL.SUPPLIER]),
      timeGMT:           fmt(row[COL.TIME_GMT]),
      tsAdminSent:       fmt(row[COL.TS_ADMIN_SENT]),
      tsCoReceived:      fmt(row[COL.TS_CO_RECEIVED]),
      tsTechReceived:    fmt(row[COL.TS_TECH_RECEIVED]),
      tsNotFixed:        fmt(row[COL.TS_NOT_FIXED]),
      tsFixed:           fmt(row[COL.TS_FIXED])
    }));

    return {
      success: true,
      totalReports,
      centers,
      malfunctions,
      companies,
      fixedPerGovernorate,
      tableRows
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
