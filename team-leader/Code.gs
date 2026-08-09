// This script is executed directly in the spreadsheet file (bound script)
// Use getActiveSpreadsheet() instead of openById()

const DATACACHE_SHEET = 'DataCache';
const TIMESTAMP_SHEET = 'CacheTimestamp';

const TARGET_TO_SOURCE_MAP = {
  'CapitalSubmited':              'Capital',
  'HawallySubmited':              'Hawally',
  'FarwaniyaSubmited':            'Farwaniya',
  'JahraSubmited':                'Jahra',
  'AhmadiSubmited':               'Ahmadi',
  'MubarakAlkabeerSubmited':      'MubarakAlKabeer',
  'AdanCenterSubmited':           'AdanCenter',
  'AmiriCenterSubmited':          'AmiriCenter',
  'BneidALGarSubmited':           'BneidALGar',
  'FarwaniyaCenterSubmited':      'FarwaniyaCenter',
  'JaberCenterSubmited':          'JaberCenter',
  'JahraCenterSubmited':          'JahraCenter',
  'NasserAlSaeedSubmited':        'NasserAlSaeed',
  'NewJahraDentalCenterSubmited': 'NewJahraDentalCenter',
  'SpecializedCenterSubmited':    'SpecializedCenter',
  'AhmadiProgramSubmited':        'AhmadiProgram',
  'CapitalProgramSubmited':       'CapitalProgram',
  'HawallyProgramSubmited':       'HawallyProgram',
  'FarwaniyaProgramSubmited':     'FarwaniyaProgram',
  'JahraProgramSubmited':         'JahraProgram',
  'MubarakALKabeerProgramSubmited': 'MubarakALKabeerProgram',
  'ALSabahHospitalSubmited':      'ALSabahHospital'
};

// ─── doGet ────────────────────────────────────────────────────────────────────
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Team Leader Login')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ─── Login ────────────────────────────────────────────────────────────────────
function doLogin(username, password) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName('Database 2');
  const data = dbSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === username && data[i][2] === password) {
      return {
        success: true,
        governate: data[i][0],
        center: data[i][4],
        tlName: data[i][3]
      };
    }
  }
  return { success: false };
}

// ═════════════════════════════════════════════════════════════════════════════
// DATACACHE
// ═════════════════════════════════════════════════════════════════════════════
function parseDateDDMMYYYY(dateStr) {
  if (!dateStr) return null;
  if (dateStr instanceof Date) return isNaN(dateStr) ? null : dateStr;
  if (typeof dateStr === 'number') return dateStr > 0 ? new Date(dateStr) : null;
  const s = String(dateStr).trim();
  if (!s) return null;
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4}) (\d{1,2}):(\d{1,2}):(\d{1,2})$/);
  if (!m) return null;
  const yyyy = m[3].length === 2 ? 2000 + +m[3] : +m[3];
  return new Date(yyyy, +m[2]-1, +m[1], +m[4], +m[5], +m[6]);
}

function readDataCache() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let cacheSheet = ss.getSheetByName(DATACACHE_SHEET);
  if (!cacheSheet || cacheSheet.getLastRow() < 1) {
    Logger.log('DataCache empty → full rebuild...');
    rebuildDataCache();
    cacheSheet = ss.getSheetByName(DATACACHE_SHEET);
    if (!cacheSheet || cacheSheet.getLastRow() < 1) return [];
  }
  const lastRow = cacheSheet.getLastRow();
  const lastCol = cacheSheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return [];

  const raw = cacheSheet.getRange(1, 1, lastRow, lastCol).getValues();
  return raw.map(r => r.map(cell => {
    if (cell === null || cell === undefined) return '';
    if (cell instanceof Date) {
      if (isNaN(cell)) return '';
      const dd   = String(cell.getDate()).padStart(2, '0');
      const MM   = String(cell.getMonth() + 1).padStart(2, '0');
      const yyyy = cell.getFullYear();
      const HH   = String(cell.getHours()).padStart(2, '0');
      const mm   = String(cell.getMinutes()).padStart(2, '0');
      const ss_  = String(cell.getSeconds()).padStart(2, '0');
      return `${dd}/${MM}/${yyyy} ${HH}:${mm}:${ss_}`;
    }
    return String(cell);
  }));
}

function rebuildDataCache() {
  const ALL_SHEET_NAMES = [
    'Capital','Hawally','Farwaniya','Jahra','Ahmadi','MubarakAlKabeer',
    'AdanCenter','AmiriCenter','BneidALGar','FarwaniyaCenter','JaberCenter',
    'JahraCenter','NasserAlSaeed','NewJahraDentalCenter','SpecializedCenter',
    'AhmadiProgram','CapitalProgram','HawallyProgram','FarwaniyaProgram',
    'JahraProgram','MubarakALKabeerProgram','ALSabahHospital'
  ];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let cacheSheet = ss.getSheetByName(DATACACHE_SHEET);
  if (!cacheSheet) cacheSheet = ss.insertSheet(DATACACHE_SHEET);

  const MAX_COL = 42;
  const allRows = [];

  ALL_SHEET_NAMES.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return;
    const numRows = lastRow - 2;
    const numCols = Math.min(sheet.getLastColumn(), MAX_COL);
    const displayRows = sheet.getRange(3, 1, numRows, numCols).getDisplayValues();
    const valueRows   = sheet.getRange(3, 1, numRows, numCols).getValues();

    for (let i = 0; i < displayRows.length; i++) {
      if (!displayRows[i][0] || String(displayRows[i][0]).trim() === '') continue;
      const safeRow = displayRows[i].map((cell, colIdx) => {
        if (cell === null || cell === undefined) return '';
        if (valueRows[i][colIdx] instanceof Date) {
          const d = valueRows[i][colIdx];
          if (isNaN(d)) return '';
          return String(d.getDate()).padStart(2,'0')+'/'+
                 String(d.getMonth()+1).padStart(2,'0')+'/'+
                 d.getFullYear()+' '+
                 String(d.getHours()).padStart(2,'0')+':'+
                 String(d.getMinutes()).padStart(2,'0')+':'+
                 String(d.getSeconds()).padStart(2,'0');
        }
        return String(cell);
      });
      allRows.push(safeRow);
    }
  });

  cacheSheet.clearContents();
  if (allRows.length > 0) {
    cacheSheet.getRange(1, 1, allRows.length, allRows[0].length).setValues(allRows);
  }

  let tsSheet = ss.getSheetByName(TIMESTAMP_SHEET);
  if (!tsSheet) tsSheet = ss.insertSheet(TIMESTAMP_SHEET);
  const nowEpoch = Date.now();
  tsSheet.getRange('A1').setValue(nowEpoch);

  Logger.log('DataCache full rebuild (TML): ' + allRows.length + ' row');
  return nowEpoch;
}

function rebuildDataCachePartial(targetSheetName) {
  const sourceSheetName = TARGET_TO_SOURCE_MAP[targetSheetName];
  if (!sourceSheetName) {
    Logger.log('Source sheet not found for: ' + targetSheetName);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tsSheet = ss.getSheetByName(TIMESTAMP_SHEET);
    if (tsSheet) tsSheet.getRange('A1').setValue(Date.now());
    return;
  }

  SpreadsheetApp.flush();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const MAX_COL = 42;

  let cacheSheet = ss.getSheetByName(DATACACHE_SHEET);
  if (!cacheSheet || cacheSheet.getLastRow() < 1) {
    Logger.log('DataCache doesnt exist yet — just update timestamp');
    let tsSheet = ss.getSheetByName(TIMESTAMP_SHEET);
    if (!tsSheet) tsSheet = ss.insertSheet(TIMESTAMP_SHEET);
    tsSheet.getRange('A1').setValue(Date.now());
    return;
  }

  const sourceSheet = ss.getSheetByName(sourceSheetName);
  if (!sourceSheet) return;

  const sourceLastRow = sourceSheet.getLastRow();
  const sourceIdentifiers = new Set();
  const newSourceRows = [];

  if (sourceLastRow >= 3) {
    const numRows = sourceLastRow - 2;
    const numCols = Math.min(sourceSheet.getLastColumn(), MAX_COL);
    const displayRows = sourceSheet.getRange(3, 1, numRows, numCols).getDisplayValues();
    const valueRows   = sourceSheet.getRange(3, 1, numRows, numCols).getValues();

    for (let i = 0; i < displayRows.length; i++) {
      if (!displayRows[i][0] || String(displayRows[i][0]).trim() === '') continue;
      const id = String(displayRows[i][0]).trim();
      sourceIdentifiers.add(id);

      const safeRow = displayRows[i].map((cell, colIdx) => {
        if (cell === null || cell === undefined) return '';
        if (valueRows[i][colIdx] instanceof Date) {
          const d = valueRows[i][colIdx];
          if (isNaN(d)) return '';
          return String(d.getDate()).padStart(2,'0')+'/'+
                 String(d.getMonth()+1).padStart(2,'0')+'/'+
                 d.getFullYear()+' '+
                 String(d.getHours()).padStart(2,'0')+':'+
                 String(d.getMinutes()).padStart(2,'0')+':'+
                 String(d.getSeconds()).padStart(2,'0');
        }
        return String(cell);
      });
      newSourceRows.push(safeRow);
    }
  }

  const cacheLastRow = cacheSheet.getLastRow();
  const cacheLastCol = cacheSheet.getLastColumn();
  let existingRows = [];
  if (cacheLastRow > 0 && cacheLastCol > 0) {
    const rawCache = cacheSheet.getRange(1, 1, cacheLastRow, cacheLastCol).getValues();
    existingRows = rawCache.filter(r => {
      const id = String(r[0] || '').trim();
      return id !== '' && !sourceIdentifiers.has(id);
    });
  }

  const allRows = [...existingRows, ...newSourceRows];
  const maxCols = allRows.reduce((m, r) => Math.max(m, r.length), 0);
  const normalized = allRows.map(r => {
    const copy = [...r];
    while (copy.length < maxCols) copy.push('');
    return copy;
  });

  cacheSheet.clearContents();
  if (normalized.length > 0) {
    cacheSheet.getRange(1, 1, normalized.length, normalized[0].length).setValues(normalized);
  }

  let tsSheet = ss.getSheetByName(TIMESTAMP_SHEET);
  if (!tsSheet) tsSheet = ss.insertSheet(TIMESTAMP_SHEET);
  tsSheet.getRange('A1').setValue(Date.now());

  Logger.log('DataCache partial rebuild (TML): ' + sourceSheetName + ', total: ' + normalized.length);
}

function setupDataCache() {
  rebuildDataCache();
  Logger.log('Setup done.');
}

// ═════════════════════════════════════════════════════════════════════════════
// MASTER FETCH: all tables in 1 server call
// ═════════════════════════════════════════════════════════════════════════════
function getAllTableData(governate, center) {
  const rows = readDataCache();

  // Single pass over the cache: each row is checked against both table
  // conditions instead of re-scanning the whole array once per table.
  const table1 = [];
  const table2Items = [];

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];

    // TABLE 1: Pending TML review — requires an exact center match (no
    // bypass for a falsy center, unlike table 2 below)
    if (r[3] === center && r[15] === 'pending') {
      table1.push([
        ...r.slice(0, 14),
        r[31] || '',  // Technician Status
        r[34] || '',  // Reason
        r[35] || ''   // Fixed Notes
      ]);
    }

    // TABLE 2: Teamleader Review (Fixed / Not Fixed) — falsy center matches all
    if ((!center || r[3] === center) &&
        (r[15] === 'Fixed' || r[15] === 'Not Fixed')) {
      table2Items.push({
        row: [
          ...r.slice(0, 14),
          r[31] || '',  // Technician Status
          r[33] || '',  // TimeStamp of Not Fixed
          r[34] || '',  // Reasons
          r[36] || '',  // TimeStamp of Fixed
          r[35] || '',  // Fixed Notes
          r[15] || '',  // Teamleader Review
          r[16] || '',  // Review TimeStamp
          r[37] || '',   // Teamleader Name
          r[41] || ''   // pdf
        ],
        sortKey: parseDateDDMMYYYY(r[16]) || new Date(0)
      });
    }
  }

  table1.sort((a, b) => (parseDateDDMMYYYY(b[1]) || 0) - (parseDateDDMMYYYY(a[1]) || 0));
  table2Items.sort((a, b) => b.sortKey - a.sortKey);
  const table2 = table2Items.map(item => item.row);

  return { table1, table2 };
}

// ─── Submit ───────────────────────────────────────────────────────────────────
function submitData(rows, governate, tlName) {
  const targetSheetName = getTargetSheetName(governate);
  const sourceSheetName = getSheetNameByGovernate(governate);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName(targetSheetName);
  const sourceSheet = ss.getSheetByName(sourceSheetName);

  rows.forEach(row => {
    const code      = row[13];
    const timestamp = new Date();

    const targetData = targetSheet.getDataRange().getValues();
    let rowIndex = -1;
    for (let i = 1; i < targetData.length; i++) {
      if (targetData[i][12] === code) { rowIndex = i + 1; break; }
    }

    if (rowIndex > 0) {
      if (row[14] === 'Not Fixed') {
        // Sent back for a second round: wipe the working columns and mark
        // it as a re-submission rather than recording a normal TL review.
        targetSheet.getRange(rowIndex, 4, 1, 21).clearContent(); // D:X
        targetSheet.getRange(rowIndex, 25, 1, 1).setValue(code + '2ndtime'); // Y
        targetSheet.getRange(rowIndex, 26, 1, 1).setValue('Not Fixed');      // Z
        targetSheet.getRange(rowIndex, 27, 1, 1).setValue(timestamp);       // AA
        targetSheet.getRange(rowIndex, 28, 1, 1).setValue(tlName);          // AB
        targetSheet.getRange(rowIndex, 30, 1, 1).setValue('2nd time');      // AD
      } else {
        targetSheet.getRange(rowIndex, 1, 1, 1).setValue(code);
        targetSheet.getRange(rowIndex, 2, 1, 1).setValue(row[14]);
        targetSheet.getRange(rowIndex, 3, 1, 1).setValue(timestamp);
        targetSheet.getRange(rowIndex, 24, 1, 1).setValue(tlName);
      }
    } else {
      // rowIndex -1 not valid for setRange, skip or append
      Logger.log('Row not found for code: ' + code);
    }

    // Update source sheet col L
    const sourceData = sourceSheet.getDataRange().getValues();
    for (let i = 1; i < sourceData.length; i++) {
      if (sourceData[i][0] === row[0] && sourceData[i][1] === row[1]) {
        sourceSheet.getRange(i + 1, 12).setValue(timestamp);
        break;
      }
    }
  });

  rebuildDataCachePartial(targetSheetName);
}

// ─── Helper functions ─────────────────────────────────────────────────────────
function getSheetNameByGovernate(governate) {
  const sheetMap = {
    'Capital': 'Capital',
    'Hawally': 'Hawally',
    'Farwaniya': 'Farwaniya',
    'Jahra': 'Jahra',
    'Ahmadi': 'Ahmadi',
    'Mubarak Alkabeer': 'MubarakAlKabeer',
    'Adan Center': 'AdanCenter',
    'Amiri Center': 'AmiriCenter',
    'Bneid AL Gar': 'BneidALGar',
    'Farwaniya Center': 'FarwaniyaCenter',
    'Jaber Center': 'JaberCenter',
    'Jahra Center': 'JahraCenter',
    'Nasser Al Saeed': 'NasserAlSaeed',
    'New Jahra Dental Center': 'NewJahraDentalCenter',
    'Specialized Center': 'SpecializedCenter',
    'Ahmadi Program': 'AhmadiProgram',
    'Capital Program': 'CapitalProgram',
    'Hawally Program': 'HawallyProgram',
    'Farwaniya Program': 'FarwaniyaProgram',
    'Jahra Program': 'JahraProgram',
    'Mubarak ALKabeer Program': 'MubarakALKabeerProgram',
    'AL Sabah Hospital': 'ALSabahHospital'
  };
  return sheetMap[governate] || 'Capital';
}

function getTargetSheetName(governate) {
  const targetSheetMap = {
    'Capital': 'CapitalSubmited',
    'Hawally': 'HawallySubmited',
    'Farwaniya': 'FarwaniyaSubmited',
    'Jahra': 'JahraSubmited',
    'Ahmadi': 'AhmadiSubmited',
    'Mubarak Alkabeer': 'MubarakAlkabeerSubmited',
    'Adan Center': 'AdanCenterSubmited',
    'Amiri Center': 'AmiriCenterSubmited',
    'Bneid AL Gar': 'BneidALGarSubmited',
    'Farwaniya Center': 'FarwaniyaCenterSubmited',
    'Jaber Center': 'JaberCenterSubmited',
    'Jahra Center': 'JahraCenterSubmited',
    'Nasser Al Saeed': 'NasserAlSaeedSubmited',
    'New Jahra Dental Center': 'NewJahraDentalCenterSubmited',
    'Specialized Center': 'SpecializedCenterSubmited',
    'Ahmadi Program': 'AhmadiProgramSubmited',
    'Capital Program': 'CapitalProgramSubmited',
    'Hawally Program': 'HawallyProgramSubmited',
    'Farwaniya Program': 'FarwaniyaProgramSubmited',
    'Jahra Program': 'JahraProgramSubmited',
    'Mubarak ALKabeer Program': 'MubarakALKabeerProgramSubmited',
    'AL Sabah Hospital': 'ALSabahHospitalSubmited'
  };
  return targetSheetMap[governate] || 'CapitalSubmited';
}

// ─── Notification (use CacheTimestamp) ─────────────────────────────────────
function checkForNotification(lastKnownTs) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tsSheet = ss.getSheetByName(TIMESTAMP_SHEET);
    if (!tsSheet) return { hasNewData: false, currentTs: 0 };
    const currentTs = tsSheet.getRange('A1').getValue();
    const tsNum = typeof currentTs === 'number' ? currentTs : 0;
    return {
      hasNewData: lastKnownTs > 0 && tsNum > lastKnownTs,
      currentTs: tsNum
    };
  } catch(e) {
    return { hasNewData: false, currentTs: 0 };
  }
}

function clearNotificationData() {
  // Still there for compatibility
}
