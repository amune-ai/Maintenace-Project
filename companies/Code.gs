const SPREADSHEET_ID = "1QGFrQHb7c6NAZTlhxMWClPox1TcQUrwVMeIlI2UMY8I";//'Input Regions Dev 3 (Send Email)'; // Replace with your Spreadsheet ID

// === أسطر الواتساب المضافة ===
var INSTANCE_ID = 'instance171817';
var TOKEN = 'z5gky4yzrcq1z6wb';
var GROUP_ID = '120363409425308350@g.us';

function sendWhatsAppUltra(row, destination, company) {
  var url = "https://api.ultramsg.com/" + INSTANCE_ID + "/messages/chat";
  var cleanTo = destination.toString().includes('@g.us') ? destination : destination.toString().replace(/[^\d]/g, '');
  var message = "🦷 *Maintenance Assignment* 🦷\n\n🎫 *No:* " + row[0] + "\n🏢 *Company:* " + company + "\n📍 *Region:* " + row[2] + "\n🏥 *Clinic:* " + row[3] + "\n🚪 *Room:* " + row[6] + "\n⚙️ *Model:* " + row[9] + "\n🔢 *Serial:* " + row[10] + "\n⚠️ *Malfunction:* " + row[8] + "\n👤 *Reporter:* " + row[4] + "\n📞 *Contact:* " + row[5];
  var payload = { "token": TOKEN, "to": cleanTo, "body": message };
  UrlFetchApp.fetch(url, { "method": "post", "payload": payload, "muteHttpExceptions": true });
}
// ==========================

// ─── Sheet name maps ──────────────────────────────────────────────────────────
const SHEET_MAP = {
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
  'Mubarak AL Kabeer Program': 'MubarakALKabeerProgram',
  'AL Sabah Hospital': 'ALSabahHospital'
};

const TARGET_SHEET_MAP = {
  'capital': 'CapitalSubmited',
  'hawally': 'HawallySubmited',
  'farwaniya': 'FarwaniyaSubmited',
  'jahra': 'JahraSubmited',
  'ahmadi': 'AhmadiSubmited',
  'mubarak alkabeer': 'MubarakAlKabeerSubmited',
  'adan center': 'AdanCenterSubmited',
  'amiri center': 'AmiriCenterSubmited',
  'bneid al gar': 'BneidALGarSubmited',
  'farwaniya center': 'FarwaniyaCenterSubmited',
  'jaber center': 'JaberCenterSubmited',
  'jahra center': 'JahraCenterSubmited',
  'nasser al saeed': 'NasserAlSaeedSubmited',
  'new jahra dental center': 'NewJahraDentalCenterSubmited',
  'specialized center': 'SpecializedCenterSubmited',
  'ahmadi program': 'AhmadiProgramSubmited',
  'capital program': 'CapitalProgramSubmited',
  'hawally program': 'HawallyProgramSubmited',
  'farwaniya program': 'FarwaniyaProgramSubmited',
  'jahra program': 'JahraProgramSubmited',
  'mubarak al kabeer program': 'MubarakALKabeerProgramSubmited',
  'al sabah hospital': 'ALSabahHospitalSubmited'
};

// ─── Reverse map: target sheet → source sheet ─────────────────────────────────
const TARGET_TO_SOURCE_MAP = {
  'CapitalSubmited':               'Capital',
  'HawallySubmited':               'Hawally',
  'FarwaniyaSubmited':             'Farwaniya',
  'JahraSubmited':                 'Jahra',
  'AhmadiSubmited':                'Ahmadi',
  'MubarakAlKabeerSubmited':       'MubarakAlKabeer',
  'AdanCenterSubmited':            'AdanCenter',
  'AmiriCenterSubmited':           'AmiriCenter',
  'BneidALGarSubmited':            'BneidALGar',
  'FarwaniyaCenterSubmited':       'FarwaniyaCenter',
  'JaberCenterSubmited':           'JaberCenter',
  'JahraCenterSubmited':           'JahraCenter',
  'NasserAlSaeedSubmited':         'NasserAlSaeed',
  'NewJahraDentalCenterSubmited': 'NewJahraDentalCenter',
  'SpecializedCenterSubmited':     'SpecializedCenter',
  'AhmadiProgramSubmited':         'AhmadiProgram',
  'CapitalProgramSubmited':        'CapitalProgram',
  'HawallyProgramSubmited':        'HawallyProgram',
  'FarwaniyaProgramSubmited':      'FarwaniyaProgram',
  'JahraProgramSubmited':          'JahraProgram',
  'MubarakALKabeerProgramSubmited': 'MubarakALKabeerProgram',
  'ALSabahHospitalSubmited':       'ALSabahHospital'
};

const ALL_SHEET_NAMES  = Object.values(SHEET_MAP);
const DATACACHE_SHEET  = 'DataCache';
const TIMESTAMP_SHEET  = 'CacheTimestamp';

// ─── doGet ────────────────────────────────────────────────────────────────────
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Companies Login')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ─── Login ────────────────────────────────────────────────────────────────────
function doLogin(username, password) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const data = ss.getSheetByName('CompanyDatabase4').getDataRange().getValues();

  const governates = [...new Set(
    data.slice(1).map(r => r[0]).filter(g => g && g.toString().trim() !== '')
  )].join(', ');

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === username && data[i][2] === password) {
      return {
        success: true,
        governate: data[i][0],
        adminName: data[i][3],
        centers: data[i][4].split(',').map(c => c.trim()),
        governateDropdown: governates.split(',').map(c => c.trim())
      };
    }
  }
  return { success: false };
}

// ═════════════════════════════════════════════════════════════════════════════
// DATACACHE: Shared sheet used by all webApps
// ═════════════════════════════════════════════════════════════════════════════

function rebuildDataCache() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  let cacheSheet = ss.getSheetByName(DATACACHE_SHEET);
  if (!cacheSheet) {
    cacheSheet = ss.insertSheet(DATACACHE_SHEET);
    Logger.log('New DataCache Sheet created');
  }

  const allRows = [];
  const MAX_COL = 47;

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
          const dd   = String(d.getDate()).padStart(2, '0');
          const MM   = String(d.getMonth() + 1).padStart(2, '0');
          const yyyy = d.getFullYear();
          const HH   = String(d.getHours()).padStart(2, '0');
          const mm   = String(d.getMinutes()).padStart(2, '0');
          const ss_  = String(d.getSeconds()).padStart(2, '0');
          return `${dd}/${MM}/${yyyy} ${HH}:${mm}:${ss_}`;
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

  Logger.log('DataCache full rebuild: ' + allRows.length + ' baris, ts=' + nowEpoch);
  return nowEpoch;
}

function rebuildDataCachePartial(targetSheetName) {
  const sourceSheetName = TARGET_TO_SOURCE_MAP[targetSheetName];
  if (!sourceSheetName) {
    Logger.log('rebuild Data Cache Partial: source sheet not found for ' + targetSheetName + ' — fallback full rebuild');
    rebuildDataCache();
    return;
  }

  SpreadsheetApp.flush();

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const MAX_COL = 47;

  let cacheSheet = ss.getSheetByName(DATACACHE_SHEET);
  if (!cacheSheet || cacheSheet.getLastRow() < 1) {
    Logger.log('DataCache is missing/empty — full rebuild');
    rebuildDataCache();
    return;
  }

  const sourceSheet = ss.getSheetByName(sourceSheetName);
  if (!sourceSheet) {
    Logger.log('Source sheet not found: ' + sourceSheetName);
    return;
  }

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
          const dd   = String(d.getDate()).padStart(2, '0');
          const MM   = String(d.getMonth() + 1).padStart(2, '0');
          const yyyy = d.getFullYear();
          const HH   = String(d.getHours()).padStart(2, '0');
          const mm   = String(d.getMinutes()).padStart(2, '0');
          const ss_  = String(d.getSeconds()).padStart(2, '0');
          return `${dd}/${MM}/${yyyy} ${HH}:${mm}:${ss_}`;
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
  const nowEpoch = Date.now();
  tsSheet.getRange('A1').setValue(nowEpoch);

  Logger.log('DataCache partial rebuild: ' + newSourceRows.length + ' row from ' + sourceSheetName + ', total cache: ' + normalized.length + ', ts=' + nowEpoch);
  return nowEpoch;
}

function readDataCache() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let cacheSheet = ss.getSheetByName(DATACACHE_SHEET);

  if (!cacheSheet || cacheSheet.getLastRow() < 1) {
    Logger.log('DataCache is empty/does not exist → full rebuild...');
    rebuildDataCache();
    cacheSheet = ss.getSheetByName(DATACACHE_SHEET);
    if (!cacheSheet || cacheSheet.getLastRow() < 1) {
      Logger.log('DataCache is still empty after rebuild — return empty');
      return { display: [], sortKeys: [] };
    }
  }

  const lastRow = cacheSheet.getLastRow();
  const lastCol = cacheSheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return { display: [], sortKeys: [] };

  const raw = cacheSheet.getRange(1, 1, lastRow, lastCol).getValues();

  const display = raw.map(r => r.map(cell => {
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

  const sortKeys = display.map(r => {
    const d = parseDateDDMMYYYY(r[36]);
    return d ? d.getTime() : 0;
  });

  return { display, sortKeys };
}

// ═════════════════════════════════════════════════════════════════════════════
// MASTER FETCH: filter from DataCache, send all tables at once
// ═════════════════════════════════════════════════════════════════════════════
// TL Notes: combines AR (Not Fixed) + AU (2nd time) into one display string
function tlNotes(r) {
  return [r[43], r[46]].filter(v => v && String(v).trim() !== '').join(' | ');
}

function getAllTableData(governate, center, currentAdminName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const cacheResult = readDataCache();
  if (!cacheResult) {
    Logger.log('readDataCache return null — return empty result');
    return { table1: [], table2: { tableData: [], companiesList: [] }, table3: [], table4: [], table5: [], tableTML: [] };
  }
  const { display: combinedDisplay, sortKeys: combinedValues } = cacheResult;

  // Single pass over the cache: each row is checked against all table
  // conditions instead of re-scanning the whole array once per table.
  const table1 = [];
  const table2Rows = [];
  const table3 = [];
  const table4 = [];
  const table5Items = [];
  const tableTML = [];

  for (let i = 0; i < combinedDisplay.length; i++) {
    const r = combinedDisplay[i];
    const centerMatch = !center || r[3] === center;

    if (r[23] === 'pending' && r[19] === currentAdminName) {
      table1.push([...r.slice(0, 14), r[19] || '', r[20] || '', r[21] || '', tlNotes(r)]);
    }

    if (
      (r[27] === 'Received' || r[27] === 'pending') &&
      centerMatch &&
      r[19] === currentAdminName &&
      (!r[26] || r[26].trim() === '')
    ) {
      table2Rows.push([
        r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7], r[8], r[9],
        r[10], r[11], r[12], r[13],
        r[19] || '', r[20] || '', r[21] || '', r[23] || '',
        r[24] || '', r[25] || '', r[28] || '', tlNotes(r)
      ]);
    }

    if (
      r[27] === 'Received' && r[31] === 'pending' &&
      centerMatch &&
      r[19] === currentAdminName
    ) {
      table3.push([
        r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7], r[8], r[9],
        r[10], r[11], r[12], r[13],
        r[19] || '', r[20] || '', r[21] || '',
        r[24] || '', r[25] || '', r[28] || '', r[29] || '', tlNotes(r)
      ]);
    }

    if (r[31] === 'Not Fixed' && centerMatch && r[19] === currentAdminName) {
      table4.push([
        r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7], r[8], r[9],
        r[10], r[11], r[12], r[13],
        r[19] || '', r[20] || '', r[21] || '',
        r[24] || '', r[25] || '', r[28] || '', r[29] || '',
        r[33] || '', r[34] || '', tlNotes(r)
      ]);
    }

    if (r[31] === 'Fixed' && centerMatch && r[19] === currentAdminName) {
      table5Items.push({
        row: [
          r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7], r[8], r[9],
          r[10], r[11], r[12], r[13],
          r[19] || '', r[20] || '', r[21] || '',
          r[24] || '', r[25] || '', r[28] || '', r[29] || '',
          r[33] || '',  // TimeStamp of Not Fixed
          r[34] || '',  // Reasons
          r[36] || '', r[35] || '', r[41] || '',  // pdf
          tlNotes(r)    // TL Notes
        ],
        sortKey: combinedValues[i]
      });
    }

    const tml = r[15];
    if (
      centerMatch &&
      (tml === 'Fixed' || tml === 'Not Fixed') &&
      r[19] === currentAdminName
    ) {
      tableTML.push([
        ...r.slice(0, 14),
        r[31] || '', r[33] || '', r[34] || '', r[36] || '', r[35] || '',
        r[15] || '', r[16] || '', r[37] || '', tlNotes(r)
      ]);
    }
  }

  table1.sort((a, b) => (parseDateDDMMYYYY(b[15]) || 0) - (parseDateDDMMYYYY(a[15]) || 0));
  table2Rows.sort((a, b) => (parseDateDDMMYYYY(b[19]) || 0) - (parseDateDDMMYYYY(a[19]) || 0));
  table3.sort((a, b) => (parseDateDDMMYYYY(b[20]) || 0) - (parseDateDDMMYYYY(a[20]) || 0));
  table4.sort((a, b) => (parseDateDDMMYYYY(b[21]) || 0) - (parseDateDDMMYYYY(a[21]) || 0));
  table5Items.sort((a, b) => b.sortKey - a.sortKey);
  const table5 = table5Items.map(item => item.row);
  tableTML.sort((a, b) => (parseDateDDMMYYYY(b[20]) || 0) - (parseDateDDMMYYYY(a[20]) || 0));

  let companiesList = [];
  try {
    const companiesSheet = ss.getSheetByName('Companies');
    const headers = companiesSheet.getRange(1, 1, 1, companiesSheet.getLastColumn()).getValues()[0];
    const adminColIndex = headers.indexOf(currentAdminName);
    if (adminColIndex !== -1) {
      const lastRow = companiesSheet.getLastRow();
      if (lastRow > 1) {
        companiesList = companiesSheet
          .getRange(2, adminColIndex + 1, lastRow - 1, 1)
          .getValues()
          .map(r => r[0])
          .filter(v => v && v !== '');
      }
    }
  } catch (e) { Logger.log('Companies sheet error: ' + e); }

  return {
    table1,
    table2: { tableData: table2Rows, companiesList },
    table3,
    table4,
    table5,
    tableTML
  };
}

function parseDateDDMMYYYY(dateStr) {
  if (!dateStr) return null;
  if (dateStr instanceof Date) return isNaN(dateStr) ? null : dateStr;
  if (typeof dateStr === 'number') return dateStr > 0 ? new Date(dateStr) : null;
  const s = String(dateStr).trim();
  if (!s) return null;
  const m = s.match(/^(\d{2})\/(\d{2})\/(\d{4}) (\d{2}):(\d{2}):(\d{2})$/);
  if (!m) return null;
  return new Date(+m[3], +m[2]-1, +m[1], +m[4], +m[5], +m[6]);
}

function submitStatusA(rows, governate, adminName, status, saveUsername) {
  const region = rows[0][2] ? rows[0][2].trim() : '';
  const targetSheetName = TARGET_SHEET_MAP[region.toLowerCase()] || 'CapitalSubmited';

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const targetSheet = ss.getSheetByName(targetSheetName);

  rows.forEach(row => {
    const code      = row[13];
    const companies = row[14];
    const timestamp = new Date();

    const colD = targetSheet.getRange(1, 4, targetSheet.getLastRow(), 1).getValues();
    let rowIndex = -1;
    for (let i = 1; i < colD.length; i++) {
      if (colD[i][0] === code) { rowIndex = i + 1; break; }
    }

    if (rowIndex > 0) {
      if (status === 'Reject') {
        const reason = row[row.length - 1];
        targetSheet.getRange(rowIndex, 4, 1, 7).clearContent();
        targetSheet.getRange(rowIndex, 25, 1, 1).setValue(code);
        targetSheet.getRange(rowIndex, 26, 1, 1).setValue(reason);
        targetSheet.getRange(rowIndex, 27, 1, 1).setValue(timestamp);
        targetSheet.getRange(rowIndex, 28, 1, 1).setValue(adminName);
      } else {
        targetSheet.getRange(rowIndex, 9,  1, 1).setValue(code);
        targetSheet.getRange(rowIndex, 10, 1, 1).setValue(status);
        targetSheet.getRange(rowIndex, 11, 1, 2).setValues([[companies, timestamp]]);
        targetSheet.getRange(rowIndex, 25, 1, 4).clearContent();
      }
    } else {
      if (status !== 'Reject') {
        targetSheet.appendRow([
          '', '', '', code || '', '', '', '', '', code || '',
          status || '', companies || '', timestamp || ''
        ]);
      }
    }
    ss.getSheetByName('comp log').appendRow([saveUsername, timestamp, status, companies, rows[0][2]]);
  });

  rebuildDataCachePartial(targetSheetName);
}

function submitStatusB(rows, governate, adminName, status, saveUsername) {
  const region = rows[0][2] ? rows[0][2].trim() : '';
  const targetSheetName = TARGET_SHEET_MAP[region.toLowerCase()] || 'CapitalSubmited';

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const targetSheet = ss.getSheetByName(targetSheetName);

  rows.forEach(row => {
    const code      = row[13];
    const companies = row[14];
    const timestamp = new Date();

    const colD = targetSheet.getRange(1, 4, targetSheet.getLastRow(), 1).getValues();
    let rowIndex = -1;
    for (let i = 1; i < colD.length; i++) {
      if (colD[i][0] === code) { rowIndex = i + 1; break; }
    }

    if (rowIndex > 0) {
      targetSheet.getRange(rowIndex, 4,  1, 1).setValue(code);
      targetSheet.getRange(rowIndex, 13, 1, 1).setValue(code);
      targetSheet.getRange(rowIndex, 14, 1, 1).setValue(status);
      targetSheet.getRange(rowIndex, 15, 1, 2).setValues([[companies, timestamp]]);
    } else {
      targetSheet.appendRow([
        '', '', '', code || '', '', '', '', '', '', '', '', '',
        code || '', status || '', companies || '', timestamp || ''
      ]);
    }
    ss.getSheetByName('comp log').appendRow([saveUsername, timestamp, status, companies]);

    // --- WhatsApp and Email Section ---
    if (companies && companies.trim() !== '') {
      const techEmailSheet = ss.getSheetByName('TechsEmails');
      if (techEmailSheet && techEmailSheet.getLastRow() > 1) {
        const techRange = techEmailSheet.getRange(2, 2, techEmailSheet.getLastRow() - 1, 3);
        const techValues = techRange.getValues();

        for (let i = 0; i < techValues.length; i++) {
          const techName  = (techValues[i][0] || '').toString().trim();
          const techEmail = (techValues[i][1] || '').toString().trim();
          const techPhone = (techValues[i][2] || '').toString().trim();

          if (techName === companies) {
            // === أسطر إرسال الواتساب المضافة ===
            if (techPhone !== '') sendWhatsAppUltra(row, techPhone, companies);
            if (GROUP_ID !== '') sendWhatsAppUltra(row, GROUP_ID, companies);
            // ===================================

            if (techEmail) {
              const htmlBody = `
                <table border="1" style="border-collapse:collapse;width:100%">
                  <tr style="background-color:#f2f2f2"><th style="padding:8px">Field</th><th style="padding:8px">Value</th></tr>
                  <tr><td style="padding:8px">Response Number</td><td style="padding:8px">${row[0]||''}</td></tr>
                  <tr><td style="padding:8px">Timestamp</td><td style="padding:8px">${row[1]||''}</td></tr>
                  <tr><td style="padding:8px">Region</td><td style="padding:8px">${row[2]||''}</td></tr>
                  <tr><td style="padding:8px">Clinic</td><td style="padding:8px">${row[3]||''}</td></tr>
                  <tr><td style="padding:8px">Room</td><td style="padding:8px">${row[6]||''}</td></tr>
                  <tr><td style="padding:8px">Reporter Name</td><td style="padding:8px">${row[4]||''}</td></tr>
                  <tr><td style="padding:8px">Phone</td><td style="padding:8px">${row[5]||''}</td></tr>
                  <tr><td style="padding:8px">Model</td><td style="padding:8px">${row[9]||''}</td></tr>
                  <tr><td style="padding:8px">Serial</td><td style="padding:8px">${row[10]||''}</td></tr>
                  <tr><td style="padding:8px">Supplier</td><td style="padding:8px">${row[11]||''}</td></tr>
                  <tr><td style="padding:8px">Malfunction</td><td style="padding:8px">${row[8]||''}</td></tr>
                  <tr><td style="padding:8px">Assigned</td><td style="padding:8px">${companies}</td></tr>
                  <tr><td style="padding:8px">Admin</td><td style="padding:8px">${adminName}</td></tr>
                </table>
                <p><strong>Status:</strong> Assigned to your team</p>
                <p>Please review and take necessary action.</p>`;
              GmailApp.sendEmail(techEmail, 'Equipment Issue Assigned - ' + row[0], '', { htmlBody });
              try {
                techEmailSheet.getRange('E2').setValue('Remaining quota: ' + MailApp.getRemainingDailyQuota());
              } catch(e) {
                techEmailSheet.getRange('E2').setValue('Quota check failed');
              }
            }
            break;
          }
        }
      }
    }
  });

  rebuildDataCachePartial(targetSheetName);
}

function getSheetNameByGovernate(governate) {
  return SHEET_MAP[governate] || 'Capital';
}

function getTargetSheetName(governate) {
  return TARGET_SHEET_MAP[(governate || '').toLowerCase()] || 'CapitalSubmited';
}

function getReasonsFromSheet() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('TechDatabase4');
    return sheet.getRange('I2:I').getValues().flat().filter(v => v && v.trim() !== '');
  } catch (e) {
    throw new Error('Failed to fetch reasons: ' + e.message);
  }
}

function checkForNotification(lastKnownTs) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const tsSheet = ss.getSheetByName(TIMESTAMP_SHEET);
    if (!tsSheet) return { hasNewData: false, currentTs: 0 };

    const currentTs = tsSheet.getRange('A1').getValue();
    const tsNum = typeof currentTs === 'number' ? currentTs : 0;

    return {
      hasNewData: lastKnownTs > 0 && tsNum > lastKnownTs,
      currentTs: tsNum
    };
  } catch (e) {
    return { hasNewData: false, currentTs: 0 };
  }
}

function clearNotificationData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const notiSheet = ss.getSheetByName('Notification Company');
    if (!notiSheet) return;
    const lastRow = notiSheet.getLastRow();
    if (lastRow < 2) return;
    notiSheet.getRange(2, 1, lastRow - 1, 2).clearContent();
  } catch (e) {
    Logger.log('Error clearing notification: ' + e);
  }
}

function setupDataCache() {
  rebuildDataCache();
  Logger.log('Setup is complete. The Data Cache and Timestamp Cache have been created.');
}
