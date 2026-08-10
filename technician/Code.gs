const SPREADSHEET_ID = "1QGFrQHb7c6NAZTlhxMWClPox1TcQUrwVMeIlI2UMY8I";
let templateDoc1 = "1MS3mpdYxqwm1r4o3g2DvVKUs__FKyk1GPdvkvobnVgY";
let templateDoc2 = "1aU4PTLO8gsuxHrv2HNhCP4QyxMkZGN_8PxZMpMyOkzg";
let folderID = "14ojLRUqZUbIQ9up3Il0QLMbJs9xA2L7t"; //https://drive.google.com/drive/folders/14ojLRUqZUbIQ9up3Il0QLMbJs9xA2L7t?usp=sharing
const DATACACHE_SHEET = 'DataCache';
const TIMESTAMP_SHEET = 'CacheTimestamp';

const TARGET_SHEET_MAP_TECH = {
  'capital':                'CapitalSubmited',
  'hawally':                'HawallySubmited',
  'farwaniya':              'FarwaniyaSubmited',
  'jahra':                  'JahraSubmited',
  'ahmadi':                 'AhmadiSubmited',
  'mubarak alkabeer':       'MubarakAlKabeerSubmited',
  'adan center':            'AdanCenterSubmited',
  'amiri center':           'AmiriCenterSubmited',
  'bneid al gar':            'BneidALGarSubmited',
  'farwaniya center':       'FarwaniyaCenterSubmited',
  'jaber center':           'JaberCenterSubmited',
  'jahra center':           'JahraCenterSubmited',
  'nasser al saeed':         'NasserAlSaeedSubmited',
  'new jahra dental center':'NewJahraDentalCenterSubmited',
  'specialized center':     'SpecializedCenterSubmited',
  'ahmadi program':         'AhmadiProgramSubmited',
  'capital program':        'CapitalProgramSubmited',
  'hawally program':        'HawallyProgramSubmited',
  'farwaniya program':      'FarwaniyaProgramSubmited',
  'jahra program':          'JahraProgramSubmited',
  'Mubarak AL Kabeer Program':'MubarakALKabeerProgramSubmited',
  'al sabah hospital':       'ALSabahHospitalSubmited'
};

const TARGET_TO_SOURCE_MAP = {
  'CapitalSubmited':              'Capital',
  'HawallySubmited':              'Hawally',
  'FarwaniyaSubmited':            'Farwaniya',
  'JahraSubmited':                'Jahra',
  'AhmadiSubmited':               'Ahmadi',
  'MubarakAlKabeerSubmited':      'MubarakAlKabeer',
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
  'MubarakALKabeerProgramSubmited':'MubarakALKabeerProgram',
  'ALSabahHospitalSubmited':      'ALSabahHospital'
};

const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

// ─── doGet ────────────────────────────────────────────────────────────────────
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Technician Login')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ─── Login ────────────────────────────────────────────────────────────────────
function doLogin(username, password) {
  const dbSheet = ss.getSheetByName('TechDatabase4');
  const data = dbSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === username && data[i][2] === password) {
      const centers = data[i][4].split(',').map(c => c.trim());
      return {
        success: true,
        governate: data[i][0],
        adminName: data[i][3],
        centers: centers
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

  let cacheSheet = ss.getSheetByName(DATACACHE_SHEET);
  if (!cacheSheet) cacheSheet = ss.insertSheet(DATACACHE_SHEET);

  const MAX_COL = 47;
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

  Logger.log('DataCache full rebuild (Tech): ' + allRows.length + ' row');
  return nowEpoch;
}

function rebuildDataCachePartial(targetSheetName) {
  const sourceSheetName = TARGET_TO_SOURCE_MAP[targetSheetName];
  if (!sourceSheetName) {
    Logger.log('Source sheet not found for: ' + targetSheetName);
    const tsSheet = ss.getSheetByName(TIMESTAMP_SHEET);
    if (tsSheet) tsSheet.getRange('A1').setValue(Date.now());
    return;
  }

  SpreadsheetApp.flush();

  const MAX_COL = 47;
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

  Logger.log('DataCache partial rebuild (Tech): ' + sourceSheetName + ', total: ' + normalized.length);
}

function setupDataCache() {
  rebuildDataCache();
  Logger.log('Setup complete.');
}

// ─── Generate PDF V.2 ─────────────────────────────────────────────────────────────
// ─── Generate PDF with full data from Main Database ─────────────────────────────────
function generatePdf(rowDataFromSubmit, ts, fn) {
  try {
    const templateId = templateDoc1;
    const folder = DriveApp.getFolderById(folderID);

    // Full headers (25 columns) - updated with TimeStamp of Fixed and Fixed Notes
    const headers = [
      'Response', 'Time GMT', 'Governorate', 'Center',
      'Reporter Name', 'Mobile Number', 'Room',
      'Area of Malfunction', 'Malfunction', 'Model',
      'Serial', 'Supplier', 'Recap Info', 'Code',
      'Company Name', 'TimeStamp Admin Sent', 'Admin Name',
      'Company Received By', 'TimeStamp Company Received',
      'Technician Name', 'TimeStamp Technician Received',
      'TimeStamp of Not Fixed', 'Reasons',
      'TimeStamp of Fixed', 'Fixed Notes'
    ];

    // Get complete data from Main Database
    const table2Sheet = ss.getSheetByName(DATACACHE_SHEET)
    const completeRowData = getCompleteRowFromTable2(table2Sheet, rowDataFromSubmit, ts, fn);

    const templateFile = DriveApp.getFileById(templateId);
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy HH:mm:ss');
    const fileName = completeRowData[3] + '_' + (completeRowData[13] || 'Unknown') + '_' + timestamp;

    // Copy template
    const newDoc = templateFile.makeCopy(fileName, folder);
    Utilities.sleep(1000);

    const doc = DocumentApp.openById(newDoc.getId());
    const body = doc.getBody();

    // Replace all placeholders with complete data from Main Database
    for (let i = 0; i < headers.length; i++) {
      const placeholder = '{' + headers[i] + '}';
      let value = completeRowData[i] || '';

      // Check if value is a timestamp/date and format it properly
      if (value instanceof Date) {
        // Format date as dd/MM/yyyy HH:mm:ss
        value = Utilities.formatDate(value, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
      } else {
        // For non-date values, convert to string normally
        value = String(value);
      }

      body.replaceText(placeholder, value);
    }

    body.replaceText('{GeneratedAt}', timestamp);

    doc.saveAndClose();

    // Convert to PDF
    const pdfBlob = newDoc.getAs('application/pdf');
    const pdfFile = folder.createFile(pdfBlob).setName(fileName + '.pdf');
    Utilities.sleep(500);

    // Delete temporary .doc files
    newDoc.setTrashed(true);

    const fileId = pdfFile.getId();
    const pdfUrl = 'https://drive.google.com/file/d/' + fileId + '/view';

    return pdfUrl;

  } catch(e) {
    Logger.log('=== generatePdf ERROR ===');
    Logger.log('Error: ' + e.message);
    Logger.log('Stack: ' + e.stack);
    throw new Error('Failed to generate PDF: ' + e.message);
  }
}

// ─── Helper function to get complete row data from Table 2 ─────────────────────
function getCompleteRowFromTable2(table2Sheet, rowDataFromSubmit, ts, fn) {
  // Force any pending writes to commit before reading the latest data
  SpreadsheetApp.flush();

  // Use getDisplayValues() to preserve timestamp format (not getValues())
  const displayData = table2Sheet.getDataRange().getDisplayValues();
  const headers = displayData[0]; // First row contains headers

  // Find matching row in Table 2 using Code column as unique identifier
  const uniqueIdentifier = rowDataFromSubmit[13]; // Code column (index 13) is unique

  Logger.log('Looking for Code: ' + uniqueIdentifier);
  Logger.log('Total rows in sheet: ' + displayData.length);

  let matchedRow = null;
  for (let i = 1; i < displayData.length; i++) {
    if (displayData[i][13] === uniqueIdentifier) {
      // Found matching row in Table 2
      Logger.log('Found matching row at index: ' + i);
      matchedRow = displayData[i];
      break;
    }
  }

  if (!matchedRow) {
    Logger.log('Warning: No matching row found in Table 2 for Code: ' + uniqueIdentifier);
    Logger.log('Last 5 Code values in column 13: ' + displayData.slice(-5).map(row => row[13]));
    return rowDataFromSubmit;
  }

  // Define column mapping from Table 2 headers to template headers
  // Format: headerName -> columnIndex in Table 2 (0-based indexing)
  const columnMapping = {
    'Response': 0,                          // Column A
    'Time GMT': 1,                          // Column B
    'Governorate': 2,                       // Column C
    'Center': 3,                            // Column D
    'Reporter Name': 4,                     // Column E
    'Mobile Number': 5,                     // Column F
    'Room': 6,                              // Column G
    'Area of Malfunction': 7,               // Column H
    'Malfunction': 8,                       // Column I
    'Model': 9,                             // Column J
    'Serial': 10,                           // Column K
    'Supplier': 11,                         // Column L
    'Recap Info': 12,                       // Column M
    'Code': 13,                             // Column N
    'Company Name': 19,                     // Column T
    'TimeStamp Admin Sent': 20,             // Column U
    'Admin Name': 21,                       // Column V
    'Company Received By': 24,              // Column Y
    'TimeStamp Company Received': 25,       // Column Z
    'Technician Name': 28,                  // Column AC
    'TimeStamp Technician Received': 29,    // Column AD
    'TimeStamp of Not Fixed': 33,           // Column AH
    'Reasons': 34                           // Column AI
    // Note: 'TimeStamp of Fixed' and 'Fixed Notes' will be taken from parameters ts and fn
  };

  // Build complete row data using the mapping
  const templateHeaders = [
    'Response', 'Time GMT', 'Governorate', 'Center',
    'Reporter Name', 'Mobile Number', 'Room',
    'Area of Malfunction', 'Malfunction', 'Model',
    'Serial', 'Supplier', 'Recap Info', 'Code',
    'Company Name', 'TimeStamp Admin Sent', 'Admin Name',
    'Company Received By', 'TimeStamp Company Received',
    'Technician Name', 'TimeStamp Technician Received',
    'TimeStamp of Not Fixed', 'Reasons',
    'TimeStamp of Fixed', 'Fixed Notes'
  ];

  // Extract data from matched row using column mapping
  const completeRowData = [];
  for (let i = 0; i < templateHeaders.length; i++) {
    const headerName = templateHeaders[i];

    // Check if this is one of the parameters from submit
    if (headerName === 'TimeStamp of Fixed') {
      // Use ts parameter (timestamp from Table 3 submit)
      completeRowData.push(ts || '');
      Logger.log('Added TimeStamp of Fixed from parameter: ' + ts);
    } else if (headerName === 'Fixed Notes') {
      // Use fn parameter (fixed notes from Table 3 submit)
      completeRowData.push(fn || '');
      Logger.log('Added Fixed Notes from parameter: ' + fn);
    } else {
      // Use data from matched row based on column mapping
      const columnIndex = columnMapping[headerName];
      const cellValue = matchedRow[columnIndex] || '';
      completeRowData.push(cellValue);
    }
  }

  Logger.log('Successfully extracted ' + completeRowData.length + ' columns from matched row and parameters');

  return completeRowData;
}

// ═════════════════════════════════════════════════════════════════════════════
// MASTER FETCH: all tables in 1 server call
// ═════════════════════════════════════════════════════════════════════════════
function getAllTableData(governate, center, currentAdminName) {
  const rows = readDataCache();

  // TL Notes: combines AQ, AR, AT, AU (skips AS) into one display string
  function tlNotes(r) {
    return [r[42], r[43], r[45], r[46]].filter(v => v && String(v).trim() !== '').join(' | ');
  }

  // Single pass over the cache: each row is checked against all 4 table
  // conditions instead of re-scanning the whole array once per table.
  const table1Items = [];
  const table2Items = [];
  const table3Items = [];
  const tableTMLItems = [];

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (r[28] !== currentAdminName) continue;

    // TABLE 1: Tech Received — col[27]='Received', col[30]=''
    if (r[27] === 'Received' && r[30] === '') {
      table1Items.push({
        row: [
          ...r.slice(0, 14),
          r[19] || '',  // Company Name
          r[20] || '',  // TimeStamp Admin Sent
          r[21] || '',  // Admin Name
          r[24] || '',  // Company Received By
          r[25] || '',  // TimeStamp Company Received
          r[28] || '',  // Technician Name
          r[29] || '',  // TimeStamp Technician Received
          tlNotes(r)    // TL Notes
        ],
        sortKey: parseDateDDMMYYYY(r[29]) || new Date(0)
      });
    }

    // TABLE 2: Fixed — col[31]='Fixed'
    if (r[31] === 'Fixed') {
      table2Items.push({
        row: [
          ...r.slice(0, 14),
          r[19] || '',  // Company Name
          r[20] || '',  // TimeStamp Admin Sent
          r[21] || '',  // Admin Name
          r[24] || '',  // Company Received By
          r[25] || '',  // TimeStamp Company Received
          r[28] || '',  // Technician Name
          r[29] || '',  // TimeStamp Technician Received
          r[33] || '',  // TimeStamp of Not Fixed
          r[34] || '',  // Reasons
          r[36] || '',  // TimeStamp of Fixed
          r[35] || '',   // Fixed Notes
          r[41] || '',  // pdf
          tlNotes(r)    // TL Notes
        ],
        sortKey: parseDateDDMMYYYY(r[36]) || new Date(0)
      });
    }

    // TABLE 3: Not Fixed — col[31]='Not Fixed'
    if (r[31] === 'Not Fixed') {
      table3Items.push({
        row: [
          ...r.slice(0, 14),
          r[19] || '',  // Company Name
          r[20] || '',  // TimeStamp Admin Sent
          r[21] || '',  // Admin Name
          r[24] || '',  // Company Received By
          r[25] || '',  // TimeStamp Company Received
          r[28] || '',  // Technician Name
          r[29] || '',  // TimeStamp Technician Received
          r[33] || '',  // TimeStamp of Not Fixed
          r[34] || '',  // Reasons
          tlNotes(r)    // TL Notes
        ],
        sortKey: parseDateDDMMYYYY(r[33]) || new Date(0)
      });
    }

    // TABLE TML: Teamleader Review — col[15]='Fixed' or 'Not Fixed'
    if (r[15] === 'Fixed' || r[15] === 'Not Fixed') {
      tableTMLItems.push({
        row: [
          ...r.slice(0, 14),
          r[31] || '',  // Technician Status
          r[33] || '',  // TimeStamp of Not Fixed
          r[34] || '',  // Reasons
          r[36] || '',  // TimeStamp of Fixed
          r[35] || '',  // Fixed Notes
          r[15] || '',  // Teamleader Review
          r[16] || '',  // Review TimeStamp
          r[37] || '',  // Teamleader Name
          tlNotes(r)    // TL Notes
        ],
        sortKey: parseDateDDMMYYYY(r[16]) || new Date(0)
      });
    }
  }

  table1Items.sort((a, b) => b.sortKey - a.sortKey);
  const table1 = table1Items.map(item => item.row);
  table2Items.sort((a, b) => b.sortKey - a.sortKey);
  const table2 = table2Items.map(item => item.row);
  table3Items.sort((a, b) => b.sortKey - a.sortKey);
  const table3 = table3Items.map(item => item.row);
  tableTMLItems.sort((a, b) => b.sortKey - a.sortKey);
  const tableTML = tableTMLItems.map(item => item.row);

  return { table1, table2, table3, tableTML };
}

// ─── Helper: get target sheet name ───────────────────────────────────────────
function getTargetSheetNameByRegion(region) {
  const key = (region || '').toString().trim().toLowerCase();
  return TARGET_SHEET_MAP_TECH[key] || 'CapitalSubmited';
}

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
    'Mubarak AL Kabeer Program': 'MubarakALKabeerProgram',
    'AL Sabah Hospital': 'ALSabahHospital'
  };
  return sheetMap[governate] || 'Capital';
}

// ─── Submit Data (Fixed / Not Fixed from Table 1) ─────────────────────────────
function submitData(rows, governate, tlName) {
  const targetSheetName = getTargetSheetNameByRegion(rows[0][2]);
  const targetSheet = ss.getSheetByName(targetSheetName);

  rows.forEach(row => {
    const code       = row[13];
    const fixedNotes = row[21];
    const timestamp  = new Date();

    const targetData = targetSheet.getDataRange().getValues();
    let rowIndex = -1;

    for (let i = 2; i < targetData.length; i++) {   // start from row 3
      if (targetData[i][3] === code) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex > 0) {
      targetSheet.getRange(rowIndex, 18, 1, 1).setValue(row[14]);

      if (row[14] === 'Fixed') {
        let pdfUrl = '';

        try {
          pdfUrl = generatePdf(row, timestamp, fixedNotes);        // ← now directly string
          console.log('PDF successfully created: ' + pdfUrl);
        } catch(e) {
          Logger.log('PDF generation failed for code ' + code + ': ' + e.message);
          pdfUrl = 'PDF Error: ' + e.message.substring(0, 50); // so we know the error
        }

        targetSheet.getRange(rowIndex, 17, 1, 1).setValue(code);
        targetSheet.getRange(rowIndex, 18, 1, 1).setValue('Fixed');
        targetSheet.getRange(rowIndex, 22, 1, 1).setValue(fixedNotes);
        targetSheet.getRange(rowIndex, 23, 1, 1).setValue(timestamp);
        targetSheet.getRange(rowIndex, 29, 1, 1).setValue(pdfUrl);

      } else {
        // Not Fixed
        targetSheet.getRange(rowIndex, 17, 1, 1).setValue(code);
        targetSheet.getRange(rowIndex, 18, 1, 1).setValue('Not Fixed');
        targetSheet.getRange(rowIndex, 20, 1, 1).setValue(timestamp);
        targetSheet.getRange(rowIndex, 21, 1, 1).setValue(row[21]);
        targetSheet.getRange(rowIndex, 29, 1, 1).setValue(''); // empty
      }
    }
  });

  rebuildDataCachePartial(targetSheetName);
  return "Success";
}

// ─── Submit Data 2 (Fixed dari Table 3) ──────────────────────────────────────
function submitData2(rows, governate, tlName) {
  const targetSheetName = getTargetSheetNameByRegion(rows[0][2]);
  const targetSheet = ss.getSheetByName(targetSheetName);

  rows.forEach(row => {
    const code      = row[13];
    const timestamp = new Date();

    const targetData = targetSheet.getDataRange().getValues();
    let rowIndex = -1;
    for (let i = 1; i < targetData.length; i++) {
      if (targetData[i][16] === code) { rowIndex = i + 1; break; }
    }

    if (rowIndex > 0) {

      // Generate PDF for Fixed from Table 3 (template 2)
      let pdfUrl = '';

      try {
        pdfUrl = generatePdf(row, timestamp, row[23]);        // ← now directly string
        console.log('PDF created successfully: ' + pdfUrl);
      } catch(e) {
        Logger.log('PDF generation failed for code ' + code + ': ' + e.message);
        pdfUrl = 'PDF Error: ' + e.message.substring(0, 50); // o we know the error
      }

      targetSheet.getRange(rowIndex, 18, 1, 1).setValue('Fixed');
      targetSheet.getRange(rowIndex, 20, 1, 1).setValue(row[21]);
      targetSheet.getRange(rowIndex, 21, 1, 1).setValue(row[22]);
      targetSheet.getRange(rowIndex, 22, 1, 1).setValue(row[23]);
      targetSheet.getRange(rowIndex, 23, 1, 1).setValue(timestamp);
      targetSheet.getRange(rowIndex, 29, 1, 1).setValue(pdfUrl);
    }
  });

  rebuildDataCachePartial(targetSheetName);
  return "Success"; // send back to client
}

// ─── Get reasons ──────────────────────────────────────────────────────────────
function getReasonsFromSheet() {
  try {
    const sheet = ss.getSheetByName('TechDatabase4');
    return sheet.getRange('H2:H').getValues()
      .flat()
      .filter(v => v && v.toString().trim() !== '');
  } catch(e) {
    throw new Error('Failed to fetch reasons: ' + e.message);
  }
}

// ─── Notification (use CacheTimestamp) ─────────────────────────────────────
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
  } catch(e) {
    return { hasNewData: false, currentTs: 0 };
  }
}

function clearNotificationData() {
  // Still there for compatibility
};

//Just use to get authorization from Drive API
function authorizeDrive() {
  DriveApp.getRootFolder();
  const testDoc = DocumentApp.create('test_auth_delete_me');
  DriveApp.getFileById(testDoc.getId()).setTrashed(true);
  Logger.log('Authorization successful');
}
