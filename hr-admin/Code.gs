const SPREADSHEET_ID = "1QGFrQHb7c6NAZTlhxMWClPox1TcQUrwVMeIlI2UMY8I";//'file name: NwCopy Test1 Dental Primary Care - Dev 2 (Send Email)'; // Replace with your Spreadsheet ID

const DATACACHE_SHEET = 'DataCache';
const TIMESTAMP_SHEET = 'CacheTimestamp';

const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

// ─── doGet ────────────────────────────────────────────────────────────────────
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('HR Admin Login')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ─── Login ────────────────────────────────────────────────────────────────────
function doLogin(username, password) {
  try {
    const dbSheet = ss.getSheetByName('HRAdmin');
    const data = dbSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === username && data[i][1] === password) {
        const centers = data[i][3] ? data[i][3].split(',').map(c => c.trim()) : [];
        return {
          success: true,
          governate: data[i][0],
          adminName: data[i][2] || '',
          centers: centers
        };
      }
    }
    return { success: false };
  } catch(e) {
    console.error('Error in doLogin:', e);
    return { success: false };
  }
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

  Logger.log('DataCache full rebuild (HRAdmin): ' + allRows.length + ' baris');
  return nowEpoch;
}

function setupDataCache() {
  rebuildDataCache();
  Logger.log('Setup selesai.');
}

// ═════════════════════════════════════════════════════════════════════════════
// MASTER FETCH: all tables in 1 server call
// ═════════════════════════════════════════════════════════════════════════════
function getAllTableData(governate, center) {
  const rows = readDataCache();

  // TL Notes: combines AR (Not Fixed) + AU (2nd time) into one display string
  function tlNotes(r) {
    return [r[43], r[46]].filter(v => v && String(v).trim() !== '').join(' | ');
  }

  // Single pass over the cache: each row is checked against all 7 table
  // conditions instead of re-scanning the whole array once per table.
  const table1 = [];
  const table2 = [];
  const table3 = [];
  const table4 = [];
  const table5 = [];
  const table6Items = [];
  const tableTMLItems = [];

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (center && r[3] !== center) continue;

    // TABLE 1: Rejected — col[18]='pending', col[38] (company reject) or
    // col[43]/AR (TL not-fixed 2nd-time) has content
    if (r[18] === 'pending' && ((r[38] && r[38].trim() !== '') || (r[43] && r[43].trim() !== ''))) {
      table1.push([
        ...r.slice(0, 14),
        r[19] || '',  // Company Name
        r[38] || '',  // Rejected Reason
        r[39] || '',  // Rejected TimeStamp
        r[40] || '',  // Rejected Company
        tlNotes(r)    // TL Notes
      ]);
    }

    // TABLE 2: Sent + pending company — col[18]='Sent', col[23]='pending'
    if (r[18] === 'Sent' && r[23] === 'pending') {
      table2.push([
        ...r.slice(0, 14),
        r[19] || '',  // Company Name
        r[20] || '',  // TimeStamp Admin Sent
        r[21] || '',  // Admin Name
        tlNotes(r)    // TL Notes
      ]);
    }

    // TABLE 3: Company Received + tech pending — col[23]='Received', col[27]='pending'
    if (r[23] === 'Received' && r[27] === 'pending') {
      table3.push([
        ...r.slice(0, 14),
        r[19] || '',  // Company Name
        r[20] || '',  // TimeStamp Admin Sent
        r[21] || '',  // Admin Name
        r[24] || '',  // Company Received By
        r[25] || '',  // TimeStamp Company Received
        tlNotes(r)    // TL Notes
      ]);
    }

    // TABLE 4: Tech Received + pending fix — col[27]='Received', col[31]='pending'
    if (r[27] === 'Received' && r[31] === 'pending') {
      table4.push([
        ...r.slice(0, 14),
        r[19] || '',  // Company Name
        r[20] || '',  // TimeStamp Admin Sent
        r[21] || '',  // Admin Name
        r[24] || '',  // Company Received By
        r[25] || '',  // TimeStamp Company Received
        r[28] || '',  // Technician Name
        r[29] || '',  // TimeStamp Technician Received
        tlNotes(r)    // TL Notes
      ]);
    }

    // TABLE 5: Not Fixed — col[31]='Not Fixed'
    if (r[31] === 'Not Fixed') {
      table5.push([
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
      ]);
    }

    // TABLE 6: Fixed — col[31]='Fixed'
    if (r[31] === 'Fixed') {
      table6Items.push({
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
          r[14] || '',  // Teamleader Review
          r[16] || '',  // Review TimeStamp
          r[37] || '',  // Teamleader Name
          tlNotes(r)    // TL Notes
        ],
        sortKey: parseDateDDMMYYYY(r[16]) || new Date(0)
      });
    }
  }

  table1.sort((a, b) => (parseDateDDMMYYYY(b[1]) || 0) - (parseDateDDMMYYYY(a[1]) || 0));
  table2.sort((a, b) => (parseDateDDMMYYYY(b[15]) || 0) - (parseDateDDMMYYYY(a[15]) || 0));
  table3.sort((a, b) => (parseDateDDMMYYYY(b[18]) || 0) - (parseDateDDMMYYYY(a[18]) || 0));
  table4.sort((a, b) => (parseDateDDMMYYYY(b[20]) || 0) - (parseDateDDMMYYYY(a[20]) || 0));
  table5.sort((a, b) => (parseDateDDMMYYYY(b[21]) || 0) - (parseDateDDMMYYYY(a[21]) || 0));
  table6Items.sort((a, b) => b.sortKey - a.sortKey);
  const table6 = table6Items.map(item => item.row);
  tableTMLItems.sort((a, b) => b.sortKey - a.sortKey);
  const tableTML = tableTMLItems.map(item => item.row);

  return { table1, table2, table3, table4, table5, table6, tableTML };
}

// ─── Notification (pakai CacheTimestamp) ─────────────────────────────────────
function checkForNotification(lastKnownTs) {
  try {
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
