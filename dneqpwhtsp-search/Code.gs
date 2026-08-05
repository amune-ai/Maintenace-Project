//File name NwCopy Test1 Dental Primary Care - Dev 2 (Send Email)
var SHEET_ID = '1QGFrQHb7c6NAZTlhxMWClPox1TcQUrwVMeIlI2UMY8I'; // Your Google Sheet ID for target submit

var SHEET_ID2 = '1e4J3f8TIhzT2bJFGlxqqOYxGMpsS6kTBE-uLRz_m9q0'; // Your Google Sheet ID

// --- أسطر خدمة الواتساب (المتغيرات) ---
var INSTANCE_ID = 'instance171817';
var TOKEN = 'z5gky4yzrcq1z6wb';
var GROUP_ID = '120363409425308350@g.us';
// ------------------------------------

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
  'MubarakALKabeerProgramSubmited': 'MubarakALKabeerProgram',
  'ALSabahHospitalSubmited':      'ALSabahHospital'
};

const ALL_SHEET_NAMES  = Object.values(SHEET_MAP);
const DATACACHE_SHEET  = 'DataCache';
const TIMESTAMP_SHEET  = 'CacheTimestamp';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Maintenance Reports')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getData() {
  var ss = SpreadsheetApp.openById(SHEET_ID2);
  var sheet = ss.getSheetByName('Dependent Dropdown3');

  var rangeMain = sheet.getRange('A2:H' + sheet.getLastRow());
  var valuesMain = rangeMain.getDisplayValues().filter(row => row.some(cell => cell !== ''));

  var rangeAdditional = sheet.getRange('L2:L' + sheet.getLastRow());
  var valuesAdditional = rangeAdditional.getValues()
    .flat()
    .filter(val => val !== '')
    .sort();
  var uniqueAdditionalOptions = [...new Set(valuesAdditional)];

  return {
    data: valuesMain,
    additionalOptions: uniqueAdditionalOptions
  };
}

function submitForm(formData) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var dataSheet = ss.getSheetByName('Data');

    var region = (formData.region || '').toString().trim();
    var regionShort = '';

    if (region) {
      var words = region.split(/\s+/).filter(word => word.length > 0);
      if (words.length >= 1) {
        var maxWords = Math.min(words.length, 4);
        for (var i = 0; i < maxWords; i++) {
          regionShort += words[i].substring(0, 3);
        }
      }
    }

    var lastRow = dataSheet.getLastRow();
    var sequence = 10001;

    if (lastRow >= 2) {
      var responseNumbers = dataSheet.getRange("A2:A" + lastRow).getValues();
      var regions = dataSheet.getRange("C2:C" + lastRow).getValues();
      var maxSequence = 10000;

      for (var i = 0; i < responseNumbers.length; i++) {
        var respNum = responseNumbers[i][0] ? responseNumbers[i][0].toString().trim() : '';
        var reg = regions[i][0] ? regions[i][0].toString().trim() : '';

        if (reg === region && respNum.startsWith('#')) {
          var numPart = respNum.substring(1, respNum.length - regionShort.length);
          if (!isNaN(numPart)) {
            var num = parseInt(numPart, 10);
            if (num > maxSequence) {
              maxSequence = num;
            }
          }
        }
      }
      sequence = maxSequence + 1;
    }

    var responseNumber = `#${sequence}${regionShort}`;

  var timestamp = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'dd/MM/yyyy HH:mm:ss'
  );

  var generatedCode = responseNumber +
                      (formData.room || '').toString().trim().substring(0, 3) +
                      (formData.serial || '');

  dataSheet.appendRow([
    responseNumber,           // A
    timestamp,                // B
    formData.region,          // C
    formData.clinic,          // D
    formData.room,            // E
    formData.name,            // F
    formData.manufacturer,    // G
    formData.model,           // H
    formData.serial,          // I
    formData.supplier,        // J
    formData.reporterName,    // K
    formData.phone,           // L
    formData.additionalField, // M
    generatedCode             // N
  ]);

  var regionUpper = region.toUpperCase().replace(/\s+/g, '');
  var targetSheetName = 'AsimahSubmited';

  if (regionUpper === 'CAPITAL') {
    targetSheetName = 'CapitalSubmited';
  } else if (regionUpper === 'HAWALLY') {
    targetSheetName = 'HawallySubmited';
  } else if (regionUpper === 'FARWANIYA') {
    targetSheetName = 'FarwaniyaSubmited';
  } else if (regionUpper === 'JAHRA') {
    targetSheetName = 'JahraSubmited';
  } else if (regionUpper === 'MUBARAKALKABEER') {
    targetSheetName = 'MubarakAlKabeerSubmited';
  } else if (regionUpper === 'AHMADI') {
    targetSheetName = 'AhmadiSubmited';
  } else if (regionUpper === 'ADANCENTER') {
    targetSheetName = 'AdanCenterSubmited';
    } else if (regionUpper === 'AMIRICENTER') {
    targetSheetName = 'AmiriCenterSubmited';
} else if (regionUpper === 'BNEIDALGAR') {
    targetSheetName = 'BneidALGarSubmited';
} else if (regionUpper === 'FARWANIYACENTER') {
    targetSheetName = 'FarwaniyaCenterSubmited';
} else if (regionUpper === 'JABERCENTER') {
    targetSheetName = 'JaberCenterSubmited';
} else if (regionUpper === 'JAHRACENTER') {
    targetSheetName = 'JahraCenterSubmited';
} else if (regionUpper === 'NASSERALSAEED') {
    targetSheetName = 'NasserAlSaeedSubmited';
} else if (regionUpper === 'NEWJAHRADENTALCENTER') {
    targetSheetName = 'NewJahraDentalCenterSubmited';
} else if (regionUpper === 'SPECIALIZEDCENTER') {
    targetSheetName = 'SpecializedCenterSubmited';
} else if (regionUpper === 'AHMADIPROGRAM') {
    targetSheetName = 'AhmadiProgramSubmited';
} else if (regionUpper === 'CAPITALPROGRAM') {
    targetSheetName = 'CapitalProgramSubmited';
} else if (regionUpper === 'HAWALLYPROGRAM') {
    targetSheetName = 'HawallyProgramSubmited';
} else if (regionUpper === 'FARWANIYAPROGRAM') {
    targetSheetName = 'FarwaniyaProgramSubmited';
} else if (regionUpper === 'JAHRAPROGRAM') {
    targetSheetName = 'JahraProgramSubmited';
} else if (regionUpper === 'MUBARAKALKABEERPROGRAM') {
    targetSheetName = 'MubarakALKabeerProgramSubmited';
} else if (regionUpper === 'ALSABAHHOSPITAL') {
    targetSheetName = 'ALSabahHospitalSubmited';
  }

  var targetSheet = ss.getSheetByName(targetSheetName);

  if (!targetSheet) {
    Logger.log('Target sheet not found: ' + targetSheetName);
    return;
  }

  var lastRowTarget = targetSheet.getLastRow();
  var nextRow = lastRowTarget + 1;

  var valuesToWrite = [
    '',
    '',
    '',
    generatedCode,
    'Sent',
    formData.supplier || '',
    timestamp
  ];

  targetSheet.getRange(nextRow, 1, 1, valuesToWrite.length).setValues([valuesToWrite]);

    let notiTimestamp = new Date();
    let notiSheetHRAdmin = ss.getSheetByName('Notification HRAdmin');
    let notiSheetCompany = ss.getSheetByName('Notification Company');
    let notiSheetAdmin = ss.getSheetByName('Notification Admin');
    let notiSheetTML = ss.getSheetByName('Notification TML');
    let notiSheetAdminS = ss.getSheetByName('Notification AdminS');

    if(notiSheetHRAdmin) notiSheetHRAdmin.appendRow([notiTimestamp, "Data from Dental WebApp"]);
    if(notiSheetCompany) notiSheetCompany.appendRow([notiTimestamp, "Data from Dental WebApp"]);
    if(notiSheetAdmin) notiSheetAdmin.appendRow([notiTimestamp, "Data from Dental WebApp"]);
    if(notiSheetTML) notiSheetTML.appendRow([notiTimestamp, "Data from Dental WebApp"]);
    if(notiSheetAdminS) notiSheetAdminS.appendRow([notiTimestamp, "Data from Dental WebApp"]);

  var supplier = formData.supplier;
  if (supplier) {
    var emailSheet = ss.getSheetByName('CompaniesEmails');
    if (emailSheet) {
      var lastRowEmail = emailSheet.getLastRow();
      if (lastRowEmail > 1) {
        var range = emailSheet.getRange(2, 1, lastRowEmail - 1, 3); // Columns A:C
        var values = range.getValues();
        for (var i = 0; i < values.length; i++) {
          if (values[i][0].toString().trim() === supplier.toString().trim()) {
            var supplierEmail = values[i][1];
            var supplierPhone = values[i][2]; // القيمة من العمود C

            // --- أسطر إرسال الواتساب المضافة ---
            if (supplierPhone) {
              sendWhatsAppUltra(formData, responseNumber, supplierPhone, timestamp);
            }
            if (GROUP_ID && GROUP_ID !== '') {
              sendWhatsAppUltra(formData, responseNumber, GROUP_ID, timestamp);
            }
            // ---------------------------------

            if (supplierEmail && supplierEmail.toString().trim() !== '') {
              var htmlBody = `
                <table border="1" style="border-collapse: collapse; width: 100%;">
                  <tr style="background-color: #f2f2f2;">
                    <th style="padding: 8px; text-align: left;">Field</th>
                    <th style="padding: 8px; text-align: left;">Value</th>
                  </tr>
                  <tr><td style="padding: 8px;">Response Number</td><td style="padding: 8px;">${responseNumber}</td></tr>
                  <tr><td style="padding: 8px;">Timestamp</td><td style="padding: 8px;">${timestamp}</td></tr>
                  <tr><td style="padding: 8px;">Region</td><td style="padding: 8px;">${formData.region || ''}</td></tr>
                  <tr><td style="padding: 8px;">Clinic</td><td style="padding: 8px;">${formData.clinic || ''}</td></tr>
                  <tr><td style="padding: 8px;">Room</td><td style="padding: 8px;">${formData.room || ''}</td></tr>
                  <tr><td style="padding: 8px;">Name</td><td style="padding: 8px;">${formData.name || ''}</td></tr>
                  <tr><td style="padding: 8px;">Manufacturer</td><td style="padding: 8px;">${formData.manufacturer || ''}</td></tr>
                  <tr><td style="padding: 8px;">Model</td><td style="padding: 8px;">${formData.model || ''}</td></tr>
                  <tr><td style="padding: 8px;">Serial</td><td style="padding: 8px;">${formData.serial || ''}</td></tr>
                  <tr><td style="padding: 8px;">Supplier</td><td style="padding: 8px;">${formData.supplier || ''}</td></tr>
                  <tr><td style="padding: 8px;">Reporter Name</td><td style="padding: 8px;">${formData.reporterName || ''}</td></tr>
                  <tr><td style="padding: 8px;">Phone</td><td style="padding: 8px;">${formData.phone || ''}</td></tr>
                  <tr><td style="padding: 8px;">Malfunction</td><td style="padding: 8px;">${formData.additionalField || ''}</td></tr>
                </table>
              `;
              GmailApp.sendEmail(supplierEmail, 'Notification From Dental Equipments WebApp', '', {
                htmlBody: htmlBody
              });

              var newRemaining = MailApp.getRemainingDailyQuota();
              emailSheet.getRange('D2').setValue('Your remain daily quota: ' + newRemaining);
              break;
            }
          }
        }
      }
    }
  }

rebuildDataCachePartial(targetSheetName);
return "Success";
}

// --- دالة إرسال الواتساب الكاملة المضافة ---
function sendWhatsAppUltra(formData, responseNumber, destination, time) {
  var url = "https://api.ultramsg.com/" + INSTANCE_ID + "/messages/chat";
  var cleanTo = destination.toString().includes('@g.us') ? destination : destination.toString().replace(/[^\d]/g, '');

  var message = "🦷 *New Maintenance Request* 🦷\n\n" +
                "🎫 *Request No.:* " + responseNumber + "\n" +
                "📅 *Timestamp:* " + time + "\n" +
                "--------------------------\n" +
                "📍 *Governorate:* " + formData.region + "\n" +
                "🏥 *Center:* " + formData.clinic + "\n" +
                "🚪 *Clinic:* " + formData.room + "\n" +
                "⚙️ *Name:* " + formData.name + "\n" +
                "🏭 *Manufacturer:* " + formData.manufacturer + "\n" +
                "📦 *Model:* " + formData.model + "\n" +
                "🔢 *Serial:* " + formData.serial + "\n" +
                "⚠️ *Malfunction:* " + formData.additionalField + "\n" +
                "--------------------------\n" +
                "👤 *Reporter Name:* " + formData.reporterName + "\n" +
                "📞 *Phone:* " + formData.phone;

  var payload = { "token": TOKEN, "to": cleanTo, "body": message };
  UrlFetchApp.fetch(url, { "method": "post", "payload": payload, "muteHttpExceptions": true });
}
// -----------------------------------------

function rebuildDataCache() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  let cacheSheet = ss.getSheetByName(DATACACHE_SHEET);
  if (!cacheSheet) {
    cacheSheet = ss.insertSheet(DATACACHE_SHEET);
    Logger.log('New DataCache Sheet created');
  }

  const allRows = [];
  const MAX_COL = 41;

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

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const MAX_COL = 41;

  let cacheSheet = ss.getSheetByName(DATACACHE_SHEET);
  if (!cacheSheet || cacheSheet.getLastRow() < 1) {
    Logger.log('Cache data does not exist/is empty — full rebuild');
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

  Logger.log('DataCache partial rebuild: ' + newSourceRows.length + ' line of ' + sourceSheetName + ', total cache: ' + normalized.length + ', ts=' + nowEpoch);
  return nowEpoch;
}
