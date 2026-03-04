// ────────── CONFIG ──────────
var FILE_NAME      = 'GSADUs Catalog_Registry.csv';   // CSV filename in Drive
var CATALOG_SHEET  = 'Catalog';          // sheet tab with your ADU_Catalog data
var KEY_COLUMN     = 'Model';            // unique key column header
var HEADER_ROW     = 2;                  // row number where headers live
// ────────────────────────────

/**
 * 1) Syncs the ADU_Catalog sheet to the CSV by:
 *    • Appending only truly new rows
 *    • Updating rows with changed values
 *    • Counting unchanged rows
 *    • Reporting rows in the sheet missing from CSV
 *    Uses batch read/write for speed and full-row comparison
 */
function updateCatalogFromCsv() {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var sheet    = ss.getSheetByName(CATALOG_SHEET);
  if (!sheet) throw new Error('Sheet "' + CATALOG_SHEET + '" not found');

  var lastCol  = sheet.getLastColumn();
  var lastRow  = sheet.getLastRow();
  if (lastRow <= HEADER_ROW) {
    SpreadsheetApp.getUi().alert('No data found below header row ' + HEADER_ROW);
    return;
  }

  var numRows   = lastRow - HEADER_ROW + 1;
  var rangeA1   = sheet.getRange(HEADER_ROW, 1, numRows, lastCol);
  var sheetVals = rangeA1.getValues();
  var headers   = sheetVals[0];
  var dataRows  = sheetVals.slice(1);

  // normalize sheet headers (trim and remove BOM if any)
  headers = headers.map(function(h){
    return (h===null || h===undefined) ? '' : String(h).replace(/^\uFEFF/, '').replace(/^"+|"+$/g, '').trim();
  });

  var headerIdx = {};
  headers.forEach(function(h,i){ headerIdx[h] = i; });
  if (headerIdx[KEY_COLUMN] === undefined) {
    throw new Error('Key column "' + KEY_COLUMN + '" not found in headers');
  }

  var existingMap = {};
  dataRows.forEach(function(r,i){
    var key = normalizeKey(r[headerIdx[KEY_COLUMN]]);
    if (key) existingMap[key] = i;
  });

  var files = DriveApp.getFilesByName(FILE_NAME);
  if (!files.hasNext()) throw new Error('CSV file "' + FILE_NAME + '" not found');
  var file = files.next();
  var raw = file.getBlob().getDataAsString();
  // log a short raw preview to help diagnose BOM/hidden chars
  Logger.log('Raw CSV start (200 chars): %s', raw.slice(0,200));
  var csvAll = Utilities.parseCsv(raw);
  var csvHdr = csvAll.shift();
  // normalize csv headers (trim and remove BOM if any)
  csvHdr = csvHdr.map(function(h){
    return (h===null || h===undefined) ? '' : String(h).replace(/^\uFEFF/, '').replace(/^"+|"+$/g, '').trim();
  });
  // diagnostic logs / quick UI alert to show exactly what headers were parsed
  Logger.log('Parsed csvHdr: %s', JSON.stringify(csvHdr));
  Logger.log('CSV rows (excluding header): %s', csvAll.length);
  if (csvAll.length) Logger.log('First CSV data row: %s', JSON.stringify(csvAll[0]));
  try {
    SpreadsheetApp.getUi().alert('CSV headers detected:\n' + csvHdr.join(', '));
  } catch (e) {
    // UI may not be available in some contexts; ignore
  }
  var csvIdx = {};
  csvHdr.forEach(function(h,i){ csvIdx[h] = i; });
  if (csvIdx[KEY_COLUMN] === undefined) {
    throw new Error('Key column "' + KEY_COLUMN + '" not found in CSV');
  }

  var countNew=0, countUpdated=0, countUnchanged=0;
  var matchedKeys={}, sheetOut=sheetVals.slice();

  csvAll.forEach(function(r){
    var key = normalizeKey(r[csvIdx[KEY_COLUMN]]);
    if (!key) return;
    var idx = existingMap[key];
    var rowData = headers.map(function(h){
      return (csvIdx[h]!==undefined) ? r[csvIdx[h]] : '';
    });
    if (idx!==undefined) {
      var existingRow = dataRows[idx].map(String);
      var newRow      = rowData.map(String);
      if (JSON.stringify(existingRow) === JSON.stringify(newRow)) {
        countUnchanged++;
      } else {
        sheetOut[idx+1] = rowData;
        countUpdated++;
      }
      matchedKeys[key]=true;
    } else {
      sheetOut.push(rowData);
      countNew++;
      matchedKeys[key]=true;
    }
  });

  var countMissing = Object.keys(existingMap).reduce(function(acc,k){
    return acc + (!matchedKeys[k]?1:0);
  },0);

  sheet.getRange(HEADER_ROW,1,sheetOut.length,lastCol).setValues(sheetOut);

  SpreadsheetApp.getUi().alert(
    'New entries: '       + countNew       + '\n' +
    'Updated entries: '   + countUpdated   + '\n' +
    'Unchanged entries: ' + countUnchanged + '\n' +
    'Missing entries: '   + countMissing
  );
}

/**
 * Sets up a daily trigger at 2 AM to auto-run the CSV import.
 */
function scheduleDailyCsvUpdate() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'updateCatalogFromCsv') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('updateCatalogFromCsv')
           .timeBased()
           .everyDays(1)
           .atHour(2)
           .create();
}

/**
 * Adds the custom menu on spreadsheet open.
 */
function onOpen(){
  SpreadsheetApp.getUi()
    .createMenu('ADU Catalog')
    .addItem('Update from CSV','updateCatalogFromCsv')
    .addItem('Schedule daily CSV update','scheduleDailyCsvUpdate')
    .addToUi();
}

function normalizeKey(v){ return (v===null || v===undefined) ? '' : String(v).trim(); }