// Multi-CSV import dispatcher + generic importer for Google Sheets
// Paste into Apps Script editor. Edit IMPORT_MAP for your CSV→sheet mappings.

// ------------------ CONFIG: map CSV files to target sheets ------------------
var IMPORT_MAP = [
  {
    // Registry CSV -> Catalog sheet (example)
    fileName: 'GSADUs Catalog_Registry.csv',
    sheetName: 'Catalog',
    headerRow: 2,           // row index (1-based) where headers live in the sheet
    keyColumn: 'Model',     // header name used as unique key (must exist in CSV or map)
    headerMap: null,        // optional { 'CSV Name' : 'Sheet Header' } mapping
    mode: 'sync'            // 'sync' (default) or 'replace' - replace writes CSV rows wholesale
  },
  {
    // Elements CSV -> MG_Elem sheet (example)
    fileName: 'GSADUs Catalog_Elements.csv',
    sheetName: 'MG_Elem',
    headerRow: 1,
    keyColumn: 'ModelGroup', // adjust to match your MG_Elem sheet header
    headerMap: null,         // specify if CSV header names differ from sheet
    mode: 'replace'          // large CSV: replace entire target table (faster for big files)
  }
];
// ---------------------------------------------------------------------------

// Top-level dispatcher you can assign to a button or call manually.
// dryRun true => no writes, just return summaries
function runAllImports(dryRun) {
  var results = [];
  var start = new Date();

  for (var i = 0; i < IMPORT_MAP.length; i++) {
    var m = IMPORT_MAP[i];
    try {
      var res = importCsvToSheet(m, { dryRun: !!dryRun, normalizeCase: true, pickLatest: true });
      results.push({ name: m.fileName, sheet: m.sheetName, ok: true, result: res });
    } catch (e) {
      // Catch anything unexpected to keep dispatcher resilient
      Logger.log('Importer for %s failed: %s', m.fileName, e.stack || e.message || e);
      results.push({ name: m.fileName, sheet: m.sheetName, ok: false, error: String(e) });
    }
  }

  var duration = (new Date() - start) / 1000;
  // compact UI summary
  var lines = [];
  results.forEach(function(r){
    if (!r.ok) {
      lines.push(r.name + ' -> ' + r.sheet + ': ERROR - ' + r.error);
      return;
    }
    if (r.result.skipped) {
      lines.push(r.name + ' -> ' + r.sheet + ': SKIPPED - ' + r.result.reason);
    } else {
      lines.push(r.name + ' -> ' + r.sheet + ': new=' + r.result.new + ', updated=' + r.result.updated + ', unchanged=' + r.result.unchanged + ', missing=' + (r.result.missing||0));
    }
  });
  lines.push('Total time: ' + duration + 's');

  try {
    SpreadsheetApp.getUi().alert('Import summary:\n' + lines.join('\n'));
  } catch (e) { /* ignore in non-interactive runs */ }

  // persistent log
  writeImportLog(results, duration);

  return results;
}

// Generic importer: imports one CSV to one sheet according to mapping
// mapping: { fileName, sheetName, headerRow, keyColumn, headerMap (optional) }
// opts: { dryRun:boolean, normalizeCase:boolean, pickLatest:boolean }
function importCsvToSheet(mapping, opts) {
  opts = opts || {};
  var dryRun = !!opts.dryRun;
  var normalizeCase = !!opts.normalizeCase;
  var pickLatest = opts.pickLatest !== false; // default true

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(mapping.sheetName);
  if (!sh) {
    return { skipped: true, reason: 'Sheet not found: ' + mapping.sheetName };
  }

  // Read sheet headers
  var headerRow = mapping.headerRow || 1;
  var lastCol = sh.getLastColumn();
  var lastRow = sh.getLastRow();
  if (lastCol < 1) {
    return { skipped: true, reason: 'Sheet has no columns: ' + mapping.sheetName };
  }
  var numRows = Math.max(1, lastRow - headerRow + 1);
  var range = sh.getRange(headerRow, 1, numRows, lastCol);
  var sheetVals = range.getValues();
  var sheetHeaders = sheetVals[0];

  // Normalize sheet headers (strip BOM, surrounding quotes, trim)
  sheetHeaders = sheetHeaders.map(function(h){
    return (h === null || h === undefined) ? '' : String(h).replace(/^\uFEFF/, '').replace(/^"+|"+$/g, '').trim();
  });

  // Build header map for quick index by sheet header name
  var headerIdx = {};
  for (var i = 0; i < sheetHeaders.length; i++) headerIdx[sheetHeaders[i]] = i;

  // Validate key column exists in target sheet headers (or accept if headerMap will provide mapping)
  var keyColumn = mapping.keyColumn;
  var headerMap = mapping.headerMap || null;

  // Find CSV file in Drive (pick latest if duplicates)
  var files = DriveApp.getFilesByName(mapping.fileName);
  if (!files.hasNext()) {
    return { skipped: true, reason: 'CSV file not found: ' + mapping.fileName };
  }
  var file = pickLatest ? pickLatestFile(files) : files.next();
  if (!file) return { skipped: true, reason: 'CSV file not found (pickLatest failed): ' + mapping.fileName };

  var raw = file.getBlob().getDataAsString();
  // Parse CSV (simple comma-based parse)
  var csvAll = Utilities.parseCsv(raw);
  if (!csvAll || csvAll.length < 1) return { skipped: true, reason: 'CSV empty or parse failed: ' + mapping.fileName };

  // Extract and normalize CSV headers
  var csvHdr = csvAll.shift();
  csvHdr = csvHdr.map(function(h){
    return (h === null || h === undefined) ? '' : String(h).replace(/^\uFEFF/, '').replace(/^"+|"+$/g, '').trim();
  });

  // Optionally build a csv->sheet header name mapping using headerMap
  // Effective sheet header name for a given CSV header:
  //   - if headerMap provided and headerMap[csvHeader] exists -> use that
  //   - else if the csvHeader matches a sheet header -> use csvHeader
  //   - else if a sheet header exactly matches after normalization -> use that
  var csvToSheetHeader = {};
  for (var ci = 0; ci < csvHdr.length; ci++) {
    var ch = csvHdr[ci];
    var mapped = null;
    if (headerMap && headerMap[ch]) mapped = headerMap[ch];
    else if (headerIdx.hasOwnProperty(ch)) mapped = ch;
    // last resort: try case-insensitive match
    else {
      for (var shn in headerIdx) {
        if (shn && ch && shn.toLowerCase() === ch.toLowerCase()) { mapped = shn; break; }
      }
    }
    if (mapped) csvToSheetHeader[ch] = mapped;
    else csvToSheetHeader[ch] = null; // CSV column will be ignored unless sheet has a same-named header later
  }

  // Build csvIdx: index of CSV column for each effective sheet header
  var csvIdx = {};
  for (var k = 0; k < sheetHeaders.length; k++) {
    csvIdx[sheetHeaders[k]] = undefined;
  }
  for (var j = 0; j < csvHdr.length; j++) {
    var csvH = csvHdr[j];
    var sheetH = csvToSheetHeader[csvH];
    if (sheetH && headerIdx.hasOwnProperty(sheetH)) {
      csvIdx[sheetH] = j;
    }
  }

  // If mapping.mode === 'replace', perform a wholesale replace of the target table
  if (mapping.mode === 'replace') {
    // Build rows mapped to sheetHeaders order
    var rowsMapped = csvAll.map(function(csvRow) {
      return sheetHeaders.map(function(shName) {
        var cidx = csvIdx[shName];
        return (cidx !== undefined && cidx !== void 0) ? (csvRow[cidx] || '') : '';
      });
    });

    var sheetOutReplace = [sheetHeaders].concat(rowsMapped);
    var countNewR = rowsMapped.length;
    var countUpdatedR = 0, countUnchangedR = 0, countMissingR = 0;

    if (!dryRun) {
      // write starting at headerRow; resize to number of rows we have
      sh.getRange(headerRow, 1, sheetOutReplace.length, lastCol).setValues(sheetOutReplace);
      // if sheet previously had more rows below, optionally clear the remaining rows
      var prevTotalRows = sh.getMaxRows();
      var writtenRows = headerRow - 1 + sheetOutReplace.length;
      if (prevTotalRows > writtenRows) {
        // clear leftover area (optional: leave as-is to preserve sheet size)
        var leftoverRows = prevTotalRows - writtenRows;
        sh.getRange(writtenRows + 1, 1, leftoverRows, lastCol).clearContent();
      }
    }

    return { new: countNewR, updated: countUpdatedR, unchanged: countUnchangedR, missing: countMissingR, skipped: false };
  }

  // Ensure we can find the key column index in either sheetHeaders or headerMap
  var effectiveKeyHeader = keyColumn;
  if (headerMap) {
    // If a headerMap maps SOME CSV header to the sheet keyColumn, we're ok.
    // Otherwise, check if keyColumn exists in sheetHeaders.
    if (!headerIdx.hasOwnProperty(keyColumn)) {
      return { skipped: true, reason: 'Key column not found in sheet headers: ' + keyColumn };
    }
  } else {
    if (!headerIdx.hasOwnProperty(keyColumn)) {
      // try case-insensitive in sheet headers
      var found = null;
      for (var shn2 in headerIdx) {
        if (shn2 && shn2.toLowerCase() === keyColumn.toLowerCase()) { found = shn2; break; }
      }
      if (found) effectiveKeyHeader = found;
      else return { skipped: true, reason: 'Key column not found in sheet headers: ' + keyColumn };
    }
  }

  // Build existingMap from sheet data rows (normalize keys)
  var sheetDataRows = sheetVals.slice(1); // rows under headerRow
  var existingMap = {};
  for (var r = 0; r < sheetDataRows.length; r++) {
    var row = sheetDataRows[r];
    var keyVal = row[ headerIdx[effectiveKeyHeader] ];
    var nk = normalizeKey(keyVal, normalizeCase);
    if (nk) existingMap[nk] = r;
  }

  // Iterate CSV rows, create new/updated logic
  var countNew = 0, countUpdated = 0, countUnchanged = 0;
  var matchedKeys = {};
  var sheetOut = sheetVals.slice(); // copy of the section we're going to write back

  for (var rr = 0; rr < csvAll.length; rr++) {
    var csvRow = csvAll[rr];
    // Resolve CSV key value: find which CSV column maps to the sheet key header
    var csvKeyIdx = -1;
    // if headerMap pointed some csv header to the sheet key, prefer that
    for (var chdr in csvToSheetHeader) {
      if (csvToSheetHeader[chdr] === effectiveKeyHeader) {
        // find index j where csvHdr[j] === chdr
        for (var j2 = 0; j2 < csvHdr.length; j2++) if (csvHdr[j2] === chdr) { csvKeyIdx = j2; break; }
      }
    }
    // fallback: if csv column with same name as sheet key exists
    if (csvKeyIdx === -1 && csvIdx[effectiveKeyHeader] !== undefined) csvKeyIdx = csvIdx[effectiveKeyHeader];

    var csvKeyRaw = (csvKeyIdx >= 0) ? csvRow[csvKeyIdx] : '';
    var keyNormalized = normalizeKey(csvKeyRaw, normalizeCase);
    if (!keyNormalized) continue; // skip rows without key

    var existingIdx = (existingMap.hasOwnProperty(keyNormalized) ? existingMap[keyNormalized] : undefined);

    // build rowData using sheetHeaders order; pick value from CSV if csvIdx[sheetHeader] is set
    var rowData = [];
    for (var hi = 0; hi < sheetHeaders.length; hi++) {
      var shName = sheetHeaders[hi];
      var cidx = csvIdx[shName];
      rowData.push( (cidx !== undefined && cidx !== void 0) ? (csvRow[cidx] || '') : '' );
    }

    if (existingIdx !== undefined) {
      var existingRow = sheetDataRows[existingIdx].map(String);
      var newRow = rowData.map(String);
      if (JSON.stringify(existingRow) === JSON.stringify(newRow)) {
        countUnchanged++;
      } else {
        sheetOut[existingIdx + 1] = rowData;
        countUpdated++;
      }
      matchedKeys[keyNormalized] = true;
    } else {
      sheetOut.push(rowData);
      countNew++;
      matchedKeys[keyNormalized] = true;
    }
  }

  // count missing (present in sheet but not in CSV)
  var countMissing = 0;
  for (var kkey in existingMap) {
    if (!matchedKeys[kkey]) countMissing++;
  }

  // Write back unless dryRun
  if (!dryRun) {
    // write from headerRow, col 1, sheetOut.length rows, lastCol columns
    sh.getRange(headerRow, 1, sheetOut.length, lastCol).setValues(sheetOut);
  }

  return { new: countNew, updated: countUpdated, unchanged: countUnchanged, missing: countMissing, skipped:false };
}

// Helper: pick latest file from Files iterator (by LastUpdated or DateCreated)
function pickLatestFile(filesIterator) {
  var chosen = null;
  var chosenDate = 0;
  while (filesIterator.hasNext()) {
    var f = filesIterator.next();
    var d = 0;
    try { d = f.getLastUpdated().getTime(); } catch (e) {
      try { d = f.getDateCreated().getTime(); } catch (ee) { d = 0; }
    }
    if (d >= chosenDate) { chosen = f; chosenDate = d; }
  }
  return chosen;
}

// Normalize key values - trim and optionally lower-case
function normalizeKey(v, lower) {
  if (v === null || v === undefined) return '';
  var s = String(v).replace(/^\uFEFF/, '').trim();
  return lower ? s.toLowerCase() : s;
}

  /**
   * Run a single mapping by index (0-based) — useful for debugging long runs.
   * dryRun true => no writes. Returns the importer result and logs timing.
   */
  function runImportByIndex(idx, dryRun) {
    if (typeof idx !== 'number') {
      Logger.log('runImportByIndex: idx must be numeric (0-based). Got: %s', idx);
      return { error: 'invalid index' };
    }
    var mapping = IMPORT_MAP[idx];
    if (!mapping) {
      Logger.log('runImportByIndex: no mapping at index %s', idx);
      return { error: 'mapping not found' };
    }
    var t0 = new Date();
    try {
      var res = importCsvToSheet(mapping, { dryRun: !!dryRun, normalizeCase: true, pickLatest: true });
      var dt = (new Date() - t0) / 1000;
      Logger.log('runImportByIndex: %s -> %s finished in %s s; result: %s', mapping.fileName, mapping.sheetName, dt, JSON.stringify(res));
      return res;
    } catch (e) {
      var dt2 = (new Date() - t0) / 1000;
      Logger.log('runImportByIndex ERROR: %s after %s s: %s', mapping.fileName, dt2, e.stack || e.message || e);
      return { error: String(e) };
    }
  }

  /**
   * Run a single mapping by CSV fileName (convenience wrapper).
   */
  function runImportByFileName(fileName, dryRun) {
    for (var i = 0; i < IMPORT_MAP.length; i++) {
      if (IMPORT_MAP[i].fileName === fileName) return runImportByIndex(i, dryRun);
    }
    Logger.log('runImportByFileName: mapping not found for %s', fileName);
    return { error: 'mapping not found' };
  }

// Append an execution summary to a "CSV Import Log" sheet
function writeImportLog(results, seconds) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logName = 'CSV Import Log';
  var sh = ss.getSheetByName(logName);
  if (!sh) {
    sh = ss.insertSheet(logName);
    sh.appendRow(['Timestamp','Duration_s','Summary']);
  }
  var ts = new Date();
  sh.appendRow([ ts, seconds, JSON.stringify(results) ]);
}