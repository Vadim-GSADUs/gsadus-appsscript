/**
 * GSADUs Tools (V10)
 *
 * Columns are resolved dynamically from row 1 headers — the script does NOT
 * depend on a fixed column order. Add, remove, or reorder columns freely;
 * as long as row 1 header names match, everything continues to work.
 *
 * Required headers in the Supplier tab:
 *   Design_Bundle, Category, Supplier_URL, Supplier, Product_Name,
 *   Product_Size, File_ID, Drive_URL, Filename, Sync_Status
 *
 * Optional headers (read by MoodBoard.js, ignored here):
 *   VScale, HScale
 */

// ── Constants ────────────────────────────────────────────────────────────────

const MATERIALS_FOLDER_ID = '1hc2moJgK51YPqYxcmm_Zgry5YxbsbGAs';
const TEMPLATE_ID         = '1oGLgK-aCvKVh1EIhADQsqeqWQLlUaCTo4AkmtAY9dU4';
const BUNDLES_FOLDER_ID   = '1v7vLPjvPdMA42wGA9XqC_29DNtZP21Gk'; // Interior Design Bundles folder
const BUNDLES_JSON_NAME   = 'bundles_library.json';

// ── Menu ─────────────────────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('GSADUs Tools')
    .addItem('1. Pull from Order Template',  'pullFromOrderTemplate')
    .addItem('2. Sync Material Assets',      'syncMaterialAssets')
    .addItem('3. Export to JSON',            'exportToJson')
    .addSeparator()
    .addItem('Compute Missing Scales',       'computeMissingScales')
    .addItem('Audit Materials Folder',       'auditMaterialsFolder')
    .addItem('Format Active Sheet',          'formatActiveSheetColumns')
    .addToUi();
}

// ── STEP 1 ───────────────────────────────────────────────────────────────────

/**
 * Reads the ORDER TEMPLATE > Bundles tab and writes hyperlinked product names
 * into the Supplier_URL column as =HYPERLINK() formulas.
 *
 * Source layout:
 *   Row 1 : bundle name headers (C=Subway, F=Harbor, I=Navy, L=Olive, O=Antique, R=Villa)
 *   Col A : category labels (rows 2–7): Flooring, Bathroom Floor Tile, etc.
 */
function pullFromOrderTemplate() {
  const ss            = SpreadsheetApp.getActiveSpreadsheet();
  const supplierSheet = ss.getSheetByName('Supplier');
  const remoteSheet   = SpreadsheetApp.openById(TEMPLATE_ID).getSheetByName('Bundles');

  const colMap = getColMap_(supplierSheet);
  if (!validateCols_(colMap, ['Supplier_URL'], 'Step 1')) return;

  // Read remote sheet in one batch
  const numRows   = remoteSheet.getLastRow();
  const numCols   = remoteSheet.getLastColumn();
  const fullRange = remoteSheet.getRange(1, 1, numRows, numCols);
  const values    = fullRange.getValues();
  const formulas  = fullRange.getFormulas();
  const richText  = fullRange.getRichTextValues();

  // Bundle name → 0-based column index (header row 1 of remote sheet)
  const bundleColMap = {};
  values[0].forEach((v, i) => {
    const name = String(v).trim().toUpperCase();
    if (name) bundleColMap[name] = i;
  });

  // Category label → 0-based row index (col A, rows 2+)
  const catRowMap = {};
  for (let r = 1; r < values.length; r++) {
    const cat = String(values[r][0]).trim().toUpperCase();
    if (cat && !catRowMap[cat]) catRowMap[cat] = r;
  }

  const CATEGORIES = [
    'Flooring', 'Bathroom Floor Tile', 'Shower Wall Tile',
    'Shower Pan Tile', 'Kitchen Backsplash', 'Cabinet Color'
  ];
  const BUNDLES = ['Subway', 'Harbor', 'Navy', 'Olive', 'Antique', 'Villa'];

  // Existing Supplier rows (data only, skip header) → "BUNDLE|CATEGORY" → 1-based sheet row
  const lastRow = Math.max(1, supplierSheet.getLastRow());
  const rowMap  = {};
  if (lastRow > 1) {
    supplierSheet.getRange(2, 1, lastRow - 1, 2).getValues().forEach((row, i) => {
      const k = `${String(row[0]).trim().toUpperCase()}|${String(row[1]).trim().toUpperCase()}`;
      if (k !== '|') rowMap[k] = i + 2; // +2: 1-based + skip header row
    });
  }

  const supplierUrlCol = colMap['Supplier_URL'] + 1; // 1-based for getRange
  let created = 0, written = 0;

  BUNDLES.forEach(bundle => {
    const bKey   = bundle.toUpperCase();
    const colIdx = bundleColMap[bKey];

    CATEGORIES.forEach(cat => {
      const cKey   = cat.toUpperCase();
      const mapKey = `${bKey}|${cKey}`;

      // Ensure row exists
      let sheetRow = rowMap[mapKey];
      if (!sheetRow) {
        supplierSheet.appendRow([bundle, cat]);
        sheetRow = supplierSheet.getLastRow();
        rowMap[mapKey] = sheetRow;
        created++;
      }

      if (colIdx === undefined) return;
      const rIdx = catRowMap[cKey];
      if (rIdx === undefined) return;

      const text = String(values[rIdx][colIdx]).trim();
      if (!text) return;

      // Extract URL — try rich text first, then formula
      let url = null;
      const rt = richText[rIdx][colIdx];
      if (rt) {
        url = rt.getLinkUrl();
        if (!url) {
          const runs = rt.getRuns();
          for (let i = 0; i < runs.length; i++) {
            const u = runs[i].getLinkUrl();
            if (u) { url = u; break; }
          }
        }
      }
      if (!url) {
        const f = formulas[rIdx][colIdx];
        if (f) {
          const m = f.match(/HYPERLINK\("([^"]+)"/i);
          if (m) url = m[1];
        }
      }

      const cell = supplierSheet.getRange(sheetRow, supplierUrlCol);
      url ? cell.setFormula(`=HYPERLINK("${url}","${text.replace(/"/g, '""')}")`)
          : cell.setValue(text);
      written++;
    });
  });

  ss.toast(
    `Created ${created} rows, wrote ${written} Supplier_URL cells.`,
    'Step 1 Complete', 8
  );
}

// ── STEP 2 ───────────────────────────────────────────────────────────────────

/**
 * For each Supplier row where Supplier + Product_Name are filled:
 *
 *   1. Computes canonical filename: Supplier_Product-Name.ext
 *   2. Dedup check — if another row already processed the same Supplier+Product,
 *      reuses that Drive file (no duplicate files in Materials folder).
 *   3. Otherwise resolves the file from Drive_URL, which accepts:
 *        a. Google Drive URL  → https://drive.google.com/file/d/FILE_ID/...
 *        b. Windows path      → G:\Shared drives\...\filename.jpg
 *        c. Relative path     → Materials\filename.jpg
 *      Renames the file to canonical name, moves it into Materials\ if needed.
 *   4. Writes File_ID, Drive_URL (canonical), Filename, Sync_Status columns.
 *
 * CRITICAL: Supplier_URL contains =HYPERLINK() formulas. getValues() returns
 * only the display text — never write that value back to the cell or the
 * formula will be destroyed. This script only reads Supplier_URL display text
 * for dedup matching; it never writes back to that column.
 */
function syncMaterialAssets() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sheet   = ss.getSheetByName('Supplier');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { ss.toast('No data rows found.', 'Step 2'); return; }

  const colMap = getColMap_(sheet);
  const REQUIRED = ['Supplier', 'Product_Name', 'Supplier_URL', 'File_ID', 'Drive_URL', 'Filename', 'Sync_Status'];
  if (!validateCols_(colMap, REQUIRED, 'Step 2')) return;

  const numDataRows = lastRow - 1;
  const numCols     = sheet.getLastColumn();

  // Read all columns at once.
  // Supplier_URL (col C or wherever) returns display text via getValues() — safe to
  // read for dedup matching, never written back.
  const values = sheet.getRange(2, 1, numDataRows, numCols).getValues();

  // Resolve 1-based column positions for output writes
  const FILE_ID_COL     = colMap['File_ID']     + 1;
  const DRIVE_URL_COL   = colMap['Drive_URL']   + 1;
  const FILENAME_COL    = colMap['Filename']     + 1;
  const SYNC_STATUS_COL = colMap['Sync_Status'] + 1;

  // Seed output arrays from existing values (unchanged rows keep their data)
  const outFileId     = values.map(row => [String(row[colMap['File_ID']])    .trim()]);
  const outDriveUrl   = values.map(row => [String(row[colMap['Drive_URL']])  .trim()]);
  const outFilename   = values.map(row => [String(row[colMap['Filename']])   .trim()]);
  const outSyncStatus = Array.from({ length: numDataRows }, () => ['']); // cleared each run

  const materialsFolder = DriveApp.getFolderById(MATERIALS_FOLDER_ID);

  // ── PRE-PASS: propagate existing Drive data to rows sharing the same Supplier_URL display text ──
  const urlFillMap = {};
  for (let i = 0; i < numDataRows; i++) {
    const fileId   = String(values[i][colMap['File_ID']])    .trim();
    const driveUrl = String(values[i][colMap['Drive_URL']])  .trim();
    const filename = String(values[i][colMap['Filename']])   .trim();
    const urlText  = String(values[i][colMap['Supplier_URL']]).trim().toLowerCase();
    if (fileId && urlText && !urlFillMap[urlText]) {
      urlFillMap[urlText] = { fileId, driveUrl, filename };
    }
  }
  let prefilled = 0;
  for (let i = 0; i < numDataRows; i++) {
    if (String(values[i][colMap['File_ID']]).trim()) continue; // already has File_ID
    const urlText = String(values[i][colMap['Supplier_URL']]).trim().toLowerCase();
    if (urlText && urlFillMap[urlText]) {
      const fill = urlFillMap[urlText];
      values[i][colMap['File_ID']]   = fill.fileId;
      values[i][colMap['Drive_URL']] = fill.driveUrl;
      values[i][colMap['Filename']]  = fill.filename;
      prefilled++;
    }
  }

  // Dedup map built progressively — ensures every file is renamed to current canonical convention
  const dedupMap = {};
  let processed = 0, deduped = 0, skipped = 0, errors = 0;

  for (let i = 0; i < numDataRows; i++) {
    const supplier       = String(values[i][colMap['Supplier']])    .trim();
    const product        = String(values[i][colMap['Product_Name']]).trim();
    const existingFileId  = String(values[i][colMap['File_ID']])    .trim();
    const existingDriveUrl = String(values[i][colMap['Drive_URL']]) .trim();

    if (!supplier || !product) { skipped++; continue; }

    const key      = dedupKey(supplier, product);
    const basename = canonicalBasename(supplier, product);

    // Same Supplier+Product already handled this run → propagate
    if (dedupMap[key]) {
      const asset        = dedupMap[key];
      outFileId[i][0]     = asset.fileId;
      outDriveUrl[i][0]   = asset.driveUrl;
      outFilename[i][0]   = asset.filename;
      outSyncStatus[i][0] = 'Synced (shared): ' + timestamp();
      deduped++;
      continue;
    }

    // Resolve file: try Drive_URL first, then File_ID
    let file = null;
    try {
      if (existingDriveUrl) file = resolveFile(existingDriveUrl, materialsFolder);
      if (!file && existingFileId) file = DriveApp.getFileById(existingFileId);
    } catch (e) {
      outSyncStatus[i][0] = '⚠ Resolve error: ' + e.message;
      errors++;
      continue;
    }

    if (!file) {
      outSyncStatus[i][0] = '⚠ No file found — paste a Drive URL or path into the Drive_URL column';
      skipped++;
      continue;
    }

    try {
      const currentName = file.getName();
      const ext         = currentName.includes('.') ? currentName.split('.').pop() : 'jpg';
      const canonical   = `${basename}.${ext}`;

      if (currentName !== canonical) file.setName(canonical);
      ensureInFolder(file, materialsFolder);

      const asset = { fileId: file.getId(), driveUrl: file.getUrl(), filename: canonical };
      dedupMap[key] = asset;

      outFileId[i][0]     = asset.fileId;
      outDriveUrl[i][0]   = asset.driveUrl;
      outFilename[i][0]   = asset.filename;
      outSyncStatus[i][0] = 'Synced: ' + timestamp();
      processed++;

    } catch (e) {
      outSyncStatus[i][0] = '⚠ Error: ' + e.message;
      errors++;
    }
  }

  // Write each output column independently — safe regardless of column order
  sheet.getRange(2, FILE_ID_COL,     numDataRows, 1).setValues(outFileId);
  sheet.getRange(2, DRIVE_URL_COL,   numDataRows, 1).setValues(outDriveUrl);
  sheet.getRange(2, FILENAME_COL,    numDataRows, 1).setValues(outFilename);
  sheet.getRange(2, SYNC_STATUS_COL, numDataRows, 1).setValues(outSyncStatus);

  ss.toast(
    `Processed: ${processed}  |  Shared (dedup): ${deduped}  |  Pre-filled (same URL): ${prefilled}  |  Skipped: ${skipped}  |  Errors: ${errors}`,
    'Step 2 Complete', 12
  );
}

// ── STEP 3 ───────────────────────────────────────────────────────────────────

/**
 * Exports the Supplier sheet to bundles_library.json in the Interior Design
 * Bundles Drive folder. Schema:
 *
 *   { _meta, bundles: [ { name, materials: [ { category, supplier,
 *     product_name, product_size, product_url, drive_file_id,
 *     drive_url, filename } ] } ] }
 *
 * product_url is extracted from the =HYPERLINK() formula in Supplier_URL.
 * sync_status is intentionally omitted — it's internal GSheet bookkeeping.
 * If a file already exists in the target folder it is trashed before writing.
 */
function exportToJson() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sheet   = ss.getSheetByName('Supplier');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { ss.toast('No data rows found.', 'Step 3'); return; }

  const colMap = getColMap_(sheet);
  const REQUIRED = ['Design_Bundle', 'Category', 'Supplier_URL', 'Supplier',
                    'Product_Name', 'Product_Size', 'File_ID', 'Drive_URL', 'Filename'];
  if (!validateCols_(colMap, REQUIRED, 'Step 3')) return;

  const numDataRows = lastRow - 1;
  const numCols     = sheet.getLastColumn();

  const values   = sheet.getRange(2, 1, numDataRows, numCols).getValues();
  const formulas = sheet.getRange(2, colMap['Supplier_URL'] + 1, numDataRows, 1).getFormulas();

  const BUNDLE_ORDER = ['Subway', 'Harbor', 'Navy', 'Olive', 'Antique', 'Villa'];
  const bundleMap    = {};
  BUNDLE_ORDER.forEach(name => { bundleMap[name] = { name, materials: [] }; });

  for (let i = 0; i < numDataRows; i++) {
    const bundle   = String(values[i][colMap['Design_Bundle']]).trim();
    const category = String(values[i][colMap['Category']]).trim();
    if (!bundle || !category) continue;

    // Extract product URL from =HYPERLINK("url","text") formula
    let productUrl = null;
    const formula  = formulas[i][0];
    if (formula) {
      const m = formula.match(/HYPERLINK\("([^"]+)"/i);
      if (m) productUrl = m[1];
    }

    const material = {
      category:      category,
      supplier:      String(values[i][colMap['Supplier']])    .trim() || null,
      product_name:  String(values[i][colMap['Product_Name']]).trim() || null,
      product_size:  String(values[i][colMap['Product_Size']]).trim() || null,
      product_url:   productUrl,
      drive_file_id: String(values[i][colMap['File_ID']])    .trim() || null,
      drive_url:     String(values[i][colMap['Drive_URL']])  .trim() || null,
      filename:      String(values[i][colMap['Filename']])   .trim() || null,
    };

    if (!bundleMap[bundle]) bundleMap[bundle] = { name: bundle, materials: [] };
    bundleMap[bundle].materials.push(material);
  }

  const bundles = BUNDLE_ORDER.filter(n => bundleMap[n]).map(n => bundleMap[n]);
  Object.keys(bundleMap).forEach(n => { if (!BUNDLE_ORDER.includes(n)) bundles.push(bundleMap[n]); });

  const HARDWARE_FINISHES = [
    { name: 'Matte Black',     image_url: null },
    { name: 'Brushed Nickel',  image_url: null },
    { name: 'Champagne Gold',  image_url: null },
    { name: 'Polished Chrome', image_url: null },
  ];

  const output = {
    _meta: {
      last_sync:      new Date().toISOString(),
      source:         ss.getName(),
      schema_version: '1.1',
    },
    hardware: HARDWARE_FINISHES,
    bundles,
  };

  const folder   = DriveApp.getFolderById(BUNDLES_FOLDER_ID);
  const existing = folder.getFilesByName(BUNDLES_JSON_NAME);
  while (existing.hasNext()) existing.next().setTrashed(true);
  folder.createFile(BUNDLES_JSON_NAME, JSON.stringify(output, null, 2), MimeType.PLAIN_TEXT);

  const totalMaterials = bundles.reduce((n, b) => n + b.materials.length, 0);
  ss.toast(
    `Exported ${bundles.length} bundles · ${totalMaterials} materials → ${BUNDLES_JSON_NAME}`,
    'Step 3 Complete', 8
  );
}

// ── Mood Board — moved to Moodboard/MoodBoard.js (gslides-bound script) ──────
// Script ID: 1mAkLKNGRybcALsujtsPHNz4aYxViOEfUygH3ZNYPC9qHZvi9_CLMTViS

// ── Compute Scales ────────────────────────────────────────────────────────────

/**
 * For each Supplier row where File_ID is set and exactly one of VScale/HScale
 * is filled, computes the missing value using the image's native aspect ratio:
 *
 *   HScale = VScale × (nativeW / nativeH)
 *   VScale = HScale / (nativeW / nativeH)
 *
 * Assumes the image is an orthographic swatch framing the full material region
 * uniformly (no crop / perspective). Values are read by MoodBoard.js and fed
 * into Architextures (https://architextures.org) for accurate texture tiling.
 *
 * Rows with both scales set or neither set are left alone. Image dimensions
 * are parsed from the file's header bytes (JPEG + PNG supported) — no Drive
 * advanced service required.
 */
function computeMissingScales() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sheet   = ss.getSheetByName('Supplier');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { ss.toast('No data rows found.', 'Scales'); return; }

  const colMap = getColMap_(sheet);
  if (!validateCols_(colMap, ['File_ID', 'VScale', 'HScale'], 'Compute Scales')) return;

  const numDataRows = lastRow - 1;
  const numCols     = sheet.getLastColumn();
  const values      = sheet.getRange(2, 1, numDataRows, numCols).getValues();

  const V_COL = colMap['VScale'] + 1;
  const H_COL = colMap['HScale'] + 1;

  const dimCache = {}; // fileId → { w, h } | null
  const errors   = [];
  let computed = 0;

  for (let i = 0; i < numDataRows; i++) {
    const fileId = String(values[i][colMap['File_ID']]).trim();
    if (!fileId) continue;

    const vRaw = values[i][colMap['VScale']];
    const hRaw = values[i][colMap['HScale']];
    const vVal = (vRaw === '' || vRaw === null) ? null : Number(vRaw);
    const hVal = (hRaw === '' || hRaw === null) ? null : Number(hRaw);

    if (vVal && hVal) continue;
    if (!vVal && !hVal) continue;

    let dims;
    try {
      if (!(fileId in dimCache)) dimCache[fileId] = getImageDimensions_(fileId);
      dims = dimCache[fileId];
    } catch (e) {
      errors.push(`Row ${i + 2}: ${e.message}`);
      continue;
    }
    if (!dims) {
      errors.push(`Row ${i + 2}: unsupported image format`);
      continue;
    }

    const ar = dims.w / dims.h;
    if (vVal && !hVal) {
      sheet.getRange(i + 2, H_COL).setValue(round2_(vVal * ar));
    } else {
      sheet.getRange(i + 2, V_COL).setValue(round2_(hVal / ar));
    }
    computed++;
  }

  const msg = `Computed ${computed} scale(s).` +
              (errors.length ? `  Errors: ${errors.length} — see Logger.` : '');
  if (errors.length) errors.forEach(e => Logger.log(e));
  ss.toast(msg, 'Scales Complete', 10);
}

/**
 * Reads native pixel dimensions from the image header bytes. Supports JPEG
 * and PNG (covers the vast majority of material photos). Returns { w, h } or
 * null for unsupported formats.
 */
function getImageDimensions_(fileId) {
  const bytes = DriveApp.getFileById(fileId).getBlob().getBytes();
  const u = (i) => bytes[i] & 0xFF;

  // PNG: signature 89 50 4E 47 0D 0A 1A 0A, IHDR width/height at bytes 16–23
  if (bytes.length >= 24 &&
      u(0) === 0x89 && u(1) === 0x50 && u(2) === 0x4E && u(3) === 0x47) {
    const w = (u(16) << 24) | (u(17) << 16) | (u(18) << 8) | u(19);
    const h = (u(20) << 24) | (u(21) << 16) | (u(22) << 8) | u(23);
    return { w, h };
  }

  // JPEG: starts FF D8; scan segments for an SOF marker (C0–CF, excluding
  // C4 DHT, C8 JPG, CC DAC) to read height/width from the frame header.
  if (u(0) === 0xFF && u(1) === 0xD8) {
    let i = 2;
    while (i < bytes.length - 8) {
      if (u(i) !== 0xFF) return null;
      const marker = u(i + 1);
      if (marker >= 0xC0 && marker <= 0xCF &&
          marker !== 0xC4 && marker !== 0xC8 && marker !== 0xCC) {
        const h = (u(i + 5) << 8) | u(i + 6);
        const w = (u(i + 7) << 8) | u(i + 8);
        return { w, h };
      }
      const len = (u(i + 2) << 8) | u(i + 3);
      if (len < 2) return null;
      i += 2 + len;
    }
  }

  return null;
}

function round2_(n) { return Math.round(n * 100) / 100; }

// ── Audit ─────────────────────────────────────────────────────────────────────

/**
 * Compares the Materials Drive folder against the Filename column of the
 * Supplier sheet. Reports any files in Drive that are NOT referenced by the
 * sheet (orphans). Prompts to trash them after showing the list.
 */
function auditMaterialsFolder() {
  const ui     = SpreadsheetApp.getUi();
  const sheet  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Supplier');
  const lastRow = sheet.getLastRow();

  const colMap = getColMap_(sheet);
  if (!validateCols_(colMap, ['Filename'], 'Audit')) return;

  // Build set of expected filenames from the Filename column
  const expected = new Set();
  if (lastRow > 1) {
    sheet.getRange(2, colMap['Filename'] + 1, lastRow - 1, 1).getValues()
      .forEach(row => {
        const name = String(row[0]).trim();
        if (name) expected.add(name);
      });
  }

  const folder  = DriveApp.getFolderById(MATERIALS_FOLDER_ID);
  const iter    = folder.getFiles();
  const orphans = [];

  while (iter.hasNext()) {
    const f = iter.next();
    if (!expected.has(f.getName())) {
      orphans.push({ name: f.getName(), id: f.getId() });
    }
  }

  if (orphans.length === 0) {
    ui.alert('Audit Complete', '✓ No orphaned files — Materials folder matches the sheet exactly.', ui.ButtonSet.OK);
    return;
  }

  const list    = orphans.map(o => `  • ${o.name}`).join('\n');
  const confirm = ui.alert(
    `Audit: ${orphans.length} orphaned file(s) found`,
    `These files are in the Materials folder but not referenced in the Supplier sheet:\n\n${list}\n\nMove them to Trash?`,
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) {
    ui.alert('No files were deleted.');
    return;
  }

  orphans.forEach(o => DriveApp.getFileById(o.id).setTrashed(true));
  ui.alert('Done', `Moved ${orphans.length} file(s) to Trash.`, ui.ButtonSet.OK);
}

// ── Helpers ───────────────────────────────────────────────────────────────────

/**
 * Builds a { headerName: 0-based-index } map from row 1 of the given sheet.
 * Columns with blank headers are skipped. Matching is case-sensitive.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {Object.<string, number>}
 */
function getColMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => {
    const key = String(h).trim();
    if (key) map[key] = i;
  });
  return map;
}

/**
 * Checks that all required header names are present in colMap.
 * Shows an alert and returns false if any are missing.
 * @param {Object.<string, number>} colMap
 * @param {string[]}                required
 * @param {string}                  contextName  Shown in the alert title
 * @returns {boolean}
 */
function validateCols_(colMap, required, contextName) {
  const missing = required.filter(h => colMap[h] === undefined);
  if (missing.length === 0) return true;
  SpreadsheetApp.getUi().alert(
    contextName + ' — Missing Column(s)',
    'The following headers were not found in row 1 of the Supplier sheet:\n\n  ' +
    missing.join(', ') +
    '\n\nCheck that row 1 contains the exact header names listed above.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  return false;
}

/**
 * Resolves a Drive File object from any of:
 *   - Google Drive URL:  https://drive.google.com/file/d/FILE_ID/...
 *   -                    https://drive.google.com/open?id=FILE_ID
 *   - Windows/relative path: extracts filename, searches Materials folder then all Drive
 */
function resolveFile(input, materialsFolder) {
  input = input.trim().replace(/^"+|"+$/g, '').trim();

  let m = input.match(/\/d\/([a-zA-Z0-9_-]{25,})/);
  if (m) return DriveApp.getFileById(m[1]);

  m = input.match(/[?&]id=([a-zA-Z0-9_-]{25,})/);
  if (m) return DriveApp.getFileById(m[1]);

  const filename = input.replace(/\\/g, '/').split('/').filter(Boolean).pop();
  if (!filename) return null;

  let iter = materialsFolder.getFilesByName(filename);
  if (iter.hasNext()) return iter.next();

  iter = DriveApp.getFilesByName(filename);
  if (iter.hasNext()) return iter.next();

  return null;
}

/**
 * Ensures a file is inside the target folder.
 * In a Shared Drive, a file has exactly one parent; addFile + removeFile handles the move.
 */
function ensureInFolder(file, targetFolder) {
  const targetId  = targetFolder.getId();
  const parents   = file.getParents();
  const parentIds = [];
  while (parents.hasNext()) parentIds.push(parents.next().getId());

  if (parentIds.includes(targetId)) return;

  targetFolder.addFile(file);
  parentIds.forEach(pid => {
    if (pid !== targetId) {
      try { DriveApp.getFolderById(pid).removeFile(file); } catch (_) {}
    }
  });
}

function dedupKey(supplier, product) {
  return `${supplier.trim().toUpperCase()}|${product.trim().toUpperCase()}`;
}

/**
 * Canonical basename:  "Republic Floor" + "Sharc North Forest"
 *                    → "RepublicFloor_Sharc-North-Forest"
 */
function canonicalBasename(supplier, product) {
  const s = supplier.trim().replace(/\s+/g, '');
  const p = product.trim().replace(/\s+/g, '-').replace(/[^a-zA-Z0-9\-]/g, '').replace(/-+/g, '-');
  return `${s}_${p}`;
}

function timestamp() {
  return new Date().toLocaleString();
}

function formatActiveSheetColumns() {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (!sheet.getLastColumn()) return;
  const colMap = getColMap_(sheet);
  if (colMap['Supplier_URL'] !== undefined) sheet.setColumnWidth(colMap['Supplier_URL'] + 1, 250);
  if (colMap['Drive_URL']    !== undefined) sheet.setColumnWidth(colMap['Drive_URL']    + 1, 250);
}

// ── Debug (run from Apps Script editor) ──────────────────────────────────────

function debugRemoteSheet() {
  const remoteSheet = SpreadsheetApp.openById(TEMPLATE_ID).getSheetByName('Bundles');
  const numRows     = remoteSheet.getLastRow();
  const numCols     = remoteSheet.getLastColumn();
  const values      = remoteSheet.getRange(1, 1, Math.min(numRows, 10), Math.min(numCols, 21)).getValues();
  const richText    = remoteSheet.getRange(1, 1, Math.min(numRows, 10), Math.min(numCols, 21)).getRichTextValues();

  Logger.log(`Remote Bundles: ${numRows} rows × ${numCols} cols`);
  Logger.log('=== Row 1 headers ===');
  values[0].forEach((v, i) => { if (String(v).trim()) Logger.log(`  col ${i}: "${v}"`); });
  Logger.log('=== Col A rows 1-8 ===');
  for (let r = 0; r < Math.min(8, values.length); r++) Logger.log(`  row ${r + 1}: "${values[r][0]}"`);
  Logger.log('=== C2 (Subway/Flooring) ===');
  const rt = richText[1][2];
  Logger.log(`  value: "${values[1][2]}"  linkUrl: "${rt ? rt.getLinkUrl() : 'n/a'}"`);
}
