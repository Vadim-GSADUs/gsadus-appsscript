/**
 * GSADUs Tools (V11 — schema 1.2)
 *
 * Two tabs, each a source of truth for its own entity (see docs/SOURCES.md):
 *
 *   Supplier  — one row per (Design_Bundle, Category): the material/bundle data.
 *               Required headers: Design_Bundle, Category, Supplier_URL, Supplier,
 *               Product_Name, Product_Size
 *
 *   Images    — one row per image, keyed to a material by Material_Key
 *               (= canonicalBasename(Supplier, Product_Name)). Holds the
 *               image-specific specs. Headers in IMAGES_HEADERS below.
 *
 * Columns are resolved dynamically from the header row — order is not significant.
 *
 * Image files are normalized to a canonical PNG master OUTSIDE this script
 * (PNGTools → Conversion subtab). This script records whatever format is present;
 * it never transcodes.
 */

// ── Constants ────────────────────────────────────────────────────────────────

const SUPPLIER_SHEET      = 'Supplier';
const IMAGES_SHEET        = 'Images';
const MATERIALS_FOLDER_ID = '1hc2moJgK51YPqYxcmm_Zgry5YxbsbGAs';
const TEMPLATE_ID         = '1oGLgK-aCvKVh1EIhADQsqeqWQLlUaCTo4AkmtAY9dU4';
const BUNDLES_FOLDER_ID   = '1v7vLPjvPdMA42wGA9XqC_29DNtZP21Gk'; // Interior Design Bundles folder
const BUNDLES_JSON_NAME   = 'bundles_library.json';
const SCHEMA_VERSION      = '1.2';

// Standalone read-only Mood Board dashboard — a SEPARATE Apps Script web app
// (source: AppsScript/GSADUs Materials/MoodBoardV1). Launched from the menu below.
const MOODBOARD_WEBAPP_URL =
  'https://script.google.com/macros/s/AKfycbxYJkXjXr2XqCP_lF7KEFgLAmwacFQgm_ZLIQaTqn5j-1RphctR2fyqEaL3bfIybu0p/exec';

// Images tab (schema 1.2): one row per image. See docs/SOURCES.md §7.
const IMAGES_HEADERS = [
  'Material_Key', 'Image_Type', 'Source_URL', 'Source_Format',
  'File_ID', 'Drive_URL', 'Filename', 'Format', 'Width_px', 'Height_px',
  'VScale', 'HScale', 'Sync_Status', 'Notes',
];
const IMAGE_TYPE_TOKENS = { 'Material_Image': 'material', 'Showcase_Image': 'showcase' };

// ── Menu ─────────────────────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('GSADUs Tools')
    .addItem('1. Pull from Order Template',  'pullFromOrderTemplate')
    .addItem('2. Sync Image Assets',         'syncImageAssets')
    .addItem('3. Export to JSON',            'exportToJson')
    .addSeparator()
    .addItem('Open Mood Board',              'openMoodBoardDashboard')
    .addSeparator()
    .addItem('Seed Images from Supplier',    'seedImagesFromSupplier')
    .addItem('Compute Missing Scales',       'computeMissingScales')
    .addItem('Audit Materials Folder',       'auditMaterialsFolder')
    .addItem('Format Active Sheet',          'formatActiveSheetColumns')
    .addToUi();
}

/**
 * Opens the standalone read-only Mood Board dashboard (a separate web app — see
 * MOODBOARD_WEBAPP_URL) in a new browser tab. A modal dialog is used because
 * server-side code can't open a tab directly; it auto-opens via window.open and
 * also shows a button as a fallback when the popup is blocked.
 */
function openMoodBoardDashboard() {
  const url  = MOODBOARD_WEBAPP_URL;
  const html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top"><meta charset="utf-8">' +
    '<style>body{font-family:Roboto,Arial,sans-serif;margin:18px;color:#202124}' +
    'p{font-size:13px;color:#5f6368;margin:0 0 14px}' +
    'a.btn{display:inline-block;background:#1a1a2e;color:#fff;text-decoration:none;' +
    'padding:10px 18px;border-radius:8px;font-weight:600;font-size:13px}</style></head>' +
    '<body><p>Opening the Design Bundles mood board in a new tab…</p>' +
    '<a class="btn" href="' + url + '" target="_blank" rel="noopener" ' +
    'onclick="google.script.host.close()">Open Dashboard</a>' +
    '<script>window.open(' + JSON.stringify(url) + ',"_blank");</script></body></html>'
  ).setWidth(380).setHeight(130);
  SpreadsheetApp.getUi().showModalDialog(html, 'Mood Board Dashboard');
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

// ── Images tab — ensure / seed ────────────────────────────────────────────────

/**
 * Returns the Images sheet, creating it with the canonical header row if absent
 * (or backfilling the header row if the tab exists but is empty).
 */
function ensureImagesSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(IMAGES_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(IMAGES_SHEET);
  }
  if (sheet.getLastRow() === 0 || sheet.getLastColumn() === 0) {
    sheet.getRange(1, 1, 1, IMAGES_HEADERS.length).setValues([IMAGES_HEADERS]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * One-time (idempotent) migration: creates a Material_Image row in the Images tab
 * for every unique material in the Supplier tab that currently carries image data
 * (File_ID / Drive_URL / Filename), copying VScale/HScale across. Safe to re-run —
 * skips any material that already has a Material_Image row.
 *
 * The live Supplier tab was cleaned on 2026-06-25: legacy image columns
 * (File_ID/Drive_URL/Filename/Sync_Status) were removed, leaving A:H as the
 * current schema. This helper is retained for old copies/backups where those
 * columns still exist; on the cleaned live sheet it is a safe no-op.
 */
function seedImagesFromSupplier() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const supplier = ss.getSheetByName(SUPPLIER_SHEET);
  if (!supplier) { ss.toast(`No "${SUPPLIER_SHEET}" tab found.`, 'Seed Images'); return; }

  const sCol = getColMap_(supplier);
  if (!validateCols_(sCol, ['Supplier', 'Product_Name'], 'Seed Images')) return;

  const sLast = supplier.getLastRow();
  if (sLast < 2) { ss.toast('No Supplier rows.', 'Seed Images'); return; }

  const sVals = supplier.getRange(2, 1, sLast - 1, supplier.getLastColumn()).getValues();

  const images = ensureImagesSheet_();
  const iCol   = getColMap_(images);

  // Existing (Material_Key|Image_Type) already present → don't duplicate
  const existing = {};
  const iLast = images.getLastRow();
  if (iLast > 1) {
    images.getRange(2, 1, iLast - 1, images.getLastColumn()).getValues().forEach(r => {
      const k = `${String(r[iCol['Material_Key']]).trim()}|${String(r[iCol['Image_Type']]).trim()}`;
      if (k !== '|') existing[k] = true;
    });
  }

  const has = (name) => sCol[name] !== undefined;
  const newRows = [];
  const seen = {};

  for (let i = 0; i < sVals.length; i++) {
    const sup  = String(sVals[i][sCol['Supplier']]).trim();
    const prod = String(sVals[i][sCol['Product_Name']]).trim();
    if (!sup || !prod) continue;

    const key = canonicalBasename(sup, prod);
    if (seen[key]) continue;                       // one Material_Image per material
    seen[key] = true;
    if (existing[`${key}|Material_Image`]) continue;

    const fileId   = has('File_ID')   ? String(sVals[i][sCol['File_ID']]).trim()   : '';
    const driveUrl = has('Drive_URL') ? String(sVals[i][sCol['Drive_URL']]).trim() : '';
    const filename = has('Filename')  ? String(sVals[i][sCol['Filename']]).trim()  : '';
    if (!fileId && !driveUrl && !filename) continue; // nothing to migrate

    const vscale = has('VScale') ? sVals[i][sCol['VScale']] : '';
    const hscale = has('HScale') ? sVals[i][sCol['HScale']] : '';
    const fmt    = filename ? extOf_(filename) : '';

    const row = IMAGES_HEADERS.map(h => {
      switch (h) {
        case 'Material_Key': return key;
        case 'Image_Type':   return 'Material_Image';
        case 'File_ID':      return fileId;
        case 'Drive_URL':    return driveUrl;
        case 'Filename':     return filename;
        case 'Format':       return fmt;
        case 'VScale':       return vscale;
        case 'HScale':       return hscale;
        case 'Sync_Status':  return filename ? 'Seeded: ' + timestamp() : '';
        default:             return '';
      }
    });
    newRows.push(row);
  }

  if (newRows.length) {
    images.getRange(images.getLastRow() + 1, 1, newRows.length, IMAGES_HEADERS.length).setValues(newRows);
  }
  ss.toast(`Seeded ${newRows.length} Material_Image row(s) into "${IMAGES_SHEET}".`, 'Seed Images', 8);
}

// ── STEP 2 ───────────────────────────────────────────────────────────────────

/**
 * For each row in the Images tab (keyed by Material_Key + Image_Type):
 *   1. Resolves the image file from Drive_URL (URL or path) then File_ID.
 *   2. Renames it to the canonical convention and moves it into Materials\:
 *        {Material_Key}__{type}[-{n}].{ext}
 *      e.g. RepublicFloor_Verona-Light__material.png
 *           Roca_Nordico-Snow-UP-12x24__showcase-2.jpg
 *   3. Reads native dimensions + format and writes File_ID / Drive_URL / Filename /
 *      Format / Width_px / Height_px / Sync_Status.
 *
 * Each Images row is a distinct image (no cross-row dedup — a product reused across
 * bundles is a single Material_Key with its images recorded once). Transcoding is
 * done in PNGTools, not here; this records whatever format the file already is.
 */
function syncImageAssets() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(IMAGES_SHEET);
  if (!sheet) {
    ss.toast(`No "${IMAGES_SHEET}" tab. Run "Seed Images from Supplier" first.`, 'Sync Image Assets');
    return;
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { ss.toast('No image rows found.', 'Sync Image Assets'); return; }

  const colMap = getColMap_(sheet);
  const REQUIRED = ['Material_Key', 'Image_Type', 'File_ID', 'Drive_URL', 'Filename',
                    'Format', 'Width_px', 'Height_px', 'Sync_Status'];
  if (!validateCols_(colMap, REQUIRED, 'Sync Image Assets')) return;

  const numDataRows = lastRow - 1;
  const numCols     = sheet.getLastColumn();
  const values      = sheet.getRange(2, 1, numDataRows, numCols).getValues();

  const FILE_ID_COL     = colMap['File_ID']     + 1;
  const DRIVE_URL_COL   = colMap['Drive_URL']   + 1;
  const FILENAME_COL    = colMap['Filename']    + 1;
  const FORMAT_COL      = colMap['Format']      + 1;
  const WIDTH_COL       = colMap['Width_px']    + 1;
  const HEIGHT_COL      = colMap['Height_px']   + 1;
  const SYNC_STATUS_COL = colMap['Sync_Status'] + 1;

  // Seed outputs from existing values; unchanged rows keep their data.
  const outFileId = values.map(r => [String(r[colMap['File_ID']]).trim()]);
  const outDrive  = values.map(r => [String(r[colMap['Drive_URL']]).trim()]);
  const outName   = values.map(r => [String(r[colMap['Filename']]).trim()]);
  const outFormat = values.map(r => [String(r[colMap['Format']]).trim()]);
  const outWidth  = values.map(r => [r[colMap['Width_px']]]);
  const outHeight = values.map(r => [r[colMap['Height_px']]]);
  const outStatus = Array.from({ length: numDataRows }, () => ['']); // cleared each run

  const materialsFolder = DriveApp.getFolderById(MATERIALS_FOLDER_ID);
  const folderIndex = buildMaterialsFolderIndex_(materialsFolder); // basename → live file
  const seqCount = {}; // (Material_Key|Image_Type) → running count for -n suffixing

  let processed = 0, skipped = 0, errors = 0;

  for (let i = 0; i < numDataRows; i++) {
    const materialKey = String(values[i][colMap['Material_Key']]).trim();
    const imageType   = String(values[i][colMap['Image_Type']]).trim();
    const driveUrl    = String(values[i][colMap['Drive_URL']]).trim();
    const fileId      = String(values[i][colMap['File_ID']]).trim();

    if (!materialKey || !imageType) { skipped++; continue; }

    const groupKey = `${materialKey}|${imageType}`;
    seqCount[groupKey] = (seqCount[groupKey] || 0) + 1;
    const seq = seqCount[groupKey];
    const canonicalBase = canonicalImageBasename_(materialKey, imageType, seq);

    let file = null;
    try {
      // 1) Materials folder is authoritative — find the row's file by its
      //    canonical basename among LIVE files (getFiles excludes trashed).
      //    Self-heals rows whose stored File_ID/Drive_URL went stale, e.g. after
      //    a PNGTools format conversion replaced the file.
      file = folderIndex[canonicalBase.toLowerCase()] || null;
      // 2) First-time registration only: a pasted Drive_URL/path or File_ID for
      //    a file not yet canonically placed. Ignore trashed results.
      if (!file && driveUrl) {
        const f = resolveFile(driveUrl, materialsFolder);
        if (f && !f.isTrashed()) file = f;
      }
      if (!file && fileId) {
        const f = DriveApp.getFileById(fileId);
        if (f && !f.isTrashed()) file = f;
      }
    } catch (e) {
      outStatus[i][0] = '⚠ Resolve error: ' + e.message;
      errors++;
      continue;
    }
    if (!file) {
      outStatus[i][0] = '⚠ No file — paste a Drive URL/path into Drive_URL';
      skipped++;
      continue;
    }

    try {
      const currentName = file.getName();
      const ext         = extOf_(currentName);
      const canonical   = `${canonicalBase}.${ext}`;
      if (currentName !== canonical) file.setName(canonical);
      ensureInFolder(file, materialsFolder);

      let dims = null;
      try { dims = getImageDimensions_(file.getId()); } catch (_) { dims = null; }

      outFileId[i][0] = file.getId();
      outDrive[i][0]  = file.getUrl();
      outName[i][0]   = canonical;
      outFormat[i][0] = ext;
      if (dims) { outWidth[i][0] = dims.w; outHeight[i][0] = dims.h; }
      outStatus[i][0] = 'Synced: ' + timestamp();
      processed++;
    } catch (e) {
      outStatus[i][0] = '⚠ Error: ' + e.message;
      errors++;
    }
  }

  sheet.getRange(2, FILE_ID_COL,     numDataRows, 1).setValues(outFileId);
  sheet.getRange(2, DRIVE_URL_COL,   numDataRows, 1).setValues(outDrive);
  sheet.getRange(2, FILENAME_COL,    numDataRows, 1).setValues(outName);
  sheet.getRange(2, FORMAT_COL,      numDataRows, 1).setValues(outFormat);
  sheet.getRange(2, WIDTH_COL,       numDataRows, 1).setValues(outWidth);
  sheet.getRange(2, HEIGHT_COL,      numDataRows, 1).setValues(outHeight);
  sheet.getRange(2, SYNC_STATUS_COL, numDataRows, 1).setValues(outStatus);

  ss.toast(
    `Processed: ${processed}  |  Skipped: ${skipped}  |  Errors: ${errors}`,
    'Sync Image Assets', 12
  );
}

// ── STEP 3 ───────────────────────────────────────────────────────────────────

/**
 * Exports Supplier (material/bundle data) joined with Images (by Material_Key) to
 * bundles_library.json. Schema 1.2:
 *
 *   { _meta, hardware, bundles: [ { name, materials: [ {
 *       category, supplier, product_name, product_size, product_url, material_key,
 *       images: [ { type, file_id, drive_url, filename, format, width, height,
 *                   vscale, hscale, source_url } ]
 *   } ] } ] }
 *
 * product_url is extracted from the =HYPERLINK() formula in Supplier_URL. The old
 * single-image fields on the material row are gone — images live under images[].
 */
function exportToJson() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const supplier = ss.getSheetByName(SUPPLIER_SHEET);
  if (!supplier) { ss.toast(`No "${SUPPLIER_SHEET}" tab found.`, 'Step 3'); return; }
  const sLast = supplier.getLastRow();
  if (sLast < 2) { ss.toast('No Supplier rows.', 'Step 3'); return; }

  const sCol = getColMap_(supplier);
  const REQUIRED = ['Design_Bundle', 'Category', 'Supplier_URL', 'Supplier', 'Product_Name', 'Product_Size'];
  if (!validateCols_(sCol, REQUIRED, 'Step 3')) return;

  const numRows = sLast - 1;
  const sVals   = supplier.getRange(2, 1, numRows, supplier.getLastColumn()).getValues();
  const sForm   = supplier.getRange(2, sCol['Supplier_URL'] + 1, numRows, 1).getFormulas();

  const imageIndex = buildImageIndex_(); // Material_Key → [ image, … ]

  const BUNDLE_ORDER = ['Subway', 'Harbor', 'Navy', 'Olive', 'Antique', 'Villa'];
  const bundleMap    = {};
  BUNDLE_ORDER.forEach(name => { bundleMap[name] = { name, materials: [] }; });

  for (let i = 0; i < numRows; i++) {
    const bundle   = String(sVals[i][sCol['Design_Bundle']]).trim();
    const category = String(sVals[i][sCol['Category']]).trim();
    if (!bundle || !category) continue;

    let productUrl = null;
    const formula  = sForm[i][0];
    if (formula) {
      const m = formula.match(/HYPERLINK\("([^"]+)"/i);
      if (m) productUrl = m[1];
    }

    const supplierName = String(sVals[i][sCol['Supplier']]).trim();
    const productName  = String(sVals[i][sCol['Product_Name']]).trim();
    const materialKey  = (supplierName && productName) ? canonicalBasename(supplierName, productName) : '';

    const material = {
      category:      category,
      supplier:      supplierName || null,
      product_name:  productName || null,
      product_size:  String(sVals[i][sCol['Product_Size']]).trim() || null,
      product_url:   productUrl,
      material_key:  materialKey || null,
      images:        (materialKey && imageIndex[materialKey]) ? imageIndex[materialKey] : [],
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
      schema_version: SCHEMA_VERSION,
    },
    hardware: HARDWARE_FINISHES,
    bundles,
  };

  const folder   = DriveApp.getFolderById(BUNDLES_FOLDER_ID);
  const existing = folder.getFilesByName(BUNDLES_JSON_NAME);
  while (existing.hasNext()) existing.next().setTrashed(true);
  folder.createFile(BUNDLES_JSON_NAME, JSON.stringify(output, null, 2), MimeType.PLAIN_TEXT);

  const totalMaterials = bundles.reduce((n, b) => n + b.materials.length, 0);
  const totalImages    = bundles.reduce((n, b) => n + b.materials.reduce((m, x) => m + x.images.length, 0), 0);
  ss.toast(
    `Exported ${bundles.length} bundles · ${totalMaterials} materials · ${totalImages} images → ${BUNDLES_JSON_NAME}`,
    'Step 3 Complete', 8
  );
}

/**
 * Reads the Images tab and returns { Material_Key: [ imageObject, … ] }.
 * Returns {} when the tab is absent or empty.
 */
function buildImageIndex_() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(IMAGES_SHEET);
  const index = {};
  if (!sheet) return index;
  const last = sheet.getLastRow();
  if (last < 2) return index;

  const col = getColMap_(sheet);
  if (col['Material_Key'] === undefined || col['Image_Type'] === undefined) return index;

  const vals = sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).getValues();
  const num = (v) => (v === '' || v === null) ? null : Number(v);
  const str = (v) => { const s = String(v == null ? '' : v).trim(); return s || null; };

  vals.forEach(r => {
    const key = String(r[col['Material_Key']]).trim();
    if (!key) return;
    const img = {
      type:       str(r[col['Image_Type']]),
      file_id:    col['File_ID']    !== undefined ? str(r[col['File_ID']])    : null,
      drive_url:  col['Drive_URL']  !== undefined ? str(r[col['Drive_URL']])  : null,
      filename:   col['Filename']   !== undefined ? str(r[col['Filename']])   : null,
      format:     col['Format']     !== undefined ? str(r[col['Format']])     : null,
      width:      col['Width_px']   !== undefined ? num(r[col['Width_px']])   : null,
      height:     col['Height_px']  !== undefined ? num(r[col['Height_px']])  : null,
      vscale:     col['VScale']     !== undefined ? num(r[col['VScale']])     : null,
      hscale:     col['HScale']     !== undefined ? num(r[col['HScale']])     : null,
      source_url: col['Source_URL'] !== undefined ? str(r[col['Source_URL']]) : null,
    };
    (index[key] = index[key] || []).push(img);
  });
  return index;
}

// ── Compute Scales ────────────────────────────────────────────────────────────

/**
 * For each Images row of type Material_Image where File_ID is set and exactly one
 * of VScale/HScale is filled, computes the missing value from the image's native
 * aspect ratio:
 *
 *   HScale = VScale × (nativeW / nativeH)
 *   VScale = HScale / (nativeW / nativeH)
 *
 * VScale/HScale apply to Material_Image only (proportional texture tiling — fed to
 * Architextures / Mood Board). Rows with both or neither set are left alone. Image
 * dimensions are parsed from header bytes (JPEG + PNG).
 */
function computeMissingScales() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(IMAGES_SHEET);
  if (!sheet) { ss.toast(`No "${IMAGES_SHEET}" tab found.`, 'Scales'); return; }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { ss.toast('No image rows found.', 'Scales'); return; }

  const colMap = getColMap_(sheet);
  if (!validateCols_(colMap, ['Image_Type', 'File_ID', 'VScale', 'HScale'], 'Compute Scales')) return;

  const numDataRows = lastRow - 1;
  const numCols     = sheet.getLastColumn();
  const values      = sheet.getRange(2, 1, numDataRows, numCols).getValues();

  const V_COL = colMap['VScale'] + 1;
  const H_COL = colMap['HScale'] + 1;

  const dimCache = {}; // fileId → { w, h } | null
  const errors   = [];
  let computed = 0;

  for (let i = 0; i < numDataRows; i++) {
    if (String(values[i][colMap['Image_Type']]).trim() !== 'Material_Image') continue;
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
 * Compares the Materials Drive folder against the Filename column of the Images
 * tab. Reports any files in Drive that are NOT referenced by the sheet (orphans).
 * Prompts to trash them after showing the list.
 */
function auditMaterialsFolder() {
  const ui     = SpreadsheetApp.getUi();
  const sheet  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(IMAGES_SHEET);
  if (!sheet) { ui.alert('Audit', `No "${IMAGES_SHEET}" tab found.`, ui.ButtonSet.OK); return; }
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
    ui.alert('Audit Complete', '✓ No orphaned files — Materials folder matches the Images tab exactly.', ui.ButtonSet.OK);
    return;
  }

  const list    = orphans.map(o => `  • ${o.name}`).join('\n');
  const confirm = ui.alert(
    `Audit: ${orphans.length} orphaned file(s) found`,
    `These files are in the Materials folder but not referenced in the Images tab:\n\n${list}\n\nMove them to Trash?`,
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
    'The following headers were not found in row 1:\n\n  ' +
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

/**
 * Canonical material key / basename:
 *   "Republic Floor" + "Sharc North Forest" → "RepublicFloor_Sharc-North-Forest"
 */
function canonicalBasename(supplier, product) {
  const s = supplier.trim().replace(/\s+/g, '');
  const p = product.trim().replace(/\s+/g, '-').replace(/[^a-zA-Z0-9\-]/g, '').replace(/-+/g, '-');
  return `${s}_${p}`;
}

/**
 * Builds { basename-without-ext (lowercased) : File } from the LIVE files in the
 * Materials folder (getFiles excludes trashed files). When several live files
 * share a basename, the canonical PNG wins. The Materials folder is the
 * authoritative store for image bytes, so Sync uses this to relink rows whose
 * stored File_ID/Drive_URL went stale (e.g. after a format conversion replaced
 * the file and left the old id trashed-but-resolvable).
 */
function buildMaterialsFolderIndex_(folder) {
  const index = {};
  const iter = folder.getFiles();
  while (iter.hasNext()) {
    const f    = iter.next();
    const name = f.getName();
    const dot  = name.lastIndexOf('.');
    const base = (dot > 0 ? name.slice(0, dot) : name).toLowerCase();
    const ext  = (dot > 0 ? name.slice(dot + 1) : '').toLowerCase();
    const existing = index[base];
    if (!existing) {
      index[base] = f;
    } else {
      const existingExt = existing.getName().split('.').pop().toLowerCase();
      if (ext === 'png' && existingExt !== 'png') index[base] = f; // prefer canonical PNG
    }
  }
  return index;
}

/**
 * Canonical image basename:  {Material_Key}__{type}[-{n}]
 *   ("RepublicFloor_Verona-Light", "Material_Image", 1) → "RepublicFloor_Verona-Light__material"
 *   (…, "Showcase_Image", 2)                            → "…__showcase-2"
 */
function canonicalImageBasename_(materialKey, imageType, seq) {
  const token = IMAGE_TYPE_TOKENS[imageType] ||
    String(imageType).trim().toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '') ||
    'image';
  const suffix = (seq && seq > 1) ? `-${seq}` : '';
  return `${materialKey}__${token}${suffix}`;
}

/** Lowercase file extension (no dot), defaulting to 'jpg'. */
function extOf_(filename) {
  const n = String(filename || '');
  return n.includes('.') ? n.split('.').pop().toLowerCase() : 'jpg';
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
  if (colMap['Source_URL']   !== undefined) sheet.setColumnWidth(colMap['Source_URL']   + 1, 250);
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
