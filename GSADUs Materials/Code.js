/**
 * GSADUs Tools (V9)
 *
 * Supplier tab column layout:
 *   A  Design_Bundle   — written by Step 1
 *   B  Category        — written by Step 1
 *   C  Supplier_URL    — written by Step 1 (HYPERLINK formula)
 *   D  Supplier        — manual  (e.g. "Roca", "Republic Floor")
 *   E  Product_Name    — manual  (e.g. "Nordico Snow UP 12x24")
 *   F  Sourced_Url     — manual, optional reference (not touched by script)
 *   G  File_ID         — written by Step 2
 *   H  Drive_URL       — INPUT for Step 2 (any format); overwritten with canonical URL
 *   I  Filename        — written by Step 2 (canonical name)
 *   J  Sync_Status     — written by Step 2
 */

// ── Constants ────────────────────────────────────────────────────────────────

const MATERIALS_FOLDER_ID = '1hc2moJgK51YPqYxcmm_Zgry5YxbsbGAs';
const TEMPLATE_ID         = '1oGLgK-aCvKVh1EIhADQsqeqWQLlUaCTo4AkmtAY9dU4';

// 0-based column indices for the Supplier sheet
const COL = {
  BUNDLE:       0,  // A
  CATEGORY:     1,  // B
  SUPPLIER_URL: 2,  // C
  SUPPLIER:     3,  // D
  PRODUCT_NAME: 4,  // E
  SOURCED_URL:  5,  // F  (reference only — never written by script)
  FILE_ID:      6,  // G
  DRIVE_URL:    7,  // H
  FILENAME:     8,  // I
  SYNC_STATUS:  9   // J
};

const NUM_COLS = 10; // A:J

// ── Menu ─────────────────────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('GSADUs Tools')
    .addItem('1. Pull from Order Template',  'pullFromOrderTemplate')
    .addItem('2. Sync Material Assets',      'syncMaterialAssets')
    .addSeparator()
    .addItem('Audit Materials Folder',       'auditMaterialsFolder')
    .addItem('Format Active Sheet',          'formatActiveSheetColumns')
    .addToUi();
}

// ── STEP 1 ───────────────────────────────────────────────────────────────────

/**
 * Reads the ORDER TEMPLATE > Bundles tab and writes hyperlinked product names
 * into Supplier col C (Supplier_URL) as =HYPERLINK() formulas.
 *
 * Source layout:
 *   Row 1 : bundle name headers (C=Subway, F=Harbor, I=Navy, L=Olive, O=Antique, R=Villa)
 *   Col A : category labels (rows 2–7): Flooring, Bathroom Floor Tile, etc.
 */
function pullFromOrderTemplate() {
  const ss            = SpreadsheetApp.getActiveSpreadsheet();
  const supplierSheet = ss.getSheetByName('Supplier');
  const remoteSheet   = SpreadsheetApp.openById(TEMPLATE_ID).getSheetByName('Bundles');

  // Read remote sheet in one batch
  const numRows   = remoteSheet.getLastRow();
  const numCols   = remoteSheet.getLastColumn();
  const fullRange = remoteSheet.getRange(1, 1, numRows, numCols);
  const values    = fullRange.getValues();
  const formulas  = fullRange.getFormulas();
  const richText  = fullRange.getRichTextValues();

  // Bundle name → 0-based column index (header row 1)
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

  // Existing Supplier rows → "BUNDLE|CATEGORY" → 1-based sheet row
  const lastRow = Math.max(1, supplierSheet.getLastRow());
  const rowMap  = {};
  supplierSheet.getRange(1, 1, lastRow, 2).getValues().forEach((row, i) => {
    const k = `${String(row[0]).trim().toUpperCase()}|${String(row[1]).trim().toUpperCase()}`;
    if (k !== '|') rowMap[k] = i + 1;
  });

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

      const cell = supplierSheet.getRange(sheetRow, COL.SUPPLIER_URL + 1);
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
 * For each Supplier row where D (Supplier) + E (Product_Name) are filled:
 *
 *   1. Computes canonical filename: Supplier_Product-Name.ext
 *   2. Dedup check — if another row already processed the same Supplier+Product,
 *      reuses that Drive file (no duplicate files in Materials folder).
 *   3. Otherwise resolves the file from H (Drive_URL), which accepts:
 *        a. Google Drive URL  → https://drive.google.com/file/d/FILE_ID/...
 *        b. Windows path      → G:\Shared drives\...\filename.jpg
 *        c. Relative path     → Materials\filename.jpg
 *      Renames the file to canonical name, moves it into Materials\ if needed.
 *   4. Writes G (File_ID), H (canonical Drive URL), I (Filename), J (Sync_Status).
 *
 * Sourced_Url (col F) is never touched — it's a reference-only field.
 */
function syncMaterialAssets() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sheet   = ss.getSheetByName('Supplier');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { ss.toast('No data rows found.', 'Step 2'); return; }

  const numDataRows = lastRow - 1;

  // Read only D:I (cols 4-9) — inputs for Step 2.
  // Deliberately skipping A:C so that =HYPERLINK() formulas in col C
  // are never read as plain text and written back, which would destroy them.
  const inputs = sheet.getRange(2, 4, numDataRows, 6).getValues();
  // inputs[i]: [0]=D Supplier, [1]=E Product_Name, [2]=F Sourced_Url,
  //            [3]=G File_ID,  [4]=H Drive_URL,    [5]=I Filename (existing)

  // Read col C display text (Supplier_URL label) — for Supplier_URL-based dedup.
  // Read-only: we never write back to col C.
  const supplierUrlTexts = sheet.getRange(2, 3, numDataRows, 1).getValues();
  // supplierUrlTexts[i][0] = display text, e.g. "Verona Light"

  // Output array for G:J (cols 7-10) — only these columns are written back.
  // outputs[i]: [0]=G File_ID, [1]=H Drive_URL, [2]=I Filename, [3]=J Sync_Status
  const outputs = inputs.map(row => [row[3], row[4], row[5], '']); // seed G,H,I from existing

  const materialsFolder = DriveApp.getFolderById(MATERIALS_FOLDER_ID);

  // ── PRE-PASS: propagate existing Drive data to rows sharing the same Supplier_URL ──
  // Build a map from Supplier_URL display text → existing drive data for rows that
  // already have a File_ID. Then fill any rows missing File_ID that share the same text.
  const urlFillMap = {}; // supplierUrlText (lowercase) → { fileId, driveUrl, filename }
  for (let i = 0; i < inputs.length; i++) {
    const fileId   = String(inputs[i][3]).trim();
    const driveUrl = String(inputs[i][4]).trim();
    const filename = String(inputs[i][5]).trim();
    const urlText  = String(supplierUrlTexts[i][0]).trim().toLowerCase();
    if (fileId && urlText && !urlFillMap[urlText]) {
      urlFillMap[urlText] = { fileId, driveUrl, filename };
    }
  }
  // Fill missing G/H from matched Supplier_URL rows
  let prefilled = 0;
  for (let i = 0; i < inputs.length; i++) {
    if (String(inputs[i][3]).trim()) continue; // already has File_ID
    const urlText = String(supplierUrlTexts[i][0]).trim().toLowerCase();
    if (urlText && urlFillMap[urlText]) {
      const fill = urlFillMap[urlText];
      inputs[i][3] = fill.fileId;   // G
      inputs[i][4] = fill.driveUrl; // H
      inputs[i][5] = fill.filename; // I
      prefilled++;
    }
  }

  // Dedup map built progressively — keyed "SUPPLIER|PRODUCT" → { fileId, driveUrl, filename }
  // Not pre-seeded: ensures every file is renamed to current canonical convention each run.
  const dedupMap = {};

  let processed = 0, deduped = 0, skipped = 0, errors = 0;

  for (let i = 0; i < inputs.length; i++) {
    const supplier = String(inputs[i][0]).trim();
    const product  = String(inputs[i][1]).trim();
    // inputs[i][2] = Sourced_Url — reference only, never used by script
    const existingFileId  = String(inputs[i][3]).trim();
    const existingDriveUrl = String(inputs[i][4]).trim();

    if (!supplier || !product) { skipped++; continue; }

    const key      = dedupKey(supplier, product);
    const basename = canonicalBasename(supplier, product);

    // Same Supplier+Product already handled this run → propagate
    if (dedupMap[key]) {
      const asset = dedupMap[key];
      outputs[i] = [asset.fileId, asset.driveUrl, asset.filename, 'Synced (shared): ' + timestamp()];
      deduped++;
      continue;
    }

    // Resolve file: try H (Drive URL / path) first, then G (File_ID)
    let file = null;
    try {
      if (existingDriveUrl) file = resolveFile(existingDriveUrl, materialsFolder);
      if (!file && existingFileId) file = DriveApp.getFileById(existingFileId);
    } catch (e) {
      outputs[i][3] = '⚠ Resolve error: ' + e.message;
      errors++;
      continue;
    }

    if (!file) {
      outputs[i][3] = '⚠ No file found — paste a Drive URL or path into col H';
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

      outputs[i] = [asset.fileId, asset.driveUrl, asset.filename, 'Synced: ' + timestamp()];
      processed++;

    } catch (e) {
      outputs[i][3] = '⚠ Error: ' + e.message;
      errors++;
    }
  }

  // Write only G:J — col C (Supplier_URL) and A:F are never touched
  sheet.getRange(2, 7, numDataRows, 4).setValues(outputs);

  ss.toast(
    `Processed: ${processed}  |  Shared (dedup): ${deduped}  |  Pre-filled (same URL): ${prefilled}  |  Skipped: ${skipped}  |  Errors: ${errors}`,
    'Step 2 Complete', 12
  );
}

// ── Audit ─────────────────────────────────────────────────────────────────────

/**
 * Compares the Materials Drive folder against col I (Filename) of the Supplier sheet.
 * Reports any files in Drive that are NOT referenced by the sheet (orphans).
 * Prompts to trash them after showing the list.
 */
function auditMaterialsFolder() {
  const ui     = SpreadsheetApp.getUi();
  const sheet  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Supplier');
  const lastRow = sheet.getLastRow();

  // Build set of expected filenames from col I (1-based col 9)
  const expected = new Set();
  if (lastRow > 1) {
    sheet.getRange(2, COL.FILENAME + 1, lastRow - 1, 1).getValues()
      .forEach(row => {
        const name = String(row[0]).trim();
        if (name) expected.add(name);
      });
  }

  // Scan Materials folder for actual files
  const folder  = DriveApp.getFolderById(MATERIALS_FOLDER_ID);
  const iter    = folder.getFiles();
  const orphans = []; // { name, id }

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

  // Show list and ask to delete
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
 * Resolves a Drive File object from any of:
 *   - Google Drive URL:  https://drive.google.com/file/d/FILE_ID/...
 *   -                    https://drive.google.com/open?id=FILE_ID
 *   - Windows/relative path: extracts filename, searches Materials folder then all Drive
 */
function resolveFile(input, materialsFolder) {
  // Strip surrounding double quotes added by Windows Ctrl+Shift+C copy
  input = input.trim().replace(/^"+|"+$/g, '').trim();

  // Drive URL with /d/ID
  let m = input.match(/\/d\/([a-zA-Z0-9_-]{25,})/);
  if (m) return DriveApp.getFileById(m[1]);

  // Drive URL with ?id= or &id=
  m = input.match(/[?&]id=([a-zA-Z0-9_-]{25,})/);
  if (m) return DriveApp.getFileById(m[1]);

  // Path (Windows or relative) — extract filename after last slash or backslash
  const filename = input.replace(/\\/g, '/').split('/').filter(Boolean).pop();
  if (!filename) return null;

  // Search Materials folder first
  let iter = materialsFolder.getFilesByName(filename);
  if (iter.hasNext()) return iter.next();

  // Fallback: search all of Drive
  iter = DriveApp.getFilesByName(filename);
  if (iter.hasNext()) return iter.next();

  return null;
}

/**
 * Ensures a file is inside the target folder.
 * In a Shared Drive, a file has exactly one parent; addFile + removeFile handles the move.
 */
function ensureInFolder(file, targetFolder) {
  const targetId = targetFolder.getId();
  const parents  = file.getParents();
  const parentIds = [];
  while (parents.hasNext()) parentIds.push(parents.next().getId());

  if (parentIds.includes(targetId)) return; // already there

  targetFolder.addFile(file);
  parentIds.forEach(pid => {
    if (pid !== targetId) {
      try { DriveApp.getFolderById(pid).removeFile(file); } catch (_) {}
    }
  });
}

/** Dedup map key */
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
  sheet.setColumnWidth(COL.SUPPLIER_URL + 1, 250);  // C
  sheet.setColumnWidth(COL.DRIVE_URL + 1,    250);  // H
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
