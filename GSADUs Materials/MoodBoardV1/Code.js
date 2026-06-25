/**
 * Design Bundles — Standalone Read-Only Dashboard (server controller)
 *
 * Serves a read-only mood-board / catalog view of the canonical GSADUs Materials
 * database (schema 1.2). The HTML front end is fed a flat projection: one row per
 * Supplier material, with its CANONICAL image joined in from the Images tab
 * (keyed by Material_Key = canonicalBasename(Supplier, Product_Name)).
 *
 * Why a join: under schema 1.2 the Images tab holds the canonical PNG masters.
 * Supplier's legacy File_ID/Drive_URL/Filename columns were removed from the
 * live sheet on 2026-06-25. This dashboard reads only the canonical source. See
 * docs/SOURCES.md in the GSADUs Materials project.
 *
 * Read-only by design — there are no write paths. Editing the canonical database
 * happens in the bound GSADUs Tools script on the sheet itself, never here.
 */

// ── Configuration ─────────────────────────────────────────────────────────────

// Canonical "GSADUs Materials" spreadsheet (shared drive). Schema 1.2.
const SPREADSHEET_ID = '1JT5NJED-NiqOIuC6b-tq78e5mXDUonD-R7EAvaaZriM';
const SUPPLIER_SHEET  = 'Supplier';
const IMAGES_SHEET    = 'Images';

// Lowercased header names the front end resolves by (cleanHeaders.indexOf(...)).
// Order here defines the column index of each field in the returned rows.
const VIEW_HEADERS = [
  'design_bundle', 'category', 'supplier_url', 'supplier', 'product_name',
  'product_size', 'vscale', 'hscale', 'file_id', 'drive_url', 'image', 'showcase',
];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Design Bundles')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── Read ──────────────────────────────────────────────────────────────────────

/**
 * Returns the dashboard dataset as [{ row: VIEW_HEADERS }, { index_, row }, …]:
 * index 0 is the header row, each subsequent entry is one Supplier material with
 * its canonical Material_Image joined from the Images tab.
 */
function getSheetData() {
  const ss       = SpreadsheetApp.openById(SPREADSHEET_ID);
  const supplier = ss.getSheetByName(SUPPLIER_SHEET);
  if (!supplier) throw new Error(`Sheet "${SUPPLIER_SHEET}" not found in ${ss.getName()}.`);

  const lastRow = supplier.getLastRow();
  const result  = [{ row: VIEW_HEADERS.slice() }];
  if (lastRow < 2) return result;

  const sCol = getColMap_(supplier);
  ['Design_Bundle', 'Category', 'Supplier', 'Product_Name'].forEach(h => {
    if (sCol[h] === undefined) throw new Error(`Supplier tab missing required header "${h}".`);
  });

  const numRows = lastRow - 1;
  const values  = supplier.getRange(2, 1, numRows, supplier.getLastColumn()).getValues();
  const urlForm = sCol['Supplier_URL'] !== undefined
    ? supplier.getRange(2, sCol['Supplier_URL'] + 1, numRows, 1).getFormulas()
    : null;

  const imageIndex = buildImageIndex_(ss); // Material_Key → { material, showcase }

  const get = (rowVals, header) =>
    sCol[header] !== undefined ? String(rowVals[sCol[header]]).trim() : '';

  for (let i = 0; i < numRows; i++) {
    const rowVals = values[i];
    const bundle  = get(rowVals, 'Design_Bundle');
    const category = get(rowVals, 'Category');
    if (!bundle && !category) continue; // skip blank rows

    const supplierName = get(rowVals, 'Supplier');
    const productName  = get(rowVals, 'Product_Name');
    const materialKey  = (supplierName && productName)
      ? canonicalBasename(supplierName, productName) : '';

    // Real URL out of the =HYPERLINK("url","text") formula (display value is just text).
    let supplierUrl = get(rowVals, 'Supplier_URL');
    if (urlForm) {
      const m = String(urlForm[i][0]).match(/HYPERLINK\("([^"]+)"/i);
      if (m) supplierUrl = m[1];
    }

    const img      = materialKey ? imageIndex[materialKey] : null;
    const mat      = img ? img.material : null;
    const show     = img ? img.showcase : null;
    const fileId   = mat ? mat.file_id  : '';
    const driveUrl = mat ? mat.drive_url : '';
    const imageUrl    = fileId ? `https://drive.google.com/thumbnail?id=${fileId}&sz=w600` : '';
    // Showcase = the installed/in-context photo (Image_Type "Showcase_Image").
    // Larger render width since it's shown as a hero image, not a swatch.
    const showcaseUrl = (show && show.file_id)
      ? `https://drive.google.com/thumbnail?id=${show.file_id}&sz=w1200` : '';

    result.push({
      index_: i + 2, // canonical Supplier sheet row (stable React key)
      row: [
        bundle, category, supplierUrl, supplierName, productName,
        get(rowVals, 'Product_Size'), get(rowVals, 'VScale'), get(rowVals, 'HScale'),
        fileId, driveUrl, imageUrl, showcaseUrl,
      ],
    });
  }

  return result;
}

// ── Helpers ───────────────────────────────────────────────────────────────────

/**
 * Reads the Images tab and returns
 *   { Material_Key: { material: {file_id,drive_url}|null, showcase: {…}|null } }
 * capturing the first Material_Image (the swatch) and the first Showcase_Image
 * (the installed photo) per material. Returns {} if the tab is absent or empty.
 */
function buildImageIndex_(ss) {
  const sheet = ss.getSheetByName(IMAGES_SHEET);
  const index = {};
  if (!sheet) return index;
  const last = sheet.getLastRow();
  if (last < 2) return index;

  const col = getColMap_(sheet);
  if (col['Material_Key'] === undefined) return index;

  const vals = sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).getValues();
  const ref  = (r) => ({
    file_id:   col['File_ID']   !== undefined ? String(r[col['File_ID']]).trim()   : '',
    drive_url: col['Drive_URL'] !== undefined ? String(r[col['Drive_URL']]).trim() : '',
  });
  vals.forEach(r => {
    const key = String(r[col['Material_Key']]).trim();
    if (!key) return;
    const type  = col['Image_Type'] !== undefined ? String(r[col['Image_Type']]).trim() : 'Material_Image';
    const entry = index[key] || (index[key] = { material: null, showcase: null });
    if (type === 'Showcase_Image') {
      if (!entry.showcase) entry.showcase = ref(r); // first showcase wins
    } else {
      if (!entry.material) entry.material = ref(r); // blank/Material_Image = swatch
    }
  });
  return index;
}

/**
 * Builds a { headerName: 0-based-index } map from row 1 of the given sheet.
 * Blank headers are skipped. Matching is case-sensitive.
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
 * Canonical material key / basename — must match the bound GSADUs Tools script:
 *   "Republic Floor" + "Verona Light" → "RepublicFloor_Verona-Light"
 */
function canonicalBasename(supplier, product) {
  const s = supplier.trim().replace(/\s+/g, '');
  const p = product.trim().replace(/\s+/g, '-').replace(/[^a-zA-Z0-9\-]/g, '').replace(/-+/g, '-');
  return `${s}_${p}`;
}
