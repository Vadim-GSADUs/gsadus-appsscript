/**
 * GSADUs Mood Board (container-bound to Design Bundles - Mood Board.gslides)
 *
 * Reads material data from the GSADUs Materials Google Sheet and updates
 * one slide per design bundle in this presentation.
 *
 * Columns are resolved dynamically from row 1 headers — the script does NOT
 * depend on a fixed column order. Add or reorder columns freely; as long as
 * row 1 header names match, everything continues to work.
 *
 * Required headers in the Supplier tab:
 *   Design_Bundle, Category, Supplier_URL, Product_Name, Product_Size, File_ID
 *
 * Optional headers (used when present):
 *   VScale  — real-world height (in) the image represents
 *   HScale  — real-world width (in) the image represents
 *
 * Images and label text boxes are tagged with setTitle() so their positions
 * are preserved across re-syncs. Move elements freely after the first sync —
 * subsequent syncs update content and size but keep your layout.
 *
 * SETUP: set SHEET_ID below to the ID of your GSADUs Materials spreadsheet
 *        (found in its URL: docs.google.com/spreadsheets/d/SHEET_ID/edit)
 */

// ── Configuration ─────────────────────────────────────────────────────────────

const SHEET_ID            = '1JT5NJED-NiqOIuC6b-tq78e5mXDUonD-R7EAvaaZriM';
const MATERIALS_FOLDER_ID = '1hc2moJgK51YPqYxcmm_Zgry5YxbsbGAs';

const BUNDLE_ORDER = ['Subway', 'Harbor', 'Navy', 'Olive', 'Antique', 'Villa'];

const CATEGORIES = [
  'Flooring',          'Bathroom Floor Tile', 'Shower Wall Tile',
  'Shower Pan Tile',   'Kitchen Backsplash',  'Cabinet Color',
];

// Real-world reference: REF_INCHES of material height maps to the maximum
// card image height before bundle-wide fitting is applied.
const REF_INCHES = 24;
const SCALE_MULTIPLIER = 4;

// Slide geometry (points). Standard Google Slides: 720 × 540.
const MB_SLIDE_W   = 720;
const MB_SLIDE_H   = 540;
const MB_TITLE_H   =  44;
const MB_GRID_TOP  =  44;   // grid starts immediately below title bar
const MB_CELL_W    = 240;   // 720 / 3 cols
const MB_CELL_H    = 248;   // (540 - 44) / 2 rows
const MB_LABEL_H   =  48;   // floating label text box height (below image)
const MB_LABEL_GAP =   4;   // gap between image bottom and label top
const MB_IMG_MAX_H = MB_CELL_H - MB_LABEL_H - MB_LABEL_GAP; // 196 pt

// ── Menu ──────────────────────────────────────────────────────────────────────

function onOpen() {
  SlidesApp.getUi()
    .createMenu('GSADUs')
    .addItem('Sync Mood Board', 'syncMoodBoard')
    .addToUi();
}

// ── Sync Mood Board ───────────────────────────────────────────────────────────

/**
 * Updates each bundle slide from the Supplier sheet.
 * Positions are preserved on re-sync via title tags.
 * First run (no tags detected) clears the slide and inserts at default positions.
 *
 * NOTE: fetches up to 36 image blobs (6 bundles × 6 categories).
 * Expect 30–90 seconds execution time depending on file sizes.
 */
function syncMoodBoard() {
  const pres  = SlidesApp.getActivePresentation();
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Supplier');

  if (!sheet) {
    SlidesApp.getUi().alert('Could not open Supplier sheet. Check SHEET_ID in MoodBoard.js.');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { SlidesApp.getUi().alert('No data rows found in Supplier sheet.'); return; }

  const colMap = getColMap_(sheet);
  const REQUIRED = ['Design_Bundle', 'Category', 'Supplier_URL', 'Product_Name', 'Product_Size', 'File_ID'];
  if (!validateMbCols_(colMap, REQUIRED)) return;

  const numDataRows = lastRow - 1;
  const numCols     = sheet.getLastColumn();

  const values   = sheet.getRange(2, 1, numDataRows, numCols).getValues();
  const formulas = sheet.getRange(2, colMap['Supplier_URL'] + 1, numDataRows, 1).getFormulas();

  // Build bundle → [material] map
  const bundleData = {};
  BUNDLE_ORDER.forEach(name => { bundleData[name] = []; });

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

    // VScale / HScale — optional columns, treat blank or zero as null
    const vRaw   = colMap['VScale'] !== undefined ? values[i][colMap['VScale']] : null;
    const hRaw   = colMap['HScale'] !== undefined ? values[i][colMap['HScale']] : null;
    const vscale = (vRaw && Number(vRaw) > 0) ? Number(vRaw) : null;
    const hscale = (hRaw && Number(hRaw) > 0) ? Number(hRaw) : null;

    const entry = {
      category:     category,
      product_name: String(values[i][colMap['Product_Name']]).trim() || null,
      product_size: String(values[i][colMap['Product_Size']]).trim() || null,
      product_url:  productUrl,
      file_id:      String(values[i][colMap['File_ID']]).trim()      || null,
      vscale:       vscale,
      hscale:       hscale,
    };

    if (!bundleData[bundle]) bundleData[bundle] = [];
    bundleData[bundle].push(entry);
  }

  // ── Ensure exactly BUNDLE_ORDER.length slides ──
  let slideCount = pres.getSlides().length;
  while (slideCount < BUNDLE_ORDER.length) { pres.appendSlide(); slideCount++; }
  while (pres.getSlides().length > BUNDLE_ORDER.length) {
    const all = pres.getSlides();
    all[all.length - 1].remove();
  }

  // ── Update each slide ──
  BUNDLE_ORDER.forEach((bundleName, idx) => {
    const slide = pres.getSlides()[idx];
    updateBundleSlide_(slide, bundleName, bundleData[bundleName] || []);
  });

  pres.saveAndClose();
  SlidesApp.openById(pres.getId());
}

// ── Slide updater ─────────────────────────────────────────────────────────────

/**
 * Updates one bundle slide using tag-based position preservation.
 *
 * Each material image is tagged: image.setTitle(category)
 * Each label text box is tagged: label.setTitle('label:' + category)
 * The title bar is tagged:       titleBar.setTitle('title')
 *
 * On first run (no 'title' tag found), clears the slide and inserts fresh.
 * On re-sync, finds each element by tag and updates content/size in place.
 */
function updateBundleSlide_(slide, bundleName, materials) {
  // category (uppercase) → first matching material entry
  const catMap = {};
  materials.forEach(m => {
    const k = m.category.trim().toUpperCase();
    if (!catMap[k]) catMap[k] = m;
  });

  // Detect whether this slide already has our tag-based layout
  const initialized = slide.getPageElements().some(el => {
    try { return el.getTitle() === 'title'; } catch (_) { return false; }
  });

  if (!initialized) {
    // First run or migration from old overlay layout — clear and rebuild fresh
    slide.getPageElements().forEach(el => el.remove());
    const tb = slide.insertTextBox(bundleName.toUpperCase(), 0, 0, MB_SLIDE_W, MB_TITLE_H);
    tb.getFill().setSolidFill('#1a1a2e');
    tb.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
    const tt = tb.getText();
    tt.getTextStyle().setFontSize(20).setBold(true).setForegroundColor('#ffffff');
    tt.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    tb.setTitle('title');
    tb.setDescription(bundleName);
  } else {
    // Update title bar text only
    const titleEl = findElByTitle_(slide.getShapes(), 'title');
    if (titleEl) titleEl.getText().setText(bundleName.toUpperCase());
  }

  const renderItems = [];

  CATEGORIES.forEach((cat, idx) => {
    const col = idx % 3;
    const row = Math.floor(idx / 3);

    // Default grid positions — fits slide bounds exactly with label below image
    const defaultLeft = col * MB_CELL_W;
    const defaultTop  = MB_GRID_TOP + row * MB_CELL_H;

    const mat = catMap[cat.toUpperCase()];

    // Find existing visual element (image or placeholder shape) and label
    const existingMat   = findMaterialEl_(slide, cat);
    const existingLabel = findElByTitle_(slide.getShapes(), 'label:' + cat);

    // Preserve position if element already placed by user
    const left = existingMat ? existingMat.getLeft() : defaultLeft;
    const top  = existingMat ? existingMat.getTop()  : defaultTop;

    // Remove existing material element — will be replaced below
    if (existingMat) existingMat.remove();

    const item = {
      cat: cat,
      mat: mat,
      left: left,
      top: top,
      existingLabel: existingLabel,
      visual: null,
      nativeW: null,
      nativeH: null,
      imgW: MB_CELL_W,
      imgH: MB_IMG_MAX_H,
    };

    if (mat && mat.file_id) {
      try {
        const blob = DriveApp.getFileById(mat.file_id).getBlob();
        // Insert at natural size to read native pixel dimensions for aspect ratio
        const img = slide.insertImage(blob);
        img.setTitle(cat);
        img.setDescription(bundleName);
        item.visual = img;
        item.nativeW = img.getWidth();
        item.nativeH = img.getHeight();
      } catch (_) {
        // File fetch failed — show error placeholder
        const ph = insertMbPlaceholder_(slide, left, top, MB_CELL_W, MB_IMG_MAX_H, '#cccccc');
        ph.setTitle(cat);
        ph.setDescription(bundleName);
        item.visual = ph;
      }
    } else {
      // No file_id yet — placeholder (darker if row exists, lighter if absent)
      const color = mat ? '#cccccc' : '#e0e0e0';
      const ph = insertMbPlaceholder_(slide, left, top, MB_CELL_W, MB_IMG_MAX_H, color);
      ph.setTitle(cat);
      ph.setDescription(bundleName);
      item.visual = ph;
    }

    renderItems.push(item);
  });

  const pointsPerInch = computeBundleScale_(renderItems);

  renderItems.forEach(item => {
    if (item.nativeW && item.nativeH) {
      const sz = computeImageSize_(item.mat && item.mat.vscale, item.mat && item.mat.hscale, item.nativeW, item.nativeH, pointsPerInch);
      item.visual.setLeft(item.left);
      item.visual.setTop(item.top);
      item.visual.setWidth(sz.w);
      item.visual.setHeight(sz.h);
      item.imgW = sz.w;
      item.imgH = sz.h;
    }

    // Label text box — co-located directly below the image
    const labelLeft = item.left;
    const labelTop  = item.top + item.imgH + MB_LABEL_GAP;

    if (item.existingLabel) {
      item.existingLabel.setLeft(labelLeft);
      item.existingLabel.setTop(labelTop);
      item.existingLabel.setWidth(MB_CELL_W);
      item.existingLabel.setHeight(MB_LABEL_H);
      setLabelText_(item.existingLabel, item.cat, item.mat);
    } else {
      const lb = slide.insertTextBox('', labelLeft, labelTop, MB_CELL_W, MB_LABEL_H);
      lb.setTitle('label:' + item.cat);
      lb.setDescription(bundleName);
      lb.getFill().setTransparent();
      lb.getBorder().setTransparent();
      setLabelText_(lb, item.cat, item.mat);
    }
  });
}

// ── Helpers ───────────────────────────────────────────────────────────────────

/**
 * Computes real-world width/height in inches using the provided scale columns.
 * If only one axis is provided, the other is inferred from the image aspect ratio.
 * Returns null when no scale metadata is present.
 *
 * @param {number|null} vscale
 * @param {number|null} hscale
 * @param {number}      nativeW
 * @param {number}      nativeH
 * @returns {{ w: number, h: number } | null}
 */
function getRealWorldSize_(vscale, hscale, nativeW, nativeH) {
  if (!vscale && !hscale) return null;

  const aspect = nativeW / nativeH;

  if (vscale && hscale) {
    return { w: hscale, h: vscale };
  }

  if (vscale) {
    return { w: vscale * aspect, h: vscale };
  }

  return { w: hscale, h: hscale / aspect };
}

/**
 * Computes a shared points-per-inch factor for a bundle slide so all scaled
 * images preserve their exact proportions relative to each other.
 *
 * @param {Array<Object>} renderItems
 * @returns {number}
 */
function computeBundleScale_(renderItems) {
  let pointsPerInch = MB_IMG_MAX_H / REF_INCHES;

  renderItems.forEach(item => {
    if (!item.nativeW || !item.nativeH) return;

    const realSize = getRealWorldSize_(item.mat && item.mat.vscale, item.mat && item.mat.hscale, item.nativeW, item.nativeH);
    if (!realSize) return;

    pointsPerInch = Math.min(
      pointsPerInch,
      MB_CELL_W / realSize.w,
      MB_IMG_MAX_H / realSize.h
    );
  });

  return pointsPerInch * SCALE_MULTIPLIER;
}

/**
 * Fits an image into the card image area while preserving aspect ratio.
 *
 * @param {number} nativeW
 * @param {number} nativeH
 * @returns {{ w: number, h: number }}
 */
function fitImageToCard_(nativeW, nativeH) {
  const fitScale = Math.min(MB_CELL_W / nativeW, MB_IMG_MAX_H / nativeH);
  return {
    w: Math.round(nativeW * fitScale),
    h: Math.round(nativeH * fitScale),
  };
}

/**
 * Computes image display dimensions that preserve aspect ratio and fit within
 * the card image area (MB_CELL_W × MB_IMG_MAX_H).
 *
 * When scale metadata is present, a shared points-per-inch factor is applied so
 * relative proportions are preserved across the entire bundle slide.
 * If neither scale is set, the image simply fits within the card bounds.
 *
 * @param {number|null} vscale  Real-world height in inches the image represents
 * @param {number|null} hscale  Real-world width in inches the image represents
 * @param {number}      nativeW Natural image width in points (from insertImage)
 * @param {number}      nativeH Natural image height in points (from insertImage)
 * @param {number}      pointsPerInch Shared bundle scale factor in slide points
 * @returns {{ w: number, h: number }}
 */
function computeImageSize_(vscale, hscale, nativeW, nativeH, pointsPerInch) {
  const realSize = getRealWorldSize_(vscale, hscale, nativeW, nativeH);
  if (!realSize) return fitImageToCard_(nativeW, nativeH);

  return {
    w: Math.round(realSize.w * pointsPerInch),
    h: Math.round(realSize.h * pointsPerInch),
  };
}

/**
 * Writes styled multi-line label text into a text box.
 *   Line 1: category name — 8 pt, italic, gray
 *   Line 2: product name  — 10 pt, bold, dark, hyperlinked if URL available
 *   Line 3: product size  — 8 pt, gray (omitted if not set)
 *
 * @param {SlidesApp.Shape} tb
 * @param {string}          cat  Category label
 * @param {Object|null}     mat  Material entry (may be null for empty categories)
 */
function setLabelText_(tb, cat, mat) {
  const productName = (mat && mat.product_name) ? mat.product_name : '\u2014';
  const productSize = (mat && mat.product_size) ? mat.product_size : null;
  const productUrl  = (mat && mat.product_url)  ? mat.product_url  : null;

  const lines = [cat, productName];
  if (productSize) lines.push(productSize);

  const text = tb.getText();
  text.setText(lines.join('\n'));

  // Category: small italic gray
  text.getRange(0, cat.length).getTextStyle()
    .setFontSize(8).setItalic(true).setForegroundColor('#888888');

  // Product name: bold dark, optionally hyperlinked
  const nameStart = cat.length + 1;
  const nameEnd   = nameStart + productName.length;
  const nameStyle = text.getRange(nameStart, nameEnd).getTextStyle()
    .setFontSize(10).setBold(true).setForegroundColor('#1a1a2e');
  if (productUrl) nameStyle.setLinkUrl(productUrl);

  // Size: small gray
  if (productSize) {
    const sizeStart = nameEnd + 1;
    text.getRange(sizeStart, sizeStart + productSize.length).getTextStyle()
      .setFontSize(8).setForegroundColor('#777777');
  }
}

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
 * Shows an alert via SlidesApp.getUi() and returns false if any are missing.
 * @param {Object.<string, number>} colMap
 * @param {string[]}                required
 * @returns {boolean}
 */
function validateMbCols_(colMap, required) {
  const missing = required.filter(h => colMap[h] === undefined);
  if (missing.length === 0) return true;
  SlidesApp.getUi().alert(
    'Sync Mood Board — Missing Column(s)',
    'The following headers were not found in row 1 of the Supplier sheet:\n\n  ' +
    missing.join(', ') +
    '\n\nCheck that row 1 contains the exact header names listed above.',
    SlidesApp.getUi().ButtonSet.OK
  );
  return false;
}

/**
 * Returns the first element in the array whose title matches, or null.
 * @param {PageElement[]} elements
 * @param {string}        title
 * @returns {PageElement|null}
 */
function findElByTitle_(elements, title) {
  for (let i = 0; i < elements.length; i++) {
    try { if (elements[i].getTitle() === title) return elements[i]; } catch (_) {}
  }
  return null;
}

/**
 * Finds a material visual element (real image or placeholder shape) for the
 * given category. Checks slide images first, then shapes.
 * @param {SlidesApp.Slide} slide
 * @param {string}          cat
 * @returns {PageElement|null}
 */
function findMaterialEl_(slide, cat) {
  const img = findElByTitle_(slide.getImages(), cat);
  if (img) return img;
  return findElByTitle_(slide.getShapes(), cat);
}

/**
 * Inserts a solid-color placeholder rectangle and returns the shape.
 * @returns {SlidesApp.Shape}
 */
function insertMbPlaceholder_(slide, x, y, w, h, hexColor) {
  const rect = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y, w, h);
  rect.getFill().setSolidFill(hexColor);
  rect.getBorder().setTransparent();
  return rect;
}
