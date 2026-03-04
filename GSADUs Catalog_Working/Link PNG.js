// ===== Link PNG.gs =====
// ────────── CONFIG ──────────
var CATALOG_SHEET = 'Catalog';           // sheet tab with your ADU_Catalog data
var KEY_COLUMN    = 'Model';             // unique key column header
var HEADER_ROW    = 2;                   // row number where headers live
var PNG_FOLDER_ID = '12pptHnUNdUL_nLJUAvt9NHX_SawvKJ57';  // Drive folder ID for floorplan PNGs
// ────────────────────────────

/**
 * Creates clickable hyperlinks in the "Floorplan_PNG" column
 * for each model whose PNG exists in the specified folder.
 */
function linkFloorplanPngs() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CATALOG_SHEET);
  if (!sheet) throw new Error('Sheet "' + CATALOG_SHEET + '" not found');

  // Fetch all PNG files and map by base filename (without extension)
  var folder = DriveApp.getFolderById(PNG_FOLDER_ID);
  var files  = folder.getFilesByType(MimeType.PNG);
  var urlMap = {};
  while (files.hasNext()) {
    var f    = files.next();
    var name = f.getName().replace(/\.png$/i, '');
    urlMap[name] = f.getUrl();
  }

  // Read header row to locate columns
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];
  var modelIdx = headers.indexOf(KEY_COLUMN);
  var pngIdx   = headers.indexOf('Floorplan_PNG');
  if (modelIdx < 0 || pngIdx < 0) {
    throw new Error('Ensure headers "' + KEY_COLUMN + '" and "Floorplan_PNG" exist in row ' + HEADER_ROW);
  }

  // Read all model keys from the sheet
  var lastRow   = sheet.getLastRow();
  var numModels = lastRow - HEADER_ROW;
  var models    = sheet.getRange(HEADER_ROW + 1, modelIdx + 1, numModels, 1).getValues();

  // Build hyperlink formulas for each model
  var linkFormulas = models.map(function(r) {
    var key = r[0];
    var url = urlMap[key] || '';
    return [ url ? '=HYPERLINK("' + url + '", "View PNG")' : '' ];
  });

  // Write all hyperlinks in one batch
  sheet.getRange(HEADER_ROW + 1, pngIdx + 1, linkFormulas.length, 1)
       .setValues(linkFormulas);

  // Notify user
  SpreadsheetApp.getUi().alert(
    'Linked ' + linkFormulas.filter(r => r[0]).length + ' floorplan PNGs.'
  );
}

/**
 * Adds the "Link Floorplan PNGs" menu item on open.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('ADU Catalog')
    .addItem('Link Floorplan PNGs','linkFloorplanPngs')
    .addToUi();
}
