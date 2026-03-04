// if your Config tab is named something else, change this
var CONFIG_SHEET_NAME = 'Config';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Publish')
    .addItem('Publish Catalog', 'publishCatalog')
    .addToUi();
}

function publishCatalog() {
  var ss     = SpreadsheetApp.getActive();
  var config = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!config) {
    var names = ss.getSheets().map(s=>s.getName()).join(', ');
    SpreadsheetApp.getUi().alert(
      '❌ Couldn’t find sheet “' + CONFIG_SHEET_NAME + '”.\n' +
      'Available tabs: ' + names
    );
    return;
  }

  // 1) Read your Sheets→Publish table
  var pubTbl = extractTable(config, 'Sheets', 'Publish');
  var toPublish = pubTbl
    .filter(r => r[1] === true)
    .map(r => r[0]);
  if (!toPublish.length) {
    SpreadsheetApp.getUi().alert('No sheets checked TRUE under "Publish".');
    return;
  }

  // 2) Read your Name→URL table
  var pathTbl = extractTable(config, 'Name', 'URL');
  var matchRow = pathTbl.find(r => r[0] === 'Path_GSADUs_Catalog_Published');
  if (!matchRow) {
    SpreadsheetApp.getUi().alert(
      'Couldn’t find "Path_GSADUs_Catalog_Published" in Name→URL table.'
    );
    return;
  }
  var folderUrl = matchRow[1];
  var idMatch   = folderUrl.match(/[-\w]{25,}/);
  if (!idMatch) {
    SpreadsheetApp.getUi().alert('Invalid Folder URL:\n' + folderUrl);
    return;
  }
  var folderId = idMatch[0];

  // —— BEFORE CREATING: delete any existing “GSADUs Catalog Published” files in that folder ——
  var folder = DriveApp.getFolderById(folderId);
  var existing = folder.getFilesByName('GSADUs Catalog Published');
  while (existing.hasNext()) {
    var oldFile = existing.next();
    oldFile.setTrashed(true);
  }

  // 3) Create the new “GSADUs Catalog Published” spreadsheet
  var newSS   = SpreadsheetApp.create('GSADUs Catalog Published');
  var newFile = DriveApp.getFileById(newSS.getId());

  // 4) Copy only the approved sheets
  toPublish.forEach(function(name){
    var sh = ss.getSheetByName(name);
    if (sh) sh.copyTo(newSS).setName(name);
  });
  // remove the default “Sheet1” or any extras
  newSS.getSheets().forEach(function(sh){
    if (toPublish.indexOf(sh.getName()) === -1) {
      newSS.deleteSheet(sh);
    }
  });

  // —— STRIP OUT embedded drawings/images (buttons) ——
  newSS.getSheets().forEach(function(sh) {
    if (typeof sh.getDrawings === 'function') {
      sh.getDrawings().forEach(function(d){ d.remove(); });
    }
    if (typeof sh.getImages === 'function') {
      sh.getImages().forEach(function(img){ img.remove(); });
    }
  });

  // 5) Move into the Shared-drive folder by ID
  newFile.moveTo(folder);

  SpreadsheetApp.getUi().alert(
    '✅ Published to "' + newSS.getName() + '"' +
    '\nin folder:\n' + folderUrl
  );
}


/**
 * extractTable(sheet, h1, h2)
 *  • Finds headers h1 and h2 anywhere in row 1
 *  • Returns every [cellUnderH1, cellUnderH2] down until the
 *    first row where both are blank.
 */
function extractTable(sheet, h1, h2) {
  var data    = sheet
                 .getRange(1,1, sheet.getLastRow(), sheet.getLastColumn())
                 .getValues();
  var headers = data[0];
  var c1 = headers.indexOf(h1), c2 = headers.indexOf(h2);
  if (c1 < 0 || c2 < 0) {
    throw new Error('Missing headers "'+h1+'" and/or "'+h2+'" in row 1.');
  }
  var out = [];
  for (var i = 1; i < data.length; i++) {
    var v1 = data[i][c1], v2 = data[i][c2];
    if ((v1 === null || v1 === '') && (v2 === null || v2 === '')) {
      break;
    }
    out.push([v1, v2]);
  }
  return out;
}
