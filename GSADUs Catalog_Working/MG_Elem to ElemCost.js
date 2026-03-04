function importNewElemCostEntries() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("MG_Elem");
  const targetSheet = ss.getSheetByName("Costs");

  if (!sourceSheet || !targetSheet) {
    SpreadsheetApp.getUi().alert("Sheet 'MG_Elem' or 'Costs' not found.");
    return;
  }

  // Source: all data
  const sourceData = sourceSheet.getDataRange().getValues();
  const sourceHeaders = sourceData[0];

  // Target: starts at row 3
  const targetStartRow = 3;
  const targetDataRange = targetSheet.getRange(targetStartRow, 1, targetSheet.getLastRow() - 2, targetSheet.getLastColumn());
  const targetData = targetDataRange.getValues();
  const targetHeaders = targetData[0];

  // Column indices
  const sCategory = sourceHeaders.indexOf("Category");
  const sFamily = sourceHeaders.indexOf("Family");
  const sTypeName = sourceHeaders.indexOf("Type Name");

  const tCategory = targetHeaders.indexOf("Category");
  const tFamily = targetHeaders.indexOf("Family");
  const tTypeName = targetHeaders.indexOf("Type Name");

  if (sCategory === -1 || sFamily === -1 || sTypeName === -1 ||
      tCategory === -1 || tFamily === -1 || tTypeName === -1) {
    SpreadsheetApp.getUi().alert("Missing required columns (Category, Family, Type Name).");
    return;
  }

  // Build unique keys from MG_Elem (deduplicate BEFORE checking target)
  const uniqueSourceKeys = new Set();
  const uniqueSourceEntries = [];

  for (let i = 1; i < sourceData.length; i++) {
    const row = sourceData[i];
    const key = [row[sCategory], row[sFamily], row[sTypeName]].join("|");
    if (!uniqueSourceKeys.has(key)) {
      uniqueSourceKeys.add(key);
      uniqueSourceEntries.push([row[sCategory], row[sFamily], row[sTypeName]]);
    }
  }

  // Existing keys in ElemCost
  const existingKeys = new Set(
    targetData.slice(1).map(row =>
      [row[tCategory], row[tFamily], row[tTypeName]].join("|")
    )
  );

  // Filter to only new keys
  const rowsToInsert = uniqueSourceEntries.filter(entry => {
    const key = entry.join("|");
    return !existingKeys.has(key);
  }).map(entry => {
    const totalCols = targetHeaders.length;
    return entry.concat(Array(totalCols - entry.length).fill(""));
  });

  if (rowsToInsert.length === 0) {
    SpreadsheetApp.getUi().alert("No new unique entries to add.");
    return;
  }

  // Append new entries
  const nextRow = targetSheet.getLastRow() + 1;
  targetSheet.getRange(nextRow, 1, rowsToInsert.length, targetHeaders.length).setValues(rowsToInsert);

  SpreadsheetApp.getUi().alert(`${rowsToInsert.length} unique entries added to 'ElemCost'.`);
}
