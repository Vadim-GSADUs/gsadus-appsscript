function refreshProposalsFromDrive() {
  const ss    = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(CONFIG.SHEET_PROPOSALS);
  if (!sheet) throw new Error('Sheet "' + CONFIG.SHEET_PROPOSALS + '" not found.');

  // Ensure header row is present (A1:E1)
  const headerRange = sheet.getRange(1, 1, 1, 5);
  const headerValues = headerRange.getValues()[0];
  const expectedHeader = [
    CONFIG.PROP_COL.KEY,
    CONFIG.PROP_COL.PROPOSAL,
    CONFIG.PROP_COL.URL,
    CONFIG.PROP_COL.NAME,
    CONFIG.PROP_COL.STREET
  ];

  let headerChanged = false;
  for (let i = 0; i < expectedHeader.length; i++) {
    if (headerValues[i] !== expectedHeader[i]) {
      headerValues[i] = expectedHeader[i];
      headerChanged = true;
    }
  }
  if (headerChanged) {
    headerRange.setValues([headerValues]);
  }

  // Clear ONLY existing data rows (row 2 down), keep header & formatting
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 5).clearContent();
  }

  // Build fresh rows from Drive
  const root = DriveApp.getFolderById(CONFIG.ROOT_PROPOSAL_FOLDER_ID);
  const it   = root.getFolders();
  const rows = [];

  while (it.hasNext()) {
    const f          = it.next();
    const name       = f.getName();
    const url        = f.getUrl();
    const proposal   = extractProposalNumber_(name);
    const streetOnly = extractStreetFromFolderName_(name);

    // Compute numeric Key from Proposal (e.g. "PP12" -> 12)
    const key = (proposal && typeof proposal === 'string')
      ? (proposal.replace(/[^0-9]/g, '') || '')
      : '';

    // If key is numeric string, convert to number for proper VALUE semantics
    const keyValue = key === '' ? '' : Number(key);

    rows.push([
      keyValue,
      proposal,
      url,
      name,
      streetOnly
    ]);
  }

  // Write data rows contiguously from row 2
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, 5).setValues(rows);
  }

  // If the sheet previously had more data rows than we just wrote,
  // make sure there are no leftover values below our new data.
  const newLastRow = rows.length + 1; // header + data
  const maxRow      = sheet.getMaxRows();
  if (maxRow > newLastRow + 1) {
    // Clear contents in the rectangle below, but do not delete rows.
    sheet.getRange(newLastRow + 1, 1, maxRow - newLastRow, 5).clearContent();
  }
}

/** "PP12 3836 - 3838 Westporter Dr" -> "PP12" */
function extractProposalNumber_(folderName) {
  if (!folderName) return '';
  const m = folderName.match(/(PP\d+)/i);
  return m ? m[1].toUpperCase() : '';
}

/** "PP12 3836 - 3838 Westporter Dr" -> "3836 - 3838 Westporter Dr" */
function extractStreetFromFolderName_(folderName) {
  if (!folderName) return '';
  // Strip leading proposal number if present
  const cleaned = folderName.replace(/^(PP\d+\s*)/i, '').trim();
  return cleaned;
}
