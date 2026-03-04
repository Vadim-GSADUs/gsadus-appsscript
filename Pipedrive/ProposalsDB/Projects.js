// ---------------------------
// Projects - Drive folder index for project folders
// ---------------------------

/**
 * Refresh the Projects sheet from Drive folder structure.
 * Scans ROOT_PROJECT_FOLDER_ID and builds table with project metadata.
 * Links projects to proposals by matching street addresses.
 */
function refreshProjectsFromDrive() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(CONFIG.SHEET_PROJECTS);
  
  // Create sheet if it doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_PROJECTS);
  }

  // Ensure header row is present (A1:F1) - 6 columns
  const headerRange = sheet.getRange(1, 1, 1, 6);
  const headerValues = headerRange.getValues()[0];
  const expectedHeader = [
    CONFIG.PROJ_COL.KEY,
    CONFIG.PROJ_COL.PROJECT,
    CONFIG.PROJ_COL.URL,
    CONFIG.PROJ_COL.NAME,
    CONFIG.PROJ_COL.STREET,
    CONFIG.PROJ_COL.PROPOSAL
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
    sheet.getRange(2, 1, lastRow - 1, 6).clearContent();
  }

  // Build fresh rows from Drive
  const root = DriveApp.getFolderById(CONFIG.ROOT_PROJECT_FOLDER_ID);
  const it = root.getFolders();
  const rows = [];

  while (it.hasNext()) {
    const f = it.next();
    const name = f.getName();
    const url = f.getUrl();
    const project = extractProjectNumber_(name);
    const streetOnly = extractStreetFromProjectFolderName_(name);

    // Compute numeric Key from Project # (e.g. "P45" -> 45)
    const key = (project && typeof project === 'string')
      ? (project.replace(/[^0-9]/g, '') || '')
      : '';

    // If key is numeric string, convert to number for proper VALUE semantics
    const keyValue = key === '' ? '' : Number(key);

    rows.push([
      keyValue,
      project,
      url,
      name,
      streetOnly,
      '' // PP# will be filled in next step via matching
    ]);
  }

  // Write data rows contiguously from row 2
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, 6).setValues(rows);
  }

  // If the sheet previously had more data rows than we just wrote,
  // make sure there are no leftover values below our new data.
  const newLastRow = rows.length + 1; // header + data
  const maxRow = sheet.getMaxRows();
  if (maxRow > newLastRow + 1) {
    // Clear contents in the rectangle below, but do not delete rows.
    sheet.getRange(newLastRow + 1, 1, maxRow - newLastRow, 6).clearContent();
  }

  // Now link Projects to Proposals by matching street addresses
  linkProjectsToProposals_();
  
  Logger.log('refreshProjectsFromDrive: Complete. Projects written: ' + rows.length);
}

/**
 * Extract project number from folder name.
 * "P22 525 Amberly Ct" -> "P22"
 * "P0 0000 Address Ln. - Project Template" -> "P0"
 * @private
 */
function extractProjectNumber_(folderName) {
  if (!folderName) return '';
  const m = folderName.match(/(P\d+)/i);
  return m ? m[1].toUpperCase() : '';
}

/**
 * Extract street address from project folder name.
 * "P22 525 Amberly Ct" -> "525 Amberly Ct"
 * Strips leading project number if present.
 * @private
 */
function extractStreetFromProjectFolderName_(folderName) {
  if (!folderName) return '';
  // Strip leading project number (P followed by digits and optional space)
  const cleaned = folderName.replace(/^(P\d+\s*)/i, '').trim();
  // Remove template suffix if present
  return cleaned.replace(/\s*-\s*Project Template$/i, '').trim();
}

/**
 * Link Projects to Proposals by matching street addresses.
 * Populates PP# column in Projects sheet based on Street Only matches.
 * @private
 */
function linkProjectsToProposals_() {
  const ss = SpreadsheetApp.getActive();
  const projectsSheet = ss.getSheetByName(CONFIG.SHEET_PROJECTS);
  const proposalsSheet = ss.getSheetByName(CONFIG.SHEET_PROPOSALS);

  if (!projectsSheet || !proposalsSheet) {
    Logger.log('linkProjectsToProposals_: Required sheets not found');
    return;
  }

  // Read Proposals data to build street -> PP# map
  const proposalsData = proposalsSheet.getDataRange().getValues();
  if (proposalsData.length <= 1) {
    Logger.log('linkProjectsToProposals_: No proposals data to link');
    return;
  }

  const proposalsHeaders = proposalsData[0];
  const propStreetIdx = proposalsHeaders.indexOf(CONFIG.PROP_COL.STREET);
  const propPPIdx = proposalsHeaders.indexOf(CONFIG.PROP_COL.PROPOSAL);

  if (propStreetIdx === -1 || propPPIdx === -1) {
    Logger.log('linkProjectsToProposals_: Required columns not found in Proposals');
    return;
  }

  // Build map: normalized street -> PP#
  const streetToPP = {};
  for (let i = 1; i < proposalsData.length; i++) {
    const street = String(proposalsData[i][propStreetIdx] || '').trim();
    const pp = proposalsData[i][propPPIdx];
    
    if (street && pp) {
      const normalizedStreet = normalizeStreetAddress_(street);
      // Store the PP# (last one wins if duplicates)
      streetToPP[normalizedStreet] = pp;
    }
  }

  // Read Projects data
  const projectsData = projectsSheet.getDataRange().getValues();
  if (projectsData.length <= 1) {
    Logger.log('linkProjectsToProposals_: No projects data to link');
    return;
  }

  const projectsHeaders = projectsData[0];
  const projStreetIdx = projectsHeaders.indexOf(CONFIG.PROJ_COL.STREET);
  const projPPIdx = projectsHeaders.indexOf(CONFIG.PROJ_COL.PROPOSAL);

  if (projStreetIdx === -1 || projPPIdx === -1) {
    Logger.log('linkProjectsToProposals_: Required columns not found in Projects');
    return;
  }

  // Update PP# column in Projects based on street match
  let matchCount = 0;
  for (let i = 1; i < projectsData.length; i++) {
    const street = String(projectsData[i][projStreetIdx] || '').trim();
    
    if (!street) continue;
    
    const normalizedStreet = normalizeStreetAddress_(street);
    const matchedPP = streetToPP[normalizedStreet];
    
    if (matchedPP) {
      // Write the matched PP# to the cell
      projectsSheet.getRange(i + 1, projPPIdx + 1).setValue(matchedPP);
      matchCount++;
    }
  }

  Logger.log('linkProjectsToProposals_: Matched ' + matchCount + ' projects to proposals');
}

/**
 * Normalize street address for matching.
 * Removes common variations, case differences, extra spaces.
 * @private
 */
function normalizeStreetAddress_(street) {
  if (!street) return '';
  
  let normalized = String(street).trim().toLowerCase();
  
  // Standardize common abbreviations
  normalized = normalized
    .replace(/\bstreet\b/g, 'st')
    .replace(/\bstr\b/g, 'st')
    .replace(/\bdrive\b/g, 'dr')
    .replace(/\bavenue\b/g, 'ave')
    .replace(/\bav\b/g, 'ave')
    .replace(/\broad\b/g, 'rd')
    .replace(/\blane\b/g, 'ln')
    .replace(/\bcourt\b/g, 'ct')
    .replace(/\bcircle\b/g, 'cir')
    .replace(/\bway\b/g, 'way')
    .replace(/\bcrescent\b/g, 'crescent');
  
  // Remove extra spaces, punctuation, and normalize
  normalized = normalized
    .replace(/[.,\/#!$%\^&\*;:{}=\-_`~()]/g, ' ') // Replace punctuation with space
    .replace(/\s+/g, ' ') // Collapse multiple spaces
    .trim();
  
  return normalized;
}

/**
 * Convenience function to refresh both Proposals and Projects.
 * Called from "Pull From Drive" menu item.
 */
function refreshAllDriveFolders() {
  refreshProposalsFromDrive();
  refreshProjectsFromDrive();
  
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Drive Refresh Complete',
    'Updated both Proposals and Projects sheets from Drive.',
    ui.ButtonSet.OK
  );
}
