/**
 * FOLDER DISCOVERY ENGINE
 * Scans root Drive folder for new project folders and adds them to "Folders" sheet.
 * Optimized for timed triggers - minimal API calls, batch operations.
 */

const DISCOVERY_CONFIG = {
  ROOT_FOLDER_ID: "1NOLgBO5xZu4EXFJ0PuS56Xa_cQ-SEvpW",
  SHEET_NAME: "Folders",
  STATUSES_SHEET_NAME: "Statuses",
  // Regex to extract Project ID (e.g., "P161" from "P161 1590 3rd Ave")
  PROJECT_ID_REGEX: /^(P\d+)\s+/i,
  // Folder IDs to ignore (e.g., templates)
  IGNORED_FOLDER_IDS: [
    "1NuPtrNKKkULG4p_ReTW3m7qkhOUZrRmf" // P0 - Project File Template
  ],
  // Statuses tab column indices (0-based)
  STATUSES_PREFIX_COL: 0,  // Column A - "P" or "PP"
  STATUSES_NUMBER_COL: 1,  // Column B - number
  STATUSES_VALUE_COL: 8    // Column I - Status value
};

/**
 * Main discovery function - safe for timed triggers.
 * Finds new project folders and adds them to the Folders sheet.
 * @param {boolean} silent - If true, skips UI alerts (for triggers)
 * @returns {Object} Results summary
 */
function discoverNewProjects(silent) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DISCOVERY_CONFIG.SHEET_NAME);

  if (!sheet) {
    if (!silent) SpreadsheetApp.getUi().alert('Sheet "Folders" not found.');
    return { error: "Sheet not found" };
  }

  // 1. Get existing folder URLs from sheet (single read)
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const urlIdx = headers.indexOf("Folder URL");
  const pIdIdx = headers.indexOf("Project ID");
  const nameIdx = headers.indexOf("Full Folder Name");
  const addressIdx = headers.indexOf("Address");

  if (urlIdx < 0) {
    if (!silent) SpreadsheetApp.getUi().alert('Column "Folder URL" not found.');
    return { error: "Folder URL column not found" };
  }

  // Build Set of existing folder IDs for O(1) lookup
  const existingFolderIds = new Set();
  for (let r = 1; r < data.length; r++) {
    const url = String(data[r][urlIdx] || "").trim();
    const folderId = extractFolderIdFromUrl_(url);
    if (folderId) existingFolderIds.add(folderId);
  }

  // 2. Get all subfolders from root (single API call via iterator)
  const rootFolder = DriveApp.getFolderById(DISCOVERY_CONFIG.ROOT_FOLDER_ID);
  const folderIterator = rootFolder.getFolders();
  const ignoredIds = new Set(DISCOVERY_CONFIG.IGNORED_FOLDER_IDS);

  const newFolders = [];

  while (folderIterator.hasNext()) {
    const folder = folderIterator.next();
    const folderId = folder.getId();

    // Skip ignored folders (e.g., templates)
    if (ignoredIds.has(folderId)) continue;

    // Skip if already in sheet
    if (existingFolderIds.has(folderId)) continue;

    const folderName = folder.getName();
    const folderUrl = folder.getUrl();

    // Extract Project ID and Address from folder name
    const parsed = parseFolderName_(folderName);

    newFolders.push({
      folderId: folderId,
      folderName: folderName,
      folderUrl: folderUrl,
      projectId: parsed.projectId,
      address: parsed.address
    });
  }

  // Sort alpha-numerically by folder name (matches Drive sort order)
  newFolders.sort((a, b) => a.folderName.localeCompare(b.folderName, undefined, { numeric: true, sensitivity: 'base' }));

  // 3. Batch write new rows (single write operation)
  if (newFolders.length > 0) {
    const newRows = newFolders.map(f => {
      const row = new Array(headers.length).fill("");

      if (pIdIdx >= 0) row[pIdIdx] = f.projectId;
      if (nameIdx >= 0) row[nameIdx] = f.folderName;
      if (urlIdx >= 0) row[urlIdx] = f.folderUrl;
      if (addressIdx >= 0) row[addressIdx] = f.address;
      // Statuses column left blank - driven by "Statuses" tab

      return row;
    });

    // Append all new rows at once
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, newRows.length, headers.length).setValues(newRows);
  }

  const result = {
    scanned: existingFolderIds.size + newFolders.length,
    existing: existingFolderIds.size,
    added: newFolders.length,
    newProjects: newFolders.map(f => f.projectId || f.folderName)
  };

  if (!silent) {
    if (newFolders.length === 0) {
      SpreadsheetApp.getUi().alert("No new projects found.\n\nScanned: " + result.scanned + " folders");
    } else {
      SpreadsheetApp.getUi().alert(
        "Discovery Complete!\n\n" +
        "Added: " + result.added + " new project(s)\n" +
        "Total folders: " + result.scanned + "\n\n" +
        "New projects:\n" + result.newProjects.join("\n")
      );
    }
  }

  return result;
}

/**
 * Wrapper for manual menu trigger (shows UI)
 */
function discoverNewProjectsManual() {
  discoverNewProjects(false);
}

/**
 * Wrapper for timed trigger (silent, no UI)
 */
function discoverNewProjectsTrigger() {
  const result = discoverNewProjects(true);
  Logger.log("Folder Discovery: " + JSON.stringify(result));
  return result;
}

/**
 * Parse folder name to extract Project ID and Address.
 * Expected format: "P### Address" (e.g., "P161 1590 3rd Ave")
 */
function parseFolderName_(folderName) {
  const name = String(folderName || "").trim();
  const match = name.match(DISCOVERY_CONFIG.PROJECT_ID_REGEX);

  if (match) {
    return {
      projectId: match[1].toUpperCase(),
      address: name.substring(match[0].length).trim()
    };
  }

  // Fallback: use full name as project ID if no pattern match
  return {
    projectId: name,
    address: ""
  };
}

/**
 * Extract folder ID from various Google Drive URL formats.
 */
function extractFolderIdFromUrl_(url) {
  const s = String(url || "");

  // Format: /folders/<ID>
  const m1 = s.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (m1 && m1[1]) return m1[1];

  // Format: ?id=<ID>
  const m2 = s.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (m2 && m2[1]) return m2[1];

  // Already an ID
  if (/^[a-zA-Z0-9_-]{10,}$/.test(s)) return s;

  return null;
}

/**
 * Syncs statuses from the "Statuses" tab to the "Folders" tab.
 * Matches by Project ID (Statuses: ColA + ColB = Folders: Project ID)
 * @param {boolean} silent - If true, skips UI alerts (for triggers)
 * @returns {Object} Results summary
 */
function syncStatusesFromTab(silent) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const foldersSheet = ss.getSheetByName(DISCOVERY_CONFIG.SHEET_NAME);
  const statusesSheet = ss.getSheetByName(DISCOVERY_CONFIG.STATUSES_SHEET_NAME);

  if (!foldersSheet) {
    if (!silent) SpreadsheetApp.getUi().alert('Sheet "Folders" not found.');
    return { error: "Folders sheet not found" };
  }

  if (!statusesSheet) {
    if (!silent) SpreadsheetApp.getUi().alert('Sheet "Statuses" not found.');
    return { error: "Statuses sheet not found" };
  }

  // 1. Build status map from Statuses tab (Project ID -> Status)
  const statusData = statusesSheet.getDataRange().getValues();
  const statusMap = new Map();

  for (let r = 1; r < statusData.length; r++) {
    const row = statusData[r];
    const prefix = String(row[DISCOVERY_CONFIG.STATUSES_PREFIX_COL] || "").trim();
    const number = String(row[DISCOVERY_CONFIG.STATUSES_NUMBER_COL] || "").trim();
    const status = String(row[DISCOVERY_CONFIG.STATUSES_VALUE_COL] || "").trim();

    if (!prefix || !number) continue;

    const projectId = (prefix + number).toUpperCase();
    statusMap.set(projectId, status);
  }

  // 2. Read Folders tab and update statuses
  const foldersData = foldersSheet.getDataRange().getValues();
  const headers = foldersData[0];

  const pIdIdx = headers.indexOf("Project ID");
  const statusIdx = headers.indexOf("Statuses");

  if (pIdIdx < 0) {
    if (!silent) SpreadsheetApp.getUi().alert('Column "Project ID" not found in Folders.');
    return { error: "Project ID column not found" };
  }

  if (statusIdx < 0) {
    if (!silent) SpreadsheetApp.getUi().alert('Column "Statuses" not found in Folders.');
    return { error: "Statuses column not found" };
  }

  let updated = 0;
  let notFound = 0;
  let unchanged = 0;

  // 3. Update statuses in memory
  for (let r = 1; r < foldersData.length; r++) {
    const projectId = String(foldersData[r][pIdIdx] || "").trim().toUpperCase();
    if (!projectId) continue;

    const newStatus = statusMap.get(projectId);

    if (newStatus === undefined) {
      notFound++;
      continue;
    }

    const currentStatus = String(foldersData[r][statusIdx] || "").trim();

    if (currentStatus === newStatus) {
      unchanged++;
      continue;
    }

    foldersData[r][statusIdx] = newStatus;
    updated++;
  }

  // 4. Batch write if any updates
  if (updated > 0) {
    foldersSheet.getRange(2, 1, foldersData.length - 1, headers.length)
      .setValues(foldersData.slice(1));
  }

  const result = {
    updated: updated,
    unchanged: unchanged,
    notFound: notFound,
    totalStatusEntries: statusMap.size
  };

  if (!silent) {
    SpreadsheetApp.getUi().alert(
      "Status Sync Complete!\n\n" +
      "Updated: " + result.updated + "\n" +
      "Unchanged: " + result.unchanged + "\n" +
      "Not found in Statuses tab: " + result.notFound
    );
  }

  return result;
}

/**
 * Combined function: Discover new projects, then sync statuses.
 * Ideal for timed triggers.
 * @param {boolean} silent - If true, skips UI alerts
 * @returns {Object} Combined results
 */
function discoverAndSyncStatuses(silent) {
  const discoveryResult = discoverNewProjects(true);
  const statusResult = syncStatusesFromTab(true);

  const result = {
    discovery: discoveryResult,
    statusSync: statusResult
  };

  if (!silent) {
    SpreadsheetApp.getUi().alert(
      "Discovery & Status Sync Complete!\n\n" +
      "--- New Projects ---\n" +
      "Added: " + (discoveryResult.added || 0) + "\n\n" +
      "--- Status Updates ---\n" +
      "Updated: " + (statusResult.updated || 0) + "\n" +
      "Unchanged: " + (statusResult.unchanged || 0)
    );
  }

  Logger.log("Discover & Sync: " + JSON.stringify(result));
  return result;
}

/**
 * Manual menu wrapper for status sync
 */
function syncStatusesManual() {
  syncStatusesFromTab(false);
}

/**
 * Manual menu wrapper for combined discover + status sync
 */
function discoverAndSyncManual() {
  discoverAndSyncStatuses(false);
}

/**
 * Timed trigger wrapper for combined discover + status sync
 */
function discoverAndSyncTrigger() {
  return discoverAndSyncStatuses(true);
}
