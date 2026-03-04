/**
 * GSADUs Smart Sync Engine 2.0
 * - Batch in-memory processing
 * - Shadow-ledger persistence fixes (prevents perpetual "2 files updated")
 * - Push (Sheet -> Drive) and Pull (Drive -> Sheet)
 * - Detailed HTML report
 */

const CONFIG = {
  SHEET_NAME: "Folders",
  SHADOW_SHEET_NAME: "Folders_Shadow",
  CONFIG_SHEET_NAME: "Config",
  FILENAME: "Project Info.csv",
  SUBFOLDER_NAME: "2. Supporting Documents",
  STATUS_COL_NAME: "Statuses",

  // Treat these as NOT active if they ever appear:
  INACTIVE_STATUSES: ["completed", "closed", "archived", "inactive", "cancelled", "canceled"]
};

/**
 * Public menu wrappers
 */
function pushActiveProjects() { showCombinedReport([], pushDataUpdates(true, true)); }
function pushAllProjects()    { showCombinedReport([], pushDataUpdates(true, false)); }
function pullActiveProjects() { showCombinedReport(pullProjectUpdates(true, true), []); }
function pullAllProjects()    { showCombinedReport(pullProjectUpdates(true, false), []); }

function syncActiveProjects() {
  const pullResults = pullProjectUpdates(true, true);
  const pushResults = pushDataUpdates(true, true);
  showCombinedReport(pullResults, pushResults);
}

function syncAllProjects() {
  const pullResults = pullProjectUpdates(true, false);
  const pushResults = pushDataUpdates(true, false);
  showCombinedReport(pullResults, pushResults);
}

/**
 * PUSH: Sheet -> Drive CSV
 * silent=true means do not show UI inside this function (caller handles reporting)
 * activeOnly=true filters by Statuses
 */
function pushDataUpdates(silent, activeOnly) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30 * 1000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) throw new Error(`Sheet not found: ${CONFIG.SHEET_NAME}`);

    const shadowSheet = getOrInitShadowSheet(ss);

    // Bulk load
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    const headers = data[0];

    const pIdIdx = headers.indexOf("Project ID");
    const urlIdx = headers.indexOf("Folder URL");
    const statusIdx = headers.indexOf(CONFIG.STATUS_COL_NAME);

    if (pIdIdx < 0) throw new Error(`Column not found: Project ID`);
    if (urlIdx < 0) throw new Error(`Column not found: Folder URL`);

    // Config map (which headers to include + how to normalize)
    const configMap = getConfigMap(ss); // key: sheet header, value: {csvLabel, type, mode}

    // Build shadow map (Project ID -> state)
    const shadowMap = readShadowMap_(shadowSheet);

    const now = new Date();
    const results = [];
    const newShadow = new Map(shadowMap); // copy so we can persist safely

    // Process rows in memory; only do Drive I/O when needed
    for (let r = 1; r < data.length; r++) {
      const row = data[r];
      const pId = String(row[pIdIdx] || "").trim();
      if (!pId) continue;

      if (activeOnly && !isActiveProject_(row, statusIdx)) {
        results.push({ id: pId, type: "PUSH", status: "Skipped", details: "Inactive (filtered)" });
        continue;
      }

      const folderUrl = String(row[urlIdx] || "").trim();
      if (!folderUrl) {
        results.push({ id: pId, type: "PUSH", status: "Error", details: "Missing Folder URL" });
        continue;
      }

      const payload = generateVerticalCsvPayload_(row, headers, configMap);
      const newHash = md5Hex_(payload.csvContent);

      const prev = shadowMap.get(pId) || { fileId: "", hash: "", lastSynced: null };
      let currentFileId = prev.fileId || "";

      const needsUpdate = (!prev.hash || prev.hash !== newHash || !currentFileId);

      if (!needsUpdate) {
        results.push({ id: pId, type: "PUSH", status: "Skipped", details: "No changes" });
        continue;
      }

      try {
        const file = updateOrCreateFile_(pId, folderUrl, payload.csvContent, currentFileId);
        currentFileId = file.getId();

        newShadow.set(pId, {
          fileId: currentFileId,
          hash: newHash,
          lastSynced: now
        });

        results.push({
          id: pId,
          type: "PUSH",
          status: "Updated",
          details: prev.fileId ? "Data changed" : "Link repaired / created"
        });

      } catch (e) {
        results.push({ id: pId, type: "PUSH", status: "Error", details: String(e.message || e) });
      }
    }

    // Persist shadow in one write (fixes phantom loop)
    writeShadowMap_(shadowSheet, newShadow);

    if (!silent) showCombinedReport([], results);
    return results;

  } finally {
    lock.releaseLock();
  }
}

/**
 * PULL: Drive CSV -> Sheet
 * - Uses file.getLastUpdated vs shadow.lastSynced (with buffer) to decide if a pull is needed
 * - Performs header-level diff and only writes sheet once (batch)
 */
function pullProjectUpdates(silent, activeOnly) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30 * 1000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) throw new Error(`Sheet not found: ${CONFIG.SHEET_NAME}`);

    const shadowSheet = getOrInitShadowSheet(ss);

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    const headers = data[0];

    const pIdIdx = headers.indexOf("Project ID");
    const statusIdx = headers.indexOf(CONFIG.STATUS_COL_NAME);
    if (pIdIdx < 0) throw new Error(`Column not found: Project ID`);

    const headerIndex = buildHeaderIndex_(headers);
    const shadowMap = readShadowMap_(shadowSheet);

    const results = [];
    const updatedRows = []; // {rowIndex0, newRowArray}
    const now = new Date();

    // 1-minute buffer to avoid clock skew / immediate push-after-pull
    const BUFFER_MS = 60 * 1000;

    for (let r = 1; r < data.length; r++) {
      const row = data[r];
      const pId = String(row[pIdIdx] || "").trim();
      if (!pId) continue;

      if (activeOnly && !isActiveProject_(row, statusIdx)) {
        results.push({ id: pId, type: "PULL", status: "Skipped", details: "Inactive (filtered)" });
        continue;
      }

      const shadow = shadowMap.get(pId);
      if (!shadow || !shadow.fileId) {
        results.push({ id: pId, type: "PULL", status: "Skipped", details: "No linked file" });
        continue;
      }

      try {
        const file = DriveApp.getFileById(shadow.fileId);
        const lastUpdated = file.getLastUpdated();
        const lastSynced = shadow.lastSynced ? new Date(shadow.lastSynced) : null;

        // If never synced, allow pull (rare)
        const threshold = lastSynced ? new Date(lastSynced.getTime() + BUFFER_MS) : new Date(0);

        if (lastUpdated <= threshold) {
          results.push({ id: pId, type: "PULL", status: "Skipped", details: "Sheet is up to date" });
          continue;
        }

        const csvText = file.getBlob().getDataAsString();
        const kv = parseVerticalCsvToMap_(csvText); // Field -> Value

        // Clone row for modifications
        const newRow = row.slice();
        const changedHeaders = [];

        // Apply CSV->Sheet updates for matching headers
        Object.keys(kv).forEach(field => {
          const colIdx = headerIndex[field];
          if (colIdx === undefined) return; // CSV label doesn't exist in sheet

          const oldVal = normalizeForCompare_(newRow[colIdx]);
          const newVal = normalizeForCompare_(kv[field]);

          if (oldVal !== newVal) {
            newRow[colIdx] = kv[field];
            changedHeaders.push(field);
          }
        });

        if (changedHeaders.length === 0) {
          results.push({ id: pId, type: "PULL", status: "Skipped", details: "Timestamp changed; content identical" });

          // Still update lastSynced so we don't keep re-checking this modification endlessly.
          // Hash is not updated here because Pull is not based on hash; Push will recompute.
          shadowMap.set(pId, { fileId: shadow.fileId, hash: shadow.hash, lastSynced: now });
          continue;
        }

        updatedRows.push({ rowIndex0: r, newRow });

        // Update shadow lastSynced after successful pull stage
        shadowMap.set(pId, { fileId: shadow.fileId, hash: shadow.hash, lastSynced: now });

        results.push({
          id: pId,
          type: "PULL",
          status: "Updated",
          details: `Headers: ${changedHeaders.join(", ")}`
        });

      } catch (e) {
        results.push({ id: pId, type: "PULL", status: "Error", details: String(e.message || e) });
      }
    }

    // Batch write sheet once
    if (updatedRows.length > 0) {
      // Modify in-memory data then write the whole body range once
      updatedRows.forEach(u => { data[u.rowIndex0] = u.newRow; });
      sheet.getRange(2, 1, data.length - 1, headers.length).setValues(data.slice(1));
    }

    // Persist shadow updates (lastSynced) in one write
    writeShadowMap_(shadowSheet, shadowMap);

    if (!silent) showCombinedReport(results, []);
    return results;

  } finally {
    lock.releaseLock();
  }
}

/**
 * Fix missing File IDs in shadow by locating the CSV in the project folder.
 * This repairs the "phantom updates" cause without pushing content changes.
 */
function repairPhantomLinks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  const shadowSheet = getOrInitShadowSheet(ss);

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    SpreadsheetApp.getUi().alert("No project rows found.");
    return;
  }

  const headers = data[0];
  const pIdIdx = headers.indexOf("Project ID");
  const urlIdx = headers.indexOf("Folder URL");
  if (pIdIdx < 0 || urlIdx < 0) throw new Error("Required columns missing: Project ID and/or Folder URL");

  const shadowMap = readShadowMap_(shadowSheet);
  let fixed = 0;
  let scanned = 0;

  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const pId = String(row[pIdIdx] || "").trim();
    if (!pId) continue;

    const st = shadowMap.get(pId);
    const needsRepair = (!st || !st.fileId);

    if (!needsRepair) continue;

    scanned++;
    const folderUrl = String(row[urlIdx] || "").trim();
    if (!folderUrl) continue;

    try {
      const file = findCsvInProjectFolder_(folderUrl);
      if (file) {
        shadowMap.set(pId, {
          fileId: file.getId(),
          hash: (st && st.hash) ? st.hash : "",
          lastSynced: (st && st.lastSynced) ? st.lastSynced : new Date(0)
        });
        fixed++;
      }
    } catch (e) {
      // ignore per-row errors; user can review logs if needed
      Logger.log(`Repair error for ${pId}: ${e.message || e}`);
    }
  }

  writeShadowMap_(shadowSheet, shadowMap);
  SpreadsheetApp.getUi().alert(`Repair complete.\nScanned: ${scanned}\nFixed: ${fixed}`);
}

/**
 * Rebuilds a basic Config tab using current Folders headers.
 * Keeps the real schema: Sheet Header / CSV Label / Sync Mode / Data Type
 */
function initializeConfigTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const folders = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!folders) throw new Error(`Sheet not found: ${CONFIG.SHEET_NAME}`);

  const headers = folders.getRange(1, 1, 1, folders.getLastColumn()).getValues()[0];

  let cfg = ss.getSheetByName(CONFIG.CONFIG_SHEET_NAME);
  if (!cfg) cfg = ss.insertSheet(CONFIG.CONFIG_SHEET_NAME);
  cfg.clear();

  cfg.getRange(1, 1, 1, 4).setValues([["Sheet Header", "CSV Label", "Sync Mode", "Data Type"]]);
  cfg.setFrozenRows(1);

  const rows = headers.map(h => {
    const header = String(h || "").trim();
    if (!header) return null;

    // Reasonable defaults:
    let mode = "MANUAL";
    if (header === "Project ID") mode = "SYSTEM";
    if (header === "Full Folder Name" || header === "Folder URL" || header === "Statuses") mode = "HIDDEN";

    // Data type guess:
    let type = "String";
    if (header.toLowerCase().includes("date")) type = "Date";
    if (header === "Folder URL") type = "URL";
    if (header === "Lat" || header === "Long") type = "Number";

    return [header, header, mode, type];
  }).filter(Boolean);

  if (rows.length) cfg.getRange(2, 1, rows.length, 4).setValues(rows);
}

/* -------------------------
 * Helpers / Internals
 * ------------------------- */

function getOrInitShadowSheet(ss) {
  let sh = ss.getSheetByName(CONFIG.SHADOW_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.SHADOW_SHEET_NAME);
    sh.getRange(1, 1, 1, 4).setValues([["Project ID", "CSV File ID", "Sync Hash", "Last Synced"]]);
    sh.setFrozenRows(1);
    sh.hideSheet();
  }
  return sh;
}

/**
 * Reads Config tab:
 * columns: Sheet Header | CSV Label | Sync Mode | Data Type
 * includes SYSTEM and MANUAL; excludes HIDDEN
 */
function getConfigMap(ss) {
  const cfgSheet = ss.getSheetByName(CONFIG.CONFIG_SHEET_NAME);
  const map = new Map();

  if (!cfgSheet) return map;

  const values = cfgSheet.getDataRange().getValues();
  if (values.length < 2) return map;

  const header = values[0].map(v => String(v || "").trim());
  const idxSheetHeader = header.indexOf("Sheet Header");
  const idxCsvLabel = header.indexOf("CSV Label");
  const idxMode = header.indexOf("Sync Mode");
  const idxType = header.indexOf("Data Type");

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const sheetHeader = idxSheetHeader >= 0 ? String(row[idxSheetHeader] || "").trim() : "";
    if (!sheetHeader) continue;

    const csvLabel = idxCsvLabel >= 0 ? String(row[idxCsvLabel] || "").trim() : sheetHeader;
    const mode = idxMode >= 0 ? String(row[idxMode] || "").trim().toUpperCase() : "MANUAL";
    const type = idxType >= 0 ? String(row[idxType] || "").trim() : "String";

    if (mode === "HIDDEN") continue;

    map.set(sheetHeader, { csvLabel, mode, type });
  }

  return map;
}

function readShadowMap_(shadowSheet) {
  const values = shadowSheet.getDataRange().getValues();
  const map = new Map();
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const pId = String(row[0] || "").trim();
    if (!pId) continue;
    map.set(pId, {
      fileId: String(row[1] || "").trim(),
      hash: String(row[2] || "").trim(),
      lastSynced: row[3] ? new Date(row[3]) : null
    });
  }
  return map;
}

function writeShadowMap_(shadowSheet, shadowMap) {
  const rows = [];
  shadowMap.forEach((v, k) => {
    rows.push([k, v.fileId || "", v.hash || "", v.lastSynced || ""]);
  });

  rows.sort((a, b) => String(a[0]).localeCompare(String(b[0])));

  // Clear existing body and write once
  const last = shadowSheet.getLastRow();
  if (last > 1) shadowSheet.getRange(2, 1, last - 1, 4).clearContent();

  if (rows.length > 0) shadowSheet.getRange(2, 1, rows.length, 4).setValues(rows);
}

function buildHeaderIndex_(headers) {
  const idx = {};
  for (let i = 0; i < headers.length; i++) {
    const h = String(headers[i] || "").trim();
    if (!h) continue;
    idx[h] = i;
  }
  return idx;
}

function isActiveProject_(row, statusIdx) {
  if (statusIdx < 0) return true; // if no status column, don't filter
  const status = String(row[statusIdx] || "").trim().toLowerCase();
  if (!status) return true;

  return CONFIG.INACTIVE_STATUSES.indexOf(status) === -1;
}

/**
 * Create the vertical CSV content ("Field,Value") from a row.
 * Uses Config to choose which headers to include and how to normalize.
 */
function generateVerticalCsvPayload_(row, headers, configMap) {
  const lines = [];

  for (let c = 0; c < headers.length; c++) {
    const sheetHeader = String(headers[c] || "").trim();
    if (!sheetHeader) continue;

    const cfg = configMap.get(sheetHeader);
    if (!cfg) continue;

    const field = cfg.csvLabel || sheetHeader;
    const val = formatValueForCsv_(row[c], cfg.type);

    lines.push(csvEscape_(field) + "," + csvEscape_(val));
  }

  return { csvContent: lines.join("\n") };
}

function formatValueForCsv_(value, type) {
  if (value === null || value === undefined) return "";
  if (value instanceof Date) {
    // normalize date output for stable hashing + predictable CSV
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }

  const t = String(type || "String").toLowerCase();
  if (t === "date") return String(value).trim();
  if (t === "number") return String(value).trim();
  if (t === "url") return String(value).trim();

  return String(value).trim();
}

function csvEscape_(s) {
  let v = String(s === null || s === undefined ? "" : s);
  v = v.replace(/"/g, '""');
  if (/[",\n\r]/.test(v)) v = `"${v}"`;
  return v;
}

function md5Hex_(content) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, content, Utilities.Charset.UTF_8);
  return bytes.map(b => {
    const v = (b < 0) ? b + 256 : b;
    return v.toString(16).padStart(2, "0");
  }).join("");
}

/**
 * Update by fileId if possible; otherwise locate folder and create/update the CSV.
 */
function updateOrCreateFile_(projectId, folderUrl, content, existingFileId) {
  // fast path: update by file id
  if (existingFileId) {
    try {
      const f = DriveApp.getFileById(existingFileId);
      f.setContent(content);
      return f;
    } catch (e) {
      Logger.log(`File ID invalid for ${projectId}: ${existingFileId}. Falling back to folder search.`);
    }
  }

  const folderId = extractFolderId_(folderUrl);
  const folder = DriveApp.getFolderById(folderId);

  let sub = null;
  const it = folder.getFoldersByName(CONFIG.SUBFOLDER_NAME);
  if (it.hasNext()) sub = it.next();
  else sub = folder.createFolder(CONFIG.SUBFOLDER_NAME);

  const files = sub.getFilesByName(CONFIG.FILENAME);
  if (files.hasNext()) {
    const f = files.next();
    f.setContent(content);
    return f;
  }

  return sub.createFile(CONFIG.FILENAME, content, MimeType.CSV);
}

function findCsvInProjectFolder_(folderUrl) {
  const folderId = extractFolderId_(folderUrl);
  const folder = DriveApp.getFolderById(folderId);

  // Try in subfolder first
  const subs = folder.getFoldersByName(CONFIG.SUBFOLDER_NAME);
  if (subs.hasNext()) {
    const sub = subs.next();
    const fIt = sub.getFilesByName(CONFIG.FILENAME);
    if (fIt.hasNext()) return fIt.next();
  }

  // Then root
  const rootFiles = folder.getFilesByName(CONFIG.FILENAME);
  if (rootFiles.hasNext()) return rootFiles.next();

  return null;
}

function extractFolderId_(url) {
  // Handles: https://drive.google.com/drive/folders/<ID>...
  // and: https://drive.google.com/open?id=<ID>
  const s = String(url || "");
  const m1 = s.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (m1 && m1[1]) return m1[1];

  const m2 = s.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (m2 && m2[1]) return m2[1];

  // If it's already an ID
  if (/^[a-zA-Z0-9_-]{10,}$/.test(s)) return s;

  throw new Error(`Unable to extract folder ID from URL: ${url}`);
}

/**
 * Parses a vertical CSV ("Field,Value") into an object.
 * Uses Utilities.parseCsv to properly handle quoted values/newlines.
 */
function parseVerticalCsvToMap_(csvText) {
  const rows = Utilities.parseCsv(csvText || "");
  const out = {};
  rows.forEach(r => {
    if (!r || r.length < 2) return;
    const key = String(r[0] || "").trim();
    if (!key) return;
    out[key] = String(r[1] || "");
  });
  return out;
}

function normalizeForCompare_(v) {
  if (v === null || v === undefined) return "";
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  return String(v).trim();
}

/**
 * Rich HTML dialog
 */
function showCombinedReport(pullResults, pushResults) {
  pullResults = pullResults || [];
  pushResults = pushResults || [];

  const touchedPull = pullResults.filter(r => r.status !== "Skipped").length;
  const touchedPush = pushResults.filter(r => r.status !== "Skipped").length;

  let html = ''
    + '<style>'
    + 'body{font-family:sans-serif;padding:10px;}'
    + 'table{border-collapse:collapse;width:100%;}'
    + 'th,td{border:1px solid #ddd;padding:6px;text-align:left;font-size:12px;vertical-align:top;}'
    + 'th{background:#f4f4f4;}'
    + '.Updated{background:#e6fffa;}'
    + '.Error{background:#ffe6e6;}'
    + '.Skipped{color:#666;}'
    + '</style>';

  html += `<h3>Smart Sync Report</h3>`;
  html += `<p><strong>Pulled:</strong> ${touchedPull} | <strong>Pushed:</strong> ${touchedPush}</p>`;
  html += '<table><tr><th>Project</th><th>Type</th><th>Status</th><th>Details</th></tr>';

  const all = pullResults.concat(pushResults);
  if (all.length === 0) {
    html += '<tr><td colspan="4">No actions taken.</td></tr>';
  } else {
    all.forEach(r => {
      const cls = r.status || "";
      html += `<tr class="${cls}"><td>${escHtml_(r.id || "")}</td><td>${escHtml_(r.type || "")}</td><td>${escHtml_(r.status || "")}</td><td>${escHtml_(r.details || "")}</td></tr>`;
    });
  }

  html += '</table><br><button onclick="google.script.host.close()" style="padding:8px 16px;cursor:pointer;">Close</button>';

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(750).setHeight(520),
    "GSADUs Sync Results"
  );
}

function escHtml_(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
