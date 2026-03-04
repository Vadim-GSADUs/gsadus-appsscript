/**
 * AUDIT SYSTEM
 * Scans MASTER_DB for quality issues (Duplicates, Empty Fields, AI Logic)
 */

function runAudit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName('MASTER_DB');
  if (!master) {
    SpreadsheetApp.getUi().alert("MASTER_DB sheet not found.");
    return;
  }
  
  const data = master.getDataRange().getValues();
  const headers = data[0];
  
  // Mappings
  const itemIdIdx = headers.indexOf('Item_ID');
  const urlIdx = headers.indexOf('Product_URL');
  const nameIdx = headers.indexOf('Product_Name');
  const divIdx = headers.indexOf('Division');
  const catIdx = headers.indexOf('Category');
  const tierIdx = headers.indexOf('Tier'); // Added Tier
  const descIdx = headers.indexOf('Description');
  
  if (urlIdx === -1 || nameIdx === -1) {
    SpreadsheetApp.getUi().alert("Could not find required columns (Product_URL or Product_Name).");
    return;
  }

  // Load Config for Validation
  const config = getSelectionData();
  const validTiers = config.tiers; 
  const categoryMap = {}; 
  
  config.categoryData.forEach(cat => {
    const key = `${cat[0]}|${cat[2]}`; 
    categoryMap[key] = cat[1];
  });

  let report = [];
  let urlTracker = {};
  let idTracker = {}; // Check for duplicate IDs
  let previousId = ""; // Check for sort order

  // 1. HARD CHECKS (Fast loops)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowNum = i + 1;

    // SKIP if no Item_ID
    if (itemIdIdx !== -1 && !row[itemIdIdx]) continue;

    const itemId = row[itemIdIdx] ? String(row[itemIdIdx]).trim() : "";
    const name = row[nameIdx];
    const division = row[divIdx];
    const category = row[catIdx];
    const tier = tierIdx !== -1 ? row[tierIdx] : null;
    const url = row[urlIdx];
    
    // Check A: Item_ID Integrity
    const idRegex = /^\d{2}-\d{2}-\d{2}$/;
    
    // 1. Duplicate ID Check (CRITICAL)
    if (idTracker[itemId]) {
      report.push({ row: rowNum, issue: "Duplicate Item_ID", details: `ID ${itemId} already exists at Row ${idTracker[itemId]}` });
    } else {
      idTracker[itemId] = rowNum;
    }

    // 2. Sort Order Check (Items should be sequential)
    if (previousId && itemId < previousId) {
       report.push({ row: rowNum, issue: "Unsorted Row", details: `${itemId} appears after ${previousId}` });
    }
    previousId = itemId;

    // 3. Structure Format
    if (!idRegex.test(itemId)) {
      report.push({ row: rowNum, issue: "Invalid Item_ID", details: `Format must be XX-XX-XX. Found: ${itemId}` });
    } else {
      // ... logic checks for div/cat match ...
      const expectedDivPrefix = division ? division.split(' : ')[0] : "";
      if (expectedDivPrefix && !itemId.startsWith(expectedDivPrefix)) {
        report.push({ row: rowNum, issue: "ID Mismatch (Div)", details: `ID ${itemId} does not match Division ${expectedDivPrefix}` });
      }
      
      if (division && category) {
        const expectedCatId = categoryMap[`${division}|${category}`];
        if (expectedCatId && !itemId.startsWith(expectedCatId + "-")) {
             report.push({ row: rowNum, issue: "ID Mismatch (Cat)", details: `Category maps to ${expectedCatId}, but ID is ${itemId}` });
        }
      }
    }
    
    // Check B: Valid Tier
    if (tier && !validTiers.includes(tier)) {
       report.push({ row: rowNum, issue: "Invalid Tier", details: `Value '${tier}' not in System Config` });
    }

    // Check C: Duplicate URLs
    if (url && typeof url === 'string') {
      const cleanUrl = url.trim();
      if (cleanUrl) {
        if (urlTracker[cleanUrl]) {
          report.push({ row: rowNum, issue: "Duplicate URL", details: `Same as Row ${urlTracker[cleanUrl]}` });
        } else {
          urlTracker[cleanUrl] = rowNum;
        }
      }
    }

    // Check B: Missing Data
    if (!row[nameIdx]) report.push({ row: rowNum, issue: "Missing Name", details: "Product Name is empty" });
    if (divIdx !== -1 && !row[divIdx]) report.push({ row: rowNum, issue: "Missing Division", details: "Division not assigned" });

    // Check C: Length Limits
    if (row[nameIdx] && row[nameIdx].length > 100) {
      report.push({ row: rowNum, issue: "Name Too Long", details: `Current: ${row[nameIdx].length} chars (Max 100)` });
    }
  }

  // 2. AI CHECKS (Batch process for efficiency)
  // We pick a few random rows or "suspect" rows to spot check with Gemini
  // For this example, let's check the last 5 added items
  const lastRow = data.length;
  // If data.length is small (e.g. only headers), don't run AI check
  if (lastRow > 1) {
    const startCheck = Math.max(1, lastRow - 5);
    
    // Show a toast that AI audit is starting
    SpreadsheetApp.getActiveSpreadsheet().toast("Auditing last 5 entries with AI...", "Audit System");
    
    for (let i = startCheck; i < lastRow; i++) {
      const row = data[i];
      
      // SKIP if no Item_ID
      if (itemIdIdx !== -1 && !row[itemIdIdx]) continue;

      // Only audit if we have minimum data
      if (row[nameIdx]) {
        const validation = aiAuditRow(row[nameIdx], row[divIdx], row[catIdx], row[descIdx]);
        
        if (!validation.passed) {
          report.push({ 
            row: i + 1, 
            issue: "AI Flagged", 
            details: validation.reason 
          });
        }
      }
    }
  }

  // 3. SHOW REPORT
  showAuditReport(report);
}

/**
 * AI Auditor: Asks Gemini if the classification makes sense
 */
function aiAuditRow(name, division, category, desc) {
  if (!name) return { passed: true }; // Skip empty rows

  const prompt = `
  You are a Data Auditor. Check this product entry for logical errors.
  
  PRODUCT: ${name}
  ASSIGNED TO: ${division} > ${category}
  DESC: ${desc}
  
  RULES:
  1. Does the product actually belong in this Division? (e.g., A "Toilet" should NOT be in "Electrical")
  2. Is the description professional? (No "ALL CAPS", no typos)
  
  RESPONSE:
  Return JSON ONLY: { "passed": true/false, "reason": "Short explanation if false (max 10 words)" }
  `;
  
  try {
    const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) return { passed: true }; // Skip if no key

    // Use verified v1beta endpoint with gemini-flash-latest
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key=${apiKey}`;
    
    const payload = { 
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: { 
        temperature: 0.1
      }
    };
    
    const response = UrlFetchApp.fetch(url, {
      method: "POST", 
      contentType: "application/json", 
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    
    const responseCode = response.getResponseCode();
    if (responseCode !== 200) return { passed: true }; // Fail open

    const json = JSON.parse(response.getContentText());
    if (!json.candidates || !json.candidates[0]) return { passed: true };

    let text = json.candidates[0].content.parts[0].text;
    text = text.replace(/^```json/i, "").replace(/```$/i, "").trim();
    return JSON.parse(text);

  } catch (e) {
    Logger.log("Audit Error: " + e.toString());
    return { passed: true }; // Fail open if AI breaks
  }
}

/**
 * Display Report in a simple Modal
 */
function showAuditReport(issues) {
  let html = `
    <html>
      <head>
        <style>
          body { font-family: sans-serif; font-size: 13px; padding: 10px; }
          h3 { margin-top: 0; color: #1b5e20; }
          table { width: 100%; border-collapse: collapse; margin-top: 10px; }
          th { text-align: left; background: #e0e0e0; padding: 6px; border: 1px solid #ccc; font-size: 11px; }
          td { border: 1px solid #c0c0c0; padding: 6px; vertical-align: top; }
          .red { color: #d32f2f; font-weight: bold; }
          .row-id { font-family: monospace; text-align: center; color: #666; }
        </style>
      </head>
      <body>
        <h3>Audit Report</h3>
  `;
  
  if (issues.length === 0) {
    html += "<p style='color:green; font-weight:bold; font-size:14px;'>✅ No issues found!</p><p>Great job keeping the database clean.</p>";
  } else {
    html += `<p>Found <strong>${issues.length}</strong> potential issues:</p>`;
    html += "<table><tr><th width='40'>Row</th><th width='100'>Issue</th><th>Details</th></tr>";
    issues.forEach(item => {
      html += `<tr><td class="row-id">${item.row}</td><td class="red">${item.issue}</td><td>${item.details}</td></tr>`;
    });
    html += "</table>";
  }
  
  html += `
        <div style="margin-top:15px; text-align:right;">
          <button onclick="google.script.host.close()" style="padding: 6px 12px; cursor: pointer;">Close</button>
        </div>
      </body>
    </html>
  `;
  
  const ui = HtmlService.createHtmlOutput(html).setWidth(600).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(ui, 'Audit Results');
}