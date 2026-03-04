/**
 * GSADU Product Management System - v6 (Dynamic Header Mapping)
 * Supports: Tri-Partite IDs, Gap-Filling, Category Descriptions, and Dynamic Column Lookups
 */

/**
 * Diagnostic: Lists available Gemini models for this API key
 */
function checkAvailableModels() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    SpreadsheetApp.getUi().alert("No API Key found.");
    return;
  }
  
  const url = `https://generativelanguage.googleapis.com/v1beta/models?key=${apiKey}`;
  
  try {
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const json = JSON.parse(response.getContentText());
    
    if (json.error) {
      SpreadsheetApp.getUi().alert("API Error: " + json.error.message);
      return;
    }
    
    const models = json.models
      .filter(m => m.supportedGenerationMethods.includes("generateContent"))
      .map(m => m.name.replace('models/', ''))
      .sort();
      
    Logger.log("Available Models:\n" + models.join('\n'));
    SpreadsheetApp.getUi().alert("Available Models (check logs for full list):\n\n" + models.slice(0, 15).join('\n'));
  } catch (e) {
    SpreadsheetApp.getUi().alert("Fetch Error: " + e.toString());
  }
}

/**
 * Helper: Maps a spreadsheet to an array of objects keyed by header names.
 * This prevents column-mismatch errors if columns are moved or added.
 * Converts Date objects to strings for JSON serialization.
 */
function getMappedData(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);
  
  return rows.map(row => {
    let obj = {};
    headers.forEach((header, i) => {
      if (header) {
        // Convert Date objects to ISO strings for JSON serialization
        let value = row[i];
        if (value instanceof Date) {
          value = value.toISOString();
        }
        obj[header] = value;
      }
    });
    return obj;
  });
}

/**
 * Helper: Find column index by header name (1-based for getRange)
 */
function getColByName(sheet, name) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const index = headers.indexOf(name);
  return index === -1 ? null : index + 1;
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('GSADU Tools')
    .addItem('Open Product Entry', 'showSidebar')
    .addItem('Open Category Manager', 'showCategoryManager')
    .addItem('Bulk AI Importer', 'showBulkImporter')
    .addSeparator()
    .addItem('Run Data Audit', 'runAudit')
    .addSeparator()
    .addItem('Sort Products by Item ID', 'sortMasterDB')
    .addItem('Export Data as JSON', 'showDataExport')
    .addSeparator()
    .addItem('Refresh System', 'showSidebar') 
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('GSADU Product Manager')
    .setWidth(450);
  SpreadsheetApp.getUi().showSidebar(html);
}

function showCategoryManager() {
  const html = HtmlService.createHtmlOutputFromFile('CategoryManager')
    .setTitle('GSADU Category Manager')
    .setWidth(500);
  SpreadsheetApp.getUi().showSidebar(html);
}

function showDataExport() {
  const html = HtmlService.createHtmlOutputFromFile('DataExport')
    .setTitle('GSADU Data Export')
    .setWidth(600);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Exports all sheet data as a JSON object for debugging/reference
 * Uses dynamic header mapping - column positions don't matter
 */
function exportAllData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = getMappedData('SYSTEM_CONFIG');
    const master = getMappedData('MASTER_DB');
    const changelog = getMappedData('CHANGE_LOG');
    
    // Get sidebar view for debugging
    let sidebarView = null;
    try {
      sidebarView = getSelectionData();
    } catch (e) {
      sidebarView = { error: e.toString() };
    }
    
    return {
      exportDate: new Date().toISOString(),
      spreadsheetId: ss.getId(),
      spreadsheetName: ss.getName(),
      SYSTEM_CONFIG: config,
      MASTER_DB: master,
      CHANGE_LOG: changelog,
      sidebarView: sidebarView
    };
  } catch (e) {
    return { 
      error: e.toString(),
      stack: e.stack 
    };
  }
}

/**
 * Gets selection data for Sidebar dropdowns
 * Uses dynamic header mapping - resilient to column changes
 */
function getSelectionData() {
  const configData = getMappedData('SYSTEM_CONFIG');
  
  // Extract simple lists by header name, remove duplicates
  const divisions = [...new Set(configData.map(r => r['DIV_LABEL']).filter(Boolean))];
  const tiers = [...new Set(configData.map(r => r['TIERS']).filter(Boolean))];
  const bundles = [...new Set(configData.map(r => r['BUNDLE_NAME']).filter(Boolean))];
  const finishes = [...new Set(configData.map(r => r['FINISH_COLORS']).filter(Boolean))];

  // Format Category Data for Sidebar's array expectation:
  // [Parent, ID, Name, Desc, Gap, ShowFields]
  const categoryData = configData
    .filter(r => r['PARENT_DIV']) 
    .map(r => [
      r['PARENT_DIV'],    // [0] Parent Division
      r['CAT_ID'],        // [1] Category ID
      r['CAT_NAME'],      // [2] Category Name
      r['CAT_DESC'],      // [3] Description
      "",                 // [4] Gap (maintains Sidebar.html index 5)
      r['SHOW_FIELDS'] || ""  // [5] ShowFields
    ]);

  return {
    divisions: divisions,
    categoryData: categoryData,
    tiers: tiers,
    bundles: bundles,
    finishes: finishes
  };
}

function generateNextItemID(catID) {
  if (!catID) return "ERROR-NO-CATEGORY";
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName('MASTER_DB');
  
  // Get Item_ID column dynamically
  const itemIdCol = getColByName(master, 'Item_ID');
  if (!itemIdCol) return catID + "-01";
  
  const lastRow = master.getLastRow();
  if (lastRow < 2) return catID + "-01";

  const ids = master.getRange(2, itemIdCol, lastRow - 1, 1).getValues().flat();
  let existingNums = ids
    .filter(id => id.toString().startsWith(catID + "-"))
    .map(id => {
      let parts = id.toString().split('-');
      return parts.length === 3 ? parseInt(parts[2]) : null;
    })
    .filter(num => num !== null && !isNaN(num))
    .sort((a, b) => a - b);

  let nextNum = 1;
  for (let i = 0; i < existingNums.length; i++) {
    if (existingNums[i] === nextNum) { nextNum++; } 
    else if (existingNums[i] > nextNum) { break; }
  }
  return catID + "-" + nextNum.toString().padStart(2, '0');
}

function addCategory(divLabel, name, desc, showFields = "") {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = ss.getSheetByName('SYSTEM_CONFIG');
  const divId = divLabel.split(' : ')[0];
  
  // Get column indices dynamically
  const catIdCol = getColByName(config, 'CAT_ID');
  const parentDivCol = getColByName(config, 'PARENT_DIV');
  const catNameCol = getColByName(config, 'CAT_NAME');
  const catDescCol = getColByName(config, 'CAT_DESC');
  const showFieldsCol = getColByName(config, 'SHOW_FIELDS');
  
  if (!catIdCol || !parentDivCol) {
    return "Error: Required columns CAT_ID or PARENT_DIV not found";
  }
  
  // Find existing category IDs for this division
  const lastRow = config.getLastRow();
  const existingIds = config.getRange(2, catIdCol, lastRow - 1, 1).getValues().flat()
    .filter(id => id.toString().startsWith(divId + "-"))
    .map(id => parseInt(id.toString().split('-')[1]))
    .sort((a,b) => a-b);

  // Gap-filling for category number
  let nextNum = 1;
  for (let n of existingIds) {
    if (n === nextNum) nextNum++;
    else if (n > nextNum) break;
  }
  const newCatId = divId + "-" + nextNum.toString().padStart(2, '0');
  
  // Find first empty row in PARENT_DIV column
  const parentData = config.getRange(2, parentDivCol, lastRow - 1, 1).getValues();
  let targetRow = lastRow + 1;
  for (let i = 0; i < parentData.length; i++) {
    if (parentData[i][0] === "") {
      targetRow = i + 2;
      break;
    }
  }
  
  // Write values to their respective columns
  config.getRange(targetRow, parentDivCol).setValue(divLabel);
  config.getRange(targetRow, catIdCol).setValue(newCatId);
  if (catNameCol) config.getRange(targetRow, catNameCol).setValue(name);
  if (catDescCol) config.getRange(targetRow, catDescCol).setValue(desc);
  if (showFieldsCol) config.getRange(targetRow, showFieldsCol).setValue(showFields);
  
  return "Successfully created " + newCatId + ": " + name;
}

function saveToDatabase(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName('MASTER_DB');
  const log = ss.getSheetByName('CHANGE_LOG');
  
  // Validate required fields
  if (!payload.itemId) {
    throw new Error("Item ID is required. Please select a category first.");
  }
  
  // Get column indices dynamically
  const cols = {
    Item_ID: getColByName(master, 'Item_ID'),
    Division: getColByName(master, 'Division'),
    Category: getColByName(master, 'Category'),
    Product_Name: getColByName(master, 'Product_Name'),
    Description: getColByName(master, 'Description'),
    Tier: getColByName(master, 'Tier'),
    Bundle: getColByName(master, 'Bundle'),
    Finish: getColByName(master, 'Finish'),
    Image_URL: getColByName(master, 'Image_URL'),
    Product_URL: getColByName(master, 'Product_URL'),
    Last_Updated: getColByName(master, 'Last_Updated')
  };
  
  // Find actual last row with data in Item_ID column (ignores data validation)
  let targetRow = 2; // Start after header
  if (cols.Item_ID) {
    const itemIdData = master.getRange(2, cols.Item_ID, master.getMaxRows() - 1, 1).getValues();
    for (let i = 0; i < itemIdData.length; i++) {
      if (itemIdData[i][0] === "" || itemIdData[i][0] === null) {
        targetRow = i + 2;
        break;
      }
      targetRow = i + 3; // Move past this row
    }
  }
  
  // Write each value to its column
  if (cols.Item_ID) master.getRange(targetRow, cols.Item_ID).setValue("'" + payload.itemId);
  if (cols.Division) master.getRange(targetRow, cols.Division).setValue(payload.division);
  if (cols.Category) master.getRange(targetRow, cols.Category).setValue(payload.category);
  if (cols.Product_Name) master.getRange(targetRow, cols.Product_Name).setValue(payload.productName);
  if (cols.Description) master.getRange(targetRow, cols.Description).setValue(payload.description || "");
  if (cols.Tier) master.getRange(targetRow, cols.Tier).setValue(payload.tier);
  if (cols.Bundle) master.getRange(targetRow, cols.Bundle).setValue(payload.bundle || "");
  if (cols.Finish) master.getRange(targetRow, cols.Finish).setValue(payload.finish || "");
  if (cols.Image_URL) master.getRange(targetRow, cols.Image_URL).setValue(payload.imageUrl || "");
  if (cols.Product_URL) master.getRange(targetRow, cols.Product_URL).setValue(payload.productUrl || "");
  if (cols.Last_Updated) master.getRange(targetRow, cols.Last_Updated).setValue(new Date());
  
  // Auto-sort MASTER_DB by Item_ID after insert
  sortMasterDB();
  
  // Log the action
  log.appendRow([new Date(), Session.getActiveUser().getEmail(), "ADD", payload.itemId, payload.productName]);
  return "Successfully added " + payload.itemId;
}

/**
 * Sorts MASTER_DB by Item_ID column (ascending)
 * Safe to call anytime - preserves header row
 */
function sortMasterDB() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName('MASTER_DB');
  const itemIdCol = getColByName(master, 'Item_ID');
  
  if (!itemIdCol) return;
  
  const lastRow = master.getLastRow();
  const lastCol = master.getLastColumn();
  
  if (lastRow < 3) return; // Need at least 2 data rows to sort
  
  // Sort data range (excluding header row)
  const dataRange = master.getRange(2, 1, lastRow - 1, lastCol);
  dataRange.sort({ column: itemIdCol, ascending: true });
}

function mergeCategories(keepId, mergeId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName('MASTER_DB');
  const config = ss.getSheetByName('SYSTEM_CONFIG');
  
  // Get column indices dynamically from config
  const catIdCol = getColByName(config, 'CAT_ID');
  const catNameCol = getColByName(config, 'CAT_NAME');
  
  // Get keep category name from config
  const configData = getMappedData('SYSTEM_CONFIG');
  const keepRow = configData.find(r => r['CAT_ID'] === keepId);
  const keepName = keepRow ? keepRow['CAT_NAME'] : "";
  
  // Get MASTER_DB column indices
  const itemIdCol = getColByName(master, 'Item_ID');
  const categoryCol = getColByName(master, 'Category');
  
  if (!itemIdCol || !categoryCol) {
    return "Error: Required columns not found in MASTER_DB";
  }
  
  const lastRow = master.getLastRow();
  if (lastRow > 1) {
    const itemIds = master.getRange(2, itemIdCol, lastRow - 1, 1).getValues();
    for (let i = 0; i < itemIds.length; i++) {
      if (itemIds[i][0].toString().startsWith(mergeId + "-")) {
        let itemPart = itemIds[i][0].toString().split('-')[2];
        let newId = keepId + "-" + itemPart;
        master.getRange(i + 2, itemIdCol).setValue("'" + newId);
        master.getRange(i + 2, categoryCol).setValue(keepName);
      }
    }
  }
  deleteCategoryFromConfig(mergeId);
  return "Merged items into " + keepId;
}

function deleteCategoryFromConfig(catId) {
  const config = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SYSTEM_CONFIG');
  
  // Get all category-related column indices
  const catIdCol = getColByName(config, 'CAT_ID');
  const parentDivCol = getColByName(config, 'PARENT_DIV');
  const catNameCol = getColByName(config, 'CAT_NAME');
  const catDescCol = getColByName(config, 'CAT_DESC');
  const showFieldsCol = getColByName(config, 'SHOW_FIELDS');
  
  if (!catIdCol) {
    return "Error: CAT_ID column not found";
  }
  
  const lastRow = config.getLastRow();
  const data = config.getRange(2, catIdCol, lastRow - 1, 1).getValues();
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === catId) {
      const row = i + 2;
      // Clear each category-related column
      if (parentDivCol) config.getRange(row, parentDivCol).clearContent();
      config.getRange(row, catIdCol).clearContent();
      if (catNameCol) config.getRange(row, catNameCol).clearContent();
      if (catDescCol) config.getRange(row, catDescCol).clearContent();
      if (showFieldsCol) config.getRange(row, showFieldsCol).clearContent();
      break;
    }
  }
  return "Removed " + catId;
}

// =============================================================================
// AI BULK IMPORTER FUNCTIONS
// =============================================================================

/**
 * Shows the Bulk AI Importer sidebar
 */
function showBulkImporter() {
  const html = HtmlService.createHtmlOutputFromFile('BulkImporter')
    .setTitle('GSADU AI Bulk Importer')
    .setWidth(600);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Core AI Function: Sends product info to Gemini Flash latest
 * Uses your categories from SYSTEM_CONFIG for context
 */
/**
 * CLEAN VERSION: Extracts Open Graph tags ONLY.
 * No regex guessing, no URL reading.
 */
function fetchLinkPreview(url) {
  // 1. Try "Free" Scraping first (OG Tags) with robust fallback
  const scrapeResult = tryScraping(url);
  
  // If we got a valid title and image, use it. 
  // If we got a title but no image, we might still want to try search if specifically requested, 
  // but for now let's say if we have a title, we are good enough to avoid "Nuclear Option".
  // However, the user specifically wants the IMAGE.
  if (scrapeResult.success && scrapeResult.image && scrapeResult.title) {
    return scrapeResult;
  }
  
  // 2. If scraping failed or missed the image, use Google Custom Search API (The "Nuclear Option")
  // You need a Custom Search Engine (CSE) configured to search "The entire web" or specific shopping sites.
  // Set the property in File > Project Settings > Script Properties
  const scriptProps = PropertiesService.getScriptProperties();
  const SEARCH_ENGINE_ID = scriptProps.getProperty('GOOGLE_SEARCH_ENGINE_ID') || 'YOUR_SEARCH_ENGINE_ID_HERE'; 
  const API_KEY = scriptProps.getProperty('GOOGLE_SEARCH_API_KEY') || scriptProps.getProperty('GEMINI_API_KEY');

  // Logic: Only try this if we have a Search Engine ID, otherwise skip to save quota/errors
  if (SEARCH_ENGINE_ID && SEARCH_ENGINE_ID !== 'YOUR_SEARCH_ENGINE_ID_HERE') {
    try {
      // We search for the URL itself to find Google's indexed entry for it, requesting image result
      const apiUrl = `https://www.googleapis.com/customsearch/v1?q=${encodeURIComponent(url)}&cx=${SEARCH_ENGINE_ID}&key=${API_KEY}&searchType=image&num=1`;
      
      const response = UrlFetchApp.fetch(apiUrl, { muteHttpExceptions: true });
      const json = JSON.parse(response.getContentText());
      
      if (json.items && json.items.length > 0) {
        const googleImage = json.items[0].link;
        // If scraping failed completely, use Google Title. If scraping worked but lacked image, keep scraped title.
        const title = (scrapeResult.success && scrapeResult.title) ? scrapeResult.title : (json.items[0].title || "Product Found via Google");
        
        return {
          success: true,
          title: decodeHtml(title),
          description: "Image retrieved via Google Search",
          image: googleImage,
          url: url
        };
      }
    } catch (e) {
      Logger.log("Google Search fallback failed: " + e.toString());
    }
  }

  // 3. Fallback / Total Failure
  // If we at least scraped a title, return that without image
  if (scrapeResult.success) {
    return scrapeResult;
  }

  return { 
    success: false, // This will trigger the "URL Analysis Only" fallback in aiProcessLink
    title: "Manual Entry Required", 
    image: "", 
    url: url,
    description: "Could not scrape or search for this product."
  };
}

/**
 * Helper: The standard OG Tag scraper
 */
function tryScraping(url) {
  try {
    const params = {
      'muteHttpExceptions': true,
      'followRedirects': true,
      'headers': {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8'
      }
    };
    
    const response = UrlFetchApp.fetch(url, params);
    const contentText = response.getContentText();
    const responseCode = response.getResponseCode();

    if (responseCode !== 200 && responseCode !== 206) {
       return { success: false, error: "Status " + responseCode };
    }

    // Truncate to first 100k chars for regex performance
    const html = contentText.length > 100000 ? contentText.substring(0, 100000) : contentText;

    const ogImage = html.match(/<meta property="og:image" content="(.*?)"/i)?.[1];
    const ogTitle = html.match(/<meta property="og:title" content="(.*?)"/i)?.[1];
    const ogDesc  = html.match(/<meta property="og:description" content="(.*?)"/i)?.[1];
    
    // Fallback standard tags
    const title = ogTitle || html.match(/<title>(.*?)<\/title>/i)?.[1];
    const desc = ogDesc || html.match(/<meta name="description" content="(.*?)"/i)?.[1];

    if (!title && !ogImage) return { success: false, error: "No tags found" };

    return { 
      success: true, 
      title: decodeHtml(title || ""), 
      image: ogImage || "", 
      description: decodeHtml(desc || ""), 
      url: url 
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * AI Function: Reasons based purely on the scrapped context
 */
function callGeminiAI(productTitle, description, imageUrl, productUrl) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  const selectionData = getSelectionData();
  const categoryMenu = selectionData.categoryData.map(c => `- ${c[0]} | ${c[2]}`).join('\n');

  // We give the AI the specific context it needs
  const prompt = `
  Analyze this product information extracted from a retailer's website.
  
  PRODUCT CONTEXT:
  Title: "${productTitle}"
  Description: "${description}"
  URL: "${productUrl || ''}"
  
  TASK:
  1. Identify the core product. (HINT: If Title is "Access Denied" or generic, extract the product name from the URL slug).
  2. Assign it to the correct Division and Category from the list below.
  3. Create a clean, professional name (Brand + Product Type + Key Spec). MUST BE 50 CHARACTERS OR LESS.
  4. Use the "short_desc" field to include key details (dimensions, material, technical specs) that did not fit in the name.
  5. Assign a "confidence" score (0-100) reflecting how certain you are about the Division/Category match.
  
  VALID CATEGORIES:
  ${categoryMenu}
  
  Return JSON ONLY: { "division": "...", "category": "...", "cleaned_name": "...", "short_desc": "...", "confidence": 85 }
  `;

  // Use verified v1beta endpoint with gemini-flash-latest to ensure compatibility
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key=${apiKey}`;
  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { responseMimeType: "application/json" }
  };

  try {
    const response = UrlFetchApp.fetch(url, {
      method: "POST",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    
    const json = JSON.parse(response.getContentText());
    if (json.error) throw new Error(json.error.message);
    
    let aiText = json.candidates[0].content.parts[0].text;
    aiText = aiText.replace(/```json/g, "").replace(/```/g, ""); // Clean markdown
    return JSON.parse(aiText);
    
  } catch (e) {
    // If AI fails, return a manual review object
    return {
      division: "01 : Pre-Construction Work",
      category: "Review Required",
      cleaned_name: productTitle,
      short_desc: "AI Error: " + e.message
    };
  }
}

// Helper to clean up HTML entities (e.g., &amp; -> &)
function decodeHtml(html) {
  if (!html) return "";
  return html.replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, "\"").replace(/&#39;/g, "'");
}

/**
 * Bridge function: Processes a single URL through scraping + AI
 * Called from BulkImporter.html
 */
function aiProcessLink(url) {
  // 1. Get the link metadata via Open Graph scraping
  let metadata = fetchLinkPreview(url);
  
  if (!metadata.success) {
     // Fallback: If scraping errors (DNS/Timeouts), still let AI try the URL
     console.warn("Link scrape failed (" + metadata.error + "), falling back to AI URL analysis.");
     metadata = {
       title: "URL Analysis Only",
       description: "Site could not be reached. Analyzing URL text.",
       image: ""
     };
  }
  
  // 2. Use Gemini to categorize it
  const aiResult = callGeminiAI(metadata.title, metadata.description, metadata.image, url);
  
  // 3. Add the image URL from scraping if available
  aiResult.image_url = metadata.image || "";
  aiResult.source_url = url;
  
  return aiResult;
}

/**
 * Saves a product from AI import (generates Item_ID automatically)
 * Called from BulkImporter.html
 */
function saveAIProduct(data) {
  // Find the matching category ID from the category name
  const selectionData = getSelectionData();
  const categoryMatch = selectionData.categoryData.find(c => 
    c[0] === data.division && c[2] === data.category
  );
  
  if (!categoryMatch) {
    throw new Error(`Category not found: ${data.division} > ${data.category}`);
  }
  
  const catId = categoryMatch[1]; // The CAT_ID like "07-01"
  const itemId = generateNextItemID(catId);
  
  const payload = {
    itemId: itemId,
    division: data.division,
    category: data.category,
    productName: data.cleaned_name,
    description: data.short_desc,
    tier: data.tier || "Standard",
    imageUrl: data.image_url || "",
    productUrl: data.source_url
  };
  
  return saveToDatabase(payload);
}

/**
 * Scrapes the Open Graph image from a URL.
 * @param {string} url The product URL to scrape.
 * @return The image URL string.
 * @customfunction
 */
function GET_OG_IMAGE(url) {
  if (!url) return "";
  try {
    const params = {
      'muteHttpExceptions': true,
      'headers': {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
      }
    };
    
    const response = UrlFetchApp.fetch(url, params);
    const contentText = response.getContentText();
    // Truncate to first 100k chars to avoid regex freezing on large pages
    const html = contentText.length > 100000 ? contentText.substring(0, 100000) : contentText;

    // Look for og:image meta tag
    const match = html.match(/<meta[^>]*property=["']og:image["'][^>]*content=["']([^"']+)["']/i) 
               || html.match(/<meta[^>]*content=["']([^"']+)["'][^>]*property=["']og:image["']/i);
    return match ? match[1] : "No Image Found";
  } catch (e) {
    return "Error: " + e.message;
  }
}