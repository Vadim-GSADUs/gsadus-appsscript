# GSADU Product Management System - Refactoring Plan

## Project Overview

**Purpose:** Evolve a static PDF product catalog into a dynamic Google Sheets-based product management system for ADU (Accessory Dwelling Unit) construction materials.

**Google Sheet URL:** https://docs.google.com/spreadsheets/d/1weTU_afdOWqd8XBIup5FAND1JveAowQQEqGmrZBv0Ho/edit?gid=352330289#gid=352330289

**Current State:** Partially implemented Apps Script with Sidebar for product entry and Category Manager for administration. Two test entries exist in MASTER_DB.

---

## Architecture Decisions (Finalized)

### ID System
- **Format:** `DIV-CAT-ITEM` (e.g., `02-01-03`)
- **Logic:** Gap-filling algorithm assigns lowest available item number within a category
- **Rationale:** Supports reorganization, category merges, and surgical item insertion

### Data Model

#### MASTER_DB Schema (11 columns)
| Column | Type | Required | Description |
|--------|------|----------|-------------|
| Item_ID | String | Yes | Tri-partite ID (DIV-CAT-ITEM) |
| Division | String | Yes | Division label (e.g., "02 : Site Improvements") |
| Category | String | Yes | Category name (e.g., "Fencing & Gates") |
| Product_Name | String | Yes | Short product name |
| Description | String | No | Detailed specs, model info, notes |
| Tier | String | Yes | "Standard" (default), "Premium", or "Optional" |
| Bundle | String | No | B1-B6 or blank (only for applicable categories) |
| Finish | String | No | Finish color or blank (only for applicable categories) |
| Image_URL | String | No | Direct link to product image |
| Product_URL | String | No | Link to product/vendor page |
| Last_Updated | Date | Auto | Timestamp of last modification |

#### SYSTEM_CONFIG Schema
**Divisions (Columns A-C):**
| Column | Header | Description |
|--------|--------|-------------|
| A | DIV_ID | Two-digit division number (01-20) |
| B | DIV_NAME | Division name |
| C | DIV_LABEL | Formatted label "ID : Name" |

**Categories (Columns D-H + new I):**
| Column | Header | Description |
|--------|--------|-------------|
| D | PARENT_DIV | Parent division label |
| E | CAT_ID | Category ID (DIV-CAT format) |
| F | CAT_NAME | Category name |
| G | CAT_DESC | Category description |
| H | SHOW_FIELDS | NEW - Comma-delimited list: "Bundle", "Finish", or "Bundle,Finish" |

**Tiers (Column J):**
| Column | Header | Description |
|--------|--------|-------------|
| J | TIERS | Tier options: Standard, Premium, Optional |

**Bundles (Columns L-M):**
| Column | Header | Description |
|--------|--------|-------------|
| L | BUNDLE_ID | Bundle code (B1-B6) |
| M | BUNDLE_NAME | Bundle name |

**Finishes (Column O):**
| Column | Header | Description |
|--------|--------|-------------|
| O | FINISH_COLORS | Finish color options |

### Dynamic UI Behavior
- Sidebar reads `SHOW_FIELDS` column from SYSTEM_CONFIG
- Bundle dropdown: Only visible when category's SHOW_FIELDS contains "Bundle"
- Finish dropdown: Only visible when category's SHOW_FIELDS contains "Finish"
- Tier dropdown: Always visible, defaults to "Standard"

---

## Phase 1: Schema Refactoring

### 1.1 Update SYSTEM_CONFIG Sheet

**Add new column:**
- Insert column H: `SHOW_FIELDS`
- Populate based on category requirements:

| CAT_ID | CAT_NAME | SHOW_FIELDS |
|--------|----------|-------------|
| 16-01 | Flooring Materials | Bundle |
| 16-02 | Countertops | Bundle |
| 16-03 | Tile & Backsplash | Bundle,Finish |
| 18-01 | Cabinetry | Bundle |
| 07-03 | Fixtures & Hardware | Finish |
| (all others) | ... | (blank) |

**Categories requiring SHOW_FIELDS values:**
- DIV-16 (Floor Coverings & Tile): All categories need "Bundle", 16-03 also needs "Finish" for Schluter edges
- DIV-18 (Finish Carpentry): 18-01 Cabinetry needs "Bundle"
- DIV-07 (Plumbing): 07-03 Fixtures & Hardware needs "Finish" for fixture finishes
- DIV-06 (Electrical): Potentially 06-03 Lighting: Decorative needs "Finish"
- Review other divisions for finish requirements

### 1.2 Update MASTER_DB Sheet

**Replace existing headers with:**
```
Item_ID | Division | Category | Product_Name | Description | Tier | Bundle | Finish | Image_URL | Product_URL | Last_Updated
```

**Migration steps:**
1. Backup current MASTER_DB data
2. Clear existing headers
3. Add new headers (Row 1)
4. Migrate any existing test data to new structure
5. Format columns appropriately (Date for Last_Updated)

### 1.3 Update CHANGE_LOG Sheet (if needed)

Current structure appears adequate. Verify columns:
```
Timestamp | User | Action | Item_ID | Product_Name
```

---

## Phase 2: Code Refactoring

### 2.1 Code.js - getSelectionData()

**Current behavior:** Returns divisions, categoryData, tiers, bundles, finishes as flat arrays.

**New behavior:** 
- categoryData must include SHOW_FIELDS column
- Return structure:

```javascript
return {
  divisions: [...],           // Column C values
  categoryData: [...],        // Columns D-H (Parent, ID, Name, Desc, ShowFields)
  tiers: [...],               // Column J values
  bundles: [...],             // Column L values (BUNDLE_ID)
  finishes: [...]             // Column O values
};
```

**Code changes:**
```javascript
function getSelectionData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = ss.getSheetByName('SYSTEM_CONFIG');
  
  const divisions = config.getRange("C2:C21").getValues().flat().filter(String);
  // Updated range to include SHOW_FIELDS (column H)
  const categoryData = config.getRange("D2:H100").getValues().filter(row => row[0] !== "");
  const tiers = config.getRange("J2:J10").getValues().flat().filter(String);
  const bundles = config.getRange("L2:L10").getValues().flat().filter(String);
  const finishes = config.getRange("O2:O20").getValues().flat().filter(String);
  
  return {
    divisions: divisions,
    categoryData: categoryData,  // Now includes [Parent, ID, Name, Desc, ShowFields]
    tiers: tiers,
    bundles: bundles,
    finishes: finishes
  };
}
```

### 2.2 Code.js - saveToDatabase()

**Current behavior:** Saves with old schema including UID.

**New behavior:** Save with new 11-column schema.

```javascript
function saveToDatabase(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName('MASTER_DB');
  const log = ss.getSheetByName('CHANGE_LOG');
  
  const rowData = [
    "'" + payload.itemId,      // Item_ID (text format)
    payload.division,           // Division
    payload.category,           // Category
    payload.productName,        // Product_Name
    payload.description,        // Description
    payload.tier,               // Tier
    payload.bundle || "",       // Bundle (nullable)
    payload.finish || "",       // Finish (nullable)
    payload.imageUrl,           // Image_URL
    payload.productUrl,         // Product_URL
    new Date()                  // Last_Updated
  ];
  
  master.appendRow(rowData);
  log.appendRow([new Date(), Session.getActiveUser().getEmail(), "ADD", payload.itemId, payload.productName]);
  return "Successfully added " + payload.itemId;
}
```

### 2.3 Code.js - generateNextItemID()

**No changes required.** Current gap-filling logic is correct.

### 2.4 Code.js - mergeCategories()

**Update column references** to match new MASTER_DB schema:
- Item_ID is now column 1 (index 0)
- Category is now column 3 (index 2)

```javascript
function mergeCategories(keepId, mergeId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName('MASTER_DB');
  const config = ss.getSheetByName('SYSTEM_CONFIG');
  
  // Get keep category name from config (column F = index 2 in E:H range)
  const configData = config.getRange("E2:F100").getValues();
  const keepName = configData.find(r => r[0] === keepId)?.[1] || "";
  
  const lastRow = master.getLastRow();
  if (lastRow > 1) {
    const data = master.getRange(2, 1, lastRow - 1, 3).getValues(); // Item_ID, Division, Category
    for (let i = 0; i < data.length; i++) {
      if (data[i][0].toString().startsWith(mergeId + "-")) {
        let itemPart = data[i][0].toString().split('-')[2];
        let newId = keepId + "-" + itemPart;
        master.getRange(i + 2, 1).setValue("'" + newId);  // Update Item_ID
        master.getRange(i + 2, 3).setValue(keepName);      // Update Category
      }
    }
  }
  deleteCategoryFromConfig(mergeId);
  return "Merged items into " + keepId;
}
```

### 2.5 Code.js - addCategory()

**Update range reference** for SHOW_FIELDS column:

```javascript
function addCategory(divLabel, name, desc, showFields = "") {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = ss.getSheetByName('SYSTEM_CONFIG');
  const divId = divLabel.split(' : ')[0];
  
  // Find existing category IDs for this division
  const existingIds = config.getRange("E2:E100").getValues().flat()
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
  
  // Find first empty row
  const colD = config.getRange("D2:D100").getValues();
  let targetRow = 100;
  for (let i = 0; i < colD.length; i++) {
    if (colD[i][0] === "") {
      targetRow = i + 2;
      break;
    }
  }
  
  // Write 5 columns: Parent, ID, Name, Desc, ShowFields
  config.getRange(targetRow, 4, 1, 5).setValues([[divLabel, newCatId, name, desc, showFields]]);
  return "Successfully created " + newCatId + ": " + name;
}
```

### 2.6 Code.js - deleteCategoryFromConfig()

**Update range** to clear SHOW_FIELDS column too:

```javascript
function deleteCategoryFromConfig(catId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SYSTEM_CONFIG');
  const data = sheet.getRange("E2:E100").getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === catId) {
      sheet.getRange(i + 2, 4, 1, 5).clearContent();  // Clear D through H
      break;
    }
  }
  return "Removed " + catId;
}
```

---

## Phase 3: Sidebar.html Refactoring

### 3.1 Updated HTML Structure

```html
<!DOCTYPE html>
<html>
  <head>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/materialize/1.0.0/css/materialize.min.css">
    <style>
      body { padding: 15px; font-size: 13px; background-color: #f9f9f9; }
      .section-label { font-weight: bold; color: #1b5e20; border-bottom: 2px solid #1b5e20; margin-top: 15px; margin-bottom: 5px; padding-bottom: 3px; }
      .helper-text { font-size: 11px; color: #666; font-style: italic; background: #e8f5e9; padding: 8px; border-radius: 4px; margin-bottom: 10px; }
      .input-field label { pointer-events: none; transform: translateY(-14px) scale(0.8); transform-origin: 0 0; color: #1b5e20 !important; }
      select { display: block; width: 100%; height: 35px; margin-bottom: 10px; border: 1px solid #ccc; border-radius: 4px; }
      textarea { width: 100%; height: 60px; margin-bottom: 10px; border: 1px solid #ccc; border-radius: 4px; padding: 8px; resize: vertical; }
      .btn-save { width: 100%; background-color: #1b5e20 !important; margin-top: 20px; font-weight: bold; height: 45px; }
      .id-view { background-color: #eee !important; border: none !important; font-weight: bold; color: #d32f2f; padding-left: 8px !important; }
      .conditional-field { display: none; }
      .conditional-field.visible { display: block; }
    </style>
  </head>
  <body>
    <h6 style="text-align:center; font-weight:bold;">GSADU PRODUCT ENTRY</h6>
    
    <!-- Section 1: Categorization -->
    <p class="section-label">1. Categorization</p>
    <label>Division</label>
    <select id="division" onchange="filterCategories()">
      <option value="" disabled selected>Select Division</option>
    </select>
    <label>Category</label>
    <select id="category" onchange="updateCategoryInfo()">
      <option value="" disabled selected>Select Division First</option>
    </select>
    <div id="catDesc" class="helper-text">Description will appear here...</div>
    <div class="input-field">
      <input type="text" id="itemId" class="id-view" readonly>
      <label>Item ID (DIV-CAT-ITEM)</label>
    </div>
    
    <!-- Section 2: Product Details -->
    <p class="section-label">2. Product Details</p>
    <div class="input-field">
      <input type="text" id="productName" placeholder="Short product name">
      <label>Product Name</label>
    </div>
    <label>Description</label>
    <textarea id="description" placeholder="Detailed specs, model info, notes..."></textarea>
    <label>Selection Tier</label>
    <select id="tier">
      <!-- Options populated dynamically, Standard selected by default -->
    </select>
    
    <!-- Section 3: Design Options (Conditional) -->
    <div id="designSection" class="conditional-field">
      <p class="section-label">3. Design Options</p>
      <div id="bundleField" class="conditional-field">
        <label>Bundle</label>
        <select id="bundle">
          <option value="">None</option>
        </select>
      </div>
      <div id="finishField" class="conditional-field">
        <label>Finish</label>
        <select id="finish">
          <option value="">None</option>
        </select>
      </div>
    </div>
    
    <!-- Section 4: Resources -->
    <p class="section-label" id="resourcesLabel">3. Resources</p>
    <div class="input-field">
      <input type="text" id="imageUrl" placeholder="Direct link to product image">
      <label>Image URL</label>
    </div>
    <div class="input-field">
      <input type="text" id="productUrl" placeholder="Link to product page">
      <label>Product URL</label>
    </div>
    
    <button class="btn btn-save" onclick="submitData()">Save Product</button>
    
    <script>
      let fullConfig = {};
      let currentShowFields = [];
      
      // Initialize on load
      google.script.run.withSuccessHandler(data => {
        fullConfig = data;
        
        // Populate divisions
        const divSelect = document.getElementById('division');
        data.divisions.forEach(v => divSelect.add(new Option(v, v)));
        
        // Populate tiers with Standard as default
        const tierSelect = document.getElementById('tier');
        data.tiers.forEach((v, i) => {
          const opt = new Option(v, v);
          if (v === 'Standard') opt.selected = true;
          tierSelect.add(opt);
        });
        
        // Populate bundles
        const bundleSelect = document.getElementById('bundle');
        data.bundles.forEach(v => bundleSelect.add(new Option(v, v)));
        
        // Populate finishes
        const finishSelect = document.getElementById('finish');
        data.finishes.forEach(v => finishSelect.add(new Option(v, v)));
        
      }).getSelectionData();
      
      function filterCategories() {
        const div = document.getElementById('division').value;
        const catSelect = document.getElementById('category');
        catSelect.innerHTML = '<option value="" disabled selected>Select Category</option>';
        
        // categoryData: [Parent, ID, Name, Desc, ShowFields]
        fullConfig.categoryData
          .filter(c => c[0] === div)
          .forEach(c => {
            let opt = new Option(c[2], c[1]);  // Name as text, ID as value
            opt.dataset.desc = c[3];           // Description
            opt.dataset.showFields = c[4] || "";  // ShowFields
            catSelect.add(opt);
          });
        
        // Reset conditional fields
        hideDesignFields();
      }
      
      function updateCategoryInfo() {
        const catSelect = document.getElementById('category');
        const opt = catSelect.options[catSelect.selectedIndex];
        
        // Update description
        document.getElementById('catDesc').innerText = opt.dataset.desc || "No description.";
        
        // Generate next Item ID
        google.script.run
          .withSuccessHandler(id => document.getElementById('itemId').value = id)
          .generateNextItemID(opt.value);
        
        // Handle conditional design fields
        const showFields = opt.dataset.showFields || "";
        updateDesignFields(showFields);
      }
      
      function updateDesignFields(showFields) {
        const fields = showFields.split(',').map(f => f.trim().toLowerCase());
        currentShowFields = fields;
        
        const designSection = document.getElementById('designSection');
        const bundleField = document.getElementById('bundleField');
        const finishField = document.getElementById('finishField');
        const resourcesLabel = document.getElementById('resourcesLabel');
        
        const showBundle = fields.includes('bundle');
        const showFinish = fields.includes('finish');
        
        // Toggle visibility
        bundleField.classList.toggle('visible', showBundle);
        finishField.classList.toggle('visible', showFinish);
        designSection.classList.toggle('visible', showBundle || showFinish);
        
        // Update section numbering
        resourcesLabel.textContent = (showBundle || showFinish) ? '4. Resources' : '3. Resources';
        
        // Reset values if hidden
        if (!showBundle) document.getElementById('bundle').value = "";
        if (!showFinish) document.getElementById('finish').value = "";
      }
      
      function hideDesignFields() {
        document.getElementById('designSection').classList.remove('visible');
        document.getElementById('bundleField').classList.remove('visible');
        document.getElementById('finishField').classList.remove('visible');
        document.getElementById('resourcesLabel').textContent = '3. Resources';
        document.getElementById('bundle').value = "";
        document.getElementById('finish').value = "";
        currentShowFields = [];
      }
      
      function submitData() {
        const payload = {
          division: document.getElementById('division').value,
          category: document.getElementById('category').options[document.getElementById('category').selectedIndex].text,
          itemId: document.getElementById('itemId').value,
          productName: document.getElementById('productName').value,
          description: document.getElementById('description').value,
          tier: document.getElementById('tier').value,
          bundle: document.getElementById('bundle').value,
          finish: document.getElementById('finish').value,
          imageUrl: document.getElementById('imageUrl').value,
          productUrl: document.getElementById('productUrl').value
        };
        
        // Validation
        if (!payload.division || !payload.category || !payload.productName) {
          alert("Required: Division, Category, and Product Name");
          return;
        }
        
        google.script.run
          .withSuccessHandler(msg => { alert(msg); location.reload(); })
          .withFailureHandler(err => alert("Error: " + err))
          .saveToDatabase(payload);
      }
    </script>
  </body>
</html>
```

---

## Phase 4: CategoryManager.html Updates

### 4.1 Add SHOW_FIELDS to Category Creation

```html
<!-- In the "Add New Category" card, add: -->
<label>Show Fields (Optional)</label>
<input type="text" id="showFields" placeholder="Bundle, Finish, or Bundle,Finish">
<p style="font-size:10px; color:#888; margin-top:-8px;">Leave blank for most categories</p>
```

### 4.2 Update handleAdd() Function

```javascript
function handleAdd() {
  const div = document.getElementById('addDiv').value;
  const name = document.getElementById('newName').value;
  const desc = document.getElementById('newDesc').value;
  const showFields = document.getElementById('showFields').value;
  
  if(!div || !name) return alert("Required: Division and Name.");
  
  google.script.run
    .withSuccessHandler(msg => { alert(msg); load(); })
    .addCategory(div, name, desc, showFields);
}
```

---

## Phase 5: Implementation Checklist

### Pre-Implementation
- [ ] Backup current Google Sheet (File > Make a copy)
- [ ] Document current MASTER_DB data (if any beyond test entries)

### Step 1: SYSTEM_CONFIG Updates
- [ ] Insert column H with header "SHOW_FIELDS"
- [ ] Shift existing columns if needed (Tiers, Bundles, Finishes)
- [ ] Populate SHOW_FIELDS for applicable categories:
  - [ ] 16-01: Bundle
  - [ ] 16-02: Bundle
  - [ ] 16-03: Bundle,Finish
  - [ ] 18-01: Bundle
  - [ ] 07-03: Finish
  - [ ] Review and add others as needed

### Step 2: MASTER_DB Updates
- [ ] Clear existing test data (or migrate manually)
- [ ] Replace headers with new 11-column schema
- [ ] Format Last_Updated column as Date

### Step 3: Code.js Updates
- [ ] Update getSelectionData() - expand categoryData range
- [ ] Update saveToDatabase() - new payload structure
- [ ] Update mergeCategories() - fix column references
- [ ] Update addCategory() - add showFields parameter
- [ ] Update deleteCategoryFromConfig() - expand clear range

### Step 4: Sidebar.html Updates
- [ ] Replace entire file with new version
- [ ] Test division/category dropdowns
- [ ] Test conditional Bundle/Finish fields
- [ ] Test form submission

### Step 5: CategoryManager.html Updates
- [ ] Add SHOW_FIELDS input field
- [ ] Update handleAdd() function

### Step 6: Testing
- [ ] Test adding product to category WITHOUT Bundle/Finish
- [ ] Test adding product to category WITH Bundle only
- [ ] Test adding product to category WITH Finish only
- [ ] Test adding product to category WITH Bundle AND Finish
- [ ] Test category merge function
- [ ] Test category delete function
- [ ] Test category add function with SHOW_FIELDS

---

## Phase 6: Future Enhancements (Out of Scope)

- Edit existing products (currently append-only)
- Delete products from MASTER_DB
- Bulk import from CSV/existing data
- Image preview in sidebar
- Search/filter in sidebar
- Export to PDF (recreate formatted catalog)
- Pricing and vendor tracking (optional columns)

---

## Reference: Column Mappings

### SYSTEM_CONFIG Column Reference
| Column | Letter | Current Header |
|--------|--------|----------------|
| DIV_ID | A | DIV_ID |
| DIV_NAME | B | DIV_NAME |
| DIV_LABEL | C | DIV_LABEL |
| PARENT_DIV | D | PARENT_DIV |
| CAT_ID | E | CAT_ID |
| CAT_NAME | F | CAT_NAME |
| CAT_DESC | G | CAT_DESC |
| SHOW_FIELDS | H | **NEW** |
| (gap) | I | |
| TIERS | J | TIERS |
| (gap) | K | |
| BUNDLE_ID | L | BUNDLE_ID |
| BUNDLE_NAME | M | BUNDLE_NAME |
| (gap) | N | |
| FINISH_COLORS | O | FINISH_COLORS |

### MASTER_DB Column Reference (New)
| Column | Letter | Header |
|--------|--------|--------|
| 1 | A | Item_ID |
| 2 | B | Division |
| 3 | C | Category |
| 4 | D | Product_Name |
| 5 | E | Description |
| 6 | F | Tier |
| 7 | G | Bundle |
| 8 | H | Finish |
| 9 | I | Image_URL |
| 10 | J | Product_URL |
| 11 | K | Last_Updated |

---

## Document History

| Date | Version | Changes |
|------|---------|---------|
| 2025-01-21 | 1.0 | Initial planning document created |

---

## How to Use This Document

**For continuation with AI assistant:**
1. Share this entire document at the start of a new session
2. Reference specific phases/steps for targeted work
3. Update checklist as items are completed

**For manual implementation:**
1. Follow Phase 5 checklist in order
2. Test after each major step
3. Rollback to backup if issues arise
