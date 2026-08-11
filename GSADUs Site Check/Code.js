/**
 * GSADUs: MODULAR GENERATOR (v13 - Independent Zones)
 * ---------------------------------------------------
 * - Runs separate update cycles for Map vs. Images.
 * - Uses Element Titles ("GS_MAP", "GS_MODEL") to track and replace items.
 * - Logo: Strict 5.75" x 0.47" sizing.
 */

const CONFIG = {
  FOLDER_ID: '1tnai9lsArYCHGU62b56QLv1pMCKrglA_', 
  LOGO_ID:   '1VAi8GTuBRUSxNHA8mDt4Fi9O_yB7S3tT', 
  
  // PAGE SETUP (17x11)
  PAGE_WIDTH_PT: 17 * 72, 
  PAGE_HEIGHT_PT: 11 * 72,
  
  // LEFT ZONE (Site Plan)
  MAP_SIZE_PT: 10.5 * 72, 
  MAP_LEFT_X: 0.25 * 72,  
  
  // RIGHT ZONE (Specs Column)
  RIGHT_COL_X: 11 * 72,       
  RIGHT_COL_WIDTH: 5.75 * 72, 
  
  // LOGO SPECS (Strict)
  LOGO_Y: 0.25 * 72,
  LOGO_W: 5.75 * 72,
  LOGO_H: 0.47 * 72, 

  // SCALING MATH
  PTS_PER_FT: 3.9307,
  MODEL_PX_PER_FT: 38.79 
};

function onOpen() {
  SlidesApp.getUi().createMenu('GSADUs').addItem('Open Generator', 'showSidebar').addToUi();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('ADU Generator').setWidth(340);
  SlidesApp.getUi().showSidebar(html);
}

// --- DATA FETCHERS (Unchanged) ---
function getModelList() {
  try {
    var folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
    var files = folder.getFiles();
    var modelSet = new Set();
    while (files.hasNext()) {
      var name = files.next().getName();
      if (name.includes(".png") || name.includes(".jpg")) {
        var parts = name.split(" ");
        if (parts.length > 0) modelSet.add(parts[0]);
      }
    }
    var models = [];
    modelSet.forEach(m => models.push(m));
    return models.sort();
  } catch (e) { throw new Error("Folder Error: " + e.message); }
}

function getImagesForModel(modelName) {
  var folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
  var files = folder.getFiles();
  var imageList = [];
  var strictMatch = modelName + " ";

  while (files.hasNext()) {
    var file = files.next();
    var fName = file.getName();
    
    if (fName.indexOf(strictMatch) === 0 && (fName.includes(".png") || fName.includes(".jpg"))) {
      var isChecked = false;
      var sortOrder = 99;
      // SORTING LOGIC
      if (fName.includes("3D Plan")) { isChecked = true; sortOrder = 1; }   
      else if (fName.includes("Southeast")) { isChecked = true; sortOrder = 2; } 
      else if (fName.includes("Southwest")) { sortOrder = 3; }
      else if (fName.includes("Northeast")) { sortOrder = 4; }
      else if (fName.includes("Northwest")) { sortOrder = 5; }
      else if (fName.includes("Floorplan")) { isChecked = true; sortOrder = 6; } 
      
      imageList.push({ name: fName, id: file.getId(), isChecked: isChecked, sort: sortOrder });
    }
  }
  return imageList.sort((a, b) => a.sort - b.sort);
}

function fetchImageBase64(fileId) {
  var file = DriveApp.getFileById(fileId);
  return Utilities.base64Encode(file.getBlob().getBytes());
}

// --- ZONE UPDATERS ---

/** * ZONE 1: MAP UPDATE 
 * Deletes old map, inserts new one. Does NOT touch models.
 */
function updateMapOnly(inputString) {
  var slide = SlidesApp.getActivePresentation().getSlides()[0];
  
  // 1. Find and Delete existing Map
  var elements = slide.getPageElements();
  elements.forEach(function(e) {
    if (e.getTitle() === "GS_MAP") e.remove();
  });

  // 2. Insert New Map
  var location = inputString || "313 Lenka Ct, Roseville, CA"; 
  var mapBlob = Maps.newStaticMap()
    .setCenter(location)
    .setZoom(21)
    .setMapType(Maps.StaticMap.Type.SATELLITE)
    .setSize(1000, 1000)
    .getBlob();
    
  var mapImg = slide.insertImage(mapBlob);
  mapImg.setTitle("GS_MAP"); // TAG IT!
  
  // Position
  mapImg.setWidth(CONFIG.MAP_SIZE_PT);
  mapImg.setHeight(CONFIG.MAP_SIZE_PT);
  var mapTop = (CONFIG.PAGE_HEIGHT_PT - CONFIG.MAP_SIZE_PT) / 2; 
  mapImg.setLeft(CONFIG.MAP_LEFT_X);
  mapImg.setTop(mapTop);
  mapImg.sendToBack();
  
  // Opacity
  Utilities.sleep(200); 
  try {
    var resource = { requests: [{ updateImageProperties: {
          objectId: mapImg.getObjectId(),
          imageProperties: { transparency: 0.2 }, 
          fields: "imageProperties.transparency"
    }}]};
    Slides.Presentations.batchUpdate(resource, SlidesApp.getActivePresentation().getId());
  } catch (e) {}

  return "Map Updated.";
}

/** * ZONE 2: LOGO CHECK 
 * Ensures logo exists and is sized strictly.
 */
function ensureLogo() {
  var slide = SlidesApp.getActivePresentation().getSlides()[0];
  var exists = false;
  
  slide.getPageElements().forEach(function(e) {
    if (e.getTitle() === "GS_LOGO") exists = true;
  });

  if (!exists) {
    try {
      var logoFile = DriveApp.getFileById(CONFIG.LOGO_ID);
      var logoImg = slide.insertImage(logoFile.getBlob());
      logoImg.setTitle("GS_LOGO"); // TAG IT!
      
      // STRICT SIZING
      logoImg.setWidth(CONFIG.LOGO_W);
      logoImg.setHeight(CONFIG.LOGO_H);
      logoImg.setLeft(CONFIG.RIGHT_COL_X);
      logoImg.setTop(CONFIG.LOGO_Y);
    } catch(e) { console.warn("Logo missing"); }
  }
}

/** * ZONE 3: PREPARE IMAGE UPDATE
 * Clears old model images to prepare for new selection.
 * Returns Layout Data (Map Center & Stack Start Y).
 */
function prepareImageUpdate() {
  var slide = SlidesApp.getActivePresentation().getSlides()[0];
  
  // 1. Delete ONLY Model Images
  var elements = slide.getPageElements();
  elements.forEach(function(e) {
    if (e.getTitle() === "GS_MODEL") e.remove();
  });

  // 2. Ensure Logo (so we know where to stack below)
  ensureLogo();

  // 3. Calculate Layout Anchors
  var mapTop = (CONFIG.PAGE_HEIGHT_PT - CONFIG.MAP_SIZE_PT) / 2; 
  var mapCenter = { 
    x: CONFIG.MAP_LEFT_X + (CONFIG.MAP_SIZE_PT/2), 
    y: mapTop + (CONFIG.MAP_SIZE_PT/2) 
  };
  
  // Stack starts below Logo (Y + H + Padding)
  var startY = CONFIG.LOGO_Y + CONFIG.LOGO_H + 20;

  return { mapCenter: mapCenter, currentRightY: startY };
}

/** * ZONE 3: PLACE SINGLE IMAGE
 */
function placeMeasuredImage(fileId, nativeW, nativeH, layoutData, fileName, currentRightY) {
  var slide = SlidesApp.getActivePresentation().getSlides()[0];
  var file = DriveApp.getFileById(fileId);
  var img = slide.insertImage(file.getBlob());
  img.setTitle("GS_MODEL"); // TAG IT!
  
  var ratio = nativeW / nativeH;

  if (fileName.includes("3D Plan")) {
    // LEFT (Overlay)
    var realFeetWidth = nativeW / CONFIG.MODEL_PX_PER_FT;
    var targetWidthPt = realFeetWidth * CONFIG.PTS_PER_FT;
    var targetHeightPt = targetWidthPt / ratio;
    
    img.setWidth(targetWidthPt);
    img.setHeight(targetHeightPt);
    img.setLeft(layoutData.mapCenter.x - (targetWidthPt / 2));
    img.setTop(layoutData.mapCenter.y - (targetHeightPt / 2));
    
    return currentRightY; 
    
  } else {
    // RIGHT (Stack)
    var targetWidthPt = CONFIG.RIGHT_COL_WIDTH;
    var targetHeightPt = targetWidthPt / ratio;
    
    img.setWidth(targetWidthPt);
    img.setHeight(targetHeightPt);
    img.setLeft(CONFIG.RIGHT_COL_X);
    img.setTop(currentRightY); 
    
    return currentRightY + targetHeightPt + 20; 
  }
}