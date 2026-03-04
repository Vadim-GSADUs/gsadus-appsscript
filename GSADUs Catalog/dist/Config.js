const PATHS = {
    // Images
    IMAGE_ROOT: 'Support/PNG',
    // CSV input
    CSV_CATALOG: 'Support/CSV',
    // PDFs (optional)
    PDF_ROOT: 'Support/PDF',
    // AppSheet data (optional)
    APPSHEET_DATA: 'appsheet/data/GSADUsCatalog-434555248'
};
const CFG = Object.freeze({
    // Single Drive anchor: "Working" folder
    ROOT_ID: '1vYB2hmB4WfqksvMZDrxSR8l6SqK1ffT1',
    // Human-readable relative paths from ROOT_ID
    PATHS,
    // CSV selection rules
    CSV_CATALOG_BASENAME: 'GSADUs Catalog_Registry',
    // Sheets
    CATALOG_RAW_TAB: 'Catalog_Raw',
    CATALOG_TAB: 'Catalog',
    MODEL_HEADER: 'Model',
    // Optional publish target
    PRODUCTION_SHEET_ID: ''
});
// quick deploy ping
function __ping() {
    Logger.log('ts ok');
}
