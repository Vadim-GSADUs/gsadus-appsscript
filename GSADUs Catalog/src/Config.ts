// Config (typed)
interface Paths { [k: string]: string }
interface Config {
  ROOT_ID: string;
  PATHS: Paths;
  CSV_CATALOG_BASENAME: string;
  CATALOG_RAW_TAB: string;
  CATALOG_TAB: string;
  MODEL_HEADER: string;
  PRODUCTION_SHEET_ID: string;
}

const PATHS: Paths = {
  // Images
  IMAGE_ROOT: 'Support/PNG',

  // CSV input
  CSV_CATALOG: 'Support/CSV',

  // PDFs (optional)
  PDF_ROOT: 'Support/PDF',

  // AppSheet data (optional)
  APPSHEET_DATA: 'appsheet/data/GSADUsCatalog-434555248'
};

const CFG: Config = Object.freeze({
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
