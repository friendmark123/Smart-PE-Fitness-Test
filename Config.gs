// =============================================================================
// CONFIGURATION & CONSTANTS
// =============================================================================
const SPREADSHEET_ID = "1rAZyKOI7-C1VHgv6EPXQ5KiD039-DLcaM4xsYEXcSaA";
const FOLDER_PDF_ID   = "1oYAw15QVtc7bc80gXU85geK1tqXHZ39n";

const DOC_TEMPLATE_ID_INDIVIDUAL = "1foHM3r9xpZg9d34X7AFb6tAvvkmlGORWZdpQ5eL2TfE";
const DOC_TEMPLATE_ID_CLASS      = "1G6Wi2wynRGpyETcHOjvvN2vR8htABlnKGvDf_KGEFf8";
const DOC_TEMPLATE_ID_ANNUAL     = "1lijYX7evzSlnTbXLIGpd3aPhouc_6eiIMCyUiELtL3s";

function getDb() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}