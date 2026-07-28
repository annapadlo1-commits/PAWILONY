/** Inventory PRO Enterprise v2.9.0 — modele raportowe. */
function createInventoryReportModel_(data) {
  const items = Array.isArray(data.items) ? data.items : [];
  return {
    schemaVersion: '2.10.0',
    metadata: {
      applicationVersion: CONFIG.VERSION,
      spreadsheetName: SpreadsheetApp.getActiveSpreadsheet().getName(),
      sourceSheetName: data.sourceSheetName,
      generatedAt: data.generatedAt || new Date(),
      generatedBy: getCurrentUserEmail_(),
      durationMs: Number(data.durationMs || 0)
    },
    items: items,
    normal: items.filter(item => item.type === CONFIG.PRODUCT_TYPES.NORMAL),
    keg: items.filter(item => item.type === CONFIG.PRODUCT_TYPES.KEG),
    location: items.filter(item => item.type === CONFIG.PRODUCT_TYPES.LOCATION),
    warnings: Array.isArray(data.warnings) ? data.warnings : [],
    validationIssues: Array.isArray(data.validationIssues) ? data.validationIssues : [],
    summary: {},
    statistics: {}
  };
}

function createReportIssue_(severity, product, message) {
  return {
    severity: String(severity || 'WARNING').toUpperCase(),
    product: String(product || ''),
    message: String(message || '')
  };
}