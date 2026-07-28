/** Inventory PRO Enterprise v2.9.0 — walidacja raportu bez zapisu do źródła. */
function validateInventoryReport(sourceSheetName) {
  const report = generateInventoryReport(sourceSheetName);
  return {
    valid: !report.validationIssues.some(issue => issue.severity === 'ERROR'),
    sourceSheetName: report.metadata.sourceSheetName,
    products: report.summary.products,
    issues: report.validationIssues,
    durationMs: report.metadata.durationMs
  };
}

function showReportingValidation() {
  const result = validateInventoryReport();
  const lines = [
    'Arkusz: ' + result.sourceSheetName,
    'Produkty: ' + result.products,
    'Status: ' + (result.valid ? 'POPRAWNY' : 'BŁĘDY'),
    'Czas: ' + result.durationMs + ' ms',
    ''
  ];
  if (!result.issues.length) lines.push('Nie znaleziono problemów.');
  else result.issues.slice(0, 30).forEach(issue => lines.push(issue.severity + ' — ' + issue.product + ': ' + issue.message));
  if (result.issues.length > 30) lines.push('... oraz ' + (result.issues.length - 30) + ' kolejnych problemów.');
  SpreadsheetApp.getUi().alert('🍕 Walidacja Reporting Engine 2.9.0', lines.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
  return result;
}