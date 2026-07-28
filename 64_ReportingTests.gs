/** Inventory PRO Enterprise v2.10.2 — testy silnika raportowego. */
function runReportingEngineTests() {
  const tests = [
    testReportingSummaryColumns290_,
    testReportingAllowedCategories290_,
    testReportingNoFallbackCategory290_,
    testReportingReadOnlyContract290_
  ];
  const results = tests.map(test => {
    try { test(); return { name: test.name, passed: true, message: 'PASS' }; }
    catch (error) { return { name: test.name, passed: false, message: error.message || String(error) }; }
  });
  const passed = results.filter(result => result.passed).length;
  SpreadsheetApp.getUi().alert(
    '🍕 Reporting Engine — Testy',
    'Testy: ' + results.length + '\nPASS: ' + passed + '\nFAIL: ' + (results.length - passed) + '\n\n' + results.map(result => (result.passed ? '✓ ' : '✗ ') + result.name + ': ' + result.message).join('\n'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  return results;
}

function testReportingSummaryColumns290_() {
  ['NORMAL', 'KEG', 'LOCATION'].forEach(function(type) {
    if (getInventorySummaryLayout_(type).finalTotal !== getConfiguredInventoryLayout_(type).finalTotal) {
      throw new Error(type + ' musi czytać kolumnę finalną skonfigurowaną dla lokalu.');
    }
  });
}

function testReportingAllowedCategories290_() {
  if (getBusinessCategories().indexOf('POZOSTAŁE') !== -1) throw new Error('POZOSTAŁE nie może być kategorią biznesową.');
  ['WINO', 'BITTER', 'SOFTY'].forEach(category => {
    if (getBusinessCategories().indexOf(category) === -1) throw new Error('Brak kategorii: ' + category);
  });
}

function testReportingNoFallbackCategory290_() {
  let failed = false;
  try { createConfigurationProduct_('Test', 'test', '', 'NORMAL', 999); }
  catch (error) { failed = true; }
  if (!failed) throw new Error('Produkt bez fizycznej kategorii powinien zostać odrzucony.');
}

function testReportingReadOnlyContract290_() {
  const source = String(generateInventoryReport_).toLowerCase();
  ['setvalue(', 'setvalues(', 'clearcontent(', 'appendrow('].forEach(token => {
    if (source.indexOf(token) !== -1) throw new Error('Silnik raportowy zawiera operację zapisu: ' + token);
  });
}
