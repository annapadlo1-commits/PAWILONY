/**
 * Ujednolica wyłącznie sposób wyświetlania liczb.
 * Nie zmienia wartości ani formuł.
 */
function normalizeInventoryNumberFormats_() {
  const sheet = getSheetByConfiguredName_(CONFIG.SHEETS.INVENTORY);
  if (!sheet) throw new Error('Nie znaleziono arkusza: ' + CONFIG.SHEETS.INVENTORY + '.');
  let formattedCells = 0;

  scanInventoryProducts_().forEach(function(product) {
    const columns = {};
    if (isDirectFinalInventoryProduct_(product)) {
      columns.B = true;
    } else {
      const layout = getConfiguredInventoryLayout_(product.type) || {};
      Object.keys(layout).forEach(function(key) {
        const values = Array.isArray(layout[key]) ? layout[key] : [layout[key]];
        values.forEach(function(value) {
          const match = String(value || '').trim().toUpperCase().match(/^([A-Z]+)$/);
          if (match) columns[match[1]] = true;
        });
      });
    }

    Object.keys(columns).forEach(function(letter) {
      // „General” respektuje lokalizację arkusza i nie wyświetla 12 jako „12,”.
      sheet.getRange(letter + product.inventoryRow).setNumberFormat('General');
      formattedCells += 1;
    });
  });

  SpreadsheetApp.flush();
  return { formattedCells: formattedCells };
}

function normalizeInventoryNumberFormats() {
  const result = normalizeInventoryNumberFormats_();
  SpreadsheetApp.getUi().alert(
    'Inventory PRO',
    'Uporządkowano format liczb w ' + result.formattedCells +
      ' komórkach. Wartości i formuły nie zostały zmienione.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
