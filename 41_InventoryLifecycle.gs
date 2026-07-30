/**
 * Inventory PRO Enterprise v2.1.2 LTS
 * Archiwizacja i czyszczenie inwentury.
 */

function startNewInventory() {
  const sheet = getSheetByConfiguredName_(CONFIG.SHEETS.INVENTORY);
  if (!sheet) throw new Error('Nie znaleziono arkusza Inwentura.');
  if (hasCurrentInventoryData_(sheet)) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Najpierw zakończ bieżącą inwentaryzację. Otwieram końcowy przegląd.','Inventory PRO',8);
    showFinalReview();
    return;
  }
  clearInventoryStatusColors_(sheet);
  startInventorySession_();
  appendApplicationEvent_('INVENTORY_STARTED','Rozpoczęto nową inwentaryzację',{});
  SpreadsheetApp.getActiveSpreadsheet().toast('Nowa inwentaryzacja jest aktywna.','Inventory PRO',6);
}

function hasCurrentInventoryData_(sheet) {
  const products = buildProductCatalog();
  return products.some(product => {
    if (!product.inventoryRow) return false;
    const directFinal = getDirectFinalInventoryColumn_(product);
    if (directFinal) {
      return sheet.getRange(directFinal + product.inventoryRow).getValue() !== '';
    }
    const c=product.columns||{};
    const cols=product.type===CONFIG.PRODUCT_TYPES.LOCATION?[c.warehouse,c.darkroom,c.fridges]:[c.weight,c.quantity];
    return cols.filter(Boolean).some(col=>sheet.getRange(col+product.inventoryRow).getValue() !== '');
  });
}

function clearCurrentInventory() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Wyczysc biezaca inwenture',
    'Dane zostana wyczyszczone bez tworzenia archiwum. Kontynuowac?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  const sheet = getSheetByConfiguredName_(CONFIG.SHEETS.INVENTORY);

  if (!sheet) {
    throw new Error('Nie znaleziono arkusza Inwentura.');
  }

  const result = clearCurrentInventoryData_(sheet);
  clearUnknownTareGrossApprovals_();
  clearInventoryStatusColors_(sheet);

  logInfo(
    'InventoryLifecycle',
    'clearCurrentInventory',
    'Wyczyszczono biezaca inwenture',
    result
  );

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Wyczyszczono ' + result.clearedCells + ' pol.',
    'Inventory PRO',
    8
  );
}

function clearCurrentInventoryData_(sheet) {
  const products = buildProductCatalog();
  const a1Set = {};
  const prepared = [];

  products.forEach(product => {
    if (!product.inventoryRow) return;

    const type = String(product.type || '').trim().toUpperCase();
    const directFinal = getDirectFinalInventoryColumn_(product);
    const configuredColumns = getInputColumnsForProductType_(type);
    const candidates = directFinal
      ? [directFinal]
      : type === CONFIG.PRODUCT_TYPES.LOCATION
      ? [configuredColumns.warehouse, configuredColumns.darkroom, configuredColumns.fridges]
      : [configuredColumns.weight, configuredColumns.quantity];

    candidates.filter(Boolean).forEach(column => {
      const safeColumn = assertSafeInventoryTargetColumn_(product, column);
      const a1 = safeColumn + product.inventoryRow;
      if (a1Set[a1]) return;
      const range = sheet.getRange(a1);
      const formula = range.getFormula();
      if (formula) {
        throw new Error(
          'Reset inwentaryzacji zablokowany: komórka wejściowa ' + a1 +
          ' zawiera formułę. Sprawdź konfigurację produktu „' + product.name + '”.'
        );
      }
      a1Set[a1] = true;
      prepared.push({ a1: a1, range: range, previousValue: range.getValue() });
    });
  });

  if (prepared.length) {
    sheet.getRangeList(prepared.map(item => item.a1)).clearContent();
  }

  // Formuły są częścią konstrukcji arkusza, nie danymi inwentaryzacji.
  // Odtwarzamy je po każdym czyszczeniu, również gdy zostały wcześniej
  // przypadkowo usunięte albo pochodziły ze starego trybu opakowania.
  const restoredFormulas = applyCanonicalFormulasToProductsBatch_(sheet, products);

  SpreadsheetApp.flush();

  const residualCells = prepared
    .filter(item => item.range.getValue() !== '')
    .map(item => item.a1);
  if (residualCells.length) {
    throw new Error(
      'Nie uruchomiono nowej inwentaryzacji: po czyszczeniu pozostały dane w ' +
      residualCells.slice(0, 12).join(', ') +
      (residualCells.length > 12 ? ' oraz ' + (residualCells.length - 12) + ' kolejnych komórkach.' : '.')
    );
  }

  return {
    clearedCells: prepared.length,
    clearedAddresses: prepared.map(item => item.a1),
    restoredFormulaCells: restoredFormulas.formulaCells,
    residualCells: []
  };
}

function createUniqueArchiveName_(spreadsheet, baseName) {
  let name = baseName;
  let counter = 2;

  while (spreadsheet.getSheetByName(name)) {
    name = baseName + ' (' + counter + ')';
    counter++;
  }

  return name;
}
