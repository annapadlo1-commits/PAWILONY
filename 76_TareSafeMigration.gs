/**
 * Jednorazowa, kontrolowana migracja formuły otwartego opakowania.
 * Zmienia wyłącznie dokładną starszą formułę „brutto - tara”.
 * Nie uruchamia ogólnej automatycznej naprawy formuł.
 */
function migrateTareSafeFormulas() {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  const sheet = getSheetByConfiguredName_(CONFIG.SHEETS.INVENTORY);
  if (!sheet) {
    lock.releaseLock();
    throw new Error('Nie znaleziono arkusza: ' + CONFIG.SHEETS.INVENTORY + '.');
  }

  const changes = [];
  try {
    scanInventoryProducts_().forEach(product => {
      if (isDirectFinalInventoryProduct_(product)) return;
      const type = String(product.type || '').trim().toUpperCase();
      if (type === CONFIG.PRODUCT_TYPES.LOCATION) return;
      const layout = getConfiguredInventoryLayout_(type);
      const contract = getInventoryFormulaContract_(product)
        .find(item => item.column === normalizeColumnLetter_(layout.openNet));
      if (!contract) throw new Error('Brak kontraktu formuły dla produktu „' + product.name + '”.');

      const row = Number(product.inventoryRow);
      const range = sheet.getRange(row, contract.columnNumber);
      const current = range.getFormula();
      const legacy = '=' + layout.grossWeight + row + '-' + layout.emptyContainerWeight + row;
      const normalizedCurrent = normalizeInventoryFormula_(current);
      if (normalizedCurrent === normalizeInventoryFormula_(contract.formula)) return;
      if (normalizedCurrent !== normalizeInventoryFormula_(legacy)) {
        throw new Error(
          'Migrację zablokowano: komórka ' + contract.column + row +
          ' nie zawiera oczekiwanej starszej formuły.'
        );
      }
      changes.push({range: range, previous: current, expected: contract.formula, r1c1: contract.r1c1});
    });

    changes.forEach(change => change.range.setFormulaR1C1(change.r1c1));
    SpreadsheetApp.flush();
    changes.forEach(change => {
      if (normalizeInventoryFormula_(change.range.getFormula()) !== normalizeInventoryFormula_(change.expected)) {
        throw new Error('Kontrola migracji nie powiodła się dla ' + change.range.getA1Notation() + '.');
      }
    });
    return {success: true, changedCells: changes.length};
  } catch (error) {
    changes.slice().reverse().forEach(change => {
      if (normalizeInventoryFormula_(change.range.getFormula()) === normalizeInventoryFormula_(change.expected)) {
        change.range.setFormula(change.previous);
      }
    });
    SpreadsheetApp.flush();
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function migrateTareSafeFormulasWithDialog() {
  const result = migrateTareSafeFormulas();
  SpreadsheetApp.getUi().alert(
    'Inventory PRO — bezpieczne liczenie tary',
    'Migracja zakończona. Zmienione komórki: ' + result.changedCells + '.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  return result;
}
