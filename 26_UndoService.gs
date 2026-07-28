/**
 * Inventory PRO 4.3.4 — konfliktowe cofanie importu.
 * Ręczne zmiany wykonane po imporcie nie są nadpisywane.
 */

function undoLastImport() {
  const ui = SpreadsheetApp.getUi();
  const importId = getLastActiveImportId_();
  if (!importId) {
    ui.alert('Nie znaleziono importu do cofnięcia.');
    return;
  }

  const response = ui.alert(
    'Cofnij ostatni import',
    'Czy na pewno cofnąć import:\n' + importId + '?\n\n' +
      'Komórki zmienione później ręcznie zostaną pominięte i zaraportowane jako konflikty.',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  const result = undoImportById_(importId);
  ui.alert(
    'Cofanie importu',
    'Przywrócono pozycji: ' + result.restoredCount +
      '\nPrzywrócono komórek: ' + result.restoredCellCount +
      '\nKonflikty — późniejsze zmiany ręczne: ' + result.conflictCount +
      '\nPominięto nieprawidłowych: ' + result.skippedCount,
    ui.ButtonSet.OK
  );
  return result;
}

function undoImportById_(importId) {
  const lock = LockService.getDocumentLock();
  try {
    lock.waitLock(30000);
    const inventorySheet = getSheetByConfiguredName_(CONFIG.SHEETS.INVENTORY);
    const historySheet = getOrCreateTechnicalHistorySheet_();
    if (!inventorySheet) throw new Error('Nie znaleziono arkusza INWENTURA.');

    const lastRow = historySheet.getLastRow();
    if (lastRow < 2) {
      return {
        importId: importId,
        restoredCount: 0,
        restoredCellCount: 0,
        conflictCount: 0,
        skippedCount: 0
      };
    }

    const values = historySheet.getRange(2, 1, lastRow - 1, 16).getValues();
    const plan = buildUndoPlan_(values, importId);
    const formulaCellIndex = buildInventoryFormulaCellIndex_(scanInventoryProducts_());
    const evaluated = evaluateUndoPlanAgainstInventory_(inventorySheet, plan, formulaCellIndex);
    const undoneAt = new Date();
    const undoneBy = getCurrentUserEmail_();

    const inventorySnapshots = evaluated.applicableChanges.map(change => {
      const range = inventorySheet.getRange(change.column + change.row);
      return { range: range, value: range.getValue(), formula: range.getFormula() };
    });
    const allAuditRows = Array.from(new Set(
      evaluated.applicableAuditRows
        .concat(evaluated.conflictAuditRows, plan.invalidAuditRows)
    ));
    const auditSnapshots = allAuditRows.map(historyRow => ({
      historyRow: historyRow,
      values: historySheet.getRange(historyRow, 13, 1, 3).getValues()
    }));

    try {
      evaluated.applicableChanges.forEach(change => {
        const range = inventorySheet.getRange(change.column + change.row);
        if (isUndoBlankValue_(change.previousValue)) range.clearContent();
        else range.setValue(change.previousValue);
      });
      SpreadsheetApp.flush();
      refreshInventoryStatusColors_(inventorySheet);

      setUndoAuditStatus_(historySheet, evaluated.applicableAuditRows, 'UNDONE', undoneAt, undoneBy);
      setUndoAuditStatus_(historySheet, evaluated.conflictAuditRows, 'UNDO_CONFLICT', undoneAt, undoneBy);
      setUndoAuditStatus_(historySheet, plan.invalidAuditRows, 'UNDO_SKIPPED', undoneAt, undoneBy);
      SpreadsheetApp.flush();
    } catch (error) {
      inventorySnapshots.forEach(snapshot => {
        if (snapshot.formula) snapshot.range.setFormula(snapshot.formula);
        else if (isUndoBlankValue_(snapshot.value)) snapshot.range.clearContent();
        else snapshot.range.setValue(snapshot.value);
      });
      auditSnapshots.forEach(snapshot => {
        historySheet.getRange(snapshot.historyRow, 13, 1, 3).setValues(snapshot.values);
      });
      SpreadsheetApp.flush();
      throw error;
    }

    const result = {
      importId: importId,
      restoredCount: evaluated.applicableAuditRows.length,
      restoredCellCount: evaluated.applicableChanges.length,
      conflictCount: evaluated.conflictChanges.length,
      conflictAuditRows: evaluated.conflictAuditRows.length,
      skippedCount: plan.invalidAuditRows.length,
      conflicts: evaluated.conflictChanges.map(change => ({
        cell: change.column + change.row,
        expectedImportValue: change.expectedNewValue,
        currentValue: change.liveValue
      }))
    };

    logInfo(
      'UndoService',
      'undoImportById_',
      'Zakończono cofanie importu',
      result
    );
    return result;
  } finally {
    lock.releaseLock();
  }
}

function isUndoBlankValue_(value) {
  return value === '' || value === null || value === undefined;
}

function setUndoAuditStatus_(historySheet, rows, status, timestamp, user) {
  (rows || []).forEach(historyRow => {
    historySheet.getRange(historyRow, 13, 1, 3)
      .setValues([[status, timestamp, user]]);
  });
}

/**
 * Jeżeli kilka pozycji jednego importu zapisało się do tej samej komórki,
 * przywracamy stan sprzed pierwszego zapisu, ale oczekujemy wartości po
 * ostatnim zapisie z grupy.
 */
function buildUndoPlan_(historyValues, importId) {
  const groups = {};
  const auditRows = [];
  const invalidAuditRows = [];

  (historyValues || []).forEach((row, index) => {
    const rowImportId = String(row[0] || '').trim();
    const status = String(row[12] || '').trim();
    const alreadyUndone = row[13];
    if (rowImportId !== importId || status !== 'SAVED' || alreadyUndone) return;

    const historyRow = index + 2;
    const targetRow = Number(row[8]);
    const targetColumn = normalizeColumnLetter_(row[9]);
    if (!Number.isInteger(targetRow) || targetRow < 1 || !targetColumn) {
      invalidAuditRows.push(historyRow);
      return;
    }

    const key = targetColumn + targetRow;
    if (!groups[key]) {
      groups[key] = {
        row: targetRow,
        column: targetColumn,
        previousValue: row[10],
        expectedNewValue: row[11],
        auditRows: []
      };
    } else {
      groups[key].expectedNewValue = row[11];
    }
    groups[key].auditRows.push(historyRow);
    auditRows.push(historyRow);
  });

  return {
    changes: Object.keys(groups).map(key => groups[key]),
    auditRows: auditRows,
    invalidAuditRows: invalidAuditRows
  };
}

function evaluateUndoPlanAgainstInventory_(inventorySheet, plan, formulaCellIndex) {
  const applicableChanges = [];
  const conflictChanges = [];
  const applicableAuditRows = [];
  const conflictAuditRows = [];
  const formulaIndex = formulaCellIndex || {};

  (plan.changes || []).forEach(change => {
    const a1 = change.column + change.row;
    const range = inventorySheet.getRange(a1);
    const liveFormula = range.getFormula();
    const liveValue = range.getValue();
    const isProtectedFormulaCell = Boolean(formulaIndex[a1] || liveFormula);
    const matchesImport = inventoryCellValuesEqual_(liveValue, change.expectedNewValue);

    if (!isProtectedFormulaCell && matchesImport) {
      applicableChanges.push(change);
      applicableAuditRows.push.apply(applicableAuditRows, change.auditRows || []);
    } else {
      const conflict = Object.assign({}, change, {
        liveValue: liveValue,
        liveFormula: liveFormula,
        reason: isProtectedFormulaCell ? 'FORMULA_PROTECTED' : 'VALUE_CHANGED'
      });
      conflictChanges.push(conflict);
      conflictAuditRows.push.apply(conflictAuditRows, change.auditRows || []);
    }
  });

  return {
    applicableChanges: applicableChanges,
    conflictChanges: conflictChanges,
    applicableAuditRows: applicableAuditRows,
    conflictAuditRows: conflictAuditRows
  };
}
