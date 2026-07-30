/**
 * Inventory PRO 4.3.6 SAFE MODE — bezpieczna odbudowa formuł PAWILONÓW.
 *
 * Naprawa nie nadpisuje konfliktowych wartości. Komórki puste, spłaszczone
 * z wynikiem zgodnym z obliczeniem oraz błędne formuły mogą zostać naprawione
 * po utworzeniu pełnej kopii zakładki.
 */

function buildInventoryFormulaRepairPlan_(sheet, products, audit) {
  const formulaAudit = audit || buildInventoryFormulaAudit_(sheet, products || []);
  const values = sheet.getRange(
    1, 1, Math.max(sheet.getLastRow(), 1), Math.max(getInventoryLayoutMaxColumn_(), 1)
  ).getValues();
  const byCell = {};

  [
    ['MISSING', formulaAudit.missing],
    ['FLATTENED_SAFE', formulaAudit.flattened],
    ['LEGACY_EQUIVALENT', formulaAudit.legacy],
    ['INVALID_FORMULA', formulaAudit.invalid],
    ['CALCULATION_ERROR', formulaAudit.calculationErrors]
  ].forEach(group => group[1].forEach(issue => {
      if (byCell[issue.cell]) return;
      const rowValues = values[issue.row - 1] || [];
      const currentValue = rowValues[issue.columnNumber - 1];
      byCell[issue.cell] = {
        row: issue.row,
        column: issue.column,
        columnNumber: issue.columnNumber,
        a1: issue.cell,
        formula: issue.expectedFormula,
        r1c1: issue.expectedR1C1,
        previousFormula: issue.actualFormula || '',
        previousValue: currentValue,
        product: issue.product,
        category: issue.category,
        type: issue.type
        ,classification: group[0]
      };
    }));

  return Object.keys(byCell).map(key => byCell[key]).sort((left, right) =>
    left.columnNumber - right.columnNumber || left.row - right.row
  );
}

function buildFormulaRepairSegments_(plan) {
  const groups = {};
  (plan || []).forEach(change => {
    const key = change.columnNumber + '|' + change.r1c1;
    if (!groups[key]) groups[key] = [];
    groups[key].push(change);
  });

  const segments = [];
  Object.keys(groups).forEach(key => {
    const changes = groups[key].sort((a, b) => a.row - b.row);
    let start = null;
    let previous = null;
    changes.forEach(change => {
      if (!start || change.row !== previous.row + 1) {
        if (start) {
          segments.push({
            startRow: start.row,
            endRow: previous.row,
            columnNumber: start.columnNumber,
            r1c1: start.r1c1
          });
        }
        start = change;
      }
      previous = change;
    });
    if (start) {
      segments.push({
        startRow: start.row,
        endRow: previous.row,
        columnNumber: start.columnNumber,
        r1c1: start.r1c1
      });
    }
  });

  return segments.sort((left, right) =>
    left.columnNumber - right.columnNumber || left.startRow - right.startRow
  );
}

function createFormulaRepairBackupSheet_(sheet) {
  const spreadsheet = sheet.getParent();
  const timestamp = Utilities.formatDate(
    new Date(),
    spreadsheet.getSpreadsheetTimeZone() || 'Europe/Warsaw',
    'yyyyMMdd-HHmmss'
  );
  const base = 'BACKUP FORMULY ' + timestamp;
  let name = base;
  let suffix = 2;
  while (spreadsheet.getSheetByName(name)) {
    name = base + '-' + suffix;
    suffix++;
  }
  const copy = sheet.copyTo(spreadsheet).setName(name);
  copy.hideSheet();
  return name;
}

function preflightInventoryFormulaRepairPlan_(sheet, plan) {
  (plan || []).forEach(change => {
    const range = sheet.getRange(change.row, change.columnNumber);
    const liveFormula = range.getFormula();
    const liveValue = range.getValue();
    const sameFormula = normalizeInventoryFormula_(liveFormula) ===
      normalizeInventoryFormula_(change.previousFormula || '');
    const sameValue = change.previousFormula
      ? true
      : inventoryCellValuesEqual_(liveValue, change.previousValue);
    if (!sameFormula || !sameValue) {
      throw new Error(
        'Komórka ' + change.a1 +
        ' została zmieniona po audycie. Naprawę przerwano bez nadpisywania ręcznej zmiany.'
      );
    }
  });
}

function validateInventoryFormulaRepairPlanTargets_(plan) {
  (plan || []).forEach(change => {
    const classifications = [
      'MISSING', 'FLATTENED_SAFE', 'LEGACY_EQUIVALENT',
      'INVALID_FORMULA', 'CALCULATION_ERROR'
    ];
    if (classifications.indexOf(String(change.classification || '')) === -1) {
      throw new Error(
        'Naprawę formuł przerwano bez zmian: pozycja ' +
        (change.a1 || '?') + ' nie ma bezpiecznej klasyfikacji.'
      );
    }
    const type = String(change.type || '').trim().toUpperCase();
    const layout = getConfiguredInventoryLayout_(type);
    if (!layout) {
      throw new Error(
        'Naprawę formuł przerwano: brak układu kolumn dla typu ' +
        (type || '(pusty)') + ' w pozycji ' + (change.a1 || '?') + '.'
      );
    }
    const target = normalizeColumnLetter_(change.column);
    const allowed = (layout.formulaColumns || []).map(normalizeColumnLetter_);
    const configuredNumber = inventoryColumnLetterToNumber_(target);
    if (
      !target ||
      allowed.indexOf(target) === -1 ||
      configuredNumber !== Number(change.columnNumber)
    ) {
      throw new Error(
        'Naprawę formuł przerwano bez zmian: komórka ' +
        (change.a1 || '?') + ' wskazuje kolumnę ' + (target || '?') +
        ', niedozwoloną dla typu ' + type + '. Dozwolone kolumny formuł: ' +
        allowed.join(', ') + '.'
      );
    }
  });
  return true;
}

function applyInventoryFormulaRepairPlan_(sheet, plan) {
  // Cały plan jest sprawdzany przed pierwszym zapisem. Dzięki temu plan
  // utworzony dla innego układu arkusza nie może przesunąć formuł ani danych.
  validateInventoryFormulaRepairPlanTargets_(plan);
  preflightInventoryFormulaRepairPlan_(sheet, plan);
  const segments = buildFormulaRepairSegments_(plan);
  segments.forEach(segment => {
    sheet.getRange(
      segment.startRow,
      segment.columnNumber,
      segment.endRow - segment.startRow + 1,
      1
    ).setFormulaR1C1(segment.r1c1);
  });
  return segments.length;
}

function rollbackInventoryFormulaRepairPlan_(sheet, plan) {
  (plan || []).slice().reverse().forEach(change => {
    const range = sheet.getRange(change.row, change.columnNumber);
    const liveFormula = range.getFormula();
    if (
      normalizeInventoryFormula_(liveFormula) !==
      normalizeInventoryFormula_(change.formula)
    ) {
      // Komórka nie została jeszcze zapisana albo została później zmieniona
      // ręcznie. W obu sytuacjach rollback nie może jej nadpisywać.
      return;
    }
    if (change.previousFormula) {
      range.setFormula(change.previousFormula);
    } else if (
      change.previousValue === '' ||
      change.previousValue === null ||
      change.previousValue === undefined
    ) {
      range.clearContent();
    } else {
      range.setValue(change.previousValue);
    }
  });
}

function writeInventoryFormulaAuditReport_(audit) {
  const sheet = getOrCreateConfiguredSheet_(CONFIG.SHEETS.FORMULA_AUDIT);
  const headers = [
    'Status', 'Komórka', 'Produkt', 'Kategoria', 'Typ',
    'Formuła oczekiwana', 'Formuła obecna', 'Wartość obecna', 'Wynik oczekiwany'
  ];
  const rows = [];
  const append = function(status, issues) {
    (issues || []).forEach(issue => rows.push([
      status,
      issue.cell,
      issue.product,
      issue.category,
      issue.type,
      issue.expectedFormula,
      issue.actualFormula || '',
      issue.currentValue === undefined || issue.currentValue === null ? '' : issue.currentValue,
      issue.expectedValue === undefined || issue.expectedValue === null ? '' : issue.expectedValue
    ]));
  };

  append('KONFLIKT', audit.conflicts);
  append('BRAK', audit.missing);
  append('SPŁASZCZONA_ZGODNA', audit.flattened);
  append('STARSZA_POPRAWNA_FORMUŁA', audit.legacy);
  append('BŁĘDNA_FORMUŁA', audit.invalid);
  append('BŁĄD_OBLICZENIA', audit.calculationErrors);

  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#f6b26b');
  if (rows.length) sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
  if (!sheet.isSheetHidden()) sheet.hideSheet();
  return rows.length;
}

function formatFormulaConflictCells_(audit, limit) {
  return (audit.conflicts || []).slice(0, limit || 12)
    .map(issue => issue.cell + ' (' + issue.product + ')')
    .join(', ');
}

function repairInventoryFormulas_(options) {
  const settings = options || {};
  const lock = LockService.getDocumentLock();
  const startedAt = Date.now();
  let plan = [];
  let sheet = null;
  let backupSheetName = '';

  try {
    lock.waitLock(30000);
    sheet = settings.sheet || getSheetByConfiguredName_(CONFIG.SHEETS.INVENTORY);
    if (!sheet) throw new Error('Nie znaleziono arkusza: ' + CONFIG.SHEETS.INVENTORY);
    const products = settings.products || scanInventoryProducts_();
    const auditBefore = settings.audit || buildInventoryFormulaAudit_(sheet, products);
    writeInventoryFormulaAuditReport_(auditBefore);

    if (auditBefore.hasBlockingConflicts && settings.failOnConflicts !== false) {
      throw new Error(
        'Wykryto ' + auditBefore.conflictFormulaCells +
        ' konfliktowych komórek. Automatyczna naprawa została zablokowana. ' +
        'Przykłady: ' + formatFormulaConflictCells_(auditBefore, 8) + '.'
      );
    }

    plan = buildInventoryFormulaRepairPlan_(sheet, products, auditBefore);
    if (!plan.length) {
      return {
        success: auditBefore.safe,
        changedCells: 0,
        segments: 0,
        backupSheetName: '',
        conflictsRemaining: auditBefore.conflictFormulaCells,
        audit: auditBefore,
        durationMs: Date.now() - startedAt
      };
    }

    if (settings.createBackup !== false) {
      backupSheetName = createFormulaRepairBackupSheet_(sheet);
    }

    const segments = applyInventoryFormulaRepairPlan_(sheet, plan);
    SpreadsheetApp.flush();
    const auditAfter = buildInventoryFormulaAudit_(sheet, products);
    writeInventoryFormulaAuditReport_(auditAfter);

    const requireFullySafe = settings.requireFullySafe !== false;
    if (requireFullySafe && !auditAfter.safe) {
      throw new Error(
        'Kontrola po naprawie nie przeszła: brakujące=' + auditAfter.missingFormulaCells +
        ', spłaszczone=' + auditAfter.flattenedFormulaCells +
        ', konflikty=' + auditAfter.conflictFormulaCells +
        ', starsze=' + auditAfter.legacyFormulaCells +
        ', nieprawidłowe=' + auditAfter.invalidFormulaCells +
        ', błędy=' + auditAfter.errorFormulaCells + '.'
      );
    }

    const result = {
      success: auditAfter.safe || (!requireFullySafe && auditAfter.repairableFormulaCells === 0),
      fullySafe: auditAfter.safe,
      changedCells: plan.length,
      segments: segments,
      backupSheetName: backupSheetName,
      conflictsRemaining: auditAfter.conflictFormulaCells,
      audit: auditAfter,
      durationMs: Date.now() - startedAt,
      source: settings.source || 'manual'
    };

    logInfo(
      'FormulaRepair',
      'repairInventoryFormulas_',
      auditAfter.safe
        ? 'Odbudowano formuły PAWILONÓW.'
        : 'Odbudowano bezpieczne formuły; konflikty pozostawiono bez zmian.',
      {
        sheet: sheet.getName(),
        changedCells: result.changedCells,
        segments: result.segments,
        backupSheetName: result.backupSheetName,
        conflictsRemaining: result.conflictsRemaining,
        source: result.source
      },
      result.durationMs
    );
    return result;
  } catch (error) {
    if (sheet && plan.length) {
      try {
        rollbackInventoryFormulaRepairPlan_(sheet, plan);
        SpreadsheetApp.flush();
      } catch (rollbackError) {
        logError('FormulaRepair', 'repairInventoryFormulas_.rollback', rollbackError, null, 0);
      }
    }
    logError(
      'FormulaRepair',
      'repairInventoryFormulas_',
      error,
      { changedCellsPlanned: plan.length, backupSheetName: backupSheetName },
      Date.now() - startedAt
    );
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function applyCanonicalFormulasToProductRow_(sheet, product, packagingProfileOverride) {
  const contracts = getInventoryFormulaContract_(product, packagingProfileOverride);
  contracts.forEach(contract => {
    sheet.getRange(product.inventoryRow, contract.columnNumber).setFormula(contract.formula);
  });
  return contracts.length;
}

function verifyCanonicalFormulasForProductRow_(sheet, product) {
  return getInventoryFormulaContract_(product).map(contract => {
    const actual = sheet.getRange(product.inventoryRow, contract.columnNumber).getFormula();
    return {
      cell: contract.column + product.inventoryRow,
      expected: contract.formula,
      actual: actual,
      valid: normalizeInventoryFormula_(actual) === normalizeInventoryFormula_(contract.formula)
    };
  });
}

function repairInventoryFormulasWithDialog() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Inventory PRO — naprawa formuł wyłączona',
    'Automatyczna naprawa formuł jest wyłączona w wersji 4.3.7 SAFE MODE. ' +
      'Nie wykonano żadnego zapisu.',
    ui.ButtonSet.OK
  );
  return { success: false, disabled: true, changedCells: 0 };
}
