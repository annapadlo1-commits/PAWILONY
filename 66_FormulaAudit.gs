/**
 * Inventory PRO 4.3.4 — audyt kontraktu formuł PAWILONÓW.
 *
 * Audytuje wyłącznie fizyczne wiersze produktów wykryte przez
 * scanInventoryProducts_(). Rozróżnia bezpiecznie spłaszczone wyniki od
 * konfliktów, których nie wolno automatycznie nadpisywać.
 */

function getInventoryFormulaContract_(product, packagingProfileOverride) {
  if (isDirectFinalInventoryProduct_(product)) return [];
  const row = Number(product && product.inventoryRow) || 0;
  const type = String(product && product.type || '').trim().toUpperCase();
  const layout = getConfiguredInventoryLayout_(type);
  if (!row || !layout) return [];

  if (type === CONFIG.PRODUCT_TYPES.LOCATION) {
    const sources = [layout.warehouse, layout.darkroom, layout.fridges].filter(Boolean);
    return validateInventoryFormulaContract_(product, layout, [buildInventoryFormulaContractEntry_(
      layout.finalTotal,
      'SUM',
      sources,
      row
    )]);
  }

  if (type === CONFIG.PRODUCT_TYPES.KEG) {
    return validateInventoryFormulaContract_(product, layout, [
      buildPackagingAwareFormulaContractEntry_(product, layout, row, packagingProfileOverride),
      buildInventoryFormulaContractEntry_(
        layout.fullUnitsVolume,
        'PRODUCT',
        [layout.fullUnits, layout.unitCapacity],
        row
      ),
      buildInventoryFormulaContractEntry_(
        layout.finalTotal,
        'SUM',
        [layout.openNet, layout.fullUnitsVolume],
        row
      )
    ]);
  }

  return validateInventoryFormulaContract_(product, layout, [
    buildPackagingAwareFormulaContractEntry_(product, layout, row, packagingProfileOverride),
    buildInventoryFormulaContractEntry_(
      layout.fullUnitsVolume,
      'PRODUCT',
      [layout.fullUnits, layout.unitCapacity],
      row
    ),
    buildInventoryFormulaContractEntry_(
      layout.finalTotal,
      'SUM',
      [layout.openNet, layout.prepNet, layout.fullUnitsVolume].filter(Boolean),
      row
    )
  ]);
}

/**
 * Fail-closed guard. A formula path may not write anything unless the
 * semantic targets, configured formula columns and generated contract agree.
 */
function validateInventoryFormulaContract_(product, layout, contracts) {
  const type = String(product && product.type || '').trim().toUpperCase();
  const expected = type === CONFIG.PRODUCT_TYPES.LOCATION
    ? [layout.finalTotal]
    : [layout.openNet, layout.fullUnitsVolume, layout.finalTotal];
  const normalizeList = list => (list || []).map(normalizeColumnLetter_).filter(Boolean);
  const wanted = normalizeList(expected);
  const configured = normalizeList(layout.formulaColumns);
  const generated = normalizeList((contracts || []).map(item => item.column));
  const same = (left, right) =>
    left.length === right.length && left.every((value, index) => value === right[index]);

  if (!same(wanted, configured) || !same(wanted, generated)) {
    throw new Error(
      'Zatrzymano zapis formuł dla „' + String(product && product.name || '') +
      '”. Kontrakt arkusza nie jest zgodny z CONFIG.INVENTORY_LAYOUTS (' +
      type + ': oczekiwane ' + wanted.join(', ') + '; konfiguracja ' +
      configured.join(', ') + '; wygenerowane ' + generated.join(', ') + ').'
    );
  }

  (contracts || []).forEach(entry => {
    const column = normalizeColumnLetter_(entry.column);
    if (!column || Number(entry.columnNumber) !== inventoryColumnLetterToNumber_(column) ||
        String(entry.formula || '').charAt(0) !== '=' ||
        String(entry.r1c1 || '').charAt(0) !== '=') {
      throw new Error('Zatrzymano zapis: nieprawidłowy kontrakt formuły w kolumnie ' + column + '.');
    }
  });
  return contracts;
}

function buildPackagingAwareFormulaContractEntry_(product, layout, row, packagingProfileOverride) {
  const target = normalizeColumnLetter_(layout.openNet);
  const spec = buildPackagingAwareOpenNetFormulaSpec_(
    product, layout, row, target, packagingProfileOverride
  );
  return {
    column: target,
    columnNumber: inventoryColumnLetterToNumber_(target),
    operation: spec.operation,
    sourceColumns: [layout.grossWeight, layout.emptyContainerWeight].filter(Boolean),
    formula: spec.formula,
    r1c1: spec.r1c1
  };
}

function buildInventoryFormulaContractEntry_(targetColumn, operation, sourceColumns, row) {
  const target = normalizeColumnLetter_(targetColumn);
  const sources = (sourceColumns || []).map(normalizeColumnLetter_).filter(Boolean);
  const a1Sources = sources.map(column => column + row);
  let formula = '';

  if (operation === 'DIFFERENCE') {
    // Brak wagi brutto oznacza brak otwartego opakowania, a nie wagę 0.
    // Zapis bez funkcji jest odporny na lokalny separator formuł.
    formula = '=(' + a1Sources[0] + '-' + a1Sources[1] + ')*(' + a1Sources[0] + '<>"")';
  } else if (operation === 'PRODUCT') {
    formula = '=' + a1Sources[0] + '*' + a1Sources[1];
  } else {
    formula = buildSeparatorSafeSumFormulaA1_(sources, row);
  }

  const targetNumber = inventoryColumnLetterToNumber_(target);
  const r1c1Sources = sources.map(column =>
    buildRelativeR1C1Reference_(targetNumber, inventoryColumnLetterToNumber_(column))
  );
  let r1c1 = '';
  if (operation === 'DIFFERENCE') {
    r1c1 = '=(' + r1c1Sources[0] + '-' + r1c1Sources[1] + ')*(' + r1c1Sources[0] + '<>"")';
  } else if (operation === 'PRODUCT') {
    r1c1 = '=' + r1c1Sources[0] + '*' + r1c1Sources[1];
  } else {
    r1c1 = buildSeparatorSafeSumFormulaR1C1_(targetNumber, sources);
  }

  return {
    column: target,
    columnNumber: targetNumber,
    operation: operation,
    sourceColumns: sources,
    formula: formula,
    r1c1: r1c1
  };
}


function buildSeparatorSafeSumFormulaA1_(sources, row) {
  const cols = (sources || []).map(normalizeColumnLetter_).filter(Boolean);
  if (cols.length === 3 && cols.join('|') === 'E|G|J') {
    return '=SUM(E' + row + ':G' + row + ')+J' + row;
  }
  const numbers = cols.map(inventoryColumnLetterToNumber_);
  const contiguous = numbers.every((value, index) => index === 0 || value === numbers[index - 1] + 1);
  if (contiguous && cols.length > 1) {
    return '=SUM(' + cols[0] + row + ':' + cols[cols.length - 1] + row + ')';
  }
  return '=' + cols.map(column => column + row).join('+');
}

function buildSeparatorSafeSumFormulaR1C1_(targetColumnNumber, sources) {
  const cols = (sources || []).map(normalizeColumnLetter_).filter(Boolean);
  if (cols.length === 3 && cols.join('|') === 'E|G|J') {
    return '=SUM(RC[-6]:RC[-4])+RC[-1]';
  }
  const numbers = cols.map(inventoryColumnLetterToNumber_);
  const contiguous = numbers.every((value, index) => index === 0 || value === numbers[index - 1] + 1);
  const refs = numbers.map(number => buildRelativeR1C1Reference_(targetColumnNumber, number));
  if (contiguous && refs.length > 1) {
    return '=SUM(' + refs[0] + ':' + refs[refs.length - 1] + ')';
  }
  return '=' + refs.join('+');
}

function buildRelativeR1C1Reference_(targetColumnNumber, sourceColumnNumber) {
  const offset = Number(sourceColumnNumber) - Number(targetColumnNumber);
  if (!offset) return 'RC';
  return 'RC[' + offset + ']';
}

function getCanonicalInventoryFormula_(product, column) {
  const wanted = normalizeColumnLetter_(column);
  const entry = getInventoryFormulaContract_(product).find(item => item.column === wanted);
  return entry ? entry.formula : '';
}

function normalizeInventoryFormula_(formula) {
  return String(formula || '')
    .replace(/\s+/g, '')
    .replace(/;/g, ',')
    .toUpperCase();
}

function stripInventoryFormulaOuterParentheses_(formula) {
  let normalized = normalizeInventoryFormula_(formula);
  while (normalized.indexOf('=(') === 0 && normalized.charAt(normalized.length - 1) === ')') {
    normalized = '=' + normalized.slice(2, -1);
  }
  return normalized;
}

function isLegacyEquivalentInventoryFormula_(formula, contract, row) {
  if (!formula || !contract) return false;
  const sources = (contract.sourceColumns || []).filter(Boolean);
  if (sources.length < 2) return false;
  const legacy = '=' + sources
    .map(column => column + row)
    .join('+');
  return stripInventoryFormulaOuterParentheses_(formula) ===
    stripInventoryFormulaOuterParentheses_(legacy);
}

function isSpreadsheetFormulaError_(displayValue) {
  return /^#(?:REF!|VALUE!|DIV\/0!|N\/A|NAME\?|NUM!|NULL!|ERROR!)$/i
    .test(String(displayValue || '').trim());
}

function isInventoryAuditBlank_(value) {
  return value === '' || value === null || value === undefined;
}

function inventoryAuditNumericValue_(value, ignoreText) {
  if (isInventoryAuditBlank_(value)) return 0;
  if (typeof value === 'number') return Number.isFinite(value) ? value : NaN;
  const normalized = String(value).trim().replace(',', '.');
  if (!normalized) return 0;
  const parsed = Number(normalized);
  if (Number.isFinite(parsed)) return parsed;
  return ignoreText ? 0 : NaN;
}

function calculateCanonicalInventoryValue_(contract, rowValues) {
  const values = (contract.sourceColumns || []).map(column => {
    const index = inventoryColumnLetterToNumber_(column) - 1;
    return rowValues && index >= 0 ? rowValues[index] : '';
  });

  if (contract.operation === 'DIFFERENCE') {
    if (isInventoryAuditBlank_(values[0])) return 0;
    const left = inventoryAuditNumericValue_(values[0], false);
    const right = inventoryAuditNumericValue_(values[1], false);
    return Number.isFinite(left) && Number.isFinite(right) ? left - right : NaN;
  }
  if (contract.operation === 'PRODUCT') {
    const left = inventoryAuditNumericValue_(values[0], false);
    const right = inventoryAuditNumericValue_(values[1], false);
    return Number.isFinite(left) && Number.isFinite(right) ? left * right : NaN;
  }
  return values.reduce((sum, value) => sum + inventoryAuditNumericValue_(value, true), 0);
}

function inventoryAuditValuesEqual_(left, right) {
  const tolerance = Number(CONFIG.FORMULA_POLICY && CONFIG.FORMULA_POLICY.NUMERIC_TOLERANCE) || 0.000000001;
  const a = Number(left);
  const b = Number(right);
  return Number.isFinite(a) && Number.isFinite(b) && Math.abs(a - b) <= tolerance;
}

function buildInventoryFormulaIssue_(product, contract, actualFormula, currentValue, displayedValue, expectedValue) {
  const row = Number(product.inventoryRow);
  return {
    cell: contract.column + row,
    row: row,
    column: contract.column,
    columnNumber: contract.columnNumber,
    product: product.name,
    category: product.category,
    type: product.type,
    expectedFormula: contract.formula,
    expectedR1C1: contract.r1c1,
    expectedValue: Number.isFinite(expectedValue) ? expectedValue : '',
    actualFormula: actualFormula || '',
    currentValue: currentValue,
    displayedValue: displayedValue,
    operation: contract.operation,
    sourceColumns: contract.sourceColumns.slice()
  };
}

function buildInventoryFormulaAudit_(sheet, products) {
  const physicalProducts = Array.isArray(products) ? products : [];
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const width = Math.max(getInventoryLayoutMaxColumn_(), 1);
  const range = sheet.getRange(1, 1, lastRow, width);
  const formulas = range.getFormulas();
  const values = range.getValues();
  const displayValues = range.getDisplayValues();

  const missing = [];
  const flattened = [];
  const conflicts = [];
  const legacy = [];
  const invalid = [];
  const calculationErrors = [];
  let expectedFormulaCells = 0;
  let presentFormulaCells = 0;

  physicalProducts.forEach(product => {
    const row = Number(product.inventoryRow);
    const rowValues = values[row - 1] || [];
    getInventoryFormulaContract_(product).forEach(contract => {
      expectedFormulaCells++;
      const columnIndex = contract.columnNumber - 1;
      const actual = formulas[row - 1] && formulas[row - 1][columnIndex] || '';
      const currentValue = rowValues[columnIndex];
      const displayed = displayValues[row - 1] && displayValues[row - 1][columnIndex] || '';
      const expectedValue = calculateCanonicalInventoryValue_(contract, rowValues);
      const issue = buildInventoryFormulaIssue_(
        product, contract, actual, currentValue, displayed, expectedValue
      );

      if (actual) {
        presentFormulaCells++;
        if (normalizeInventoryFormula_(actual) !== normalizeInventoryFormula_(contract.formula)) {
          if (isLegacyEquivalentInventoryFormula_(actual, contract, row)) legacy.push(issue);
          else invalid.push(issue);
        }
        if (isSpreadsheetFormulaError_(displayed)) {
          calculationErrors.push(issue);
        }
        return;
      }

      if (isInventoryAuditBlank_(currentValue)) {
        missing.push(issue);
        return;
      }

      const currentNumeric = inventoryAuditNumericValue_(currentValue, false);
      if (Number.isFinite(expectedValue) && Number.isFinite(currentNumeric) &&
          inventoryAuditValuesEqual_(currentNumeric, expectedValue)) {
        flattened.push(issue);
      } else {
        conflicts.push(issue);
      }
    });
  });

  const repairableKeys = {};
  missing.concat(flattened, legacy, invalid, calculationErrors).forEach(issue => {
    repairableKeys[issue.cell] = true;
  });

  return {
    safe: missing.length === 0 && flattened.length === 0 && conflicts.length === 0 &&
      invalid.length === 0 && calculationErrors.length === 0,
    operationallySafe: missing.length === 0 && flattened.length === 0 && conflicts.length === 0 &&
      invalid.length === 0 && calculationErrors.length === 0,
    repairableWithoutConflicts: conflicts.length === 0,
    hasBlockingConflicts: conflicts.length > 0,
    sheetName: sheet.getName(),
    products: physicalProducts.length,
    expectedFormulaCells: expectedFormulaCells,
    presentFormulaCells: presentFormulaCells,
    missingFormulaCells: missing.length,
    flattenedFormulaCells: flattened.length,
    conflictFormulaCells: conflicts.length,
    legacyFormulaCells: legacy.length,
    invalidFormulaCells: invalid.length,
    errorFormulaCells: calculationErrors.length,
    repairableFormulaCells: Object.keys(repairableKeys).length,
    missing: missing,
    flattened: flattened,
    conflicts: conflicts,
    legacy: legacy,
    invalid: invalid,
    calculationErrors: calculationErrors
  };
}

function buildInventoryFormulaCellIndex_(products) {
  const index = {};
  (products || []).forEach(product => {
    getInventoryFormulaContract_(product).forEach(contract => {
      index[contract.column + product.inventoryRow] = {
        product: product.name,
        type: product.type,
        category: product.category,
        formula: contract.formula
      };
    });
  });
  return index;
}

function auditInventoryFormulaCoverage_(options) {
  const settings = options || {};
  const sheet = settings.sheet || getSheetByConfiguredName_(CONFIG.SHEETS.INVENTORY);
  if (!sheet) throw new Error('Nie znaleziono arkusza: ' + CONFIG.SHEETS.INVENTORY);
  const products = settings.products || scanInventoryProducts_();
  const audit = buildInventoryFormulaAudit_(sheet, products);

  logInfo(
    'FormulaAudit',
    'auditInventoryFormulaCoverage_',
    audit.safe ? 'Kontrakt formuł PAWILONÓW jest kompletny.' : 'Wykryto problemy z formułami PAWILONÓW.',
    {
      sheet: audit.sheetName,
      expected: audit.expectedFormulaCells,
      present: audit.presentFormulaCells,
      missing: audit.missingFormulaCells,
      flattened: audit.flattenedFormulaCells,
      conflicts: audit.conflictFormulaCells,
      legacy: audit.legacyFormulaCells,
      invalid: audit.invalidFormulaCells,
      calculationErrors: audit.errorFormulaCells
    }
  );
  return audit;
}

function formatInventoryFormulaAudit_(audit) {
  const lines = [
    'Arkusz: ' + audit.sheetName,
    'Produkty: ' + audit.products,
    'Oczekiwane formuły: ' + audit.expectedFormulaCells,
    'Obecne formuły: ' + audit.presentFormulaCells,
    'Puste komórki bez formuły: ' + audit.missingFormulaCells,
    'Spłaszczone, zgodne wyniki: ' + audit.flattenedFormulaCells,
    'Konflikty wymagające decyzji: ' + audit.conflictFormulaCells,
    'Poprawne formuły starszego typu (+): ' + audit.legacyFormulaCells,
    'Nieprawidłowe formuły: ' + audit.invalidFormulaCells,
    'Błędy obliczeń: ' + audit.errorFormulaCells,
    '',
    audit.safe ? 'Status: OK' :
      (audit.hasBlockingConflicts ? 'Status: BLOKADA — wymagane rozstrzygnięcie konfliktów' : 'Status: MOŻLIWA BEZPIECZNA NAPRAWA')
  ];

  const sample = audit.conflicts.concat(
    audit.missing, audit.flattened, audit.invalid, audit.calculationErrors
  ).slice(0, 15);
  if (sample.length) {
    lines.push('', 'Przykładowe komórki:');
    sample.forEach(issue => lines.push(
      '- ' + issue.cell + ' — ' + issue.product +
      (issue.expectedValue !== '' ? ' (oczekiwany wynik: ' + issue.expectedValue + ')' : '')
    ));
  }
  return lines.join('\n');
}

function auditInventoryFormulaCoverageWithDialog() {
  return runSafely_(
    'FormulaAudit',
    'auditInventoryFormulaCoverageWithDialog',
    function() {
      const audit = auditInventoryFormulaCoverage_();
      SpreadsheetApp.getUi().alert(
        'Inventory PRO — audyt formuł PAWILONÓW',
        formatInventoryFormulaAudit_(audit),
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return audit;
    },
    'Nie udało się przeprowadzić audytu formuł PAWILONÓW.'
  );
}
