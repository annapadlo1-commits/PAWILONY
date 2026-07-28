/**
 * Inventory PRO 4.3.4 — bezpieczny, selektywny zapis PAWILONÓW.
 * Odczyt buforów służy wyłącznie do obliczania sum, a zapis obejmuje tylko
 * konkretne komórki wejściowe. Kolumny formuł są chronione kontraktem typu.
 */

function saveImportItems(items) {
  const submission = prepareImportSubmission_(items);
  const activeItems = submission.items;
  const excludedCount = submission.excludedCount;

  const lock = LockService.getDocumentLock();
  const startedAt = Date.now();
  const importId = createUniqueId_('IMP');
  let inventorySheet = null;
  let writePlan = [];
  let inventoryWritten = false;

  try {
    lock.waitLock(30000);
    const sheet = getSheetByConfiguredName_(CONFIG.SHEETS.INVENTORY);
    inventorySheet = sheet;
    if (!sheet) throw new Error('Nie znaleziono arkusza: ' + CONFIG.SHEETS.INVENTORY);

    const containsNewProduct = activeItems.some(item =>
      item && item.include && String(item.status || '').toUpperCase() === 'NEW_PRODUCT'
    );
    if (containsNewProduct) {
      SpreadsheetApp.flush();
      invalidateProductCatalogCache_();
    }

    const runtimeContext = buildRuntimeContext_();
    const productIndex = runtimeContext.productIndex;
    const qualitySettings = loadQualitySettings_();
    const lastRow = sheet.getLastRow();
    const usedColumns = collectUsedTargetColumns_(activeItems, productIndex);
    const columnBuffers = loadColumnBuffers_(sheet, usedColumns, lastRow);

    const results = [];
    let savedCount = 0;
    let skippedCount = 0;

    activeItems.forEach(item => {
      const result = prepareSingleImportWrite_(
        item, productIndex, columnBuffers, qualitySettings
      );
      result.importId = importId;
      result.location = item.location || '';
      results.push(result);
      if (result.saved) savedCount++;
      else skippedCount++;
    });

    annotateSavedDuplicateResults_(results, qualitySettings);
    writePlan = buildSparseWritePlan_(results);
    // Flaga jest ustawiana przed pierwszym setValue. Dzięki temu błąd
    // pojedynczego zapisu w połowie planu nadal uruchomi selektywny rollback.
    inventoryWritten = writePlan.length > 0;
    writeSparseWritePlan_(sheet, writePlan);

    SpreadsheetApp.flush();
    const statusColors = refreshInventoryStatusColors_(sheet);
    appendImportHistory_(importId, results, sheet.getName());

    let learnedAliasesCount = 0;
    let protectedFamilyAliasesCount = 0;
    try {
      const aliasSuggestions = collectAliasSuggestions_(activeItems);
      protectedFamilyAliasesCount = countProtectedAliasSuggestions_(
        aliasSuggestions,
        runtimeContext
      );
      learnedAliasesCount = saveAliasesBatch_(aliasSuggestions);
    } catch (aliasError) {
      logWarning(
        'InventoryWriter', 'saveImportItems',
        'Import zapisano, ale nie udało się zapisać aliasów.',
        { message: normalizeError_(aliasError).message }
      );
    }

    const response = {
      success: true,
      importId: importId,
      savedCount: savedCount,
      skippedCount: skippedCount,
      excludedCount: excludedCount,
      learnedAliasesCount: learnedAliasesCount,
      protectedFamilyAliasesCount: protectedFamilyAliasesCount,
      duplicateGroupCount: countSavedDuplicateGroups_(results),
      warningCount: results.filter(result => result.qualityWarning).length,
      coloredProductsCount: statusColors.colored,
      formulaErrorColorCount: statusColors.errors,
      results: results,
      durationMs: Date.now() - startedAt
    };

    logInfo('InventoryWriter', 'saveImportItems', 'Import zapisany', {
      importId: importId,
      savedCount: savedCount,
      skippedCount: skippedCount,
      changedCells: writePlan.length
    }, response.durationMs);

    return response;
  } catch (error) {
    if (inventoryWritten && inventorySheet && writePlan.length) {
      try {
        rollbackSparseWritePlan_(inventorySheet, writePlan);
        SpreadsheetApp.flush();
      } catch (rollbackError) {
        logError('InventoryWriter', 'saveImportItems.rollback', rollbackError, {
          importId: importId
        }, 0);
      }
    }
    logError('InventoryWriter', 'saveImportItems', error, { importId: importId }, Date.now() - startedAt);
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function prepareImportSubmission_(items) {
  if (!Array.isArray(items) || !items.length) {
    throw new Error('Brak pozycji do zapisania.');
  }
  const activeItems = items.filter(item => item && item.include === true);
  if (!activeItems.length) {
    throw new Error('Nie wybrano żadnej pozycji do importu.');
  }
  const invalid = activeItems.filter(item =>
    !String(item.selectedProduct || '').trim() ||
    !Number.isFinite(Number(item.value)) ||
    ['PARSE_ERROR', 'NOT_FOUND', 'ERROR', 'AMBIGUOUS'].includes(
      String(item.status || '').toUpperCase()
    ) ||
    String(item.qualityLevel || '').toUpperCase() === 'ERROR' ||
    (
      String(item.productType || '').toUpperCase() === 'LOCATION' &&
      !String(item.location || '').trim()
    )
  );
  if (invalid.length) {
    throw new Error(
      'Pozostały aktywne błędy w zaznaczonych pozycjach: ' + invalid.length + '.'
    );
  }
  return {
    items: activeItems,
    excludedCount: items.length - activeItems.length
  };
}

function buildSparseWritePlan_(results) {
  const byCell = {};
  (results || []).forEach(result => {
    if (!result || !result.saved || !result.row || !result.column) return;
    const column = String(result.column).toUpperCase();
    const key = column + String(result.row);
    if (!byCell[key]) {
      byCell[key] = {
        a1: key,
        row: Number(result.row),
        column: column,
        previousValue: result.previousValue,
        newValue: result.newValue,
        product: result.product || '',
        productType: result.productType || ''
      };
    } else {
      byCell[key].newValue = result.newValue;
    }
  });

  return Object.keys(byCell).map(key => byCell[key]).sort((a, b) =>
    a.column.localeCompare(b.column) || a.row - b.row
  );
}

function writeSparseWritePlan_(sheet, plan) {
  const prepared = (plan || []).map(change => {
    validateSparseWriteTarget_(change);
    const range = sheet.getRange(change.a1);
    const formula = range.getFormula();
    if (formula) {
      throw new Error('Zablokowano zapis do komórki z formułą: ' + change.a1 + '.');
    }
    const liveValue = range.getValue();
    if (!inventoryCellValuesEqual_(liveValue, change.previousValue)) {
      throw new Error(
        'Komórka ' + change.a1 + ' została zmieniona równolegle. ' +
        'Import przerwano bez nadpisywania ręcznej zmiany.'
      );
    }
    change.previousNumberFormat = range.getNumberFormat();
    change.appliedNumberFormat = Number.isInteger(Number(change.newValue)) ? '0' : 'General';
    return { range: range, change: change };
  });

  prepared.forEach(item => {
    item.range
      .setValue(item.change.newValue)
      .setNumberFormat(item.change.appliedNumberFormat);
  });
}

function validateSparseWriteTarget_(change) {
  const productContext = {
    name: change && change.product || '',
    product: change && change.product || '',
    type: change && change.productType || ''
  };
  const requestedColumn = normalizeColumnLetter_(change && change.column);
  const safeColumn = assertSafeInventoryTargetColumn_(
    productContext,
    requestedColumn
  );
  if (safeColumn !== requestedColumn) {
    throw new Error(
      'Niespójna kolumna zapisu ' + requestedColumn +
      ' dla produktu „' + (productContext.name || 'produkt') + '”.'
    );
  }
  return safeColumn;
}

function rollbackSparseWritePlan_(sheet, plan) {
  (plan || []).slice().reverse().forEach(change => {
    const range = sheet.getRange(change.a1);
    const liveValue = range.getValue();
    if (inventoryCellValuesEqual_(liveValue, change.previousValue)) {
      // Ta pozycja nie została jeszcze zapisana przed wystąpieniem błędu.
      return;
    }
    if (inventoryCellValuesEqual_(liveValue, change.newValue)) {
      if (change.previousValue === '' || change.previousValue === null || change.previousValue === undefined) {
        range.clearContent();
      } else {
        range.setValue(change.previousValue);
      }
      if (change.previousNumberFormat) range.setNumberFormat(change.previousNumberFormat);
      return;
    }
    logWarning(
      'InventoryWriter', 'rollbackSparseWritePlan_',
      'Nie cofnięto komórki zmienionej po imporcie.',
      { cell: change.a1, expected: change.newValue, actual: liveValue }
    );
  });
}

function inventoryCellValuesEqual_(left, right) {
  const leftBlank = left === '' || left === null || left === undefined;
  const rightBlank = right === '' || right === null || right === undefined;
  if (leftBlank || rightBlank) return leftBlank && rightBlank;
  if (typeof left === 'number' || typeof right === 'number') {
    const a = Number(left);
    const b = Number(right);
    return Number.isFinite(a) && Number.isFinite(b) && Math.abs(a - b) < 0.000000001;
  }
  return String(left) === String(right);
}

function collectUsedTargetColumns_(
  items,
  productIndex
) {
  const columns = {};

  items.forEach(item => {
    if (!item.include || !item.selectedProduct) {
      return;
    }

    const product = productIndex[
      normalizeText(item.selectedProduct)
    ];

    if (!product) {
      return;
    }

    const value = Number(item.value);

    if (!Number.isFinite(value)) {
      return;
    }

    const column = resolveTargetColumn_(
      product,
      value,
      item.location
    );

    if (column) {
      const safeColumn = assertSafeInventoryTargetColumn_(product, column);
      columns[safeColumn] = true;
    }
  });

  return Object.keys(columns);
}

function loadColumnBuffers_(
  sheet,
  columns,
  lastRow
) {
  const buffers = {};

  columns.forEach(column => {
    buffers[column] = sheet
      .getRange(column + '1:' + column + lastRow)
      .getValues();
  });

  return buffers;
}

function prepareSingleImportWrite_(
  item,
  productIndex,
  columnBuffers,
  qualitySettings
) {
  const originalInput = String(
    item.originalInput || ''
  ).trim();

  if (!item.include) {
    return createWriteResult_(
      originalInput,
      false,
      'Pominieto przez uzytkownika'
    );
  }

  const productName = String(
    item.selectedProduct || ''
  ).trim();

  if (!productName) {
    return createWriteResult_(
      originalInput,
      false,
      'Nie wybrano produktu'
    );
  }

  const product = productIndex[
    normalizeText(productName)
  ];

  if (!product) {
    return createWriteResult_(
      originalInput,
      false,
      'Produktu nie ma w aktywnym katalogu'
    );
  }

  if (!product.inventoryRow) {
    return createWriteResult_(
      originalInput,
      false,
      'Nie znaleziono wiersza produktu'
    );
  }

  const value = Number(item.value);

  if (!Number.isFinite(value)) {
    return createWriteResult_(
      originalInput,
      false,
      'Nieprawidlowa wartosc'
    );
  }

  const quality = evaluateImportQuality_(
    product,
    value,
    item.location || '',
    qualitySettings
  );

  if (quality.blocking) {
    const blocked = createWriteResult_(
      originalInput,
      false,
      quality.message || 'Wartosc zablokowana przez ustawienia'
    );
    blocked.product = product.name;
    blocked.addedValue = value;
    blocked.qualityWarning = true;
    blocked.qualityLevel = quality.level;
    blocked.qualityFlags = quality.flags || [];
    return blocked;
  }

  const targetColumn = resolveTargetColumn_(
    product,
    value,
    item.location
  );
  const safeTargetColumn = targetColumn
    ? assertSafeInventoryTargetColumn_(product, targetColumn)
    : '';

  if (
    !safeTargetColumn ||
    !columnBuffers[safeTargetColumn]
  ) {
    return createWriteResult_(
      originalInput,
      false,
      'Nie ustalono kolumny docelowej'
    );
  }

  const bufferIndex = product.inventoryRow - 1;
  const previousRaw =
    columnBuffers[safeTargetColumn][bufferIndex][0];

  const previousNumericValue =
    typeof previousRaw === 'number' &&
    Number.isFinite(previousRaw)
      ? previousRaw
      : 0;

  const newValue = previousNumericValue + value;

  columnBuffers[safeTargetColumn][bufferIndex][0] =
    newValue;

  return {
    originalInput: originalInput,
    product: product.name,
    saved: true,
    row: product.inventoryRow,
    column: safeTargetColumn,
    productType: String(product.type || '').toUpperCase(),
    // Zachowujemy dokładny stan komórki. Pusta komórka musi po cofnięciu
    // ponownie być pusta, a nie zawierać techniczne zero użyte do obliczeń.
    previousValue: previousRaw === null || previousRaw === undefined ? '' : previousRaw,
    addedValue: value,
    newValue: newValue,
    qualityWarning: quality.warning,
    qualityLevel: quality.level,
    qualityFlags: quality.flags || [],
    duplicateCount: 1,
    duplicateValues: [value],
    duplicateTotal: value,
    duplicateWarning: false,
    message:
      'Zapisano do ' +
      safeTargetColumn +
      product.inventoryRow +
      (quality.warning ? ' | ' + quality.message : '')
  };
}

function resolveTargetColumn_(product, value, location) {
  const directFinal = getDirectFinalInventoryColumn_(product);
  if (directFinal) return directFinal;
  const type = String(product.type || '').toUpperCase();
  const runtimeColumns = normalizeRuntimeProductColumns_(type, product.columns, product);
  const isWholeNumber = Number.isInteger(value);

  if (type === CONFIG.PRODUCT_TYPES.LOCATION) {
    const area = resolveLocationArea_(location);
    if (!area || !area.columnKey) return '';
    return String(runtimeColumns[area.columnKey] || '').toUpperCase();
  }

  const quantityColumn = String(runtimeColumns.quantity || '').toUpperCase();
  const weightColumn = String(runtimeColumns.weight || '').toUpperCase();
  const category = normalizeText(product.category || '');

  // KAWA jest liczona wagowo również przy liczbie całkowitej.
  if (category === 'kawa' && weightColumn) return weightColumn;
  if (weightColumn && !quantityColumn) return weightColumn;
  if (quantityColumn && !weightColumn) return quantityColumn;
  return isWholeNumber ? quantityColumn : weightColumn;
}

function createWriteResult_(
  originalInput,
  saved,
  message
) {
  return {
    originalInput: originalInput,
    product: '',
    productType: '',
    saved: saved,
    row: null,
    column: '',
    previousValue: null,
    addedValue: null,
    newValue: null,
    qualityWarning: false,
    qualityLevel: saved ? 'OK' : 'ERROR',
    qualityFlags: [],
    duplicateCount: 1,
    duplicateValues: [],
    duplicateTotal: null,
    duplicateWarning: false,
    message: message
  };
}


function annotateSavedDuplicateResults_(results, qualitySettings) {
  const groups = {};
  const threshold = qualitySettings && qualitySettings.duplicateWarningCount
    ? qualitySettings.duplicateWarningCount
    : 2;

  results.forEach((result, index) => {
    if (!result.saved || !result.product || !result.row || !result.column) return;
    const key = [normalizeText(result.product), result.row, result.column].join('|');
    if (!groups[key]) groups[key] = [];
    groups[key].push(index);
  });

  Object.keys(groups).forEach(key => {
    const indexes = groups[key];
    const values = indexes.map(index => Number(results[index].addedValue));
    const total = values.reduce((sum, value) => sum + value, 0);
    const isDuplicate = indexes.length >= threshold;

    indexes.forEach(index => {
      const result = results[index];
      result.duplicateCount = indexes.length;
      result.duplicateValues = values;
      result.duplicateTotal = total;
      result.duplicateWarning = isDuplicate;

      if (isDuplicate) {
        const duplicateMessage = 'DUPLIKAT: zsumowano ' + indexes.length +
          ' wpisy (' + values.join(' + ') + ' = ' + total + ')';
        result.message = result.message
          ? result.message + ' | ' + duplicateMessage
          : duplicateMessage;
      }
    });
  });
}

function countSavedDuplicateGroups_(results) {
  const groups = {};
  results.forEach(result => {
    if (!result.duplicateWarning || !result.product) return;
    groups[[normalizeText(result.product), result.row, result.column].join('|')] = true;
  });
  return Object.keys(groups).length;
}
