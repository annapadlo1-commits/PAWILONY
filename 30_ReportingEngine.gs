/**
 * Inventory PRO Enterprise v2.10.2
 * Centralny, tylko-do-odczytu silnik raportowania.
 *
 * Źródło prawdy dla stanu końcowego:
 * - NORMAL   -> K
 * - KEG      -> J
 * - LOCATION -> E
 *
 * Moduł nie zapisuje niczego w arkuszu źródłowym.
 */

function generateInventoryReport(sourceSheetName) {
  return generateInventoryReport_(sourceSheetName);
}

function generateInventoryReport_(sourceSheetName) {
  const started = Date.now();
  const sheetName = resolveReportingSourceSheetName_(sourceSheetName);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error('Nie znaleziono arkusza raportowego: ' + sheetName + '.');

  const lastRow = Math.max(sheet.getLastRow(), 1);
  const lastColumn = Math.max(sheet.getLastColumn(), 11);
  const values = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
  const displayValues = sheet.getRange(1, 1, lastRow, lastColumn).getDisplayValues();
  const categoryByRow = buildStrictInventoryCategoryMapFromSheet_(sheet, displayValues);
  let catalog = buildProductCatalog();
  if (!catalog.length) {
    invalidateProductCatalogCache_();
    catalog = buildProductCatalogUncached_();
  }
  if (!catalog.length && scanInventoryProducts_().length) {
    throw new Error(
      'Katalog raportowy jest pusty mimo obecności produktów w INWENTURZE. ' +
      'Odśwież konfigurację SŁOWNIKA przed eksportem.'
    );
  }
  const settings = loadQualitySettings_();

  const items = [];
  const warnings = [];
  const validationIssues = [];

  catalog.forEach(product => {
    const row = Number(product.inventoryRow);
    if (!row || row > lastRow) {
      validationIssues.push(createReportIssue_('ERROR', product.name, 'Brak poprawnego wiersza INWENTURA.'));
      return;
    }

    const type = String(product.type || '').trim().toUpperCase();
    if (!getConfiguredInventoryLayout_(type)) {
      validationIssues.push(createReportIssue_('ERROR', product.name, 'Nieobsługiwany typ: ' + (type || 'BRAK') + '.'));
      return;
    }

    const physicalCategory = categoryByRow[row] || '';
    if (!physicalCategory) {
      validationIssues.push(createReportIssue_('ERROR', product.name, 'Nie znaleziono fizycznej kategorii w arkuszu źródłowym.'));
      return;
    }

    const item = readInventorySummaryItemFromMatrix_(values, product, physicalCategory);
    item.flags = [];
    item.reviewReason = classifyFinalReviewRequirement_(item);
    applySummaryWarnings_(item, settings);
    if (!item.hasValue) item.flags.push('BRAK WARTOŚCI');
    if (item.reviewReason === 'FULL_UNITS_ONLY') item.flags.push('TYLKO PEŁNE SZTUKI');
    if (item.reviewReason === 'UNIT_CAPACITY_MISSING') {
      item.flags.push('BRAK POJEMNOŚCI PEŁNEJ BUTELKI / KEGA');
      validationIssues.push(createReportIssue_(
        'ERROR',
        product.name,
        'Wpisano pełne sztuki, ale brakuje pojemności jednej pełnej butelki / kega.'
      ));
    }
    if (item.reviewReason === 'PACKAGING_DATA_MISSING') item.flags.push('BRAK DANYCH OPAKOWANIA');
    if (item.reviewReason === 'REFERENCE_RESET_REQUIRED') item.flags.push('NOWY PUNKT REFERENCYJNY');
    item.flags = Array.from(new Set(item.flags));
    if (String(item.finalTotal || '').trim().charAt(0) === '#') {
      item.flags.push('BŁĄD FORMUŁY');
      validationIssues.push(createReportIssue_(
        'ERROR', product.name,
        'Stan końcowy zawiera błąd formuły. Uruchom naprawę profili opakowań.'
      ));
    }

    if (normalizeBusinessCategory_(product.category) !== physicalCategory) {
      validationIssues.push(createReportIssue_(
        'WARNING',
        product.name,
        'Kategoria w SŁOWNIKU „' + String(product.category || 'BRAK') + '” różni się od fizycznej kategorii „' + physicalCategory + '”. Raport użył kategorii z arkusza.'
      ));
    }

    if (item.flags.length) {
      warnings.push({
        key: item.key,
        product: item.product,
        category: item.category,
        type: item.type,
        finalTotal: item.finalTotal,
        unit: item.unit,
        flags: item.flags.slice(),
        values: buildLegacyReviewValues_(item)
      });
    }
    items.push(item);
  });

  const report = createInventoryReportModel_({
    sourceSheetName: sheetName,
    generatedAt: new Date(),
    items: items,
    warnings: warnings,
    validationIssues: validationIssues,
    durationMs: Date.now() - started
  });
  report.summary = buildReportingSummary_(items);
  report.statistics = buildInventoryStatistics_(items);
  return report;
}

function readInventorySummaryItemFromMatrix_(values, product, category) {
  const type = String(product.type || '').trim().toUpperCase();
  const row = Number(product.inventoryRow);
  const rowValues = values[row - 1] || [];
  if (isDirectFinalInventoryProduct_(product)) {
    const directValue = matrixValueOrBlank_(rowValues, 'B');
    const directItem = {
      key: product.normalizedName || normalizeText(product.name),
      product: product.name,
      category: category,
      type: type,
      inventoryRow: row,
      unit: 'szt.',
      finalTotal: directValue,
      total: directValue,
      details: { directFinal: directValue },
      cells: { finalTotal: 'B' + row, directFinal: 'B' + row }
    };
    directItem.values = { 'Stan końcowy': directValue };
    directItem.hasValue = directValue !== '';
    return directItem;
  }
  const layout = getInventorySummaryLayout_(type);

  const item = {
    key: product.normalizedName || normalizeText(product.name),
    product: product.name,
    category: category,
    type: type,
    inventoryRow: row,
    unit: layout.unit,
    finalTotal: matrixValueOrBlank_(rowValues, layout.finalTotal),
    total: matrixValueOrBlank_(rowValues, layout.finalTotal),
    details: {},
    cells: {}
  };
  addSummaryCellAddress_(item.cells, 'finalTotal', layout.finalTotal, row);

  if (type === CONFIG.PRODUCT_TYPES.LOCATION) {
    item.details = {
      warehouse: matrixValueOrBlank_(rowValues, layout.warehouse),
      darkroom: matrixValueOrBlank_(rowValues, layout.darkroom),
      fridges: matrixValueOrBlank_(rowValues, layout.fridges)
    };
    addSummaryCellAddress_(item.cells, 'warehouse', layout.warehouse, row);
    addSummaryCellAddress_(item.cells, 'darkroom', layout.darkroom, row);
    addSummaryCellAddress_(item.cells, 'fridges', layout.fridges, row);
  } else {
    item.details = {
      grossWeight: matrixValueOrBlank_(rowValues, layout.grossWeight),
      emptyContainerWeight: matrixValueOrBlank_(rowValues, layout.emptyContainerWeight),
      openNet: matrixValueOrBlank_(rowValues, layout.openNet),
      prepNet: matrixValueOrBlank_(rowValues, layout.prepNet),
      fullUnits: matrixValueOrBlank_(rowValues, layout.fullUnits),
      unitCapacity: matrixValueOrBlank_(rowValues, layout.unitCapacity),
      fullUnitsVolume: matrixValueOrBlank_(rowValues, layout.fullUnitsVolume)
    };
    addSummaryCellAddress_(item.cells, 'grossWeight', layout.grossWeight, row);
    addSummaryCellAddress_(item.cells, 'fullUnits', layout.fullUnits, row);
    item.packaging = getProductPackagingProfile_(product);
    applyPackagingAwareReportingValues_(item);
  }

  item.values = buildLegacyReviewValues_(item);
  item.hasValue = isReportingItemCompleted_(item);
  return item;
}

function applyPackagingAwareReportingValues_(item) {
  const details = item && item.details || {};
  const profile = item && item.packaging;
  let fullVolume = details.fullUnitsVolume;
  if (fullVolume === '' || fullVolume === null || fullVolume === undefined) {
    const units = Number(details.fullUnits);
    const capacity = Number(details.unitCapacity);
    fullVolume = Number.isFinite(units) && Number.isFinite(capacity) ? units * capacity : 0;
    details.fullUnitsVolume = fullVolume;
  }
  const prep = Number(details.prepNet);
  const extras = (Number.isFinite(prep) ? prep : 0) +
    (Number.isFinite(Number(fullVolume)) ? Number(fullVolume) : 0);
  if (details.grossWeight === '') {
    details.openNet = 0;
    item.finalTotal = extras;
    item.total = extras;
    return item;
  }

  let openNet = Number(details.openNet);
  if (profile && profile.mode === CONFIG.PACKAGING_MODES.GROSS_REFERENCE) {
    openNet = calculateReferenceRemainingQuantity_(details.grossWeight, profile);
  } else if (
    profile && profile.mode === CONFIG.PACKAGING_MODES.TARE_UNKNOWN ||
    details.emptyContainerWeight === ''
  ) {
    openNet = 0;
  } else {
    const gross = Number(details.grossWeight);
    const tare = Number(details.emptyContainerWeight);
    openNet = Number.isFinite(gross) && Number.isFinite(tare) ? Math.max(0, gross - tare) : 0;
  }
  details.openNet = Number.isFinite(Number(openNet)) ? Number(openNet) : 0;
  const total = details.openNet + extras;
  item.finalTotal = total;
  item.total = total;
  return item;
}

function applyTareSafeReportingValues_(item) {
  return applyPackagingAwareReportingValues_(item);
}

function matrixValueOrBlank_(rowValues, columnLetter) {
  const column = normalizeColumnLetter_(columnLetter);
  if (!column) return '';
  const index = columnLetterToNumber290_(column) - 1;
  const value = rowValues[index];
  if (value === '' || value === null || value === undefined) return '';
  const number = Number(value);
  return Number.isFinite(number) ? number : value;
}

function columnLetterToNumber290_(letters) {
  return String(letters || '').toUpperCase().split('').reduce((total, character) => total * 26 + character.charCodeAt(0) - 64, 0);
}

function isReportingItemCompleted_(item) {
  const details = item.details || {};
  if (item.type === CONFIG.PRODUCT_TYPES.LOCATION) {
    return getLocationAreaDefinitions_().some(area => details[area.columnKey] !== '');
  }
  if (hasMissingUnitCapacityForFullUnits_(details)) return false;
  return details.grossWeight !== '' || details.prepNet !== '' || details.fullUnits !== '';
}

function hasMissingUnitCapacityForFullUnits_(details) {
  const source = details || {};
  const units = source.fullUnits === '' ? 0 : Number(source.fullUnits);
  if (!Number.isFinite(units) || units <= 0) return false;
  const capacity = source.unitCapacity === '' ? NaN : Number(source.unitCapacity);
  return !Number.isFinite(capacity) || capacity <= 0;
}

function classifyFinalReviewRequirement_(item) {
  const details = item && item.details || {};
  if (isDirectFinalInventoryProduct_({name: item && item.product})) {
    return item.finalTotal === '' ? 'MISSING_PRODUCT' : '';
  }
  if (item.type === CONFIG.PRODUCT_TYPES.LOCATION) {
    return getLocationAreaDefinitions_().some(area => details[area.columnKey] !== '')
      ? ''
      : 'MISSING_PRODUCT';
  }
  const grossMissing = details.grossWeight === '';
  const unitsProvided = details.fullUnits !== '';
  const prepProvided = details.prepNet !== '';
  if (hasMissingUnitCapacityForFullUnits_(details)) return 'UNIT_CAPACITY_MISSING';
  if (grossMissing && unitsProvided) return 'FULL_UNITS_ONLY';
  if (grossMissing && !unitsProvided && !prepProvided) return 'MISSING_PRODUCT';
  const profile = item.packaging || {};
  if (!grossMissing && (
    profile.mode === CONFIG.PACKAGING_MODES.TARE_UNKNOWN ||
    profile.mode === CONFIG.PACKAGING_MODES.TARE_KNOWN && details.emptyContainerWeight === ''
  )) return 'PACKAGING_DATA_MISSING';
  if (
    !grossMissing &&
    profile.mode === CONFIG.PACKAGING_MODES.GROSS_REFERENCE &&
    Number(details.grossWeight) > Number(profile.referenceGrossWeight)
  ) return 'REFERENCE_RESET_REQUIRED';
  return '';
}

function buildFinalReviewCategoryGroups_(items) {
  const groups = {};
  (items || []).forEach(item => {
    if (!item.reviewReason) return;
    const category = String(item.category || 'BEZ KATEGORII');
    if (!groups[category]) {
      groups[category] = {
        category: category,
        total: 0,
        missingProducts: 0,
        fullUnitsOnly: 0,
        unitCapacityMissing: 0,
        packagingDataMissing: 0,
        referenceResetRequired: 0,
        products: []
      };
    }
    const group = groups[category];
    group.total++;
    if (item.reviewReason === 'MISSING_PRODUCT') group.missingProducts++;
    if (item.reviewReason === 'FULL_UNITS_ONLY') group.fullUnitsOnly++;
    if (item.reviewReason === 'UNIT_CAPACITY_MISSING') group.unitCapacityMissing++;
    if (item.reviewReason === 'PACKAGING_DATA_MISSING') group.packagingDataMissing++;
    if (item.reviewReason === 'REFERENCE_RESET_REQUIRED') group.referenceResetRequired++;
    group.products.push({
      key: item.key,
      product: item.product,
      type: item.type,
      reason: item.reviewReason,
      finalTotal: item.finalTotal,
      unit: item.unit
    });
  });
  return Object.keys(groups).sort((a, b) => a.localeCompare(b, 'pl'))
    .map(category => groups[category]);
}

function buildReportingSummary_(items) {
  const categories = {};
  items.forEach(item => {
    const category = item.category;
    if (!categories[category]) categories[category] = { products: 0, completed: 0, missing: 0 };
    categories[category].products++;
    if (item.hasValue) categories[category].completed++;
    else categories[category].missing++;
  });
  return {
    products: items.length,
    completed: items.filter(item => item.hasValue).length,
    missing: items.filter(item => !item.hasValue).length,
    warningProducts: items.filter(item => item.flags && item.flags.some(flag => flag !== 'BRAK WARTOŚCI')).length,
    categories: categories,
    reviewCategories: buildFinalReviewCategoryGroups_(items).length,
    reviewProducts: items.filter(item => item.reviewReason).length
  };
}

function resolveReportingSourceSheetName_(sourceSheetName) {
  const requested = String(sourceSheetName || '').trim();
  if (requested) {
    const requestedSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets().find(sheet => normalizeText(sheet.getName()) === normalizeText(requested));
    if (requestedSheet) return requestedSheet.getName();
  }
  const inventory = getSheetByConfiguredName_(CONFIG.SHEETS.INVENTORY);
  if (!inventory) throw new Error('Nie znaleziono bieżącego arkusza INWENTURA.');
  return inventory.getName();
}
