/**
 * Inventory PRO 4.3.4 — automatyczna konfiguracja produktów PAWILONÓW.
 * Mapowanie kolumn jest odczytywane z fizycznych nagłówków arkusza,
 * a CONFIG.INVENTORY_LAYOUT pozostaje bezpiecznym fallbackiem.
 */

function rebuildProductConfiguration() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Odbuduj konfigurację PAWILONÓW',
    'Skrypt wyczyści konfigurację w kolumnach D:L arkusza SŁOWNIK i zbuduje ją ponownie z arkusza ' +
      CONFIG.SHEETS.INVENTORY + '. Aliasy w A:B pozostaną bez zmian. Przed zmianą powstanie ukryta kopia SŁOWNIKA. Kontynuować?',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  try {
    const scan = scanInventoryProductsWithDiagnostics_();
    const existingCount = loadProductConfigurations().length;
    const validation = validateScannedProductConfiguration_(scan, existingCount);
    if (validation.errors.length) {
      throw new Error('Skan konfiguracji został zablokowany:\n- ' + validation.errors.join('\n- '));
    }

    if (validation.requiresCountDropConfirmation) {
      const second = ui.alert(
        'Podejrzana zmiana liczby produktów',
        'Obecna konfiguracja: ' + existingCount + ' produktów.\n' +
          'Nowy skan: ' + scan.products.length + ' produktów.\n\n' +
          'To może oznaczać nierozpoznany nagłówek lub usunięte wiersze. Kontynuować mimo różnicy?',
        ui.ButtonSet.YES_NO
      );
      if (second !== ui.Button.YES) return { success: false, cancelled: true };
    }

    const backupSheetName = createDictionaryConfigurationBackup_();
    writeFullProductConfiguration_(scan.products);
    ensureConfiguredDirectFinalAliases_(scan.products);
    invalidateProductCatalogCache_();
    ui.alert(
      'Configuration Builder',
      'Zbudowano konfigurację dla ' + scan.products.length + ' produktów.\n' +
        'Kopia bezpieczeństwa: ' + backupSheetName +
        (validation.warnings.length ? '\nOstrzeżenia: ' + validation.warnings.length : ''),
      ui.ButtonSet.OK
    );
    return {
      success: true,
      products: scan.products.length,
      backupSheetName: backupSheetName,
      diagnostics: scan.diagnostics,
      warnings: validation.warnings
    };
  } catch (error) {
    ui.alert('Błąd Configuration Builder', error.message || String(error), ui.ButtonSet.OK);
    throw error;
  }
}

function syncProductConfiguration() {
  const ui = SpreadsheetApp.getUi();
  try {
    const scan = scanInventoryProductsWithDiagnostics_();
    const validation = validateScannedProductConfiguration_(scan, 0);
    if (validation.errors.length) {
      throw new Error('Skan konfiguracji został zablokowany:\n- ' + validation.errors.join('\n- '));
    }
    const existingIndex = getExistingConfigurationIndex_();
    const newProducts = scan.products.filter(product => !existingIndex[product.normalizedName]);
    if (!newProducts.length) {
      ui.alert('Synchronizacja zakończona', 'Nie znaleziono nowych produktów.', ui.ButtonSet.OK);
      return { success: true, added: 0 };
    }
    appendProductConfigurations_(newProducts);
    ensureConfiguredDirectFinalAliases_(scan.products);
    invalidateProductCatalogCache_();
    ui.alert('Synchronizacja zakończona', 'Dodano nowych produktów: ' + newProducts.length, ui.ButtonSet.OK);
    return { success: true, added: newProducts.length, diagnostics: scan.diagnostics };
  } catch (error) {
    ui.alert('Błąd synchronizacji', error.message || String(error), ui.ButtonSet.OK);
    throw error;
  }
}

function scanInventoryProducts_() {
  return scanInventoryProductsWithDiagnostics_().products;
}

function scanInventoryProductsWithDiagnostics_() {
  const sheet = getSheetByConfiguredName_(CONFIG.SHEETS.INVENTORY);
  if (!sheet) throw new Error('Nie znaleziono arkusza: ' + CONFIG.SHEETS.INVENTORY);

  const lastRow = sheet.getLastRow();
  const lastColumn = Math.max(sheet.getLastColumn(), getInventoryLayoutMaxColumn_());
  if (lastRow < 1) return { products: [], diagnostics: { errors: [], warnings: [] } };

  const dataRange = sheet.getRange(1, 1, lastRow, lastColumn);
  const values = dataRange.getDisplayValues();
  const mergedHeaderRows = buildMergedHeaderRowMap_(dataRange);

  let normalColumns = detectInventoryInputColumnsFromHeaderRow_(
    values[0] || [],
    CONFIG.PRODUCT_TYPES.NORMAL,
    getInputColumnsForProductType_(CONFIG.PRODUCT_TYPES.NORMAL)
  );
  let kegColumns = getInputColumnsForProductType_(CONFIG.PRODUCT_TYPES.KEG);
  let locationColumns = getInputColumnsForProductType_(CONFIG.PRODUCT_TYPES.LOCATION);

  const products = [];
  const usedNames = {};
  const diagnostics = {
    errors: [],
    warnings: [],
    headerRows: [],
    headerContinuationRows: [],
    skippedRows: []
  };
  let currentCategory = '';
  let currentType = CONFIG.PRODUCT_TYPES.NORMAL;

  for (let rowIndex = 0; rowIndex < values.length; rowIndex++) {
    const sheetRow = rowIndex + 1;
    const rowValues = values[rowIndex] || [];
    const productName = String(rowValues[0] || '').trim();
    const rowText = rowValues.filter(value => String(value || '').trim()).join(' ').trim();
    if (!rowText) continue;

    const detectedHeader = detectSectionHeader_(
      rowText, sheetRow, mergedHeaderRows, productName, rowValues
    );
    if (detectedHeader.isHeader) {
      currentCategory = detectedHeader.category;
      currentType = detectedHeader.type;
      diagnostics.headerRows.push(sheetRow);

      if (currentType === CONFIG.PRODUCT_TYPES.LOCATION) {
        locationColumns = mergeDetectedProductColumns_(
          locationColumns,
          detectInventoryInputColumnsFromHeaderRow_(rowValues, currentType, locationColumns)
        );
      } else if (currentType === CONFIG.PRODUCT_TYPES.KEG) {
        kegColumns = mergeDetectedProductColumns_(
          kegColumns,
          detectInventoryInputColumnsFromHeaderRow_(rowValues, currentType, kegColumns)
        );
      } else {
        normalColumns = mergeDetectedProductColumns_(
          normalColumns,
          detectInventoryInputColumnsFromHeaderRow_(rowValues, currentType, normalColumns)
        );
      }
      continue;
    }

    if (currentCategory && isInventoryHeaderContinuationRow_(rowValues, currentType)) {
      diagnostics.headerContinuationRows.push(sheetRow);
      if (currentType === CONFIG.PRODUCT_TYPES.LOCATION) {
        locationColumns = mergeDetectedProductColumns_(
          locationColumns,
          detectInventoryInputColumnsFromHeaderRow_(rowValues, currentType, locationColumns)
        );
      } else if (currentType === CONFIG.PRODUCT_TYPES.KEG) {
        kegColumns = mergeDetectedProductColumns_(
          kegColumns,
          detectInventoryInputColumnsFromHeaderRow_(rowValues, currentType, kegColumns)
        );
      } else {
        normalColumns = mergeDetectedProductColumns_(
          normalColumns,
          detectInventoryInputColumnsFromHeaderRow_(rowValues, currentType, normalColumns)
        );
      }
      continue;
    }

    if (!productName || isIgnoredInventoryText_(productName)) {
      diagnostics.skippedRows.push(sheetRow);
      continue;
    }

    const normalizedName = normalizeText(productName);
    if (!normalizedName) continue;
    if (usedNames[normalizedName]) {
      diagnostics.warnings.push('Zduplikowana nazwa produktu „' + productName + '” w wierszu ' + sheetRow + '.');
      continue;
    }
    if (!currentCategory) {
      diagnostics.errors.push('Produkt „' + productName + '” w wierszu ' + sheetRow + ' nie ma fizycznej kategorii.');
      continue;
    }

    const effectiveType = inferInventoryProductType_(
      currentType, currentCategory, productName, rowValues
    );
    const directFinalColumn = getDirectFinalInventoryColumn_({ name: productName });
    const columns = directFinalColumn
      ? { quantity: directFinalColumn, weight: '', warehouse: '', darkroom: '', fridges: '' }
      : effectiveType === CONFIG.PRODUCT_TYPES.LOCATION
        ? locationColumns
        : effectiveType === CONFIG.PRODUCT_TYPES.KEG
          ? kegColumns
          : normalColumns;

    const mapping = validateProductColumnMapping_(
      effectiveType,
      columns,
      { name: productName, type: effectiveType }
    );
    if (!mapping.valid) {
      diagnostics.errors.push(
        'Nieprawidłowe mapowanie produktu „' + productName + '” (wiersz ' + sheetRow + '): ' +
        mapping.errors.join(' ')
      );
      continue;
    }

    usedNames[normalizedName] = true;
    products.push(createConfigurationProduct_(
      productName, normalizedName, currentCategory, effectiveType, sheetRow, mapping.columns
    ));
  }

  return { products: products, diagnostics: diagnostics };
}

function buildMergedHeaderRowMap_(dataRange) {
  const map = {};
  if (!dataRange || typeof dataRange.getMergedRanges !== 'function') return map;
  dataRange.getMergedRanges().forEach(range => {
    if (range.getNumColumns() <= 1) return;
    for (let row = range.getRow(); row <= range.getLastRow(); row++) map[row] = true;
  });
  return map;
}

function detectSectionHeader_(rowText, sheetRow, mergedHeaderRows, firstCell, rowValues) {
  const primary = String(firstCell === undefined ? rowText : firstCell || '').trim();
  const normalizedPrimary = normalizeText(primary);
  const normalizedRow = normalizeText(rowText);
  const exactCategory = normalizeBusinessCategory_(primary);
  const isMergedHeader = Boolean((mergedHeaderRows || {})[sheetRow]);

  const isExactCategory = exactCategory && normalizeText(exactCategory) === normalizedPrimary;
  const isKeg = normalizedPrimary === 'keg' || normalizedPrimary === 'piwo keg' ||
    normalizedPrimary === 'piwa keg';
  const isBottleBeer = normalizedPrimary.indexOf('piwo butel') === 0 ||
    normalizedPrimary.indexOf('piwa butel') === 0;
  const isSoft = exactCategory === 'SOFTY' || normalizedPrimary === 'softy' ||
    normalizedPrimary.indexOf('softy na szt') === 0;
  const isGenericBeer = exactCategory === 'PIWO' || normalizedPrimary === 'piwo' ||
    normalizedPrimary === 'piwa';

  if (isKeg) return { isHeader: true, category: 'PIWO KEG', type: CONFIG.PRODUCT_TYPES.KEG };
  if (isBottleBeer) return { isHeader: true, category: 'PIWO BUTELKI', type: CONFIG.PRODUCT_TYPES.LOCATION };
  if (isSoft) return { isHeader: true, category: 'SOFTY', type: CONFIG.PRODUCT_TYPES.LOCATION };
  if (isGenericBeer) return { isHeader: true, category: 'PIWO', type: CONFIG.PRODUCT_TYPES.NORMAL };
  if (isExactCategory || (isMergedHeader && exactCategory)) {
    return { isHeader: true, category: exactCategory, type: CONFIG.PRODUCT_TYPES.NORMAL };
  }

  const rowCategory = isMergedHeader ? normalizeBusinessCategory_(normalizedRow) : '';
  if (rowCategory) {
    return {
      isHeader: true,
      category: rowCategory,
      type: rowCategory === 'SOFTY' || rowCategory === 'PIWO BUTELKI'
        ? CONFIG.PRODUCT_TYPES.LOCATION
        : rowCategory === 'PIWO KEG'
          ? CONFIG.PRODUCT_TYPES.KEG
          : CONFIG.PRODUCT_TYPES.NORMAL
    };
  }
  return { isHeader: false, category: '', type: '' };
}

function isLikelyCategoryHeader_(rowText) {
  const normalized = normalizeText(rowText);
  if (!normalized) return false;
  return Boolean(normalizeBusinessCategory_(normalized));
}

function isIgnoredInventoryText_(text) {
  const normalized = normalizeText(text);
  if (!normalized) return true;
  if (resolveLocationHeaderArea_(normalized)) return true;
  return [
    'produkt', 'produkty', 'nazwa produktu', 'wpis uzytkownika',
    'razem', 'suma', 'waga szt w butelce kegu', 'waga w kegu',
    'waga pusty keg', 'waga bez kega', 'pelne kegi', 'pojemnosc',
    'calosc l', 'calosc'
  ].includes(normalized);
}


function isExactInventoryHeaderText_(value) {
  const normalized = normalizeText(value);
  if (!normalized) return true;
  if (isIgnoredInventoryText_(normalized)) return true;

  const exactHeaders = [
    'bitter', 'brandy', 'gin', 'likier', 'piwo keg', 'piwo butelki',
    'rum', 'tequila', 'wermut', 'whisky', 'wino', 'wodka', 'premixy',
    'softy', 'softy na szt', 'kawa'
  ];
  return exactHeaders.includes(normalized);
}

function validateScannedProductConfiguration_(scan, existingCount) {
  const products = scan && scan.products || [];
  const diagnostics = scan && scan.diagnostics || { errors: [], warnings: [] };
  const errors = (diagnostics.errors || []).slice();
  const warnings = (diagnostics.warnings || []).slice();

  if (!products.length) errors.push('Nie znaleziono żadnych produktów w arkuszu INWENTURA.');
  products.forEach(product => {
    const mapping = validateProductColumnMapping_(product.type, product.columns, product);
    if (!mapping.valid) {
      errors.push('Produkt „' + product.name + '”: ' + mapping.errors.join(' '));
    }
    if (isExactInventoryHeaderText_(product.name)) {
      errors.push('Nagłówek został błędnie rozpoznany jako produkt: „' + product.name + '”.');
    }
  });

  const current = Number(existingCount) || 0;
  const drop = Math.max(current - products.length, 0);
  const relativeDrop = current ? drop / current : 0;
  const policy = CONFIG.CONFIGURATION_BUILDER || {};
  const requiresCountDropConfirmation = current > 0 &&
    drop > (Number(policy.MAX_ABSOLUTE_COUNT_DROP) || 10) &&
    relativeDrop > (Number(policy.MAX_RELATIVE_COUNT_DROP) || 0.20);
  if (requiresCountDropConfirmation) {
    warnings.push('Nowy skan zmniejsza liczbę produktów o ' + drop + ' (' + Math.round(relativeDrop * 100) + '%).');
  }

  return {
    valid: errors.length === 0,
    errors: errors,
    warnings: warnings,
    requiresCountDropConfirmation: requiresCountDropConfirmation
  };
}

function createConfigurationProduct_(name, normalizedName, category, type, inventoryRow, detectedColumns) {
  const columns = cloneProductColumns_(detectedColumns || getInputColumnsForProductType_(type));
  return {
    name: name,
    normalizedName: normalizedName,
    type: type,
    category: requireBusinessCategory_(category, name, inventoryRow),
    inventoryRow: inventoryRow,
    columns: columns,
    active: true
  };
}

function createDictionaryConfigurationBackup_() {
  const sheet = getDictionarySheet_();
  const spreadsheet = sheet.getParent();
  const timestamp = Utilities.formatDate(
    new Date(),
    spreadsheet.getSpreadsheetTimeZone() || 'Europe/Warsaw',
    'yyyyMMdd-HHmmss'
  );
  const base = 'BACKUP SLOWNIK ' + timestamp;
  let name = base;
  let suffix = 2;
  while (spreadsheet.getSheetByName(name)) {
    name = base + '-' + suffix;
    suffix++;
  }
  sheet.copyTo(spreadsheet).setName(name).hideSheet();
  return name;
}

function writeFullProductConfiguration_(products) {
  const sheet = getDictionarySheet_();
  ensureConfigurationHeaders_(sheet);
  const rowsToClear = Math.max(sheet.getMaxRows() - 1, 1);
  sheet.getRange(2, 4, rowsToClear, 9).clearContent();
  const output = products.map(configurationToRow_);
  if (output.length) sheet.getRange(2, 4, output.length, 9).setValues(output);
  formatConfigurationTable_(sheet, output.length);
}

function appendProductConfigurations_(products) {
  const sheet = getDictionarySheet_();
  ensureConfigurationHeaders_(sheet);
  const output = products.map(configurationToRow_);
  if (!output.length) return;
  const startRow = findNextConfigurationRow_(sheet);
  sheet.getRange(startRow, 4, output.length, 9).setValues(output);
  formatConfigurationTable_(sheet, startRow + output.length - 2);
}

function configurationToRow_(product) {
  return [
    product.name,
    product.type,
    product.category,
    product.columns.quantity,
    product.columns.weight,
    product.columns.warehouse,
    product.columns.darkroom,
    product.columns.fridges,
    product.active ? 'TAK' : 'NIE'
  ];
}

function getExistingConfigurationIndex_() {
  const index = {};
  loadProductConfigurations().forEach(product => { index[product.normalizedName] = true; });
  return index;
}

function ensureConfiguredDirectFinalAliases_(products) {
  const definitions = getDirectFinalProductDefinitions_();
  if (!definitions.length) return 0;
  const catalog = Array.isArray(products) ? products : [];
  const existing = loadAliases();
  const rows = [];
  definitions.forEach(definition => {
    const names = getDirectFinalProductNames_(definition);
    const product = catalog.find(item => names.indexOf(normalizeText(item && item.name || '')) >= 0);
    if (!product) return;
    names.forEach(alias => {
      if (!alias || existing[alias]) return;
      rows.push([alias, product.name]);
      existing[alias] = product.name;
    });
  });
  if (!rows.length) return 0;
  const sheet = getDictionarySheet_();
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 2).setValues(rows);
  return rows.length;
}

function findNextConfigurationRow_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 2;
  const values = sheet.getRange(2, 4, lastRow - 1, 1).getDisplayValues();
  for (let index = 0; index < values.length; index++) {
    if (!String(values[index][0] || '').trim()) return index + 2;
  }
  return lastRow + 1;
}

function ensureConfigurationHeaders_(sheet) {
  const labels = getLocationColumnLabelMap_();
  sheet.getRange(1, 4, 1, 9)
    .setValues([[
      'Produkt',
      'Typ',
      'Kategoria',
      'Kolumna sztuk',
      'Kolumna wagi',
      labels.warehouse || 'Lokalizacja 1',
      labels.darkroom || 'Lokalizacja 2',
      labels.fridges || 'Lokalizacja 3',
      'Aktywny'
    ]])
    .setFontWeight('bold')
    .setBackground('#d9ead3');
}

function formatConfigurationTable_(sheet, productCount) {
  sheet.autoResizeColumns(4, 9);
  if (productCount > 0) {
    sheet.getRange(2, 4, productCount, 9).setVerticalAlignment('middle');
  }
}

function cleanCategoryName_(rowText) {
  return String(rowText || '').replace(/\s+/g, ' ').trim();
}
