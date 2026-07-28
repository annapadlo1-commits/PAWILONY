/**
 * Inventory PRO 4.3.4
 * Bezpieczny workflow tworzenia nowych produktów z poziomu importu.
 */

function getProductResolverData(inputName) {
  const context = buildRuntimeContext_();
  return buildProductResolverPayload_(context, inputName);
}

function buildProductResolverPayload_(context, inputName) {
  const normalized = normalizeText(inputName);
  const source = String(inputName || '').trim();
  const suggestions = source ? getRecognitionShortlist_(source, context)
    .map(product => scoreRecognitionCandidate_(source, product))
    .sort((a, b) => b.score - a.score || a.product.name.localeCompare(b.product.name))
    .slice(0, 12)
    .map(item => ({
      name: item.product.name,
      type: item.product.type,
      category: item.product.category,
      score: Math.round(item.score),
      inventoryRow: item.product.inventoryRow || null
    })) : [];

  return {
    inputName: String(inputName || '').trim(),
    normalizedInput: normalized,
    suggestions: suggestions,
    products: context.catalog.map(product => ({
      name: product.name,
      type: product.type,
      category: product.category,
      aliases: (product.aliases || []).slice(),
      inventoryRow: product.inventoryRow || null
    })).sort((a, b) => a.name.localeCompare(b.name)),
    categories: Array.from(new Set(context.catalog.map(product => product.category).filter(Boolean))).sort(),
    productTypes: Object.keys(CONFIG.PRODUCT_TYPES).map(key => CONFIG.PRODUCT_TYPES[key]),
    locations: getLocationUiOptions_()
  };
}

function createNewProductFromImport(request) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  let inventorySheet = null;
  let productName = '';
  let insertion = null;
  let dictionaryResult = null;
  let auditResult = null;
  let packagingResult = null;

  try {
    const data = request || {};
    const name = String(data.name || '').trim();
    productName = name;
    const referenceName = String(data.referenceProduct || '').trim();
    const requestedType = String(data.productType || '').trim().toUpperCase();
    const requestedCategory = String(data.category || '').trim();

    validateNewProductName_(name);

    const context = buildRuntimeContext_();
    if (context.productIndex[normalizeText(name)]) {
      throw new Error('Produkt o tej nazwie już istnieje: ' + name);
    }
    if (context.aliasIndex[normalizeText(name)]) {
      throw new Error('Ta nazwa jest już aliasem istniejącego produktu. Wybierz produkt zamiast tworzyć nowy.');
    }

    const reference = referenceName
      ? context.productIndex[normalizeText(referenceName)]
      : null;
    if (referenceName && !reference) {
      throw new Error('Nie znaleziono produktu referencyjnego: ' + referenceName);
    }

    const productType = String(reference ? reference.type : requestedType).toUpperCase();
    const category = normalizeBusinessCategory_(reference ? reference.category : requestedCategory);
    if (!Object.values(CONFIG.PRODUCT_TYPES).includes(productType)) {
      throw new Error('Wybierz prawidłowy typ produktu.');
    }
    if (!category) throw new Error('Wybierz kategorię produktu.');

    inventorySheet = getSheetByConfiguredName_(CONFIG.SHEETS.INVENTORY);
    if (!inventorySheet) throw new Error('Nie znaleziono arkusza inwentury.');

    const columns = resolveNewProductColumns_(reference, productType);
    insertion = insertProductInventoryRow_(
      inventorySheet, name, reference, productType, category, columns, {
        unitCapacity: data.unitCapacity,
        packagingMode: data.packaging && data.packaging.mode,
        referenceNetQuantity: data.packaging && data.packaging.referenceNetQuantity
      }
    );

    dictionaryResult = ensureNewProductInDictionary_({
      name: name,
      productType: productType,
      category: category,
      columns: columns,
      sourceAlias: data.sourceAlias || ''
    });

    const pendingProduct = {
      name: name,
      normalizedName: normalizeText(name),
      type: productType,
      category: category,
      inventoryRow: insertion.row
    };
    packagingResult = saveProductPackagingProfile_(pendingProduct, data.packaging || {
      mode: CONFIG.PACKAGING_MODES.TARE_UNKNOWN
    });
    // Punkt referencyjny opisuje profil opakowania, ale nie jest osobnym
    // pomiarem bieżącej inwentury. W ścieżce importu wartość z analizowanego
    // wiersza zapisze za chwilę InventoryWriter. Wstępne wpisanie tej samej
    // wagi tutaj powodowało jej podwojenie (np. 1,2 + 1,2 = 2,4).
    applyProductPackagingProfileToInventory_(pendingProduct, packagingResult.profile, {
      workflow: 'IMPORT'
    });
    const finalization = finalizeNewProductInventoryRow_(
      inventorySheet,
      pendingProduct,
      packagingResult.profile,
      insertion
    );

    auditResult = appendNewProductAudit_(
      name,
      reference ? reference.name : '',
      productType,
      category,
      insertion.row,
      data.sourceInput || ''
    );

    SpreadsheetApp.flush();
    invalidateProductCatalogCache_();

    const refreshed = buildRuntimeContext_();
    const product = refreshed.productIndex[normalizeText(name)];
    if (!product || !product.inventoryRow) {
      throw new Error('Produkt został dodany, ale nie udało się odświeżyć katalogu. Uruchom synchronizację słownika.');
    }

    const formulaVerification = verifyCanonicalFormulasForProductRow_(inventorySheet, product);
    const invalidFormula = formulaVerification.find(item => !item.valid);
    if (invalidFormula) {
      throw new Error('Nowy produkt nie otrzymał poprawnej formuły w komórce ' + invalidFormula.cell + '.');
    }

    return {
      success: true,
      product: {
        name: product.name,
        type: product.type,
        category: product.category,
        inventoryRow: product.inventoryRow
      },
      formulasApplied: finalization.formulasApplied,
      formulaVerification: formulaVerification,
      dictionaryUpdated: Boolean(dictionaryResult && dictionaryResult.configurationRow),
      dictionaryConfigurationRow: dictionaryResult.configurationRow,
      dictionaryAliasesAdded: dictionaryResult.addedAliases,
      packagingMode: packagingResult.profile.mode,
      packagingProfile: packagingResult.profile,
      newProductAuditAdded: Boolean(auditResult && auditResult.added),
      message: 'Dodano nowy produkt: ' + product.name + ', formuły i konfigurację słownika.'
    };
  } catch (error) {
    const rollback = { audit: null, dictionary: null, inventory: null, errors: [] };

    if (packagingResult && packagingResult.saved) {
      try {
        rollbackProductPackagingProfile_(productName, null);
      } catch (rollbackError) {
        rollback.errors.push('profil opakowania: ' + normalizeError_(rollbackError).message);
      }
    }
    if (auditResult && auditResult.added) {
      try {
        rollback.audit = rollbackNewProductAudit_(productName, auditResult);
      } catch (rollbackError) {
        rollback.errors.push('audyt: ' + normalizeError_(rollbackError).message);
      }
    }
    if (dictionaryResult) {
      try {
        rollback.dictionary = rollbackNewProductDictionaryEntry_(productName, dictionaryResult);
      } catch (rollbackError) {
        rollback.errors.push('słownik: ' + normalizeError_(rollbackError).message);
      }
    }
    if (inventorySheet && insertion && insertion.row) {
      try {
        rollback.inventory = rollbackInsertedInventoryProductRow_(
          inventorySheet, insertion.row, productName
        );
      } catch (rollbackError) {
        rollback.errors.push('inwentura: ' + normalizeError_(rollbackError).message);
      }
    }

    try {
      SpreadsheetApp.flush();
      invalidateProductCatalogCache_();
    } catch (flushError) {
      rollback.errors.push('odświeżenie: ' + normalizeError_(flushError).message);
    }

    logError(
      'ProductCreationService',
      'createNewProductFromImport',
      error,
      { product: productName, rollback: rollback },
      0
    );
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function validateNewProductName_(name) {
  if (!name) throw new Error('Nazwa produktu nie może być pusta.');
  if (name.length < 2) throw new Error('Nazwa produktu jest za krótka.');
  if (/^\d+(?:[.,]\d+)?$/.test(name)) throw new Error('Nazwa produktu nie może składać się tylko z liczby.');
}

function resolveNewProductColumns_(reference, productType) {
  const defaults = defaultColumnsForProductType_(productType);
  if (!reference || !reference.columns) return defaults;
  const validation = validateProductColumnMapping_(productType, reference.columns);
  if (!validation.valid) {
    logWarning(
      'ProductCreationService',
      'resolveNewProductColumns_',
      'Nie skopiowano nieprawidłowego mapowania produktu referencyjnego; użyto profilu PAWILONÓW.',
      { product: reference.name, errors: validation.errors }
    );
    return defaults;
  }
  return validation.columns;
}

function findProductInsertionContext_(reference, productType, category) {
  const products = scanInventoryProducts_();
  if (reference && reference.inventoryRow) {
    const style = products.find(product =>
      Number(product.inventoryRow) === Number(reference.inventoryRow)
    ) || reference;
    return { insertAfter: Number(reference.inventoryRow), styleReference: style };
  }

  const matching = products.filter(product =>
    String(product.type || '').toUpperCase() === String(productType || '').toUpperCase() &&
    normalizeBusinessCategory_(product.category) === normalizeBusinessCategory_(category)
  );
  if (!matching.length) {
    throw new Error(
      'Nie znaleziono sekcji „' + category + '” typu ' + productType +
      '. Wybierz produkt referencyjny z docelowej sekcji.'
    );
  }
  matching.sort((a, b) => Number(a.inventoryRow) - Number(b.inventoryRow));
  const last = matching[matching.length - 1];
  return { insertAfter: Number(last.inventoryRow), styleReference: last };
}

function insertProductInventoryRow_(sheet, name, reference, productType, category, columns, options) {
  const lastColumn = Math.max(sheet.getLastColumn(), getInventoryLayoutMaxColumn_());
  const context = findProductInsertionContext_(reference, productType, category);
  const insertAfter = Number(context.insertAfter) || 0;
  if (insertAfter < 1) throw new Error('Nie ustalono bezpiecznego miejsca wstawienia produktu.');

  const targetRow = insertAfter + 1;
  let rowInserted = false;
  try {
    sheet.insertRowAfter(insertAfter);
    rowInserted = true;
    const styleReference = context.styleReference;

    if (styleReference && styleReference.inventoryRow) {
      let sourceRow = Number(styleReference.inventoryRow);
      if (sourceRow > insertAfter) sourceRow += 1;
      const sourceRange = sheet.getRange(sourceRow, 1, 1, lastColumn);
      const targetRange = sheet.getRange(targetRow, 1, 1, lastColumn);
      sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
      targetRange.setNumberFormats(sourceRange.getNumberFormats());
      sheet.setRowHeight(targetRow, sheet.getRowHeight(sourceRow));
    }

    sheet.getRange(targetRow, 1).setValue(name);
    const product = {
      name: name,
      normalizedName: normalizeText(name),
      type: productType,
      category: category,
      inventoryRow: targetRow,
      columns: cloneProductColumns_(columns || defaultColumnsForProductType_(productType))
    };
    if (isDirectFinalInventoryProduct_(product)) {
      product.columns = {
        quantity: getDirectFinalInventoryColumn_(product),
        weight: '',
        warehouse: '',
        darkroom: '',
        fridges: ''
      };
    }
    clearNewProductInputCells_(sheet, targetRow, productType, product.columns);
    applyNewProductUnitCapacity_(
      sheet,
      product,
      styleReference,
      options && options.unitCapacity,
      options && options.packagingMode,
      options && options.referenceNetQuantity
    );

    return {
      row: targetRow,
      styleSourceRow: styleReference && styleReference.inventoryRow
        ? Number(styleReference.inventoryRow)
        : 0
    };
  } catch (error) {
    if (rowInserted) {
      try {
        rollbackInsertedInventoryProductRow_(sheet, targetRow, name);
        SpreadsheetApp.flush();
      } catch (rollbackError) {
        logError(
          'ProductCreationService',
          'insertProductInventoryRow_.rollback',
          rollbackError,
          { product: name, row: targetRow },
          0
        );
      }
    }
    throw error;
  }
}

function finalizeNewProductInventoryRow_(sheet, product, profile, insertion) {
  const row = Number(product && product.inventoryRow) || 0;
  if (!sheet || !row) throw new Error('Nie można dokończyć tworzenia wiersza produktu.');

  const lastColumn = Math.max(sheet.getLastColumn(), getInventoryLayoutMaxColumn_());
  let sourceRow = Number(insertion && insertion.styleSourceRow) || 0;
  if (sourceRow >= row) sourceRow += 1;
  if (sourceRow > 0 && sourceRow <= sheet.getLastRow()) {
    const sourceRange = sheet.getRange(sourceRow, 1, 1, lastColumn);
    const targetRange = sheet.getRange(row, 1, 1, lastColumn);
    sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
    targetRange.setNumberFormats(sourceRange.getNumberFormats());
    sheet.setRowHeight(row, sheet.getRowHeight(sourceRow));
    sheet.getRange(row, 1).setValue(product.name);
  }

  const formulasApplied = isDirectFinalInventoryProduct_(product)
    ? 0
    : applyCanonicalFormulasToProductRow_(sheet, product);
  SpreadsheetApp.flush();
  const verification = isDirectFinalInventoryProduct_(product)
    ? []
    : verifyCanonicalFormulasForProductRow_(sheet, product);
  const invalid = verification.filter(item => !item.valid);
  if (invalid.length) {
    throw new Error(
      'Nie udało się ustawić formuł nowego produktu: ' +
      invalid.map(item => item.cell).join(', ') + '.'
    );
  }
  return { formulasApplied: formulasApplied, formulaVerification: verification };
}

function resolveNewProductUnitCapacity_(explicitCapacity, grossReferenceQuantity, referenceCapacity) {
  const explicit = normalizePackagingNumber_(explicitCapacity);
  if (Number.isFinite(explicit) && explicit > 0) return explicit;
  const grossReference = normalizePackagingNumber_(grossReferenceQuantity);
  if (Number.isFinite(grossReference) && grossReference > 0) return grossReference;
  const inherited = normalizePackagingNumber_(referenceCapacity);
  return Number.isFinite(inherited) && inherited > 0 ? inherited : '';
}

function applyNewProductUnitCapacity_(sheet, product, reference, explicitCapacity, packagingMode, referenceNetQuantity) {
  if (!sheet || !product || product.type === CONFIG.PRODUCT_TYPES.LOCATION) {
    return {applied: false, value: ''};
  }
  const layout = getConfiguredInventoryLayout_(product.type) || {};
  if (!layout.unitCapacity) return {applied: false, value: ''};
  let referenceCapacity = '';
  if (reference && reference.inventoryRow) {
    const referenceLayout = getConfiguredInventoryLayout_(reference.type) || {};
    if (referenceLayout.unitCapacity) {
      referenceCapacity = sheet.getRange(
        referenceLayout.unitCapacity + Number(reference.inventoryRow)
      ).getValue();
    }
  }
  const grossReferenceQuantity =
    String(packagingMode || '').toUpperCase() === String(CONFIG.PACKAGING_MODES.GROSS_REFERENCE)
      ? referenceNetQuantity
      : '';
  const capacity = resolveNewProductUnitCapacity_(
    explicitCapacity,
    grossReferenceQuantity,
    referenceCapacity
  );
  const target = sheet.getRange(layout.unitCapacity + Number(product.inventoryRow));
  if (capacity === '') target.clearContent();
  else target.setValue(capacity);
  return {applied: capacity !== '', value: capacity};
}

function rollbackInsertedInventoryProductRow_(sheet, row, expectedName) {
  if (!sheet || !row) return { removed: false, conflict: 'Brak arkusza lub wiersza.' };
  const currentName = String(sheet.getRange(Number(row), 1).getDisplayValue() || '').trim();
  if (normalizeText(currentName) !== normalizeText(expectedName)) {
    const conflict =
      'Nie usunięto wiersza ' + row + ', ponieważ jego nazwa zmieniła się z „' +
      expectedName + '” na „' + (currentName || 'PUSTO') + '”.';
    logWarning(
      'ProductCreationService',
      'rollbackInsertedInventoryProductRow_',
      conflict,
      { row: row, expectedName: expectedName, currentName: currentName }
    );
    return { removed: false, conflict: conflict };
  }
  sheet.deleteRow(Number(row));
  return { removed: true, row: Number(row) };
}

function clearNewProductInputCells_(sheet, row, productType, columns) {
  const source = cloneProductColumns_(columns);
  const candidates = [
    source.quantity,
    source.weight,
    source.warehouse,
    source.darkroom,
    source.fridges
  ].filter(Boolean);

  Array.from(new Set(candidates)).forEach(column => {
    const normalizedColumn = normalizeColumnLetter_(column);
    const directFinalColumn = normalizeColumnLetter_(columns && columns.quantity);
    const isDirectFinalColumn = directFinalColumn === 'B' && normalizedColumn === 'B';
    if (!isDirectFinalColumn && !isAllowedInputColumnForProductType_(productType, normalizedColumn)) {
      throw new Error(
        'Kolumna ' + normalizedColumn + ' nie jest polem wejściowym typu ' + productType +
        ' w profilu PAWILONÓW.'
      );
    }
    if (isFormulaColumnForProductType_(productType, normalizedColumn)) {
      throw new Error(
        'Kolumna ' + normalizedColumn + ' jest kolumną obliczeniową typu ' + productType + '.'
      );
    }
    const cell = sheet.getRange(normalizedColumn + row);
    if (cell.getFormula()) {
      throw new Error(
        'Kolumna wejściowa ' + normalizedColumn + ' wskazuje formułę w nowym wierszu ' + row + '.'
      );
    }
    cell.clearContent();
  });
}

function defaultColumnsForProductType_(productType) {
  return getInputColumnsForProductType_(productType);
}

function appendProductConfiguration_(name, productType, category, reference) {
  const sheet = getDictionarySheet_();
  const columns = resolveNewProductColumns_(reference, productType);
  const row = [
    name,
    productType,
    category,
    columns.quantity || '',
    columns.weight || '',
    columns.warehouse || '',
    columns.darkroom || '',
    columns.fridges || '',
    'TAK'
  ];
  sheet.getRange(
    sheet.getLastRow() + 1,
    CONFIG.DICTIONARY.CONFIG_START_COLUMN,
    1,
    CONFIG.DICTIONARY.CONFIG_COLUMN_COUNT
  ).setValues([row]);
}

function getOrCreateNewProductsSheet_() {
  const sheet = getOrCreateConfiguredSheet_(CONFIG.SHEETS.NEW_PRODUCTS);
  const headers = [
    'Timestamp', 'User', 'Produkt', 'Produkt referencyjny', 'Typ',
    'Kategoria', 'Wiersz inwentury', 'Wpis z importu'
  ];
  const current = sheet.getRange(1, 1, 1, headers.length).getDisplayValues()[0];
  if (current.join('|') !== headers.join('|')) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sheet;
}

function appendNewProductAudit_(name, referenceName, productType, category, inventoryRow, sourceInput) {
  const sheet = getOrCreateNewProductsSheet_();
  const normalizedName = normalizeText(name);
  const lastRow = sheet.getLastRow();

  if (lastRow >= 2) {
    const existing = sheet.getRange(2, 3, lastRow - 1, 5).getValues();
    const duplicate = existing.some(row =>
      normalizeText(row[0]) === normalizedName &&
      Number(row[4] || 0) === Number(inventoryRow || 0)
    );
    if (duplicate) {
      logInfo(
        'ProductCreationService',
        'appendNewProductAudit_',
        'Pominięto duplikat wpisu w historii nowych produktów.',
        { name: name, inventoryRow: inventoryRow }
      );
      return { added: false, row: null };
    }
  }

  const auditRow = sheet.getLastRow() + 1;
  sheet.getRange(auditRow, 1, 1, 8).setValues([[
    new Date(),
    Session.getActiveUser().getEmail() || '',
    name,
    referenceName,
    productType,
    normalizeBusinessCategory_(category),
    inventoryRow,
    sourceInput
  ]]);
  return { added: true, row: auditRow };
}

function rollbackNewProductAudit_(productName, auditResult) {
  const result = auditResult || {};
  if (!result.added || !result.row) return { removed: false };
  const sheet = getOrCreateNewProductsSheet_();
  const range = sheet.getRange(Number(result.row), 1, 1, 8);
  const values = range.getDisplayValues()[0];
  if (normalizeText(values[2]) !== normalizeText(productName)) {
    const conflict =
      'Nie usunięto wpisu audytu w wierszu ' + result.row +
      ', ponieważ wskazuje już inny produkt.';
    logWarning(
      'ProductCreationService',
      'rollbackNewProductAudit_',
      conflict,
      { expectedProduct: productName, actualProduct: values[2] }
    );
    return { removed: false, conflict: conflict };
  }
  range.clearContent();
  return { removed: true, row: Number(result.row) };
}
