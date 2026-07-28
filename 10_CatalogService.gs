function getDictionarySheet_() {
  const sheet = getSheetByConfiguredName_(
    CONFIG.SHEETS.DICTIONARY
  );

  if (!sheet) {
    throw new Error(
      'Nie znaleziono arkusza: ' +
      CONFIG.SHEETS.DICTIONARY
    );
  }

  return sheet;
}

function detectDictionaryCodeContamination_(sheet) {
  const target = sheet || getDictionarySheet_();
  const lastRow = target.getLastRow();
  if (lastRow < 2) return [];
  const values = target.getRange(2, 2, lastRow - 1, 1).getDisplayValues();
  const signatures = [
    /^\s*\/\*\*?/, /^\s*\*\/?\s*$/, /^\s*function\s+[A-Za-z_$]/,
    /^\s*(const|let|var)\s+[A-Za-z_$]/, /SpreadsheetApp\./,
    /getRange\s*\(/, /setValues?\s*\(/, /setFormula/,
    /=>/, /throw\s+new\s+Error/, /return\s+[A-Za-z_$].*\(/,
    /^\s*[{}][;,]?\s*$/
  ];
  const issues = [];
  values.forEach((row, index) => {
    const value = String(row[0] || '').trim();
    if (!value) return;
    if (signatures.some(pattern => pattern.test(value))) {
      issues.push({ row: index + 2, value: value });
    }
  });
  return issues;
}

function assertDictionaryIsSafe_() {
  const sheet = getDictionarySheet_();
  const issues = detectDictionaryCodeContamination_(sheet);
  if (!issues.length) return true;
  const examples = issues.slice(0, 5).map(item => 'B' + item.row).join(', ');
  throw new Error(
    'SLOWNIK jest uszkodzony: wykryto fragmenty kodu w kolumnie B (' +
    issues.length + ' komórek; przykłady: ' + examples + '). ' +
    'Import został zatrzymany, aby nie zapisywać błędnych danych.'
  );
}

function loadAliases() {
  const sheet = getDictionarySheet_();
  assertDictionaryIsSafe_();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return {};
  }

  const values = sheet
    .getRange(2, 1, lastRow - 1, 2)
    .getValues();

  const result = {};

  values.forEach(row => {
    const alias = normalizeText(row[0]);
    const product = String(row[1] || '').trim();

    if (alias && product) {
      result[alias] = product;
    }
  });

  return result;
}

function loadProductConfigurations() {
  const sheet = getDictionarySheet_();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  return sheet
    .getRange(2, 4, lastRow - 1, 9)
    .getValues()
    .map((row, index) => ({
      dictionaryRow: index + 2,
      name: String(row[0] || '').trim(),
      normalizedName: normalizeText(row[0]),
      type: String(row[1] || 'NORMAL')
        .trim()
        .toUpperCase(),
      category: String(row[2] || '').trim(),
      columns: {
        quantity: String(row[3] || '')
          .trim()
          .toUpperCase(),
        weight: String(row[4] || '')
          .trim()
          .toUpperCase(),
        warehouse: String(row[5] || '')
          .trim()
          .toUpperCase(),
        darkroom: String(row[6] || '')
          .trim()
          .toUpperCase(),
        fridges: String(row[7] || '')
          .trim()
          .toUpperCase()
      },
      active: ['tak', 'true', '1'].includes(
        normalizeText(row[8])
      )
    }))
    .map(product => {
      product.columns = normalizeRuntimeProductColumns_(product.type, product.columns, product);
      return product;
    })
    .filter(product => product.name && product.active);
}

function saveAliasesBatch_(aliasSuggestions) {
  if (!Array.isArray(aliasSuggestions) || !aliasSuggestions.length) return 0;

  const sheet = getDictionarySheet_();
  const existing = loadAliases();
  const context = buildRuntimeContext_();
  const rows = [];
  const pending = {};

  aliasSuggestions.forEach(suggestion => {
    const validation = validateAliasSuggestion_(suggestion, context, existing);
    if (!validation.valid) {
      logWarning('Dictionary', 'saveAliasesBatch_', 'Alias pominiety: ' + validation.reason, suggestion);
      return;
    }

    const key = normalizeText(validation.alias);
    if (pending[key]) return;
    rows.push([validation.alias, validation.product]);
    pending[key] = true;
    existing[key] = validation.product;
  });

  if (!rows.length) return 0;
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 2).setValues(rows);
  invalidateProductCatalogCache_();
  return rows.length;
}

/**
 * Dodaje lub aktualizuje pelny rekord nowego produktu w arkuszu SLOWNIK.
 * Tabele aliasow (A:B) i konfiguracji (D:L) sa niezalezne, dlatego
 * wyszukujemy pierwszy pusty wiersz osobno dla kazdej z nich.
 */
function ensureNewProductInDictionary_(productData) {
  const data = productData || {};
  const name = String(data.name || '').trim();
  const productType = String(data.productType || '').trim().toUpperCase();
  const category = String(data.category || '').trim();
  const columns = data.columns || {};
  const sourceAlias = String(data.sourceAlias || '').trim();

  if (!name) throw new Error('Brak nazwy produktu do zapisu w slowniku.');

  const sheet = getDictionarySheet_();
  const normalizedName = normalizeText(name);
  const mutation = {
    configurationRow: null,
    configurationCreated: false,
    addedAliases: [],
    addedAliasRows: []
  };

  try {
    const configurations = loadProductConfigurations();
    let configuration = configurations.find(item => item.normalizedName === normalizedName);

    if (!configuration) {
      const configRow = findFirstEmptyDictionaryRow_(
        sheet,
        CONFIG.DICTIONARY.CONFIG_START_COLUMN,
        CONFIG.DICTIONARY.FIRST_DATA_ROW
      );

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
        configRow,
        CONFIG.DICTIONARY.CONFIG_START_COLUMN,
        1,
        CONFIG.DICTIONARY.CONFIG_COLUMN_COUNT
      ).setValues([row]);

      configuration = { dictionaryRow: configRow, name: name };
      mutation.configurationCreated = true;
    }
    mutation.configurationRow = configuration.dictionaryRow || null;

    const aliasesToEnsure = [name, sourceAlias]
      .map(value => String(value || '').trim())
      .filter(Boolean);

    const existingAliases = loadAliases();
    aliasesToEnsure.forEach(alias => {
      const normalizedAlias = normalizeText(alias);
      if (!normalizedAlias) return;

      const existingTarget = existingAliases[normalizedAlias];
      if (existingTarget) {
        if (normalizeText(existingTarget) !== normalizedName) {
          logWarning(
            'Dictionary',
            'ensureNewProductInDictionary_',
            'Nie nadpisano konfliktowego aliasu nowego produktu.',
            { alias: alias, existingProduct: existingTarget, newProduct: name }
          );
        }
        return;
      }

      const aliasRow = findFirstEmptyDictionaryRow_(
        sheet,
        CONFIG.DICTIONARY.ALIAS_COLUMN,
        CONFIG.DICTIONARY.FIRST_DATA_ROW
      );
      sheet.getRange(aliasRow, CONFIG.DICTIONARY.ALIAS_COLUMN, 1, 2)
        .setValues([[alias, name]]);
      existingAliases[normalizedAlias] = name;
      mutation.addedAliases.push(alias);
      mutation.addedAliasRows.push({ row: aliasRow, alias: alias, product: name });
    });

    SpreadsheetApp.flush();
    invalidateProductCatalogCache_();
    return mutation;
  } catch (error) {
    try {
      rollbackNewProductDictionaryEntry_(name, mutation);
    } catch (rollbackError) {
      logError(
        'Dictionary',
        'ensureNewProductInDictionary_.rollback',
        rollbackError,
        { product: name, mutation: mutation },
        0
      );
    }
    throw error;
  }
}

/**
 * Cofnięcie dotyczy wyłącznie rekordów utworzonych przez bieżącą operację.
 * Przed wyczyszczeniem każda komórka jest ponownie porównywana, dzięki czemu
 * ręczna zmiana wykonana równolegle nie zostanie nadpisana.
 */
function rollbackNewProductDictionaryEntry_(productName, dictionaryResult) {
  const result = dictionaryResult || {};
  if (!result.configurationCreated && !(result.addedAliasRows || []).length) {
    return { configurationRemoved: false, aliasesRemoved: 0, conflicts: [] };
  }

  const sheet = getDictionarySheet_();
  const targetKey = normalizeText(productName);
  const conflicts = [];
  let aliasesRemoved = 0;

  (result.addedAliasRows || []).slice().reverse().forEach(item => {
    const range = sheet.getRange(Number(item.row), CONFIG.DICTIONARY.ALIAS_COLUMN, 1, 2);
    const values = range.getDisplayValues()[0];
    if (
      normalizeText(values[0]) === normalizeText(item.alias) &&
      normalizeText(values[1]) === targetKey
    ) {
      range.clearContent();
      aliasesRemoved++;
    } else {
      conflicts.push('Alias w wierszu ' + item.row + ' został zmieniony po utworzeniu.');
    }
  });

  let configurationRemoved = false;
  if (result.configurationCreated && result.configurationRow) {
    const range = sheet.getRange(
      Number(result.configurationRow),
      CONFIG.DICTIONARY.CONFIG_START_COLUMN,
      1,
      CONFIG.DICTIONARY.CONFIG_COLUMN_COUNT
    );
    const values = range.getDisplayValues()[0];
    if (normalizeText(values[0]) === targetKey) {
      range.clearContent();
      configurationRemoved = true;
    } else {
      conflicts.push(
        'Konfiguracja w wierszu ' + result.configurationRow + ' została zmieniona po utworzeniu.'
      );
    }
  }

  SpreadsheetApp.flush();
  invalidateProductCatalogCache_();
  return {
    configurationRemoved: configurationRemoved,
    aliasesRemoved: aliasesRemoved,
    conflicts: conflicts
  };
}

function findFirstEmptyDictionaryRow_(sheet, column, firstDataRow) {
  const startRow = Math.max(Number(firstDataRow) || 2, 2);
  const lastRow = Math.max(sheet.getLastRow(), startRow);
  const values = sheet.getRange(startRow, column, lastRow - startRow + 1, 1).getDisplayValues();

  for (let index = 0; index < values.length; index++) {
    if (!String(values[index][0] || '').trim()) return startRow + index;
  }
  return lastRow + 1;
}
