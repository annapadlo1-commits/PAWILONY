/**
 * Inventory PRO 2.8 - Alias Manager
 * Centralne, bezpieczne zarzadzanie biblioteka aliasow.
 */

function showAliasManager() {
  const html = renderInventoryTemplate_('UI_AliasManager')
    .setWidth(1220)
    .setHeight(780);

  SpreadsheetApp.getUi().showModalDialog(html, '🍕 Alias Manager');
}

function getAliasManagerData() {
  return runSafely_(
    'AliasManager',
    'getAliasManagerData',
    function() {
      const sheet = getDictionarySheet_();
      const products = loadAllProductConfigurationsForManager_()
        .map(item => ({
          name: item.name,
          active: item.active,
          category: item.category,
          type: item.type
        }))
        .sort((a, b) => a.name.localeCompare(b.name, 'pl'));

      const productMap = {};
      products.forEach(product => {
        productMap[normalizeText(product.name)] = product;
      });

      const lastRow = sheet.getLastRow();
      const values = lastRow >= 2
        ? sheet.getRange(2, 1, lastRow - 1, 3).getValues()
        : [];

      const aliases = [];
      const grouped = {};

      values.forEach((row, index) => {
        const alias = String(row[0] || '').trim();
        const product = String(row[1] || '').trim();
        if (!alias && !product) return;

        const aliasKey = normalizeText(alias);
        const productKey = normalizeText(product);
        const target = productMap[productKey] || null;
        const record = {
          row: index + 2,
          alias: alias,
          normalizedAlias: aliasKey,
          product: product,
          productExists: Boolean(target),
          productActive: target ? target.active : false,
          category: target ? target.category : '',
          type: target ? target.type : '',
          isExactName: Boolean(target && aliasKey === productKey),
          conflict: false,
          conflictTargets: [],
          createdAt: row[2] instanceof Date && !Number.isNaN(row[2].getTime())
            ? row[2].toISOString()
            : ''
        };

        aliases.push(record);
        if (aliasKey) {
          if (!grouped[aliasKey]) grouped[aliasKey] = [];
          grouped[aliasKey].push(record);
        }
      });

      Object.keys(grouped).forEach(aliasKey => {
        const records = grouped[aliasKey];
        const targets = [];
        records.forEach(record => {
          const key = normalizeText(record.product);
          if (key && !targets.some(item => normalizeText(item) === key)) {
            targets.push(record.product);
          }
        });
        if (targets.length > 1) {
          records.forEach(record => {
            record.conflict = true;
            record.conflictTargets = targets.slice();
          });
        }
      });

      const lastAliasIndex = values.reduce((last, row, index) =>
        (String(row[0] || '').trim() || String(row[1] || '').trim()) ? index : last, -1);

      aliases.sort((a, b) => {
        if (a.conflict !== b.conflict) return a.conflict ? -1 : 1;
        if (a.productExists !== b.productExists) return a.productExists ? 1 : -1;
        return a.alias.localeCompare(b.alias, 'pl');
      });

      return {
        version: CONFIG.VERSION,
        aliases: aliases,
        products: products,
        summary: {
          aliases: aliases.length,
          uniqueAliases: Object.keys(grouped).length,
          conflicts: Object.keys(grouped).filter(key => {
            const targets = {};
            grouped[key].forEach(item => {
              const targetKey = normalizeText(item.product);
              if (targetKey) targets[targetKey] = true;
            });
            return Object.keys(targets).length > 1;
          }).length,
          orphaned: aliases.filter(item => !item.productExists).length,
          inactiveTargets: aliases.filter(item => item.productExists && !item.productActive).length,
          exactNameAliases: aliases.filter(item => item.isExactName).length,
          blankRows: Math.max(0, lastAliasIndex + 1 - aliases.length)
        }
      };
    },
    'Nie udalo sie wczytac biblioteki aliasow.'
  );
}

function addAliasFromAliasManager(payload) {
  return runSafely_(
    'AliasManager',
    'addAliasFromAliasManager',
    function() {
      const data = validateAliasManagerPayload_(payload, false);
      const existing = loadAliasRowsForManager_();
      const sameAlias = existing.filter(item => item.normalizedAlias === data.normalizedAlias);

      if (sameAlias.length) {
        const sameTarget = sameAlias.some(item =>
          normalizeText(item.product) === normalizeText(data.product)
        );
        if (sameTarget) {
          return { success: true, message: 'Ten alias juz wskazuje wybrany produkt.' };
        }
        throw new Error(
          'Alias jest juz przypisany do: ' +
          sameAlias.map(item => item.product).filter(Boolean).join(', ')
        );
      }

      const context = buildRuntimeContext_();
      const validation = validateAliasSuggestion_(
        { alias: data.alias, product: data.product },
        context,
        loadAliases()
      );
      if (!validation.valid) {
        throw new Error(validation.reason || 'Alias zostal odrzucony przez walidacje.');
      }

      const sheet = getDictionarySheet_();
      const row = findFirstEmptyDictionaryRow_(
        sheet,
        CONFIG.DICTIONARY.ALIAS_COLUMN,
        CONFIG.DICTIONARY.FIRST_DATA_ROW
      );
      sheet.getRange(row, 1, 1, 3).setValues([[data.alias, data.product, new Date()]]);
      invalidateProductCatalogCache_();

      logInfo('AliasManager', 'addAliasFromAliasManager', 'Dodano alias', data);
      return { success: true, message: 'Alias zostal dodany.', row: row };
    },
    'Nie udalo sie dodac aliasu.'
  );
}

function updateAliasFromAliasManager(payload) {
  return runSafely_(
    'AliasManager',
    'updateAliasFromAliasManager',
    function() {
      const data = validateAliasManagerPayload_(payload, true);
      const sheet = getDictionarySheet_();
      const existing = loadAliasRowsForManager_();
      const current = existing.find(item => item.row === data.row);
      if (!current) throw new Error('Nie znaleziono aliasu w wierszu ' + data.row + '.');

      const collision = existing.find(item =>
        item.row !== data.row &&
        item.normalizedAlias === data.normalizedAlias &&
        normalizeText(item.product) !== normalizeText(data.product)
      );
      if (collision) {
        throw new Error('Ten alias wskazuje juz inny produkt: ' + collision.product);
      }

      const duplicate = existing.find(item =>
        item.row !== data.row &&
        item.normalizedAlias === data.normalizedAlias &&
        normalizeText(item.product) === normalizeText(data.product)
      );
      if (duplicate) {
        throw new Error('Identyczny alias juz istnieje w wierszu ' + duplicate.row + '.');
      }

      const context = buildRuntimeContext_();
      const aliasesWithoutCurrent = loadAliases();
      delete aliasesWithoutCurrent[current.normalizedAlias];
      const validation = validateAliasSuggestion_(
        { alias: data.alias, product: data.product },
        context,
        aliasesWithoutCurrent
      );
      if (!validation.valid) {
        throw new Error(validation.reason || 'Alias zostal odrzucony przez walidacje.');
      }

      sheet.getRange(data.row, 1, 1, 2).setValues([[data.alias, data.product]]);
      invalidateProductCatalogCache_();

      logInfo('AliasManager', 'updateAliasFromAliasManager', 'Zmieniono alias', {
        row: data.row,
        oldAlias: current.alias,
        oldProduct: current.product,
        newAlias: data.alias,
        newProduct: data.product
      });
      return { success: true, message: 'Alias zostal zaktualizowany.' };
    },
    'Nie udalo sie zaktualizowac aliasu.'
  );
}

function deleteAliasFromAliasManager(row) {
  return runSafely_(
    'AliasManager',
    'deleteAliasFromAliasManager',
    function() {
      const rowNumber = Number(row);
      if (!Number.isInteger(rowNumber) || rowNumber < 2) {
        throw new Error('Nieprawidlowy numer wiersza aliasu.');
      }

      const sheet = getDictionarySheet_();
      const values = sheet.getRange(rowNumber, 1, 1, 2).getDisplayValues()[0];
      const alias = String(values[0] || '').trim();
      const product = String(values[1] || '').trim();
      if (!alias && !product) throw new Error('Alias jest juz pusty.');

      sheet.getRange(rowNumber, 1, 1, 3).clearContent();
      compactAliasDictionary_();
      invalidateProductCatalogCache_();

      logInfo('AliasManager', 'deleteAliasFromAliasManager', 'Usunieto alias', {
        row: rowNumber,
        alias: alias,
        product: product
      });
      return { success: true, message: 'Alias zostal usuniety.' };
    },
    'Nie udalo sie usunac aliasu.'
  );
}

function deleteDuplicateAliasesFromAliasManager() {
  return runSafely_(
    'AliasManager',
    'deleteDuplicateAliasesFromAliasManager',
    function() {
      const sheet = getDictionarySheet_();
      const rows = loadAliasRowsForManager_();
      const seen = {};
      const rowsToClear = [];

      rows.forEach(item => {
        const key = item.normalizedAlias + '|' + normalizeText(item.product);
        if (!item.normalizedAlias || !normalizeText(item.product)) return;
        if (seen[key]) rowsToClear.push(item.row);
        else seen[key] = item.row;
      });

      rowsToClear.forEach(row => sheet.getRange(row, 1, 1, 3).clearContent());
      compactAliasDictionary_();
      if (rowsToClear.length) invalidateProductCatalogCache_();

      return {
        success: true,
        removed: rowsToClear.length,
        message: rowsToClear.length
          ? 'Usunieto zduplikowane rekordy: ' + rowsToClear.length
          : 'Nie znaleziono identycznych duplikatow.'
      };
    },
    'Nie udalo sie usunac duplikatow aliasow.'
  );
}

function validateAliasManagerPayload_(payload, requireRow) {
  const data = payload || {};
  const alias = String(data.alias || '').trim();
  const product = String(data.product || '').trim();
  const normalizedAlias = normalizeText(alias);
  const products = loadAllProductConfigurationsForManager_();
  const target = products.find(item => item.normalizedName === normalizeText(product));

  if (!alias) throw new Error('Alias jest wymagany.');
  if (!normalizedAlias || normalizedAlias.length < 2) {
    throw new Error('Alias jest zbyt krotki.');
  }
  if (!product) throw new Error('Produkt docelowy jest wymagany.');
  if (!target) throw new Error('Produkt docelowy nie istnieje w katalogu: ' + product);

  const result = {
    alias: alias,
    normalizedAlias: normalizedAlias,
    product: target.name
  };

  if (requireRow) {
    const row = Number(data.row);
    if (!Number.isInteger(row) || row < 2) throw new Error('Nieprawidlowy wiersz aliasu.');
    result.row = row;
  }
  return result;
}

function loadAliasRowsForManager_() {
  const sheet = getDictionarySheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  return sheet.getRange(2, 1, lastRow - 1, 3).getValues()
    .map((row, index) => ({
      row: index + 2,
      alias: String(row[0] || '').trim(),
      normalizedAlias: normalizeText(row[0]),
      product: String(row[1] || '').trim(),
      createdAt: row[2] instanceof Date && !Number.isNaN(row[2].getTime())
        ? row[2].toISOString()
        : ''
    }))
    .filter(item => item.alias || item.product);
}


function compactAliasDictionary() {
  return runSafely_('AliasManager','compactAliasDictionary',function(){
    return compactAliasDictionary_('chronological');
  },'Nie udało się uporządkować aliasów.');
}

function organizeAliasDictionary(orderMode) {
  return runSafely_('AliasManager','organizeAliasDictionary',function(){
    return compactAliasDictionary_(orderMode || 'chronological');
  },'Nie udało się uporządkować aliasów.');
}

function compactAliasDictionary_(orderMode) {
  const sheet = getDictionarySheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { aliases:0, removedBlankRows:0 };
  const range = sheet.getRange(2,1,lastRow-1,3);
  const values = range.getValues();
  const records = values.map((row, index) => ({
    values: [row[0], row[1], row[2]],
    alias: String(row[0] || '').trim(),
    product: String(row[1] || '').trim(),
    createdAt: row[2] instanceof Date && !Number.isNaN(row[2].getTime()) ? row[2] : null,
    originalIndex: index
  })).filter(record => record.alias || record.product);
  const mode = ['chronological','product','alias'].includes(String(orderMode || '').toLowerCase())
    ? String(orderMode).toLowerCase()
    : 'chronological';
  const rows = sortAliasRecordsForManager_(records, mode).map(record => record.values);
  const removed = values.length - rows.length;
  range.clearContent();
  if (rows.length) sheet.getRange(2,1,rows.length,3).setValues(rows);
  if (!String(sheet.getRange(1,3).getValue() || '').trim()) sheet.getRange(1,3).setValue('Dodano');
  invalidateProductCatalogCache_();
  const labels = { chronological:'kolejność dodania', product:'produkt → alias', alias:'alias A–Z' };
  return { aliases: rows.length, removedBlankRows: removed, orderMode:mode,
    message:'Uporządkowano '+rows.length+' aliasów ('+labels[mode]+'). Usunięto pustych wierszy: '+removed+'.' };
}

function sortAliasRecordsForManager_(records, mode) {
  return (records || []).slice().sort((a, b) => {
    if (mode === 'product') {
      return a.product.localeCompare(b.product, 'pl', { sensitivity:'base' }) ||
        a.alias.localeCompare(b.alias, 'pl', { sensitivity:'base' }) ||
        a.originalIndex - b.originalIndex;
    }
    if (mode === 'alias') {
      return a.alias.localeCompare(b.alias, 'pl', { sensitivity:'base' }) ||
        a.product.localeCompare(b.product, 'pl', { sensitivity:'base' }) ||
        a.originalIndex - b.originalIndex;
    }
    if (!a.createdAt && !b.createdAt) return a.originalIndex - b.originalIndex;
    if (!a.createdAt) return -1;
    if (!b.createdAt) return 1;
    return a.createdAt.getTime() - b.createdAt.getTime() || a.originalIndex - b.originalIndex;
  });
}