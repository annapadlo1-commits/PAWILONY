/**
 * Inventory PRO Enterprise v2.1.3 Recovery
 * Jedno wczytanie katalogu na cala analize.
 */

function analyzeImportText(inputText) {
  const startedAt = Date.now();
  const text = String(inputText || '').trim();

  if (!text) {
    throw new Error('Wklej tekst inwentaryzacji.');
  }

  const contextStartedAt = Date.now();
  const runtimeContext = buildRuntimeContext_();
  const contextReadyAt = Date.now();
  const qualitySettings = loadQualitySettings_();
  const parseStartedAt = Date.now();
  const parsedItems = parseInventoryText(text, runtimeContext);
  const parsedAt = Date.now();

  const items = parsedItems.map((parsedItem, index) => {
    if (
      parsedItem.status !== 'OK' ||
      !parsedItem.product ||
      parsedItem.value === null
    ) {
      return {
        id: index + 1,
        originalInput: parsedItem.originalInput || '',
        parsedProduct: parsedItem.product || '',
        value: parsedItem.value,
        status: 'PARSE_ERROR',
        message: parsedItem.message || 'Nie rozpoznano wpisu',
        selectedProduct: '',
        productType: '',
        category: '',
        location: parsedItem.location || '',
        requiresLocation: false,
        candidates: [],
        include: false,
        qualityWarning: false,
        qualityLevel: 'ERROR',
        learnAlias: false
      };
    }

    const parserMatch = matchProductForParser_(parsedItem.product, runtimeContext);
    const match = parserMatch.match;

    if (match.matched && match.product) {
      return createReadyImportItem_(
        index,
        parsedItem,
        match,
        qualitySettings
      );
    }

    if (
      match.status === 'AMBIGUOUS' &&
      Array.isArray(match.candidates)
    ) {
      return {
        id: index + 1,
        originalInput: parsedItem.originalInput,
        parsedProduct: parsedItem.product,
        value: parsedItem.value,
        status: 'AMBIGUOUS',
        message: 'Wybierz produkt z listy',
        selectedProduct: '',
        productType: '',
        category: '',
        location: parsedItem.location || '',
        requiresLocation: false,
        qualityWarning: false,
        qualityLevel: 'OK',
        learnAlias: false,
        candidates: match.candidates.map(candidate => ({
          name: candidate.product.name,
          type: candidate.product.type,
          category: candidate.product.category,
          score: candidate.score
        })),
        include: true
      };
    }

    return {
      id: index + 1,
      originalInput: parsedItem.originalInput,
      parsedProduct: parsedItem.product,
      value: parsedItem.value,
      status: 'NOT_FOUND',
      message: match.message || 'Nie znaleziono produktu',
      selectedProduct: '',
      productType: '',
      category: '',
      location: parsedItem.location || '',
      requiresLocation: false,
      candidates: [],
      include: false,
      qualityWarning: false,
      qualityLevel: 'ERROR',
      learnAlias: false
    };
  });

  annotatePreviousInventoryValues_(items, runtimeContext);
  annotatePreviewDuplicates_(items, qualitySettings);
  const outputItems = CONFIG.REVIEW && CONFIG.REVIEW.AUTO_MERGE_DUPLICATES
    ? mergePreviewDuplicates_(items)
    : items;
  outputItems.forEach((item, index) => item.id = index + 1);
  const completedAt = Date.now();

  return {
    success: true,
    itemCount: outputItems.length,
    sourceItemCount: items.length,
    readyCount: outputItems.filter(item =>
      ['EXACT', 'ALIAS', 'VARIANT', 'SMART', 'AUTO'].includes(item.status)
    ).length,
    ambiguousCount: outputItems.filter(
      item => item.status === 'AMBIGUOUS'
    ).length,
    errorCount: outputItems.filter(item =>
      ['PARSE_ERROR', 'NOT_FOUND'].includes(item.status)
    ).length,
    locationCount: outputItems.filter(
      item => item.requiresLocation
    ).length,
    durationMs: Date.now() - startedAt,
    performance: {
      contextMs: contextReadyAt - contextStartedAt,
      parseMs: parsedAt - parseStartedAt,
      enrichMs: completedAt - parsedAt,
      totalMs: completedAt - startedAt,
      matcher: runtimeContext.performanceStats || {}
    },
    duplicateGroupCount: outputItems.filter(item => item.autoMerged).length || countDuplicateGroups_(items),
    qualitySettings: qualitySettings,
    resolverData: buildProductResolverPayload_(runtimeContext, ''),
    items: outputItems
  };
}

/**
 * Szybka inwentaryzacja nie ponownie zgaduje produktu z tekstu.
 * Użytkownik wybiera rekord katalogu, a serwer nadal waliduje wszystkie pola
 * i zwraca dokładnie ten sam model podglądu co zwykły parser.
 */
function analyzeQuickInventoryItems(entries) {
  if (!Array.isArray(entries) || !entries.length) {
    throw new Error('Dodaj co najmniej jedną pozycję szybkiej inwentaryzacji.');
  }
  const context = buildRuntimeContext_();
  const qualitySettings = loadQualitySettings_();
  const items = entries.map(function(entry, index) {
    const name = String(entry && entry.productName || '').trim();
    const product = context.productIndex[normalizeText(name)];
    const value = Number(entry && entry.value);
    const location = String(entry && entry.location || '').trim();
    if (!product) throw new Error('Nie znaleziono produktu w katalogu: ' + name);
    if (!Number.isFinite(value) || value < 0) {
      throw new Error('Nieprawidłowa wartość dla produktu: ' + product.name);
    }
    if (product.type === CONFIG.PRODUCT_TYPES.LOCATION && !location) {
      throw new Error('Wybierz lokalizację dla produktu: ' + product.name);
    }
    const parsedItem = {
      originalInput: product.name + ' ' + value + (location ? ' ' + location : ''),
      product: product.name,
      value: value,
      location: location
    };
    return createReadyImportItem_(index, parsedItem, {
      matched: true,
      product: product,
      status: 'MANUAL',
      message: 'Wybrano z katalogu'
    }, qualitySettings);
  });
  annotatePreviousInventoryValues_(items, context);
  annotatePreviewDuplicates_(items, qualitySettings);
  const output = CONFIG.REVIEW && CONFIG.REVIEW.AUTO_MERGE_DUPLICATES
    ? mergePreviewDuplicates_(items) : items;
  output.forEach(function(item, index) { item.id = index + 1; });
  return { success: true, itemCount: output.length, items: output };
}

function annotatePreviousInventoryValues_(items, runtimeContext) {
  const context = runtimeContext || buildRuntimeContext_();
  const sheet = getSheetByConfiguredName_(CONFIG.SHEETS.INVENTORY);
  if (!sheet || !items || !items.length) return;
  const productIndex = context.productIndex || {};
  const requests = [];
  const columns = {};

  items.forEach((item, index) => {
    if (!item.selectedProduct || !Number.isFinite(Number(item.value))) return;
    const product = productIndex[normalizeText(item.selectedProduct)];
    if (!product || !product.inventoryRow) return;
    const column = resolveTargetColumn_(product, Number(item.value), item.location || '');
    if (!column) return;
    columns[column] = true;
    requests.push({ index: index, row: product.inventoryRow, column: column });
  });

  const buffers = {};
  const lastRow = sheet.getLastRow();
  Object.keys(columns).forEach(column => {
    buffers[column] = sheet.getRange(column + '1:' + column + lastRow).getValues();
  });

  requests.forEach(request => {
    const raw = buffers[request.column] && buffers[request.column][request.row - 1]
      ? buffers[request.column][request.row - 1][0]
      : null;
    if (raw === '' || raw === null || raw === undefined) return;
    const previous = Number(raw);
    if (!Number.isFinite(previous)) return;
    const item = items[request.index];
    const current = Number(item.value);
    const delta = current - previous;
    const percent = previous === 0 ? null : (delta / Math.abs(previous)) * 100;
    item.previousValue = previous;
    item.previousDelta = delta;
    item.previousChangePercent = percent;
    const largeAbsolute = Math.abs(delta) >= Number(CONFIG.REVIEW.PREVIOUS_CHANGE_WARNING_ABSOLUTE || 10);
    const largePercent = percent !== null && Math.abs(percent) >= Number(CONFIG.REVIEW.PREVIOUS_CHANGE_WARNING_PERCENT || 250);
    item.previousValueWarning = Boolean(largeAbsolute && (largePercent || previous === 0));
    if (item.previousValueWarning) {
      item.qualityWarning = true;
      item.qualityLevel = item.qualityLevel === 'ERROR' ? 'ERROR' : 'WARNING';
      item.qualityFlags = (item.qualityFlags || []).concat(['DUZA_ZMIANA_WZGLEDEM_POPRZEDNIEJ']);
      item.message = [item.message, 'Duża zmiana: było ' + previous + ', jest ' + current]
        .filter(Boolean).join(' | ');
    }
  });
}

function mergePreviewDuplicates_(items) {
  const groups = {};
  const groupedIndexes = {};
  (items || []).forEach((item, index) => {
    if (!item.include || !item.selectedProduct || !Number.isFinite(Number(item.value)) ||
        ['PARSE_ERROR', 'NOT_FOUND', 'ERROR', 'AMBIGUOUS'].includes(item.status)) return;
    const key = buildPreviewDuplicateKey_(item);
    if (!groups[key]) groups[key] = [];
    groups[key].push({ item: item, index: index });
  });

  const mergedByFirstIndex = {};
  Object.keys(groups).forEach(key => {
    const group = groups[key];
    if (group.length < 2) return;
    const sourceItems = group.map(entry => JSON.parse(JSON.stringify(entry.item)));
    const merged = JSON.parse(JSON.stringify(group[0].item));
    merged.value = sourceItems.reduce((sum, item) => sum + Number(item.value), 0);
    merged.originalInput = sourceItems.map(item => item.originalInput).join(' + ');
    merged.autoMerged = true;
    merged.sourceItems = sourceItems;
    merged.duplicateCount = sourceItems.length;
    merged.duplicateValues = sourceItems.map(item => Number(item.value));
    merged.duplicateInputs = sourceItems.map(item => item.originalInput || item.parsedProduct || '');
    merged.duplicateTotal = merged.value;
    merged.duplicateWarning = true;
    if (Number.isFinite(Number(merged.previousValue))) {
      merged.previousValue = Number(merged.previousValue);
      merged.previousDelta = merged.value - merged.previousValue;
      merged.previousChangePercent = merged.previousValue === 0
        ? null
        : (merged.previousDelta / Math.abs(merged.previousValue)) * 100;
      const largeAbsolute = Math.abs(merged.previousDelta) >= Number(CONFIG.REVIEW.PREVIOUS_CHANGE_WARNING_ABSOLUTE || 10);
      const largePercent = merged.previousChangePercent !== null &&
        Math.abs(merged.previousChangePercent) >= Number(CONFIG.REVIEW.PREVIOUS_CHANGE_WARNING_PERCENT || 250);
      merged.previousValueWarning = Boolean(largeAbsolute && (largePercent || merged.previousValue === 0));
    }
    merged.baseMessage = 'Połączono automatycznie ' + sourceItems.length + ' wpisy';
    merged.message = merged.baseMessage;
    mergedByFirstIndex[group[0].index] = merged;
    group.forEach(entry => groupedIndexes[entry.index] = true);
  });

  const output = [];
  (items || []).forEach((item, index) => {
    if (mergedByFirstIndex[index]) output.push(mergedByFirstIndex[index]);
    else if (!groupedIndexes[index]) output.push(item);
  });
  return output;
}

function createReadyImportItem_(
  index,
  parsedItem,
  match,
  qualitySettings
) {
  const product = match.product;
  const quality = evaluateImportQuality_(
    product,
    Number(parsedItem.value),
    parsedItem.location || '',
    qualitySettings
  );

  return {
    id: index + 1,
    originalInput: parsedItem.originalInput,
    parsedProduct: parsedItem.product,
    value: parsedItem.value,
    status: match.status,
    message: quality.warning
      ? quality.message
      : match.message,
    selectedProduct: product.name,
    productType: product.type,
    category: product.category,
    location: parsedItem.location || '',
    requiresLocation:
      product.type === CONFIG.PRODUCT_TYPES.LOCATION,
    candidates: [],
    include: true,
    qualityWarning: quality.warning,
    qualityLevel: quality.level,
    qualityFlags: quality.flags || [],
    duplicateCount: 1,
    duplicateValues: [parsedItem.value],
    duplicateTotal: parsedItem.value,
    duplicateWarning: false,
    learnAlias: false
  };
}


function annotatePreviewDuplicates_(items, qualitySettings) {
  const groups = {};
  const threshold = qualitySettings && qualitySettings.duplicateWarningCount
    ? qualitySettings.duplicateWarningCount
    : 2;

  items.forEach((item, index) => {
    if (!item.include || !item.selectedProduct || !Number.isFinite(Number(item.value))) return;
    const key = buildPreviewDuplicateKey_(item);
    if (!groups[key]) groups[key] = [];
    groups[key].push(index);
  });

  Object.keys(groups).forEach(key => {
    const indexes = groups[key];
    const values = indexes.map(index => Number(items[index].value));
    const total = values.reduce((sum, value) => sum + value, 0);
    const isDuplicate = indexes.length >= threshold;

    indexes.forEach(index => {
      const item = items[index];
      item.duplicateCount = indexes.length;
      item.duplicateValues = values;
      item.duplicateInputs = indexes.map(i => items[i].originalInput || items[i].parsedProduct || '');
      item.duplicateTotal = total;
      item.duplicateWarning = isDuplicate;
      if (isDuplicate) {
        const duplicateMessage = 'Duplikat: ' + indexes.length + ' wpisy, suma ' + total;
        item.message = item.message ? item.message + ' | ' + duplicateMessage : duplicateMessage;
      }
    });
  });
}

function buildPreviewDuplicateKey_(item) {
  const value = Number(item.value);
  const valueMode = Number.isInteger(value) ? 'WHOLE' : 'WEIGHT';
  return [
    normalizeText(item.selectedProduct),
    String(item.productType || ''),
    normalizeText(item.location || ''),
    valueMode
  ].join('|');
}

function countDuplicateGroups_(items) {
  const seen = {};
  items.forEach(item => {
    if (!item.duplicateWarning) return;
    seen[buildPreviewDuplicateKey_(item)] = true;
  });
  return Object.keys(seen).length;
}


/** v2.9.2: ponowna analiza pojedynczego wiersza Smart Review. */
function reanalyzeSingleImportItem(inputText) {
  const result = analyzeImportText(inputText);
  if (!result.items || result.items.length !== 1) {
    throw new Error('Wpis powinien zawierać dokładnie jedną pozycję. Rozpoznano: ' + (result.items ? result.items.length : 0));
  }
  return result.items[0];
}
