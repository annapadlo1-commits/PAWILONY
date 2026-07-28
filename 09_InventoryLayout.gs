/** Inventory PRO 4.3 — jedno źródło prawdy dla układu INWENTURY. */


/** Produkty skonfigurowane jako bezpośredni stan końcowy (bez formuł). */
function isDirectFinalInventoryProduct_(product) {
  const name = normalizeText(product && product.name || product && product.product || '');
  const row = Number(product && product.inventoryRow) || 0;
  return getDirectFinalProductDefinitions_().some(definition =>
    getDirectFinalProductNames_(definition).indexOf(name) >= 0 ||
      (row > 0 && row === Number(definition && definition.inventoryRow))
  );
}

function getDirectFinalInventoryColumn_(product) {
  const name = normalizeText(product && product.name || product && product.product || '');
  const row = Number(product && product.inventoryRow) || 0;
  const definition = getDirectFinalProductDefinitions_().find(item =>
    getDirectFinalProductNames_(item).indexOf(name) >= 0 ||
      (row > 0 && row === Number(item && item.inventoryRow))
  );
  return normalizeColumnLetter_(definition && definition.column || '');
}

function getDirectFinalProductNames_(definition) {
  return [definition && definition.name].concat(definition && definition.aliases || [])
    .map(normalizeText)
    .filter(Boolean);
}

function getDirectFinalProductDefinitions_() {
  const configured = CONFIG.DIRECT_FINAL_PRODUCTS;
  return Array.isArray(configured) ? configured.slice() : [];
}


function getConfiguredInventoryLayout_(productType) {
  const type = String(productType || '').trim().toUpperCase();
  const configured = CONFIG.INVENTORY_LAYOUT && CONFIG.INVENTORY_LAYOUT[type];
  if (!configured) return null;
  const copy = {};
  Object.keys(configured).forEach(key => {
    copy[key] = Array.isArray(configured[key]) ? configured[key].slice() : configured[key];
  });
  return copy;
}

function getInputColumnsForProductType_(productType) {
  const type = String(productType || '').trim().toUpperCase();
  const layout = getConfiguredInventoryLayout_(type) || {};
  if (type === CONFIG.PRODUCT_TYPES.LOCATION) {
    return {
      quantity: '',
      weight: '',
      warehouse: normalizeColumnLetter_(layout.warehouse),
      darkroom: normalizeColumnLetter_(layout.darkroom),
      fridges: normalizeColumnLetter_(layout.fridges)
    };
  }
  return {
    quantity: normalizeColumnLetter_(layout.fullUnits),
    weight: normalizeColumnLetter_(layout.grossWeight),
    warehouse: '',
    darkroom: '',
    fridges: ''
  };
}

function getFormulaColumnsForProductType_(productType) {
  const layout = getConfiguredInventoryLayout_(productType) || {};
  return (layout.formulaColumns || []).map(normalizeColumnLetter_).filter(Boolean);
}

function getAllowedInputColumnsForProductType_(productType) {
  const columns = getInputColumnsForProductType_(productType);
  return Array.from(new Set([
    columns.quantity,
    columns.weight,
    columns.warehouse,
    columns.darkroom,
    columns.fridges
  ].map(normalizeColumnLetter_).filter(Boolean)));
}

function isFormulaColumnForProductType_(productType, column) {
  const wanted = normalizeColumnLetter_(column);
  return Boolean(wanted && getFormulaColumnsForProductType_(productType).indexOf(wanted) >= 0);
}

function isAllowedInputColumnForProductType_(productType, column) {
  const wanted = normalizeColumnLetter_(column);
  return Boolean(wanted && getAllowedInputColumnsForProductType_(productType).indexOf(wanted) >= 0);
}

function assertSafeInventoryTargetColumn_(product, column) {
  const directFinal = getDirectFinalInventoryColumn_(product);
  const wantedDirect = normalizeColumnLetter_(column);
  if (directFinal) {
    if (wantedDirect !== directFinal) {
      const productName = String(product && (product.name || product.product) || 'produkt finalny').trim();
      throw new Error('Produkt „' + productName + '” może być zapisywany wyłącznie do kolumny ' + directFinal + '.');
    }
    return directFinal;
  }
  const type = String(product && product.type || '').trim().toUpperCase();
  const wanted = normalizeColumnLetter_(column);
  const name = String(product && product.name || '').trim() || 'nieznany produkt';
  if (!wanted) throw new Error('Nie ustalono kolumny docelowej dla produktu „' + name + '”.');
  if (isFormulaColumnForProductType_(type, wanted)) {
    throw new Error(
      'Zablokowano zapis produktu „' + name + '” do kolumny obliczeniowej ' + wanted +
      ' (typ ' + type + '). Odbuduj konfigurację SŁOWNIKA.'
    );
  }
  if (!isAllowedInputColumnForProductType_(type, wanted)) {
    throw new Error(
      'Kolumna ' + wanted + ' nie jest dozwolonym polem wejściowym dla typu ' + type +
      ' i produktu „' + name + '”. Odbuduj konfigurację SŁOWNIKA.'
    );
  }
  return wanted;
}

function validateProductColumnMapping_(productType, columns, product) {
  const type = String(productType || '').trim().toUpperCase();
  const source = cloneProductColumns_(columns);
  if (isDirectFinalInventoryProduct_(product)) {
    const directColumn = getDirectFinalInventoryColumn_(product);
    const errors = [];
    if (source.quantity !== directColumn) {
      errors.push('Wyjątek finalny musi mieć kolumnę sztuk ' + directColumn + '.');
    }
    ['weight', 'warehouse', 'darkroom', 'fridges'].forEach(key => {
      if (source[key]) errors.push('Wyjątek finalny nie może używać pola „' + key + '”.');
    });
    return {
      valid: errors.length === 0,
      errors: errors,
      columns: source
    };
  }
  const errors = [];
  const used = {};
  const allowed = getAllowedInputColumnsForProductType_(type);
  const formulas = getFormulaColumnsForProductType_(type);

  Object.keys(source).forEach(key => {
    const column = normalizeColumnLetter_(source[key]);
    if (!column) return;
    if (formulas.indexOf(column) >= 0) {
      errors.push('Pole „' + key + '” wskazuje kolumnę formuły ' + column + '.');
    }
    if (allowed.indexOf(column) < 0) {
      errors.push('Pole „' + key + '” wskazuje niedozwoloną kolumnę ' + column + '.');
    }
    if (used[column]) {
      errors.push('Kolumna ' + column + ' została przypisana jednocześnie do pól „' + used[column] + '” i „' + key + '”.');
    }
    used[column] = key;
  });

  if (type === CONFIG.PRODUCT_TYPES.LOCATION) {
    if (!source.warehouse && !source.darkroom && !source.fridges) {
      errors.push('Produkt LOCATION nie ma żadnej kolumny lokalizacji.');
    }
    if (source.quantity || source.weight) {
      errors.push('Produkt LOCATION nie może korzystać z pól ilości/wagi.');
    }
  } else {
    if (!source.quantity && !source.weight) {
      errors.push('Produkt ' + type + ' nie ma kolumny sztuk ani wagi.');
    }
    if (source.warehouse || source.darkroom || source.fridges) {
      errors.push('Produkt ' + type + ' nie może korzystać z kolumn lokalizacji.');
    }
  }

  return { valid: errors.length === 0, errors: errors, columns: source };
}

/**
 * Starsze rekordy SŁOWNIKA mogą zawierać układ kolumn z innego lokalu.
 * W czasie działania nigdy nie ufamy takiemu mapowaniu: poprawny rekord
 * zachowujemy, a niepoprawny zastępujemy układem z CONFIG bieżącego lokalu.
 */
function normalizeRuntimeProductColumns_(productType, columns, product) {
  const type = String(productType || '').trim().toUpperCase();
  if (isDirectFinalInventoryProduct_(product)) {
    return {
      quantity: getDirectFinalInventoryColumn_(product),
      weight: '',
      warehouse: '',
      darkroom: '',
      fridges: ''
    };
  }
  const validation = validateProductColumnMapping_(type, columns, product);
  return validation.valid
    ? validation.columns
    : cloneProductColumns_(getInputColumnsForProductType_(type));
}

function cloneProductColumns_(columns) {
  const source = columns || {};
  return {
    quantity: normalizeColumnLetter_(source.quantity),
    weight: normalizeColumnLetter_(source.weight),
    warehouse: normalizeColumnLetter_(source.warehouse),
    darkroom: normalizeColumnLetter_(source.darkroom),
    fridges: normalizeColumnLetter_(source.fridges)
  };
}

function mergeDetectedProductColumns_(baseColumns, detectedColumns) {
  const merged = cloneProductColumns_(baseColumns);
  const detected = cloneProductColumns_(detectedColumns);
  Object.keys(merged).forEach(key => {
    if (detected[key]) merged[key] = detected[key];
  });
  return merged;
}

function detectInventoryInputColumnsFromHeaderRow_(rowValues, productType, fallbackColumns) {
  const type = String(productType || '').trim().toUpperCase();
  const detected = cloneProductColumns_(fallbackColumns || getInputColumnsForProductType_(type));
  const values = rowValues || [];

  if (type === CONFIG.PRODUCT_TYPES.LOCATION) {
    values.forEach((value, index) => {
      const area = resolveLocationHeaderArea_(value);
      if (area && area.columnKey) detected[area.columnKey] = columnNumberToLetter_(index + 1);
    });
    return detected;
  }

  values.forEach((value, index) => {
    const normalized = normalizeText(value || '');
    if (!normalized) return;
    const column = columnNumberToLetter_(index + 1);
    if (isGrossInventoryHeader_(normalized)) detected.weight = column;
    if (isFullUnitsInventoryHeader_(normalized, type)) detected.quantity = column;
  });
  return detected;
}

function isInventoryHeaderContinuationRow_(rowValues, productType) {
  const values = rowValues || [];
  const firstCell = String(values[0] || '').trim();
  if (firstCell) return false;
  const type = String(productType || '').trim().toUpperCase();
  if (type === CONFIG.PRODUCT_TYPES.LOCATION) {
    return values.some(value => Boolean(resolveLocationHeaderArea_(value)));
  }
  return values.some(value => {
    const normalized = normalizeText(value || '');
    return isGrossInventoryHeader_(normalized) || isFullUnitsInventoryHeader_(normalized, type);
  });
}

function isGrossInventoryHeader_(normalizedHeader) {
  const value = String(normalizedHeader || '');
  return value.indexOf('waga szt w butelce') >= 0 ||
    value.indexOf('waga brutto') >= 0 ||
    value.indexOf('waga w kegu') >= 0;
}

function isFullUnitsInventoryHeader_(normalizedHeader, productType) {
  const value = String(normalizedHeader || '');
  if (String(productType || '').toUpperCase() === CONFIG.PRODUCT_TYPES.KEG) {
    return value.indexOf('pelne kegi') >= 0 || value.indexOf('pelne keg') >= 0;
  }
  return value.indexOf('pelne btlk szt') >= 0 ||
    value.indexOf('pelne butelki szt') >= 0 ||
    value.indexOf('pelne szt') >= 0;
}

function inferInventoryProductType_(sectionType, category, productName, rowValues) {
  const normalizedName = normalizeText(productName || '');
  const normalizedCategory = normalizeText(category || '');
  const currentType = String(sectionType || CONFIG.PRODUCT_TYPES.NORMAL).toUpperCase();

  if (currentType === CONFIG.PRODUCT_TYPES.LOCATION) return CONFIG.PRODUCT_TYPES.LOCATION;
  if (currentType === CONFIG.PRODUCT_TYPES.KEG) return CONFIG.PRODUCT_TYPES.KEG;
  if (normalizedCategory === 'piwo keg') return CONFIG.PRODUCT_TYPES.KEG;
  if (normalizedCategory === 'piwo butelki' || normalizedCategory === 'softy') {
    return CONFIG.PRODUCT_TYPES.LOCATION;
  }
  // W PAWILONACH słowo „KEG” może być częścią nazwy produktu liczonego
  // standardowo (np. wino 18 l). Typ zmienia wyłącznie fizyczna sekcja lub
  // awaryjna, ogólna sekcja PIWO.
  if (normalizedCategory === 'piwo' && (
      /(^|\s)keg(\s|$)/.test(normalizedName) || isLikelyKegInventoryRow_(rowValues)
    )) {
    return CONFIG.PRODUCT_TYPES.KEG;
  }
  return CONFIG.PRODUCT_TYPES.NORMAL;
}

function isLikelyKegInventoryRow_(rowValues) {
  if (!Array.isArray(rowValues) || !rowValues.length) return false;
  const layout = getConfiguredInventoryLayout_(CONFIG.PRODUCT_TYPES.KEG) || {};
  const capacityIndex = inventoryColumnLetterToNumber_(layout.unitCapacity) - 1;
  if (capacityIndex < 0 || capacityIndex >= rowValues.length) return false;
  const raw = String(rowValues[capacityIndex] || '').trim().replace(',', '.');
  const capacity = Number(raw);
  return Number.isFinite(capacity) && capacity > 5;
}

function getInventoryLayoutMaxColumn_() {
  let maximum = 1;
  Object.keys(CONFIG.INVENTORY_LAYOUT || {}).forEach(type => {
    const layout = CONFIG.INVENTORY_LAYOUT[type] || {};
    Object.keys(layout).forEach(key => {
      if (key === 'formulaColumns' || key === 'unit') return;
      maximum = Math.max(maximum, inventoryColumnLetterToNumber_(layout[key]));
    });
    (layout.formulaColumns || []).forEach(column => {
      maximum = Math.max(maximum, inventoryColumnLetterToNumber_(column));
    });
  });
  return maximum;
}

function inventoryColumnLetterToNumber_(letters) {
  return String(letters || '').toUpperCase().split('').reduce((total, character) => {
    const code = character.charCodeAt(0) - 64;
    return code >= 1 && code <= 26 ? total * 26 + code : total;
  }, 0);
}

function columnNumberToLetter_(columnNumber) {
  let number = Number(columnNumber) || 0;
  let result = '';
  while (number > 0) {
    const remainder = (number - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    number = Math.floor((number - 1) / 26);
  }
  return result;
}
