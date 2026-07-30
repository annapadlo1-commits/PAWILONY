/**
 * Inventory PRO 4.7.0 — profile opakowań produktów.
 * Dane techniczne są oddzielone od SŁOWNIKA, aby nie zmieniać jego kontraktu.
 */
const PRODUCT_PACKAGING_HEADERS_ = [
  'KLUCZ PRODUKTU', 'PRODUKT', 'TRYB', 'TARA',
  'WAGA REFERENCYJNA BRUTTO', 'ILOŚĆ POCZĄTKOWA',
  'PRZELICZNIK MASY', 'DATA REFERENCJI', 'AKTUALIZACJA', 'UŻYTKOWNIK'
];
let PRODUCT_PACKAGING_RUNTIME_CACHE_ = null;
const UNKNOWN_TARE_GROSS_APPROVAL_PROPERTY_ = 'INVENTORY_UNKNOWN_TARE_GROSS_APPROVALS';
let UNKNOWN_TARE_GROSS_APPROVAL_RUNTIME_CACHE_ = null;

function loadUnknownTareGrossApprovals_() {
  if (UNKNOWN_TARE_GROSS_APPROVAL_RUNTIME_CACHE_) {
    return UNKNOWN_TARE_GROSS_APPROVAL_RUNTIME_CACHE_;
  }
  const session = ensureActiveInventorySession_();
  const properties = PropertiesService.getDocumentProperties();
  let stored = {};
  try { stored = JSON.parse(properties.getProperty(UNKNOWN_TARE_GROSS_APPROVAL_PROPERTY_) || '{}'); }
  catch (error) { stored = {}; }
  UNKNOWN_TARE_GROSS_APPROVAL_RUNTIME_CACHE_ = stored.sessionId === session.id && stored.products
    ? stored
    : {sessionId: session.id, products: {}};
  return UNKNOWN_TARE_GROSS_APPROVAL_RUNTIME_CACHE_;
}

function isUnknownTareGrossApproved_(product) {
  const key = normalizeText(product && product.name);
  return Boolean(key && loadUnknownTareGrossApprovals_().products[key]);
}

function saveUnknownTareGrossApproval_(product, approved) {
  const state = loadUnknownTareGrossApprovals_();
  const key = normalizeText(product && product.name);
  if (!key) throw new Error('Brak produktu dla decyzji o tarze.');
  if (approved) state.products[key] = true;
  else delete state.products[key];
  PropertiesService.getDocumentProperties().setProperty(
    UNKNOWN_TARE_GROSS_APPROVAL_PROPERTY_, JSON.stringify(state)
  );
  UNKNOWN_TARE_GROSS_APPROVAL_RUNTIME_CACHE_ = state;
}

function clearUnknownTareGrossApprovals_() {
  PropertiesService.getDocumentProperties().deleteProperty(UNKNOWN_TARE_GROSS_APPROVAL_PROPERTY_);
  UNKNOWN_TARE_GROSS_APPROVAL_RUNTIME_CACHE_ = null;
}

function getProductPackagingSheet_(createIfMissing) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.PRODUCT_PACKAGING);
  if (!sheet && createIfMissing) {
    sheet = spreadsheet.insertSheet(CONFIG.SHEETS.PRODUCT_PACKAGING);
    sheet.getRange(1, 1, 1, PRODUCT_PACKAGING_HEADERS_.length)
      .setValues([PRODUCT_PACKAGING_HEADERS_])
      .setFontWeight('bold')
      .setBackground('#f4cccc');
    sheet.setFrozenRows(1);
    sheet.hideSheet();
  }
  return sheet;
}

function invalidateProductPackagingCache_() {
  PRODUCT_PACKAGING_RUNTIME_CACHE_ = null;
}

function normalizePackagingNumber_(value) {
  if (value === '' || value === null || value === undefined) return '';
  const number = Number(String(value).trim().replace(',', '.'));
  return Number.isFinite(number) ? number : NaN;
}

function validateProductPackagingProfile_(profile, product) {
  const source = profile || {};
  const type = String(product && product.type || '').trim().toUpperCase();
  if (type === CONFIG.PRODUCT_TYPES.LOCATION || isDirectFinalInventoryProduct_(product)) {
    return { mode: '', emptyWeight: '', referenceGrossWeight: '', referenceNetQuantity: '', massConversionFactor: '', referenceDate: '' };
  }

  const validModes = Object.keys(CONFIG.PACKAGING_MODES).map(key => CONFIG.PACKAGING_MODES[key]);
  const mode = String(source.mode || CONFIG.PACKAGING_MODES.TARE_UNKNOWN).trim().toUpperCase();
  if (!validModes.includes(mode)) throw new Error('Nieprawidłowy tryb opakowania: ' + mode + '.');

  const result = {
    mode: mode,
    emptyWeight: normalizePackagingNumber_(source.emptyWeight),
    referenceGrossWeight: normalizePackagingNumber_(source.referenceGrossWeight),
    referenceNetQuantity: normalizePackagingNumber_(source.referenceNetQuantity),
    massConversionFactor: normalizePackagingNumber_(source.massConversionFactor),
    referenceDate: String(source.referenceDate || '').trim()
  };

  if (mode === CONFIG.PACKAGING_MODES.TARE_KNOWN) {
    if (!Number.isFinite(result.emptyWeight) || result.emptyWeight <= 0) {
      throw new Error('Dla trybu „tara znana” podaj dodatnią wagę pustej butelki lub kega.');
    }
    result.referenceGrossWeight = '';
    result.referenceNetQuantity = '';
    result.massConversionFactor = '';
    result.referenceDate = '';
  } else if (mode === CONFIG.PACKAGING_MODES.TARE_UNKNOWN) {
    result.emptyWeight = '';
    result.referenceGrossWeight = '';
    result.referenceNetQuantity = '';
    result.massConversionFactor = '';
    result.referenceDate = '';
  } else {
    if (!Number.isFinite(result.referenceGrossWeight) || result.referenceGrossWeight <= 0) {
      throw new Error('Podaj dodatnią początkową wagę opakowania razem z produktem.');
    }
    if (!Number.isFinite(result.referenceNetQuantity) || result.referenceNetQuantity < 0) {
      throw new Error('Podaj początkową ilość produktu.');
    }
    if (result.massConversionFactor === '') result.massConversionFactor = 1;
    if (!Number.isFinite(result.massConversionFactor) || result.massConversionFactor <= 0) {
      throw new Error('Przelicznik masy musi być dodatni.');
    }
    result.emptyWeight = '';
    if (!result.referenceDate) result.referenceDate = new Date().toISOString().slice(0, 10);
  }
  return result;
}

function loadProductPackagingProfiles_() {
  if (PRODUCT_PACKAGING_RUNTIME_CACHE_) return PRODUCT_PACKAGING_RUNTIME_CACHE_;
  const sheet = getProductPackagingSheet_(false);
  const map = {};
  if (sheet && sheet.getLastRow() >= 2) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, PRODUCT_PACKAGING_HEADERS_.length)
      .getValues()
      .forEach((row, index) => {
        const key = normalizeText(row[0] || row[1]);
        if (!key) return;
        map[key] = {
          row: index + 2,
          productName: String(row[1] || '').trim(),
          mode: String(row[2] || '').trim().toUpperCase(),
          emptyWeight: row[3] === '' ? '' : Number(row[3]),
          referenceGrossWeight: row[4] === '' ? '' : Number(row[4]),
          referenceNetQuantity: row[5] === '' ? '' : Number(row[5]),
          massConversionFactor: row[6] === '' ? '' : Number(row[6]),
          referenceDate: row[7] instanceof Date
            ? Utilities.formatDate(row[7], Session.getScriptTimeZone(), 'yyyy-MM-dd')
            : String(row[7] || '').trim()
        };
      });
  }
  PRODUCT_PACKAGING_RUNTIME_CACHE_ = map;
  return map;
}

function getProductPackagingProfile_(product) {
  const type = String(product && product.type || '').trim().toUpperCase();
  if (!product || type === CONFIG.PRODUCT_TYPES.LOCATION || isDirectFinalInventoryProduct_(product)) return null;
  if (product.packaging && product.packaging.mode) return Object.assign({}, product.packaging);
  const stored = loadProductPackagingProfiles_()[normalizeText(product.name)];
  if (stored) return Object.assign({}, stored);
  return {
    productName: product.name,
    mode: CONFIG.PACKAGING_MODES.TARE_KNOWN,
    emptyWeight: '',
    referenceGrossWeight: '',
    referenceNetQuantity: '',
    massConversionFactor: '',
    referenceDate: '',
    inheritedFromInventory: true
  };
}

function saveProductPackagingProfile_(product, profile, originalName) {
  const validated = validateProductPackagingProfile_(profile, product);
  if (!validated.mode) return {saved: false, profile: validated};
  const sheet = getProductPackagingSheet_(true);
  const profiles = loadProductPackagingProfiles_();
  const oldKey = normalizeText(originalName || product.name);
  const newKey = normalizeText(product.name);
  const existing = profiles[oldKey] || profiles[newKey];
  const row = existing && existing.row ? existing.row : Math.max(2, sheet.getLastRow() + 1);
  sheet.getRange(row, 1, 1, PRODUCT_PACKAGING_HEADERS_.length).setValues([[
    newKey,
    product.name,
    validated.mode,
    validated.emptyWeight,
    validated.referenceGrossWeight,
    validated.referenceNetQuantity,
    validated.massConversionFactor,
    validated.referenceDate,
    new Date(),
    getCurrentUserEmail_()
  ]]);
  if (oldKey !== newKey && profiles[oldKey] && profiles[oldKey].row !== row) {
    sheet.getRange(profiles[oldKey].row, 1, 1, PRODUCT_PACKAGING_HEADERS_.length).clearContent();
  }
  invalidateProductPackagingCache_();
  return {saved: true, row: row, profile: validated};
}

function rollbackProductPackagingProfile_(productName, previous) {
  const sheet = getProductPackagingSheet_(false);
  if (!sheet) return;
  const current = loadProductPackagingProfiles_()[normalizeText(productName)];
  if (!current || !current.row) return;
  if (!previous) {
    sheet.getRange(current.row, 1, 1, PRODUCT_PACKAGING_HEADERS_.length).clearContent();
  } else {
    sheet.getRange(current.row, 1, 1, PRODUCT_PACKAGING_HEADERS_.length).setValues([[
      normalizeText(previous.productName || productName),
      previous.productName || productName,
      previous.mode || '',
      previous.emptyWeight === undefined ? '' : previous.emptyWeight,
      previous.referenceGrossWeight === undefined ? '' : previous.referenceGrossWeight,
      previous.referenceNetQuantity === undefined ? '' : previous.referenceNetQuantity,
      previous.massConversionFactor === undefined ? '' : previous.massConversionFactor,
      previous.referenceDate || '',
      new Date(),
      getCurrentUserEmail_()
    ]]);
  }
  invalidateProductPackagingCache_();
}

function applyProductPackagingProfileToInventory_(product, profile, options) {
  if (!product || !product.inventoryRow || !profile || !profile.mode) return {applied: false};
  const sheet = getSheetByConfiguredName_(CONFIG.SHEETS.INVENTORY);
  const layout = getConfiguredInventoryLayout_(product.type);
  if (!sheet || !layout || !layout.emptyContainerWeight) return {applied: false};
  const tareCell = sheet.getRange(layout.emptyContainerWeight + product.inventoryRow);
  if (profile.mode === CONFIG.PACKAGING_MODES.TARE_KNOWN) tareCell.setValue(profile.emptyWeight);
  else tareCell.clearContent();
  const settings = options || {};
  if (
    shouldSeedPackagingCurrentMeasurement_(settings) &&
    profile.mode === CONFIG.PACKAGING_MODES.GROSS_REFERENCE &&
    layout.grossWeight
  ) {
    const grossCell = sheet.getRange(layout.grossWeight + product.inventoryRow);
    if (grossCell.getValue() === '') grossCell.setValue(profile.referenceGrossWeight);
  }
  invalidateProductPackagingCache_();
  // Profil przekazujemy jawnie. Dzięki temu przebudowa formuły nie może użyć
  // starego profilu osadzonego w obiekcie produktu ani wpisu ze starego cache.
  const formulas = applyCanonicalFormulasToProductRow_(sheet, product, profile);
  SpreadsheetApp.flush();
  return {applied: true, formulas: formulas};
}

function shouldSeedPackagingCurrentMeasurement_(options) {
  const settings = options || {};
  // Import jest właścicielem bieżącego pomiaru i zapisuje go dokładnie raz
  // przez InventoryWriter. Profil opakowania nie może wyprzedzać tego zapisu.
  if (String(settings.workflow || '').trim().toUpperCase() === 'IMPORT') return false;
  return settings.seedCurrentMeasurement === true;
}

function applyProductPackagingProfilesToInventoryBatch_(sheet, products, profiles) {
  const inventory = sheet || getSheetByConfiguredName_(CONFIG.SHEETS.INVENTORY);
  if (!inventory) throw new Error('Nie znaleziono arkusza: ' + CONFIG.SHEETS.INVENTORY + '.');
  const profileMap = profiles || {};
  const byTareColumn = {};
  const catalog = (products || []).filter(product =>
    product && product.inventoryRow &&
    product.type !== CONFIG.PRODUCT_TYPES.LOCATION &&
    !isDirectFinalInventoryProduct_(product)
  );

  catalog.forEach(product => {
    const layout = getConfiguredInventoryLayout_(product.type);
    const column = normalizeColumnLetter_(layout && layout.emptyContainerWeight);
    const profile = profileMap[normalizeText(product.name)] ||
      getProductPackagingProfile_(product);
    if (!column || !profile) return;
    if (!byTareColumn[column]) byTareColumn[column] = [];
    byTareColumn[column].push({
      row: Number(product.inventoryRow),
      value: profile.mode === CONFIG.PACKAGING_MODES.TARE_KNOWN
        ? profile.emptyWeight
        : ''
    });
  });

  const lastRow = inventory.getLastRow();
  Object.keys(byTareColumn).forEach(column => {
    const range = inventory.getRange(column + '1:' + column + lastRow);
    const values = range.getValues();
    const formulas = range.getFormulas();
    const payload = values.map((row, index) => [
      formulas[index][0] || row[0]
    ]);
    byTareColumn[column].forEach(change => {
      payload[change.row - 1][0] = change.value;
    });
    range.setValues(payload);
  });

  const formulaResult = applyCanonicalFormulasToProductsBatch_(
    inventory, catalog, profileMap
  );
  invalidateProductPackagingCache_();
  SpreadsheetApp.flush();
  return {
    repairedProducts: catalog.length,
    tareColumns: Object.keys(byTareColumn).length,
    formulaCells: formulaResult.formulaCells,
    formulaColumns: formulaResult.formulaColumns
  };
}

function buildPackagingAwareOpenNetFormulaSpec_(
  product, layout, row, targetColumn, packagingProfileOverride
) {
  const profile = packagingProfileOverride ||
    getProductPackagingProfile_(product) ||
    {mode: CONFIG.PACKAGING_MODES.TARE_KNOWN};
  const targetNumber = inventoryColumnLetterToNumber_(targetColumn);
  const gross = layout.grossWeight + row;
  const tare = layout.emptyContainerWeight + row;
  const grossR1C1 = buildRelativeR1C1Reference_(targetNumber, inventoryColumnLetterToNumber_(layout.grossWeight));
  const tareR1C1 = buildRelativeR1C1Reference_(targetNumber, inventoryColumnLetterToNumber_(layout.emptyContainerWeight));

  if (profile.mode === CONFIG.PACKAGING_MODES.TARE_UNKNOWN) {
    if (isUnknownTareGrossApproved_(product)) {
      return {
        operation: 'GROSS_AS_NET_APPROVED',
        formula: '=' + gross,
        r1c1: '=' + grossR1C1
      };
    }
    return {
      operation: 'TARE_UNKNOWN',
      formula: '=0*(' + gross + '<>"")',
      r1c1: '=0*(' + grossR1C1 + '<>"")'
    };
  }
  if (profile.mode === CONFIG.PACKAGING_MODES.GROSS_REFERENCE) {
    return {
      operation: 'GROSS_REFERENCE',
      formula: buildReferenceRemainingFormula_(gross, profile),
      r1c1: buildReferenceRemainingFormula_(grossR1C1, profile)
    };
  }
  return {
    operation: 'DIFFERENCE',
    formula: '=(' + gross + '-' + tare + ')*(' + gross + '<>"")*(' + tare + '<>"")',
    r1c1: '=(' + grossR1C1 + '-' + tareR1C1 + ')*(' + grossR1C1 + '<>"")*(' + tareR1C1 + '<>"")'
  };
}

function buildReferenceRemainingFormula_(grossReference, profile) {
  const referenceGross = Number(profile.referenceGrossWeight);
  const initialQuantity = Number(profile.referenceNetQuantity);
  const factor = Number(profile.massConversionFactor || 1);
  const referenceGrossFormula = buildLocaleSafeFormulaNumber_(referenceGross);
  const initialQuantityFormula = buildLocaleSafeFormulaNumber_(initialQuantity);
  const factorFormula = buildLocaleSafeFormulaNumber_(factor);
  const loss = '((' + referenceGrossFormula + '-' + grossReference + ')+ABS(' + referenceGrossFormula + '-' + grossReference + '))/2';
  const rawRemaining = '(' + initialQuantityFormula + '-(' + loss + ')*' + factorFormula + ')';
  const remaining = '((' + rawRemaining + ')+ABS(' + rawRemaining + '))/2';
  return '=' + remaining + '*(' + grossReference + '<>"")';
}

/**
 * Apps Script setFormula() przyjmuje angielskie nazwy funkcji, ale literały
 * dziesiętne są interpretowane według ustawień regionalnych arkusza. Zapis
 * liczby jako ułamka (np. 0,7 -> 7/10) działa jednakowo w każdej lokalizacji.
 */
function buildLocaleSafeFormulaNumber_(value) {
  const number = Number(value);
  if (!Number.isFinite(number)) {
    throw new Error('Nieprawidłowa liczba profilu opakowania.');
  }
  if (Math.trunc(number) === number) return String(number);
  const text = Math.abs(number).toFixed(9).replace(/0+$/, '').replace(/\.$/, '');
  const parts = text.split('.');
  const denominator = Math.pow(10, (parts[1] || '').length);
  const numerator = Math.round(Math.abs(number) * denominator) * (number < 0 ? -1 : 1);
  return '(' + numerator + '/' + denominator + ')';
}

function calculateReferenceRemainingQuantity_(currentGrossWeight, profile) {
  const current = Number(currentGrossWeight);
  const reference = Number(profile && profile.referenceGrossWeight);
  const initial = Number(profile && profile.referenceNetQuantity);
  const factor = Number(profile && profile.massConversionFactor || 1);
  if (![current, reference, initial, factor].every(Number.isFinite)) return '';
  return Math.max(0, initial - Math.max(0, reference - current) * factor);
}

function migrateProductPackagingProfilesFromInventory() {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const sheet = getSheetByConfiguredName_(CONFIG.SHEETS.INVENTORY);
    const packagingSheet = getProductPackagingSheet_(true);
    const existing = loadProductPackagingProfiles_();
    const rows = [];
    scanInventoryProducts_().forEach(product => {
      if (product.type === CONFIG.PRODUCT_TYPES.LOCATION || isDirectFinalInventoryProduct_(product)) return;
      if (existing[normalizeText(product.name)]) return;
      const layout = getConfiguredInventoryLayout_(product.type);
      const tare = sheet.getRange(layout.emptyContainerWeight + product.inventoryRow).getValue();
      const number = normalizePackagingNumber_(tare);
      rows.push([
        normalizeText(product.name), product.name,
        Number.isFinite(number) && number > 0 ? CONFIG.PACKAGING_MODES.TARE_KNOWN : CONFIG.PACKAGING_MODES.TARE_UNKNOWN,
        Number.isFinite(number) && number > 0 ? number : '',
        '', '', '', '', new Date(), getCurrentUserEmail_()
      ]);
    });
    if (rows.length) {
      packagingSheet.getRange(packagingSheet.getLastRow() + 1, 1, rows.length, PRODUCT_PACKAGING_HEADERS_.length).setValues(rows);
    }
    invalidateProductPackagingCache_();
    const inventory = scanInventoryProducts_();
    const currentProfiles = loadProductPackagingProfiles_();
    const repairResult = applyProductPackagingProfilesToInventoryBatch_(
      sheet, inventory, currentProfiles
    );
    const formats = normalizeInventoryNumberFormats_();
    SpreadsheetApp.flush();
    return {
      success: true,
      createdProfiles: rows.length,
      repairedProducts: repairResult.repairedProducts,
      repairedFormulaCells: repairResult.formulaCells,
      formattedCells: formats.formattedCells
    };
  } finally {
    lock.releaseLock();
  }
}

function migrateProductPackagingProfilesWithDialog() {
  const result = migrateProductPackagingProfilesFromInventory();
  SpreadsheetApp.getUi().alert(
    'Inventory PRO — profile opakowań',
    'Gotowe. Utworzone profile: ' + result.createdProfiles +
      '. Naprawione produkty: ' + result.repairedProducts +
      '. Uporządkowane formaty liczb: ' + result.formattedCells + '.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  return result;
}
