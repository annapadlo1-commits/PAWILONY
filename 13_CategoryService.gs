/**
 * Inventory PRO Enterprise v2.8.2
 * Centralna normalizacja i naprawa kategorii biznesowych.
 */

const INVENTORY_BUSINESS_CATEGORIES_ = Object.freeze([
  'BITTER', 'BRANDY', 'GIN', 'LIKIER', 'PIWO', 'PIWO KEG', 'PIWO BUTELKI',
  'RUM', 'TEQUILA', 'WERMUT', 'WHISKY', 'WINO', 'WÓDKA',
  'PREMIXY', 'SOFTY', 'KAWA'
]);

function getBusinessCategories() {
  return INVENTORY_BUSINESS_CATEGORIES_.slice();
}

function normalizeBusinessCategory_(value) {
  const raw = String(value || '').trim();
  const normalized = normalizeText(raw);
  if (!normalized) return '';

  const rules = [
    { pattern: /softy/, value: 'SOFTY' },
    { pattern: /piwo\s+butelki/, value: 'PIWO BUTELKI' },
    { pattern: /piwo\s+keg/, value: 'PIWO KEG' },
    { pattern: /^piw[oa]$/, value: 'PIWO' },
    { pattern: /\bbitter\b/, value: 'BITTER' },
    { pattern: /\bbrandy\b/, value: 'BRANDY' },
    { pattern: /\bgin\b/, value: 'GIN' },
    { pattern: /\blikier\b/, value: 'LIKIER' },
    { pattern: /\brum\b/, value: 'RUM' },
    { pattern: /\btequila\b/, value: 'TEQUILA' },
    { pattern: /\bwermut\b/, value: 'WERMUT' },
    { pattern: /\bwhisky\b/, value: 'WHISKY' },
    { pattern: /\bwino\b/, value: 'WINO' },
    { pattern: /\bwodka\b/, value: 'WÓDKA' },
    { pattern: /\bpremixy\b/, value: 'PREMIXY' },
    { pattern: /\bkawa\b/, value: 'KAWA' }
  ];

  for (let i = 0; i < rules.length; i++) {
    if (rules[i].pattern.test(normalized)) return rules[i].value;
  }

  const exact = INVENTORY_BUSINESS_CATEGORIES_.find(category =>
    normalizeText(category) === normalized
  );
  return exact || '';
}

function repairDictionaryCategoriesFromInventory() {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const dictionary = getDictionarySheet_();
    const inventory = getSheetByConfiguredName_(CONFIG.SHEETS.INVENTORY);
    if (!dictionary || !inventory) throw new Error('Brak arkusza SŁOWNIK lub INWENTURA.');

    const inventoryValues = inventory.getDataRange().getDisplayValues();
    const categoryByRow = buildStrictInventoryCategoryMapFromSheet_(inventory, inventoryValues);
    const startRow = CONFIG.DICTIONARY.FIRST_DATA_ROW;
    const lastRow = dictionary.getLastRow();
    if (lastRow < startRow) return { updated: 0, unresolved: 0 };

    const configRange = dictionary.getRange(
      startRow,
      CONFIG.DICTIONARY.CONFIG_START_COLUMN,
      lastRow - startRow + 1,
      CONFIG.DICTIONARY.CONFIG_COLUMN_COUNT
    );
    const configs = configRange.getValues();
    let updated = 0;
    let unresolved = 0;

    configs.forEach(row => {
      const productName = String(row[0] || '').trim();
      if (!productName) return;
      const current = normalizeBusinessCategory_(row[2]);
      const inventoryRow = findInventoryProductRow_(inventoryValues, productName);
      const inferred = inventoryRow ? (categoryByRow[inventoryRow] || '') : '';
      const resolved = inferred || current;
      if (!resolved) {
        unresolved++;
        return;
      }
      if (String(row[2] || '').trim() !== resolved) {
        row[2] = resolved;
        updated++;
      }
    });

    if (updated) configRange.setValues(configs);
    SpreadsheetApp.flush();
    invalidateProductCatalogCache_();
    return { updated: updated, unresolved: unresolved };
  } finally {
    lock.releaseLock();
  }
}

function buildInventoryCategoryMap_(values) {
  const map = {};
  let current = '';
  for (let r = 0; r < values.length; r++) {
    const rowText = values[r].join(' ').trim();
    const category = normalizeBusinessCategory_(rowText);
    if (category) current = category;
    map[r + 1] = current;
  }
  return map;
}

function findInventoryProductRow_(values, productName) {
  const target = normalizeText(productName);
  for (let r = 0; r < values.length; r++) {
    if (normalizeText(values[r][0]) === target) return r + 1;
  }
  return 0;
}


function requireBusinessCategory_(value, productName, inventoryRow) {
  const category = normalizeBusinessCategory_(value);
  if (!category || category === 'POZOSTAŁE') {
    throw new Error('Nie można utworzyć konfiguracji produktu „' + String(productName || '') + '” w wierszu ' + String(inventoryRow || '?') + ': brak fizycznej kategorii w INWENTURA.');
  }
  return category;
}

function buildStrictInventoryCategoryMapFromSheet_(sheet, displayValues) {
  const values = displayValues || sheet.getDataRange().getDisplayValues();
  const range = sheet.getRange(1, 1, Math.max(values.length, 1), Math.max(sheet.getLastColumn(), 1));
  const mergedRows = buildMergedHeaderRowMap_(range);
  const map = {};
  let current = '';
  for (let r = 0; r < values.length; r++) {
    const sheetRow = r + 1;
    const row = values[r] || [];
    const firstCell = String(row[0] || '').trim();
    const rowText = row.filter(Boolean).join(' ').trim();
    const exactFirst = normalizeBusinessCategory_(firstCell);
    const header = detectSectionHeader_(rowText, sheetRow, mergedRows);
    if (header.isHeader && header.category) current = header.category;
    else if (exactFirst && normalizeText(firstCell) === normalizeText(exactFirst)) current = exactFirst;
    map[sheetRow] = current;
  }
  return map;
}

function repairInvalidDictionaryCategories290() {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dictionary = getDictionarySheet_();
    const inventory = getSheetByConfiguredName_(CONFIG.SHEETS.INVENTORY);
    if (!dictionary || !inventory) throw new Error('Brak arkusza SŁOWNIK lub INWENTURA.');

    const zone = ss.getSpreadsheetTimeZone();
    const backupName = createUniqueArchiveName_(ss, ('BACKUP SLOWNIK KATEGORIE ' + Utilities.formatDate(new Date(), zone, 'yyyy-MM-dd HH-mm-ss')).slice(0, 90));
    dictionary.copyTo(ss).setName(backupName).hideSheet();

    const values = inventory.getDataRange().getDisplayValues();
    const categoryByRow = buildStrictInventoryCategoryMapFromSheet_(inventory, values);
    const productRows = {};
    values.forEach((row, index) => {
      const name = normalizeText(row[0]);
      if (name && !productRows[name]) productRows[name] = index + 1;
    });

    const startRow = CONFIG.DICTIONARY.FIRST_DATA_ROW;
    const count = Math.max(dictionary.getLastRow() - startRow + 1, 0);
    if (!count) return { updated: 0, unresolved: 0, backupSheet: backupName };
    const range = dictionary.getRange(startRow, CONFIG.DICTIONARY.CONFIG_START_COLUMN, count, CONFIG.DICTIONARY.CONFIG_COLUMN_COUNT);
    const rows = range.getValues();
    let updated = 0;
    let unresolved = 0;
    const changes = [];

    rows.forEach(row => {
      const productName = String(row[0] || '').trim();
      if (!productName) return;
      const inventoryRow = productRows[normalizeText(productName)] || 0;
      const physicalCategory = inventoryRow ? categoryByRow[inventoryRow] : '';
      if (!physicalCategory) { unresolved++; return; }
      const oldCategory = String(row[2] || '').trim();
      if (oldCategory !== physicalCategory) {
        row[2] = physicalCategory;
        updated++;
        changes.push([productName, oldCategory || 'BRAK', physicalCategory, inventoryRow]);
      }
    });

    if (updated) range.setValues(rows);
    SpreadsheetApp.flush();
    invalidateProductCatalogCache_();
    writeCategoryRepairAudit290_(changes, unresolved);
    return { updated: updated, unresolved: unresolved, backupSheet: backupName };
  } finally {
    lock.releaseLock();
  }
}

function repairInvalidDictionaryCategories290WithDialog() {
  const ui = SpreadsheetApp.getUi();
  const answer = ui.alert('Napraw kategorie 2.9.0', 'Kategorie w SŁOWNIKU zostaną porównane z fizycznymi sekcjami w INWENTURA. Powstanie ukryta kopia bezpieczeństwa. Formuły i dane INWENTURA nie zostaną zmienione. Kontynuować?', ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;
  const result = repairInvalidDictionaryCategories290();
  ui.alert('Naprawa zakończona', 'Poprawiono: ' + result.updated + '\nNierozwiązane: ' + result.unresolved + '\nKopia: ' + result.backupSheet, ui.ButtonSet.OK);
  return result;
}

function writeCategoryRepairAudit290_(changes, unresolved) {
  const sheet = getOrCreateConfiguredSheet_(CONFIG.SHEETS.DATA_AUDIT);
  sheet.clearContents();
  const headers = ['Produkt', 'Kategoria przed', 'Kategoria po', 'Wiersz INWENTURA'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#f6b26b');
  if (changes.length) sheet.getRange(2, 1, changes.length, headers.length).setValues(changes);
  sheet.getRange(Math.max(changes.length + 3, 3), 1).setValue('Nierozwiązane: ' + unresolved);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
}