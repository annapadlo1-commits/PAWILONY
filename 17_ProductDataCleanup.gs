/**
 * Inventory PRO Enterprise v2.7.1
 * Bezpieczny audyt i porzadkowanie danych produktowych.
 */

// normalizeBusinessCategory_ przeniesiono do CategoryService.gs


function buildNewProductAuditKey_(name, inventoryRow) {
  return normalizeText(name) + '|' + String(Number(inventoryRow || 0));
}

function runProductDataAudit() {
  const result = buildProductDataAudit_();
  writeProductDataAuditSheet_(result);
  SpreadsheetApp.getUi().alert(
    'Inventory PRO - Audyt danych',
    'Konflikty nazw: ' + result.duplicateNames.length + '\n' +
    'Konflikty aliasow: ' + result.aliasConflicts.length + '\n' +
    'Kategorie do normalizacji: ' + result.categoryIssues.length + '\n' +
    'Duplikaty historii nowych produktow: ' + result.newProductAuditDuplicates.length + '\n' +
    'Techniczne wiersze w Historii: ' + result.legacyHistoryRows.length + '\n\n' +
    'Szczegoly zapisano w arkuszu "' + CONFIG.SHEETS.DATA_AUDIT + '".',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  return result;
}

function repairProductData() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Inventory PRO - Napraw dane',
    'Operacja znormalizuje kategorie, usunie powtorzone wpisy w "Nowe produkty" i przeniesie stare logi techniczne z "Historia" do "Historia legacy". Nie laczy produktow automatycznie. Kontynuowac?',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return { cancelled: true };

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const result = {
      normalizedCategories: normalizeDictionaryCategories_(),
      removedNewProductAuditDuplicates: removeDuplicateNewProductAudits_(),
      movedLegacyHistoryRows: moveLegacyTechnicalHistoryRows_()
    };
    SpreadsheetApp.flush();
    invalidateProductCatalogCache_();
    runProductDataAudit();
    ui.alert('Naprawa zakonczona',
      'Kategorie: ' + result.normalizedCategories + '\n' +
      'Usuniete duplikaty nowych produktow: ' + result.removedNewProductAuditDuplicates + '\n' +
      'Przeniesione stare logi: ' + result.movedLegacyHistoryRows,
      ui.ButtonSet.OK);
    return result;
  } finally {
    lock.releaseLock();
  }
}

function buildProductDataAudit_() {
  const configs = loadProductConfigurations();
  const aliases = loadAliases();
  const duplicateNames = [];
  const byName = {};
  configs.forEach(product => {
    const key = normalizeText(product.name);
    if (!byName[key]) byName[key] = [];
    byName[key].push(product.name);
  });
  Object.keys(byName).forEach(key => {
    if (byName[key].length > 1) duplicateNames.push({ key: key, products: byName[key] });
  });

  const aliasConflicts = [];
  Object.keys(aliases).forEach(alias => {
    const target = aliases[alias];
    const product = byName[normalizeText(alias)];
    if (product && normalizeText(target) !== normalizeText(alias)) {
      aliasConflicts.push({ alias: alias, target: target, nameCollision: product.join(', ') });
    }
  });

  const categoryIssues = configs
    .filter(product => product.category && normalizeBusinessCategory_(product.category) !== product.category)
    .map(product => ({ product: product.name, current: product.category, proposed: normalizeBusinessCategory_(product.category) }));

  return {
    duplicateNames: duplicateNames,
    aliasConflicts: aliasConflicts,
    categoryIssues: categoryIssues,
    newProductAuditDuplicates: findDuplicateNewProductAudits_(),
    legacyHistoryRows: findLegacyTechnicalHistoryRows_()
  };
}

function findDuplicateNewProductAudits_() {
  const sheet = getOrCreateNewProductsSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  const seen = {};
  const duplicates = [];
  values.forEach((row, index) => {
    const key = buildNewProductAuditKey_(row[2], row[6]);
    if (seen[key]) duplicates.push({ row: index + 2, key: key, product: row[2], inventoryRow: row[6] });
    else seen[key] = index + 2;
  });
  return duplicates;
}

function removeDuplicateNewProductAudits_() {
  const sheet = getOrCreateNewProductsSheet_();
  const duplicates = findDuplicateNewProductAudits_();
  duplicates.sort((a, b) => b.row - a.row).forEach(item => sheet.deleteRow(item.row));
  return duplicates.length;
}

function findLegacyTechnicalHistoryRows_() {
  const sheet = getOrCreateBusinessHistorySheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, Math.max(16, sheet.getLastColumn())).getValues();
  const result = [];
  values.forEach((row, index) => {
    const id = String(row[0] || '').trim();
    if (id && !/^IMP-/i.test(id)) result.push({ row: index + 2, values: row });
  });
  return result;
}

function moveLegacyTechnicalHistoryRows_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const history = getOrCreateBusinessHistorySheet_();
  const rows = findLegacyTechnicalHistoryRows_();
  if (!rows.length) return 0;
  let legacy = spreadsheet.getSheetByName(CONFIG.SHEETS.HISTORY_LEGACY);
  if (!legacy) legacy = spreadsheet.insertSheet(CONFIG.SHEETS.HISTORY_LEGACY);
  if (legacy.getLastRow() === 0) {
    const width = Math.max(16, history.getLastColumn());
    legacy.getRange(1, 1, 1, width).setValues([history.getRange(1, 1, 1, width).getValues()[0]]);
    legacy.setFrozenRows(1);
  }
  const width = Math.max(16, history.getLastColumn());
  legacy.getRange(legacy.getLastRow() + 1, 1, rows.length, width).setValues(rows.map(item => item.values.slice(0, width)));
  rows.sort((a, b) => b.row - a.row).forEach(item => history.deleteRow(item.row));
  return rows.length;
}

function normalizeDictionaryCategories_() {
  const sheet = getDictionarySheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.DICTIONARY.FIRST_DATA_ROW) return 0;
  const categoryColumn = CONFIG.DICTIONARY.CONFIG_START_COLUMN + 2;
  const count = lastRow - CONFIG.DICTIONARY.FIRST_DATA_ROW + 1;
  const range = sheet.getRange(CONFIG.DICTIONARY.FIRST_DATA_ROW, categoryColumn, count, 1);
  const values = range.getValues();
  let changed = 0;
  values.forEach(row => {
    const current = String(row[0] || '').trim();
    const normalized = normalizeBusinessCategory_(current);
    if (current && normalized !== current) {
      row[0] = normalized;
      changed++;
    }
  });
  if (changed) range.setValues(values);
  return changed;
}

function writeProductDataAuditSheet_(audit) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.DATA_AUDIT);
  if (!sheet) sheet = spreadsheet.insertSheet(CONFIG.SHEETS.DATA_AUDIT);
  sheet.clear();
  const rows = [['Typ', 'Produkt / Alias', 'Aktualna wartosc', 'Propozycja / Szczegoly']];
  audit.duplicateNames.forEach(item => rows.push(['DUPLIKAT_NAZWY', item.key, item.products.join(', '), 'Wymaga recznego polaczenia w Product Manager']));
  audit.aliasConflicts.forEach(item => rows.push(['KONFLIKT_ALIASU', item.alias, item.target, 'Kolizja z nazwa: ' + item.nameCollision]));
  audit.categoryIssues.forEach(item => rows.push(['KATEGORIA', item.product, item.current, item.proposed]));
  audit.newProductAuditDuplicates.forEach(item => rows.push(['DUPLIKAT_HISTORII_NOWEGO_PRODUKTU', item.product, item.inventoryRow, 'Wiersz arkusza: ' + item.row]));
  audit.legacyHistoryRows.forEach(item => rows.push(['STARY_LOG_W_HISTORII', String(item.values[0] || ''), String(item.values[1] || ''), 'Wiersz arkusza: ' + item.row]));
  sheet.getRange(1, 1, rows.length, 4).setValues(rows);
  sheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#f9cb9c');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 4);
}
