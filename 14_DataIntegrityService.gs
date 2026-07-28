/**
 * Inventory PRO Enterprise v2.8.4
 * Final 2.8 data integrity checks and safe dictionary rebuild.
 */

function auditDictionaryConfigurationIntegrity() {
  const scanned = scanInventoryProducts_();
  const configured = loadAllDictionaryConfigurations284_();
  const scannedIndex = {};
  scanned.forEach(item => scannedIndex[item.normalizedName] = item);

  const issues = [];
  configured.forEach(item => {
    const source = scannedIndex[item.normalizedName];
    if (!source) {
      issues.push(['BRAK W INWENTURZE', item.name, item.type, item.category, '']);
      return;
    }
    const expectedCategory = normalizeBusinessCategory_(source.category);
    const currentCategory = normalizeBusinessCategory_(item.category);
    if (item.type !== source.type) {
      issues.push(['NIEPRAWIDŁOWY TYP', item.name, item.type, source.type, '']);
    }
    if (currentCategory !== expectedCategory) {
      issues.push(['NIEPRAWIDŁOWA KATEGORIA', item.name, item.category, expectedCategory, '']);
    }
    const expected = source.columns || {};
    const actual = item.columns || {};
    ['quantity','weight','warehouse','darkroom','fridges'].forEach(key => {
      if (String(actual[key] || '') !== String(expected[key] || '')) {
        issues.push(['NIEPRAWIDŁOWA KOLUMNA', item.name, key + ': ' + (actual[key] || '—'), key + ': ' + (expected[key] || '—'), '']);
      }
    });
  });

  const sheet = getOrCreateConfiguredSheet_(CONFIG.SHEETS.DATA_AUDIT);
  sheet.clearContents();
  const headers = ['Problem', 'Produkt', 'Wartość obecna', 'Wartość oczekiwana', 'Uwagi'];
  sheet.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold').setBackground('#f6b26b');
  if (issues.length) sheet.getRange(2,1,issues.length,headers.length).setValues(issues);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);

  return { productsScanned: scanned.length, configurations: configured.length, issues: issues.length };
}

function rebuildDictionaryConfigurationSafely() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Bezpieczna odbudowa SŁOWNIKA',
    'Operacja utworzy kopię arkusza SŁOWNIK, a następnie odbuduje wyłącznie tabelę konfiguracji D:L na podstawie arkusza INWENTURA. Aliasy A:B pozostaną bez zmian. Formuły i dane arkusza INWENTURA nie zostaną dotknięte. Kontynuować?',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dictionary = getDictionarySheet_();
    const zone = ss.getSpreadsheetTimeZone();
    const backupName = createUniqueArchiveName_(ss, ('BACKUP SLOWNIK ' + Utilities.formatDate(new Date(), zone, 'yyyy-MM-dd HH-mm-ss')).slice(0,90));
    dictionary.copyTo(ss).setName(backupName).hideSheet();

    const products = scanInventoryProducts_();
    if (!products.length) throw new Error('Nie znaleziono produktów w arkuszu INWENTURA.');
    writeFullProductConfiguration_(products);
    SpreadsheetApp.flush();
    invalidateProductCatalogCache_();

    const audit = auditDictionaryConfigurationIntegrity();
    ui.alert(
      'Odbudowa zakończona',
      'Produkty: ' + products.length + '\nProblemy po odbudowie: ' + audit.issues + '\nKopia bezpieczeństwa: ' + backupName,
      ui.ButtonSet.OK
    );
    return { rebuilt: products.length, issues: audit.issues, backupSheet: backupName };
  } finally {
    lock.releaseLock();
  }
}

function loadAllDictionaryConfigurations284_() {
  const sheet = getDictionarySheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2,4,lastRow-1,9).getValues().map((row,index)=>({
    dictionaryRow:index+2,
    name:String(row[0]||'').trim(),
    normalizedName:normalizeText(row[0]),
    type:String(row[1]||'').trim().toUpperCase(),
    category:String(row[2]||'').trim(),
    columns:{
      quantity:String(row[3]||'').trim().toUpperCase(),
      weight:String(row[4]||'').trim().toUpperCase(),
      warehouse:String(row[5]||'').trim().toUpperCase(),
      darkroom:String(row[6]||'').trim().toUpperCase(),
      fridges:String(row[7]||'').trim().toUpperCase()
    },
    active:['tak','true','1'].includes(normalizeText(row[8]))
  })).filter(item=>item.name);
}
