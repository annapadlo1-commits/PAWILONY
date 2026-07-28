/**
 * Inventory PRO 4.3.4 — walidacja konfiguracji, układu i formuł PAWILONÓW.
 */

function validateEnterpriseConfiguration() {
  const result = buildValidationReport_();

  logInfo(
    'Core_Validation',
    'validateEnterpriseConfiguration',
    result.valid ? 'Konfiguracja poprawna' : 'Wykryto problemy',
    result
  );

  SpreadsheetApp.getUi().alert(
    'Inventory PRO - Walidacja',
    formatValidationReport_(result),
    SpreadsheetApp.getUi().ButtonSet.OK
  );

  return result;
}

function buildValidationReport_() {
  const errors = [];
  const warnings = [];
  const requiredSheets = [
    CONFIG.SHEETS.INVENTORY,
    CONFIG.SHEETS.DICTIONARY,
    CONFIG.SHEETS.SETTINGS,
    CONFIG.SHEETS.HISTORY,
    CONFIG.SHEETS.TECH_LOG,
    CONFIG.SHEETS.IMPORT_AUDIT
  ];

  requiredSheets.forEach(sheetName => {
    if (!getSheetByConfiguredName_(sheetName)) {
      errors.push('Brak arkusza: ' + sheetName);
    }
  });

  validateInventoryLayoutConfiguration_(errors);

  const dictionarySheet = getSheetByConfiguredName_(CONFIG.SHEETS.DICTIONARY);
  if (dictionarySheet) {
    const aliasHeader = String(dictionarySheet.getRange(1, 1).getDisplayValue() || '').trim();
    const productHeader = String(dictionarySheet.getRange(1, 4).getDisplayValue() || '').trim();

    if (!aliasHeader) warnings.push('Brak nagłówka aliasów w A1 arkusza ' + dictionarySheet.getName() + '.');
    if (normalizeText(productHeader) !== 'produkt') {
      errors.push('Kolumna D arkusza ' + dictionarySheet.getName() + ' powinna mieć nagłówek Produkt.');
    }
  }

  let catalogSummary = null;
  try {
    catalogSummary = getProductCatalogSummary();
    if (catalogSummary.products === 0) errors.push('Katalog produktów jest pusty.');
    if (catalogSummary.missingInventoryRow > 0) {
      warnings.push('Produkty bez wiersza w INWENTURA: ' + catalogSummary.missingInventoryRow + '.');
    }
  } catch (error) {
    errors.push('Nie udało się zbudować katalogu: ' + normalizeError_(error).message);
  }

  let scannedProducts = null;
  try {
    scannedProducts = scanInventoryProducts_();
    if (!scannedProducts.length) errors.push('Skan arkusza INWENTURA nie wykrył produktów.');
    scannedProducts.forEach(product => {
      const mapping = validateProductColumnMapping_(product.type, product.columns, product);
      if (!mapping.valid) {
        errors.push('Nieprawidłowe mapowanie „' + product.name + '”: ' + mapping.errors.join(' '));
      }
    });
    if (catalogSummary && catalogSummary.products && scannedProducts.length) {
      const difference = Math.abs(Number(catalogSummary.products) - scannedProducts.length);
      if (difference > 10 && difference / Math.max(scannedProducts.length, 1) > 0.05) {
        warnings.push(
          'Katalog i fizyczny arkusz różnią się liczbą produktów: katalog ' +
          catalogSummary.products + ', arkusz ' + scannedProducts.length + '.'
        );
      }
    }
  } catch (error) {
    errors.push('Nie udało się przeskanować układu INWENTURA: ' + normalizeError_(error).message);
  }

  let formulaAudit = null;
  const inventorySheet = getSheetByConfiguredName_(CONFIG.SHEETS.INVENTORY);
  if (inventorySheet && scannedProducts) {
    try {
      formulaAudit = auditInventoryFormulaCoverage_({
        sheet: inventorySheet,
        products: scannedProducts
      });
      if (formulaAudit.conflictFormulaCells > 0) {
        errors.push(
          'Formuły: ' + formulaAudit.conflictFormulaCells +
          ' konfliktów wymagających ręcznego rozstrzygnięcia.'
        );
      }
      if (formulaAudit.missingFormulaCells > 0) {
        errors.push('Formuły: ' + formulaAudit.missingFormulaCells + ' pustych komórek bez formuły.');
      }
      if (formulaAudit.flattenedFormulaCells > 0) {
        errors.push('Formuły: ' + formulaAudit.flattenedFormulaCells + ' spłaszczonych wyników.');
      }
      if (formulaAudit.legacyFormulaCells > 0) {
        // Poprawne sumy starszego typu z operatorem + są równoważne
        // kontraktowi i nie obniżają Application Health.
      }
      if (formulaAudit.invalidFormulaCells > 0) {
        errors.push('Formuły: ' + formulaAudit.invalidFormulaCells + ' nieprawidłowych wzorów.');
      }
      if (formulaAudit.errorFormulaCells > 0) {
        errors.push('Formuły: ' + formulaAudit.errorFormulaCells + ' błędów obliczeniowych.');
      }
    } catch (error) {
      errors.push('Nie udało się przeprowadzić audytu formuł: ' + normalizeError_(error).message);
    }
  }

  return {
    valid: errors.length === 0,
    errors: errors,
    warnings: warnings,
    catalogSummary: catalogSummary,
    scannedProducts: scannedProducts ? scannedProducts.length : null,
    formulaAudit: formulaAudit
  };
}

function validateInventoryLayoutConfiguration_(errors) {
  [
    CONFIG.PRODUCT_TYPES.NORMAL,
    CONFIG.PRODUCT_TYPES.KEG,
    CONFIG.PRODUCT_TYPES.LOCATION
  ].forEach(type => {
    let layout = null;
    try {
      layout = getConfiguredInventoryLayout_(type);
    } catch (error) {
      errors.push('Brak poprawnego profilu układu dla typu ' + type + '.');
      return;
    }

    const mapping = validateProductColumnMapping_(type, getInputColumnsForProductType_(type));
    if (!mapping.valid) {
      errors.push('Profil ' + type + ': ' + mapping.errors.join(' '));
    }

    const formulas = getFormulaColumnsForProductType_(type);
    const inputs = getAllowedInputColumnsForProductType_(type);
    formulas.forEach(column => {
      if (inputs.includes(column)) {
        errors.push('Profil ' + type + ': kolumna ' + column + ' jest jednocześnie wejściowa i obliczeniowa.');
      }
    });

    if (!layout.finalTotal) errors.push('Profil ' + type + ' nie definiuje kolumny stanu końcowego.');
  });
}

function formatValidationReport_(result) {
  let message = result.valid ? 'Konfiguracja jest poprawna.' : 'Konfiguracja wymaga poprawy.';

  if (result.errors.length) message += '\n\nBŁĘDY:\n- ' + result.errors.join('\n- ');
  if (result.warnings.length) message += '\n\nOSTRZEŻENIA:\n- ' + result.warnings.join('\n- ');

  if (result.catalogSummary) {
    message +=
      '\n\nKATALOG:\n' +
      'Produkty: ' + result.catalogSummary.products + '\n' +
      'Aliasy: ' + result.catalogSummary.aliases + '\n' +
      'NORMAL: ' + result.catalogSummary.normal + '\n' +
      'KEG: ' + result.catalogSummary.keg + '\n' +
      'LOCATION: ' + result.catalogSummary.location;
  }

  if (result.scannedProducts !== null) {
    message += '\nFizyczne produkty w INWENTURA: ' + result.scannedProducts;
  }

  if (result.formulaAudit) {
    message +=
      '\n\nFORMUŁY:\n' +
      'Oczekiwane: ' + result.formulaAudit.expectedFormulaCells + '\n' +
      'Obecne: ' + result.formulaAudit.presentFormulaCells + '\n' +
      'Brakujące: ' + result.formulaAudit.missingFormulaCells + '\n' +
      'Spłaszczone: ' + result.formulaAudit.flattenedFormulaCells + '\n' +
      'Konflikty: ' + result.formulaAudit.conflictFormulaCells + '\n' +
      'Starsze poprawne (+): ' + result.formulaAudit.legacyFormulaCells + '\n' +
      'Nieprawidłowe: ' + result.formulaAudit.invalidFormulaCells + '\n' +
      'Błędy obliczeń: ' + result.formulaAudit.errorFormulaCells;
  }

  return message;
}

function showEnterpriseDiagnostics() {
  const validation = buildValidationReport_();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const catalog = validation.catalogSummary || {};
  const formulas = validation.formulaAudit || {};

  const lines = [
    'Wersja: ' + CONFIG.VERSION,
    'Lokal: ' + CONFIG.LOCATION.NAME,
    'Plik: ' + spreadsheet.getName(),
    'Strefa czasowa: ' + spreadsheet.getSpreadsheetTimeZone(),
    'Ustawienia regionalne: ' + spreadsheet.getSpreadsheetLocale(),
    '',
    'STATUS: ' + (validation.valid ? 'POPRAWNY' : 'WYMAGA UWAGI'),
    'Błędy: ' + validation.errors.length,
    'Ostrzeżenia: ' + validation.warnings.length,
    '',
    'KATALOG',
    'Produkty: ' + Number(catalog.products || 0),
    'Fizyczne wiersze: ' + Number(validation.scannedProducts || 0),
    'Aliasy: ' + Number(catalog.aliases || 0),
    'NORMAL: ' + Number(catalog.normal || 0),
    'KEG: ' + Number(catalog.keg || 0),
    'LOCATION: ' + Number(catalog.location || 0),
    'Brak wiersza INWENTURA: ' + Number(catalog.missingInventoryRow || 0),
    '',
    'FORMUŁY',
    'Oczekiwane: ' + Number(formulas.expectedFormulaCells || 0),
    'Obecne: ' + Number(formulas.presentFormulaCells || 0),
    'Brakujące: ' + Number(formulas.missingFormulaCells || 0),
    'Spłaszczone: ' + Number(formulas.flattenedFormulaCells || 0),
    'Konflikty: ' + Number(formulas.conflictFormulaCells || 0),
    'Starsze poprawne (+): ' + Number(formulas.legacyFormulaCells || 0),
    'Nieprawidłowe: ' + Number(formulas.invalidFormulaCells || 0),
    'Błędy obliczeń: ' + Number(formulas.errorFormulaCells || 0)
  ];

  if (validation.errors.length) lines.push('', 'BŁĘDY', '- ' + validation.errors.join('\n- '));
  if (validation.warnings.length) lines.push('', 'OSTRZEŻENIA', '- ' + validation.warnings.join('\n- '));

  const diagnostics = {
    version: CONFIG.VERSION,
    location: CONFIG.LOCATION,
    spreadsheetName: spreadsheet.getName(),
    spreadsheetId: spreadsheet.getId(),
    timezone: spreadsheet.getSpreadsheetTimeZone(),
    locale: spreadsheet.getSpreadsheetLocale(),
    validation: validation
  };

  logInfo('Core_Validation', 'showEnterpriseDiagnostics', 'Wygenerowano diagnostykę', diagnostics);
  SpreadsheetApp.getUi().alert(
    'Inventory PRO - Diagnostyka',
    lines.join('\n'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
