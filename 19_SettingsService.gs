/**
 * Inventory PRO Enterprise v2.4
 * Ustawienia walidacji edytowalne z arkusza Ustawienia.
 */

const QUALITY_SETTING_DEFINITIONS_ = [
  ['OSTRZEGAJ_PRZY_ZERO', true, 'Pokaz ostrzezenie, gdy wpisano 0 (TAK/NIE)'],
  ['BLOKUJ_WARTOSCI_UJEMNE', true, 'Nie zapisuj wartosci ujemnych (TAK/NIE)'],
  ['MAX_SZTUKI_STANDARD', 20, 'Maksymalna liczba sztuk dla produktow standardowych'],
  ['MAX_WAGA_STANDARD', 20, 'Maksymalna waga dla produktow standardowych'],
  ['MAX_PELNE_KEGI', 20, 'Maksymalna liczba pelnych kegow'],
  ['MAX_WAGA_KEGA', 100, 'Maksymalna waga kega'],
  ['MAX_SZTUKI_LOKALIZACJA', 500, 'Maksymalna liczba sztuk dla piw butelkowych i softow w jednej lokalizacji'],
  ['PROG_DUPLIKATU', 2, 'Od ilu wystapien oznaczac produkt jako duplikat']
];

function ensureQualitySettingsSheet_() {
  const spreadsheet = getInventorySpreadsheet_();
  let sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.SETTINGS);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(CONFIG.SHEETS.SETTINGS);
  }

  const headers = [['Parametr', 'Wartosc', 'Opis']];
  sheet.getRange(1, 1, 1, 3)
    .setValues(headers)
    .setFontWeight('bold')
    .setBackground('#d9ead3');
  sheet.setFrozenRows(1);

  const lastRow = sheet.getLastRow();
  const existing = {};

  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, 2)
      .getDisplayValues()
      .forEach(row => {
        const key = String(row[0] || '').trim().toUpperCase();
        if (key) existing[key] = true;
      });
  }

  const missing = QUALITY_SETTING_DEFINITIONS_.filter(definition =>
    !existing[definition[0]]
  );

  if (missing.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, missing.length, 3)
      .setValues(missing);
  }

  sheet.autoResizeColumns(1, 3);
  return sheet;
}

function loadQualitySettings_() {
  const defaults = {
    warnZero: true,
    blockNegative: true,
    normalWholeWarning: CONFIG.QUALITY.NORMAL_WHOLE_WARNING,
    normalWeightWarning: CONFIG.QUALITY.NORMAL_WEIGHT_WARNING,
    kegWholeWarning: CONFIG.QUALITY.KEG_WHOLE_WARNING,
    kegWeightWarning: CONFIG.QUALITY.KEG_WEIGHT_WARNING,
    locationWarning: CONFIG.QUALITY.LOCATION_WARNING,
    duplicateWarningCount: 2
  };

  const sheet = ensureQualitySettingsSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return defaults;

  const values = sheet.getRange(2, 1, lastRow - 1, 2).getDisplayValues();
  const map = {};
  values.forEach(row => {
    const key = String(row[0] || '').trim().toUpperCase();
    if (key) map[key] = row[1];
  });

  return {
    warnZero: parseBooleanSetting_(map.OSTRZEGAJ_PRZY_ZERO, defaults.warnZero),
    blockNegative: parseBooleanSetting_(map.BLOKUJ_WARTOSCI_UJEMNE, defaults.blockNegative),
    normalWholeWarning: parseNumberSetting_(map.MAX_SZTUKI_STANDARD, defaults.normalWholeWarning),
    normalWeightWarning: parseNumberSetting_(map.MAX_WAGA_STANDARD, defaults.normalWeightWarning),
    kegWholeWarning: parseNumberSetting_(map.MAX_PELNE_KEGI, defaults.kegWholeWarning),
    kegWeightWarning: parseNumberSetting_(map.MAX_WAGA_KEGA, defaults.kegWeightWarning),
    locationWarning: parseNumberSetting_(map.MAX_SZTUKI_LOKALIZACJA, defaults.locationWarning),
    duplicateWarningCount: Math.max(2, Math.floor(parseNumberSetting_(map.PROG_DUPLIKATU, defaults.duplicateWarningCount)))
  };
}

function parseNumberSetting_(rawValue, fallback) {
  const normalized = String(rawValue === undefined ? '' : rawValue)
    .trim()
    .replace(',', '.');
  const value = Number(normalized);
  return Number.isFinite(value) && value >= 0 ? value : fallback;
}

function parseBooleanSetting_(rawValue, fallback) {
  const normalized = normalizeText(rawValue);
  if (['tak', 'true', '1', 'yes', 'on'].includes(normalized)) return true;
  if (['nie', 'false', '0', 'no', 'off'].includes(normalized)) return false;
  return fallback;
}
