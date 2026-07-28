/**
 * Inventory PRO 4.3.4
 * Odporne wyszukiwanie arkuszy po nazwie i aliasach konfiguracji.
 */

function getConfiguredSheetKey_(configuredName) {
  const wanted = normalizeText(configuredName);
  const sheets = CONFIG.SHEETS || {};
  const keys = Object.keys(sheets);
  for (let index = 0; index < keys.length; index++) {
    if (normalizeText(sheets[keys[index]]) === wanted) return keys[index];
  }
  return '';
}

function getConfiguredSheetAliases_(configuredName) {
  const canonical = String(configuredName || '').trim();
  const aliases = [canonical];
  const key = getConfiguredSheetKey_(canonical);
  const configuredAliases = key && CONFIG.SHEET_ALIASES && CONFIG.SHEET_ALIASES[key];
  if (Array.isArray(configuredAliases)) aliases.push.apply(aliases, configuredAliases);

  const seen = {};
  return aliases.map(value => String(value || '').trim()).filter(value => {
    const normalized = normalizeText(value);
    if (!normalized || seen[normalized]) return false;
    seen[normalized] = true;
    return true;
  });
}

function isConfiguredSheetName_(actualName, configuredName) {
  const actualKey = normalizeText(actualName);
  if (!actualKey) return false;
  return getConfiguredSheetAliases_(configuredName).some(alias =>
    normalizeText(alias) === actualKey
  );
}

function getSheetByConfiguredName_(configuredName) {
  const spreadsheet = getInventorySpreadsheet_();
  const wanted = getConfiguredSheetAliases_(configuredName).map(normalizeText);
  const sheets = spreadsheet.getSheets();
  for (let index = 0; index < sheets.length; index++) {
    if (wanted.indexOf(normalizeText(sheets[index].getName())) >= 0) return sheets[index];
  }
  return null;
}

function getOrCreateConfiguredSheet_(configuredName) {
  const existing = getSheetByConfiguredName_(configuredName);
  if (existing) return existing;
  return getInventorySpreadsheet_().insertSheet(String(configuredName || '').trim());
}

function registerInventorySpreadsheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) return '';
  PropertiesService.getScriptProperties().setProperty(
    'INVENTORY_SPREADSHEET_ID',
    spreadsheet.getId()
  );
  return spreadsheet.getId();
}

function getInventorySpreadsheet_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  const id = PropertiesService.getScriptProperties()
    .getProperty('INVENTORY_SPREADSHEET_ID');
  if (!id) {
    throw new Error(
      'Aplikacja mobilna nie jest jeszcze powiązana z arkuszem. Otwórz arkusz raz po instalacji.'
    );
  }
  return SpreadsheetApp.openById(id);
}
