/**
 * Inventory PRO Enterprise v2.1
 * Centralny logger techniczny.
 */

const LOG_LEVELS = Object.freeze({
  INFO: 'INFO',
  WARN: 'WARN',
  ERROR: 'ERROR',
  DEBUG: 'DEBUG'
});

function logEvent(level, moduleName, action, message, details, durationMs) {
  const sheet = getOrCreateTechnicalLogSheet_();
  ensureTechnicalLogHeaders_(sheet);

  sheet.appendRow([
    new Date(),
    getCurrentUserEmail_(),
    String(level || LOG_LEVELS.INFO),
    String(moduleName || ''),
    String(action || ''),
    String(message || ''),
    serializeLogDetails_(details),
    durationMs === undefined || durationMs === null ? '' : Number(durationMs),
    CONFIG.VERSION
  ]);
}

function logInfo(moduleName, action, message, details, durationMs) {
  logEvent(LOG_LEVELS.INFO, moduleName, action, message, details, durationMs);
}

function logWarning(moduleName, action, message, details, durationMs) {
  logEvent(LOG_LEVELS.WARN, moduleName, action, message, details, durationMs);
}

function logError(moduleName, action, error, details, durationMs) {
  const normalizedError = normalizeError_(error);

  logEvent(
    LOG_LEVELS.ERROR,
    moduleName,
    action,
    normalizedError.message,
    Object.assign({}, details || {}, {
      name: normalizedError.name,
      stack: normalizedError.stack
    }),
    durationMs
  );
}

function getOrCreateTechnicalLogSheet_() {
  const spreadsheet = getInventorySpreadsheet_();
  let sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.TECH_LOG);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(CONFIG.SHEETS.TECH_LOG);
  }

  return sheet;
}

function ensureTechnicalLogHeaders_(sheet) {
  const headers = [[
    'Timestamp',
    'User',
    'Level',
    'Module',
    'Action',
    'Message',
    'Details',
    'Duration ms',
    'Version'
  ]];

  sheet.getRange(1, 1, 1, headers[0].length)
    .setValues(headers)
    .setFontWeight('bold')
    .setBackground('#d9ead3');

  sheet.setFrozenRows(1);
}

function getCurrentUserEmail_() {
  try {
    return Session.getActiveUser().getEmail() || 'unknown';
  } catch (error) {
    return 'unknown';
  }
}

function serializeLogDetails_(details) {
  if (details === undefined || details === null || details === '') return '';
  if (typeof details === 'string') return details;

  try {
    return JSON.stringify(details);
  } catch (error) {
    return String(details);
  }
}

function clearTechnicalLog() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Wyczysc log techniczny',
    'Czy na pewno usunac cala zawartosc arkusza Log techniczny?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  const sheet = getOrCreateTechnicalLogSheet_();
  sheet.clearContents();
  ensureTechnicalLogHeaders_(sheet);

  ui.alert('Log techniczny zostal wyczyszczony.');
}
