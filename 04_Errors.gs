/**
 * Inventory PRO Enterprise
 * Centralna obsluga bledow.
 */

function runSafely_(moduleName, action, callback, userMessage) {
  const startedAt = Date.now();

  try {
    const result = callback();
    const durationMs = Date.now() - startedAt;

    logInfo(
      moduleName,
      action,
      'Zakonczono poprawnie',
      null,
      durationMs
    );

    return result;
  } catch (error) {
    const durationMs = Date.now() - startedAt;

    logError(
      moduleName,
      action,
      error,
      null,
      durationMs
    );

    if (userMessage) {
      SpreadsheetApp.getUi().alert(
        userMessage + '\n\n' + normalizeError_(error).message
      );
    }

    throw error;
  }
}

function normalizeError_(error) {
  if (error instanceof Error) {
    return {
      name: error.name || 'Error',
      message: error.message || String(error),
      stack: error.stack || ''
    };
  }

  return {
    name: 'Error',
    message: String(error),
    stack: ''
  };
}

function assertCondition_(condition, message) {
  if (!condition) {
    throw new Error(message || 'Warunek walidacji nie zostal spelniony.');
  }
}
