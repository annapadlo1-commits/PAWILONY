/**
 * Inventory PRO Enterprise
 * Pomiar wydajnosci.
 */

function measureExecution_(moduleName, action, callback) {
  const startedAt = Date.now();

  try {
    const result = callback();
    const durationMs = Date.now() - startedAt;

    logInfo(
      moduleName,
      action,
      'Pomiar wykonania',
      null,
      durationMs
    );

    return {
      result,
      durationMs
    };
  } catch (error) {
    const durationMs = Date.now() - startedAt;

    logError(
      moduleName,
      action,
      error,
      null,
      durationMs
    );

    throw error;
  }
}
