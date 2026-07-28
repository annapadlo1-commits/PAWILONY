/**
 * Inventory PRO 4.3.6 SAFE MODE — Health Check PAWILONÓW.
 */

function runEnterpriseHealthCheck() {
  const startedAt = Date.now();
  const validation = buildValidationReport_();
  const testReport = runEnterpriseTestsSilently_();

  let catalogSummary = null;
  let score = 100;
  const issues = [];

  try {
    catalogSummary = getProductCatalogSummary();
  } catch (error) {
    score -= 25;
    issues.push('Nie udało się zbudować katalogu.');
  }

  score -= validation.errors.length * 15;
  score -= validation.warnings.length * 5;
  score -= testReport.failed * 10;
  score = Math.max(0, Math.min(100, score));

  const formulas = validation.formulaAudit || {};
  const releaseGatePassed =
    score === 100 &&
    testReport.failed === 0 &&
    validation.errors.length === 0 &&
    Number(formulas.invalidFormulaCells || 0) === 0 &&
    Number(formulas.errorFormulaCells || 0) === 0;
  const status = releaseGatePassed ? 'HEALTHY' : (score >= 70 ? 'WARNING' : 'CRITICAL');
  const result = {
    version: CONFIG.VERSION,
    location: CONFIG.LOCATION.NAME,
    status: status,
    score: score,
    tests: {
      total: testReport.total,
      passed: testReport.passed,
      failed: testReport.failed,
      results: testReport.results
    },
    validation: {
      errors: validation.errors.length,
      warnings: validation.warnings.length
    },
    formulas: validation.formulaAudit,
    catalog: catalogSummary,
    issues: issues,
    releaseGatePassed: releaseGatePassed,
    durationMs: Date.now() - startedAt
  };

  logInfo('HealthCheck', 'runEnterpriseHealthCheck', 'Health Check: ' + status, result, result.durationMs);
  SpreadsheetApp.getUi().alert(
    'Inventory PRO - Health Check',
    formatHealthCheck_(result),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  return result;
}

function runEnterpriseTestsSilently_() {
  const tests = [
    testNormalizeText_,
    testNumberParser_,
    testCompoundNumberParser_,
    testHalfNumberParser_,
    testContinuousParser_,
    testParserPunctuation_,
    testParserLocations_,
    testAliasLoader_,
    testProductCatalog_,
    testGeminiLosslessPrompt500_,
    testMobileResolverAliases500_,
    testAudioProcessorTriggerDeduplication513_,
    testAudioPipelineRecovery514_,
    testInventoryStatusColorLifecycle514_,
    testInventoryFinishClearsConfiguredInputs515_,
    testInventoryColorLiveRefresh516_,
    testAudioScheduleIdempotency520_,
    testAudioStatusReadOnly520_,
    testAudioImmediateKick520_,
    testAudioParallelWorker521_,
    testDirectFinalAliases530_,
    testMobileImportReceipt530_,
    testProductManagerBulkArchive540_,
    testQuickInventoryRejectsEmptyList500_,
    testGeminiTransientRetry500_,
    testGeminiCompressedAudio510_,
    testExactMatcher_,
    testBusinessHistorySheet_,
    testImportAuditSheet_,
    testQualityWarning_,
    testAliasSuggestionCollection_,
    testPawilonyLayoutContract432_,
    testPawilonyHeaderDetection432_,
    testPawilonyMappingGuard432_,
    testPawilonyTargetColumnRouting432_,
    testNewProductRowRollback432_,
    testCanonicalInventoryFormulas434_,
    testFormulaContractFailClosed511_,
    testTareSafeCategoryReview462_,
    testProductPackagingProfiles470_,
    testMissingPackagingReview511_,
    testDirectFinalCoffeeException434_,
    testRecoveryDictionaryContaminationGuard_,
    testFormulaRepairClassifiedPlan514_,
    testFormulaRepairConcurrency432_,
    testFormulaConflictClassification432_
  ];

  const results = tests.map(runSingleTest_);
  return {
    total: results.length,
    passed: results.filter(result => result.passed).length,
    failed: results.filter(result => !result.passed).length,
    results: results
  };
}

function formatHealthCheck_(result) {
  let message =
    'Lokal: ' + result.location + '\n' +
    'Status: ' + result.status + '\n' +
    'Application Health: ' + result.score + '%\n\n' +
    'Bramka wydania: ' + (result.releaseGatePassed ? 'PASS' : 'FAIL') + '\n' +
    'Testy PASS: ' + result.tests.passed + '/' + result.tests.total + '\n' +
    'Błędy konfiguracji: ' + result.validation.errors + '\n' +
    'Ostrzeżenia: ' + result.validation.warnings + '\n';

  if (result.catalog) {
    message +=
      '\nProdukty: ' + result.catalog.products +
      '\nAliasy: ' + result.catalog.aliases +
      '\nNORMAL: ' + result.catalog.normal +
      '\nKEG: ' + result.catalog.keg +
      '\nLOCATION: ' + result.catalog.location +
      '\nBrak wiersza: ' + result.catalog.missingInventoryRow;
  }

  if (result.formulas) {
    message +=
      '\n\nFormuły oczekiwane: ' + result.formulas.expectedFormulaCells +
      '\nBrakujące: ' + result.formulas.missingFormulaCells +
      '\nSpłaszczone: ' + result.formulas.flattenedFormulaCells +
      '\nKonflikty: ' + result.formulas.conflictFormulaCells +
      '\nStarsze poprawne (+): ' + result.formulas.legacyFormulaCells +
      '\nNieprawidłowe: ' + result.formulas.invalidFormulaCells +
      '\nBłędy obliczeń: ' + result.formulas.errorFormulaCells;
  }

  if (result.tests.failed) {
    const failed = result.tests.results.filter(item => !item.passed).slice(0, 8);
    message += '\n\nTesty FAIL:\n- ' + failed.map(item => item.name + ': ' + item.message).join('\n- ');
  }
  return message;
}
