/**
 * Inventory PRO Enterprise v2.1.3 Recovery
 * Testy regresyjne.
 */

function runAllEnterpriseTests() {
  const startedAt = Date.now();

  const tests = [
    testNormalizeText_,
    testNumberParser_,
    testCompoundNumberParser_,
    testHalfNumberParser_,
    testContinuousParser_,
    testNumericProductNamesParser_,
    testNoProductPrefixSkipping_,
    testAttachedNumberProductParser_,
    testParserPunctuation_,
    testParserLocations_,
    testParserLucanoZero_,
    testParserKolaZeroVariant_,
    testParserOrdinaryZeroValue_,
    testParserHalfLitreName_,
    testParserLongestAgeName_,
    testAliasLoader_,
    testProductCatalog_,
    testExactMatcher_,
    testBusinessHistorySheet_,
    testImportAuditSheet_,
    testValidationReport_,
    testQualityWarning_,
    testQualityZeroWarning_,
    testDuplicateAnnotation_,
    testAliasSuggestionCollection_,
    testRecognitionNumericSafety_,
    testRecognitionAttachedVariant_,
    testRecognitionAmbiguousFamily_,
    testAliasConflictProtection_,
    testNewProductNameValidation_,
    testProductResolverData_,
    testBusinessCategoryNormalization_,
    testNewProductAuditKey_,
    testSummaryColumnMapping284_,
    testSummaryInputColumns284_,
    testRuntimeColumnNormalization501_,
    testAudioQueueWatchdog501_,
    testGeminiCompressedAudio510_,
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
    testGeminiTransientRetry500_,
    testClosedBusinessCategories284_,
    testPawilonyLayoutContract432_,
    testPawilonyLocationMapping432_,
    testGenericBeerClassification432_,
    testPawilonyHeaderDetection432_,
    testPawilonyMappingGuard432_,
    testPawilonyTargetColumnRouting432_,
    testNewProductRowRollback432_,
    testSparseWritePlan432_,
    testSparseRollback432_,
    testFormulaWriteGuard432_,
    testCanonicalInventoryFormulas434_,
    testFormulaContractFailClosed511_,
    testTareSafeCategoryReview462_,
    testProductPackagingProfiles470_,
    testMissingPackagingReview511_,
    testProductCreationAndCapacityGuard484_,
    testProductManagerFallbackRow485_,
    testDirectFinalCoffeeException434_,
    testDirectFinalProductManagerMapping4313_,
    testDirectFinalSparseWriteGuard461_,
    testXlsxExportWithoutDriveApi4313_,
    testRecoveryDictionaryContaminationGuard_,
    testFormulaRepairClassifiedPlan514_,
    testFormulaRepairSegments432_,
    testFormulaRepairConcurrency432_,
    testFormulaConflictClassification432_,
    testUndoConflictEvaluation432_
  ];

  const results = tests.map(runSingleTest_);
  const passed = results.filter(
    result => result.passed
  ).length;
  const failed = results.length - passed;

  const report = {
    total: results.length,
    passed: passed,
    failed: failed,
    durationMs: Date.now() - startedAt,
    results: results
  };

  logInfo(
    'Tests',
    'runAllEnterpriseTests',
    failed === 0
      ? 'Wszystkie testy PASS'
      : 'Wykryto testy FAIL',
    report,
    report.durationMs
  );

  SpreadsheetApp.getUi().alert(
    'Inventory PRO - Testy',
    formatTestReport_(report),
    SpreadsheetApp.getUi().ButtonSet.OK
  );

  return report;
}

function runSingleTest_(testFunction) {
  const startedAt = Date.now();

  try {
    testFunction();

    return {
      name: testFunction.name,
      passed: true,
      durationMs: Date.now() - startedAt,
      message: 'PASS'
    };

  } catch (error) {
    return {
      name: testFunction.name,
      passed: false,
      durationMs: Date.now() - startedAt,
      message: normalizeError_(error).message
    };
  }
}

function formatTestReport_(report) {
  let message =
    'Testy: ' + report.total + '\n' +
    'PASS: ' + report.passed + '\n' +
    'FAIL: ' + report.failed + '\n' +
    'Czas: ' + report.durationMs + ' ms';

  const failedTests = report.results.filter(
    result => !result.passed
  );

  if (failedTests.length) {
    message += '\n\nBLEDY:\n';

    failedTests.forEach(result => {
      message +=
        '- ' +
        result.name +
        ': ' +
        result.message +
        '\n';
    });
  }

  return message;
}

function testNormalizeText_() {
  assertCondition_(
    normalizeText('  Żółta, Kola  ') ===
      'zolta kola',
    'normalizeText nie dziala poprawnie.'
  );
}

function testNumberParser_() {
  assertCondition_(
    parseNumberText_('trzy') === 3,
    'trzy powinno dac 3'
  );

  assertCondition_(
    parseNumberText_('0,52') === 0.52,
    '0,52 powinno dac 0.52'
  );

  assertCondition_(
    parseNumberText_('1,212') === 1.212,
    '1,212 powinno dac 1.212'
  );
}

function testCompoundNumberParser_() {
  assertCondition_(
    parseNumberText_(
      'dwadziescia siedem'
    ) === 27,
    'dwadziescia siedem powinno dac 27'
  );
}

function testHalfNumberParser_() {
  assertCondition_(
    parseNumberText_('pol') === 0.5,
    'pol powinno dac 0.5'
  );

  assertCondition_(
    parseNumberText_('poltora') === 1.5,
    'poltora powinno dac 1.5'
  );
}

function testContinuousParser_() {
  const context = createParserTestContext_(['Martini Bitter', 'Godet VSOP']);
  const parsed = parseInventoryText(
    'Martini Bitter 1,212 Godet VSOP 1,046',
    context
  );

  assertCondition_(
    parsed.length === 2,
    'Powinny byc 2 pozycje, sa: ' +
      parsed.length
  );

  assertCondition_(
    parsed[0].product === 'Martini Bitter' &&
    parsed[1].product === 'Godet VSOP',
    'Nazwy zostaly zle rozdzielone.'
  );
}

function testParserPunctuation_() {
  const context = createParserTestContext_(['Litowel Ciemny Lager', 'Inne Beczki Jungle IPA']);
  const parsed = parseInventoryText(
    'darkroom Litowel Ciemny Lager 24. magazyn Inne Beczki Jungle IPA 20',
    context
  );

  assertCondition_(
    parsed.length === 2,
    'Powinny byc 2 pozycje.'
  );

  assertCondition_(
    parsed[0].value === 24 &&
    parsed[1].value === 20,
    'Interpunkcja zepsula liczby.'
  );
}

function testParserLocations_() {
  const firstArea = CONFIG.LOCATION_AREAS[0];
  const secondArea = CONFIG.LOCATION_AREAS[1] || firstArea;
  const context = createParserTestContext_(['Litowel Pomelo', 'Jurajska Pomarancza']);
  const parsed = parseInventoryText(
    firstArea.aliases[0] + ' Litowel Pomelo 30 Litowel Pomelo 2 ' +
      secondArea.aliases[0] + ' Jurajska Pomarancza 120',
    context
  );

  assertCondition_(
    parsed.length === 3,
    'Powinny byc 3 pozycje.'
  );

  assertCondition_(
    parsed[0].location === firstArea.key &&
    parsed[1].location === firstArea.key &&
    parsed[2].location === secondArea.key,
    'Lokalizacje sekcyjne sa bledne.'
  );
}

function testLocationPrefixProductNameIntegrity_() {
  const context = createParserTestContext_(['Inne Beczki', 'Cieszyn Pilsner']);
  const parsed = parseInventoryText('magazyn Inne beczki Pilsner 2', context);
  assertCondition_(parsed.length === 1,
    'Nazwa po lokalizacji nie może zostać rozbita na dwa wpisy: ' + JSON.stringify(parsed));
  assertCondition_(normalizeText(parsed[0].product) === 'inne beczki pilsner',
    'Parser zgubił początek nazwy produktu: ' + parsed[0].product);
  assertCondition_(parsed[0].value === 2 && parsed[0].location === 'magazyn',
    'Parser zgubił wartość lub lokalizację.');
  const recognition = matchProductForParser_(parsed[0].product, context);
  assertCondition_(!recognition.match.matched,
    'Nie wolno automatycznie przypisać skróconej nazwy do Cieszyn Pilsner.');

  const exactContext = createParserTestContext_([
    'Inne Beczki', 'Inne Beczki Pilsner 0,5L', 'Cieszyn Pilsner'
  ]);
  const exact = parseInventoryText('magazyn Inne beczki Pilsner 2', exactContext);
  assertCondition_(exact.length === 1 &&
      normalizeText(exact[0].product) === normalizeText('Inne Beczki Pilsner 0,5L'),
    'Pełna pozycja Inne Beczki Pilsner nie została wybrana z katalogu.');
}

function testParserLucanoZero_() {
  const context = createParserTestContext_([
    'Amaro Lucano 1L',
    'Amaro Lucano 0% 0,7L'
  ]);
  const parsed = parseInventoryText('amaro lucano zero 1,234', context);
  assertCondition_(parsed.length === 1, 'Lucano zero powinno byc jedna pozycja.');
  assertCondition_(normalizeText(parsed[0].product) === normalizeText('Amaro Lucano 0% 0,7L'), 'Lucano zero wybralo zly produkt.');
  assertCondition_(parsed[0].value === 1.234, 'Lucano zero ma zla wartosc.');
}

function testParserKolaZeroVariant_() {
  const context = createParserTestContext_([
    'Fritz 200ml KOLA',
    'Fritz 200ml KOLA BEZ CUKRU'
  ]);
  const parsed = parseInventoryText('kola 0 12', context);
  assertCondition_(parsed.length === 1, 'Kola 0 12 powinna byc jedna pozycja.');
  assertCondition_(normalizeText(parsed[0].product) === normalizeText('Fritz 200ml KOLA BEZ CUKRU'), 'Kola 0 nie wybrala wariantu bez cukru.');
  assertCondition_(parsed[0].value === 12, 'Kola 0 12 ma zla wartosc.');
}

function testParserOrdinaryZeroValue_() {
  const context = createParserTestContext_(['Fritz 200ml KOLA']);
  const product = context.catalog[0];
  product.aliases.push('fritz kola');
  context.parserPhraseIndex = null;
  const parsed = parseInventoryText('fritz kola 0', context);
  assertCondition_(parsed.length === 1, 'Fritz kola 0 powinna byc jedna pozycja.');
  assertCondition_(parsed[0].value === 0, 'Zwykle zero powinno pozostac wartoscia.');
}

function testParserHalfLitreName_() {
  const context = createParserTestContext_(['Żubrówka Bison Grass 0,5L']);
  const parsed = parseInventoryText('żubrówka bison grass pół litra 1,234', context);
  assertCondition_(parsed.length === 1, 'Pol litra powinno byc jedna pozycja.');
  assertCondition_(normalizeText(parsed[0].product) === normalizeText('Żubrówka Bison Grass 0,5L'), 'Pol litra wybralo zly produkt.');
  assertCondition_(parsed[0].value === 1.234, 'Pol litra ma zla wartosc.');
}

function testParserLongestAgeName_() {
  const context = createParserTestContext_(['Osco', 'Osco 2 years old']);
  const parsed = parseInventoryText('Osco 2 years old 22', context);
  assertCondition_(parsed.length === 1, 'Osco 2 years old powinno byc jedna pozycja.');
  assertCondition_(normalizeText(parsed[0].product) === normalizeText('Osco 2 years old'), 'Parser wybral krotsze Osco.');
  assertCondition_(parsed[0].value === 22, 'Osco 2 years old ma zla wartosc.');
}

function testAliasLoader_() {
  const aliases = loadAliases();

  assertCondition_(
    aliases &&
    typeof aliases === 'object',
    'loadAliases nie zwraca obiektu.'
  );
}

function testProductCatalog_() {
  const summary = getProductCatalogSummary();

  assertCondition_(
    summary.products > 0,
    'Katalog produktow jest pusty.'
  );
}

function testExactMatcher_() {
  const context = buildRuntimeContext_();

  if (!context.catalog.length) {
    throw new Error(
      'Brak produktow do testu matcher.'
    );
  }

  const sample = context.catalog[0];

  const result = matchProduct(
    sample.name,
    context
  );

  assertCondition_(
    result.matched &&
    result.status === 'EXACT',
    'Dokladna nazwa nie daje EXACT.'
  );
}

function testBusinessHistorySheet_() {
  const sheet =
    getOrCreateBusinessHistorySheet_();

  assertCondition_(
    normalizeText(sheet.getName()) ===
      normalizeText(CONFIG.SHEETS.HISTORY),
    'Nieprawidlowy arkusz historii.'
  );
}

function testImportAuditSheet_() {
  const sheet = getOrCreateTechnicalHistorySheet_();

  assertCondition_(
    normalizeText(
      sheet.getRange(1, 1).getDisplayValue()
    ) === 'import id',
    'Audyt importow nie ma poprawnych naglowkow.'
  );
  assertCondition_(
    normalizeText(sheet.getRange(1, 17).getDisplayValue()) === 'duplicate count',
    'Audyt importow nie zawiera danych o duplikatach.'
  );
}

function testValidationReport_() {
  const result = buildValidationReport_();

  assertCondition_(
    result &&
    Array.isArray(result.errors),
    'Raport walidacji ma zla strukture.'
  );
}


function testQualityWarning_() {
  const product = {
    type: CONFIG.PRODUCT_TYPES.NORMAL
  };

  const result = evaluateImportQuality_(
    product,
    21,
    '',
    {
      warnZero: true,
      blockNegative: true,
      normalWholeWarning: 20,
      normalWeightWarning: 20,
      kegWholeWarning: 20,
      kegWeightWarning: 100,
      locationWarning: 500,
      duplicateWarningCount: 2
    }
  );

  assertCondition_(
    result.warning === true,
    'Wartosc 21 dla NORMAL powinna generowac ostrzezenie.'
  );
}


function testQualityZeroWarning_() {
  const result = evaluateImportQuality_(
    { type: CONFIG.PRODUCT_TYPES.NORMAL },
    0,
    '',
    {
      warnZero: true,
      blockNegative: true,
      normalWholeWarning: 20,
      normalWeightWarning: 20,
      kegWholeWarning: 20,
      kegWeightWarning: 100,
      locationWarning: 500,
      duplicateWarningCount: 2
    }
  );

  assertCondition_(
    result.warning === true &&
    result.flags.includes('WPROWADZONO_ZERO'),
    'Wartosc 0 powinna generowac flage WPROWADZONO_ZERO.'
  );
}

function testDuplicateAnnotation_() {
  const results = [
    {
      saved: true,
      product: 'Bacardi 8',
      row: 10,
      column: 'C',
      addedValue: 1.2,
      message: 'Zapisano'
    },
    {
      saved: true,
      product: 'Bacardi 8',
      row: 10,
      column: 'C',
      addedValue: 0.8,
      message: 'Zapisano'
    }
  ];

  annotateSavedDuplicateResults_(results, { duplicateWarningCount: 2 });

  assertCondition_(
    results[0].duplicateWarning === true &&
    results[0].duplicateCount === 2 &&
    Math.abs(results[0].duplicateTotal - 2) < 0.000001,
    'Dwa wpisy do tej samej komorki powinny zostac oznaczone jako duplikat z suma 2.'
  );
}

function testAliasSuggestionCollection_() {
  const result = collectAliasSuggestions_([
    {
      include: true,
      learnAlias: true,
      parsedProduct: 'Thomas Henry tonik',
      selectedProduct: 'Thomas Henry Tonic'
    },
    {
      include: true,
      learnAlias: false,
      parsedProduct: 'Fritz',
      selectedProduct: 'Fritz Kola'
    }
  ]);

  assertCondition_(
    result.length === 1 &&
    result[0].alias ===
      'Thomas Henry tonik',
    'Auto Learning nie zebral poprawnego aliasu.'
  );
}


function testNumericProductNamesParser_() {
  const context = createParserTestContext_([
    'Jameson 0,7L',
    'Auchentoshan 12yo',
    'Auchentoshan 18yo',
    'Auchentoshan Three Wood',
    'Bacardi 4',
    'Bacardi 8',
    'Bacardi 10',
    'Osco',
    'Osco 2 years old'
  ]);

  const parsed = parseInventoryText(
    'Jameson 12 Auchentoshan 12 1,234 Auchentoshan 18 1,234 Bacardi 4 0,890 Bacardi 8 0,987 Bacardi 10 1,123 Osco 2 years 22',
    context
  );

  assertCondition_(
    parsed.length === 7,
    'Nazwy z cyframi: powinno byc 7 pozycji, sa: ' + parsed.length
  );

  const expected = [
    ['Jameson 0,7L', 12],
    ['Auchentoshan 12yo', 1.234],
    ['Auchentoshan 18yo', 1.234],
    ['Bacardi 4', 0.89],
    ['Bacardi 8', 0.987],
    ['Bacardi 10', 1.123],
    ['Osco 2 years old', 22]
  ];

  expected.forEach((item, index) => {
    assertCondition_(
      normalizeText(parsed[index].product) === normalizeText(item[0]) &&
      parsed[index].value === item[1],
      'Nieprawidlowy podzial pozycji ' + (index + 1) +
      ': ' + JSON.stringify(parsed[index])
    );
  });
}

function testAttachedNumberProductParser_() {
  const context = createParserTestContext_([
    'Bacardi 8',
    'Bacardi 10'
  ]);

  const parsed = parseInventoryText(
    'bacardi8 0,987 bacardi10 1,123',
    context
  );

  assertCondition_(
    parsed.length === 2 &&
    normalizeText(parsed[0].product) === 'bacardi 8' &&
    parsed[0].value === 0.987 &&
    normalizeText(parsed[1].product) === 'bacardi 10' &&
    parsed[1].value === 1.123,
    'Nazwy sklejone z cyfra zostaly zle rozpoznane.'
  );
}

function createParserTestContext_(names) {
  const catalog = names.map(name => ({
    name: name,
    normalizedName: normalizeText(name),
    aliases: [],
    type: 'NORMAL',
    category: 'TEST',
    columns: {},
    active: true
  }));

  // RC3.3: parser 3.0 korzysta z tych samych indeksów co silnik
  // produkcyjny. Stary fixture testowy przekazywał tylko dwa indeksy,
  // przez co testy kończyły się TypeError przed sprawdzeniem parsera.
  const productIndex = {};
  catalog.forEach(product => {
    productIndex[product.normalizedName] = product;
  });

  return {
    catalog: catalog,
    productIndex: productIndex,
    aliasIndex: {},
    exactRecognitionIndex: buildExactRecognitionIndex_(catalog),
    technicalRecognitionIndex: buildTechnicalRecognitionIndex_(catalog),
    firstTokenIndex: buildFirstTokenRecognitionIndex_(catalog),
    recognitionMemo: {}
  };
}


function testNoProductPrefixSkipping_() {
  const context = createParserTestContext_([
    'Jameson 0,7L',
    'Auchentoshan 12yo',
    'Auchentoshan 18yo'
  ]);

  const parsed = parseInventoryText(
    'Jameson 12 Auchentoshan 12 1,234 Auchentoshan 18 1,234',
    context
  );

  assertCondition_(
    parsed.length === 3,
    'Parser polaczyl sasiednie produkty. Pozycji: ' + parsed.length
  );

  assertCondition_(
    normalizeText(parsed[0].product) === normalizeText('Jameson 0,7L') &&
    parsed[0].value === 12 &&
    normalizeText(parsed[1].product) === normalizeText('Auchentoshan 12yo') &&
    parsed[1].value === 1.234 &&
    normalizeText(parsed[2].product) === normalizeText('Auchentoshan 18yo') &&
    parsed[2].value === 1.234,
    'Parser pominal poczatek pozycji lub polaczyl produkty: ' +
      JSON.stringify(parsed)
  );
}


function testRecognitionNumericSafety_() {
  const context = createRecognitionTestContext_(['Bacardi 4', 'Bacardi 8', 'Bacardi 10']);
  const result = recognizeProduct_('bakardi 8', context);
  assertCondition_(result.matched, 'Bakardi 8 powinno zostac rozpoznane');
  assertCondition_(result.product && result.product.name === 'Bacardi 8', 'Nie wolno pomylic liczby wariantu');
}

function testRecognitionAttachedVariant_() {
  const context = createRecognitionTestContext_(['Bacardi 8']);
  const result = recognizeProduct_('Bacardi8', context);
  assertCondition_(result.matched, 'Bacardi8 powinno byc wariantem technicznym');
  assertCondition_(result.product && result.product.name === 'Bacardi 8', 'Nieprawidlowy wariant sklejony');
}

function testRecognitionAmbiguousFamily_() {
  const context = createRecognitionTestContext_(['Bacardi 4', 'Bacardi 8', 'Bacardi 10']);
  const result = recognizeProduct_('Bacardi', context);
  assertCondition_(result.status === 'AMBIGUOUS', 'Sama marka musi wymagac wyboru');
}

function testAliasConflictProtection_() {
  const context = createRecognitionTestContext_(['Bacardi 4', 'Bacardi 8']);
  const result = validateAliasSuggestion_({ alias: 'Bacardi', product: 'Bacardi 8' }, context, {});
  assertCondition_(!result.valid, 'Zbyt ogolny alias powinien byc zablokowany');
}

function createRecognitionTestContext_(names) {
  const catalog = names.map(name => ({
    name: name,
    normalizedName: normalizeText(name),
    aliases: [],
    type: 'NORMAL',
    category: 'TEST'
  }));
  const productIndex = {};
  catalog.forEach(product => productIndex[product.normalizedName] = product);
  return {
    catalog: catalog,
    productIndex: productIndex,
    aliasIndex: {},
    exactRecognitionIndex: buildExactRecognitionIndex_(catalog),
    technicalRecognitionIndex: buildTechnicalRecognitionIndex_(catalog),
    firstTokenIndex: buildFirstTokenRecognitionIndex_(catalog),
    recognitionMemo: {}
  };
}


function testNewProductNameValidation_() {
  validateNewProductName_('Nowy Produkt Testowy');
  let failed = false;
  try { validateNewProductName_('123'); } catch (error) { failed = true; }
  assertCondition_(failed, 'Nazwa skladajaca sie tylko z liczby powinna byc odrzucona.');
}

function testProductResolverData_() {
  const context = buildRuntimeContext_();
  if (!context.catalog.length) throw new Error('Brak produktow do testu resolvera.');
  const data = getProductResolverData(context.catalog[0].name);
  assertCondition_(data && Array.isArray(data.products) && data.products.length > 0, 'Resolver nie zwraca katalogu produktow.');
  assertCondition_(Array.isArray(data.suggestions) && data.suggestions.length > 0, 'Resolver nie zwraca sugestii.');
}

function testProductManagerPayloadValidation_() {
  const valid = validateProductManagerPayload_({
    originalName: 'Test Product',
    name: 'Test Product',
    type: 'NORMAL',
    category: 'TEST',
    active: true,
    columns: { quantity: 'H', weight: 'C' }
  });
  assertCondition_(valid.type === 'NORMAL' && valid.columns.quantity === 'H', 'Walidacja Product Manager zwrocila zle dane.');

  let failed = false;
  try {
    validateProductManagerPayload_({
      originalName: 'Test Product',
      name: 'Test Product',
      type: 'LOCATION',
      category: 'TEST',
      active: true,
      columns: {}
    });
  } catch (error) {
    failed = true;
  }
  assertCondition_(failed, 'LOCATION bez kolumn lokalizacji powinien byc odrzucony.');
}


function testBusinessCategoryNormalization_() {
  assertCondition_(normalizeBusinessCategory_('SOFTY na szt. MAGAZYN SZT DARKROOM SZT LODÓWKI SZT razem') === 'SOFTY',
    'Kategoria SOFTY nie zostala znormalizowana.');
  assertCondition_(normalizeBusinessCategory_('PIWO BUTELKI MAGAZYN SZT DARKROOM SZT LODÓWKI SZT RAZEM') === 'PIWO BUTELKI',
    'Kategoria PIWO BUTELKI nie zostala znormalizowana.');
}

function testNewProductAuditKey_() {
  const key = buildNewProductAuditKey_('Thomas Henry Arbuz', 273);
  assertCondition_(key === 'thomas henry arbuz|273', 'Nieprawidlowy klucz audytu nowego produktu.');
}


function testAliasManagerPayloadValidation_() {
  const products = loadAllProductConfigurationsForManager_();
  if (!products.length) throw new Error('Brak produktu do testu Alias Managera.');
  const result = validateAliasManagerPayload_({
    alias: 'alias testowy inventory pro',
    product: products[0].name
  }, false);
  assertCondition_(
    result.alias === 'alias testowy inventory pro' && result.product === products[0].name,
    'Walidacja Alias Managera zwrocila nieprawidlowy wynik.'
  );
}

function testAliasRowsForManager_() {
  const rows = loadAliasRowsForManager_();
  assertCondition_(Array.isArray(rows), 'Lista aliasow Alias Managera nie jest tablica.');
}

function testCategoryService282_() {
  assertCondition_(normalizeBusinessCategory_('SOFTY na szt. MAGAZYN SZT DARKROOM SZT LODÓWKI SZT razem') === 'SOFTY', 'Kategoria SOFTY powinna zostać oczyszczona.');
  assertCondition_(normalizeBusinessCategory_('WÓDKA') === 'WÓDKA', 'Kategoria WÓDKA powinna zachować polski zapis.');
}

function testFinalReviewSources282_() {
  const sources = getFinalReviewSources();
  assertCondition_(Array.isArray(sources), 'Lista arkuszy eksportu powinna być tablicą.');
  assertCondition_(sources.some(item => item.current), 'Lista powinna zawierać bieżącą inwentaryzację.');
}


function testSummaryColumnMapping284_() {
  ['NORMAL', 'KEG', 'LOCATION'].forEach(function(type) {
    const expected = getConfiguredInventoryLayout_(type);
    const actual = getInventorySummaryLayout_(type);
    assertCondition_(actual.finalTotal === expected.finalTotal,
      type + ' musi korzystać z kolumny finalnej skonfigurowanej dla tego lokalu.');
  });
}

function testSummaryInputColumns284_() {
  ['NORMAL', 'KEG', 'LOCATION'].forEach(function(type) {
    const expected = getConfiguredInventoryLayout_(type);
    const actual = getInventorySummaryLayout_(type);
    ['grossWeight', 'fullUnits', 'warehouse', 'darkroom', 'fridges'].forEach(function(field) {
      assertCondition_((actual[field] || '') === (expected[field] || ''),
        'Nieprawidłowa kolumna ' + field + ' dla ' + type + '.');
    });
  });
}

function testRuntimeColumnNormalization501_() {
  ['NORMAL', 'KEG', 'LOCATION'].forEach(function(type) {
    const expected = getInputColumnsForProductType_(type);
    const stale = { quantity: 'C', weight: 'C', warehouse: 'C', darkroom: 'C', fridges: 'C' };
    const actual = normalizeRuntimeProductColumns_(type, stale, { name: 'TEST', type: type });
    Object.keys(expected).forEach(function(field) {
      assertCondition_((actual[field] || '') === (expected[field] || ''),
        type + ' musi odrzucić nieprawidłowe historyczne mapowanie pola ' + field + '.');
    });
  });
}

function testAudioQueueWatchdog501_() {
  const now = Date.now();
  const stale = getInventoryAudioRecoveryDecision_({
    id: 'TEST',
    status: 'PROCESSING',
    attempts: 0,
    createdAt: now - GEMINI_AUDIO_PROCESSING_TIMEOUT_MS_ - 1000,
    updatedAt: now - GEMINI_AUDIO_PROCESSING_TIMEOUT_MS_ - 1000
  }, now);
  assertCondition_(stale.status === 'QUEUED' && stale.attempts === 1,
    'Przerwana transkrypcja musi automatycznie wrócić do kolejki.');

  const fresh = getInventoryAudioRecoveryDecision_({
    id: 'TEST',
    status: 'PROCESSING',
    attempts: 0,
    createdAt: now,
    updatedAt: now
  }, now);
  assertCondition_(fresh.status === 'PROCESSING' && fresh.attempts === 0,
    'Aktywna transkrypcja nie może zostać przedwcześnie wznowiona.');
}

function testClosedBusinessCategories284_() {
  const allowed = getBusinessCategories();
  assertCondition_(allowed.indexOf('WINO') >= 0 && allowed.indexOf('SOFTY') >= 0, 'Brak podstawowych kategorii biznesowych.');
  assertCondition_(normalizeBusinessCategory_('weissburgunder 0 0,75 0 0') === '', 'Tekst produktu nie może zostać kategorią.');
}


function testZeroProtectedProductNames292_() {
  const context = buildRuntimeContext_();
  const lucano = parseInventoryText('amaro lucano zero 3', context);
  assertCondition_(lucano.length === 1 && Number(lucano[0].value) === 3, 'Lucano zero powinno zachować wartość 3.');
  assertCondition_(normalizeText(lucano[0].product).indexOf('zero') >= 0 || normalizeText(lucano[0].product).indexOf('0') >= 0, 'Zero powinno pozostać częścią nazwy Lucano.');

  const cola = parseInventoryText('lodowki fritz kola zero 4 lodowki fritz jablko 7', context);
  assertCondition_(cola.length === 2, 'Parser powinien zwrócić dwie pozycje bez przesunięcia wartości.');
  assertCondition_(Number(cola[0].value) === 4 && Number(cola[1].value) === 7, 'Wartości 4 i 7 nie mogą przeciekać między wpisami.');
}

function testNumericZeroStillWorks292_() {
  const context = buildRuntimeContext_();
  const parsed = parseInventoryText('fritz kola 0', context);
  assertCondition_(parsed.length === 1 && Number(parsed[0].value) === 0, 'Cyfrowe 0 musi nadal działać jako rzeczywista wartość.');
}


/** v2.10.1: regresje parsera nazw zawierajacych zero. */
function runZeroSafeParserRegressionTests2101() {
  const cases = [
    { input: 'kola zero 1', expectedValue: 1, expectedCount: 1 },
    { input: 'lucano zero 2', expectedValue: 2, expectedCount: 1 },
    { input: 'fritz kola 0', expectedValue: 0, expectedCount: 1 }
  ];
  const context = buildRuntimeContext_();
  const failures = [];

  cases.forEach(testCase => {
    const result = parseInventoryText(testCase.input, context);
    if (result.length !== testCase.expectedCount ||
        !result[0] || Number(result[0].value) !== testCase.expectedValue) {
      failures.push(testCase.input + ' -> ' + JSON.stringify(result));
    }
  });

  const message = failures.length
    ? 'FAIL (' + failures.length + ')\n' + failures.join('\n')
    : 'PASS: ' + cases.length + ' testy Zero-Safe Parser';
  SpreadsheetApp.getUi().alert('Inventory PRO 2.10.1 — Zero-Safe Tests', message, SpreadsheetApp.getUi().ButtonSet.OK);
  if (failures.length) throw new Error(message);
  return { passed: cases.length, failed: 0 };
}

/** v2.10.3 — czytelne testy ZERO: tokenizer + pełny pipeline. */
function runZeroImportRegressionTests2104() {
  const cases = [
    { input: 'kola zero 1', tokenName: 'kola zero', product: 'fritz 200ml kola bez cukru', value: 1 },
    { input: 'lucano zero 2', tokenName: 'lucano zero', product: 'amaro lucano 0% 0,7l', value: 2 },
    { input: 'fritz kola 0', tokenName: 'fritz kola', product: 'fritz 200ml kola', value: 0 }
  ];
  const context = buildRuntimeContext_();
  const results = cases.map(function(testCase) {
    const parsed = parseInventoryText(testCase.input, context);
    const first = parsed[0] || {};
    const tokenName = normalizeText(first.product || '');
    const tokenPassed = parsed.length === 1 &&
      tokenName.indexOf(normalizeText(testCase.tokenName)) !== -1 &&
      Number(first.value) === Number(testCase.value);

    const canonical = getCanonicalParserProductForTest_(first.product || '', context);
    const pipelinePassed = canonical.indexOf(normalizeText(testCase.product)) !== -1 &&
      Number(first.value) === Number(testCase.value);

    return {
      input: testCase.input,
      tokenPassed: tokenPassed,
      pipelinePassed: pipelinePassed,
      actualToken: (first.product || 'BRAK') + ' | ' + String(first.value),
      actualProduct: canonical || 'BRAK'
    };
  });

  const tokenPass = results.filter(function(r) { return r.tokenPassed; }).length;
  const pipelinePass = results.filter(function(r) { return r.pipelinePassed; }).length;
  const allPassed = tokenPass === results.length && pipelinePass === results.length;
  const lines = results.map(function(result) {
    return (result.tokenPassed ? '✓' : '✗') + ' tokenizer: ' + result.input +
      ' → ' + result.actualToken + '\n' +
      (result.pipelinePassed ? '✓' : '✗') + ' produkt: ' + result.actualProduct;
  }).join('\n\n');

  SpreadsheetApp.getUi().alert(
    'Parser ZERO — testy 2.10.5',
    'Tokenizer PASS: ' + tokenPass + '/' + results.length + '\n' +
      'Pełny pipeline PASS: ' + pipelinePass + '/' + results.length + '\n\n' + lines,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  if (!allPassed) throw new Error('Nie wszystkie testy ZERO 2.10.5 przeszły.');
  return results;
}

/** Zachowuje zgodność ze starszymi pozycjami menu. */
function runZeroImportRegressionTests2103() {
  return runZeroImportRegressionTests2104();
}

function runZeroImportRegressionTests2102() {
  return runZeroImportRegressionTests2104();
}

function getCanonicalParserProductForTest_(recognitionInput, context) {
  const parserMatch = matchProductForParser_(recognitionInput, context);
  const match = parserMatch && parserMatch.match;
  if (!match) return '';
  if (match.product && match.product.name) return normalizeText(match.product.name);
  if (match.candidates && match.candidates.length &&
      match.candidates[0].product && match.candidates[0].product.name) {
    return normalizeText(match.candidates[0].product.name);
  }
  return '';
}

/** v2.10.3 — nazwy numeryczne i słowne pojemności. */
function runNumericProductNameRegressionTests2104() {
  const cases = [
    { input: 'okentoshan 18 1.408', productContains: 'auchentoshan 18', value: 1.408 },
    { input: '1.407 ardbeg 10', productContains: 'ardbeg 10', value: 1.407 },
    { input: 'bacardi 8 0,987', productContains: 'bacardi 8', value: 0.987 },
    { input: 'singleton 15 0.588', productContains: 'singleton 15', value: 0.588 },
    { input: 'ron zacapa 23 0.335', productContains: 'ron zacapa 23', value: 0.335 },
    { input: 'bison grass pół litra 1,447', productContains: 'zubrowka bison grass 0,5', value: 1.447 }
  ];
  const context = buildRuntimeContext_();
  const results = cases.map(function(testCase) {
    const parsed = parseInventoryText(testCase.input, context);
    const first = parsed[0] || {};
    const canonical = getCanonicalParserProductForTest_(first.product || '', context);
    const passed = parsed.length === 1 &&
      canonical.indexOf(normalizeText(testCase.productContains)) !== -1 &&
      Math.abs(Number(first.value) - Number(testCase.value)) < 0.000001;
    return {
      input: testCase.input,
      passed: passed,
      actual: (canonical || first.product || 'BRAK') + ' | ' + String(first.value)
    };
  });

  const passed = results.filter(function(r) { return r.passed; }).length;
  SpreadsheetApp.getUi().alert(
    'Nazwy numeryczne i pojemności — testy 2.10.5',
    'PASS: ' + passed + '/' + results.length + '\n\n' +
      results.map(function(result) {
        return (result.passed ? '✓ ' : '✗ ') + result.input + ' → ' + result.actual;
      }).join('\n'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  if (passed !== results.length) throw new Error('Nie wszystkie testy nazw numerycznych 2.10.5 przeszły.');
  return results;
}


/** Zachowuje zgodność ze starszą pozycją menu 2.10.3. */
function runNumericProductNameRegressionTests2103() {
  return runNumericProductNameRegressionTests2104();
}
/** Inventory PRO RC3.4 — dodatkowe testy longest-match. */

function runLongestMatchParserTestsRC34_() {
  const tests = [
    testLongestMatchOscoRC34_,
    testLongestMatchDoorlysRC34_,
    testLongestMatchContinuousRC34_,
    testLongestMatchNewlineRC34_,
    testLongestMatchZeroRC34_
  ];
  const results = tests.map(function(test) {
    try { test(); return { name: test.name, passed: true, message: 'PASS' }; }
    catch (error) { return { name: test.name, passed: false, message: String(error && error.message || error) }; }
  });
  const passed = results.filter(function(r) { return r.passed; }).length;
  SpreadsheetApp.getUi().alert(
    'Parser Longest Match — RC3.4',
    'PASS: ' + passed + '/' + results.length + '\n\n' +
      results.map(function(r) { return (r.passed ? '✓ ' : '✗ ') + r.name + ': ' + r.message; }).join('\n'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  return results;
}

function testLongestMatchOscoRC34_() {
  const context = createParserTestContext_(['Osco', 'Osco 2 years old']);
  const result = parseInventoryText('Osco 2 years old 22', context);
  ip34Assert_(result.length === 1, 'Oczekiwano 1 pozycji, są: ' + result.length);
  ip34Assert_(normalizeText(result[0].product) === normalizeText('Osco 2 years old'), 'Wybrano zły produkt: ' + result[0].product);
  ip34Assert_(Number(result[0].value) === 22, 'Zła wartość: ' + result[0].value);
}

function testLongestMatchDoorlysRC34_() {
  const context = createParserTestContext_(["Doorly's 12 YO", "Doorly's 3 YO"]);
  const result = parseInventoryText("Doorly's 12 YO 0,987", context);
  ip34Assert_(result.length === 1, 'Oczekiwano 1 pozycji');
  ip34Assert_(normalizeText(result[0].product) === normalizeText("Doorly's 12 YO"), 'Wybrano zły produkt');
  ip34Assert_(Number(result[0].value) === 0.987, 'Zła wartość');
}

function testLongestMatchContinuousRC34_() {
  const context = createParserTestContext_(['Bacardi 8', 'Bacardi 10', 'Osco 2 years old']);
  const result = parseInventoryText('Bacardi 8 0,987 Bacardi 10 1,123 Osco 2 years old 22', context);
  ip34Assert_(result.length === 3, 'Oczekiwano 3 pozycji, są: ' + result.length);
  ip34Assert_(Number(result[0].value) === 0.987, 'Bacardi 8: zła wartość');
  ip34Assert_(Number(result[1].value) === 1.123, 'Bacardi 10: zła wartość');
  ip34Assert_(Number(result[2].value) === 22, 'Osco: zła wartość');
}

function testLongestMatchNewlineRC34_() {
  const context = createParserTestContext_(['Bacardi 8', 'Osco 2 years old']);
  const result = parseInventoryText('Bacardi 8 0,987\nOsco 2 years old 22', context);
  ip34Assert_(result.length === 2, 'Enter powinien dać 2 pozycje');
}

function testLongestMatchZeroRC34_() {
  const context = createParserTestContext_(['Amaro Lucano 1L', 'Amaro Lucano 0% 0,7L']);
  context.catalog[1].aliases = ['amaro lucano zero', 'lucano zero'];
  delete context.__ip34Prepared;
  const result = parseInventoryText('lucano zero 1,234', context);
  ip34Assert_(result.length === 1, 'Oczekiwano 1 pozycji');
  ip34Assert_(normalizeText(result[0].product) === normalizeText('Amaro Lucano 0% 0,7L'), 'Wariant zero rozpoznany błędnie');
  ip34Assert_(Number(result[0].value) === 1.234, 'Zła wartość wariantu zero');
}

function ip34Assert_(condition, message) {
  if (!condition) throw new Error(message || 'Assertion failed');
}
/** Inventory PRO 4.3.4 — kontrakt układu i bezpieczeństwa PAWILONÓW. */
function testPawilonyLayoutContract432_() {
  Object.keys(CONFIG.INVENTORY_LAYOUT).forEach(type => {
    const expected = CONFIG.INVENTORY_LAYOUT[type];
    const actual = getInventorySummaryLayout_(type);
    ['grossWeight', 'emptyContainerWeight', 'openNet', 'prepNet', 'fullUnits',
      'unitCapacity', 'fullUnitsVolume', 'warehouse', 'darkroom', 'fridges',
      'finalTotal'].forEach(key => {
      if (Object.prototype.hasOwnProperty.call(expected, key)) {
        assertCondition_(actual[key] === expected[key],
          type + ': rozjazd konfiguracji pola ' + key + '.');
      }
    });
    assertCondition_(
      getFormulaColumnsForProductType_(type).join('|') === (expected.formulaColumns || []).join('|'),
      type + ': nieprawidłowa lista chronionych formuł.'
    );
  });
}

function testPawilonyLocationMapping432_() {
  const product = {
    type: CONFIG.PRODUCT_TYPES.LOCATION,
    category: 'SOFTY',
    columns: getInputColumnsForProductType_(CONFIG.PRODUCT_TYPES.LOCATION)
  };
  CONFIG.LOCATION_AREAS.forEach(area => {
    const expected = normalizeColumnLetter_(product.columns[area.columnKey]);
    assertCondition_(resolveTargetColumn_(product, 12, area.aliases[0]) === expected,
      area.label + ' musi zapisywać się do ' + expected + '.');
  });
  assertCondition_(resolveTargetColumn_(product, 12, '__nieznana_lokalizacja__') === '',
    'Nieznana lokalizacja nie może trafić do przypadkowej kolumny.');
}

function testGenericBeerClassification432_() {
  const bottleRow = ['', '', '', '', '', '', '', '', '0,5', '', ''];
  const kegRow = ['', '', '', '', '', '', '', '30', '', '', ''];
  assertCondition_(inferInventoryProductType_('NORMAL', 'PIWO BUTELKI', 'Birra 0,5', bottleRow) === 'LOCATION',
    'Sekcja PIWO BUTELKI musi być liczona lokalizacyjnie.');
  assertCondition_(inferInventoryProductType_('KEG', 'PIWO KEG', 'Wawerskie Jasne', kegRow) === 'KEG',
    'Sekcja PIWO KEG musi zachować typ KEG.');
  assertCondition_(inferInventoryProductType_('NORMAL', 'WINO', 'IL RACOLTO KEG 18L', bottleRow) === 'NORMAL',
    'Słowo KEG w nazwie produktu spoza sekcji PIWO nie może zmienić układu na KEG.');
}

function testPawilonyHeaderDetection432_() {
  const normalHeader = [];
  normalHeader[inventoryColumnLetterToNumber_(CONFIG.INVENTORY_LAYOUT.NORMAL.grossWeight) - 1] = 'WAGA/SZT W BUTELCE / KEGU';
  normalHeader[inventoryColumnLetterToNumber_(CONFIG.INVENTORY_LAYOUT.NORMAL.fullUnits) - 1] = 'PEŁNE BTLK szt';
  const normal = detectInventoryInputColumnsFromHeaderRow_(
    normalHeader, CONFIG.PRODUCT_TYPES.NORMAL,
    getInputColumnsForProductType_(CONFIG.PRODUCT_TYPES.NORMAL)
  );
  assertCondition_(normal.weight === CONFIG.INVENTORY_LAYOUT.NORMAL.grossWeight &&
      normal.quantity === CONFIG.INVENTORY_LAYOUT.NORMAL.fullUnits,
    'Nagłówek główny powinien wykryć skonfigurowane wejścia NORMAL.');
  assertCondition_(isExactInventoryHeaderText_('GIN') && !isExactInventoryHeaderText_('Gin Mare 0,7L'),
    'Walidator nie może mylić produktu zawierającego kategorię z nagłówkiem.');
}

function testPawilonyMappingGuard432_() {
  const normal = getInputColumnsForProductType_('NORMAL');
  const keg = getInputColumnsForProductType_('KEG');
  const location = getInputColumnsForProductType_('LOCATION');
  assertCondition_(validateProductColumnMapping_('NORMAL', normal).valid,
    'Poprawne mapowanie NORMAL zostało odrzucone.');
  assertCondition_(validateProductColumnMapping_('KEG', keg).valid,
    'Poprawne mapowanie KEG zostało odrzucone.');
  assertCondition_(validateProductColumnMapping_('LOCATION', location).valid,
    'Poprawne mapowanie LOCATION zostało odrzucone.');
  const badNormal = cloneProductColumns_(normal);
  badNormal.weight = CONFIG.INVENTORY_LAYOUT.NORMAL.formulaColumns[0];
  assertCondition_(!validateProductColumnMapping_('NORMAL', badNormal).valid,
    'Kolumna formuły nie może być wejściem NORMAL.');
}

function testPawilonyTargetColumnRouting432_() {
  const normalColumns = getInputColumnsForProductType_('NORMAL');
  const kegColumns = getInputColumnsForProductType_('KEG');
  const normal = {
    type:'NORMAL', category:'BITTER', columns:normalColumns
  };
  const keg = {
    type:'KEG', category:'PIWO KEG', columns:kegColumns
  };
  const location = {
    type:'LOCATION', category:'SOFTY', columns:getInputColumnsForProductType_('LOCATION')
  };

  assertCondition_(resolveTargetColumn_(normal, 1.25, '') === normalColumns.weight,
    'Wartość dziesiętna NORMAL musi trafić do skonfigurowanej kolumny wagi.');
  assertCondition_(resolveTargetColumn_(normal, 3, '') === normalColumns.quantity,
    'Liczba całkowita NORMAL musi trafić do skonfigurowanej kolumny sztuk.');
  assertCondition_(resolveTargetColumn_(keg, 22.4, '') === kegColumns.weight,
    'Waga KEG musi trafić do skonfigurowanej kolumny wagi.');
  assertCondition_(resolveTargetColumn_(keg, 2, '') === kegColumns.quantity,
    'Liczba pełnych KEG musi trafić do skonfigurowanej kolumny sztuk.');
  CONFIG.LOCATION_AREAS.forEach(area => {
    assertCondition_(resolveTargetColumn_(location, 4, area.aliases[0]) === location.columns[area.columnKey],
      area.label + ': błędne trasowanie lokalizacji.');
  });
}

function testNewProductRowRollback432_() {
  const state = { name:'Nowy produkt', deleted:[] };
  const sheet = {
    getRange: function(row, column) {
      return {
        getDisplayValue: function() { return column === 1 ? state.name : ''; }
      };
    },
    deleteRow: function(row) { state.deleted.push(row); }
  };

  const removed = rollbackInsertedInventoryProductRow_(sheet, 20, 'Nowy produkt');
  assertCondition_(removed.removed && state.deleted.length === 1 && state.deleted[0] === 20,
    'Rollback powinien usunąć wyłącznie własny, niezmieniony wiersz produktu.');

  state.name = 'Ręcznie zmieniony produkt';
  const conflict = rollbackInsertedInventoryProductRow_(sheet, 21, 'Nowy produkt');
  assertCondition_(!conflict.removed && state.deleted.length === 1,
    'Rollback nie może usunąć wiersza zmienionego ręcznie.');
}

function testSparseWritePlan432_() {
  const plan = buildSparseWritePlan_([
    { saved:true, row:10, column:'C', previousValue:'', newValue:2, product:'A', productType:'NORMAL' },
    { saved:true, row:10, column:'C', previousValue:2, newValue:5, product:'A', productType:'NORMAL' },
    { saved:true, row:11, column:'H', previousValue:1, newValue:4, product:'B', productType:'NORMAL' },
    { saved:false, row:null, column:'', previousValue:'', newValue:'' }
  ]);
  assertCondition_(plan.length === 2, 'Plan powinien zawierać dwie zmienione komórki.');
  const c10 = plan.filter(item => item.a1 === 'C10')[0];
  assertCondition_(c10 && c10.previousValue === '' && c10.newValue === 5 && c10.productType === 'NORMAL',
    'Duplikaty C10 muszą zachować pierwszy stan, ostatnią sumę i typ produktu.');
}

function testSparseRollback432_() {
  const values = { C10:5, H10:2, G10:'' };
  const sheet = {
    getRange: function(a1) {
      return {
        getValue: function() { return values[a1]; },
        setValue: function(value) { values[a1] = value; },
        clearContent: function() { values[a1] = ''; }
      };
    }
  };
  rollbackSparseWritePlan_(sheet, [
    { a1:'C10', previousValue:1, newValue:5 },
    { a1:'H10', previousValue:2, newValue:6 },
    { a1:'G10', previousValue:'', newValue:3 }
  ]);
  assertCondition_(values.C10 === 1, 'Zapisana komórka C10 musi zostać cofnięta.');
  assertCondition_(values.H10 === 2, 'Komórka jeszcze niezapisana nie może zostać zmieniona przez rollback.');
  assertCondition_(values.G10 === '', 'Pusta komórka jeszcze niezapisana musi pozostać pusta.');
}

function testFormulaWriteGuard432_() {
  const protectedColumn = getFormulaColumnsForProductType_('NORMAL')[0];
  const inputColumn = getInputColumnsForProductType_('NORMAL').weight;
  const protectedA1 = protectedColumn + '10';
  const inputA1 = inputColumn + '10';
  const values = {};
  const formulas = {};
  values[protectedA1] = 0;
  values[inputA1] = 0;
  formulas[protectedA1] = '';
  formulas[inputA1] = '=A1';
  const sheet = {
    getRange: function(a1) {
      return {
        getFormula: function() { return formulas[a1] || ''; },
        getValue: function() { return values[a1]; },
        setValue: function(value) { values[a1] = value; }
      };
    }
  };

  let protectedColumnBlocked = false;
  try {
    writeSparseWritePlan_(sheet, [{
      a1:protectedA1, row:10, column:protectedColumn, previousValue:0, newValue:1,
      product:'A', productType:'NORMAL'
    }]);
  } catch (error) {
    protectedColumnBlocked = String(error && error.message || error).indexOf('obliczeniowej') >= 0;
  }
  assertCondition_(protectedColumnBlocked,
    'Zapis do kolumny obliczeniowej musi zostać zablokowany nawet po spłaszczeniu formuły.');

  let liveFormulaBlocked = false;
  try {
    writeSparseWritePlan_(sheet, [{
      a1:inputA1, row:10, column:inputColumn, previousValue:0, newValue:1,
      product:'A', productType:'NORMAL'
    }]);
  } catch (error) {
    liveFormulaBlocked = String(error && error.message || error).indexOf('formułą') >= 0;
  }
  assertCondition_(liveFormulaBlocked, 'Istniejąca formuła w dozwolonej kolumnie wejściowej musi blokować zapis.');
}

function testCanonicalInventoryFormulas434_() {
  const known = {mode:CONFIG.PACKAGING_MODES.TARE_KNOWN, emptyWeight:0.5};
  const normal = { inventoryRow:3, type:'NORMAL', category:'BITTER', name:'Amaro', packaging:known };
  const keg = { inventoryRow:85, type:'KEG', category:'PIWO KEG', name:'Wawerskie', packaging:known };
  const location = { inventoryRow:89, type:'LOCATION', category:'PIWO BUTELKI', name:'Butelka' };

  [normal, keg, location].forEach(product => {
    const columns = getFormulaColumnsForProductType_(product.type);
    assertCondition_(columns.length > 0, product.type + ' nie ma kontraktu formuł.');
    columns.forEach(column => {
      const formula = getCanonicalInventoryFormula_(product, column);
      assertCondition_(String(formula || '').charAt(0) === '=' &&
          formula.indexOf(String(product.inventoryRow)) >= 0,
        'Nieprawidłowa formuła ' + column + ' dla ' + product.type + '.');
    });
  });
  const direct = { inventoryRow:90, type:'NORMAL', name:getDirectFinalProductDefinitions_()[0].name };
  assertCondition_(getInventoryFormulaContract_(direct).length === 0,
    'Wyjątek finalny nie może otrzymać formuł.');
}

function testProductPackagingProfiles470_() {
  assertCondition_(
    shouldSeedPackagingCurrentMeasurement_({workflow:'IMPORT', seedCurrentMeasurement:true}) === false,
    'Utworzenie produktu podczas importu nie może wstępnie wpisywać pomiaru referencyjnego.'
  );
  assertCondition_(
    shouldSeedPackagingCurrentMeasurement_({workflow:'IMPORT'}) === false,
    'Domyślna ścieżka importu musi pozostawić bieżącą komórkę pomiaru pustą.'
  );
  const known = validateProductPackagingProfile_({
    mode:CONFIG.PACKAGING_MODES.TARE_KNOWN, emptyWeight:'0,489'
  }, {type:CONFIG.PRODUCT_TYPES.NORMAL, name:'Fernet'});
  assertCondition_(known.emptyWeight === 0.489, 'Profil tary musi obsługiwać polski separator dziesiętny.');

  const unknown = validateProductPackagingProfile_({
    mode:CONFIG.PACKAGING_MODES.TARE_UNKNOWN
  }, {type:CONFIG.PRODUCT_TYPES.NORMAL, name:'Nowy produkt'});
  assertCondition_(unknown.mode === CONFIG.PACKAGING_MODES.TARE_UNKNOWN,
    'Produkt musi dać się utworzyć bez dostępnej tary.');

  const reference = validateProductPackagingProfile_({
    mode:CONFIG.PACKAGING_MODES.GROSS_REFERENCE,
    referenceGrossWeight:5,
    referenceNetQuantity:2,
    massConversionFactor:1,
    referenceDate:'2026-07-26'
  }, {type:CONFIG.PRODUCT_TYPES.NORMAL, name:'Butelka referencyjna'});
  assertCondition_(Math.abs(calculateReferenceRemainingQuantity_(4.6, reference) - 1.6) < 0.000001,
    'Tryb referencyjny musi odejmować ubytek masy od ilości początkowej.');
  const localeSafeFormula = buildReferenceRemainingFormula_('C10', {
    referenceGrossWeight:1.2,
    referenceNetQuantity:0.7,
    massConversionFactor:1
  });
  assertCondition_(
    localeSafeFormula.indexOf('1.2') < 0 &&
    localeSafeFormula.indexOf('0.7') < 0 &&
    localeSafeFormula.indexOf('(12/10)') >= 0 &&
    localeSafeFormula.indexOf('(7/10)') >= 0,
    'Formuła referencyjna nie może zawierać regionalnych literałów dziesiętnych.'
  );

  const unknownItem = {
    product:'Bez tary', category:'TEST', type:CONFIG.PRODUCT_TYPES.NORMAL,
    packaging:unknown, details:{grossWeight:1.2, emptyContainerWeight:'', prepNet:'', fullUnits:2, unitCapacity:0.7}
  };
  applyPackagingAwareReportingValues_(unknownItem);
  unknownItem.reviewReason = classifyFinalReviewRequirement_(unknownItem);
  assertCondition_(Math.abs(unknownItem.finalTotal - 1.4) < 0.000001 &&
    unknownItem.reviewReason === 'PACKAGING_DATA_MISSING',
    'Brak tary nie może tworzyć sztucznego wyniku z wagi brutto i musi wymagać weryfikacji.');

  const referenceItem = {
    product:'Referencja', category:'TEST', type:CONFIG.PRODUCT_TYPES.NORMAL,
    packaging:reference, details:{grossWeight:4.6, emptyContainerWeight:'', prepNet:'', fullUnits:'', unitCapacity:0.7}
  };
  applyPackagingAwareReportingValues_(referenceItem);
  assertCondition_(Math.abs(referenceItem.finalTotal - 1.6) < 0.000001,
    'Raport musi użyć obliczenia ubytku w trybie wagi referencyjnej.');
}

function testFormulaContractFailClosed511_() {
  let blocked = false;
  try {
    validateInventoryFormulaContract_(
      {name:'Kontrakt testowy', type:CONFIG.PRODUCT_TYPES.NORMAL},
      {
        openNet:'E',
        fullUnitsVolume:'J',
        finalTotal:'K',
        formulaColumns:['F', 'J', 'K']
      },
      [
        {column:'E', columnNumber:5, formula:'=1', r1c1:'=1'},
        {column:'J', columnNumber:10, formula:'=1', r1c1:'=1'},
        {column:'K', columnNumber:11, formula:'=1', r1c1:'=1'}
      ]
    );
  } catch (error) {
    blocked = true;
  }
  assertCondition_(blocked,
    'Niezgodny układ formuł musi zostać zatrzymany przed pierwszym zapisem.');

  let repairBlocked = false;
  const kegFormulaColumns = getFormulaColumnsForProductType_(CONFIG.PRODUCT_TYPES.KEG);
  const forbiddenColumn = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('').find(function(column) {
    return kegFormulaColumns.indexOf(column) === -1;
  });
  try {
    validateInventoryFormulaRepairPlanTargets_([{
      type: CONFIG.PRODUCT_TYPES.KEG,
      a1: forbiddenColumn + '215',
      column: forbiddenColumn,
      columnNumber: inventoryColumnLetterToNumber_(forbiddenColumn),
      classification: 'INVALID_FORMULA'
    }]);
  } catch (error) {
    repairBlocked = true;
  }
  assertCondition_(repairBlocked,
    'Plan naprawy nie może zapisywać formuły poza kolumnami formuł danego typu.');
}

function testMissingPackagingReview511_() {
  const missing = buildPackagingReviewForManager_({
    active:true,
    type:CONFIG.PRODUCT_TYPES.NORMAL,
    packaging:{mode:CONFIG.PACKAGING_MODES.TARE_UNKNOWN}
  });
  assertCondition_(missing.needsDecision && missing.fullUnitsStillCalculated,
    'Brak tary musi trafić do kolejki decyzji bez blokowania pełnych sztuk.');

  const known = buildPackagingReviewForManager_({
    active:true,
    type:CONFIG.PRODUCT_TYPES.NORMAL,
    packaging:{mode:CONFIG.PACKAGING_MODES.TARE_KNOWN, emptyWeight:'0,489'}
  });
  assertCondition_(!known.needsDecision,
    'Produkt z dodatnią tarą nie może trafić do kolejki braków opakowania.');
}

function testProductCreationAndCapacityGuard484_() {
  assertCondition_(
    resolveNewProductUnitCapacity_('', '', '0,7') === 0.7,
    'Nowy produkt musi dziedziczyć pojemność z produktu referencyjnego.'
  );
  assertCondition_(
    resolveNewProductUnitCapacity_('1', '0,7', '0,5') === 1,
    'Jawnie podana pojemność musi mieć pierwszeństwo przed produktem referencyjnym.'
  );
  assertCondition_(
    resolveNewProductUnitCapacity_('', '0,7', '1') === 0.7,
    'Ilość początkowa profilu brutto musi mieć pierwszeństwo przed pojemnością produktu referencyjnego.'
  );
  const missingCapacity = {
    grossWeight: '',
    prepNet: '',
    fullUnits: 2,
    unitCapacity: ''
  };
  assertCondition_(
    hasMissingUnitCapacityForFullUnits_(missingCapacity),
    'Pełne sztuki bez pojemności muszą blokować raport.'
  );
  const item = {
    product: 'Nowy produkt',
    type: CONFIG.PRODUCT_TYPES.NORMAL,
    details: missingCapacity
  };
  assertCondition_(
    classifyFinalReviewRequirement_(item) === 'UNIT_CAPACITY_MISSING' &&
    !isReportingItemCompleted_(item),
    'Brak pojemności musi być widoczny jako decyzja i pozycja nie może być kompletna.'
  );
  assertCondition_(
    !hasMissingUnitCapacityForFullUnits_({fullUnits: 0, unitCapacity: ''}),
    'Zerowa liczba pełnych sztuk nie może wymagać pojemności.'
  );
}

function testProductManagerFallbackRow485_() {
  assertCondition_(
    resolveProductManagerFallbackRow_({'fernet branca': 24}, 'fernet branca') === 24,
    'Product Manager musi zachować wiersz INWENTURY w ścieżce awaryjnej.'
  );
  assertCondition_(
    resolveProductManagerFallbackRow_(null, 'brak') === null,
    'Brak mapy wierszy nie może powodować ReferenceError.'
  );
}

function testTareSafeCategoryReview462_() {
  const onlyUnits = {
    key:'fernet', product:'Fernet Branca 0,7L', category:'BITTER',
    type:CONFIG.PRODUCT_TYPES.NORMAL, finalTotal:-0.489, total:-0.489, unit:'l',
    details:{grossWeight:'',emptyContainerWeight:0.489,openNet:-0.489,prepNet:'',fullUnits:12,unitCapacity:0.7,fullUnitsVolume:8.4}
  };
  applyTareSafeReportingValues_(onlyUnits);
  onlyUnits.reviewReason = classifyFinalReviewRequirement_(onlyUnits);
  assertCondition_(onlyUnits.details.openNet === 0 && Math.abs(onlyUnits.finalTotal - 8.4) < 0.000001,
    'Brak wagi brutto nie może odejmować tary od pełnych sztuk.');
  assertCondition_(onlyUnits.reviewReason === 'FULL_UNITS_ONLY',
    'Produkt z samymi pełnymi sztukami musi wymagać weryfikacji kategorii.');

  const missing = {
    key:'brak', product:'Brak produktu', category:'BITTER',
    type:CONFIG.PRODUCT_TYPES.NORMAL, finalTotal:-0.5, total:-0.5, unit:'l',
    details:{grossWeight:'',emptyContainerWeight:0.5,openNet:-0.5,prepNet:'',fullUnits:'',unitCapacity:0.7,fullUnitsVolume:''}
  };
  applyTareSafeReportingValues_(missing);
  missing.reviewReason = classifyFinalReviewRequirement_(missing);
  const groups = buildFinalReviewCategoryGroups_([onlyUnits, missing]);
  assertCondition_(missing.finalTotal === 0 && missing.reviewReason === 'MISSING_PRODUCT',
    'Brak produktu musi dawać stan 0 i wymagać potwierdzenia.');
  assertCondition_(groups.length === 1 && groups[0].total === 2 &&
    groups[0].missingProducts === 1 && groups[0].fullUnitsOnly === 1,
    'Pozycje muszą być grupowane i zatwierdzane według kategorii.');
}


function testDirectFinalCoffeeException434_() {
  const definitions = getDirectFinalProductDefinitions_();
  assertCondition_(definitions.length > 0, 'Lokal nie ma skonfigurowanego wyjątku finalnego.');
  definitions.forEach(definition => {
    const product = {
      inventoryRow: 275,
      type: 'NORMAL',
      category: 'DIRECT_FINAL',
      name: definition.name,
      columns: { weight: 'C', quantity: 'H' }
    };
    const directColumn = normalizeColumnLetter_(definition.column);
    assertCondition_(isDirectFinalInventoryProduct_(product), 'Nie rozpoznano wyjątku ' + definition.name + '.');
    assertCondition_(resolveTargetColumn_(product, 3, '') === directColumn,
      definition.name + ' musi zapisywać wartość do ' + directColumn + '.');
    assertCondition_(getInventoryFormulaContract_(product).length === 0,
      definition.name + ' nie może mieć formuł.');
    assertCondition_(assertSafeInventoryTargetColumn_(product, directColumn) === directColumn,
      'Kolumna ' + directColumn + ' musi być dozwolona dla ' + definition.name + '.');
    let blocked = false;
    try { assertSafeInventoryTargetColumn_(product, 'K'); } catch (error) { blocked = true; }
    assertCondition_(blocked, 'Zapis ' + definition.name + ' poza kolumną finalną musi być blokowany.');
  });
}

function testDirectFinalProductManagerMapping4313_() {
  const definition = getDirectFinalProductDefinitions_()[0];
  const coffee = { name:definition.name, type:'NORMAL' };
  const valid = validateProductColumnMapping_('NORMAL', {
    quantity:definition.column, weight:'', warehouse:'', darkroom:'', fridges:''
  }, coffee);
  assertCondition_(valid.valid, 'Product Manager musi akceptować kolumnę finalną wyjątku.');
  const invalid = validateProductColumnMapping_('NORMAL', {
    quantity:'H', weight:'C', warehouse:'', darkroom:'', fridges:''
  }, coffee);
  assertCondition_(!invalid.valid, 'Product Manager musi odrzucać zwykłe mapowanie wyjątku finalnego.');
}

function testDirectFinalSparseWriteGuard461_() {
  const definition = getDirectFinalProductDefinitions_()[0];
  const safeColumn = validateSparseWriteTarget_({
    product: definition.name,
    productType: 'NORMAL',
    column: definition.column
  });
  assertCondition_(
    safeColumn === normalizeColumnLetter_(definition.column),
    'Końcowa walidacja planu zapisu musi dopuścić skonfigurowaną kolumnę finalną.'
  );

  let blocked = false;
  const forbiddenColumn = CONFIG.INVENTORY_LAYOUT.NORMAL.formulaColumns[0];
  try {
    validateSparseWriteTarget_({
      product: 'Zwykły produkt',
      productType: 'NORMAL',
      column: forbiddenColumn
    });
  } catch (error) {
    blocked = true;
  }
  assertCondition_(
    blocked,
    'Końcowa walidacja musi blokować kolumnę formuły dla zwykłego produktu NORMAL.'
  );
}

function testXlsxExportWithoutDriveApi4313_() {
  const url = buildSpreadsheetXlsxExportUrl_('test-sheet-id');
  assertCondition_(url === 'https://docs.google.com/spreadsheets/d/test-sheet-id/export?format=xlsx',
    'Nieprawidłowy endpoint XLSX.');
  assertCondition_(url.indexOf('googleapis.com/drive/v3') === -1,
    'Eksport XLSX nie może wymagać Drive API v3.');
}

function testFormulaRepairSegments432_() {
  const plan = [
    { row:3, columnNumber:5, r1c1:'=RC[-2]-RC[-1]' },
    { row:4, columnNumber:5, r1c1:'=RC[-2]-RC[-1]' },
    { row:6, columnNumber:5, r1c1:'=RC[-2]-RC[-1]' },
    { row:3, columnNumber:10, r1c1:'=RC[-2]*RC[-1]' }
  ];
  const segments = buildFormulaRepairSegments_(plan);
  assertCondition_(segments.length === 3,
    'Plan powinien zostać podzielony na 3 bloki, otrzymano: ' + segments.length);
  assertCondition_(segments.some(segment =>
    segment.columnNumber === 5 && segment.startRow === 3 && segment.endRow === 4
  ), 'Brakuje wspólnego bloku E3:E4.');
}

function testFormulaRepairConcurrency432_() {
  const cells = {
    E3:{ formula:'', value:1.5 },
    J3:{ formula:'=H3*I3', value:1.4 }
  };
  const sheet = {
    getRange: function(row, column) {
      const a1 = (column === 5 ? 'E' : 'J') + row;
      return {
        getFormula: function() { return cells[a1].formula; },
        getValue: function() { return cells[a1].value; },
        setFormula: function(value) { cells[a1].formula = value; },
        setValue: function(value) { cells[a1].formula = ''; cells[a1].value = value; },
        clearContent: function() { cells[a1].formula = ''; cells[a1].value = ''; }
      };
    }
  };
  const plan = [{
    row:3, column:'E', columnNumber:5, a1:'E3',
    formula:'=C3-D3', previousFormula:'', previousValue:1.5
  }];

  preflightInventoryFormulaRepairPlan_(sheet, plan);
  cells.E3.value = 9;
  let blocked = false;
  try { preflightInventoryFormulaRepairPlan_(sheet, plan); }
  catch (error) { blocked = String(error && error.message || error).indexOf('zmieniona po audycie') >= 0; }
  assertCondition_(blocked,
    'Naprawa formuł musi zostać przerwana po ręcznej zmianie komórki od czasu audytu.');

  cells.E3.value = 1.5;
  cells.E3.formula = '=C3-D3';
  rollbackInventoryFormulaRepairPlan_(sheet, plan);
  assertCondition_(cells.E3.formula === '' && cells.E3.value === 1.5,
    'Rollback powinien przywrócić stan komórki zapisanej przez naprawę.');

  cells.E3.formula = '=C3-D3+1';
  cells.E3.value = 2.5;
  rollbackInventoryFormulaRepairPlan_(sheet, plan);
  assertCondition_(cells.E3.formula === '=C3-D3+1' && cells.E3.value === 2.5,
    'Rollback nie może nadpisać formuły zmienionej ręcznie po naprawie.');
}

function buildFormulaAuditFakeSheet432_(rowValues, rowFormulas, rowDisplayValues) {
  const width = getInventoryLayoutMaxColumn_();
  const empty = new Array(width).fill('');
  const values = [empty.slice(), empty.slice()];
  const formulas = [empty.slice(), empty.slice()];
  const displays = [empty.slice(), empty.slice()];
  values[1] = rowValues.slice();
  formulas[1] = rowFormulas.slice();
  displays[1] = rowDisplayValues.slice();
  return {
    getLastRow: function() { return 2; },
    getName: function() { return 'INWENTURA'; },
    getRange: function() {
      return {
        getValues: function() { return values; },
        getFormulas: function() { return formulas; },
        getDisplayValues: function() { return displays; }
      };
    }
  };
}

function testFormulaConflictClassification432_() {
  const layout = CONFIG.INVENTORY_LAYOUT.NORMAL;
  const width = getInventoryLayoutMaxColumn_();
  const row = new Array(width).fill('');
  row[inventoryColumnLetterToNumber_(layout.grossWeight) - 1] = 2;
  row[inventoryColumnLetterToNumber_(layout.emptyContainerWeight) - 1] = 0.5;
  row[inventoryColumnLetterToNumber_(layout.openNet) - 1] = 999;
  if (layout.prepNet) row[inventoryColumnLetterToNumber_(layout.prepNet) - 1] = 0;
  row[inventoryColumnLetterToNumber_(layout.fullUnits) - 1] = 2;
  row[inventoryColumnLetterToNumber_(layout.unitCapacity) - 1] = 0.7;
  const blankFormulas = new Array(width).fill('');
  const display = row.map(value => value === '' ? '' : String(value));
  const product = { inventoryRow:2, type:'NORMAL', category:'BITTER', name:'Test' };
  const conflictAudit = buildInventoryFormulaAudit_(
    buildFormulaAuditFakeSheet432_(row, blankFormulas, display), [product]
  );
  assertCondition_(conflictAudit.conflictFormulaCells === 1,
    'Wartość sprzeczna z kontraktem musi zostać konfliktem.');
  assertCondition_(conflictAudit.missingFormulaCells === layout.formulaColumns.length - 1,
    'Pozostałe puste pola formuł powinny zostać sklasyfikowane jako brakujące.');
  assertCondition_(conflictAudit.hasBlockingConflicts,
    'Konflikt musi blokować automatyczną naprawę.');

  const flattenedRow = row.slice();
  flattenedRow[inventoryColumnLetterToNumber_(layout.openNet) - 1] = 1.5;
  const flattenedDisplay = flattenedRow.map(value => value === '' ? '' : String(value));
  const flattenedAudit = buildInventoryFormulaAudit_(
    buildFormulaAuditFakeSheet432_(flattenedRow, blankFormulas, flattenedDisplay), [product]
  );
  assertCondition_(flattenedAudit.flattenedFormulaCells === 1,
    'Zgodny wynik bez formuły powinien zostać oznaczony jako spłaszczony.');

}

function testUndoConflictEvaluation432_() {
  const live = { C10:5, H11:9, E12:3 };
  const formulas = { E12:'=C12-D12' };
  const sheet = {
    getRange: function(a1) {
      return {
        getValue: function() { return live[a1]; },
        getFormula: function() { return formulas[a1] || ''; }
      };
    }
  };
  const plan = {
    changes: [
      { row:10, column:'C', previousValue:1, expectedNewValue:5, auditRows:[2] },
      { row:11, column:'H', previousValue:2, expectedNewValue:6, auditRows:[3] },
      { row:12, column:'E', previousValue:0, expectedNewValue:3, auditRows:[4] }
    ]
  };
  const evaluated = evaluateUndoPlanAgainstInventory_(sheet, plan, { E12:{ formula:'=C12-D12' } });
  assertCondition_(evaluated.applicableChanges.length === 1 && evaluated.applicableChanges[0].column === 'C',
    'Tylko niezmieniona komórka C10 powinna zostać cofnięta.');
  assertCondition_(evaluated.conflictChanges.length === 2,
    'Ręczna zmiana i komórka formuły powinny zostać konfliktami.');
  assertCondition_(evaluated.conflictChanges.some(change => change.reason === 'VALUE_CHANGED') &&
      evaluated.conflictChanges.some(change => change.reason === 'FORMULA_PROTECTED'),
    'Cofanie powinno rozróżniać zmianę ręczną i ochronę formuły.');
}
