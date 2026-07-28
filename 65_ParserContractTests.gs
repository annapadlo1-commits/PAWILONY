/**
 * Inventory PRO 3.0 — kontrakt parsera.
 *
 * Ten zestaw jest niezależny od historycznych testów wersjonowanych.
 * Każda przebudowa parsera musi przejść cały kontrakt przed wdrożeniem.
 */
function runParserContractTests() {
  const startedAt = Date.now();
  const tests = [
    parserContractLucanoZero_,
    parserContractKolaZeroVariant_,
    parserContractOrdinaryZero_,
    parserContractHalfLitre_,
    parserContractLongestAgeName_,
    parserContractContinuousInput_,
    parserContractHardNewlineBoundary_,
    parserContractLeadingValue_,
    parserContractValueAfterProduct_,
    parserContractNoCrossLineLeakage_,
    parserContractNumericVariant_,
    parserContractLocationSection_,
    parserContractQuantityUnitSuffix_
  ];

  const results = tests.map(function(testFunction) {
    const testStartedAt = Date.now();
    try {
      testFunction();
      return {
        name: testFunction.name,
        passed: true,
        durationMs: Date.now() - testStartedAt,
        message: 'PASS'
      };
    } catch (error) {
      return {
        name: testFunction.name,
        passed: false,
        durationMs: Date.now() - testStartedAt,
        message: normalizeError_(error).message
      };
    }
  });

  const passed = results.filter(function(result) { return result.passed; }).length;
  const report = {
    total: results.length,
    passed: passed,
    failed: results.length - passed,
    durationMs: Date.now() - startedAt,
    results: results
  };

  logInfo(
    'ParserContractTests',
    'runParserContractTests',
    report.failed === 0 ? 'Kontrakt parsera PASS' : 'Kontrakt parsera FAIL',
    report,
    report.durationMs
  );

  SpreadsheetApp.getUi().alert(
    'Inventory PRO — kontrakt parsera',
    formatTestReport_(report),
    SpreadsheetApp.getUi().ButtonSet.OK
  );

  return report;
}

function parserContractLucanoZero_() {
  const context = createParserTestContext_([
    'Amaro Lucano 1L',
    'Amaro Lucano 0% 0,7L'
  ]);
  context.catalog[1].aliases = ['amaro lucano zero', 'lucano zero'];
  context.parserPhraseIndex = null;

  assertParserContract_(
    parseInventoryText('lucano zero 1,234', context),
    [{ product: 'Amaro Lucano 0% 0,7L', value: 1.234 }],
    'Lucano zero'
  );
}

function parserContractKolaZeroVariant_() {
  const context = createParserTestContext_([
    'Fritz 200ml KOLA',
    'Fritz 200ml KOLA BEZ CUKRU'
  ]);
  context.catalog[0].aliases = ['fritz kola'];
  context.catalog[1].aliases = ['kola zero', 'kola 0'];
  context.parserPhraseIndex = null;

  assertParserContract_(
    parseInventoryText('kola 0 12', context),
    [{ product: 'Fritz 200ml KOLA BEZ CUKRU', value: 12 }],
    'Kola zero'
  );
}

function parserContractOrdinaryZero_() {
  const context = createParserTestContext_(['Fritz 200ml KOLA']);
  context.catalog[0].aliases = ['fritz kola'];
  context.parserPhraseIndex = null;

  assertParserContract_(
    parseInventoryText('fritz kola 0', context),
    [{ product: 'Fritz 200ml KOLA', value: 0 }],
    'Zwykle zero'
  );
}

function parserContractHalfLitre_() {
  const context = createParserTestContext_(['Żubrówka Bison Grass 0,5L']);
  context.catalog[0].aliases = [
    'żubrówka bison grass pół',
    'żubrówka bison grass pół litra'
  ];
  context.parserPhraseIndex = null;

  assertParserContract_(
    parseInventoryText('żubrówka bison grass pół litra 1,234', context),
    [{ product: 'Żubrówka Bison Grass 0,5L', value: 1.234 }],
    'Pol litra'
  );
}

function parserContractLongestAgeName_() {
  const context = createParserTestContext_(['Osco', 'Osco 2 years old']);
  assertParserContract_(
    parseInventoryText('Osco 2 years old 22', context),
    [{ product: 'Osco 2 years old', value: 22 }],
    'Najdluzsza nazwa wieku'
  );
}

function parserContractContinuousInput_() {
  const context = createParserTestContext_(['Bacardi 8', 'Bacardi 10']);
  assertParserContract_(
    parseInventoryText('Bacardi 8 0,987 Bacardi 10 1,123', context),
    [
      { product: 'Bacardi 8', value: 0.987 },
      { product: 'Bacardi 10', value: 1.123 }
    ],
    'Tekst ciagly'
  );
}

function parserContractHardNewlineBoundary_() {
  const context = createParserTestContext_(['Bacardi 8', 'Osco 2 years old']);
  assertParserContract_(
    parseInventoryText('Bacardi 8 0,987\nOsco 2 years old 22', context),
    [
      { product: 'Bacardi 8', value: 0.987 },
      { product: 'Osco 2 years old', value: 22 }
    ],
    'Granica Enter'
  );
}

function parserContractLeadingValue_() {
  const context = createParserTestContext_(['Ardbeg 10']);
  assertParserContract_(
    parseInventoryText('1,407 Ardbeg 10', context),
    [{ product: 'Ardbeg 10', value: 1.407 }],
    'Wartosc przed produktem'
  );
}

function parserContractValueAfterProduct_() {
  const context = createParserTestContext_(['Campari 0,7 l']);
  assertParserContract_(
    parseInventoryText('Campari 1,239', context),
    [{ product: 'Campari 0,7 l', value: 1.239 }],
    'Wartosc po produkcie'
  );
}

function parserContractNoCrossLineLeakage_() {
  const context = createParserTestContext_(['Campari 0,7 l', 'Bacardi 8']);
  const parsed = parseInventoryText('Campari\nBacardi 8 0,987', context);
  assertCondition_(parsed.length === 2, 'Brak wartosci nie moze polknac nastepnego wiersza.');
  assertCondition_(parsed[0].value === null, 'Campari bez wartosci powinno pozostac bledem.');
  assertCondition_(Number(parsed[1].value) === 0.987, 'Bacardi otrzymalo zla wartosc po Enterze.');
}

function parserContractNumericVariant_() {
  const context = createParserTestContext_(['Tanqueray 10']);
  assertParserContract_(
    parseInventoryText('Tanqueray 10 1,279', context),
    [{ product: 'Tanqueray 10', value: 1.279 }],
    'Numer wariantu produktu'
  );
}

function parserContractLocationSection_() {
  const context = createParserTestContext_(['BOMBILLA 0,3L']);
  assertParserContract_(
    parseInventoryText('magazyn\nBombilla 12', context),
    [{ product: 'BOMBILLA 0,3L', value: 12, location: 'magazyn' }],
    'Sekcja lokalizacji'
  );
}

function parserContractQuantityUnitSuffix_() {
  const context = createParserTestContext_([
    'Fritz 200ml KOLA BEZ CUKRU',
    'BOMBILLA 0,3L'
  ]);
  context.catalog[0].aliases = ['kola zero'];
  context.parserPhraseIndex = null;

  ['sztuk', 'sztuki', 'sztuka', 'szt.'].forEach(function(unit) {
    assertParserContract_(
      parseInventoryText('kola zero 7 ' + unit + ' bombilla 2', context),
      [
        { product: 'Fritz 200ml KOLA BEZ CUKRU', value: 7 },
        { product: 'BOMBILLA 0,3L', value: 2 }
      ],
      'Jednostka po liczbie: ' + unit
    );
  });
}

function assertParserContract_(actual, expected, label) {
  assertCondition_(
    actual.length === expected.length,
    label + ': oczekiwano ' + expected.length + ' pozycji, otrzymano ' + actual.length + '. Wynik: ' + JSON.stringify(actual)
  );

  expected.forEach(function(expectedItem, index) {
    const actualItem = actual[index] || {};
    assertCondition_(
      normalizeText(actualItem.product) === normalizeText(expectedItem.product),
      label + ': zly produkt #' + (index + 1) + ': ' + actualItem.product
    );
    assertCondition_(
      Math.abs(Number(actualItem.value) - Number(expectedItem.value)) < 0.000001,
      label + ': zla wartosc #' + (index + 1) + ': ' + actualItem.value
    );
    if (expectedItem.location) {
      assertCondition_(
        normalizeText(actualItem.location) === normalizeText(expectedItem.location),
        label + ': zla lokalizacja #' + (index + 1) + ': ' + actualItem.location
      );
    }
  });
}
