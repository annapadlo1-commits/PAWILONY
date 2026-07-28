/**
 * Inventory PRO 4.1 — test kontraktu wydajności matchera.
 * Nie zależy od arkusza ani produkcyjnego katalogu.
 */
function runParserPerformanceTests() {
  const names = [];
  for (let index = 0; index < 350; index++) {
    names.push('Produkt Marka' + index + ' Wariant ' + (index % 24));
  }
  const context = createParserTestContext_(names);
  context.trigramRecognitionIndex = buildTrigramRecognitionIndex_(context.catalog);
  context.recognitionSources = buildRecognitionSources_(context.catalog);
  context.recognitionSourceTrigrams = buildRecognitionSourceTrigrams_(context.catalog);
  context.performanceStats = {
    memoHits: 0,
    fuzzyQueries: 0,
    fuzzyShortlisted: 0,
    levenshteinCandidates: 0
  };

  const startedAt = Date.now();
  const shortlist = getRecognitionShortlist_('Produkt Mraka127 Wariant 7', context);
  const match = recognizeProduct_('Produkt Mraka127 Wariant 7', context);
  const durationMs = Date.now() - startedAt;

  assertCondition_(shortlist.length <= CONFIG.PERFORMANCE.FUZZY_SHORTLIST_SIZE,
    'Shortlista fuzzy przekracza limit: ' + shortlist.length);
  assertCondition_(match.matched && match.product && match.product.name === 'Produkt Marka127 Wariant 7',
    'Matcher nie rozpoznal kontrolowanej literowki: ' + JSON.stringify(match));
  assertCondition_(context.performanceStats.levenshteinCandidates <= CONFIG.PERFORMANCE.LEVENSHTEIN_FINALISTS,
    'Levenshtein otrzymal zbyt wielu kandydatow: ' + context.performanceStats.levenshteinCandidates);

  const report = {
    passed: true,
    durationMs: durationMs,
    shortlistSize: shortlist.length,
    levenshteinCandidates: context.performanceStats.levenshteinCandidates,
    matchedProduct: match.product.name
  };
  SpreadsheetApp.getUi().alert(
    'Inventory PRO — wydajność',
    'PASS\nCzas: ' + durationMs + ' ms\nShortlista: ' + shortlist.length +
      '\nFinaliści Levenshteina: ' + context.performanceStats.levenshteinCandidates,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  return report;
}