/**
 * Inventory PRO 5.0 — testy kontraktu aplikacji mobilnej.
 * Testy są tylko do odczytu i nie modyfikują arkusza.
 */
function testGeminiLosslessPrompt500_() {
  const prompt = buildGeminiInventoryPrompt_();
  [
    'niczego nie pomijaj',
    'Nie wymyślaj',
    'wysokiej pewności',
    'zapisz fragment możliwie dosłownie lub fonetycznie',
    'Nie wybieraj między produktami o różnych pojemnościach'
  ].forEach(function(fragment) {
    if (prompt.indexOf(fragment) === -1) {
      throw new Error('Prompt Gemini nie zawiera reguły: ' + fragment);
    }
  });
  (CONFIG.LOCATION_AREAS || []).forEach(function(area) {
    [area.label].concat(area.aliases || []).filter(Boolean).forEach(function(alias) {
      if (prompt.indexOf(alias) === -1) {
        throw new Error('Prompt Gemini nie zawiera lokalizacji lub aliasu: ' + alias);
      }
    });
  });
}

function testMobileResolverAliases500_() {
  const context = buildRuntimeContext_();
  const payload = buildProductResolverPayload_(context, '');
  if (!payload || !Array.isArray(payload.products) || !payload.products.length) {
    throw new Error('Mobilny katalog produktów jest pusty.');
  }
  payload.products.forEach(function(product) {
    if (!Array.isArray(product.aliases)) {
      throw new Error('Mobilny produkt nie udostępnia tablicy aliasów: ' + product.name);
    }
  });
}

function testAudioProcessorTriggerDeduplication513_() {
  const scheduleSource = String(scheduleInventoryAudioProcessor_);
  const processorSource = String(processPendingInventoryAudioJobs_);
  assertCondition_(
    scheduleSource.indexOf('LockService.getScriptLock') >= 0,
    'Tworzenie wyzwalacza procesora audio musi być chronione blokadą.'
  );
  assertCondition_(
    scheduleSource.indexOf('getInventoryAudioProcessorTriggers_') >= 0 &&
      scheduleSource.indexOf('triggers.slice(1)') >= 0,
    'Harmonogram musi wykrywać i usuwać zdublowane wyzwalacze.'
  );
  assertCondition_(
    processorSource.indexOf('removeInventoryAudioProcessorTriggers_') >= 0,
    'Procesor musi sprzątać uruchomiony i osierocone wyzwalacze.'
  );
  assertCondition_(
    processorSource.indexOf('nextProcessorDelay') >= 0,
    'Następny procesor musi być planowany dopiero po zwolnieniu blokady.'
  );
}

function testQuickInventoryRejectsEmptyList500_() {
  let rejected = false;
  try {
    analyzeQuickInventoryItems([]);
  } catch (error) {
    rejected = true;
  }
  if (!rejected) throw new Error('Szybka inwentaryzacja musi odrzucać pustą listę.');
}

function testGeminiTransientRetry500_() {
  ['HTTP 429', 'HTTP 500', 'HTTP 502', 'HTTP 503', 'HTTP 504', 'high demand', 'network timeout']
    .forEach(function(message) {
      if (!isTransientGeminiError_(message)) {
        throw new Error('Nie rozpoznano przejściowego błędu Gemini: ' + message);
      }
    });
  if (isTransientGeminiError_('HTTP 400: nieprawidłowe żądanie')) {
    throw new Error('Błąd trwały HTTP 400 nie może uruchamiać automatycznych prób.');
  }
  if (geminiAudioRetryDelayMs_(2) <= geminiAudioRetryDelayMs_(1)) {
    throw new Error('Odstęp automatycznych prób Gemini nie rośnie.');
  }
}

function testGeminiCompressedAudio510_() {
  const cases = [
    ['audio/mp4', 'nagranie.m4a', 'audio/mp4'],
    ['audio/mpeg', 'nagranie.mp3', 'audio/mpeg'],
    ['', 'nagranie.aac', 'audio/aac'],
    ['audio/wav', 'nagranie.wav', 'audio/wav']
  ];
  cases.forEach(function(item) {
    const actual = normalizeGeminiAudioMimeType_(item[0], item[1]);
    if (actual !== item[2]) {
      throw new Error('Nieprawidłowy MIME dla ' + item[1] + ': ' + actual);
    }
  });
}
