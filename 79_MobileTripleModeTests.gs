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
      scheduleSource.indexOf('ScriptApp.deleteTrigger') >= 0 &&
      scheduleSource.indexOf('ScriptApp.newTrigger') >= 0,
    'Harmonogram musi usuwać stary wyzwalacz i tworzyć rzeczywiście przeplanowany.'
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

function testAudioPipelineRecovery514_() {
  const queueSource = String(queueInventoryAudioJob);
  const workerSource = String(processPendingInventoryAudioJobs_);
  const recoverySource = String(getInventoryAudioRecoveryDecision_);
  assertCondition_(
    queueSource.indexOf('saveInventoryAudioJob_') >= 0 &&
      queueSource.indexOf('scheduleInventoryAudioProcessor_(1000)') >= 0,
    'Zapis nagrania musi utrwalić zadanie i natychmiast uruchomić kolejkę.'
  );
  assertCondition_(
    workerSource.indexOf('transcribeInventoryAudio') >= 0 &&
      workerSource.indexOf("job.status = 'DONE'") >= 0,
    'Ścieżka kolejki musi prowadzić przez Gemini do gotowego transkryptu.'
  );
  assertCondition_(
    workerSource.indexOf('GEMINI_AUDIO_JOB_DEADLINE_MS_') >= 0 &&
      recoverySource.indexOf('GEMINI_AUDIO_JOB_DEADLINE_MS_') >= 0,
    'Procesor i odzyskiwanie muszą egzekwować twardy limit czasu.'
  );
}

function testInventoryStatusColorLifecycle514_() {
  assertCondition_(String(saveImportItems).indexOf('refreshInventoryStatusColors_') >= 0,
    'Import i ponowny import muszą odświeżać kolory.');
  assertCondition_(String(clearCurrentInventory).indexOf('clearInventoryStatusColors_') >= 0,
    'Czyszczenie musi usuwać kolory.');
  assertCondition_(String(startNewInventory).indexOf('clearInventoryStatusColors_') >= 0,
    'Nowa inwentaryzacja musi zaczynać się bez kolorów.');
  assertCondition_(String(finalizeInventoryAndExport).indexOf('clearInventoryStatusColors_') >= 0,
    'Zakończenie inwentaryzacji musi usuwać kolory.');
  assertCondition_(String(auditInventoryFormulaCoverage_).indexOf('refreshInventoryStatusColors_') === -1,
    'Audyt nie może kolorować arkusza.');
}

function testInventoryFinishClearsConfiguredInputs515_() {
  const clearSource = String(clearCurrentInventoryData_);
  const finishSource = String(finalizeInventoryAndExport);
  assertCondition_(
    clearSource.indexOf('getInputColumnsForProductType_') >= 0 &&
      clearSource.indexOf('product.columns') === -1,
    'Czyszczenie musi pobierać kolumny wejściowe z CONFIG, nie z cache produktu.'
  );
  assertCondition_(
    clearSource.indexOf('residualCells') >= 0,
    'Czyszczenie musi sprawdzić, czy po operacji nie pozostały dane.'
  );
  assertCondition_(
    finishSource.indexOf('clearCurrentInventoryData_') <
      finishSource.indexOf('closeActiveInventorySession_'),
    'Sesję wolno zamknąć dopiero po potwierdzonym wyczyszczeniu danych.'
  );
}

function testInventoryColorLiveRefresh516_() {
  const source = String(clearInventoryStatusColors_);
  assertCondition_(
    source.indexOf("setBackground('#ffffff')") >= 0 &&
      source.indexOf("setFontColor('#000000')") >= 0,
    'Reset kolorów musi ustawiać jawne formatowanie bazowe.'
  );
  assertCondition_(
    source.indexOf('SpreadsheetApp.flush()') >= 0,
    'Reset kolorów musi wymusić odświeżenie otwartego arkusza.'
  );
}

function testAudioScheduleIdempotency520_() {
  const keep = getInventoryAudioScheduleDecision_({
    now: 100000,
    requestedDueAt: 102000,
    existingDueAt: 101000,
    triggerCount: 1
  });
  const late = getInventoryAudioScheduleDecision_({
    now: 100000,
    requestedDueAt: 101000,
    existingDueAt: 110000,
    triggerCount: 1
  });
  const stale = getInventoryAudioScheduleDecision_({
    now: 100000,
    requestedDueAt: 101000,
    existingDueAt: 60000,
    triggerCount: 1
  });
  const missing = getInventoryAudioScheduleDecision_({
    now: 100000,
    requestedDueAt: 101000,
    existingDueAt: 0,
    triggerCount: 0
  });
  assertCondition_(keep.keepExisting === true,
    'Prawidłowego wyzwalacza nie wolno przesuwać przy kolejnym odczycie.');
  assertCondition_(late.keepExisting === false,
    'Wyzwalacz późniejszy od żądanego terminu musi zostać przeplanowany.');
  assertCondition_(stale.keepExisting === false && stale.stale === true,
    'Nieaktualny wyzwalacz musi zostać zastąpiony.');
  assertCondition_(missing.keepExisting === false,
    'Brakujący wyzwalacz musi zostać utworzony.');
}

function testAudioStatusReadOnly520_() {
  const source = String(getInventoryAudioJobs);
  assertCondition_(
    source.indexOf('ensureInventoryAudioProcessorScheduled_') === -1 &&
      source.indexOf('scheduleInventoryAudioProcessor_') === -1,
    'Odczyt statusu kolejki nie może tworzyć ani przeplanowywać wyzwalaczy.'
  );
}

function testAudioImmediateKick520_() {
  const source = String(kickInventoryAudioProcessorNow);
  assertCondition_(
    source.indexOf('processPendingInventoryAudioJobs_') >= 0,
    'Interfejs musi mieć bezpośrednią ścieżkę uruchomienia procesora audio.'
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
