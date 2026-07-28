/**
 * Inventory PRO — bezpieczna transkrypcja nagrań przez Gemini.
 * Klucz jest przechowywany wyłącznie w Script Properties.
 */
const GEMINI_TRANSCRIPTION_PROPERTY_ = 'INVENTORY_PRO_GEMINI_API_KEY';
const GEMINI_TRANSCRIPTION_MODELS_ = Object.freeze([
  'gemini-3.5-flash',
  'gemini-3.1-flash-lite'
]);
const GEMINI_MAX_AUDIO_BYTES_ = 20000000;
const GEMINI_MAX_AUDIO_SECONDS_ = 605;
const GEMINI_AUDIO_JOB_PREFIX_ = 'INVENTORY_AUDIO_JOB_';
const GEMINI_AUDIO_JOB_TTL_MS_ = 24 * 60 * 60 * 1000;
const GEMINI_AUDIO_FOLDER_PROPERTY_ = 'INVENTORY_AUDIO_TEMP_FOLDER_ID';
const GEMINI_AUDIO_MAX_ATTEMPTS_ = 5;
const GEMINI_AUDIO_RETRY_BASE_MS_ = 15000;
const GEMINI_AUDIO_PROCESSING_TIMEOUT_MS_ = 5 * 60 * 1000;
const GEMINI_AUDIO_JOB_DEADLINE_MS_ = 30 * 60 * 1000;
const GEMINI_AUDIO_TRIGGER_DUE_PROPERTY_ = 'INVENTORY_AUDIO_PROCESSOR_DUE_AT';
const GEMINI_AUDIO_TRIGGER_HANDLER_ = 'processPendingInventoryAudioJobs_';

function isGeminiTranscriptionConfigured_() {
  return Boolean(getGeminiApiKey_());
}

function getGeminiApiKey_() {
  const scriptKey = PropertiesService.getScriptProperties()
    .getProperty(GEMINI_TRANSCRIPTION_PROPERTY_);
  if (scriptKey) return scriptKey;
  try {
    const documentProperties = PropertiesService.getDocumentProperties();
    return documentProperties && documentProperties
      .getProperty(GEMINI_TRANSCRIPTION_PROPERTY_) || '';
  } catch (error) {
    return '';
  }
}

function configureGeminiTranscription() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Inventory PRO — Gemini',
    'Wklej klucz Gemini API. Klucz zostanie zapisany wyłącznie w ustawieniach skryptu i nie będzie widoczny w aplikacji mobilnej.',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const key = String(response.getResponseText() || '').trim();
  if (!isSupportedGeminiKeyFormat_(key)) {
    ui.alert('Klucz jest niepełny albo ma nieobsługiwany format. Skopiuj cały klucz przyciskiem kopiowania w Google AI Studio. Obsługiwane są nowe klucze AQ. oraz starsze AIza.');
    return;
  }
  const check = checkGeminiApiKey_(key);
  if (!check.ok) {
    ui.alert('Google odrzucił klucz. Nie zapisano zmian.\n\n' + check.message);
    return;
  }
  PropertiesService.getScriptProperties()
    .setProperty(GEMINI_TRANSCRIPTION_PROPERTY_, key);
  try {
    const documentProperties = PropertiesService.getDocumentProperties();
    if (documentProperties) documentProperties
      .setProperty(GEMINI_TRANSCRIPTION_PROPERTY_, key);
  } catch (error) {
    console.warn('Nie udało się zapisać zapasowej konfiguracji dokumentu: ' + String(error));
  }
  ui.alert('Transkrypcja Gemini została skonfigurowana. Zaktualizuj wdrożenie aplikacji mobilnej do nowej wersji.');
}

function showGeminiTranscriptionStatus() {
  const key = getGeminiApiKey_();
  const check = key ? checkGeminiApiKey_(key) : { ok: false, message: 'Brak zapisanego klucza.' };
  SpreadsheetApp.getUi().alert(
    'Inventory PRO — Gemini',
    check.ok
      ? 'Klucz API jest skonfigurowany i został zaakceptowany przez Gemini.'
      : 'Transkrypcja nie jest gotowa. ' + check.message,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function isSupportedGeminiKeyFormat_(key) {
  return /^(?:AIza[\w-]{20,}|AQ\.[A-Za-z0-9_-]{20,})$/.test(String(key || '').trim());
}

function checkGeminiApiKey_(key) {
  try {
    const response = UrlFetchApp.fetch(
      'https://generativelanguage.googleapis.com/v1beta/models?pageSize=1',
      {
        method: 'get',
        headers: { 'x-goog-api-key': key },
        muteHttpExceptions: true
      }
    );
    const code = response.getResponseCode();
    if (code >= 200 && code < 300) return { ok: true, message: 'OK' };
    let body;
    try { body = JSON.parse(response.getContentText()); } catch (error) { body = {}; }
    const message = body && body.error && body.error.message;
    return { ok: false, message: 'HTTP ' + code + ': ' + (message || 'klucz nie został zaakceptowany.') };
  } catch (error) {
    return { ok: false, message: String(error && error.message || error) };
  }
}

function buildGeminiInventoryPrompt_() {
  const context = buildRuntimeContext_();
  const catalog = (context.catalog || []).map(function(product) {
    const aliases = (product.aliases || []).filter(Boolean).slice(0, 12);
    return product.name + (aliases.length ? ' | warianty mowy: ' + aliases.join(', ') : '');
  }).join('\n');
  const locations = (CONFIG.LOCATION_AREAS || []).map(function(area) {
    return area.label + ' | warianty mowy: ' +
      [area.label].concat(area.aliases || []).filter(Boolean).join(', ');
  }).join('\n');
  return [
    'Jesteś modułem transkrypcji Inventory PRO dla lokalu ' + CONFIG.LOCATION.NAME + '.',
    'Zwróć WYŁĄCZNIE transkrypt, bez komentarzy, nagłówków i Markdown.',
    'ZASADA BEZWZGLĘDNA: niczego nie pomijaj. Każde usłyszane słowo, nazwa, liczba i lokalizacja musi pozostać w wyniku.',
    'Nie wymyślaj, nie dopowiadaj i nie zastępuj nieznanego fragmentu nazwą z katalogu.',
    'Gdy rozpoznanie nazwy z katalogu jest jednoznaczne, możesz poprawić wyłącznie jej pisownię do nazwy kanonicznej.',
    'Gdy nie masz wysokiej pewności, zapisz fragment możliwie dosłownie lub fonetycznie; parser pokaże go użytkownikowi.',
    'Każdy wypowiedziany produkt umieść w osobnym wierszu razem z wypowiedzianą wartością.',
    'Zachowuj liczby dziesiętne oraz pojemności. „zero siedem” zapisuj jako 0,7; „jeden litr” jako 1 litr.',
    'Słowo „sztuk” pozostaw przy wartości. Nie dopisuj pojemności, której nie wypowiedziano.',
    'Nie wybieraj między produktami o różnych pojemnościach. Nie wymyślaj nazw ani wartości.',
    'Nazwy lokalizacji zachowuj dokładnie w miejscu wypowiedzenia, również gdy są osobnym nagłówkiem lub kontekstem dla kolejnych produktów.',
    'LOKALIZACJE I ICH WARIANTY:',
    locations,
    'KATALOG PRODUKTÓW I WARIANTY MOWY (wyłącznie pomoc w pisowni):',
    catalog
  ].join('\n');
}

function transcribeInventoryAudio(base64Audio, mimeType, durationSeconds) {
  registerInventorySpreadsheet_();
  const key = getGeminiApiKey_();
  if (!key) throw new Error('Transkrypcja Gemini nie została skonfigurowana przez administratora.');
  const normalizedMimeType = normalizeGeminiAudioMimeType_(mimeType);
  const duration = Number(durationSeconds);
  if (!Number.isFinite(duration) || duration <= 0 || duration > GEMINI_MAX_AUDIO_SECONDS_) {
    throw new Error('Nagranie ma nieprawidłową długość lub przekracza limit 10 minut.');
  }
  const cleanBase64 = String(base64Audio || '').replace(/^data:[^;]+;base64,/, '');
  if (!cleanBase64 || cleanBase64.length > Math.ceil(GEMINI_MAX_AUDIO_BYTES_ * 4 / 3) + 8) {
    throw new Error('Nagranie jest puste albo zbyt duże.');
  }
  const bytes = Utilities.base64Decode(cleanBase64);
  if (bytes.length > GEMINI_MAX_AUDIO_BYTES_) throw new Error('Nagranie przekracza bezpieczny limit 20 MB.');

  const prompt = buildGeminiInventoryPrompt_();
  const payload = {
    contents: [{ parts: [
      { text: prompt },
      { inlineData: { mimeType: normalizedMimeType, data: cleanBase64 } }
    ] }],
    generationConfig: {
      temperature: 0,
      responseMimeType: 'text/plain'
    }
  };
  let body = {};
  let code = 0;
  let selectedModel = '';
  for (let modelIndex = 0; modelIndex < GEMINI_TRANSCRIPTION_MODELS_.length; modelIndex++) {
    const model = GEMINI_TRANSCRIPTION_MODELS_[modelIndex];
    const url = 'https://generativelanguage.googleapis.com/v1beta/models/' +
      model + ':generateContent';
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      headers: { 'x-goog-api-key': key },
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    code = response.getResponseCode();
    try { body = JSON.parse(response.getContentText()); } catch (error) { body = {}; }
    if (code >= 200 && code < 300) {
      selectedModel = model;
      break;
    }
    if ([404, 429, 500, 502, 503, 504].indexOf(code) === -1) break;
  }
  if (code < 200 || code >= 300) {
    const apiMessage = body && body.error && body.error.message;
    throw new Error('Gemini nie wykonał transkrypcji (HTTP ' + code + '). ' +
      (apiMessage ? String(apiMessage).slice(0, 300) : 'Spróbuj ponownie.'));
  }
  const candidates = body.candidates || [];
  const parts = candidates[0] && candidates[0].content && candidates[0].content.parts || [];
  const transcript = parts.map(function(part) { return part.text || ''; }).join('').trim();
  if (!transcript) throw new Error('Gemini nie zwrócił transkryptu. Nagraj próbkę ponownie.');
  return {
    transcript: transcript,
    durationSeconds: duration,
    model: selectedModel
  };
}

function normalizeGeminiAudioMimeType_(mimeType, fileName) {
  const raw = String(mimeType || '').toLowerCase().split(';')[0].trim();
  const extension = String(fileName || '').toLowerCase().split('.').pop();
  const aliases = {
    'audio/x-m4a': 'audio/mp4',
    'audio/m4a': 'audio/mp4',
    'audio/mp4': 'audio/mp4',
    'video/mp4': 'audio/mp4',
    'audio/mpeg': 'audio/mpeg',
    'audio/mp3': 'audio/mpeg',
    'audio/aac': 'audio/aac',
    'audio/x-aac': 'audio/aac',
    'audio/wav': 'audio/wav',
    'audio/x-wav': 'audio/wav',
    'audio/ogg': 'audio/ogg',
    'audio/webm': 'audio/webm'
  };
  if (aliases[raw]) return aliases[raw];
  const byExtension = {
    m4a: 'audio/mp4', mp4: 'audio/mp4', mp3: 'audio/mpeg',
    aac: 'audio/aac', wav: 'audio/wav', ogg: 'audio/ogg', webm: 'audio/webm'
  };
  if (byExtension[extension]) return byExtension[extension];
  throw new Error('Nieobsługiwany format audio. Użyj M4A, MP3, AAC, WAV, OGG albo WebM.');
}

function getGeminiTranscriptionAvailability() {
  return {
    configured: isGeminiTranscriptionConfigured_(),
    maxSeconds: GEMINI_MAX_AUDIO_SECONDS_ - 5,
    maxBytes: GEMINI_MAX_AUDIO_BYTES_
  };
}

/**
 * Trwałe zadania audio. Po zapisaniu pliku transkrypcja nie zależy już od
 * otwartej karty telefonu. Token chroni odczyt zadania przed zgadywaniem ID.
 */
function queueInventoryAudioJob(base64Audio, mimeType, durationSeconds, originalName, requestedId, requestedToken) {
  registerInventorySpreadsheet_();
  if (!isGeminiTranscriptionConfigured_()) {
    throw new Error('Transkrypcja Gemini nie została skonfigurowana przez administratora.');
  }
  const duration = Number(durationSeconds);
  if (!Number.isFinite(duration) || duration <= 0 || duration > GEMINI_MAX_AUDIO_SECONDS_) {
    throw new Error('Nagranie przekracza limit 10 minut.');
  }
  const cleanBase64 = String(base64Audio || '').replace(/^data:[^;]+;base64,/, '');
  if (!cleanBase64 || cleanBase64.length > Math.ceil(GEMINI_MAX_AUDIO_BYTES_ * 4 / 3) + 8) {
    throw new Error('Nagranie jest puste albo przekracza limit 20 MB.');
  }
  const bytes = Utilities.base64Decode(cleanBase64);
  if (bytes.length > GEMINI_MAX_AUDIO_BYTES_) throw new Error('Nagranie przekracza limit 20 MB.');

  cleanupExpiredInventoryAudioJobs_();
  const jobId = /^[a-f0-9-]{20,}$/i.test(String(requestedId || '')) ? String(requestedId) : Utilities.getUuid();
  const token = String(requestedToken || '').length >= 32 ? String(requestedToken) : Utilities.getUuid() + Utilities.getUuid();
  const existing = loadInventoryAudioJob_(jobId);
  if (existing) {
    if (existing.token !== token) throw new Error('Identyfikator zadania jest już używany.');
    return publicInventoryAudioJob_(existing);
  }
  const safeName = String(originalName || 'nagranie.wav').replace(/[^\w.\-ąćęłńóśźżĄĆĘŁŃÓŚŹŻ ]/g, '_').slice(0, 80);
  const normalizedMimeType = normalizeGeminiAudioMimeType_(mimeType, safeName);
  const file = getInventoryAudioTempFolder_().createFile(Utilities.newBlob(bytes, normalizedMimeType, 'InventoryPRO-' + jobId + '-' + safeName));
  const job = {
    id: jobId,
    token: token,
    status: 'QUEUED',
    fileId: file.getId(),
    originalName: safeName,
    mimeType: normalizedMimeType,
    durationSeconds: duration,
    createdAt: Date.now(),
    updatedAt: Date.now(),
    attempts: 0,
    nextAttemptAt: 0,
    error: ''
  };
  saveInventoryAudioJob_(job);
  scheduleInventoryAudioProcessor_(1000);
  return publicInventoryAudioJob_(job);
}

function getInventoryAudioJobs(requests) {
  cleanupExpiredInventoryAudioJobs_();
  recoverStaleInventoryAudioJobs_();
  ensureInventoryAudioProcessorScheduled_();
  return (Array.isArray(requests) ? requests : []).map(function(request) {
    const job = loadInventoryAudioJob_(request && request.id);
    if (!job || job.token !== String(request && request.token || '')) {
      return { id: String(request && request.id || ''), status: 'MISSING', error: 'Zadanie wygasło lub nie istnieje.' };
    }
    return publicInventoryAudioJob_(job);
  });
}

function retryInventoryAudioJob(id, token) {
  const job = loadInventoryAudioJob_(id);
  if (!job || job.token !== String(token || '')) throw new Error('Nie znaleziono zadania.');
  if (!job.fileId) throw new Error('Plik nagrania nie jest już dostępny.');
  job.status = 'QUEUED';
  job.error = '';
  job.attempts = 0;
  job.nextAttemptAt = 0;
  job.updatedAt = Date.now();
  saveInventoryAudioJob_(job);
  scheduleInventoryAudioProcessor_(1000);
  return publicInventoryAudioJob_(job);
}

function dismissInventoryAudioJob(id, token) {
  const job = loadInventoryAudioJob_(id);
  if (!job || job.token !== String(token || '')) return { removed: false };
  deleteInventoryAudioJobFiles_(job);
  PropertiesService.getScriptProperties().deleteProperty(GEMINI_AUDIO_JOB_PREFIX_ + job.id);
  return { removed: true };
}

function processPendingInventoryAudioJobs_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  let nextProcessorDelay = 0;
  try {
    const properties = PropertiesService.getScriptProperties();
    removeInventoryAudioProcessorTriggers_();
    properties.deleteProperty(GEMINI_AUDIO_TRIGGER_DUE_PROPERTY_);
    recoverStaleInventoryAudioJobs_();
    const all = properties.getProperties();
    const keys = Object.keys(all).filter(function(key) {
      return key.indexOf(GEMINI_AUDIO_JOB_PREFIX_) === 0;
    }).sort(function(left, right) {
      let a = {}, b = {};
      try { a = JSON.parse(all[left]); } catch (error) {}
      try { b = JSON.parse(all[right]); } catch (error) {}
      return Number(a.createdAt || 0) - Number(b.createdAt || 0);
    });
    const now = Date.now();
    const workerStartedAt = Date.now();
    let processed = 0;
    for (let keyIndex = 0; keyIndex < keys.length; keyIndex++) {
      const key = keys[keyIndex];
      let job;
      try { job = JSON.parse(all[key]); } catch (error) { continue; }
      if (!job || job.status !== 'QUEUED') continue;
      if (Number(job.nextAttemptAt || 0) > now) continue;
      if (now - Number(job.createdAt || now) > GEMINI_AUDIO_JOB_DEADLINE_MS_) {
        job.status = 'ERROR';
        job.nextAttemptAt = 0;
        job.updatedAt = Date.now();
        job.error = 'Przekroczono 30-minutowy limit przetwarzania. Nagranie zachowano — użyj „Ponów”.';
        saveInventoryAudioJob_(job);
        continue;
      }
      job.status = 'PROCESSING';
      job.updatedAt = Date.now();
      saveInventoryAudioJob_(job);
      try {
        const audioFile = DriveApp.getFileById(job.fileId);
        const result = transcribeInventoryAudio(
          Utilities.base64Encode(audioFile.getBlob().getBytes()),
          job.mimeType || audioFile.getBlob().getContentType(),
          job.durationSeconds
        );
        const transcriptFile = getInventoryAudioTempFolder_().createFile(Utilities.newBlob(
          result.transcript,
          'text/plain',
          'InventoryPRO-transcript-' + job.id + '.txt'
        ));
        job.transcriptFileId = transcriptFile.getId();
        job.model = result.model || '';
        job.status = 'DONE';
        job.error = '';
      } catch (error) {
        const message = String(error && error.message || error).slice(0, 600);
        job.attempts = Number(job.attempts || 0) + 1;
        if (isTransientGeminiError_(message) && job.attempts < GEMINI_AUDIO_MAX_ATTEMPTS_) {
          const delay = geminiAudioRetryDelayMs_(job.attempts);
          job.status = 'QUEUED';
          job.nextAttemptAt = Date.now() + delay;
          job.error = 'Chwilowa niedostępność Gemini. Automatyczna próba ' +
            (job.attempts + 1) + '/' + GEMINI_AUDIO_MAX_ATTEMPTS_ +
            ' za około ' + Math.ceil(delay / 1000) + ' s. ' + message;
        } else {
          job.status = 'ERROR';
          job.nextAttemptAt = 0;
          job.error = message;
        }
      }
      job.updatedAt = Date.now();
      saveInventoryAudioJob_(job);
      processed++;
      // W jednym uruchomieniu obsłuż kilka krótkich plików, ale pozostaw
      // bezpieczny zapas przed limitem czasu Apps Script.
      if (processed >= 3 || Date.now() - workerStartedAt > 230000) break;
    }
    const pendingJobs = Object.keys(PropertiesService.getScriptProperties().getProperties()).map(function(key) {
      if (key.indexOf(GEMINI_AUDIO_JOB_PREFIX_) !== 0) return null;
      return loadInventoryAudioJob_(key.slice(GEMINI_AUDIO_JOB_PREFIX_.length));
    }).filter(function(pending) {
      return pending && pending.status === 'QUEUED';
    });
    if (pendingJobs.length) {
      const earliest = pendingJobs.reduce(function(value, pending) {
        return Math.min(value, Number(pending.nextAttemptAt || Date.now() + 1000));
      }, Number.MAX_SAFE_INTEGER);
      nextProcessorDelay = Math.max(1000, earliest - Date.now());
    }
  } finally {
    lock.releaseLock();
  }
  if (nextProcessorDelay) scheduleInventoryAudioProcessor_(nextProcessorDelay);
}

function publicInventoryAudioJob_(job) {
  let transcript = '';
  if (job.status === 'DONE' && job.transcriptFileId) {
    try { transcript = DriveApp.getFileById(job.transcriptFileId).getBlob().getDataAsString('UTF-8'); }
    catch (error) { transcript = ''; }
  }
  return {
    id: job.id,
    token: job.token,
    status: job.status,
    originalName: job.originalName || 'nagranie',
    durationSeconds: job.durationSeconds,
    createdAt: job.createdAt,
    updatedAt: job.updatedAt,
    attempts: Number(job.attempts || 0),
    nextAttemptAt: Number(job.nextAttemptAt || 0),
    transcript: transcript,
    error: job.error || ''
  };
}

function isTransientGeminiError_(message) {
  return /\bHTTP\s+(?:429|500|502|503|504)\b/i.test(String(message || '')) ||
    /high demand|temporar(?:y|ily)|timeout|timed out|network/i.test(String(message || ''));
}

function geminiAudioRetryDelayMs_(attempt) {
  const exponent = Math.max(0, Number(attempt || 1) - 1);
  return Math.min(4 * 60 * 1000, GEMINI_AUDIO_RETRY_BASE_MS_ * Math.pow(2, exponent));
}

function scheduleInventoryAudioProcessor_(delayMs) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return false;
  try {
    const properties = PropertiesService.getScriptProperties();
    const delay = Math.max(1000, Number(delayMs || 1000));
    const dueAt = Date.now() + delay;
    // Jednorazowego triggera nie da się przeplanować przez zmianę właściwości.
    // Każde żądanie harmonogramu usuwa więc rzeczywisty stary trigger i tworzy
    // nowy z aktualnym terminem. To zamyka pętlę pozostawioną przez starsze wydanie.
    getInventoryAudioProcessorTriggers_().forEach(function(trigger) {
      try { ScriptApp.deleteTrigger(trigger); } catch (error) { console.warn(String(error)); }
    });
    properties.deleteProperty(GEMINI_AUDIO_TRIGGER_DUE_PROPERTY_);

    const trigger = ScriptApp.newTrigger(GEMINI_AUDIO_TRIGGER_HANDLER_)
      .timeBased()
      .after(delay)
      .create();
    properties.setProperty(GEMINI_AUDIO_TRIGGER_DUE_PROPERTY_, String(dueAt));
    return Boolean(trigger);
  } catch (error) {
    PropertiesService.getScriptProperties().deleteProperty(GEMINI_AUDIO_TRIGGER_DUE_PROPERTY_);
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function getInventoryAudioProcessorTriggers_() {
  return ScriptApp.getProjectTriggers().filter(function(trigger) {
    return trigger.getHandlerFunction() === GEMINI_AUDIO_TRIGGER_HANDLER_;
  });
}

function removeInventoryAudioProcessorTriggers_() {
  getInventoryAudioProcessorTriggers_().forEach(function(trigger) {
    try { ScriptApp.deleteTrigger(trigger); } catch (error) { console.warn(String(error)); }
  });
  PropertiesService.getScriptProperties().deleteProperty(GEMINI_AUDIO_TRIGGER_DUE_PROPERTY_);
}

/**
 * Jednorazowa naprawa projektu po starszych wydaniach.
 * Usuwa osierocone/zdublowane wyzwalacze i uruchamia kolejkę ponownie,
 * tylko gdy faktycznie istnieją oczekujące nagrania.
 */
function repairInventoryAudioProcessorTriggers() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) throw new Error('Procesor audio jest teraz zajęty. Spróbuj ponownie za chwilę.');
  try {
    removeInventoryAudioProcessorTriggers_();
  } finally {
    lock.releaseLock();
  }
  ensureInventoryAudioProcessorScheduled_();
  return {
    ok: true,
    triggerCount: getInventoryAudioProcessorTriggers_().length
  };
}

function getInventoryAudioRecoveryDecision_(job, now) {
  const current = Object.assign({}, job || {});
  if (
    ['DONE', 'ERROR'].indexOf(current.status) === -1 &&
    Number(now || Date.now()) - Number(current.createdAt || 0) > GEMINI_AUDIO_JOB_DEADLINE_MS_
  ) {
    current.status = 'ERROR';
    current.updatedAt = Number(now || Date.now());
    current.nextAttemptAt = 0;
    current.error = 'Przekroczono 30-minutowy limit przetwarzania. Nagranie zachowano — użyj „Ponów”.';
    return current;
  }
  if (current.status !== 'PROCESSING' ||
      Number(now || Date.now()) - Number(current.updatedAt || current.createdAt || 0) < GEMINI_AUDIO_PROCESSING_TIMEOUT_MS_) {
    return current;
  }
  current.attempts = Number(current.attempts || 0) + 1;
  current.updatedAt = Number(now || Date.now());
  current.nextAttemptAt = 0;
  if (current.attempts < GEMINI_AUDIO_MAX_ATTEMPTS_) {
    current.status = 'QUEUED';
    current.error = 'Poprzednia sesja przetwarzania została przerwana. Zadanie wznowiono automatycznie (' +
      (current.attempts + 1) + '/' + GEMINI_AUDIO_MAX_ATTEMPTS_ + ').';
  } else {
    current.status = 'ERROR';
    current.error = 'Transkrypcja była wielokrotnie przerywana. Użyj „Ponów”, aby rozpocząć nową serię prób.';
  }
  return current;
}

function recoverStaleInventoryAudioJobs_() {
  const properties = PropertiesService.getScriptProperties();
  const all = properties.getProperties();
  const now = Date.now();
  Object.keys(all).filter(function(key) {
    return key.indexOf(GEMINI_AUDIO_JOB_PREFIX_) === 0;
  }).forEach(function(key) {
    let job;
    try { job = JSON.parse(all[key]); } catch (error) { return; }
    const recovered = getInventoryAudioRecoveryDecision_(job, now);
    if (recovered.status !== job.status || recovered.updatedAt !== job.updatedAt) saveInventoryAudioJob_(recovered);
  });
}

function ensureInventoryAudioProcessorScheduled_() {
  const properties = PropertiesService.getScriptProperties();
  const all = properties.getProperties();
  const queued = Object.keys(all).some(function(key) {
    if (key.indexOf(GEMINI_AUDIO_JOB_PREFIX_) !== 0) return false;
    try { return JSON.parse(all[key]).status === 'QUEUED'; } catch (error) { return false; }
  });
  if (queued) scheduleInventoryAudioProcessor_(1000);
}

function saveInventoryAudioJob_(job) {
  PropertiesService.getScriptProperties().setProperty(
    GEMINI_AUDIO_JOB_PREFIX_ + job.id,
    JSON.stringify(job)
  );
}

function loadInventoryAudioJob_(id) {
  const raw = PropertiesService.getScriptProperties().getProperty(GEMINI_AUDIO_JOB_PREFIX_ + String(id || ''));
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (error) { return null; }
}

function cleanupExpiredInventoryAudioJobs_() {
  const properties = PropertiesService.getScriptProperties();
  const all = properties.getProperties();
  const cutoff = Date.now() - GEMINI_AUDIO_JOB_TTL_MS_;
  Object.keys(all).filter(function(key) { return key.indexOf(GEMINI_AUDIO_JOB_PREFIX_) === 0; }).forEach(function(key) {
    let job;
    try { job = JSON.parse(all[key]); } catch (error) { properties.deleteProperty(key); return; }
    if (Number(job.createdAt || 0) >= cutoff) return;
    deleteInventoryAudioJobFiles_(job);
    properties.deleteProperty(key);
  });
}

function deleteInventoryAudioJobFiles_(job) {
  [job && job.fileId, job && job.transcriptFileId].filter(Boolean).forEach(function(id) {
    try { DriveApp.getFileById(id).setTrashed(true); } catch (error) { console.warn(String(error)); }
  });
}

function getInventoryAudioTempFolder_() {
  const properties = PropertiesService.getScriptProperties();
  const configuredId = properties.getProperty(GEMINI_AUDIO_FOLDER_PROPERTY_);
  if (configuredId) {
    try { return DriveApp.getFolderById(configuredId); } catch (error) { console.warn(String(error)); }
  }
  const folder = DriveApp.createFolder('Inventory PRO — pliki tymczasowe audio');
  properties.setProperty(GEMINI_AUDIO_FOLDER_PROPERTY_, folder.getId());
  return folder;
}
