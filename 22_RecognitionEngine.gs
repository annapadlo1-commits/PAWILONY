/**
 * Inventory PRO Enterprise v2.3.0
 * Bezpieczny silnik rozpoznawania nazw produktow.
 * UWAGA: modul nie dzieli tekstu i nie modyfikuje Parser.gs.
 */

function recognizeProduct_(inputName, runtimeContext) {
  const originalInput = String(inputName || '').trim();
  const normalizedInput = normalizeText(originalInput);
  const context = runtimeContext || buildRuntimeContext_();
  const memoKey = normalizeRecognitionForScore_(originalInput);

  if (!normalizedInput) {
    return recognitionResult_(false, 'EMPTY', originalInput, null, [], 0, 'Pusty wpis');
  }

  if (context.recognitionMemo && context.recognitionMemo[memoKey]) {
    if (context.performanceStats) context.performanceStats.memoHits++;
    return context.recognitionMemo[memoKey];
  }

  let result = null;

  if (context.productIndex[normalizedInput]) {
    result = recognitionResult_(true, 'EXACT', originalInput, context.productIndex[normalizedInput], [], 100, 'Dopasowanie dokladne');
  } else if (context.aliasIndex[normalizedInput]) {
    result = recognitionResult_(true, 'ALIAS', originalInput, context.aliasIndex[normalizedInput], [], 100, 'Dopasowano przez zapamietany alias');
  }

  if (!result) {
    const technicalMatches = findTechnicalVariantMatches_(originalInput, context);
    if (technicalMatches.length === 1) {
      result = recognitionResult_(true, 'VARIANT', originalInput, technicalMatches[0], [], 99, 'Bezpieczny wariant zapisu nazwy');
    } else if (technicalMatches.length > 1) {
      result = recognitionResult_(false, 'AMBIGUOUS', originalInput, null, toCandidateResults_(technicalMatches, 99), 99, 'Wariant pasuje do kilku produktow');
    }
  }

  if (!result) {
    const shortlist = getRecognitionShortlist_(originalInput, context);
    const finalistLimit = (CONFIG.PERFORMANCE && CONFIG.PERFORMANCE.LEVENSHTEIN_FINALISTS) || 5;
    const finalists = shortlist
      .map(product => ({ product: product, cheapScore: cheapRecognitionCandidateScore_(originalInput, product, context) }))
      .sort((a, b) => b.cheapScore - a.cheapScore || a.product.name.localeCompare(b.product.name))
      .slice(0, finalistLimit)
      .map(item => item.product);
    if (context.performanceStats) context.performanceStats.levenshteinCandidates += finalists.length;
    const scored = finalists
      .map(product => scoreRecognitionCandidate_(originalInput, product))
      .filter(candidate => candidate.score >= 45)
      .sort((a, b) => b.score - a.score || a.product.name.localeCompare(b.product.name));

    if (!scored.length) {
      result = recognitionResult_(false, 'NOT_FOUND', originalInput, null, [], 0, 'Nie znaleziono produktu');
    } else {
      const top = scored.slice(0, 10);
      const best = top[0];
      const second = top[1] || null;
      const margin = second ? best.score - second.score : best.score;
      if (best.score >= 94 && margin >= 9 && !best.numericConflict && !best.missingRequiredNumber) {
        result = recognitionResult_(true, 'SMART', originalInput, best.product, top, best.score, 'Inteligentne dopasowanie o wysokiej pewnosci');
      } else {
        result = recognitionResult_(false, 'AMBIGUOUS', originalInput, null, top, best.score, 'Wybierz produkt z listy');
      }
    }
  }

  if (context.recognitionMemo) context.recognitionMemo[memoKey] = result;
  return result;
}

function getRecognitionShortlist_(inputName, context) {
  const normalized = normalizeRecognitionForScore_(inputName);
  const first = normalized.split(' ').filter(Boolean)[0] || '';
  const indexed = first && context.firstTokenIndex ? context.firstTokenIndex[first] : null;
  const limit = (CONFIG.PERFORMANCE && CONFIG.PERFORMANCE.FUZZY_SHORTLIST_SIZE) || 20;
  if (indexed && indexed.length && indexed.length <= limit) return indexed;

  if (!context.trigramRecognitionIndex) {
    context.trigramRecognitionIndex = buildTrigramRecognitionIndex_(context.catalog || []);
  }
  const counts = {};
  const products = {};
  const queryGrams = recognitionTrigrams_(normalized)
    .map(gram => ({ gram: gram, size: (context.trigramRecognitionIndex[gram] || []).length }))
    .filter(item => item.size > 0)
    .sort((a, b) => a.size - b.size || a.gram.localeCompare(b.gram))
    .slice(0, 8)
    .map(item => item.gram);
  queryGrams.forEach(gram => {
    (context.trigramRecognitionIndex[gram] || []).forEach(product => {
      const key = normalizeText(product.name);
      products[key] = product;
      counts[key] = (counts[key] || 0) + 1;
    });
  });
  const shortlisted = Object.keys(counts)
    .sort((a, b) => counts[b] - counts[a] || a.localeCompare(b))
    .slice(0, limit)
    .map(key => products[key]);
  if (context.performanceStats) {
    context.performanceStats.fuzzyQueries++;
    context.performanceStats.fuzzyShortlisted += shortlisted.length;
  }
  return shortlisted.length ? shortlisted : (context.catalog || []).slice(0, limit);
}

function cheapRecognitionCandidateScore_(inputName, product, context) {
  const input = normalizeRecognitionForScore_(inputName);
  const inputGrams = recognitionTrigrams_(input);
  const inputNumbers = input.match(/\d+(?:[.,]\d+)?/g) || [];
  const inputSet = {};
  inputGrams.forEach(gram => inputSet[gram] = true);
  const productKey = normalizeText(product.name);
  const preparedSources = context.recognitionSourceTrigrams && context.recognitionSourceTrigrams[productKey];
  const sources = preparedSources ||
    [product.name].concat(product.aliases || []).map(normalizeRecognitionForScore_).map(value => ({
      value: value,
      grams: recognitionTrigrams_(value),
      numbers: value.match(/\d+(?:[.,]\d+)?/g) || []
    }));
  let best = 0;
  sources.forEach(source => {
    const grams = source.grams;
    let overlap = 0;
    grams.forEach(gram => { if (inputSet[gram]) overlap++; });
    const denominator = Math.max(1, inputGrams.length + grams.length - overlap);
    const sameNumbers = inputNumbers.length && source.numbers.length &&
      inputNumbers.join('|') === source.numbers.join('|');
    const conflictingNumbers = inputNumbers.length && source.numbers.length && !sameNumbers;
    const numericBonus = sameNumbers ? 2 : (conflictingNumbers ? -2 : 0);
    best = Math.max(best, overlap / denominator + numericBonus);
  });
  return best;
}

function recognitionResult_(matched, status, input, product, candidates, score, message) {
  return {
    matched: matched,
    status: status,
    input: input,
    product: product,
    candidates: candidates || [],
    score: Math.round(score || 0),
    message: message || ''
  };
}

function toCandidateResults_(products, score) {
  return (products || []).slice(0, 12).map(product => ({
    product: product,
    score: score
  }));
}

function findTechnicalVariantMatches_(inputName, contextOrCatalog) {
  const inputKeys = buildTechnicalRecognitionKeys_(inputName);
  if (!inputKeys.length) return [];
  const context = Array.isArray(contextOrCatalog)
    ? { catalog: contextOrCatalog, technicalRecognitionIndex: buildTechnicalRecognitionIndex_(contextOrCatalog) }
    : (contextOrCatalog || buildRuntimeContext_());
  const found = {};
  inputKeys.forEach(key => {
    (context.technicalRecognitionIndex[key] || []).forEach(product => {
      found[normalizeText(product.name)] = product;
    });
  });
  return Object.keys(found).map(key => found[key]);
}

function buildTechnicalRecognitionKeys_(value) {
  let base = String(value || '')
    .replace(/([A-Za-zÀ-ž])(?=\d)/g, '$1 ')
    .replace(/(\d)(?=[A-Za-zÀ-ž])/g, '$1 ')
    .replace(/[_\-/]+/g, ' ');

  base = normalizeRecognitionWords_(normalizeText(base));
  if (!base) return [];

  const keys = [];
  addUniqueKey_(keys, base);
  addUniqueKey_(keys, base.replace(/\s+/g, ''));

  const ageCanonical = base
    .replace(/\b(?:years old|year old|years|year|yrs|yr|yo)\b/g, 'yo')
    .replace(/\s+/g, ' ')
    .trim();
  addUniqueKey_(keys, ageCanonical);
  addUniqueKey_(keys, ageCanonical.replace(/\s+/g, ''));

  const ageOptional = ageCanonical
    .replace(/\s*yo\b/g, '')
    .replace(/\s+/g, ' ')
    .trim();
  addUniqueKey_(keys, ageOptional);
  addUniqueKey_(keys, ageOptional.replace(/\s+/g, ''));

  return keys;
}


/**
 * Wspolna funkcja wariantow uzywana rowniez przez stabilny Parser 2.2.5.
 * Zachowuje kompatybilnosc nazwy funkcji bez zmiany logiki parsera.
 */
function buildRecognitionVariants_(value) {
  return buildTechnicalRecognitionKeys_(value);
}

function addUniqueKey_(list, key) {
  if (key && !list.includes(key)) list.push(key);
}

function normalizeRecognitionWords_(text) {
  const numberWords = {
    zero: '0', one: '1', two: '2', three: '3', four: '4', five: '5',
    six: '6', seven: '7', eight: '8', nine: '9', ten: '10',
    eleven: '11', twelve: '12', thirteen: '13', fourteen: '14',
    fifteen: '15', sixteen: '16', seventeen: '17', eighteen: '18',
    nineteen: '19', twenty: '20',
    jeden: '1', jedna: '1', dwa: '2', dwie: '2', trzy: '3', cztery: '4',
    piec: '5', szesc: '6', siedem: '7', osiem: '8', dziewiec: '9',
    dziesiec: '10', jedenascie: '11', dwanascie: '12', trzynascie: '13',
    czternascie: '14', pietnascie: '15', szesnascie: '16',
    siedemnascie: '17', osiemnascie: '18', dziewietnascie: '19',
    dwadziescia: '20'
  };

  return String(text || '').split(' ').map(token => numberWords[token] || token).join(' ');
}

function scoreRecognitionCandidate_(inputName, product) {
  const input = normalizeRecognitionForScore_(inputName);
  const target = normalizeRecognitionForScore_(product.name);
  const aliases = (product.aliases || []).map(normalizeRecognitionForScore_);
  const sources = [target].concat(aliases);

  let best = { score: 0, numericConflict: false, missingRequiredNumber: false };
  sources.forEach(source => {
    const score = scoreRecognitionPair_(input, source);
    if (score.score > best.score) best = score;
  });

  return {
    product: product,
    score: Math.round(best.score),
    numericConflict: best.numericConflict,
    missingRequiredNumber: best.missingRequiredNumber
  };
}

function normalizeRecognitionForScore_(value) {
  return normalizeRecognitionWords_(normalizeText(
    String(value || '')
      .replace(/([A-Za-zÀ-ž])(?=\d)/g, '$1 ')
      .replace(/(\d)(?=[A-Za-zÀ-ž])/g, '$1 ')
      .replace(/[_\-/]+/g, ' ')
  )).replace(/\b(?:years old|year old|years|year|yrs|yr|yo)\b/g, 'yo')
    .replace(/\s+/g, ' ')
    .trim();
}

function scoreRecognitionPair_(input, target) {
  const inputTokens = input.split(' ').filter(Boolean);
  const targetTokens = target.split(' ').filter(Boolean);
  const inputNumbers = inputTokens.filter(isNumericRecognitionToken_);
  const targetNumbers = targetTokens.filter(isNumericRecognitionToken_);
  const numericConflict = inputNumbers.length > 0 && targetNumbers.length > 0 && inputNumbers.join('|') !== targetNumbers.join('|');
  const missingRequiredNumber = inputNumbers.length === 0 && targetNumbers.length > 0;

  if (numericConflict) {
    return { score: 15, numericConflict: true, missingRequiredNumber: false };
  }

  const inputWords = inputTokens.filter(token => !isNumericRecognitionToken_(token) && token !== 'yo');
  const targetWords = targetTokens.filter(token => !isNumericRecognitionToken_(token) && token !== 'yo');

  // Pokrycie tokenow musi tolerowac drobne literowki, np. bakardi -> bacardi.
  // Bierzemy najlepsze podobienstwo kazdego slowa do slow po drugiej stronie,
  // zamiast wymagac identycznego tokenu. Zgodnosc cyfr nadal jest osobnym,
  // twardym zabezpieczeniem przed pomyleniem wariantow 4 / 8 / 10.
  const coverageInput = averageBestTokenSimilarity_(inputWords, targetWords);
  const coverageTarget = averageBestTokenSimilarity_(targetWords, inputWords);
  const similarity = calculateStringSimilarity_(input, target);
  const firstTokenSimilarity = calculateStringSimilarity_(inputWords[0] || '', targetWords[0] || '');

  let score = 0;
  score += 35 * coverageInput;
  score += 20 * coverageTarget;
  score += 25 * similarity;
  score += 15 * firstTokenSimilarity;

  if (inputNumbers.length && targetNumbers.length && inputNumbers.join('|') === targetNumbers.join('|')) score += 12;
  if (input === target) score = 100;
  if (missingRequiredNumber) score -= 18;
  if (inputWords.length === 1 && targetWords.length > 1) score -= 8;

  return {
    score: Math.max(0, Math.min(100, score)),
    numericConflict: false,
    missingRequiredNumber: missingRequiredNumber
  };
}


function averageBestTokenSimilarity_(sourceTokens, targetTokens) {
  if (!sourceTokens.length || !targetTokens.length) return 0;

  const total = sourceTokens.reduce((sum, sourceToken) => {
    let best = 0;
    targetTokens.forEach(targetToken => {
      best = Math.max(best, calculateStringSimilarity_(sourceToken, targetToken));
    });
    return sum + best;
  }, 0);

  return total / sourceTokens.length;
}

function isNumericRecognitionToken_(token) {
  return /^\d+(?:[.,]\d+)?$/.test(String(token || ''));
}