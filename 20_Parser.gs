/**
 * Inventory PRO 4.0 — jedyny publiczny punkt wejścia Parsera.
 * Parser zachowuje twarde granice wierszy, a dla tekstu ciągłego używa
 * longest-match opartego o rzeczywiste nazwy i aliasy katalogowe.
 */
function parseInventoryText(inputText, runtimeContext) {
  const raw = String(inputText || '').replace(/\r\n?/g, '\n').trim();
  if (!raw) return [];
  const context = runtimeContext || buildRuntimeContext_();
  const lines = raw.split(/\n+/).map(line => line.trim()).filter(Boolean);

  // Twarde granice wierszy: żaden produkt ani wartość nie może przeciec
  // do następnego wiersza. Tryb ciągły pozostaje dla dyktowania bez Enterów.
  if (lines.length <= 1) return parseInventoryTextContinuous_(raw, context);

  const results = [];
  let currentLocation = '';
  lines.forEach(line => {
    const prepared = prepareParserText_(line);
    const lineTokens = prepared.split(/\s+/).filter(Boolean);
    const leadingLocation = readLocationAt_(lineTokens, 0);
    if (leadingLocation) currentLocation = leadingLocation.location;
    if (leadingLocation && leadingLocation.consumed === lineTokens.length) {
      return;
    }
    const parsed = parseInventoryTextContinuous_(line, context);
    parsed.forEach(item => {
      if (!item.location && currentLocation) item.location = currentLocation;
      if (item.location && item.originalInput && !normalizeText(item.originalInput).startsWith(normalizeText(item.location))) {
        item.originalInput = item.location + ' ' + item.originalInput;
      }
      results.push(item);
    });
  });
  return results;
}

/**
 * Inventory PRO Enterprise v2.1.3 Recovery
 *
 * Parser:
 * - nie wymaga Enterow,
 * - zachowuje przecinki dziesietne,
 * - rozpoznaje kropke po liczbie,
 * - obsluguje sekcyjne lokalizacje,
 * - rozdziela kolejne produkty po wartosci.
 */

function parseInventoryTextContinuous_(inputText, runtimeContext) {
  const text = prepareParserText_(inputText);

  if (!text) {
    return [];
  }

  const tokens = text.split(/\s+/).filter(Boolean);
  const results = [];
  const context = runtimeContext || buildRuntimeContext_();

  let currentLocation = '';
  let position = 0;

  while (position < tokens.length) {
    const connector = normalizeWordForParser_(tokens[position]);

    if (isConnectorWord_(connector) || isQuantityUnitWord_(connector)) {
      position++;
      continue;
    }

    const locationMatch = readLocationAt_(tokens, position);

    if (locationMatch) {
      currentLocation = locationMatch.location;
      position += locationMatch.consumed;
      continue;
    }

    // v3.0 RC2: szybka ścieżka dictionary-first. Najpierw blokujemy
    // najdłuższy dokładny alias / nazwę produktu, a dopiero potem
    // odczytujemy liczbę pozostającą poza nazwą.
    const dictionaryEntry = findDictionaryFirstInventoryEntryAt_(
      tokens, position, context
    );

    if (dictionaryEntry) {
      const effectiveLocation = dictionaryEntry.location || currentLocation;
      results.push({
        originalInput: [effectiveLocation, dictionaryEntry.originalInput].filter(Boolean).join(' '),
        product: dictionaryEntry.product,
        value: dictionaryEntry.value,
        location: effectiveLocation,
        status: dictionaryEntry.value === null ? 'ERROR' : 'OK',
        message: dictionaryEntry.value === null ? 'Nie znaleziono ilosci' : 'Rozpoznano dictionary-first'
      });
      position = dictionaryEntry.nextPosition;
      continue;
    }

    // v2.10.3: część nagrań podaje wartość przed nazwą produktu,
    // np. "1.407 ardbeg 10". Najpierw próbujemy bezpiecznie rozpoznać
    // taki układ, zanim standardowy parser uzna liczbę z nazwy za wartość.
    const leadingValueEntry = findLeadingValueInventoryEntryAt_(
      tokens,
      position,
      context
    );

    if (leadingValueEntry) {
      results.push({
        originalInput: [
          currentLocation,
          leadingValueEntry.number.originalText,
          leadingValueEntry.originalProduct
        ].filter(Boolean).join(' '),
        product: leadingValueEntry.product,
        value: leadingValueEntry.number.value,
        location: currentLocation,
        status: 'OK',
        message: 'Rozpoznano'
      });

      position = leadingValueEntry.nextPosition;
      continue;
    }

    const entryMatch = findBestInventoryEntryAt_(
      tokens,
      position,
      context
    );

    if (entryMatch) {
      results.push({
        originalInput: [
          currentLocation,
          entryMatch.originalProduct,
          entryMatch.number.originalText
        ].filter(Boolean).join(' '),
        product: entryMatch.product,
        value: entryMatch.number.value,
        location: currentLocation,
        status: 'OK',
        message: 'Rozpoznano'
      });

      position = entryMatch.nextPosition;
      continue;
    }

    const fallback = findFirstNumberBoundary_(tokens, position, context);

    if (fallback) {
      const rawProduct = cleanProductName_(
        tokens.slice(position, fallback.numberPosition)
          .map(cleanNameToken_)
          .join(' ')
      );

      if (rawProduct) {
        results.push({
          originalInput: [
            currentLocation,
            rawProduct,
            fallback.number.originalText
          ].filter(Boolean).join(' '),
          product: normalizeRecognitionInput_(rawProduct),
          value: fallback.number.value,
          location: currentLocation,
          status: 'OK',
          message: 'Rozpoznano bez dopasowania katalogowego'
        });
      }

      position = fallback.nextPosition;
      continue;
    }

    const remainingProduct = cleanProductName_(
      tokens.slice(position)
        .map(cleanNameToken_)
        .join(' ')
    );

    if (remainingProduct) {
      results.push({
        originalInput: [
          currentLocation,
          remainingProduct
        ].filter(Boolean).join(' '),
        product: normalizeRecognitionInput_(remainingProduct),
        value: null,
        location: currentLocation,
        status: 'ERROR',
        message: 'Nie znaleziono ilosci'
      });
    }

    break;
  }

  return results;
}


/**
 * Production parser 3.0 RC2: exact dictionary-first.
 * Obsluguje wartosc przed i po produkcie oraz nie pobiera liczb
 * nalezacych do nazwy (12yo, 10, 0%, 0,5L, "pol").
 */
function findDictionaryFirstInventoryEntryAt_(tokens, startPosition, context) {
  const leading = readNumberAt_(tokens, startPosition);
  if (leading) {
    const span = findLongestExactCatalogSpanAt_(tokens, startPosition + leading.consumed, context);
    if (span) {
      return {
        originalInput: tokens.slice(startPosition, span.endPosition).join(' '),
        product: span.product.name,
        value: leading.value,
        nextPosition: span.endPosition
      };
    }
  }

  const span = findLongestExactCatalogSpanAt_(tokens, startPosition, context);
  if (!span) return null;
  const explicitLocation = readLocationAt_(tokens, span.endPosition);
  const numberPosition = span.endPosition + (explicitLocation ? explicitLocation.consumed : 0);
  const number = readNumberAt_(tokens, numberPosition);
  const quantityUnitConsumed = number && isQuantityUnitWord_(
    tokens[numberPosition + number.consumed]
  ) ? 1 : 0;
  // Nie zatwierdzamy krótszej nazwy bez wartości, jeżeli po niej pozostaje
  // dalsza część wpisu. Przykład krytyczny:
  //   magazyn Inne beczki Pilsner 2
  // Katalog może zawierać także krótszą pozycję „Inne Beczki”. Wcześniej
  // parser emitował ją jako błąd, a „Pilsner 2” przypisywał do innego
  // pilsnera. Cały fragment musi przejść przez bezpieczną ocenę granicy.
  if (!number && numberPosition < tokens.length) {
    return null;
  }
  if (
    number &&
    isProtectedCatalogNumericNameToken_(tokens, startPosition, span.endPosition, context)
  ) {
    return null;
  }
  return {
    originalInput: tokens.slice(
      startPosition,
      number ? numberPosition + number.consumed + quantityUnitConsumed : numberPosition
    ).join(' '),
    product: span.product.name,
    value: number ? number.value : null,
    nextPosition: number ? numberPosition + number.consumed + quantityUnitConsumed : numberPosition,
    location: explicitLocation ? explicitLocation.location : ''
  };
}

function findLongestExactCatalogSpanAt_(tokens, startPosition, context) {
  if (startPosition >= tokens.length) return null;
  const runtime = context || buildRuntimeContext_();
  const trie = getParserPhraseTrie_(runtime);
  const maxEnd = Math.min(tokens.length, startPosition + 16);
  let node = trie;
  let best = null;

  for (let end = startPosition; end < maxEnd && node; end++) {
    const tokenParts = canonicalParserPhrase_(cleanNameToken_(tokens[end])).split(' ').filter(Boolean);
    for (let partIndex = 0; partIndex < tokenParts.length; partIndex++) {
      node = node.children[tokenParts[partIndex]] || null;
      if (!node) break;
    }
    if (node && node.products && node.products.length === 1) {
      best = {
        endPosition: end + 1,
        recognitionInput: node.products[0].name,
        product: node.products[0],
        matchedPhrase: tokens.slice(startPosition, end + 1).join(' ')
      };
    }
  }
  const nextAfterTrie = best ? normalizeWordForParser_(tokens[best.endPosition] || '') : '';
  if (best && !['zero', 'pol'].includes(nextAfterTrie)) return best;

  // Zachowujemy kompatybilny fallback dla transformacji technicznych,
  // których nie da się przejść token po tokenie w surowym tekście.
  const phraseIndex = getParserPhraseIndex_(runtime);
  best = null;

  for (let end = startPosition + 1; end <= maxEnd; end++) {
    const raw = cleanProductName_(
      tokens.slice(startPosition, end).map(cleanNameToken_).join(' ')
    );
    if (!raw) continue;

    const variants = buildParserRecognitionVariants_(raw);
    for (let variantIndex = 0; variantIndex < variants.length; variantIndex++) {
      const key = canonicalParserPhrase_(variants[variantIndex]);
      const products = phraseIndex[key] || [];
      // Kolizja dokładnego klucza nie może być rozstrzygana kolejnością
      // produktów w arkuszu. Przekazujemy ją do późniejszego matchera.
      if (products.length !== 1) continue;
      const product = products[0];
      best = {
        endPosition: end,
        recognitionInput: product.name,
        product: product,
        matchedPhrase: raw
      };
      break;
    }
  }
  return best;
}

function getParserPhraseTrie_(context) {
  const phraseIndex = getParserPhraseIndex_(context);
  if (context.parserPhraseTrie && context.parserPhraseTrieSourceIndex === phraseIndex) {
    return context.parserPhraseTrie;
  }
  context.parserPhraseTrie = buildParserPhraseTrieFromIndex_(phraseIndex);
  context.parserPhraseTrieSourceIndex = phraseIndex;
  return context.parserPhraseTrie;
}

function buildParserPhraseTrieFromIndex_(phraseIndex) {
  const root = { children: {}, products: null };
  Object.keys(phraseIndex || {}).forEach(key => {
    let node = root;
    key.split(' ').filter(Boolean).forEach(token => {
      if (!node.children[token]) node.children[token] = { children: {}, products: null };
      node = node.children[token];
    });
    node.products = phraseIndex[key];
  });
  return root;
}

function getParserPhraseIndex_(context) {
  if (context.parserPhraseIndex) return context.parserPhraseIndex;
  const index = {};
  (context.catalog || []).forEach(product => {
    [product.name].concat(product.aliases || []).forEach(value => {
      addParserPhraseVariants_(index, value, product);
    });
  });
  context.parserPhraseIndex = index;
  return index;
}

function addParserPhraseVariants_(index, value, product) {
  const rawVariants = [String(value || '')];
  // Do indeksu EXACT nie wolno dodawać wariantów technicznych usuwających
  // słowa (np. "years old"). Takie warianty są dopuszczalne wyłącznie w
  // późniejszym, ocenianym matcherze. Inaczej "Osco 2" staje się fałszywie
  // pełną nazwą "Osco 2 years old".
  buildParserRecognitionVariants_(value).forEach(key => rawVariants.push(key));

  rawVariants.forEach(variant => {
    const full = canonicalParserPhrase_(variant);
    if (full) addParserPhraseIndexProduct_(index, full, product);
    const withoutPackaging = stripParserPackagingSuffix_(full);
    if (withoutPackaging) addParserPhraseIndexProduct_(index, withoutPackaging, product);
  });
}

function addParserPhraseIndexProduct_(index, key, product) {
  if (!index[key]) index[key] = [];
  if (!index[key].some(item => normalizeText(item.name) === normalizeText(product.name))) {
    index[key].push(product);
  }
}

function canonicalParserPhrase_(value) {
  return normalizeText(normalizeRecognitionInput_(value))
    .replace(/\s+/g, ' ')
    .trim();
}

function stripParserPackagingSuffix_(canonical) {
  const tokens = String(canonical || '').split(' ').filter(Boolean);
  if (!tokens.length) return '';
  const units = ['l', 'ml', 'cl', 'kg', 'g'];
  if (!units.includes(tokens[tokens.length - 1])) return canonical;

  tokens.pop();
  if (tokens.length && /^\d+$/.test(tokens[tokens.length - 1])) tokens.pop();
  if (tokens.length && /^\d+$/.test(tokens[tokens.length - 1]) && tokens[tokens.length - 1] === '0') tokens.pop();
  return tokens.join(' ').trim();
}

/**
 * Szuka najlepszego podzialu: NAZWA PRODUKTU | WARTOSC.
 * Liczba moze nalezec do nazwy (Bacardi 8, Auchentoshan 12yo),
 * dlatego sprawdzamy wszystkie sensowne granice, zamiast konczyc
 * nazwe na pierwszej liczbie.
 */
function findBestInventoryEntryAt_(tokens, startPosition, context) {
  const maxNameTokens = 10;
  const maxPosition = Math.min(
    tokens.length,
    startPosition + maxNameTokens + 2
  );
  const candidates = [];

  for (
    let numberPosition = startPosition + 1;
    numberPosition < maxPosition;
    numberPosition++
  ) {
    const number = readNumberAt_(tokens, numberPosition);

    if (!number) {
      continue;
    }

    // v2.9.2: slowo 'zero' / zapis '0%' moze byc czescia nazwy
    // produktu. Nie wolno konczyc nazwy w tym miejscu, jezeli dluzszy
    // fragment jest poprawnym, zakotwiczonym dopasowaniem katalogowym.
    if (
      isProtectedZeroNameToken_(tokens, startPosition, numberPosition, context) ||
      isProtectedCatalogNumericNameToken_(tokens, startPosition, numberPosition, context) ||
      isProtectedVolumePhraseToken_(tokens, startPosition, numberPosition, context)
    ) {
      continue;
    }

    const originalProduct = cleanProductName_(
      tokens.slice(startPosition, numberPosition)
        .map(cleanNameToken_)
        .join(' ')
    );

    if (!originalProduct) {
      continue;
    }

    const recognitionInput = normalizeRecognitionInput_(
      originalProduct
    );
    const parserMatch = matchProductForParser_(recognitionInput, context);
    const match = parserMatch.match;

    // Krytyczna zasada: dopasowanie musi zaczynac sie dokladnie
    // w miejscu, w ktorym parser aktualnie stoi. Matcher nie moze
    // pominac poczatku wpisu i wybrac produktu znajdujacego sie
    // dopiero pozniej w analizowanym fragmencie.
    if (!isParserMatchAnchoredAtStart_(recognitionInput, match)) {
      continue;
    }

    const score = scoreParserBoundary_(
      recognitionInput,
      match,
      numberPosition - startPosition
    );

    if (score < 60) {
      continue;
    }

    const quantityUnitConsumed = isQuantityUnitWord_(
      tokens[numberPosition + number.consumed]
    ) ? 1 : 0;

    candidates.push({
      originalProduct: originalProduct,
      product: canonicalMatchedProductName_(parserMatch),
      number: number,
      nextPosition: numberPosition + number.consumed + quantityUnitConsumed,
      score: score,
      match: match
    });
  }

  if (!candidates.length) {
    return null;
  }

  candidates.sort((a, b) => {
    if (b.score !== a.score) {
      return b.score - a.score;
    }

    // Przy remisie wybieramy dluzsza rozpoznana nazwe.
    return b.nextPosition - a.nextPosition;
  });

  return candidates[0];
}


/**
 * Zabezpiecza granice pozycji.
 *
 * Dla fragmentu:
 *   Jameson 12 Auchentoshan 12 1,234
 * matcher moze uznac, ze najlepszym produktem jest Auchentoshan 12.
 * Parser nie moze jednak zaakceptowac takiego wyniku, bo nazwa produktu
 * nie zaczyna sie od pierwszego tokenu analizowanego fragmentu.
 */
function isParserMatchAnchoredAtStart_(inputName, match) {
  if (!match || ['NOT_FOUND', 'EMPTY'].includes(match.status)) {
    return false;
  }

  const inputVariants = buildRecognitionVariants_(inputName);

  if (!inputVariants.length) {
    return false;
  }

  const products = [];

  if (match.product) {
    products.push(match.product);
  }

  (match.candidates || []).forEach(candidate => {
    if (candidate && candidate.product) {
      products.push(candidate.product);
    }
  });

  return products.some(product => {
    const targetVariants = buildRecognitionVariants_(product.name);

    (product.aliases || []).forEach(alias => {
      buildRecognitionVariants_(alias).forEach(variant => {
        if (!targetVariants.includes(variant)) {
          targetVariants.push(variant);
        }
      });
    });

    const exactAnchor = inputVariants.some(inputVariant =>
      targetVariants.some(targetVariant =>
        inputVariant === targetVariant ||
        targetVariant.startsWith(inputVariant + ' ')
      )
    );
    if (exactAnchor) return true;

    // Drobna literówka na początku nazwy nie może wyłączyć poprawnego
    // dopasowania SMART. Nadal wymagamy bardzo wysokiego wyniku, zgodnych
    // cyfr i podobnego pierwszego tokenu, więc matcher nie może przeskoczyć
    // do produktu występującego dopiero dalej w tekście.
    const scored = scoreRecognitionCandidate_(inputName, product);
    const inputFirst = normalizeRecognitionForScore_(inputName).split(' ')[0] || '';
    const targetFirst = normalizeRecognitionForScore_(product.name).split(' ')[0] || '';
    return scored.score >= 94 &&
      !scored.numericConflict &&
      !scored.missingRequiredNumber &&
      calculateStringSimilarity_(inputFirst, targetFirst) >= 0.78;
  });
}

function scoreParserBoundary_(inputName, match, nameTokenCount) {
  if (!match) {
    return 0;
  }

  const statusScores = {
    EXACT: 120,
    ALIAS: 118,
    VARIANT: 116,
    SMART: 108,
    AUTO: 100,
    AMBIGUOUS: 76,
    NOT_FOUND: 0,
    EMPTY: 0
  };

  let score = statusScores[match.status] || 0;
  const targetName = match.product
    ? match.product.name
    : (
        match.candidates && match.candidates.length
          ? match.candidates[0].product.name
          : ''
      );

  const inputNumbers = extractRecognitionNumbers_(inputName);
  const targetNumbers = extractRecognitionNumbers_(targetName);

  if (inputNumbers.length) {
    const allNumbersMatch = inputNumbers.every(number =>
      targetNumbers.includes(number)
    );

    score += allNumbersMatch ? 28 : -90;
  }

  if (containsAgeMarker_(inputName)) {
    score += 14;
  }

  if (match.score) {
    score += Math.min(12, match.score / 10);
  }

  score += Math.min(6, nameTokenCount * 0.75);

  return score;
}

function findFirstNumberBoundary_(tokens, startPosition, context) {
  for (
    let position = startPosition + 1;
    position < tokens.length;
    position++
  ) {
    const number = readNumberAt_(tokens, position);

    const fallbackContext = context || buildRuntimeContext_();
    if (number &&
        !isProtectedZeroNameToken_(tokens, startPosition, position, fallbackContext) &&
        !isProtectedCatalogNumericNameToken_(tokens, startPosition, position, fallbackContext) &&
        !isProtectedVolumePhraseToken_(tokens, startPosition, position, fallbackContext)) {
      return {
        numberPosition: position,
        number: number,
        nextPosition: position + number.consumed
      };
    }
  }

  return null;
}



/**
 * v2.10.3 — obsługa wpisów, w których wartość występuje przed nazwą,
 * np. "1.407 ardbeg 10". Akceptujemy wyłącznie zakotwiczone dopasowanie
 * katalogowe, dzięki czemu liczba nie połyka przypadkowego tekstu.
 */
function findLeadingValueInventoryEntryAt_(tokens, startPosition, context) {
  const number = readNumberAt_(tokens, startPosition);
  if (!number) return null;

  const nameStart = startPosition + number.consumed;
  if (nameStart >= tokens.length) return null;

  const maxEnd = Math.min(tokens.length, nameStart + 10);
  const candidates = [];

  for (let end = nameStart + 1; end <= maxEnd; end++) {
    const originalProduct = cleanProductName_(
      tokens.slice(nameStart, end).map(cleanNameToken_).join(' ')
    );
    if (!originalProduct) continue;

    const recognitionInput = normalizeRecognitionInput_(originalProduct);
    const parserMatch = matchProductForParser_(recognitionInput, context);
    const match = parserMatch.match;
    if (!isParserMatchAnchoredAtStart_(parserMatch.recognitionInput, match)) continue;

    const score = scoreParserBoundary_(
      parserMatch.recognitionInput,
      match,
      end - nameStart
    );
    if (score < 90) continue;

    candidates.push({
      originalProduct: originalProduct,
      product: canonicalMatchedProductName_(parserMatch),
      number: number,
      nextPosition: end,
      score: score
    });
  }

  if (!candidates.length) return null;
  candidates.sort(function(a, b) {
    if (b.score !== a.score) return b.score - a.score;
    return b.nextPosition - a.nextPosition;
  });
  return candidates[0];
}

/**
 * Chroni liczby będące częścią istniejącej nazwy/aliasu produktu,
 * np. Bacardi 8, Ardbeg 10, Singleton 15, Auchentoshan 18, Zacapa 23.
 * Liczba jest chroniona tylko wtedy, gdy dłuższy fragment rzeczywiście
 * pasuje do katalogu. Dzięki temu "campari 18" nadal oznacza stan 18.
 */
function isProtectedCatalogNumericNameToken_(tokens, startPosition, numberPosition, context) {
  const number = readNumberAt_(tokens, numberPosition);
  if (!number || number.consumed !== 1) return false;

  const raw = String(tokens[numberPosition] || '').replace(/[,:;]+$/g, '');
  if (!/^\d+$/.test(raw)) return false;

  const nextWord = normalizeWordForParser_(tokens[numberPosition + 1] || '');
  if (['yo', 'year', 'years', 'old', 'rok', 'lata', 'letni', 'letnia'].includes(nextWord)) {
    return true;
  }

  // Musi istnieć kolejny token, który może być właściwą wartością.
  // Bez niego przypadek pozostaje standardową nazwą + stanem.
  if (!readNumberAt_(tokens, numberPosition + 1)) return false;

  const longerName = cleanProductName_(
    tokens.slice(startPosition, numberPosition + 1).map(cleanNameToken_).join(' ')
  );
  if (!longerName) return false;

  const parserMatch = matchProductForParser_(
    normalizeRecognitionInput_(longerName),
    context || buildRuntimeContext_()
  );
  return isParserMatchAnchoredAtStart_(parserMatch.recognitionInput, parserMatch.match) &&
    !['NOT_FOUND', 'EMPTY'].includes(String(parserMatch.match && parserMatch.match.status || ''));
}

/**
 * Chroni słowne pojemności: "pół litra", "półlitrowa". Samo słowo
 * "pół" nadal może być wartością 0.5, gdy nie tworzy nazwy produktu.
 */
function isProtectedVolumePhraseToken_(tokens, startPosition, numberPosition, context) {
  const normalized = normalizeWordForParser_(tokens[numberPosition] || '');
  if (normalized !== 'pol') return false;

  const next = normalizeWordForParser_(tokens[numberPosition + 1] || '');
  const joined = normalizeWordForParser_(
    [tokens[numberPosition], tokens[numberPosition + 1]].filter(Boolean).join(' ')
  );
  const isVolumePhrase = ['litra', 'litr', 'litrowa', 'litrowy', 'l'].includes(next) ||
    /^pollitrow/.test(joined.replace(/\s+/g, ''));
  if (!isVolumePhrase) return false;

  const end = numberPosition + (next ? 2 : 1);
  const longerName = cleanProductName_(
    tokens.slice(startPosition, end).map(cleanNameToken_).join(' ')
  );
  if (!longerName) return false;

  const parserMatch = matchProductForParser_(
    normalizeRecognitionInput_(longerName),
    context || buildRuntimeContext_()
  );
  return isParserMatchAnchoredAtStart_(parserMatch.recognitionInput, parserMatch.match) &&
    !['NOT_FOUND', 'EMPTY'].includes(String(parserMatch.match && parserMatch.match.status || ''));
}

/**
 * Chroni warianty produktów zawierające zero/0% przed potraktowaniem
 * ich jako ilości. Zwykłe cyfrowe 0 pozostaje wartością, chyba że po nim
 * występuje kolejna liczba i cały dłuższy fragment pasuje do produktu.
 */
function isProtectedZeroNameToken_(tokens, startPosition, numberPosition, context) {
  const raw = String(tokens[numberPosition] || '').toLowerCase().replace(/[,:;]+$/g, '');
  const normalized = normalizeWordForParser_(raw);
  const textualZero = normalized === 'zero' || raw === '0%' || raw === '0,0%' || raw === '0.0%';
  const followingNumber = readNumberAt_(tokens, numberPosition + 1);
  const numericZeroWithFollowingValue = (raw === '0' || raw === '0,0' || raw === '0.0') &&
    Boolean(followingNumber);

  if (!textualZero && !numericZeroWithFollowingValue) return false;

  // v2.10.2: gdy po slowie "zero" wystepuje kolejna liczba, zero jest
  // elementem nazwy, a kolejna liczba jest wartoscia pozycji.
  // Przyklady: "kola zero 1", "lucano zero 2", "tanqueray 0% 3".
  if (textualZero && followingNumber) return true;

  const longerName = cleanProductName_(
    tokens.slice(startPosition, numberPosition + 1).map(cleanNameToken_).join(' ')
  );
  if (!longerName) return false;

  const parserMatch = matchProductForParser_(normalizeRecognitionInput_(longerName), context || buildRuntimeContext_());
  const match = parserMatch.match;
  return isParserMatchAnchoredAtStart_(parserMatch.recognitionInput, match) &&
    !['NOT_FOUND', 'EMPTY'].includes(match.status);
}


/**
 * v2.10.1 — dopasowanie nazw zawierajacych "zero" przed ekstrakcja wartosci.
 * Nie zmienia tekstu uzytkownika. Tworzy jedynie warianty rozpoznawcze:
 *   lucano zero -> lucano 0
 *   kola zero / cola zero -> kola bez cukru / cola bez cukru
 * Ostatecznie wybiera najlepsze zakotwiczone dopasowanie katalogowe.
 */

function buildParserRecognitionVariants_(value) {
  const source = normalizeRecognitionInput_(value);
  const variants = [source];
  const normalized = normalizeText(source);

  if (/\bzero\b/.test(normalized)) {
    variants.push(source.replace(/\bzero\b/gi, '0'));
    variants.push(source.replace(/\bzero\b/gi, '0%'));
    variants.push(source.replace(/\bzero\b/gi, 'bez cukru'));
  }

  // "kola 0 12": 0 jest wariantem nazwy, 12 jest wartoscia.
  // "fritz kola 0" pozostaje zwykla kola ze stanem zero.
  if (/^(?:kola|cola)\s+0$/i.test(source)) {
    variants.unshift(source.replace(/\b0\b/g, 'bez cukru'));
    variants.unshift('fritz 200ml ' + source.replace(/\b0\b/g, 'bez cukru'));
  }

  // Najczestszy wariant dyktowania softow: "kola zero" / "cola zero".
  if (/\b(?:kola|cola)\s+zero\b/i.test(source)) {
    variants.unshift(source.replace(/\bzero\b/gi, 'bez cukru'));
    variants.unshift('fritz 200ml ' + source.replace(/\bzero\b/gi, 'bez cukru'));
  }

  // Lucano Zero wystepuje w katalogu jako Amaro Lucano 0% 0,7L.
  if (/\blucano\s+zero\b/i.test(source) && !/\bamaro\b/i.test(source)) {
    variants.unshift('amaro ' + source.replace(/\bzero\b/gi, '0%'));
  }

  // v2.10.3: słowne warianty pojemności są częścią nazwy produktu,
  // a nie stanem magazynowym. ZERO pozostaje obsługiwane powyżej bez zmian.
  if (/\b(?:pol|pół)\s+(?:litra|litr|litrowa|litrowy|l)\b/i.test(source)) {
    variants.unshift(source.replace(/\b(?:pol|pół)\s+(?:litra|litr|litrowa|litrowy|l)\b/gi, '0,5l'));
    variants.unshift(source.replace(/\b(?:pol|pół)\s+(?:litra|litr|litrowa|litrowy|l)\b/gi, '500ml'));
  }
  if (/\b(?:pol|pół)litrow(?:a|y|e)?\b/i.test(source)) {
    variants.unshift(source.replace(/\b(?:pol|pół)litrow(?:a|y|e)?\b/gi, '0,5l'));
    variants.unshift(source.replace(/\b(?:pol|pół)litrow(?:a|y|e)?\b/gi, '500ml'));
  }

  return variants
    .map(normalizeRecognitionInput_)
    .filter((variant, index, all) => variant && all.indexOf(variant) === index);
}

function parserMatchStrength_(match) {
  if (!match) return -1000;
  const statusWeight = {
    EXACT: 1000,
    ALIAS: 950,
    VARIANT: 900,
    SMART: 800,
    AUTO: 750,
    AMBIGUOUS: 500,
    NOT_FOUND: -1000,
    EMPTY: -1000
  };
  const status = String(match.status || '');
  const candidateScore = Number(match.score ||
    (match.candidates && match.candidates[0] && match.candidates[0].score) || 0);
  return (statusWeight[status] || 0) + candidateScore;
}

function canonicalMatchedProductName_(parserMatch) {
  const match = parserMatch && parserMatch.match;
  if (match && match.product && match.product.name) return match.product.name;
  return parserMatch && parserMatch.recognitionInput
    ? parserMatch.recognitionInput
    : '';
}

function normalizeRecognitionInput_(value) {
  return String(value || '')
    .replace(/([A-Za-zÀ-ž])(?=\d)/g, '$1 ')
    .replace(/(\d)(?=[A-Za-zÀ-ž])/g, '$1 ')
    .replace(/\s+/g, ' ')
    .trim();
}

function extractRecognitionNumbers_(value) {
  const normalized = normalizeText(
    normalizeRecognitionInput_(value)
  );
  const matches = normalized.match(/\d+(?:[.,]\d+)?/g);

  return matches || [];
}

function containsAgeMarker_(value) {
  return /\b(?:yo|year|years|old|rok|lata|letni|letnia)\b/i.test(
    normalizeRecognitionInput_(value)
  );
}

function prepareParserText_(inputText) {
  return String(inputText || '')
    .replace(/\r\n?/g, '\n')
    .replace(/(\d)\s*,\s*(\d)/g, '$1.$2')
    .replace(/[\n;]+/g, ' ')
    .replace(/,\s*/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function cleanNameToken_(token) {
  return String(token || '')
    .replace(/^[,;:]+/, '')
    .replace(/[,;:]+$/, '')
    .trim();
}

function cleanProductName_(value) {
  return String(value || '')
    .replace(/\s+/g, ' ')
    .replace(/^(i|oraz|potem|dalej|nastepnie)\s+/i, '')
    .replace(/\s+(i|oraz|potem|dalej|nastepnie)$/i, '')
    .trim();
}

function isConnectorWord_(value) {
  return [
    'i',
    'oraz',
    'potem',
    'dalej',
    'nastepnie'
  ].includes(value);
}

/**
 * Jednostka wypowiadana po liczbie nie jest początkiem kolejnego produktu.
 * Obsługujemy najczęstsze formy z polskich transkrypcji głosowych.
 */
function isQuantityUnitWord_(value) {
  return [
    'szt',
    'sztuka',
    'sztuke',
    'sztuki',
    'sztuk'
  ].includes(normalizeWordForParser_(value));
}

function readLocationAt_(tokens, position) {
  const source = Array.isArray(tokens) ? tokens : [];
  const definitions = getLocationAreaDefinitions_();
  let best = null;

  definitions.forEach(area => {
    const candidates = [area.key, area.label].concat(area.aliases || []);
    candidates.forEach(candidate => {
      const normalized = normalizeWordForParser_(candidate || '');
      if (!normalized) return;
      const consumed = normalized.split(/\s+/).filter(Boolean).length;
      const phraseTokens = source
        .slice(position, position + consumed)
        .filter(token => token !== null && token !== undefined && String(token).trim() !== '');

      // Nie wolno uznać jednoelementowego końca wiersza za dwuwyrazowy alias.
      if (phraseTokens.length !== consumed) return;
      const phrase = normalizeWordForParser_(phraseTokens.join(' '));
      if (phrase !== normalized) return;
      if (!best || consumed > best.consumed) {
        best = { location: area.key, consumed: consumed };
      }
    });
  });

  return best;
}

function readNumberAt_(tokens, position) {
  const twoOriginal = [
    tokens[position],
    tokens[position + 1]
  ].filter(Boolean).join(' ');

  const twoValue = parseNumberText_(twoOriginal);

  if (twoValue !== null) {
    return {
      value: twoValue,
      consumed: 2,
      originalText: twoOriginal
    };
  }

  const oneOriginal = tokens[position] || '';
  const oneValue = parseNumberText_(oneOriginal);

  if (oneValue !== null) {
    return {
      value: oneValue,
      consumed: 1,
      originalText: oneOriginal
    };
  }

  return null;
}

function parseInventoryEntry_(entry) {
  const parsed = parseInventoryText(entry);

  return parsed.length
    ? parsed[0]
    : createParserResult_(
        String(entry || '').trim(),
        '',
        null,
        'ERROR',
        'Pusty wpis'
      );
}

function parseNumberText_(valueText) {
  const normalized = normalizeNumberText_(valueText);

  if (!normalized) {
    return null;
  }

  if (/^-?\d+(?:\.\d+)?$/.test(normalized)) {
    const numberValue = Number(normalized);

    return Number.isFinite(numberValue)
      ? numberValue
      : null;
  }

  const words = {
    zero: 0,
    jeden: 1,
    jedna: 1,
    dwa: 2,
    dwie: 2,
    trzy: 3,
    cztery: 4,
    piec: 5,
    szesc: 6,
    siedem: 7,
    osiem: 8,
    dziewiec: 9,
    dziesiec: 10,
    jedenascie: 11,
    dwanascie: 12,
    trzynascie: 13,
    czternascie: 14,
    pietnascie: 15,
    szesnascie: 16,
    siedemnascie: 17,
    osiemnascie: 18,
    dziewietnascie: 19,
    dwadziescia: 20,
    trzydziesci: 30,
    czterdziesci: 40,
    piecdziesiat: 50,
    szescdziesiat: 60,
    siedemdziesiat: 70,
    osiemdziesiat: 80,
    dziewiecdziesiat: 90,
    pol: 0.5,
    poltora: 1.5
  };

  if (
    Object.prototype.hasOwnProperty.call(
      words,
      normalized
    )
  ) {
    return words[normalized];
  }

  const parts = normalized.split(' ');

  if (parts.length === 2) {
    const tens = words[parts[0]];
    const units = words[parts[1]];

    if (
      Number.isInteger(tens) &&
      tens >= 20 &&
      tens <= 90 &&
      tens % 10 === 0 &&
      Number.isInteger(units) &&
      units >= 1 &&
      units <= 9
    ) {
      return tens + units;
    }
  }

  return null;
}

function normalizeNumberText_(value) {
  const polishMap = {
    'ą': 'a',
    'ć': 'c',
    'ę': 'e',
    'ł': 'l',
    'ń': 'n',
    'ó': 'o',
    'ś': 's',
    'ż': 'z',
    'ź': 'z'
  };

  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[ąćęłńóśżź]/g, char =>
      polishMap[char] || char
    )
    .replace(/,/g, '.')
    .replace(/^[^\d\w-]+/, '')
    .replace(/[.;:!?]+$/, '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function normalizeWordForParser_(value) {
  return normalizeNumberText_(value)
    .replace(/\./g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function createParserResult_(
  originalInput,
  product,
  value,
  status,
  message
) {
  return {
    originalInput: originalInput,
    product: product,
    value: value,
    location: '',
    status: status,
    message: message
  };
}


/**
 * Inventory PRO 2.10.5 — wbudowany matcher parsera.
 * Trzymany w Parser.gs, aby częściowa aktualizacja projektu nie mogła
 * pozostawić parsera bez wymaganej funkcji globalnej.
 */
function matchProductForParser_(recognitionInput, context) {
  const variants = buildParserRecognitionVariants_(recognitionInput);
  let best = null;

  variants.forEach((variant, index) => {
    const match = matchProduct(variant, context);
    if (!isParserMatchAnchoredAtStart_(variant, match)) return;

    const score = parserMatchStrength_(match) - index * 0.01;
    if (!best || score > best.score ||
        (score === best.score && variant.length > best.recognitionInput.length)) {
      best = { recognitionInput: variant, match: match, score: score };
    }
  });

  if (best) return best;
  const fallbackInput = variants[0] || normalizeRecognitionInput_(recognitionInput);
  const fallbackMatch = matchProduct(fallbackInput, context);
  if (!isParserMatchAnchoredAtStart_(fallbackInput, fallbackMatch)) {
    const candidates = [];
    if (fallbackMatch && fallbackMatch.product) {
      candidates.push({ product: fallbackMatch.product, score: fallbackMatch.score || 0 });
    }
    (fallbackMatch && fallbackMatch.candidates || []).forEach(candidate => {
      if (!candidate || !candidate.product) return;
      if (!candidates.some(item => normalizeText(item.product.name) === normalizeText(candidate.product.name))) {
        candidates.push(candidate);
      }
    });
    return {
      recognitionInput: fallbackInput,
      match: recognitionResult_(
        false,
        candidates.length ? 'AMBIGUOUS' : 'NOT_FOUND',
        fallbackInput,
        null,
        candidates,
        fallbackMatch && fallbackMatch.score || 0,
        'Nazwa zawiera nierozpoznany początek — wybierz produkt ręcznie'
      ),
      score: -1,
      safetyBlocked: true
    };
  }
  return {
    recognitionInput: fallbackInput,
    match: fallbackMatch,
    score: -1
  };
}
