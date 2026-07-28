/**
 * Inventory PRO Enterprise v2.2
 * Jeden kontekst na cala operacje.
 */

let INVENTORY_RUNTIME_SNAPSHOT_ = null;

function buildRuntimeContext_() {
  const startedAt = Date.now();
  const snapshot = getRuntimeSnapshot_();
  return {
    catalog: snapshot.catalog,
    productIndex: snapshot.productIndex,
    aliasIndex: snapshot.aliasIndex,
    exactRecognitionIndex: snapshot.exactRecognitionIndex,
    technicalRecognitionIndex: snapshot.technicalRecognitionIndex,
    firstTokenIndex: snapshot.firstTokenIndex,
    trigramRecognitionIndex: snapshot.trigramRecognitionIndex,
    recognitionSources: snapshot.recognitionSources,
    recognitionSourceTrigrams: snapshot.recognitionSourceTrigrams,
    parserPhraseIndex: snapshot.parserPhraseIndex,
    parserPhraseTrie: snapshot.parserPhraseTrie,
    parserPhraseTrieSourceIndex: snapshot.parserPhraseIndex,
    recognitionMemo: {},
    performanceStats: {
      contextBuildMs: Date.now() - startedAt,
      snapshotReused: snapshot.reused,
      memoHits: 0,
      fuzzyQueries: 0,
      fuzzyShortlisted: 0,
      levenshteinCandidates: 0
    }
  };
}

function getRuntimeSnapshot_() {
  if (INVENTORY_RUNTIME_SNAPSHOT_) {
    INVENTORY_RUNTIME_SNAPSHOT_.reused = true;
    return INVENTORY_RUNTIME_SNAPSHOT_;
  }

  const catalog = buildProductCatalog();
  const parserContext = { catalog: catalog };
  const parserPhraseIndex = getParserPhraseIndex_(parserContext);
  INVENTORY_RUNTIME_SNAPSHOT_ = {
    catalog: catalog,
    productIndex: buildProductCatalogIndex(catalog),
    aliasIndex: buildAliasProductIndex(catalog),
    exactRecognitionIndex: buildExactRecognitionIndex_(catalog),
    technicalRecognitionIndex: buildTechnicalRecognitionIndex_(catalog),
    firstTokenIndex: buildFirstTokenRecognitionIndex_(catalog),
    trigramRecognitionIndex: buildTrigramRecognitionIndex_(catalog),
    recognitionSources: buildRecognitionSources_(catalog),
    recognitionSourceTrigrams: buildRecognitionSourceTrigrams_(catalog),
    parserPhraseIndex: parserPhraseIndex,
    parserPhraseTrie: buildParserPhraseTrieFromIndex_(parserPhraseIndex),
    reused: false
  };
  return INVENTORY_RUNTIME_SNAPSHOT_;
}

function invalidateRuntimeSnapshot_() {
  INVENTORY_RUNTIME_SNAPSHOT_ = null;
}


function buildExactRecognitionIndex_(catalog) {
  const index = {};
  (catalog || []).forEach(product => {
    [product.name].concat(product.aliases || []).forEach(name => {
      const key = normalizeRecognitionInput_(name);
      if (key && !index[key]) index[key] = product;
    });
  });
  return index;
}


function buildTechnicalRecognitionIndex_(catalog) {
  const index = {};
  (catalog || []).forEach(product => {
    [product.name].concat(product.aliases || []).forEach(value => {
      buildTechnicalRecognitionKeys_(value).forEach(key => {
        if (!key) return;
        if (!index[key]) index[key] = [];
        if (!index[key].some(item => normalizeText(item.name) === normalizeText(product.name))) {
          index[key].push(product);
        }
      });
    });
  });
  return index;
}

function buildFirstTokenRecognitionIndex_(catalog) {
  const index = {};
  (catalog || []).forEach(product => {
    [product.name].concat(product.aliases || []).forEach(value => {
      const normalized = normalizeRecognitionForScore_(value);
      const first = normalized.split(' ').filter(Boolean)[0] || '';
      if (!first) return;
      if (!index[first]) index[first] = [];
      if (!index[first].some(item => normalizeText(item.name) === normalizeText(product.name))) {
        index[first].push(product);
      }
    });
  });
  return index;
}

function buildRecognitionSources_(catalog) {
  const sources = {};
  (catalog || []).forEach(product => {
    sources[normalizeText(product.name)] = [product.name]
      .concat(product.aliases || [])
      .map(normalizeRecognitionForScore_)
      .filter((value, index, all) => value && all.indexOf(value) === index);
  });
  return sources;
}

function buildRecognitionSourceTrigrams_(catalog) {
  const result = {};
  const sources = buildRecognitionSources_(catalog);
  Object.keys(sources).forEach(productKey => {
    result[productKey] = sources[productKey].map(value => ({
      value: value,
      grams: recognitionTrigrams_(value),
      numbers: (value.match(/\d+(?:[.,]\d+)?/g) || [])
    }));
  });
  return result;
}

function buildTrigramRecognitionIndex_(catalog) {
  const index = {};
  (catalog || []).forEach(product => {
    const productKey = normalizeText(product.name);
    const seen = {};
    [product.name].concat(product.aliases || []).forEach(value => {
      recognitionTrigrams_(normalizeRecognitionForScore_(value)).forEach(gram => {
        if (seen[gram]) return;
        seen[gram] = true;
        if (!index[gram]) index[gram] = [];
        index[gram].push(product);
      });
    });
  });
  return index;
}

function recognitionTrigrams_(value) {
  const compact = '  ' + String(value || '').replace(/\s+/g, ' ').trim() + '  ';
  if (!compact.trim()) return [];
  const result = [];
  const seen = {};
  for (let index = 0; index <= compact.length - 3; index++) {
    const gram = compact.slice(index, index + 3);
    if (!seen[gram]) {
      seen[gram] = true;
      result.push(gram);
    }
  }
  return result;
}