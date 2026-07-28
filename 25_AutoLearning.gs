/**
 * Inventory PRO Enterprise v2.3.0
 * Bezpieczne zapamietywanie wyborow uzytkownika.
 */
function collectAliasSuggestions_(items) {
  const suggestions = [];
  (items || []).forEach(item => {
    if (!item.include || !item.learnAlias || !item.parsedProduct || !item.selectedProduct) return;
    const alias = normalizeAliasCandidate_(item.aliasSource || item.parsedProduct || item.originalInput);
    const product = String(item.selectedProduct).trim();
    if (alias && product && normalizeText(alias) !== normalizeText(product)) {
      suggestions.push({ alias: alias, product: product });
    }
  });
  return suggestions;
}

function normalizeAliasCandidate_(value) {
  const prepared = prepareParserText_(value || '');
  const tokens = prepared.split(/\s+/).filter(Boolean);
  const location = readLocationAt_(tokens, 0);
  const withoutLocation = location ? tokens.slice(location.consumed) : tokens;

  return withoutLocation.join(' ')
    .replace(/\s+(?:i|oraz|potem|dalej|nast[eę]pnie)\s*$/i, '')
    .replace(/\s+[-+]?\d+(?:[.,]\d+)?\s*$/i, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function validateAliasSuggestion_(suggestion, runtimeContext, existingAliases) {
  const alias = String(suggestion.alias || '').trim();
  const productName = String(suggestion.product || '').trim();
  const normalizedAlias = normalizeText(alias);
  const normalizedProduct = normalizeText(productName);
  const context = runtimeContext || buildRuntimeContext_();
  const aliases = existingAliases || loadAliases();

  if (!normalizedAlias || !normalizedProduct) return { valid: false, reason: 'Pusty alias lub produkt' };
  if (normalizedAlias.length < 3 || /^\d+(?:\s+\d+)*$/.test(normalizedAlias)) {
    return { valid: false, reason: 'Alias jest zbyt krotki lub sklada sie tylko z liczb' };
  }
  if (normalizedAlias === normalizedProduct) return { valid: false, reason: 'Alias jest identyczny z nazwa produktu' };
  if (!context.productIndex[normalizedProduct]) return { valid: false, reason: 'Produkt docelowy nie istnieje w katalogu' };

  if (context.productIndex[normalizedAlias] && normalizedAlias !== normalizedProduct) {
    return { valid: false, reason: 'Alias jest pelna nazwa innego produktu' };
  }

  const familyGuard = validateCapacityFamilyAlias_(
    alias,
    productName,
    context
  );
  if (!familyGuard.valid) return familyGuard;

  if (aliases[normalizedAlias]) {
    return normalizeText(aliases[normalizedAlias]) === normalizedProduct
      ? { valid: false, reason: 'Alias juz istnieje' }
      : { valid: false, reason: 'Konflikt: alias wskazuje inny produkt' };
  }

  const familyMatches = context.catalog.filter(product =>
    product.normalizedName === normalizedAlias ||
    product.normalizedName.startsWith(normalizedAlias + ' ')
  );
  if (familyMatches.length > 1) {
    return { valid: false, reason: 'Alias jest zbyt ogolny i pasuje do kilku produktow' };
  }

  const targetNumbers = (normalizedProduct.match(/\d+/g) || []);
  const aliasNumbers = (normalizedAlias.match(/\d+/g) || []);
  const numericFamily = context.catalog.filter(product => {
    const name = normalizeText(product.name);
    const base = name.replace(/\d+/g, '').replace(/\s+/g, ' ').trim();
    const targetBase = normalizedProduct.replace(/\d+/g, '').replace(/\s+/g, ' ').trim();
    return base === targetBase;
  });
  if (!familyGuard.capacityQualified &&
      targetNumbers.length && numericFamily.length > 1 &&
      !targetNumbers.every(number => aliasNumbers.includes(number))) {
    return { valid: false, reason: 'Alias pomija liczbe odrozniajaca wariant produktu' };
  }

  return { valid: true, alias: alias, product: productName };
}

/**
 * Dwustopniowe uczenie aliasow:
 * - alias rodzinny bez pojemnosci pozostaje wyborem,
 * - alias z pojemnoscia moze wskazywac tylko odpowiadajacy jej wariant.
 */
function validateCapacityFamilyAlias_(alias, productName, runtimeContext) {
  const context = runtimeContext || buildRuntimeContext_();
  const target = extractCapacityDescriptor_(productName);
  if (!target) return { valid: true };

  const family = (context.catalog || []).filter(product => {
    const descriptor = extractCapacityDescriptor_(product.name);
    return descriptor &&
      descriptor.familyBase === target.familyBase &&
      descriptor.capacityKey !== target.capacityKey;
  });
  if (!family.length) return { valid: true };

  const aliasCapacity = extractSpokenCapacityKey_(alias);
  if (!aliasCapacity) {
    return {
      valid: false,
      protectedFamily: true,
      reason: 'Alias rodzinny bez pojemnosci pozostaje wyborem i nie zostal przypisany na stale.'
    };
  }
  if (aliasCapacity !== target.capacityKey) {
    return {
      valid: false,
      protectedFamily: true,
      reason: 'Pojemnosc w aliasie nie odpowiada wybranemu wariantowi produktu.'
    };
  }
  return { valid: true, capacityQualified: true };
}

function extractCapacityDescriptor_(value) {
  const text = normalizeText(value);
  const match = text.match(/(?:^|\s)(\d+)(?:\s+(\d+))?\s*(ml|cl|l|kg|g)(?:\s|$)/);
  if (!match) return null;
  const whole = match[0].trim();
  const integer = Number(match[1]);
  const fraction = match[2] || '';
  const amount = Number(fraction ? integer + '.' + fraction : integer);
  const unit = match[3];
  const canonical = unit === 'ml' ? ['v', amount / 1000] :
    unit === 'cl' ? ['v', amount / 100] :
    unit === 'l' ? ['v', amount] :
    unit === 'g' ? ['m', amount / 1000] : ['m', amount];
  return {
    capacityKey: canonical[0] + canonical[1],
    familyBase: text.replace(whole, ' ').replace(/\s+/g, ' ').trim()
  };
}

function extractSpokenCapacityKey_(value) {
  const text = normalizeText(value);
  const technical = extractCapacityDescriptor_(value);
  if (technical) return technical.capacityKey;
  if (/\b0\s+7$/.test(text)) return 'v0.7';
  if (/\b0\s+5$/.test(text)) return 'v0.5';
  if (/\b0\s+75$/.test(text)) return 'v0.75';
  if (/\b(?:zero\s+(?:przecinek\s+|kropka\s+)?siedem)\b/.test(text)) return 'v0.7';
  if (/\b(?:zero\s+(?:przecinek\s+|kropka\s+)?piec)\b/.test(text)) return 'v0.5';
  if (/\bpol(?:\s+litra)?\b/.test(text)) return 'v0.5';
  if (/\b(?:jeden\s+)?litr\b/.test(text)) return 'v1';
  return '';
}

function countProtectedAliasSuggestions_(suggestions, runtimeContext) {
  const context = runtimeContext || buildRuntimeContext_();
  return (suggestions || []).filter(suggestion =>
    validateCapacityFamilyAlias_(
      suggestion.alias,
      suggestion.product,
      context
    ).protectedFamily === true
  ).length;
}
