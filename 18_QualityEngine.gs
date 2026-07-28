/**
 * Inventory PRO Enterprise v2.4
 * Ostrzezenia sterowane z arkusza Ustawienia.
 */

function evaluateImportQuality_(product, value, location, settings) {
  const qualitySettings = settings || loadQualitySettings_();

  if (!product || !Number.isFinite(Number(value))) {
    return {
      level: 'ERROR',
      warning: true,
      blocking: true,
      flags: ['NIEPRAWIDLOWA_WARTOSC'],
      message: 'Nieprawidlowa wartosc'
    };
  }

  const numericValue = Number(value);
  const isWhole = Number.isInteger(numericValue);
  const type = String(product.type || '').toUpperCase();
  const flags = [];
  const messages = [];
  let blocking = false;
  let threshold = null;
  let label = '';

  if (numericValue < 0) {
    flags.push('WARTOSC_UJEMNA');
    messages.push('Wartosc ujemna jest niedozwolona');
    blocking = Boolean(qualitySettings.blockNegative);
  }

  if (numericValue === 0 && qualitySettings.warnZero) {
    flags.push('WPROWADZONO_ZERO');
    messages.push('Wprowadzono 0');
  }

  if (type === CONFIG.PRODUCT_TYPES.LOCATION) {
    threshold = qualitySettings.locationWarning;
    label = 'Za duza liczba sztuk dla lokalizacji';
  } else if (type === CONFIG.PRODUCT_TYPES.KEG) {
    threshold = isWhole
      ? qualitySettings.kegWholeWarning
      : qualitySettings.kegWeightWarning;
    label = isWhole
      ? 'Za duza liczba pelnych kegow'
      : 'Za duza waga kega';
  } else {
    threshold = isWhole
      ? qualitySettings.normalWholeWarning
      : qualitySettings.normalWeightWarning;
    label = isWhole
      ? 'Za duza liczba sztuk'
      : 'Za duza wartosc wagi';
  }

  if (numericValue > threshold) {
    flags.push('PRZEKROCZONY_PROG');
    messages.push(label + ': ' + numericValue + ' (prog: ' + threshold + ')');
  }

  return {
    level: blocking ? 'ERROR' : (flags.length ? 'WARNING' : 'OK'),
    warning: flags.length > 0,
    blocking: blocking,
    threshold: threshold,
    flags: flags,
    message: messages.join(' | ')
  };
}
