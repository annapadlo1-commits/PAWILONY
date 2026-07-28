/**
 * Inventory PRO Enterprise v2.3.0
 * Publiczny interfejs dopasowania. Logika znajduje sie w RecognitionEngine.gs.
 */
function matchProduct(inputName, runtimeContext) {
  return recognizeProduct_(inputName, runtimeContext);
}

function calculateStringSimilarity_(textA, textB) {
  const a = String(textA || '');
  const b = String(textB || '');
  if (a === b) return 1;
  if (!a.length || !b.length) return 0;
  return 1 - (levenshteinDistance_(a, b) / Math.max(a.length, b.length));
}

function levenshteinDistance_(textA, textB, maxDistance) {
  let a = String(textA || '');
  let b = String(textB || '');
  if (a.length > b.length) {
    const swap = a; a = b; b = swap;
  }
  const limit = Number.isFinite(Number(maxDistance)) ? Number(maxDistance) : Infinity;
  if (b.length - a.length > limit) return limit + 1;

  let previous = new Array(a.length + 1);
  let current = new Array(a.length + 1);
  for (let column = 0; column <= a.length; column++) previous[column] = column;
  for (let row = 1; row <= b.length; row++) {
    current[0] = row;
    let rowMinimum = current[0];
    for (let column = 1; column <= a.length; column++) {
      const cost = b.charAt(row - 1) === a.charAt(column - 1) ? 0 : 1;
      current[column] = Math.min(current[column - 1] + 1, previous[column] + 1, previous[column - 1] + cost);
      rowMinimum = Math.min(rowMinimum, current[column]);
    }
    if (rowMinimum > limit) return limit + 1;
    const swap = previous; previous = current; current = swap;
  }
  return previous[a.length];
}