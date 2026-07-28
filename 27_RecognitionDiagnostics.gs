/** Inventory PRO v2.3 - szybka diagnostyka silnika rozpoznawania. */
function debugRecognitionEngine() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Diagnostyka rozpoznawania', 'Wpisz sama nazwe produktu (bez ilosci):', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const result = matchProduct(response.getResponseText(), buildRuntimeContext_());
  const candidates = (result.candidates || []).slice(0, 8).map(c => c.product.name + ' (' + Math.round(c.score) + '%)').join('\n');
  ui.alert('Wynik: ' + result.status, 'Pewnosc: ' + (result.score || 0) + '%\nProdukt: ' + (result.product ? result.product.name : '-') + '\n\nKandydaci:\n' + (candidates || '-'), ui.ButtonSet.OK);
}

function getParserDiagnostics(inputText) {
  const source = String(inputText || '').trim();
  if (!source) throw new Error('Wklej tekst do diagnostyki.');
  const startedAt = Date.now();
  const context = buildRuntimeContext_();
  const prepared = prepareParserText_(source);
  const tokens = prepared.split(/\s+/).filter(Boolean);
  const tokenDetails = [];
  for (let position = 0; position < tokens.length;) {
    const location = readLocationAt_(tokens, position);
    if (location) {
      tokenDetails.push({ position: position, text: tokens.slice(position, position + location.consumed).join(' '), type: 'LOCATION', value: location.location });
      position += location.consumed;
      continue;
    }
    const number = readNumberAt_(tokens, position);
    if (number) {
      tokenDetails.push({ position: position, text: number.originalText, type: 'NUMBER', value: number.value });
      position += number.consumed;
      continue;
    }
    tokenDetails.push({ position: position, text: tokens[position], type: 'TEXT', value: '' });
    position++;
  }

  const parsedStartedAt = Date.now();
  const parsed = parseInventoryText(source, context);
  const parsedAt = Date.now();
  const items = parsed.map((item, index) => {
    const parserMatch = item.product ? matchProductForParser_(item.product, context) : null;
    const match = parserMatch && parserMatch.match;
    return {
      id: index + 1,
      originalInput: item.originalInput || '',
      parsedProduct: item.product || '',
      value: item.value,
      location: item.location || '',
      parserStatus: item.status,
      parserMessage: item.message || '',
      matchStatus: match ? match.status : 'BRAK',
      score: match ? Number(match.score || 0) : 0,
      selectedProduct: match && match.product ? match.product.name : '',
      candidates: match ? (match.candidates || []).slice(0, 5).map(candidate => ({
        name: candidate.product.name,
        score: Math.round(candidate.score || 0),
        numericConflict: Boolean(candidate.numericConflict),
        missingRequiredNumber: Boolean(candidate.missingRequiredNumber)
      })) : []
    };
  });

  return {
    version: CONFIG.VERSION,
    sourceLength: source.length,
    tokenCount: tokens.length,
    itemCount: items.length,
    preparedText: prepared,
    tokens: tokenDetails,
    items: items,
    performance: {
      contextMs: parsedStartedAt - startedAt,
      parserMs: parsedAt - parsedStartedAt,
      totalMs: Date.now() - startedAt,
      matcher: context.performanceStats || {}
    }
  };
}