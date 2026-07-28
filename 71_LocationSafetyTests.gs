/** Testy 4.3.7: lokalizacje, kontekst i pierwszeństwo lokalizacji bez dostępu do arkusza. */
function runLocationSafetyTests437() {
  const context = createParserTestContext_([
    'Fritz Kola', 'Bombilla', 'Litovel Ciemny Lager', 'Inne Beczki Pilsner'
  ]);
  const cases = [
    { input: 'magazyn Fritz Kola dark rum 4', product: 'Fritz Kola', value: 4, location: 'darkroom', category: 'SOFTY' },
    { input: 'lodówki Bombilla 7', product: 'Bombilla', value: 7, location: 'lodowki', category: 'SOFTY' },
    { input: 'magazyn Litovel Ciemny Lager darkroom 12', product: 'Litovel Ciemny Lager', value: 12, location: 'darkroom', category: 'PIWO BUTELKI' },
    { input: 'dark rum Inne Beczki Pilsner 2', product: 'Inne Beczki Pilsner', value: 2, location: 'darkroom', category: 'PIWO BUTELKI' }
  ];
  cases.forEach(testCase => {
    const parsed = parseInventoryText(testCase.input, context);
    assertCondition_(parsed.length === 1, testCase.category + ': oczekiwano jednej pozycji.');
    assertCondition_(parsed[0].product === testCase.product, testCase.category + ': błędny produkt.');
    assertCondition_(parsed[0].value === testCase.value, testCase.category + ': błędna ilość.');
    assertCondition_(parsed[0].location === testCase.location, testCase.category + ': błędna lokalizacja.');
  });
  const multiline = parseInventoryText(
    'dark rum Fritz Kola 1\nlodówki Fritz Kola 1\ndark rum Fritz Kola 1\nJurajska Pomarancza 12\nBombilla 1',
    createParserTestContext_(['Fritz Kola', 'Jurajska Pomarancza', 'Bombilla'])
  );
  assertCondition_(multiline.length === 5, 'Wielowierszowy wpis powinien zwrócić pięć pozycji.');
  assertCondition_(multiline[0].location === 'darkroom' && multiline[1].location === 'lodowki',
    'Jawne lokalizacje w pierwszych wierszach są błędne.');
  assertCondition_(multiline[2].location === 'darkroom' && multiline[3].location === 'darkroom' && multiline[4].location === 'darkroom',
    'Kontekst dark rum nie przeszedł na kolejne wiersze.');
  return { passed: true, location: CONFIG.LOCATION.ID, cases: cases.length + 1 };
}