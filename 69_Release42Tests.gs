/** Inventory PRO 4.2 — testy nowych funkcji bez zależności od arkusza. */
function runInventory42Tests() {
  const context = createParserTestContext_(['Bacardi 4', 'Bacardi 8', 'Bacardi 10']);
  const merged = mergePreviewDuplicates_([
    release42Item_('Bacardi 8', 0.987, 'pierwszy'),
    release42Item_('Bacardi 8', 1.123, 'drugi')
  ]);
  assertCondition_(merged.length === 1, 'Duplikaty nie zostaly polaczone.');
  assertCondition_(merged[0].autoMerged && merged[0].sourceItems.length === 2,
    'Brak danych pozwalajacych rozdzielic polaczona grupe.');
  assertCondition_(Math.abs(merged[0].value - 2.11) < 0.000001,
    'Nieprawidlowa suma duplikatow: ' + merged[0].value);

  const blocked = validateAliasSuggestion_({ alias: 'bakardi', product: 'Bacardi 8' }, context, {});
  assertCondition_(!blocked.valid, 'Alias bez liczby wariantu powinien zostac zablokowany.');
  const allowed = validateAliasSuggestion_({ alias: 'bakardi 8', product: 'Bacardi 8' }, context, {});
  assertCondition_(allowed.valid, 'Bezpieczny alias z numerem wariantu powinien byc dozwolony.');

  const capacityContext = createParserTestContext_([
    'Fernet Branca 0,7L',
    'Fernet Branca 1L'
  ]);
  const familyAlias = validateAliasSuggestion_(
    { alias: 'fernet branka', product: 'Fernet Branca 0,7L' },
    capacityContext,
    {}
  );
  assertCondition_(
    !familyAlias.valid && familyAlias.protectedFamily,
    'Alias rodzinny bez pojemnosci musi pozostac wyborem.'
  );
  assertCondition_(
    validateAliasSuggestion_(
      { alias: 'fernet branka zero siedem', product: 'Fernet Branca 0,7L' },
      capacityContext,
      {}
    ).valid,
    'Alias slowny z pojemnoscia powinien zostac zapamietany.'
  );
  assertCondition_(
    validateAliasSuggestion_(
      { alias: 'fernet branka litr', product: 'Fernet Branca 1L' },
      capacityContext,
      {}
    ).valid,
    'Alias z jednostka litr powinien zostac zapamietany.'
  );

  const historyRows = [
    release42HistoryRow_('IMP-TEST', 10, 'B', '', 2),
    release42HistoryRow_('IMP-TEST', 10, 'B', 2, 5),
    release42HistoryRow_('IMP-TEST', 11, 'C', '', 4),
    release42HistoryRow_('IMP-TEST', 12, 'D', '', 6)
  ];
  const undoPlan = buildUndoPlan_(historyRows, 'IMP-TEST');
  assertCondition_(undoPlan.changes.length === 3,
    'Cofanie powinno objąć trzy komórki lokalizacji.');
  const warehouse = undoPlan.changes.filter(change => change.column === 'B' && change.row === 10)[0];
  assertCondition_(warehouse && warehouse.previousValue === '' && warehouse.auditRows.length === 2,
    'Duplikaty w magazynie muszą wrócić do pustej komórki jednym planem cofnięcia.');
  assertCondition_(undoPlan.auditRows.length === 4,
    'Wszystkie pozycje importu muszą zostać oznaczone jako cofnięte.');

  const editable = buildEditableReviewCells_({
    type: CONFIG.PRODUCT_TYPES.LOCATION,
    cells: { warehouse:'B10', darkroom:'C10', fridges:'D10', finalTotal:'E10' }
  });
  assertCondition_(editable.Magazyn === 'B10' && editable.Darkroom === 'C10' && editable['Lodówki'] === 'D10',
    'Końcowy przegląd nie zwraca poprawnych pól edytowalnych lokalizacji.');

  const sortedAliases = sortAliasRecordsForManager_([
    { alias:'zeta', product:'Produkt B', createdAt:null, originalIndex:1, values:[] },
    { alias:'alfa', product:'Produkt A', createdAt:null, originalIndex:0, values:[] }
  ], 'alias');
  assertCondition_(sortedAliases[0].alias === 'alfa',
    'Porządkowanie aliasów A–Z nie działa.');
  const chronologicalAliases = sortAliasRecordsForManager_([
    { alias:'nowy', product:'P', createdAt:new Date('2026-07-16T10:00:00Z'), originalIndex:0, values:[] },
    { alias:'stary', product:'P', createdAt:null, originalIndex:4, values:[] },
    { alias:'nowszy', product:'P', createdAt:new Date('2026-07-16T11:00:00Z'), originalIndex:1, values:[] }
  ], 'chronological');
  assertCondition_(chronologicalAliases.map(item => item.alias).join('|') === 'stary|nowy|nowszy',
    'Chronologiczne porządkowanie aliasów nie zachowuje kolejności dodania.');

  const report = { passed: true, mergedValue: merged[0].value, sourceCount: merged[0].sourceItems.length };
  SpreadsheetApp.getUi().alert(
    'Inventory PRO 4.2',
    'PASS\nŁączenie duplikatów: PASS\nOchrona aliasów numerycznych: PASS\nPełne cofanie lokalizacji: PASS\nEdycja raportu: PASS\nPorządkowanie aliasów: PASS',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  return report;
}

function release42HistoryRow_(importId, row, column, previousValue, newValue) {
  return [importId, new Date(), 'test', '', 'Produkt', newValue, 'magazyn',
    'Inwentura', row, column, previousValue, newValue, 'SAVED', '', '', '4.2.1'];
}

function release42Item_(product, value, source) {
  return {
    include: true,
    selectedProduct: product,
    parsedProduct: product,
    originalInput: source,
    value: value,
    status: 'EXACT',
    productType: 'NORMAL',
    category: 'TEST',
    location: '',
    qualityLevel: 'OK',
    qualityWarning: false
  };
}
