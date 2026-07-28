/**
 * Inventory PRO Enterprise v2.9.3 — Excel PRO & PDF Manager Report.
 * Moduł tworzy raport wyłącznie na podstawie danych Reporting Engine.
 * Nie modyfikuje arkusza INWENTURA ani jego formuł.
 */

function buildExportManagerReportSheet_(sheet, payload) {
  const summary = payload.snapshot.summary || {};
  const statistics = payload.snapshot.statistics || {};
  const categories = Object.keys(statistics.byCategory || {}).sort((a, b) => a.localeCompare(b, 'pl'));

  sheet.clear();
  sheet.setHiddenGridlines(true);
  sheet.setColumnWidths(1, 6, 120);
  sheet.setColumnWidth(1, 190);
  sheet.setColumnWidth(2, 170);
  sheet.setColumnWidth(3, 170);
  sheet.setColumnWidth(4, 170);
  sheet.setColumnWidth(5, 170);
  sheet.setColumnWidth(6, 190);

  sheet.getRange('A1:F2').merge()
    .setValue('🍕 INVENTORY PRO — RAPORT KOŃCOWY')
    .setBackground('#d84315').setFontColor('#ffffff').setFontWeight('bold')
    .setFontSize(20).setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeights(1, 2, 30);

  const metadata = [
    ['Numer eksportu', payload.exportId],
    ['Arkusz źródłowy', payload.sourceSheetName || CONFIG.SHEETS.INVENTORY],
    ['Wygenerowano', new Date()],
    ['Użytkownik', getCurrentUserEmail_()],
    ['Wersja', CONFIG.VERSION],
    ['Notatka', payload.note || '—']
  ];
  sheet.getRange(4, 1, metadata.length, 2).setValues(metadata);
  sheet.getRange(4, 1, metadata.length, 1).setFontWeight('bold').setBackground('#fff2cc');
  sheet.getRange(4, 2, metadata.length, 1).setWrap(true);

  const cards = [
    ['Produkty', summary.products || 0],
    ['Uzupełnione', summary.completed || 0],
    ['Braki', summary.missing || 0],
    ['Ostrzeżenia', summary.warningProducts || 0]
  ];
  cards.forEach((card, i) => {
    const col = 3 + i;
    sheet.getRange(4, col).setValue(card[0]).setBackground('#fff2cc').setFontWeight('bold').setHorizontalAlignment('center');
    sheet.getRange(5, col).setValue(card[1]).setFontSize(20).setFontWeight('bold').setHorizontalAlignment('center');
    sheet.getRange(4, col, 2, 1).setBorder(true, true, true, true, false, false, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);
  });

  const completeness = summary.products ? Math.round((summary.completed || 0) * 1000 / summary.products) / 10 : 0;
  sheet.getRange('C7:F7').merge().setValue('Kompletność danych: ' + completeness + '%')
    .setBackground(completeness >= 95 ? '#d9ead3' : completeness >= 80 ? '#fff2cc' : '#f4cccc')
    .setFontWeight('bold').setHorizontalAlignment('center');

  let row = 12;
  sheet.getRange(row, 1, 1, 6).merge().setValue('PODSUMOWANIE WEDŁUG KATEGORII')
    .setBackground('#f6b26b').setFontWeight('bold').setHorizontalAlignment('center');
  row++;
  const headers = ['Kategoria', 'Produkty', 'Uzupełnione', 'Braki', 'Stan finalny', 'Jednostka'];
  sheet.getRange(row, 1, 1, headers.length).setValues([headers]).setBackground('#fff2cc').setFontWeight('bold');
  row++;
  const categoryRows = categories.map(category => {
    const c = statistics.byCategory[category];
    return [category, c.products, c.completed, c.missing, c.finalTotal, c.unit || ''];
  });
  if (categoryRows.length) sheet.getRange(row, 1, categoryRows.length, headers.length).setValues(categoryRows);
  row += categoryRows.length + 2;

  sheet.getRange(row, 1, 1, 6).merge().setValue('NAJWAŻNIEJSZE OSTRZEŻENIA')
    .setBackground('#f4cccc').setFontWeight('bold').setHorizontalAlignment('center');
  row++;
  const warningHeaders = ['Produkt', 'Kategoria', 'Typ', 'Stan końcowy', 'Jednostka', 'Flagi'];
  sheet.getRange(row, 1, 1, 6).setValues([warningHeaders]).setBackground('#fce8e6').setFontWeight('bold');
  row++;
  const warningRows = (payload.snapshot.warnings || []).slice(0, 20).map(w => [w.product, w.category, w.type, w.finalTotal, w.unit, (w.flags || []).join(', ')]);
  if (warningRows.length) sheet.getRange(row, 1, warningRows.length, 6).setValues(warningRows);
  else sheet.getRange(row, 1, 1, 6).merge().setValue('Brak ostrzeżeń.').setHorizontalAlignment('center');

  const lastRow = Math.max(sheet.getLastRow(), row);
  sheet.getRange(1, 1, lastRow, 6).setFontFamily('Arial').setVerticalAlignment('middle');
  sheet.getRange(4, 1, lastRow - 3, 6).setBorder(true, true, true, true, true, true, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);
  sheet.setFrozenRows(2);
}

function buildExportStatisticsSheet_(sheet, payload) {
  const stats = payload.snapshot.statistics || {};
  const headers = ['Kategoria', 'Produkty', 'Uzupełnione', 'Braki', 'Kompletność %', 'Stan finalny', 'Jednostka'];
  const rows = Object.keys(stats.byCategory || {}).sort((a, b) => a.localeCompare(b, 'pl')).map(category => {
    const c = stats.byCategory[category];
    return [category, c.products, c.completed, c.missing, c.products ? Math.round(c.completed * 1000 / c.products) / 10 : 0, c.finalTotal, c.unit || ''];
  });
  writeExportTable_(sheet, headers, rows, '#f6b26b');
}

function buildExportMetadataSheet_(sheet, payload) {
  const rows = [
    ['Pole', 'Wartość'],
    ['Export ID', payload.exportId],
    ['Arkusz źródłowy', payload.sourceSheetName || CONFIG.SHEETS.INVENTORY],
    ['Session ID', payload.session.id],
    ['Rozpoczęto', new Date(payload.session.startedAt)],
    ['Wygenerowano', new Date()],
    ['Użytkownik', getCurrentUserEmail_()],
    ['Wersja aplikacji', CONFIG.VERSION],
    ['Liczba eksportowanych produktów', payload.items.length],
    ['Wykluczone', payload.excludedCount],
    ['Notatka', payload.note || '']
  ];
  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
  sheet.getRange(1, 1, 1, 2).setBackground('#f6b26b').setFontWeight('bold');
  sheet.getRange(2, 1, rows.length - 1, 1).setBackground('#fff2cc').setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 230);
  sheet.setColumnWidth(2, 500);
  sheet.getRange(1, 1, rows.length, 2).setWrap(true).setFontFamily('Arial');
}

function exportManagerReportAsPdf_(spreadsheet, exportId) {
  const reportSheet = spreadsheet.getSheetByName('PODSUMOWANIE');
  if (!reportSheet) throw new Error('Nie znaleziono arkusza PODSUMOWANIE do eksportu PDF.');

  SpreadsheetApp.flush();
  Utilities.sleep(700);
  const params = [
    'format=pdf', 'size=A4', 'portrait=true', 'fitw=true', 'sheetnames=false',
    'printtitle=false', 'pagenumbers=true', 'gridlines=false', 'fzr=false',
    'top_margin=0.35', 'bottom_margin=0.35', 'left_margin=0.35', 'right_margin=0.35',
    'gid=' + reportSheet.getSheetId()
  ].join('&');
  const url = 'https://docs.google.com/spreadsheets/d/' + spreadsheet.getId() + '/export?' + params;
  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  if (response.getResponseCode() !== 200) {
    throw new Error('Nie udało się wygenerować PDF. Kod: ' + response.getResponseCode() + '. ' + response.getContentText().slice(0, 250));
  }
  return getExportFolder_().createFile(response.getBlob().setName(exportId + '_RAPORT.pdf'));
}