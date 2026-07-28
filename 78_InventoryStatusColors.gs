/**
 * Czytelny podgląd stanu przed raportem końcowym.
 * Kolorowana jest wyłącznie komórka z nazwą produktu (kolumna A).
 */
function refreshInventoryStatusColors_(sheet) {
  const inventory = sheet || getSheetByConfiguredName_(CONFIG.SHEETS.INVENTORY);
  if (!inventory) return {colored: 0, errors: 0};
  SpreadsheetApp.flush();

  const products = scanInventoryProducts_();
  const lastRow = Math.max(inventory.getLastRow(), 1);
  const lastColumn = Math.max(inventory.getLastColumn(), getInventoryLayoutMaxColumn_());
  const displayValues = inventory.getRange(1, 1, lastRow, lastColumn).getDisplayValues();
  const colors = {positive: [], zero: [], negative: [], error: []};
  const all = [];

  products.forEach(product => {
    const row = Number(product.inventoryRow);
    if (!row) return;
    const layout = getConfiguredInventoryLayout_(product.type) || {};
    const finalColumn = getDirectFinalInventoryColumn_(product) || layout.finalTotal;
    if (!finalColumn) return;
    const nameCell = 'A' + row;
    const columnNumber = columnLetterToNumber290_(finalColumn);
    const display = String((displayValues[row - 1] || [])[columnNumber - 1] || '').trim();
    all.push(nameCell);
    if (/^#/.test(display)) {
      colors.error.push(nameCell);
      return;
    }
    if (display === '') return;
    const value = Number(display.replace(/\s/g, '').replace(',', '.'));
    if (!Number.isFinite(value)) {
      colors.error.push(nameCell);
    } else if (value > 0) {
      colors.positive.push(nameCell);
    } else if (value < 0) {
      colors.negative.push(nameCell);
    } else {
      colors.zero.push(nameCell);
    }
  });

  if (all.length) inventory.getRangeList(all).setBackground(null).setFontColor(null);
  if (colors.positive.length) inventory.getRangeList(colors.positive).setBackground('#d9ead3').setFontColor('#274e13');
  if (colors.zero.length) inventory.getRangeList(colors.zero).setBackground('#ff1744').setFontColor('#ffffff');
  if (colors.negative.length) inventory.getRangeList(colors.negative).setBackground('#fce5cd').setFontColor('#b45f06');
  if (colors.error.length) inventory.getRangeList(colors.error).setBackground('#d9d2e9').setFontColor('#674ea7');
  return {
    colored: colors.positive.length + colors.zero.length + colors.negative.length + colors.error.length,
    errors: colors.error.length
  };
}

function refreshInventoryStatusColors() {
  return refreshInventoryStatusColors_(getSheetByConfiguredName_(CONFIG.SHEETS.INVENTORY));
}
