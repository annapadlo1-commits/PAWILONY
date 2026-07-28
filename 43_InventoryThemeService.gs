/**
 * Inventory PRO Enterprise v2.8.2
 * Bezpieczny motyw wizualny arkusza INWENTURA.
 * Nie modyfikuje wartości, formuł ani układu danych.
 */

function applyInventoryTheme() {
  const sheet = getSheetByConfiguredName_(CONFIG.SHEETS.INVENTORY);
  if (!sheet) throw new Error('Nie znaleziono arkusza INWENTURA.');

  const range = sheet.getDataRange();
  range.setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle');
  sheet.setFrozenRows(1);
  sheet.setRowHeight(1, 34);

  const values = range.getDisplayValues();
  const lastCol = Math.max(1, sheet.getLastColumn());
  const categorySet = new Set(getBusinessCategories().map(normalizeText));

  for (let r = 0; r < values.length; r++) {
    const firstCell = String((values[r] || [])[0] || '').trim();
    const normalized = normalizeText(firstCell);
    const businessCategory = normalizeBusinessCategory_(firstCell);
    const isCategory = Boolean(businessCategory) && categorySet.has(normalizeText(businessCategory)) && normalized === normalizeText(businessCategory);
    if (isCategory) {
      sheet.getRange(r + 1, 1, 1, lastCol)
        .setBackground('#FDE7C2')
        .setFontWeight('bold')
        .setFontSize(12)
        .setHorizontalAlignment('center')
        .setBorder(false, false, true, false, false, false, '#D84315', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      sheet.setRowHeight(r + 1, 28);
    }
  }

  sheet.getRange(1, 1, 1, lastCol)
    .setBackground('#FCE8E6')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setWrap(true)
    .setBorder(true, true, true, true, true, true, '#DADCE0', SpreadsheetApp.BorderStyle.SOLID);

  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 1)
      .setHorizontalAlignment('left')
      .setWrap(true);
    if (lastCol > 1) {
      sheet.getRange(2, 2, sheet.getLastRow() - 1, lastCol - 1)
        .setHorizontalAlignment('right');
    }
  }

  sheet.setColumnWidth(1, 280);
  for (let c = 2; c <= Math.min(lastCol, 12); c++) sheet.setColumnWidth(c, 120);
  sheet.setHiddenGridlines(true);
  SpreadsheetApp.flush();
  return { success: true, sheet: sheet.getName(), rows: sheet.getLastRow(), columns: lastCol };
}

function previewInventoryTheme() {
  return applyInventoryTheme();
}

function backupInventoryFormatting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = getSheetByConfiguredName_(CONFIG.SHEETS.INVENTORY);
  if (!source) throw new Error('Nie znaleziono arkusza INWENTURA.');
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH-mm-ss');
  const copy = source.copyTo(ss).setName('BACKUP STYLE ' + stamp);
  copy.hideSheet();
  return { success: true, backupSheetName: copy.getName() };
}