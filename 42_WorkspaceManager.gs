/**
 * Inventory PRO — kontrolowane widoki arkusza.
 * Domyślnie widoczna jest wyłącznie bieżąca inwentaryzacja.
 */

function showCurrentInventoryWorkspace() {
  setWorkspaceMode_('CURRENT');
  applyWorkspaceVisibility_('CURRENT');
  activateSheetByName_(CONFIG.SHEETS.INVENTORY);
  SpreadsheetApp.getActive().toast('Widok: inwentaryzacja bieżąca.', 'Inventory PRO', 4);
}

function showAllInventoriesWorkspace() {
  setWorkspaceMode_('ALL_INVENTORIES');
  applyWorkspaceVisibility_('ALL_INVENTORIES');
  activateSheetByName_(CONFIG.SHEETS.INVENTORY);
  SpreadsheetApp.getActive().toast('Widok: bieżąca i archiwalne inwentaryzacje.', 'Inventory PRO', 4);
}

function showAdministrationWorkspace() {
  setWorkspaceMode_('ADMIN');
  applyWorkspaceVisibility_('ADMIN');
  SpreadsheetApp.getActive().toast('Widok administracja: pokazano wszystkie zakładki.', 'Inventory PRO', 4);
}

// Zgodność ze starszym menu i zapisanymi skrótami.
function showUserWorkspace() { showCurrentInventoryWorkspace(); }
function showManagerWorkspace() { showAdministrationWorkspace(); }

function applySavedWorkspaceMode() {
  const mode = getWorkspaceMode_();
  applyWorkspaceVisibility_(mode);
  return mode;
}

function getWorkspaceMode_() {
  const stored = PropertiesService.getUserProperties().getProperty('INVENTORY_PRO_WORKSPACE_MODE');
  if (stored === 'ALL_INVENTORIES' || stored === 'ADMIN') return stored;
  return 'CURRENT';
}

function setWorkspaceMode_(mode) {
  PropertiesService.getUserProperties().setProperty('INVENTORY_PRO_WORKSPACE_MODE', mode);
}

function applyWorkspaceVisibility_(mode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inventory = getSheetByConfiguredName_(CONFIG.SHEETS.INVENTORY);
  if (!inventory) throw new Error('Nie znaleziono arkusza bieżącej inwentaryzacji.');

  // Najpierw pokaż arkusz główny, aby nigdy nie próbować ukryć ostatniej zakładki.
  if (inventory.isSheetHidden()) inventory.showSheet();

  ss.getSheets().forEach(function(sheet) {
    const isCurrent = sheet.getSheetId() === inventory.getSheetId();
    const isArchive = /^ARCHIWUM\b/i.test(sheet.getName());
    const shouldShow = mode === 'ADMIN' ||
      isCurrent ||
      (mode === 'ALL_INVENTORIES' && isArchive);
    if (shouldShow) {
      if (sheet.isSheetHidden()) sheet.showSheet();
    } else if (!sheet.isSheetHidden()) {
      sheet.hideSheet();
    }
  });
}
