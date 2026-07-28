function onOpen() {
  registerInventorySpreadsheet_();
  const ui = SpreadsheetApp.getUi();

  const reportingMenu = ui.createMenu('Raportowanie')
    .addItem('Końcowy przegląd i eksport', 'showFinalReview')
    .addItem('Dashboard bieżącej inwentaryzacji', 'showInventoryDashboard')
    .addItem('Porównaj inwentaryzacje', 'showInventoryAnalytics');

  const catalogMenu = ui.createMenu('Produkty i aliasy')
    .addItem('Product Manager', 'showProductManager')
    .addItem('Alias Manager', 'showAliasManager')
    .addItem('Nowe produkty', 'showNewProductsSheet_')
    .addItem('Audyt katalogu', 'runProductDataAudit');

  const diagnosticsMenu = ui.createMenu('Testy i diagnostyka')
    .addItem('Health Check', 'runEnterpriseHealthCheck')
    .addItem('Uruchom wszystkie testy', 'runAllEnterpriseTests')
    .addItem('Testy raportowania', 'runReportingEngineTests')
    .addItem('Waliduj konfigurację', 'validateEnterpriseConfiguration')
    .addItem('Pokaż diagnostykę', 'showEnterpriseDiagnostics')
    .addItem('Diagnostyka Parsera 4.3', 'showParserDiagnostics')
    .addItem('Audyt formuł', 'auditInventoryFormulaCoverageWithDialog');

  const configurationMenu = ui.createMenu('Konfiguracja')
    .addItem('Napraw / odśwież profile opakowań i formuły', 'migrateProductPackagingProfilesWithDialog')
    .addItem('Uporządkuj format liczb', 'normalizeInventoryNumberFormats')
    .addItem('Skonfiguruj transkrypcję Gemini', 'configureGeminiTranscription')
    .addItem('Sprawdź transkrypcję Gemini', 'showGeminiTranscriptionStatus');

  const maintenanceMenu = ui.createMenu('Konserwacja danych')
    .addItem('Wyczyść cache katalogu', 'clearProductCatalogCache')
    .addItem('Audyt danych produktów', 'runProductDataAudit')
    .addItem('Napraw dane produktów', 'repairProductData')
    .addItem('Odbuduj konfigurację SŁOWNIKA', 'rebuildDictionaryConfigurationSafely')
    .addItem('Napraw kategorie na podstawie INWENTURA', 'repairInvalidDictionaryCategories290WithDialog');

  const technicalMenu = ui.createMenu('Widok techniczny')
    .addItem('Pełny widok techniczny', 'showAdministrationWorkspace')
    .addItem('Wyczyść log techniczny', 'clearTechnicalLog');

  const adminMenu = ui.createMenu('Administracja')
    .addSubMenu(diagnosticsMenu)
    .addSubMenu(configurationMenu)
    .addSubMenu(maintenanceMenu)
    .addSubMenu(technicalMenu);

  const viewsMenu = ui.createMenu('Widoki i pomoc')
    .addItem('Historia zdarzeń', 'showHistory')
    .addItem('Ustawienia', 'showSettings')
    .addItem('Pomoc — przewodnik', 'showHelpGuide')
    .addItem('Widok — inwentaryzacja bieżąca', 'showCurrentInventoryWorkspace')
    .addItem('Widok — wszystkie inwentaryzacje', 'showAllInventoriesWorkspace')
    .addItem('Widok — administracja', 'showAdministrationWorkspace');

  ui.createMenu('INVENTORY PRO')
    .addItem('🍕 Import', 'showImport')
    .addItem('🍕 Rozpocznij / zakończ inwentaryzację', 'startNewInventory')
    .addItem('🍕 Wyczyść bieżącą inwenturę', 'clearCurrentInventory')
    .addSeparator()
    .addItem('Cofnij ostatni import', 'undoLastImport')
    .addSeparator()
    .addSubMenu(reportingMenu)
    .addSubMenu(catalogMenu)
    .addSubMenu(viewsMenu)
    .addSubMenu(adminMenu)
    .addSeparator()
    .addItem('🍕 O aplikacji', 'showAbout')
    .addToUi();

  // Każde otwarcie pliku zaczyna się w bezpiecznym widoku operacyjnym.
  try {
    setWorkspaceMode_('CURRENT');
    applyWorkspaceVisibility_('CURRENT');
    activateSheetByName_(CONFIG.SHEETS.INVENTORY);
  } catch (workspaceError) {
    console.error('Inventory PRO default workspace: ' + String(workspaceError));
  }

  logInfo('App', 'onOpen', 'Menu Production Edition zostało utworzone', {
    version: CONFIG.VERSION,
    location: CONFIG.LOCATION.NAME
  });
}

/**
 * Ręczna zmiana SŁOWNIKA lub nazw produktów w INWENTURZE musi od razu
 * unieważnić katalog. Bez tego Parser może pracować na danych sprzed 6 godzin.
 */
function onEdit(event) {
  try {
    const range = event && event.range;
    const sheet = range && range.getSheet();
    if (!sheet) return;
    const name = normalizeText(sheet.getName());
    if (
      name === normalizeText(CONFIG.SHEETS.DICTIONARY) ||
      (isConfiguredSheetName_(sheet.getName(), CONFIG.SHEETS.INVENTORY) && range.getColumn() === 1)
    ) {
      invalidateProductCatalogCache_();
    }
  } catch (error) {
    console.error('Inventory PRO onEdit cache invalidation: ' + String(error));
  }
}

function showNewProductsSheet_() {
  activateSheetByName_(CONFIG.SHEETS.NEW_PRODUCTS);
}

function showReport() {
  // Zgodność ze starszym menu i skrótami. Arkusz RAPORT jest technicznym
  // rejestrem importów; użytkownik otrzymuje właściwy przegląd końcowy.
  showFinalReview();
}

function showHistory() {
  const html = renderInventoryTemplate_('UI_History')
    .setWidth(1120)
    .setHeight(780);
  SpreadsheetApp.getUi().showModalDialog(html, 'Inventory PRO — Historia zdarzeń');
}


function showExportHistory() { showHistory(); }

function showSettings() {
  ensureQualitySettingsSheet_();
  activateSheetByName_(CONFIG.SHEETS.SETTINGS);
}

function activateSheetByName_(sheetName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheetByConfiguredName_(sheetName);

  if (!sheet) {
    const message = 'Nie znaleziono arkusza: ' + sheetName;

    logWarning(
      'Code',
      'activateSheetByName_',
      message,
      { sheetName }
    );

    SpreadsheetApp.getUi().alert(message);
    return;
  }

  if (sheet.isSheetHidden()) sheet.showSheet();
  spreadsheet.setActiveSheet(sheet);
}

function showAbout() {
  const html = renderInventoryTemplate_('UI_About')
    .setWidth(620)
    .setHeight(760);

  SpreadsheetApp.getUi().showModalDialog(
    html,
    'O aplikacji'
  );
}

function showHelpGuide() {
  const html = renderInventoryTemplate_('UI_Guide')
    .setWidth(920)
    .setHeight(760);

  SpreadsheetApp.getUi().showModalDialog(html, 'Inventory PRO — Pomoc');
}


function debugProductConfiguration() {
  return runSafely_(
    'Code',
    'debugProductConfiguration',
    function() {
      const products = loadProductConfigurations();
      const aliases = loadAliases();

      SpreadsheetApp.getUi().alert(
        'Inventory PRO - Debug\n\n' +
        'Aktywne produkty: ' + products.length + '\n' +
        'Aliasy: ' + Object.keys(aliases).length
      );
    },
    'Nie udalo sie sprawdzic konfiguracji.'
  );
}

function debugProductCatalog() {
  return runSafely_(
    'Code',
    'debugProductCatalog',
    function() {
      const summary = getProductCatalogSummary();

      SpreadsheetApp.getUi().alert(
        'Inventory PRO - Product Engine\n\n' +
        'Produkty: ' + summary.products + '\n' +
        'Aliasy: ' + summary.aliases + '\n' +
        'NORMAL: ' + summary.normal + '\n' +
        'KEG: ' + summary.keg + '\n' +
        'LOCATION: ' + summary.location + '\n' +
        'Brak wiersza w Inwentura: ' +
        summary.missingInventoryRow
      );
    },
    'Nie udalo sie zbudowac statystyk katalogu.'
  );
}

function debugSmartMatch() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Inventory PRO - Smart Match',
    'Wpisz nazwe produktu do przetestowania:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  return runSafely_(
    'Code',
    'debugSmartMatch',
    function() {
      const inputName = response.getResponseText();
      const result = matchProduct(inputName);

      let message =
        'Wpis: ' + inputName + '\n' +
        'Status: ' + result.status + '\n' +
        'Informacja: ' + result.message + '\n\n';

      if (result.product) {
        message +=
          'Wybrany produkt:\n' +
          result.product.name + '\n\n' +
          'Typ: ' + result.product.type + '\n' +
          'Kategoria: ' + result.product.category + '\n' +
          'Wiersz: ' +
          (result.product.inventoryRow || 'BRAK') +
          '\n' +
          'Wynik: ' + result.score;
      }

      if (
        !result.product &&
        result.candidates &&
        result.candidates.length
      ) {
        message += 'Mozliwe produkty:\n';

        result.candidates.forEach((candidate, index) => {
          message +=
            (index + 1) + '. ' +
            candidate.product.name +
            ' (' + candidate.score + ' pkt)\n';
        });
      }

      ui.alert('Wynik dopasowania', message, ui.ButtonSet.OK);
    },
    'Nie udalo sie wykonac testu Smart Match.'
  );
}
