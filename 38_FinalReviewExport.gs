/**
 * Inventory PRO Enterprise v2.6
 * Final Review & Export.
 */

function showFinalReview() {
  const html = renderInventoryTemplate_('UI_FinalReview')
    .setWidth(1320)
    .setHeight(900);

  SpreadsheetApp.getUi().showModalDialog(
    html,
    '🍕 Zakończ inwentaryzację — przegląd i eksport'
  );
}

function getFinalReviewSources() {
  const sources = SpreadsheetApp.getActiveSpreadsheet().getSheets()
    .map(sheet => sheet.getName())
    .filter(name => {
      const normalized = normalizeText(name);
      return isConfiguredSheetName_(name, CONFIG.SHEETS.INVENTORY) || /^ARCHIWUM\b/i.test(name);
    })
    .map(name => ({
      name: name,
      current: isConfiguredSheetName_(name, CONFIG.SHEETS.INVENTORY)
    }));

  sources.sort((a, b) => {
    if (a.current) return -1;
    if (b.current) return 1;
    return b.name.localeCompare(a.name, 'pl');
  });

  return sources;
}

function getFinalReviewData(sourceSheetName) {
  let diagnosticStage = 'wybór arkusza';
  try {
    const selectedSheet = resolveFinalReviewSheetName_(sourceSheetName);
    diagnosticStage = 'budowanie raportu';
    const snapshot = buildFinalInventorySnapshot_(selectedSheet);
    diagnosticStage = 'kontrola katalogu';
    const activeCatalogCount = Number(
      snapshot.diagnostics && snapshot.diagnostics.activeCatalogCount
    ) || 0;
  if (activeCatalogCount > 0 && (!snapshot.items || !snapshot.items.length)) {
    throw new Error(
      'BŁĄD KRYTYCZNY RAPORTU: aktywny katalog zawiera ' + activeCatalogCount +
      ' produktów, ale silnik raportu zwrócił 0 pozycji. ' +
      'Nie wolno kontynuować eksportu. Wdróż kompletną paczkę ' + CONFIG.VERSION + '.'
    );
  }
    diagnosticStage = 'sesja inwentaryzacji';
    const session = ensureActiveInventorySession_();
    diagnosticStage = 'pola raportu';
    const reviewItems = snapshot.items.map(item => {
      const editableCells = buildEditableReviewCells_(item);
      const copy = Object.assign({}, item);
      copy.cells = editableCells;
      return copy;
    });

    diagnosticStage = 'dane dodatkowe';
    return {
      version: CONFIG.VERSION,
      releaseSignature: 'FINAL_REPORT_DATA_CONTRACT_548',
      diagnostics: snapshot.diagnostics,
      sourceSheetName: selectedSheet,
      isCurrentInventory: isConfiguredSheetName_(selectedSheet, CONFIG.SHEETS.INVENTORY),
      sessionId: session.id,
      startedAt: session.startedAt,
      generatedAt: new Date().toISOString(),
      summary: snapshot.summary,
      items: reviewItems,
      reviewCategories: buildFinalReviewCategoryGroups_(reviewItems),
      warnings: snapshot.warnings,
      validationIssues: snapshot.validationIssues,
      newProducts: getNewProductsForSession_(session.startedAt),
      duplicates: getDuplicateGroupsForSession_(session.startedAt)
    };
  } catch (error) {
    throw new Error(
      'RAPORT 5.4.8 — etap „' + diagnosticStage + '”: ' +
      (error && error.message ? error.message : String(error))
    );
  }
}

function getFinalReviewDataJson(sourceSheetName) {
  try {
    const payload = getFinalReviewData(sourceSheetName);
    const json = JSON.stringify(payload);
    if (!json || json === 'null') {
      throw new Error('Serializacja raportu zwróciła pustą odpowiedź.');
    }
    return json;
  } catch (error) {
    throw new Error(
      'RAPORT 5.4.9 — most JSON: ' +
      (error && error.message ? error.message : String(error))
    );
  }
}

function validateFinalReviewCategoryConfirmations_(items, confirmedCategories) {
  const confirmed = {};
  (confirmedCategories || []).forEach(category => {
    confirmed[normalizeText(category)] = true;
  });
  const required = buildFinalReviewCategoryGroups_(items);
  const missing = required.filter(group => !confirmed[normalizeText(group.category)]);
  if (missing.length) {
    throw new Error(
      'Przed eksportem zatwierdź kategorie wymagające weryfikacji: ' +
      missing.map(group => group.category).join(', ') + '.'
    );
  }
  return required;
}

function approveUnknownTareGrossAsNet(productKey, sourceSheetName) {
  const selectedSheet = resolveFinalReviewSheetName_(sourceSheetName);
  if (!isConfiguredSheetName_(selectedSheet, CONFIG.SHEETS.INVENTORY)) {
    throw new Error('Decyzję o wadze brutto można zapisać tylko dla bieżącej inwentaryzacji.');
  }
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    SpreadsheetApp.flush();
    invalidateProductCatalogCache_();
    const catalog = buildProductCatalogUncached_();
    const key = normalizeText(productKey);
    const product = catalog.find(item =>
      item.normalizedName === key || normalizeText(item.name) === key
    );
    if (!product || !product.inventoryRow) throw new Error('Nie znaleziono produktu w bieżącej inwentaryzacji.');
    const profile = getProductPackagingProfile_(product);
    if (!profile || profile.mode !== CONFIG.PACKAGING_MODES.TARE_UNKNOWN) {
      throw new Error('Produkt nie ma profilu „tara nieznana”.');
    }
    const sheet = getSheetByConfiguredName_(CONFIG.SHEETS.INVENTORY);
    const layout = getConfiguredInventoryLayout_(product.type);
    const gross = sheet.getRange(layout.grossWeight + product.inventoryRow).getValue();
    if (gross === '' || !Number.isFinite(Number(gross)) || Number(gross) < 0) {
      throw new Error('Najpierw wpisz prawidłową wagę brutto produktu.');
    }
    saveUnknownTareGrossApproval_(product, true);
    applyCanonicalFormulasToProductRow_(sheet, product);
    SpreadsheetApp.flush();
    refreshInventoryStatusColors_(sheet);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Zatwierdzono całą wagę brutto dla: ' + product.name,
      'Inventory PRO',
      6
    );
    return {success: true, product: product.name, grossWeight: Number(gross)};
  } finally {
    lock.releaseLock();
  }
}

function approveUnknownTareGrossAsNetBatch(productKeys, sourceSheetName) {
  const selectedSheet = resolveFinalReviewSheetName_(sourceSheetName);
  if (!isConfiguredSheetName_(selectedSheet, CONFIG.SHEETS.INVENTORY)) {
    throw new Error('Decyzję o wadze brutto można zapisać tylko dla bieżącej inwentaryzacji.');
  }
  const requested = Array.from(new Set((productKeys || []).map(normalizeText).filter(Boolean)));
  if (!requested.length) return {success: true, approved: 0, products: []};

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    SpreadsheetApp.flush();
    const catalog = buildReportingCatalog_();
    const byKey = {};
    catalog.forEach(product => { byKey[product.normalizedName] = product; });
    const sheet = getSheetByConfiguredName_(CONFIG.SHEETS.INVENTORY);
    const state = loadUnknownTareGrossApprovals_();
    const approvedProducts = [];

    requested.forEach(key => {
      const product = byKey[key];
      if (!product || !product.inventoryRow) {
        throw new Error('Nie znaleziono produktu dla decyzji o wadze brutto: ' + key + '.');
      }
      const profile = getProductPackagingProfile_(product);
      if (!profile || profile.mode !== CONFIG.PACKAGING_MODES.TARE_UNKNOWN) {
        throw new Error('Produkt „' + product.name + '” nie ma profilu „tara nieznana”.');
      }
      const layout = getConfiguredInventoryLayout_(product.type);
      const gross = sheet.getRange(layout.grossWeight + product.inventoryRow).getValue();
      if (gross === '' || !Number.isFinite(Number(gross)) || Number(gross) < 0) {
        throw new Error('Produkt „' + product.name + '” nie ma prawidłowej wagi brutto.');
      }
      state.products[key] = true;
      approvedProducts.push(product);
    });

    PropertiesService.getDocumentProperties().setProperty(
      UNKNOWN_TARE_GROSS_APPROVAL_PROPERTY_, JSON.stringify(state)
    );
    UNKNOWN_TARE_GROSS_APPROVAL_RUNTIME_CACHE_ = state;
    approvedProducts.forEach(product => applyCanonicalFormulasToProductRow_(sheet, product));
    SpreadsheetApp.flush();
    refreshInventoryStatusColors_(sheet);
    return {
      success: true,
      approved: approvedProducts.length,
      products: approvedProducts.map(product => product.name)
    };
  } finally {
    lock.releaseLock();
  }
}

function applyFinalReviewCorrections(corrections, sourceSheetName) {
  if (!Array.isArray(corrections) || !corrections.length) {
    return { success: true, updated: 0 };
  }

  const selectedSheet = resolveFinalReviewSheetName_(sourceSheetName);
  if (!isConfiguredSheetName_(selectedSheet, CONFIG.SHEETS.INVENTORY)) {
    throw new Error('Archiwalna inwentaryzacja jest tylko do odczytu.');
  }
  const sheet = getSheetByConfiguredName_(CONFIG.SHEETS.INVENTORY);
  if (!sheet) throw new Error('Nie znaleziono arkusza: ' + selectedSheet + '.');

  const snapshot = buildFinalInventorySnapshot_(selectedSheet);
  const allowedCells = {};
  snapshot.items.forEach(item => {
    const cells = buildEditableReviewCells_(item);
    Object.keys(cells).forEach(label => {
      const a1 = String(cells[label] || '').trim().toUpperCase();
      if (!a1) return;
      allowedCells[a1] = {
        product: item.product,
        productType: String(item.type || '').toUpperCase(),
        label: label
      };
    });
  });

  const seen = {};
  const prepared = corrections.map(correction => {
    const a1 = String(correction.a1 || '').trim().toUpperCase();
    const allowed = allowedCells[a1];
    if (!/^[A-Z]+\d+$/.test(a1) || !allowed) {
      throw new Error('Komórka ' + (a1 || 'BRAK') + ' nie jest polem edytowalnym końcowego przeglądu.');
    }
    if (seen[a1]) throw new Error('Komórka ' + a1 + ' występuje w korektach więcej niż raz.');
    seen[a1] = true;

    const column = (a1.match(/^[A-Z]+/) || [''])[0];
    if (
      isFormulaColumnForProductType_(allowed.productType, column) ||
      !isAllowedInputColumnForProductType_(allowed.productType, column)
    ) {
      throw new Error(
        'Korekta ' + a1 + ' została zablokowana: kolumna ' + column +
        ' nie jest polem wejściowym typu ' + allowed.productType + '.'
      );
    }

    const raw = correction.value;
    if (raw === '' || raw === null || raw === undefined) {
      return {
        a1: a1, blank: true, value: '', product: allowed.product,
        productType: allowed.productType
      };
    }

    const value = Number(String(raw).replace(',', '.'));
    if (!Number.isFinite(value) || value < 0) {
      throw new Error('Nieprawidłowa wartość dla komórki ' + a1 + '.');
    }
    return {
      a1: a1, blank: false, value: value, product: allowed.product,
      productType: allowed.productType
    };
  });

  const snapshots = prepared.map(item => {
    const range = sheet.getRange(item.a1);
    const formula = range.getFormula();
    if (formula) {
      throw new Error('Korekta zablokowana: komórka ' + item.a1 + ' zawiera formułę.');
    }
    return {
      range: range,
      a1: item.a1,
      value: range.getValue(),
      formula: formula,
      intendedValue: item.blank ? '' : item.value
    };
  });

  try {
    prepared.forEach((item, index) => {
      const snapshotItem = snapshots[index];
      const liveFormula = snapshotItem.range.getFormula();
      const liveValue = snapshotItem.range.getValue();
      if (liveFormula || !inventoryCellValuesEqual_(liveValue, snapshotItem.value)) {
        throw new Error(
          'Komórka ' + item.a1 + ' została zmieniona równolegle. Korekty przerwano.'
        );
      }
      if (item.blank) snapshotItem.range.clearContent();
      else snapshotItem.range.setValue(item.value);
    });
    SpreadsheetApp.flush();
    refreshInventoryStatusColors_(sheet);
  } catch (error) {
    snapshots.slice().reverse().forEach(snapshotItem => {
      const liveValue = snapshotItem.range.getValue();
      const liveFormula = snapshotItem.range.getFormula();
      if (!liveFormula && inventoryCellValuesEqual_(liveValue, snapshotItem.intendedValue)) {
        if (snapshotItem.value === '' || snapshotItem.value === null || snapshotItem.value === undefined) {
          snapshotItem.range.clearContent();
        } else {
          snapshotItem.range.setValue(snapshotItem.value);
        }
      } else if (!inventoryCellValuesEqual_(liveValue, snapshotItem.value)) {
        logWarning(
          'FinalReviewExport',
          'applyFinalReviewCorrections.rollback',
          'Nie cofnięto komórki zmienionej równolegle.',
          { cell: snapshotItem.a1, liveValue: liveValue }
        );
      }
    });
    SpreadsheetApp.flush();
    throw error;
  }
  return { success: true, updated: prepared.length };
}

function finalizeInventoryAndExport(options) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    const data = options || {};
    const sourceSheetName = resolveFinalReviewSheetName_(data.sourceSheetName);
    const isCurrentInventory = isConfiguredSheetName_(sourceSheetName, CONFIG.SHEETS.INVENTORY);
    const corrections = Array.isArray(data.corrections) ? data.corrections : [];
    if (corrections.length) applyFinalReviewCorrections(corrections, sourceSheetName);

    const snapshot = buildFinalInventorySnapshot_(sourceSheetName);
    const blockingIssues = (snapshot.validationIssues || []).filter(issue =>
      String(issue.severity || '').toUpperCase() === 'ERROR'
    );
    if (blockingIssues.length) {
      throw new Error(
        'Eksport zablokowany przez błędy danych lub formuł: ' +
        blockingIssues.slice(0, 5).map(issue =>
          (issue.product ? issue.product + ' — ' : '') + issue.message
        ).join('; ') +
        (blockingIssues.length > 5 ? ' (oraz ' + (blockingIssues.length - 5) + ' kolejnych).' : '.')
      );
    }
    const confirmedReviewGroups = validateFinalReviewCategoryConfirmations_(
      snapshot.items,
      data.confirmedReviewCategories
    );
    const session = ensureActiveInventorySession_();
    const excluded = new Set((data.excludedProductKeys || []).map(String));
    const filteredItems = snapshot.items.filter(item => !excluded.has(item.key));
    const exportId = createInventoryExportId_();

    const exportSpreadsheet = createFinalExportSpreadsheet_({
      exportId: exportId,
      session: session,
      snapshot: snapshot,
      items: filteredItems,
      excludedCount: snapshot.items.length - filteredItems.length,
      note: String(data.note || '').trim(),
      confirmedReviewCategories: confirmedReviewGroups.map(group => group.category),
      sourceSheetName: sourceSheetName
    });

    const xlsxFile = exportSpreadsheetAsXlsx_(exportSpreadsheet, exportId);
    const pdfFile = data.includePdf === true ? exportManagerReportAsPdf_(exportSpreadsheet, exportId) : null;
    const archiveSheetName = isCurrentInventory ? createFinalInventoryArchive_(exportId) : sourceSheetName;

      appendExportEvent_({
      exportId: exportId,
      sessionId: session.id,
      startedAt: session.startedAt,
      finishedAt: new Date(),
      itemsCount: filteredItems.length,
      warningsCount: snapshot.warnings.length,
      excludedCount: snapshot.items.length - filteredItems.length,
      spreadsheetUrl: exportSpreadsheet.getUrl(),
      xlsxUrl: xlsxFile.getUrl(),
      pdfUrl: pdfFile ? pdfFile.getUrl() : '',
      archiveSheetName: archiveSheetName,
      note: String(data.note || '').trim(),
      sourceSheetName: sourceSheetName
    });

    if (isCurrentInventory) {
      const inventorySheet = getSheetByConfiguredName_(CONFIG.SHEETS.INVENTORY);
      clearCurrentInventoryData_(inventorySheet);
      clearInventoryStatusColors_(inventorySheet);
      closeActiveInventorySession_(exportId);
      startInventorySession_();
      appendApplicationEvent_('INVENTORY_STARTED','Rozpoczęto nową inwentaryzację po zamknięciu poprzedniej',{previousExportId:exportId});
    }

    logInfo('FinalReviewExport', 'finalizeInventoryAndExport', 'Zakończono inwentaryzację i utworzono eksport', {
      exportId: exportId,
      items: filteredItems.length,
      warnings: snapshot.warnings.length,
      confirmedReviewCategories: confirmedReviewGroups.map(group => group.category),
      xlsxFileId: xlsxFile.getId(),
      pdfFileId: pdfFile ? pdfFile.getId() : ''
    });

    return {
      success: true,
      exportId: exportId,
      xlsxUrl: xlsxFile.getUrl(),
      pdfUrl: pdfFile ? pdfFile.getUrl() : '',
      spreadsheetUrl: exportSpreadsheet.getUrl(),
      archiveSheetName: archiveSheetName,
      exportedItems: filteredItems.length,
      warnings: snapshot.warnings.length
    };
  } finally {
    lock.releaseLock();
  }
}

function buildFinalInventorySnapshot_(sourceSheetName) {
  const report = generateInventoryReport(sourceSheetName);
  return {
    items: report.items,
    warnings: report.warnings,
    summary: report.summary,
    validationIssues: report.validationIssues,
    statistics: report.statistics,
    metadata: report.metadata,
    diagnostics: report.diagnostics
  };
}

function buildLegacyReviewValues_(summaryItem) {
  const details = summaryItem.details || {};
  if (summaryItem.type === CONFIG.PRODUCT_TYPES.LOCATION) {
    const values = {};
    getLocationAreaDefinitions_().forEach(area => {
      values[area.label] = normalizeReportNumber_(details[area.columnKey]);
    });
    values['Stan końcowy'] = normalizeReportNumber_(summaryItem.finalTotal);
    return values;
  }
  if (summaryItem.type === CONFIG.PRODUCT_TYPES.KEG) {
    return {
      'Waga brutto': normalizeReportNumber_(details.grossWeight),
      'Waga pustego kega': normalizeReportNumber_(details.emptyContainerWeight),
      'Otwarty keg netto': normalizeReportNumber_(details.openNet),
      'PREP netto': normalizeReportNumber_(details.prepNet),
      'Pełne kegi': normalizeReportNumber_(details.fullUnits),
      'Pojemność kega': normalizeReportNumber_(details.unitCapacity),
      'Pełne kegi w l': normalizeReportNumber_(details.fullUnitsVolume),
      'Stan końcowy': normalizeReportNumber_(summaryItem.finalTotal)
    };
  }
  return {
    'Waga brutto': normalizeReportNumber_(details.grossWeight),
    'Waga pustej butelki': normalizeReportNumber_(details.emptyContainerWeight),
    'Otwarta zawartość netto': normalizeReportNumber_(details.openNet),
    'PREP netto': normalizeReportNumber_(details.prepNet),
    'Pełne butelki': normalizeReportNumber_(details.fullUnits),
    'Pojemność butelki': normalizeReportNumber_(details.unitCapacity),
    'Pełne butelki w l': normalizeReportNumber_(details.fullUnitsVolume),
    'Stan końcowy': normalizeReportNumber_(summaryItem.finalTotal)
  };
}

function buildEditableReviewCells_(summaryItem) {
  const cells = summaryItem.cells || {};
  if (isDirectFinalInventoryProduct_({name: summaryItem.product})) {
    return {
      'Stan końcowy': cells.directFinal || cells.finalTotal
    };
  }
  if (summaryItem.type === CONFIG.PRODUCT_TYPES.LOCATION) {
    const editable = {};
    getLocationAreaDefinitions_().forEach(area => {
      const a1 = cells[area.columnKey];
      if (a1) editable[area.label] = a1;
    });
    return editable;
  }
  if (summaryItem.type === CONFIG.PRODUCT_TYPES.KEG) {
    return {
      'Waga brutto': cells.grossWeight,
      'Pełne kegi': cells.fullUnits
    };
  }
  return {
    'Waga brutto': cells.grossWeight,
    'Pełne butelki': cells.fullUnits
  };
}

function applySummaryWarnings_(item, settings) {
  const details = item.details || {};
  if (isDirectFinalInventoryProduct_({name:item.product})) {
    return { 'Stan końcowy': normalizeReportNumber_(item.finalTotal) };
  }
  if (item.type === CONFIG.PRODUCT_TYPES.LOCATION) {
    getLocationAreaDefinitions_().forEach(area => {
      const value = details[area.columnKey];
      if (value !== '' && Number(value) > settings.locationWarning) item.flags.push('DUŻA ILOŚĆ');
    });
  } else if (item.type === CONFIG.PRODUCT_TYPES.KEG) {
    if (normalizeReportNumber_(details.grossWeight) !== '' && Number(normalizeReportNumber_(details.grossWeight)) > settings.kegWeightWarning) item.flags.push('DUŻA WAGA');
    if (normalizeReportNumber_(details.fullUnits) !== '' && Number(normalizeReportNumber_(details.fullUnits)) > settings.kegWholeWarning) item.flags.push('DUŻA ILOŚĆ');
  } else {
    if (normalizeReportNumber_(details.grossWeight) !== '' && Number(normalizeReportNumber_(details.grossWeight)) > settings.normalWeightWarning) item.flags.push('DUŻA WAGA');
    if (normalizeReportNumber_(details.fullUnits) !== '' && Number(normalizeReportNumber_(details.fullUnits)) > settings.normalWholeWarning) item.flags.push('DUŻA ILOŚĆ');
  }

  const entered = Object.keys(details).filter(key => details[key] !== '');
  if (settings.warnZero && entered.length && entered.every(key => Number(details[key]) === 0)) {
    item.flags.push('WPROWADZONO 0');
  }
}

function numericOrZero_(value) {
  const number = Number(value);
  return Number.isFinite(number) ? number : 0;
}

function createFinalExportSpreadsheet_(payload) {
  const source = SpreadsheetApp.getActiveSpreadsheet();
  const name = 'Inventory PRO ' + payload.exportId;
  const exportSpreadsheet = SpreadsheetApp.create(name);
  const summarySheet = exportSpreadsheet.getSheets()[0];
  summarySheet.setName('PODSUMOWANIE');

  buildExportManagerReportSheet_(summarySheet, payload);
  buildExportFinalSummarySheet_(exportSpreadsheet.insertSheet('PODSUMOWANIE FINALNE'), payload.items);
  buildExportNormalDetailsSheet_(exportSpreadsheet.insertSheet('SZCZEGÓŁY NORMAL'), payload.items);
  buildExportKegDetailsSheet_(exportSpreadsheet.insertSheet('SZCZEGÓŁY KEG'), payload.items);
  buildExportLocationDetailsSheet_(exportSpreadsheet.insertSheet('SZCZEGÓŁY LOCATION'), payload.items);
  buildExportWarningsSheet_(exportSpreadsheet.insertSheet('OSTRZEŻENIA'), payload.snapshot.warnings);
  buildExportNewProductsSheet_(exportSpreadsheet.insertSheet('NOWE PRODUKTY'), getNewProductsForSession_(payload.session.startedAt));
  buildExportDuplicatesSheet_(exportSpreadsheet.insertSheet('DUPLIKATY'), getDuplicateGroupsForSession_(payload.session.startedAt));
  buildExportStatisticsSheet_(exportSpreadsheet.insertSheet('STATYSTYKI'), payload);
  buildExportMetadataSheet_(exportSpreadsheet.insertSheet('METADANE'), payload);

  SpreadsheetApp.flush();
  return exportSpreadsheet;
}

function buildExportSummarySheet_(sheet, payload) {
  const source = SpreadsheetApp.getActiveSpreadsheet();
  const summary = payload.snapshot.summary;
  const values = [
    ['🍕 INVENTORY PRO — PODSUMOWANIE'],
    ['Numer eksportu', payload.exportId],
    ['Lokal / plik źródłowy', source.getName()],
    ['Arkusz źródłowy', payload.sourceSheetName || CONFIG.SHEETS.INVENTORY],
    ['Rozpoczęto', new Date(payload.session.startedAt)],
    ['Zakończono', new Date()],
    ['Użytkownik', getCurrentUserEmail_()],
    ['Wersja aplikacji', CONFIG.VERSION],
    ['Produkty w katalogu', summary.products],
    ['Pozycje uzupełnione', summary.completed],
    ['Brak wartości', summary.missing],
    ['Produkty z ostrzeżeniami', summary.warningProducts],
    ['Zatwierdzone kategorie wymagające weryfikacji', (payload.confirmedReviewCategories || []).join(', ') || 'Brak'],
    ['Wykluczono z eksportu', payload.excludedCount],
    ['Notatka', payload.note || '']
  ];

  sheet.getRange(1, 1, values.length, 2).setValues(values.map(row => row.length === 1 ? [row[0], ''] : row));
  sheet.getRange('A1:B1').merge().setBackground('#f6b26b').setFontWeight('bold').setFontSize(16).setHorizontalAlignment('center');
  sheet.getRange(2, 1, values.length - 1, 1).setFontWeight('bold').setBackground('#fff2cc');
  sheet.setColumnWidth(1, 230);
  sheet.setColumnWidth(2, 420);
  sheet.setFrozenRows(1);
}

function buildExportFinalSummarySheet_(sheet, items) {
  const headers = ['Kategoria', 'Produkt', 'Typ', 'Stan końcowy', 'Jednostka', 'Status'];
  const rows = items.map(item => [
    item.category,
    item.product,
    item.type,
    item.finalTotal,
    item.unit,
    item.flags.join(', ') || 'OK'
  ]);
  writeExportTable_(sheet, headers, rows, '#f6b26b');
}

function buildExportNormalDetailsSheet_(sheet, items) {
  const layout = getInventorySummaryLayout_(CONFIG.PRODUCT_TYPES.NORMAL);
  const headers = [
    'Kategoria', 'Produkt',
    'Tryb opakowania', 'Tara profilu', 'Waga referencyjna',
    'Ilość początkowa', 'Przelicznik masy', 'Data referencji',
    'Waga brutto ' + layout.grossWeight,
    'Waga pustej butelki ' + layout.emptyContainerWeight,
    'Otwarta zawartość ' + layout.openNet,
    'PREP netto ' + layout.prepNet,
    'Pełne sztuki ' + layout.fullUnits,
    'Pojemność ' + layout.unitCapacity,
    'Pełne butelki w l ' + layout.fullUnitsVolume,
    'Stan końcowy ' + layout.finalTotal,
    'Jednostka'
  ];
  const rows = items.filter(item => item.type === CONFIG.PRODUCT_TYPES.NORMAL).map(item => {
    const d = item.details || {};
    const p = item.packaging || {};
    if (isDirectFinalInventoryProduct_({name: item.product})) {
      return [
        item.category, item.product,
        '', '', '', '', '', '',
        '', '', '', '',
        d.directFinal,
        '', '',
        item.finalTotal,
        item.unit
      ];
    }
    return [item.category, item.product, p.mode || '', p.emptyWeight, p.referenceGrossWeight,
      p.referenceNetQuantity, p.massConversionFactor, p.referenceDate,
      d.grossWeight, d.emptyContainerWeight, d.openNet, d.prepNet, d.fullUnits,
      d.unitCapacity, d.fullUnitsVolume, item.finalTotal, item.unit];
  });
  writeExportTable_(sheet, headers, rows, '#fff2cc');
}

function buildExportKegDetailsSheet_(sheet, items) {
  const layout = getInventorySummaryLayout_(CONFIG.PRODUCT_TYPES.KEG);
  const headers = [
    'Kategoria', 'Produkt',
    'Tryb opakowania', 'Tara profilu', 'Waga referencyjna',
    'Ilość początkowa', 'Przelicznik masy', 'Data referencji',
    'Waga brutto ' + layout.grossWeight,
    'Waga pustego kega ' + layout.emptyContainerWeight,
    'Otwarty keg ' + layout.openNet,
    'PREP netto ' + layout.prepNet,
    'Pełne kegi ' + layout.fullUnits,
    'Pojemność kega ' + layout.unitCapacity,
    'Pełne kegi w l ' + layout.fullUnitsVolume,
    'Stan końcowy ' + layout.finalTotal,
    'Jednostka'
  ];
  const rows = items.filter(item => item.type === CONFIG.PRODUCT_TYPES.KEG).map(item => {
    const d = item.details || {};
    const p = item.packaging || {};
    return [item.category, item.product, p.mode || '', p.emptyWeight, p.referenceGrossWeight,
      p.referenceNetQuantity, p.massConversionFactor, p.referenceDate,
      d.grossWeight, d.emptyContainerWeight, d.openNet, d.prepNet, d.fullUnits,
      d.unitCapacity, d.fullUnitsVolume, item.finalTotal, item.unit];
  });
  writeExportTable_(sheet, headers, rows, '#cfe2f3');
}

function buildExportLocationDetailsSheet_(sheet, items) {
  const layout = getInventorySummaryLayout_(CONFIG.PRODUCT_TYPES.LOCATION);
  const areas = getLocationAreaDefinitions_();
  const headers = ['Kategoria', 'Produkt'].concat(
    areas.map(area => area.label + ' ' + String(layout[area.columnKey] || '')),
    ['Stan końcowy ' + layout.finalTotal, 'Jednostka']
  );
  const rows = items.filter(item => item.type === CONFIG.PRODUCT_TYPES.LOCATION).map(item => {
    const d = item.details || {};
    return [item.category, item.product]
      .concat(areas.map(area => d[area.columnKey]), [item.finalTotal, item.unit]);
  });
  writeExportTable_(sheet, headers, rows, '#d9ead3');
}

function buildExportWarningsSheet_(sheet, warnings) {
  const headers = ['Produkt', 'Kategoria', 'Flagi', 'Wartości'];
  const rows = warnings.map(item => [item.product, item.category, item.flags.join(', '), JSON.stringify(item.values)]);
  writeExportTable_(sheet, headers, rows, '#f4cccc');
}

function buildExportNewProductsSheet_(sheet, rows) {
  const headers = ['Data', 'Użytkownik', 'Produkt', 'Produkt referencyjny', 'Typ', 'Kategoria', 'Wiersz', 'Wpis źródłowy'];
  writeExportTable_(sheet, headers, rows, '#cfe2f3');
}

function buildExportDuplicatesSheet_(sheet, rows) {
  const headers = ['Import ID', 'Produkt', 'Liczba wystąpień', 'Wartości źródłowe', 'Suma'];
  writeExportTable_(sheet, headers, rows, '#ead1dc');
}

function writeExportTable_(sheet, headers, rows, headerColor) {
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground(headerColor);
  if (rows.length) sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
  if (sheet.getLastRow() > 1) sheet.getRange(1, 1, sheet.getLastRow(), headers.length).createFilter();
}

function exportSpreadsheetAsXlsx_(spreadsheet, exportId) {
  const url = buildSpreadsheetXlsxExportUrl_(spreadsheet.getId());
  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error('Nie udało się wygenerować XLSX. Kod: ' + response.getResponseCode() + '. ' + response.getContentText().slice(0, 300));
  }

  const blob = response.getBlob().setName(exportId + '.xlsx');
  if (!blob.getBytes().length) {
    throw new Error('Nie udało się wygenerować XLSX: serwer zwrócił pusty plik.');
  }
  return getExportFolder_().createFile(blob);
}

function buildSpreadsheetXlsxExportUrl_(spreadsheetId) {
  return 'https://docs.google.com/spreadsheets/d/' +
    encodeURIComponent(String(spreadsheetId || '')) +
    '/export?format=xlsx';
}

function getExportFolder_() {
  const sourceFile = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  const parents = sourceFile.getParents();
  const parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const folderName = 'Inventory PRO — Eksporty';
  const folders = parent.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : parent.createFolder(folderName);
}

function createFinalInventoryArchive_(exportId) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const inventory = getSheetByConfiguredName_(CONFIG.SHEETS.INVENTORY);
  const base = 'ARCHIWUM ' + exportId.replace(/^INV-/, '');
  const name = createUniqueArchiveName_(spreadsheet, base.slice(0, 90));
  inventory.copyTo(spreadsheet).setName(name);
  return name;
}

function appendExportEvent_(data) {
  appendApplicationEvent_('EXPORT', 'Zakończono inwentaryzację i utworzono eksport', {
    exportId:data.exportId, sessionId:data.sessionId, itemsCount:data.itemsCount,
    warningsCount:data.warningsCount, excludedCount:data.excludedCount,
    spreadsheetUrl:data.spreadsheetUrl, xlsxUrl:data.xlsxUrl, pdfUrl:data.pdfUrl || '',
    archiveSheetName:data.archiveSheetName, note:data.note || '', sourceSheetName:data.sourceSheetName
  });
}

function getNewProductsForSession_(startedAt) {
  const sheet = getOrCreateNewProductsSheet_();
  if (sheet.getLastRow() < 2) return [];
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  const start = new Date(startedAt).getTime();
  return rows.filter(row => row[0] instanceof Date && row[0].getTime() >= start);
}

function getDuplicateGroupsForSession_(startedAt) {
  const auditSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.IMPORT_AUDIT);
  if (auditSheet && auditSheet.getLastRow() >= 2) {
    const auditRows = auditSheet.getRange(2, 1, auditSheet.getLastRow() - 1, 21).getValues();
    const start = new Date(startedAt).getTime();
    const auditGroups = auditRows
      .filter(row => row[1] instanceof Date && row[1].getTime() >= start && Number(row[16]) >= 2)
      .map(row => [row[0], row[4], row[16], row[17], row[18]]);
    if (auditGroups.length) return auditGroups;
  }

  // Zgodność z importami wykonanymi przed 4.2.1, kiedy te informacje były
  // zapisywane w osobnym arkuszu RAPORT.
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.REPORT);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 19).getValues();
  const start = new Date(startedAt).getTime();
  return rows
    .filter(row => row[1] instanceof Date && row[1].getTime() >= start && Number(row[11]) >= 2)
    .map(row => [row[0], row[4], row[11], row[12], row[13]]);
}

function createInventoryExportId_() {
  const zone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  return 'INV-' + Utilities.formatDate(new Date(), zone, 'yyyyMMdd-HHmmss');
}

function ensureActiveInventorySession_() {
  const properties = PropertiesService.getDocumentProperties();
  const raw = properties.getProperty('INVENTORY_PRO_ACTIVE_SESSION');
  if (raw) {
    try { return JSON.parse(raw); } catch (error) { /* odbuduj */ }
  }
  const session = { id: createUniqueId_('SESSION'), startedAt: new Date().toISOString() };
  properties.setProperty('INVENTORY_PRO_ACTIVE_SESSION', JSON.stringify(session));
  return session;
}

function startInventorySession_() {
  clearUnknownTareGrossApprovals_();
  const session = { id: createUniqueId_('SESSION'), startedAt: new Date().toISOString() };
  PropertiesService.getDocumentProperties().setProperty('INVENTORY_PRO_ACTIVE_SESSION', JSON.stringify(session));
  return session;
}

function closeActiveInventorySession_(exportId) {
  const properties = PropertiesService.getDocumentProperties();
  const session = ensureActiveInventorySession_();
  session.closedAt = new Date().toISOString();
  session.exportId = exportId;
  properties.setProperty('INVENTORY_PRO_LAST_CLOSED_SESSION', JSON.stringify(session));
  properties.deleteProperty('INVENTORY_PRO_ACTIVE_SESSION');
}


function resolveFinalReviewSheetName_(requestedName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requested = String(requestedName || CONFIG.SHEETS.INVENTORY).trim();

  if (!requested || isConfiguredSheetName_(requested, CONFIG.SHEETS.INVENTORY)) {
    const inventorySheet = getSheetByConfiguredName_(CONFIG.SHEETS.INVENTORY);
    if (!inventorySheet) throw new Error('Nie znaleziono bieżącego arkusza INWENTURA.');
    return inventorySheet.getName();
  }

  const requestedKey = normalizeText(requested);
  const sheet = ss.getSheets().find(item => normalizeText(item.getName()) === requestedKey);
  if (!sheet) throw new Error('Nie znaleziono arkusza: ' + requested + '.');
  return sheet.getName();
}
