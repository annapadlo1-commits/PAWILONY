/**
 * Inventory PRO 4.3 — read-only inventory summary service.
 * Układ pochodzi wyłącznie z CONFIG.INVENTORY_LAYOUT.
 */

const INVENTORY_SUMMARY_COLUMNS_ = Object.freeze({
  NORMAL: true,
  KEG: true,
  LOCATION: true
});

function getInventorySummaryLayout_(productType) {
  const type = String(productType || '').trim().toUpperCase();
  const layout = getConfiguredInventoryLayout_(type);
  if (!layout) throw new Error('Nieobsługiwany typ produktu: ' + type + '.');
  return layout;
}

function readInventorySummaryItem_(sheet, product) {
  if (!sheet || !product || !product.inventoryRow) return null;

  const type = String(product.type || '').trim().toUpperCase();
  const row = Number(product.inventoryRow);
  if (isDirectFinalInventoryProduct_(product)) {
    const directValue = readCellNumberOrBlank_(sheet, 'B', row);
    return {
      key: product.normalizedName || normalizeText(product.name),
      product: product.name,
      category: normalizeBusinessCategory_(product.category) || String(product.category || '').trim(),
      type: type,
      inventoryRow: row,
      unit: 'szt.',
      finalTotal: directValue,
      total: directValue,
      details: { directFinal: directValue },
      cells: { finalTotal: 'B' + row, directFinal: 'B' + row },
      hasValue: directValue !== ''
    };
  }
  const layout = getInventorySummaryLayout_(type);
  const item = {
    key: product.normalizedName || normalizeText(product.name),
    product: product.name,
    category: normalizeBusinessCategory_(product.category) || String(product.category || '').trim(),
    type: type,
    inventoryRow: row,
    unit: layout.unit,
    finalTotal: readCellNumberOrBlank_(sheet, layout.finalTotal, row),
    details: {},
    cells: {}
  };

  addSummaryCellAddress_(item.cells, 'finalTotal', layout.finalTotal, row);

  if (type === CONFIG.PRODUCT_TYPES.LOCATION) {
    item.details = {
      warehouse: readCellNumberOrBlank_(sheet, layout.warehouse, row),
      darkroom: readCellNumberOrBlank_(sheet, layout.darkroom, row),
      fridges: readCellNumberOrBlank_(sheet, layout.fridges, row)
    };
    addSummaryCellAddress_(item.cells, 'warehouse', layout.warehouse, row);
    addSummaryCellAddress_(item.cells, 'darkroom', layout.darkroom, row);
    addSummaryCellAddress_(item.cells, 'fridges', layout.fridges, row);
  } else {
    item.details = {
      grossWeight: readCellNumberOrBlank_(sheet, layout.grossWeight, row),
      emptyContainerWeight: readCellNumberOrBlank_(sheet, layout.emptyContainerWeight, row),
      openNet: readCellNumberOrBlank_(sheet, layout.openNet, row),
      prepNet: readCellNumberOrBlank_(sheet, layout.prepNet, row),
      fullUnits: readCellNumberOrBlank_(sheet, layout.fullUnits, row),
      unitCapacity: readCellNumberOrBlank_(sheet, layout.unitCapacity, row),
      fullUnitsVolume: readCellNumberOrBlank_(sheet, layout.fullUnitsVolume, row)
    };
    addSummaryCellAddress_(item.cells, 'grossWeight', layout.grossWeight, row);
    addSummaryCellAddress_(item.cells, 'fullUnits', layout.fullUnits, row);
  }

  item.hasValue = item.finalTotal !== '' || Object.keys(item.details).some(key => item.details[key] !== '');
  return item;
}

function addSummaryCellAddress_(target, key, column, row) {
  const normalized = normalizeColumnLetter_(column);
  if (normalized) target[key] = normalized + row;
}

function readCellNumberOrBlank_(sheet, column, row) {
  const normalized = normalizeColumnLetter_(column);
  if (!normalized) return '';
  const value = sheet.getRange(normalized + row).getValue();
  if (value === '' || value === null || value === undefined) return '';
  const number = Number(value);
  return Number.isFinite(number) ? number : value;
}

function verifyInventoryFormulaSafety_(sheet, product) {
  if (isDirectFinalInventoryProduct_(product)) return [];
  const type = String(product.type || '').trim().toUpperCase();
  const layout = getInventorySummaryLayout_(type);
  const row = Number(product.inventoryRow);
  return (layout.formulaColumns || []).filter(Boolean).map(column => {
    const cell = normalizeColumnLetter_(column) + row;
    const range = sheet.getRange(cell);
    const formula = range.getFormula();
    return { cell: cell, formula: formula, hasFormula: Boolean(formula) };
  });
}
