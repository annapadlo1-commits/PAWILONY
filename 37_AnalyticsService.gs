/**
 * Inventory PRO Enterprise v2.10.0 — Inventory Analytics.
 * Porównania i trendy oparte wyłącznie na Reporting Engine.
 * Moduł jest tylko do odczytu i nie zmienia arkuszy źródłowych.
 */

function showInventoryAnalytics() {
  const html = renderInventoryTemplate_('UI_Analytics')
    .setWidth(1240)
    .setHeight(860);
  SpreadsheetApp.getUi().showModalDialog(html, '🍕 Inventory PRO — Porównanie inwentaryzacji');
}

function getInventoryAnalyticsSources() {
  return getFinalReviewSources();
}

function getInventoryComparisonData(baseSheetName, currentSheetName) {
  const started = Date.now();
  const baseReport = generateInventoryReport(baseSheetName);
  const currentReport = generateInventoryReport(currentSheetName);
  const baseMap = indexAnalyticsItems_(baseReport.items || []);
  const currentMap = indexAnalyticsItems_(currentReport.items || []);
  const keys = Array.from(new Set(Object.keys(baseMap).concat(Object.keys(currentMap))));
  const rows = [];

  keys.forEach(function(key) {
    const before = baseMap[key] || null;
    const after = currentMap[key] || null;
    const reference = after || before;
    if (!reference) return;

    const beforeValue = analyticsNumberOrNull_(before && before.finalTotal);
    const afterValue = analyticsNumberOrNull_(after && after.finalTotal);
    const comparable = Boolean(before && after && before.type === after.type && before.unit === after.unit);
    const delta = comparable && beforeValue !== null && afterValue !== null ? afterValue - beforeValue : null;
    const percent = delta !== null && beforeValue !== 0 ? (delta / Math.abs(beforeValue)) * 100 : null;
    let status = 'UNCHANGED';
    if (!before) status = 'NEW_PRODUCT';
    else if (!after) status = 'REMOVED_PRODUCT';
    else if (!before.hasValue && after.hasValue) status = 'NEWLY_COUNTED';
    else if (before.hasValue && !after.hasValue) status = 'MISSING_NOW';
    else if (!comparable || delta === null) status = 'NOT_COMPARABLE';
    else if (Math.abs(delta) < 0.0005) status = 'UNCHANGED';
    else if (delta > 0) status = 'INCREASE';
    else status = 'DECREASE';

    rows.push({
      key: key,
      category: reference.category || '',
      product: reference.product || '',
      type: reference.type || '',
      unit: reference.unit || '',
      before: beforeValue,
      after: afterValue,
      delta: delta === null ? null : roundAnalyticsNumber_(delta),
      percent: percent === null || !Number.isFinite(percent) ? null : roundAnalyticsNumber_(percent),
      status: status,
      beforeCompleted: Boolean(before && before.hasValue),
      afterCompleted: Boolean(after && after.hasValue)
    });
  });

  const comparableRows = rows.filter(function(row) { return row.delta !== null; });
  const categoryMap = {};
  comparableRows.forEach(function(row) {
    const categoryKey = [row.category, row.unit].join('|');
    if (!categoryMap[categoryKey]) {
      categoryMap[categoryKey] = {
        category: row.category,
        unit: row.unit,
        products: 0,
        before: 0,
        after: 0,
        delta: 0
      };
    }
    const bucket = categoryMap[categoryKey];
    bucket.products++;
    bucket.before += row.before || 0;
    bucket.after += row.after || 0;
    bucket.delta += row.delta || 0;
  });

  const categories = Object.keys(categoryMap).map(function(key) {
    const row = categoryMap[key];
    row.before = roundAnalyticsNumber_(row.before);
    row.after = roundAnalyticsNumber_(row.after);
    row.delta = roundAnalyticsNumber_(row.delta);
    row.percent = row.before !== 0 ? roundAnalyticsNumber_((row.delta / Math.abs(row.before)) * 100) : null;
    return row;
  }).sort(function(a, b) { return a.category.localeCompare(b.category, 'pl'); });

  const topIncreases = comparableRows.filter(function(row) { return row.delta > 0; })
    .sort(function(a, b) { return b.delta - a.delta; }).slice(0, 15);
  const topDecreases = comparableRows.filter(function(row) { return row.delta < 0; })
    .sort(function(a, b) { return a.delta - b.delta; }).slice(0, 15);

  return {
    version: CONFIG.VERSION,
    baseSheetName: baseReport.metadata.sourceSheetName,
    currentSheetName: currentReport.metadata.sourceSheetName,
    generatedAt: new Date().toISOString(),
    durationMs: Date.now() - started,
    summary: {
      products: rows.length,
      comparable: comparableRows.length,
      increases: rows.filter(function(row) { return row.status === 'INCREASE'; }).length,
      decreases: rows.filter(function(row) { return row.status === 'DECREASE'; }).length,
      unchanged: rows.filter(function(row) { return row.status === 'UNCHANGED'; }).length,
      newlyCounted: rows.filter(function(row) { return row.status === 'NEWLY_COUNTED'; }).length,
      missingNow: rows.filter(function(row) { return row.status === 'MISSING_NOW'; }).length,
      absoluteChange: roundAnalyticsNumber_(comparableRows.reduce(function(total, row) { return total + Math.abs(row.delta || 0); }, 0))
    },
    categories: categories,
    topIncreases: topIncreases,
    topDecreases: topDecreases,
    changes: rows.filter(function(row) { return row.status !== 'UNCHANGED'; })
      .sort(function(a, b) { return Math.abs(b.delta || 0) - Math.abs(a.delta || 0); }),
    validationIssues: (baseReport.validationIssues || []).concat(currentReport.validationIssues || [])
  };
}

function getInventoryHistoryAnalytics() {
  const sources = getFinalReviewSources();
  const rows = sources.map(function(source) {
    const report = generateInventoryReport(source.name);
    const typeTotals = { NORMAL: 0, KEG: 0, LOCATION: 0 };
    (report.items || []).forEach(function(item) {
      const value = analyticsNumberOrNull_(item.finalTotal);
      if (value !== null && Object.prototype.hasOwnProperty.call(typeTotals, item.type)) {
        typeTotals[item.type] += value;
      }
    });
    return {
      sourceSheetName: report.metadata.sourceSheetName,
      current: Boolean(source.current),
      products: Number((report.summary || {}).products || 0),
      completed: Number((report.summary || {}).completed || 0),
      missing: Number((report.summary || {}).missing || 0),
      completionPercent: report.summary && report.summary.products ? roundAnalyticsNumber_((report.summary.completed / report.summary.products) * 100) : 0,
      normalLiters: roundAnalyticsNumber_(typeTotals.NORMAL),
      kegLiters: roundAnalyticsNumber_(typeTotals.KEG),
      locationUnits: roundAnalyticsNumber_(typeTotals.LOCATION),
      validationErrors: (report.validationIssues || []).filter(function(issue) { return issue.severity === 'ERROR'; }).length
    };
  });
  return { version: CONFIG.VERSION, generatedAt: new Date().toISOString(), rows: rows };
}

function indexAnalyticsItems_(items) {
  const map = {};
  items.forEach(function(item) {
    const key = String(item.key || normalizeText(item.product || '')) + '|' + String(item.type || '');
    map[key] = item;
  });
  return map;
}

function analyticsNumberOrNull_(value) {
  if (value === '' || value === null || value === undefined) return null;
  const number = Number(value);
  return Number.isFinite(number) ? number : null;
}

function roundAnalyticsNumber_(value) {
  const number = Number(value);
  return Number.isFinite(number) ? Math.round(number * 1000) / 1000 : 0;
}