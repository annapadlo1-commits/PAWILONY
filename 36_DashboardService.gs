/**
 * Inventory PRO Enterprise v2.9.4 — Dashboard Suite.
 * Warstwa wyłącznie do odczytu oparta na Reporting Engine.
 */

function showInventoryDashboard() {
  const html = renderInventoryTemplate_('UI_Dashboard')
    .setWidth(1220)
    .setHeight(840);
  SpreadsheetApp.getUi().showModalDialog(html, '🍕 Inventory PRO — Dashboard inwentaryzacji');
}

function getInventoryDashboardSources() {
  return getFinalReviewSources();
}

function getInventoryDashboardData(sourceSheetName) {
  const report = generateInventoryReport(sourceSheetName);
  const summary = report.summary || {};
  const items = report.items || [];
  const validationIssues = report.validationIssues || [];

  const categories = Object.keys((report.statistics || {}).byCategory || {})
    .map(function(name) {
      const row = report.statistics.byCategory[name];
      return {
        category: name,
        products: row.products || 0,
        completed: row.completed || 0,
        missing: row.missing || 0,
        completionPercent: row.products ? Math.round((row.completed / row.products) * 1000) / 10 : 0,
        finalTotal: roundDashboardNumber_(row.finalTotal),
        unit: row.unit || ''
      };
    })
    .sort(function(a, b) { return a.category.localeCompare(b.category, 'pl'); });

  const typeStatistics = Object.keys((report.statistics || {}).byType || {}).map(function(type) {
    const row = report.statistics.byType[type];
    return {
      type: type,
      products: row.products || 0,
      completed: row.completed || 0,
      missing: row.missing || 0,
      completionPercent: row.products ? Math.round((row.completed / row.products) * 1000) / 10 : 0
    };
  });

  const missingItems = items.filter(function(item) { return !item.hasValue; }).map(function(item) {
    return { category: item.category, product: item.product, type: item.type, unit: item.unit };
  });

  const warnings = (report.warnings || []).map(function(warning) {
    return {
      product: warning.product,
      category: warning.category,
      type: warning.type,
      finalTotal: warning.finalTotal,
      unit: warning.unit,
      flags: warning.flags || []
    };
  });

  const topItems = items
    .filter(function(item) { return typeof item.finalTotal === 'number' && Number.isFinite(item.finalTotal); })
    .sort(function(a, b) { return b.finalTotal - a.finalTotal; })
    .slice(0, 12)
    .map(function(item) {
      return {
        category: item.category,
        product: item.product,
        type: item.type,
        finalTotal: roundDashboardNumber_(item.finalTotal),
        unit: item.unit
      };
    });

  const completed = Number(summary.completed || 0);
  const products = Number(summary.products || 0);

  return {
    version: CONFIG.VERSION,
    sourceSheetName: report.metadata.sourceSheetName,
    generatedAt: new Date(report.metadata.generatedAt).toISOString(),
    generatedBy: report.metadata.generatedBy || '',
    durationMs: report.metadata.durationMs || 0,
    summary: {
      products: products,
      completed: completed,
      missing: Number(summary.missing || 0),
      warningProducts: Number(summary.warningProducts || 0),
      validationErrors: validationIssues.filter(function(issue) { return issue.severity === 'ERROR'; }).length,
      completionPercent: products ? Math.round((completed / products) * 1000) / 10 : 0,
      qualityScore: calculateDashboardQualityScore_(report)
    },
    categories: categories,
    typeStatistics: typeStatistics,
    missingItems: missingItems,
    warnings: warnings,
    validationIssues: validationIssues,
    topItems: topItems
  };
}

function calculateDashboardQualityScore_(report) {
  const items = report.items || [];
  if (!items.length) return 0;
  const missing = items.filter(function(item) { return !item.hasValue; }).length;
  const warnings = items.filter(function(item) {
    return (item.flags || []).some(function(flag) { return flag !== 'BRAK WARTOŚCI'; });
  }).length;
  const errors = (report.validationIssues || []).filter(function(issue) { return issue.severity === 'ERROR'; }).length;
  const penalty = missing + (warnings * 0.5) + (errors * 3);
  return Math.max(0, Math.round((100 - (penalty / items.length) * 100) * 10) / 10);
}

function roundDashboardNumber_(value) {
  const number = Number(value);
  return Number.isFinite(number) ? Math.round(number * 1000) / 1000 : 0;
}