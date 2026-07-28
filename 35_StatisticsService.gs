/** Inventory PRO Enterprise v2.9.0 — statystyki wspólne dla przyszłych dashboardów. */
function buildInventoryStatistics_(items) {
  const byCategory = {};
  const byType = {};
  items.forEach(item => {
    if (!byCategory[item.category]) byCategory[item.category] = { products: 0, completed: 0, missing: 0, finalTotal: 0, unit: item.unit };
    const category = byCategory[item.category];
    category.products++;
    if (item.hasValue) category.completed++; else category.missing++;
    if (typeof item.finalTotal === 'number' && Number.isFinite(item.finalTotal)) category.finalTotal += item.finalTotal;

    if (!byType[item.type]) byType[item.type] = { products: 0, completed: 0, missing: 0 };
    byType[item.type].products++;
    if (item.hasValue) byType[item.type].completed++; else byType[item.type].missing++;
  });
  return { byCategory: byCategory, byType: byType };
}