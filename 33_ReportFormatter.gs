/** Inventory PRO Enterprise v2.9.0 — wspólne formatowanie wyników raportowych. */
function formatReportValue_(value, unit) {
  if (value === '' || value === null || value === undefined) return '';
  const number = Number(value);
  if (!Number.isFinite(number)) return String(value);
  return number.toLocaleString('pl-PL', { minimumFractionDigits: 0, maximumFractionDigits: 3 }) + (unit ? ' ' + unit : '');
}


function normalizeReportNumber_(value) {
  if (value === '' || value === null || value === undefined) return '';
  const n = Number(value);
  if (!Number.isFinite(n)) return value;
  return Math.round((n + Number.EPSILON) * 1000) / 1000;
}