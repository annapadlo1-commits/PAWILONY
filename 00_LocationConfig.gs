/** Inventory PRO 4.3 — konfiguracja pojedynczego lokalu. */
function getActiveLocationConfig() {
  return CONFIG.LOCATION;
}

function getActiveLocationId() {
  return CONFIG.LOCATION.ID;
}

function getActiveLocationName() {
  return CONFIG.LOCATION.NAME;
}

function getLocationAreaDefinitions_() {
  const configured = Array.isArray(CONFIG.LOCATION_AREAS) ? CONFIG.LOCATION_AREAS : [];
  const source = configured.length ? configured : [
    { key: 'magazyn', label: 'Magazyn', columnKey: 'warehouse', aliases: ['magazyn', 'warehouse'] },
    { key: 'darkroom', label: 'Darkroom', columnKey: 'darkroom', aliases: ['darkroom', 'dark room'] },
    { key: 'lodowki', label: 'Lodówki', columnKey: 'fridges', aliases: ['lodowki', 'lodowka', 'fridge', 'fridges'] }
  ];

  return source.map(area => ({
    key: String(area.key || '').trim(),
    label: String(area.label || area.key || '').trim(),
    columnKey: String(area.columnKey || '').trim(),
    aliases: (area.aliases || []).map(alias => String(alias || '').trim()).filter(Boolean)
  })).filter(area => area.key && area.columnKey);
}

function getLocationUiOptions_() {
  return getLocationAreaDefinitions_().map(area => ({
    value: area.key,
    label: area.label,
    columnKey: area.columnKey
  }));
}

function getLocationColumnLabelMap_() {
  const labels = {};
  getLocationAreaDefinitions_().forEach(area => { labels[area.columnKey] = area.label; });
  return labels;
}

function getNormalizedLocationCandidates_(area) {
  return [area.key, area.label]
    .concat(area.aliases || [])
    .map(candidate => normalizeText(candidate))
    .filter(Boolean);
}

/** Rozpoznawanie lokalizacji we wpisie użytkownika — dopasowanie ścisłe. */
function resolveLocationArea_(value) {
  const normalized = normalizeText(value || '');
  if (!normalized) return null;

  const definitions = getLocationAreaDefinitions_();
  for (let index = 0; index < definitions.length; index++) {
    const area = definitions[index];
    if (getNormalizedLocationCandidates_(area).indexOf(normalized) >= 0) return area;
  }
  return null;
}

/**
 * Rozpoznawanie nagłówków arkusza. Dopuszcza techniczne przyrostki typu
 * „SZT”, „SZTUKI” i „ILOŚĆ”, ale nie rozluźnia parsera wpisów użytkownika.
 */
function resolveLocationHeaderArea_(value) {
  const normalized = normalizeText(value || '')
    .replace(/\b(?:szt|sztuki|ilosc|liczba)\b/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  if (!normalized) return null;

  const definitions = getLocationAreaDefinitions_();
  for (let index = 0; index < definitions.length; index++) {
    const area = definitions[index];
    const candidates = getNormalizedLocationCandidates_(area);
    if (candidates.indexOf(normalized) >= 0) return area;
    if (candidates.some(candidate => normalized.indexOf(candidate + ' ') === 0)) return area;
  }
  return null;
}

function renderLocationOptionsHtml_() {
  return getLocationUiOptions_().map(option =>
    '<option value="' + escapeLocationHtml_(option.value) + '">' +
    escapeLocationHtml_(option.label) +
    '</option>'
  ).join('');
}

function escapeLocationHtml_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}