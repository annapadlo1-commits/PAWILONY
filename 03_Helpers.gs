/**
 * Inventory PRO Enterprise v2.1
 * Wspolne funkcje pomocnicze.
 */

function normalizeText(value) {
  const polishMap = {
    'ą': 'a',
    'ć': 'c',
    'ę': 'e',
    'ł': 'l',
    'ń': 'n',
    'ó': 'o',
    'ś': 's',
    'ż': 'z',
    'ź': 'z'
  };

  return String(value === null || value === undefined ? '' : value)
    .trim()
    .toLowerCase()
    .replace(/[ąćęłńóśżź]/g, char => polishMap[char] || char)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[.,;:()[\]{}'"`´’“”\-_/\\]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function normalizeColumnLetter_(value) {
  return String(value || '')
    .trim()
    .toUpperCase()
    .replace(/[^A-Z]/g, '');
}

function createUniqueId_(prefix) {
  const safePrefix = String(prefix || 'ID').toUpperCase();
  return safePrefix + '-' + Utilities.getUuid();
}
