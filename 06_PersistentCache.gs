/**
 * Inventory PRO Enterprise v2.2
 * Skompresowany cache katalogu produktow.
 */

function getCachedProductCatalog_() {
  const cache = CacheService.getDocumentCache();
  const key = CONFIG.CACHE.CATALOG_KEY;

  try {
    const encoded = cache.get(key);

    if (encoded) {
      const bytes = Utilities.base64Decode(encoded);
      const blob = Utilities.newBlob(bytes, 'application/gzip', 'inventory-pro-catalog.gz');
      const json = Utilities.ungzip(blob).getDataAsString('UTF-8');
      const catalog = JSON.parse(json);

      if (Array.isArray(catalog) && catalog.length) {
        return catalog;
      }
    }
  } catch (error) {
    logWarning(
      'PersistentCache',
      'getCachedProductCatalog_',
      'Nie udalo sie odczytac cache. Katalog zostanie odbudowany.',
      { message: normalizeError_(error).message }
    );
  }

  const catalog = buildProductCatalogUncached_();
  putProductCatalogCache_(catalog);

  return catalog;
}

function putProductCatalogCache_(catalog) {
  if (!Array.isArray(catalog)) return;

  try {
    const json = JSON.stringify(catalog);
    const zipped = Utilities.gzip(
      Utilities.newBlob(json, 'application/json', 'inventory-pro-catalog.json')
    );
    const encoded = Utilities.base64Encode(
      zipped.getBytes()
    );

    CacheService.getDocumentCache().put(
      CONFIG.CACHE.CATALOG_KEY,
      encoded,
      CONFIG.CACHE.TTL_SECONDS
    );
  } catch (error) {
    logWarning(
      'PersistentCache',
      'putProductCatalogCache_',
      'Nie udalo sie zapisac cache katalogu.',
      { message: normalizeError_(error).message }
    );
  }
}

function invalidateProductCatalogCache_() {
  CacheService.getDocumentCache().remove(
    CONFIG.CACHE.CATALOG_KEY
  );
  if (typeof invalidateRuntimeSnapshot_ === 'function') {
    invalidateRuntimeSnapshot_();
  }
}

function clearProductCatalogCache() {
  invalidateProductCatalogCache_();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Cache katalogu zostal wyczyszczony.',
    'Inventory PRO',
    6
  );
}