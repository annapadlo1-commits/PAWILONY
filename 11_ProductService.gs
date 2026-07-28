/**
 * Inventory PRO Enterprise v2.2
 * Katalog produktow z cache.
 */

function buildProductCatalog() {
  return getCachedProductCatalog_();
}

function buildProductCatalogUncached_() {
  const configurations = loadProductConfigurations();
  const aliases = loadAliases();
  const inventoryRows = buildInventoryRowIndex_();

  const aliasesByTarget = {};

  Object.keys(aliases).forEach(alias => {
    const targetKey = normalizeText(aliases[alias]);

    if (!aliasesByTarget[targetKey]) {
      aliasesByTarget[targetKey] = [];
    }

    aliasesByTarget[targetKey].push(alias);
  });

  return configurations.map(config => ({
    name: config.name,
    normalizedName: config.normalizedName,
    type: config.type,
    category: config.category,
    columns: config.columns,
    inventoryRow: inventoryRows[config.normalizedName] || null,
    aliases: aliasesByTarget[config.normalizedName] || [],
    active: config.active
  }));
}

function buildProductCatalogIndex(catalog) {
  const products = catalog || buildProductCatalog();
  const index = {};

  products.forEach(product => {
    if (!index[product.normalizedName]) {
      index[product.normalizedName] = product;
    }
  });

  return index;
}

function buildAliasProductIndex(catalog) {
  const products = catalog || buildProductCatalog();
  const index = {};

  products.forEach(product => {
    product.aliases.forEach(alias => {
      if (!index[alias]) {
        index[alias] = product;
      }
    });
  });

  return index;
}

function buildInventoryRowIndex_() {
  const sheet = getSheetByConfiguredName_(
    CONFIG.SHEETS.INVENTORY
  );

  if (!sheet) {
    throw new Error(
      'Nie znaleziono arkusza: ' +
      CONFIG.SHEETS.INVENTORY
    );
  }

  const lastRow = sheet.getLastRow();

  if (lastRow < 1) {
    return {};
  }

  const values = sheet
    .getRange(1, 1, lastRow, 1)
    .getDisplayValues();

  const index = {};

  values.forEach((row, arrayIndex) => {
    const normalizedName = normalizeText(row[0]);

    if (normalizedName && !index[normalizedName]) {
      index[normalizedName] = arrayIndex + 1;
    }
  });

  return index;
}

function getProductCatalogSummary() {
  const catalog = buildProductCatalog();

  return {
    products: catalog.length,
    aliases: catalog.reduce(
      (sum, product) => sum + product.aliases.length,
      0
    ),
    normal: catalog.filter(
      product => product.type === CONFIG.PRODUCT_TYPES.NORMAL
    ).length,
    keg: catalog.filter(
      product => product.type === CONFIG.PRODUCT_TYPES.KEG
    ).length,
    location: catalog.filter(
      product => product.type === CONFIG.PRODUCT_TYPES.LOCATION
    ).length,
    missingInventoryRow: catalog.filter(
      product => !product.inventoryRow
    ).length
  };
}
