const CONFIG = Object.freeze({
  EDITION: 'Production Edition',
  RELEASE_CHANNEL: 'STABLE',
  LOCATION: Object.freeze({
    ID: 'PAWILONY',
    NAME: 'PAWILONY',
    NUMBER: 1,
    NEXT_LOCATION: 'KRUCZA'
  }),
  VERSION: '5.5.1-REVIEW-FILTER-AND-GROSS-GUARD-PAWILONY',

  SHEETS: Object.freeze({
    INVENTORY: 'INWENTURA',
    DICTIONARY: 'SLOWNIK',
    REPORT: 'RAPORT',
    SETTINGS: 'USTAWIENIA',
    HISTORY: 'HISTORIA',
    TECH_LOG: 'Log techniczny',
    IMPORT_AUDIT: 'Audyt importow',
    NEW_PRODUCTS: 'Nowe produkty',
    EXPORT_HISTORY: 'Historia eksportow',
    DATA_AUDIT: 'Audyt danych',
    FORMULA_AUDIT: 'Audyt formul',
    PRODUCT_PACKAGING: 'Profile opakowan',
    HISTORY_LEGACY: 'Historia legacy'
  }),

  // Aliasów używa odporny resolver. Dzięki temu kod działa również w
  // starszych kopiach, w których zakładki miały inną wielkość liter.
  SHEET_ALIASES: Object.freeze({
    INVENTORY: Object.freeze(['INWENTURA', 'Inwentura']),
    DICTIONARY: Object.freeze(['SLOWNIK', 'Slownik', 'SŁOWNIK']),
    REPORT: Object.freeze(['RAPORT', 'Raport']),
    SETTINGS: Object.freeze(['USTAWIENIA', 'Ustawienia']),
    HISTORY: Object.freeze(['HISTORIA', 'Historia'])
  }),

  DICTIONARY: Object.freeze({
    FIRST_DATA_ROW: 2,
    ALIAS_COLUMN: 1,
    CONFIG_START_COLUMN: 4,
    CONFIG_COLUMN_COUNT: 9
  }),

  PRODUCT_TYPES: Object.freeze({
    NORMAL: 'NORMAL',
    KEG: 'KEG',
    LOCATION: 'LOCATION'
  }),

  PACKAGING_MODES: Object.freeze({
    TARE_KNOWN: 'TARE_KNOWN',
    TARE_UNKNOWN: 'TARE_UNKNOWN',
    GROSS_REFERENCE: 'GROSS_REFERENCE'
  }),

  DIRECT_FINAL_PRODUCTS: Object.freeze([
    Object.freeze({
      name: 'Czarna Fala Przelew',
      aliases: Object.freeze(['Czarna Fala Przelew', 'Czarna Fala Przelew 1 kg']),
      column: 'B',
      unit: 'szt.'
    })
  ]),

  // Jedno źródło prawdy dla fizycznego układu PAWILONÓW.
  // Kolumny formulaColumns są zawsze chronione przed importem i Product Managerem.
  INVENTORY_LAYOUT: Object.freeze({
    NORMAL: Object.freeze({
      grossWeight: 'C',
      emptyContainerWeight: 'D',
      openNet: 'E',
      prepNet: 'G',
      fullUnits: 'H',
      unitCapacity: 'I',
      fullUnitsVolume: 'J',
      finalTotal: 'K',
      formulaColumns: Object.freeze(['E', 'J', 'K']),
      unit: 'l'
    }),
    KEG: Object.freeze({
      grossWeight: 'C',
      emptyContainerWeight: 'D',
      openNet: 'E',
      prepNet: '',
      fullUnits: 'G',
      unitCapacity: 'H',
      fullUnitsVolume: 'I',
      finalTotal: 'J',
      formulaColumns: Object.freeze(['E', 'I', 'J']),
      unit: 'l'
    }),
    LOCATION: Object.freeze({
      warehouse: 'B',
      darkroom: 'C',
      fridges: 'D',
      finalTotal: 'E',
      formulaColumns: Object.freeze(['E']),
      unit: 'szt.'
    })
  }),

  LOCATION_AREAS: Object.freeze([
    Object.freeze({
      key: 'magazyn',
      label: 'Magazyn',
      columnKey: 'warehouse',
      aliases: Object.freeze(['magazyn', 'w magazynie', 'mag', 'warehouse'])
    }),
    Object.freeze({
      key: 'darkroom',
      label: 'Darkroom',
      columnKey: 'darkroom',
      aliases: Object.freeze(['darkroom', 'dark room', 'dark rum'])
    }),
    Object.freeze({
      key: 'lodowki',
      label: 'Lodówki',
      columnKey: 'fridges',
      aliases: Object.freeze([
        'lodowki', 'lodowka', 'lodówki', 'lodówka',
        'chlodnia', 'chłodnia', 'fridge', 'fridges'
      ])
    })
  ]),

  FORMULA_POLICY: Object.freeze({
    NUMERIC_TOLERANCE: 0.000000001,
    BLOCK_SETUP_ON_CONFLICT: true,
    CREATE_BACKUP_BEFORE_REPAIR: true,
    AUTOMATIC_REPAIR_ENABLED: true
  }),

  CONFIGURATION_BUILDER: Object.freeze({
    MAX_ABSOLUTE_COUNT_DROP: 10,
    MAX_RELATIVE_COUNT_DROP: 0.20
  }),

  CACHE: Object.freeze({
    CATALOG_KEY: 'inventory_pro_catalog_pawilony_v7',
    TTL_SECONDS: 21600
  }),

  PERFORMANCE: Object.freeze({
    FUZZY_SHORTLIST_SIZE: 20,
    LEVENSHTEIN_FINALISTS: 5,
    UI_RENDER_BATCH_SIZE: 60
  }),

  REVIEW: Object.freeze({
    AUTO_MERGE_DUPLICATES: true,
    PREVIOUS_CHANGE_WARNING_PERCENT: 250,
    PREVIOUS_CHANGE_WARNING_ABSOLUTE: 10
  }),

  QUALITY: Object.freeze({
    NORMAL_WHOLE_WARNING: 20,
    NORMAL_WEIGHT_WARNING: 20,
    KEG_WHOLE_WARNING: 20,
    KEG_WEIGHT_WARNING: 100,
    LOCATION_WARNING: 500
  })
});
