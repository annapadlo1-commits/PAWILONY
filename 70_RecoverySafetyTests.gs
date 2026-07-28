/**
 * Inventory PRO 4.3.6 SAFE MODE — testy zabezpieczeń niedestrukcyjnych.
 */

function testRecoveryDictionaryContaminationGuard_() {
  const fake = {
    getLastRow: function() { return 4; },
    getRange: function() {
      return {
        getDisplayValues: function() {
          return [['Amaro Lucano 1L'], ['function broken() {'], ['Campari 0,7 l']];
        }
      };
    }
  };
  const issues = detectDictionaryCodeContamination_(fake);
  if (issues.length !== 1 || issues[0].row !== 3) {
    throw new Error('Guard SLOWNIK nie wykrył kontrolnego fragmentu kodu.');
  }
  return true;
}

function testFormulaRepairHardDisabled436_() {
  let blocked = false;
  let message = '';
  try {
    repairInventoryFormulas_({ source: 'test' });
  } catch (error) {
    blocked = true;
    message = String(error && error.message || error);
  }
  if (!blocked) {
    throw new Error('Funkcja naprawy formuł nie została twardo zablokowana.');
  }
  if (message.indexOf('wyłączona') < 0 || message.indexOf('żadnego zapisu') < 0) {
    throw new Error('Blokada naprawy nie zwróciła jednoznacznego komunikatu bezpieczeństwa.');
  }
  return true;
}