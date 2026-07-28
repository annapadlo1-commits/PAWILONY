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

function testFormulaRepairClassifiedPlan514_() {
  const repairSource = String(repairInventoryFormulas_);
  const planSource = String(buildInventoryFormulaRepairPlan_);
  const applySource = String(applyInventoryFormulaRepairPlan_);
  if (planSource.indexOf('classification') < 0) {
    throw new Error('Plan naprawy formuł nie zapisuje klasyfikacji każdej pozycji.');
  }
  if (
    applySource.indexOf('validateInventoryFormulaRepairPlanTargets_') < 0 ||
    applySource.indexOf('preflightInventoryFormulaRepairPlan_') < 0
  ) {
    throw new Error('Zapis formuł nie wykonuje walidacji kontraktu i kontroli preflight.');
  }
  if (
    repairSource.indexOf('createFormulaRepairBackupSheet_') < 0 ||
    repairSource.indexOf('hasBlockingConflicts') < 0
  ) {
    throw new Error('Naprawa formuł nie wymaga kopii bezpieczeństwa i klasyfikacji konfliktów.');
  }
  return true;
}
