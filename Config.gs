// Configuration globale du traitement des emails.
var PROCESSING_CONFIG = {
  SPREADSHEET_ID: '1b6xwGgOi7MVl2dJR1gQ4PcI0uM2rJQKp-abph5_GPQA',
  LOOKBACK_DAYS: 365,
  MAX_THREADS: 100,

  // Passage en production : les libellés de test ne sont plus utilisés.
  MODE_TEST: false,

  // Libellés Gmail de production.
  LABEL_TREATED: 'Compta_Admin_Traité_PROD',
  LABEL_TO_VERIFY: 'Compta_Admin_A_Verifier_PROD',
  LABEL_EXCLUDED: 'Compta_Admin_Exclu_PROD',

  // Exclusions complémentaires hors périmètre Admin / Compta.
  LABELS_TO_EXCLUDE: [
    'Compta_Admin_Exclu_PROD',
    'Groupama/Transfert'
  ]
};

// Compatibilité : certains scripts historiques utilisent `Config.*`.
var Config = {
  PROCESSING_CONFIG: PROCESSING_CONFIG,
  SPREADSHEET_ID: '1b6xwGgOi7MVl2dJR1gQ4PcI0uM2rJQKp-abph5_GPQA',
  FOLDERS: {
    SAS_ROOT: '01_SAS_BIGUEURE_BROTHERS',
    PERSO_LOCATIONS_ROOT: '02_PERSO_LOCATIONS',
    PERSO_VIE_ROOT: '03_PERSO_VIE'
  }
};

// Ouvre le classeur de pilotage cible.
function getSpreadsheet_() {
  var fileId = Config.SPREADSHEET_ID;
  return SpreadsheetApp.openById(fileId);
}
