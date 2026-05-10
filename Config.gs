// Configuration globale du traitement des emails.
var PROCESSING_CONFIG = {
  LOOKBACK_DAYS: 365,
  MAX_THREADS: 100,
  MODE_TEST: true,
  LABEL_TREATED: 'Compta_Admin_Traité_V6',
  LABEL_TO_VERIFY: 'Compta_Admin_A_Verifier_V6',
  LABEL_EXCLUDED: 'Compta_Admin_Exclu_V6'
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
