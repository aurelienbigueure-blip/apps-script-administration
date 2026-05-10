// Configuration globale du traitement des emails.
var PROCESSING_CONFIG = {
  LOOKBACK_DAYS: 365,
  MAX_THREADS: 100,
  MODE_TEST: true,
  LABEL_TREATED: 'Compta_Admin_Traité',
  LABEL_TO_VERIFY: 'Compta_Admin_A_Verifier',
  LABEL_EXCLUDED: 'Compta_Admin_Exclu'
};

// Ouvre le classeur de pilotage cible.
function getSpreadsheet_() {
  var fileId = '1b6xwGgOi7MVl2dJR1gQ4PcI0uM2rJQKp-abph5_GPQA';
  return SpreadsheetApp.openById(fileId);
}
