/**
 * OpenAI.gs
 * Fichier volontairement réduit.
 *
 * Les fonctions IA de production sont dans AnalyseIA.gs.
 * Ne pas recréer ici :
 * - analyserActionsOuvertesRapideIA
 * - analyserActionsOuvertesAvecIA
 * - testOpenAIAnalyseUneAction
 * - testOpenAIAnalyseLigne2
 * - testOpenAIAnalyseLigneActive
 * - debugOpenAI*
 */

var OPENAI_MODEL = 'gpt-4o-mini';

/**
 * Enregistre la clé OpenAI dans les propriétés du script.
 * À utiliser uniquement si la clé doit être remplacée.
 */
function setOpenAIKey() {
  var ui = SpreadsheetApp.getUi();

  var response = ui.prompt(
    'OPENAI_API_KEY',
    'Collez votre clé API OpenAI :',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    ui.alert('Configuration annulée.');
    return;
  }

  var key = String(response.getResponseText() || '').trim();

  if (!key) {
    ui.alert('Aucune clé fournie.');
    return;
  }

  PropertiesService.getScriptProperties().setProperty('OPENAI_API_KEY', key);
  ui.alert('OPENAI_API_KEY enregistrée.');
}

/**
 * Récupère la clé OpenAI.
 */
function getOpenAIKey_() {
  return PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY') || '';
}

/**
 * Tronque un texte pour éviter les cellules trop lourdes.
 */
function truncateText_(text, maxChars) {
  var str = String(text || '');
  if (str.length <= maxChars) return str;
  return str.substring(0, maxChars);
}
