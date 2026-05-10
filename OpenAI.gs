var MAX_OPENAI_ACTIONS_PER_RUN = 1;
var OPENAI_MODEL = 'gpt-4o-mini';

function debugOpenAIKeyOnly() {
  var key = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (key) {
    Logger.log('OpenAI debug key: clé présente');
  } else {
    Logger.log('OpenAI debug key: clé absente');
  }
}

function debugOpenAISheetOnly() {
  var id = '';

  try {
    if (typeof Config !== 'undefined' && Config.SPREADSHEET_ID) {
      id = Config.SPREADSHEET_ID;
    }
  } catch (e) {}

  try {
    if (!id && typeof SPREADSHEET_ID !== 'undefined') {
      id = SPREADSHEET_ID;
    }
  } catch (e2) {}

  if (!id) {
    id = '1b6xwGgOi7MVl2dJR1gQ4PcI0uM2rJQKp-abph5_GPQA';
  }

  var ss = SpreadsheetApp.openById(id);
  var sheet = ss.getSheetByName('02_Actions');
  if (!sheet) {
    Logger.log('OpenAI debug sheet: onglet 02_Actions introuvable');
    return;
  }

  var values = sheet.getRange('A1:Z20').getValues();
  var headers = values[0] || [];
  var colStatut = headers.indexOf('Statut_action');
  if (colStatut < 0) {
    Logger.log('OpenAI debug sheet: colonne Statut_action introuvable dans A1:Z1');
    return;
  }

  var opened = 0;
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][colStatut] || '').trim() === 'Ouvert') {
      opened++;
    }
  }

  Logger.log('OpenAI debug sheet: lignes Ouvert dans 02_Actions!A1:Z20 = ' + opened);
}

function debugOpenAIMinimalCall() {
  var key = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!key) {
    Logger.log('OpenAI debug minimal call: OPENAI_API_KEY absente dans les propriétés du script.');
    return;
  }

  var body = {
    model: OPENAI_MODEL,
    max_tokens: 50,
    temperature: 0,
    response_format: { type: 'json_object' },
    messages: [
      { role: 'user', content: 'Réponds uniquement en JSON {"ok":true}' }
    ]
  };

  try {
    var response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
      method: 'post',
      contentType: 'application/json',
      headers: {
        Authorization: 'Bearer ' + key
      },
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    });

    var code = response.getResponseCode();
    var text = response.getContentText();
    Logger.log('OpenAI debug minimal call: HTTP statut = ' + code);
    Logger.log('OpenAI debug minimal call: réponse (300 chars) = ' + truncateText_(text, 300));

    if (code !== 200) {
      Logger.log('OpenAI debug minimal call: erreur HTTP = ' + code + ' ; contenu = ' + truncateText_(text, 300));
    }
  } catch (e) {
    Logger.log('OpenAI debug minimal call: exception = ' + String(e && e.message ? e.message : e));
  }
}

function setOpenAIKey() {
  var ui = SpreadsheetApp.getUi();
  ui.alert(
    'Configuration OpenAI',
    'Entrez la clé API OpenAI (stockée dans Script Properties sous OPENAI_API_KEY).',
    ui.ButtonSet.OK
  );

  var response = ui.prompt('OPENAI_API_KEY', 'Collez votre clé API OpenAI :', ui.ButtonSet.OK_CANCEL);
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
  ui.alert('OPENAI_API_KEY enregistrée dans les propriétés du script.');
}

function testOpenAIAnalyseUneAction() {
  Logger.log('testOpenAIAnalyseUneAction est neutralisée. Utiliser testAnalyserUneActionOuverteRapideIA().');
  return;
}

function testOpenAIAnalyseLigne2() {
  Logger.log('testOpenAIAnalyseLigne2 est neutralisée. Utiliser testAnalyserUneActionOuverteRapideIA().');
  return;
}

function testOpenAIAnalyseLigneActive() {
  Logger.log('Analyse IA ligne active: début');
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getActiveSheet();

  if (!sheet || sheet.getName() !== '02_Actions') {
    Logger.log('Sélectionner une ligne dans 02_Actions');
    return;
  }

  var activeRange = sheet.getActiveRange();
  if (!activeRange) {
    Logger.log('Aucune cellule active sélectionnée');
    return;
  }

  var rowNumber = activeRange.getRow();
  Logger.log('Analyse IA ligne active: ligne sélectionnée = ' + rowNumber);
  if (rowNumber === 1) {
    Logger.log('Sélectionner une ligne de données, pas l’en-tête');
    return;
  }

  var key = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!key) {
    throw new Error('OPENAI_API_KEY absente dans les propriétés du script.');
  }

  var aiCols = ensureAIColumnsForOpenAI_(sheet);
  var headers = sheet.getRange('A1:Z1').getValues()[0] || [];
  var row = sheet.getRange(rowNumber, 1, 1, 26).getValues()[0] || [];

  var payload = buildActiveActionPayload_(headers, row);
  var result = callOpenAIForAction_(payload);
  writeAIResultToActionRow_(sheet, rowNumber, aiCols, result);
  Logger.log('Analyse IA ligne active: écriture terminée');
}

function analyserActionsOuvertesAvecIA() {
  Logger.log('analyserActionsOuvertesAvecIA est neutralisée. Utiliser analyserActionsOuvertesRapideIA().');
  return;
}

function testAnalyserUneActionOuverteRapideIA() {
  analyserActionsOuvertesRapideIACore_(1);
}

function analyserActionsOuvertesRapideIA() {
  analyserActionsOuvertesRapideIACore_(5);
}

function analyserActionsOuvertesRapideIACore_(maxRowsToAnalyze) {
  Logger.log('début analyse IA rapide');
  var startTime = new Date().getTime();
  var key = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!key) {
    throw new Error('OPENAI_API_KEY absente dans les propriétés du script.');
  }

  var ss = getSpreadsheetRapideIA_();
  var sheet = ss.getSheetByName('02_Actions');
  if (!sheet) throw new Error('Onglet introuvable : 02_Actions');

  var aiCols = ensureAIColumnsRapideIA_(sheet);
  var values = sheet.getRange('A1:T120').getValues();
  var headers = values[0] || [];
  var analyseValues = sheet.getRange(1, aiCols.colAnalyse, 120, 1).getValues();

  var indexes = {
    domaine: getHeaderIndexRapideIA_(headers, 'Domaine'),
    action: getHeaderIndexRapideIA_(headers, 'Action'),
    priorite: getHeaderIndexRapideIA_(headers, 'Priorité'),
    statut: getHeaderIndexRapideIA_(headers, 'Statut_action'),
    lienGmail: getHeaderIndexRapideIA_(headers, 'Lien_Gmail'),
    commentaire: getHeaderIndexRapideIA_(headers, 'Commentaire'),
    derniereMaj: getHeaderIndexRapideIA_(headers, 'Dernière_maj')
  };

  var openRows = [];
  for (var i = 1; i < values.length; i++) {
    var statut = String(values[i][indexes.statut] || '').trim();
    if (statut === 'Ouvert') {
      openRows.push(i + 1);
    }
  }
  Logger.log('nombre de lignes ouvertes trouvées = ' + openRows.length);

  var analyzed = 0;
  for (var r = 0; r < openRows.length; r++) {
    if (analyzed >= maxRowsToAnalyze) break;
    if (new Date().getTime() - startTime > 90000) {
      Logger.log('Arrêt analyse IA rapide : limite 90 secondes atteinte');
      break;
    }

    var rowNumber = openRows[r];
    var rowIndex = rowNumber - 1;
    var existingAnalysis = String(analyseValues[rowIndex][0] || '').trim();
    if (existingAnalysis) continue;

    Logger.log('ligne en cours = ' + rowNumber);
    try {
      var payload = buildPayloadRapideIA_(values[rowIndex], indexes);
      var result = callOpenAIRapideIA_(payload, key);
      writeAIResultRapideIA_(sheet, rowNumber, aiCols, result);
    } catch (e) {
      var errorMessage = 'Erreur IA: ' + truncateText_(String(e && e.message ? e.message : e), 250);
      sheet.getRange(rowNumber, aiCols.colAnalyse).setValue(errorMessage);
      sheet.getRange(rowNumber, aiCols.colDate).setValue(new Date());
      Logger.log(errorMessage);
    }
    analyzed++;
  }

  Logger.log('nombre de lignes analysées = ' + analyzed);
  Logger.log('fin analyse IA rapide');
}

function getSpreadsheetRapideIA_() {
  try {
    if (typeof getSpreadsheet_ === 'function') {
      return getSpreadsheet_();
    }
  } catch (e) {}

  var id = '';
  try {
    if (typeof Config !== 'undefined' && Config.SPREADSHEET_ID) {
      id = Config.SPREADSHEET_ID;
    }
  } catch (e2) {}

  try {
    if (!id && typeof SPREADSHEET_ID !== 'undefined') {
      id = SPREADSHEET_ID;
    }
  } catch (e3) {}

  if (!id) {
    id = '1b6xwGgOi7MVl2dJR1gQ4PcI0uM2rJQKp-abph5_GPQA';
  }
  return SpreadsheetApp.openById(id);
}

function ensureAIColumnsRapideIA_(sheet) {
  var names = ['Analyse_IA', 'Décision_IA', 'Catégorie_IA', 'Action_recommandée_IA', 'Règle_proposée_IA', 'Date_analyse_IA'];
  var headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
  var cols = {};

  for (var i = 0; i < names.length; i++) {
    var idx = headers.indexOf(names[i]);
    if (idx < 0) {
      var newCol = sheet.getLastColumn() + 1;
      sheet.insertColumnAfter(sheet.getLastColumn());
      sheet.getRange(1, newCol).setValue(names[i]);
      headers.push(names[i]);
      cols[names[i]] = newCol;
    } else {
      cols[names[i]] = idx + 1;
    }
  }

  return {
    colAnalyse: cols['Analyse_IA'],
    colDecision: cols['Décision_IA'],
    colCategorie: cols['Catégorie_IA'],
    colActionReco: cols['Action_recommandée_IA'],
    colRegle: cols['Règle_proposée_IA'],
    colDate: cols['Date_analyse_IA']
  };
}

function getHeaderIndexRapideIA_(headers, name) {
  var index = headers.indexOf(name);
  if (index < 0) {
    throw new Error('Colonne introuvable dans 02_Actions : ' + name);
  }
  return index;
}

function buildPayloadRapideIA_(row, indexes) {
  return {
    domaine: String(row[indexes.domaine] || '').trim(),
    action: String(row[indexes.action] || '').trim(),
    priorite: String(row[indexes.priorite] || '').trim(),
    statut: String(row[indexes.statut] || '').trim(),
    lienGmail: String(row[indexes.lienGmail] || '').trim(),
    commentaire: String(row[indexes.commentaire] || '').trim(),
    derniereMaj: String(row[indexes.derniereMaj] || '').trim()
  };
}

function callOpenAIRapideIA_(payload, key) {
  var systemPrompt = [
    'Tu es un assistant de tri administratif et comptable.',
    'Tu aides à décider quoi faire d’une ligne 02_Actions.',
    'Tu dois être court, opérationnel et utile.',
    'Tu ne modifies jamais la décision humaine.',
    'Tu proposes seulement.',
    'Périmètre utile : facture, devis, reçu, justificatif, comptabilité, banque, impôts / TVA / URSSAF / MSA, réservation / Booking / PULSE, Colissimo / livraison utile, RDV santé / kiné, Google / Apps Script / projet IA si technique, mails Aurorebigus à analyser mail par mail.',
    'Hors périmètre probable : publicité, newsletter, offre d’emploi, candidature, marketing, mail pro Groupama transféré hors compta/admin personnel, avis produit, recommandation commerciale.',
    'Réponse JSON stricte : {"analyse":"phrase courte en français","decision":"A_TRAITER | A_IGNORER | EN_ATTENTE | A_VERIFIER","categorie":"Compta | Admin perso | Location | Santé | Technique | Marketing | Emploi | Pro transféré | Autre","action_recommandee":"quoi faire concrètement","regle_a_creer":"motif ou domaine à exclure/conserver si pertinent"}',
    'Exemples : hotmail.com générique => A_VERIFIER ; amazon.fr + commande => A_TRAITER ; amazon.fr + avis produit => A_IGNORER ; new/newsletter/publicité => A_IGNORER ; google.com + script IA/Apps Script => A_TRAITER ; offre d’emploi => A_IGNORER ; devis => A_TRAITER ; Booking/PULSE => A_TRAITER.'
  ].join('\n');

  var userPrompt = [
    'Domaine: ' + payload.domaine,
    'Action: ' + payload.action,
    'Priorité: ' + payload.priorite,
    'Statut_action: ' + payload.statut,
    'Lien_Gmail: ' + payload.lienGmail,
    'Commentaire: ' + payload.commentaire,
    'Dernière_maj: ' + payload.derniereMaj
  ].join('\n');

  var body = {
    model: OPENAI_MODEL,
    max_tokens: 250,
    temperature: 0.1,
    response_format: { type: 'json_object' },
    messages: [
      { role: 'system', content: systemPrompt },
      { role: 'user', content: userPrompt }
    ]
  };

  var response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + key
    },
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });

  var code = response.getResponseCode();
  var text = response.getContentText();
  Logger.log('statut HTTP OpenAI = ' + code);

  if (code !== 200) {
    throw new Error('Erreur OpenAI HTTP ' + code + ' : ' + truncateText_(text, 300));
  }

  var parsed = JSON.parse(text);
  var content = parsed && parsed.choices && parsed.choices[0] && parsed.choices[0].message
    ? parsed.choices[0].message.content
    : '';
  if (!content) {
    throw new Error('Réponse OpenAI vide.');
  }

  var result = JSON.parse(content);
  return {
    analyse: String(result.analyse || ''),
    decision: String(result.decision || ''),
    categorie: String(result.categorie || ''),
    action_recommandee: String(result.action_recommandee || ''),
    regle_a_creer: String(result.regle_a_creer || '')
  };
}

function writeAIResultRapideIA_(sheet, rowNumber, aiCols, result) {
  sheet.getRange(rowNumber, aiCols.colAnalyse).setValue(result.analyse || '');
  sheet.getRange(rowNumber, aiCols.colDecision).setValue(result.decision || '');
  sheet.getRange(rowNumber, aiCols.colCategorie).setValue(result.categorie || '');
  sheet.getRange(rowNumber, aiCols.colActionReco).setValue(result.action_recommandee || '');
  sheet.getRange(rowNumber, aiCols.colRegle).setValue(result.regle_a_creer || '');
  sheet.getRange(rowNumber, aiCols.colDate).setValue(new Date());
}

function ensureAIColumnsForOpenAI_(sheet) {
  var names = ['Analyse_IA', 'Décision_IA', 'Catégorie_IA', 'Action_recommandée_IA', 'Règle_proposée_IA', 'Date_analyse_IA'];
  var headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
  var cols = {};

  for (var i = 0; i < names.length; i++) {
    var idx = headers.indexOf(names[i]);
    if (idx < 0) {
      var newCol = sheet.getLastColumn() + 1;
      sheet.insertColumnAfter(sheet.getLastColumn());
      sheet.getRange(1, newCol).setValue(names[i]);
      headers.push(names[i]);
      cols[names[i]] = newCol;
    } else {
      cols[names[i]] = idx + 1;
    }
  }

  return {
    colAnalyse: cols['Analyse_IA'],
    colDecision: cols['Décision_IA'],
    colCategorie: cols['Catégorie_IA'],
    colActionReco: cols['Action_recommandée_IA'],
    colRegle: cols['Règle_proposée_IA'],
    colDate: cols['Date_analyse_IA']
  };
}

function callOpenAIForAction_(payload) {
  var key = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!key) throw new Error('OPENAI_API_KEY absente dans les propriétés du script.');

  var systemPrompt = [
    'Tu es un assistant de tri compta/admin.',
    'Tu aides l’utilisateur à décider quoi faire d’un mail dans 02_Actions.',
    'Tu ne décides pas à sa place.',
    'Tu proposes une analyse courte et opérationnelle.',
    'Périmètre utile : factures, devis, comptabilité, banque, TVA, URSSAF, impôts, MSA, PeopleDoc, Groupama Épargne, Airbnb / Wecasa / location, Booking / PULSE réservation, Colissimo, mails aurorebigus à analyser mail par mail, Google / Apps Script si technique.',
    'Hors périmètre probable : publicité, newsletter, offre d’emploi, candidature, marketing, mail pro Groupama transféré hors compta/admin perso.',
    'Répondre uniquement en JSON : {"analyse":"phrase courte en français","decision":"A_TRAITER | A_IGNORER | EN_ATTENTE | A_VERIFIER","categorie":"Compta | Admin perso | Location | Santé | Technique | Marketing | Emploi | Pro transféré | Autre","action_recommandee":"quoi faire concrètement","regle_a_creer":"motif ou domaine à exclure/conserver si pertinent"}'
  ].join('\n');

  var prompt = [
    'Domaine: ' + payload.domaine,
    'Action: ' + payload.action,
    'Priorité: ' + payload.priorite,
    'Statut_action: ' + payload.statut,
    'Lien_Gmail: ' + payload.lienGmail,
    'Commentaire: ' + payload.commentaire,
    'Dernière_maj: ' + payload.derniereMaj
  ].join('\n');

  var body = {
    model: OPENAI_MODEL,
    max_tokens: 300,
    temperature: 0.1,
    response_format: { type: 'json_object' },
    messages: [
      { role: 'system', content: systemPrompt },
      { role: 'user', content: prompt }
    ]
  };

  var response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + key
    },
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });

  var code = response.getResponseCode();
  var text = response.getContentText();
  Logger.log('OpenAI analyse: HTTP statut = ' + code);

  if (code !== 200) {
    throw new Error('Erreur OpenAI HTTP ' + code + ' : ' + truncateText_(text, 300));
  }

  var parsed = JSON.parse(text);
  var content = parsed && parsed.choices && parsed.choices[0] && parsed.choices[0].message
    ? parsed.choices[0].message.content
    : '';

  if (!content) {
    throw new Error('Réponse OpenAI vide.');
  }

  var result = JSON.parse(content);
  return {
    decision: String(result.decision || ''),
    analyse: String(result.analyse || ''),
    categorie: String(result.categorie || ''),
    action_recommandee: String(result.action_recommandee || ''),
    regle_a_creer: String(result.regle_a_creer || '')
  };
}

function writeAIResultToActionRow_(sheet, rowNumber, aiCols, result) {
  sheet.getRange(rowNumber, aiCols.colAnalyse).setValue(result.analyse || '');
  sheet.getRange(rowNumber, aiCols.colDecision).setValue(result.decision || '');
  sheet.getRange(rowNumber, aiCols.colCategorie).setValue(result.categorie || '');
  sheet.getRange(rowNumber, aiCols.colActionReco).setValue(result.action_recommandee || '');
  sheet.getRange(rowNumber, aiCols.colRegle).setValue(result.regle_a_creer || '');
  sheet.getRange(rowNumber, aiCols.colDate).setValue(new Date());
}

function buildActiveActionPayload_(headers, row) {
  return {
    domaine: getRowValueByHeader_(headers, row, 'Domaine'),
    action: getRowValueByHeader_(headers, row, 'Action'),
    priorite: getRowValueByHeader_(headers, row, 'Priorité'),
    statut: getRowValueByHeader_(headers, row, 'Statut_action'),
    lienGmail: getRowValueByHeader_(headers, row, 'Lien_Gmail'),
    commentaire: getRowValueByHeader_(headers, row, 'Commentaire'),
    derniereMaj: getRowValueByHeader_(headers, row, 'Dernière_maj')
  };
}

function getRowValueByHeader_(headers, row, headerName) {
  var col = headers.indexOf(headerName);
  if (col < 0) return '';
  return String(row[col] || '').trim();
}

function truncateText_(text, maxChars) {
  var str = String(text || '');
  if (str.length <= maxChars) return str;
  return str.substring(0, maxChars);
}

function getComptaSpreadsheetOpenAI_() {
  var id = '';

  try {
    if (typeof Config !== 'undefined' && Config.SPREADSHEET_ID) {
      id = Config.SPREADSHEET_ID;
    }
  } catch (e) {}

  try {
    if (!id && typeof SPREADSHEET_ID !== 'undefined') {
      id = SPREADSHEET_ID;
    }
  } catch (e2) {}

  if (!id) {
    id = '1b6xwGgOi7MVl2dJR1gQ4PcI0uM2rJQKp-abph5_GPQA';
  }

  return SpreadsheetApp.openById(id);
}
