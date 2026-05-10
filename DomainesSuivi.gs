var TRIAGE_IA_MAX_SOURCE_ROWS = 300;
var TRIAGE_IA_MAX_OUTPUT_ROWS = 150;
var DOMAINES_SUIVI_SHEET_NAME = '12_Domaines_Suivi';
var TRIAGE_IA_SHEET_NAME = '14_Triage_IA';

function workflowApresAnalyseIARapide() {
  Logger.log('début workflow');
  setupPerformanceComptaAdmin();
  Logger.log('performance OK');
  creerVueTriageIALegere();
  Logger.log('vue triage IA OK');
  majDomainesSuiviDepuisActionsIA();
  Logger.log('domaines suivi mis à jour OK');
  Logger.log('fin workflow');
}

function creerVueTriageIALegere() {
  Logger.log('creerVueTriageIALegere : début');
  var ss = getComptaSpreadsheetMiseEnForme_();
  var source = ss.getSheetByName('02_Actions');
  if (!source) throw new Error('Onglet introuvable : 02_Actions');

  var values = source.getRange(1, 1, TRIAGE_IA_MAX_SOURCE_ROWS, 26).getValues();
  var headers = values[0] || [];
  var indexes = getActionColumnIndexesForIA_(headers);

  var outputHeaders = [
    'Ligne_source',
    'Date_creation',
    'Domaine',
    'Action',
    'Priorité',
    'Statut_action',
    'Commentaire_utilisateur',
    'Décision_IA_proposée',
    'Commentaire_IA',
    'Règle_correction_proposée',
    'Analyse_IA',
    'Lien_Gmail'
  ];
  var output = [outputHeaders];

  for (var i = 1; i < values.length && output.length <= TRIAGE_IA_MAX_OUTPUT_ROWS; i++) {
    var row = values[i];
    if (rowIsEmptyDomainesSuivi_(row)) continue;

    var statut = getCellByIndex_(row, indexes.statutAction);
    var decisionIa = getCellByIndex_(row, indexes.decisionIaProposee);
    var commentaireIa = getCellByIndex_(row, indexes.commentaireIa);
    var regleIa = getCellByIndex_(row, indexes.regleCorrectionProposee);

    if (statut === 'Ouvert' || decisionIa || commentaireIa || regleIa) {
      output.push([
        i + 1,
        getCellByIndex_(row, indexes.dateCreation),
        getCellByIndex_(row, indexes.domaine),
        getCellByIndex_(row, indexes.action),
        getCellByIndex_(row, indexes.priorite),
        statut,
        getCellByIndex_(row, indexes.commentaire),
        decisionIa,
        commentaireIa,
        regleIa,
        getCellByIndex_(row, indexes.analyseIa),
        getCellByIndex_(row, indexes.lienGmail)
      ]);
    }
  }

  var sheet = ss.getSheetByName(TRIAGE_IA_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(TRIAGE_IA_SHEET_NAME);
    Logger.log('creerVueTriageIALegere : onglet créé');
  } else {
    Logger.log('creerVueTriageIALegere : onglet existant réutilisé');
  }

  var clearRows = Math.max(sheet.getMaxRows(), output.length, 2);
  var clearCols = Math.max(sheet.getMaxColumns(), outputHeaders.length);
  sheet.getRange(1, 1, Math.min(clearRows, 500), Math.min(clearCols, 20)).clearContent().clearFormat();

  sheet.getRange(1, 1, output.length, outputHeaders.length).setValues(output);
  formatTriageIALegere_(sheet, output.length, outputHeaders.length);

  Logger.log('creerVueTriageIALegere : lignes copiées = ' + (output.length - 1));
  Logger.log('creerVueTriageIALegere : fin');
}

function majDomainesSuiviDepuisActionsIA() {
  Logger.log('majDomainesSuiviDepuisActionsIA : début');
  var ss = getComptaSpreadsheetMiseEnForme_();
  var actions = ss.getSheetByName('02_Actions');
  if (!actions) throw new Error('Onglet introuvable : 02_Actions');

  var actionValues = actions.getRange(1, 1, TRIAGE_IA_MAX_SOURCE_ROWS, 26).getValues();
  var actionHeaders = actionValues[0] || [];
  var actionIndexes = getActionColumnIndexesForIA_(actionHeaders);

  var suivi = ss.getSheetByName(DOMAINES_SUIVI_SHEET_NAME);
  if (!suivi) {
    suivi = ss.insertSheet(DOMAINES_SUIVI_SHEET_NAME);
    Logger.log('majDomainesSuiviDepuisActionsIA : onglet 12_Domaines_Suivi créé');
  }

  var suiviColumns = ensureDomainesSuiviHeaders_(suivi);
  var existingKeys = getExistingDomainesSuiviKeys_(suivi, suiviColumns);
  var rowsToAdd = [];

  addSpecificDomainesSuiviRules_(rowsToAdd, existingKeys, suiviColumns);

  for (var i = 1; i < actionValues.length; i++) {
    var row = actionValues[i];
    if (rowIsEmptyDomainesSuivi_(row)) continue;

    var domaine = getCellByIndex_(row, actionIndexes.domaine);
    var statut = getCellByIndex_(row, actionIndexes.statutAction);
    var commentaire = getCellByIndex_(row, actionIndexes.commentaire);
    var decisionIa = getCellByIndex_(row, actionIndexes.decisionIaProposee);
    var commentaireIa = getCellByIndex_(row, actionIndexes.commentaireIa);
    var regleIa = getCellByIndex_(row, actionIndexes.regleCorrectionProposee);
    var analyseIa = getCellByIndex_(row, actionIndexes.analyseIa);
    var combinedIa = [commentaire, decisionIa, commentaireIa, regleIa, analyseIa].join(' ');

    if (statut === 'À ignorer' && domaine) {
      addDomainesSuiviRowIfMissing_(rowsToAdd, existingKeys, suiviColumns, {
        domaineOuMotif: domaine,
        type: 'Domaine',
        decision: 'Exclure',
        regleCible: '',
        statutRegle: 'À ajouter Classification.gs',
        priorite: 'Moyenne',
        source: '02_Actions',
        commentaire: joinNonEmptyDomainesSuivi_([commentaire, commentaireIa, regleIa], ' | ')
      });
    }

    if (statut === 'Traité' && domaine && containsUsefulSignalDomainesSuivi_(combinedIa)) {
      addDomainesSuiviRowIfMissing_(rowsToAdd, existingKeys, suiviColumns, {
        domaineOuMotif: domaine,
        type: 'Domaine',
        decision: 'Conserver',
        regleCible: '',
        statutRegle: 'À ajouter Classification.gs',
        priorite: 'Haute',
        source: '02_Actions',
        commentaire: joinNonEmptyDomainesSuivi_([commentaire, commentaireIa, regleIa, analyseIa], ' | ')
      });
    }
  }

  if (rowsToAdd.length > 0) {
    var startRow = Math.max(suivi.getLastRow(), 1) + 1;
    suivi.getRange(startRow, 1, rowsToAdd.length, suiviColumns.headers.length).setValues(rowsToAdd);
    Logger.log('majDomainesSuiviDepuisActionsIA : lignes ajoutées = ' + rowsToAdd.length);
  } else {
    Logger.log('majDomainesSuiviDepuisActionsIA : aucune nouvelle ligne à ajouter');
  }

  formatDomainesSuiviLeger_(suivi, suiviColumns.headers.length);
  Logger.log('majDomainesSuiviDepuisActionsIA : fin');
}

function getActionColumnIndexesForIA_(headers) {
  return {
    dateCreation: findHeaderIndexWithFallback_(headers, ['Date_creation', 'Date création'], 2),
    domaine: findHeaderIndexWithFallback_(headers, ['Domaine'], 5),
    action: findHeaderIndexWithFallback_(headers, ['Action'], 8),
    priorite: findHeaderIndexWithFallback_(headers, ['Priorité', 'Priorite'], 9),
    statutAction: findHeaderIndexWithFallback_(headers, ['Statut_action', 'Statut'], 12),
    lienGmail: findHeaderIndexWithFallback_(headers, ['Lien_Gmail'], 13),
    commentaire: findHeaderIndexWithFallback_(headers, ['Commentaire', 'Commentaire_utilisateur'], 19),
    explicationAction: findHeaderIndexWithFallback_(headers, ['Explication_action'], 21),
    pourquoiRemonte: findHeaderIndexWithFallback_(headers, ['Pourquoi_remonté', 'Pourquoi_remonte', 'Catégorie_IA'], 22),
    decisionIaProposee: findHeaderIndexWithFallback_(headers, ['Décision_IA_proposée', 'Decision_IA_proposee', 'Décision_IA', 'Decision_IA'], 23),
    commentaireIa: findHeaderIndexWithFallback_(headers, ['Commentaire_IA', 'Action_recommandée_IA', 'Action_recommandee_IA'], 24),
    regleCorrectionProposee: findHeaderIndexWithFallback_(headers, ['Règle_correction_proposée', 'Regle_correction_proposee', 'Règle_proposée_IA', 'Regle_proposee_IA'], 25),
    analyseIa: findHeaderIndexWithFallback_(headers, ['Analyse_IA'], 26)
  };
}

function ensureDomainesSuiviHeaders_(sheet) {
  var minimumHeaders = [
    'Date_ajout',
    'Domaine_ou_motif',
    'Type',
    'Décision',
    'Règle_cible',
    'Statut_règle',
    'Priorité',
    'Source',
    'Commentaire',
    'Dernière_maj'
  ];
  var lastCol = Math.max(sheet.getLastColumn(), minimumHeaders.length);
  var currentHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  if (rowIsEmptyDomainesSuivi_(currentHeaders)) {
    currentHeaders = minimumHeaders.slice();
    sheet.getRange(1, 1, 1, currentHeaders.length).setValues([currentHeaders]);
  } else {
    for (var i = 0; i < minimumHeaders.length; i++) {
      if (currentHeaders.indexOf(minimumHeaders[i]) === -1) {
        currentHeaders.push(minimumHeaders[i]);
        sheet.getRange(1, currentHeaders.length).setValue(minimumHeaders[i]);
      }
    }
  }

  var columns = { headers: currentHeaders };
  for (var h = 0; h < currentHeaders.length; h++) {
    columns[currentHeaders[h]] = h + 1;
  }
  return columns;
}

function getExistingDomainesSuiviKeys_(sheet, columns) {
  var existing = {};
  var lastRow = Math.min(Math.max(sheet.getLastRow(), 1), 1000);
  if (lastRow < 2) return existing;

  var values = sheet.getRange(2, 1, lastRow - 1, columns.headers.length).getValues();
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var key = buildDomainesSuiviKey_(
      row[columns['Domaine_ou_motif'] - 1],
      row[columns['Décision'] - 1],
      row[columns['Règle_cible'] - 1]
    );
    if (key !== '||') existing[key] = true;
  }
  return existing;
}

function addSpecificDomainesSuiviRules_(rowsToAdd, existingKeys, columns) {
  var rules = [
    {
      domaineOuMotif: 'hotmail.com',
      type: 'Domaine',
      decision: 'Exclure conditionnellement',
      regleCible: 'EXCL_HOTMAIL_MARKETING_JOB_TRAVEL',
      priorite: 'Moyenne',
      commentaire: 'Exclure seulement si publicité, newsletter, promotion, voyage marketing, offre d’emploi, candidature ou recrutement. Ne pas exclure si Wecasa, URSSAF, MSA, Booking, PULSE, devis, facture, reçu, comptable, RDV santé / kiné.'
    },
    {
      domaineOuMotif: 'boxraw.com',
      type: 'Domaine',
      decision: 'Exclure',
      regleCible: 'EXCL_MARKETING_BRANDS',
      priorite: 'Moyenne',
      commentaire: 'Domaine marketing / marque commerciale.'
    },
    {
      domaineOuMotif: 'alerte + annonce immobilière',
      type: 'Motif',
      decision: 'Exclure',
      regleCible: 'EXCL_ALERTS_IMMO',
      priorite: 'Moyenne',
      commentaire: 'Alertes annonces immobilières à exclure sauf document contractuel ou administratif.'
    },
    {
      domaineOuMotif: 'ur + wecasa',
      type: 'Motif',
      decision: 'Conserver',
      regleCible: 'KEEP_WECASA_PAYMENT',
      priorite: 'Haute',
      commentaire: 'Mails Wecasa avec montant détecté ; à conserver comme justificatif / suivi administratif.'
    },
    {
      domaineOuMotif: 'google.com + IA / Apps Script',
      type: 'Motif',
      decision: 'Conserver conditionnellement',
      regleCible: 'KEEP_GOOGLE_APPS_SCRIPT_TECH',
      priorite: 'Haute',
      commentaire: 'Conserver les alertes techniques liées à Apps Script / IA ; exclure les notifications Google génériques.'
    },
    {
      domaineOuMotif: 'amazon.fr',
      type: 'Domaine',
      decision: 'Exclure conditionnellement',
      regleCible: 'EXCL_AMAZON_MARKETING_AVIS',
      priorite: 'Moyenne',
      commentaire: 'Exclure les avis produit, recommandations et publicités. Conserver factures, commandes utiles, reçus, justificatifs.'
    },
    {
      domaineOuMotif: 'new',
      type: 'Motif',
      decision: 'Exclure conditionnellement',
      regleCible: 'EXCL_NEW_MARKETING',
      priorite: 'Moyenne',
      commentaire: 'Exclure si publicité, promotion, newsletter.'
    },
    {
      domaineOuMotif: 'notif-coli / Colissimo',
      type: 'Motif',
      decision: 'Conserver conditionnellement',
      regleCible: 'KEEP_COLISSIMO_USEFUL',
      priorite: 'Haute',
      commentaire: 'Conserver si livraison utile / justificatif ; ignorer si notification sans intérêt administratif.'
    }
  ];

  for (var i = 0; i < rules.length; i++) {
    rules[i].statutRegle = 'À ajouter Classification.gs';
    rules[i].source = 'Analyse IA récente';
    addDomainesSuiviRowIfMissing_(rowsToAdd, existingKeys, columns, rules[i]);
  }
}

function addDomainesSuiviRowIfMissing_(rowsToAdd, existingKeys, columns, rule) {
  var key = buildDomainesSuiviKey_(rule.domaineOuMotif, rule.decision, rule.regleCible);
  if (existingKeys[key]) return;

  var now = new Date();
  var row = [];
  for (var i = 0; i < columns.headers.length; i++) row.push('');

  setDomainesSuiviValue_(row, columns, 'Date_ajout', now);
  setDomainesSuiviValue_(row, columns, 'Domaine_ou_motif', rule.domaineOuMotif || '');
  setDomainesSuiviValue_(row, columns, 'Type', rule.type || 'Domaine');
  setDomainesSuiviValue_(row, columns, 'Décision', rule.decision || '');
  setDomainesSuiviValue_(row, columns, 'Règle_cible', rule.regleCible || '');
  setDomainesSuiviValue_(row, columns, 'Statut_règle', rule.statutRegle || 'À ajouter Classification.gs');
  setDomainesSuiviValue_(row, columns, 'Priorité', rule.priorite || 'Moyenne');
  setDomainesSuiviValue_(row, columns, 'Source', rule.source || '02_Actions');
  setDomainesSuiviValue_(row, columns, 'Commentaire', rule.commentaire || '');
  setDomainesSuiviValue_(row, columns, 'Dernière_maj', now);

  rowsToAdd.push(row);
  existingKeys[key] = true;
}

function setDomainesSuiviValue_(row, columns, header, value) {
  if (columns[header]) row[columns[header] - 1] = value;
}

function buildDomainesSuiviKey_(domaineOuMotif, decision, regleCible) {
  return [domaineOuMotif, decision, regleCible].map(function(value) {
    return normalizeDomainesSuiviText_(value);
  }).join('|');
}

function containsUsefulSignalDomainesSuivi_(text) {
  var normalized = normalizeDomainesSuiviText_(text);
  var signals = [
    'facture',
    'devis',
    'compta',
    'comptable',
    'wecasa',
    'urssaf',
    'msa',
    'colissimo',
    'booking',
    'pulse',
    'google apps script',
    'apps script',
    'projet ia',
    'rdv kine',
    'rdv kiné'
  ];
  for (var i = 0; i < signals.length; i++) {
    if (normalized.indexOf(normalizeDomainesSuiviText_(signals[i])) !== -1) return true;
  }
  return false;
}

function formatTriageIALegere_(sheet, lastRow, lastCol) {
  Logger.log('formatTriageIALegere_ : début');
  if (sheet.getFilter()) sheet.getFilter().remove();
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, lastCol)
    .setBackground('#D9EAF7')
    .setFontWeight('bold')
    .setFontColor('#1F2937')
    .setWrap(true);
  sheet.getRange(1, 1, Math.max(lastRow, 2), lastCol)
    .setFontFamily('Arial')
    .setFontSize(10)
    .setVerticalAlignment('top');
  sheet.getRange(1, 1, Math.max(lastRow, 1), lastCol).createFilter();

  var widths = [90, 140, 180, 420, 100, 130, 320, 170, 360, 360, 420, 180];
  for (var i = 0; i < widths.length; i++) {
    sheet.setColumnWidth(i + 1, widths[i]);
  }
  if (lastRow > 1) {
    sheet.getRange(2, 4, lastRow - 1, 1).setWrap(true);
    sheet.getRange(2, 9, lastRow - 1, 1).setWrap(true);
    sheet.getRange(2, 11, lastRow - 1, 1).setWrap(true);
  }
  Logger.log('formatTriageIALegere_ : fin');
}

function formatDomainesSuiviLeger_(sheet, lastCol) {
  var lastRow = Math.min(Math.max(sheet.getLastRow(), 2), 300);
  if (sheet.getFilter()) sheet.getFilter().remove();
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, lastCol)
    .setBackground('#E6D9F2')
    .setFontWeight('bold')
    .setFontColor('#1F2937')
    .setWrap(true);
  sheet.getRange(1, 1, lastRow, lastCol)
    .setFontFamily('Arial')
    .setFontSize(10)
    .setVerticalAlignment('top')
    .setWrap(true);
  sheet.getRange(1, 1, lastRow, lastCol).createFilter();
  sheet.setColumnWidth(1, 140);
  sheet.setColumnWidth(2, 240);
  sheet.setColumnWidth(3, 110);
  sheet.setColumnWidth(4, 190);
  sheet.setColumnWidth(5, 250);
  sheet.setColumnWidth(6, 190);
  sheet.setColumnWidth(7, 110);
  sheet.setColumnWidth(8, 150);
  sheet.setColumnWidth(9, 520);
  sheet.setColumnWidth(10, 140);
}

function findHeaderIndexWithFallback_(headers, names, fallbackOneBased) {
  for (var i = 0; i < names.length; i++) {
    var index = headers.indexOf(names[i]);
    if (index >= 0) return index;
  }
  return fallbackOneBased - 1;
}

function getCellByIndex_(row, index) {
  if (index < 0 || index >= row.length) return '';
  var value = row[index];
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

function rowIsEmptyDomainesSuivi_(row) {
  for (var i = 0; i < row.length; i++) {
    if (row[i] !== '' && row[i] !== null && row[i] !== undefined) return false;
  }
  return true;
}

function joinNonEmptyDomainesSuivi_(values, separator) {
  var nonEmpty = [];
  for (var i = 0; i < values.length; i++) {
    var value = String(values[i] || '').trim();
    if (value) nonEmpty.push(value);
  }
  return nonEmpty.join(separator);
}

function normalizeDomainesSuiviText_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/[àáâãäå]/g, 'a')
    .replace(/[èéêë]/g, 'e')
    .replace(/[ìíîï]/g, 'i')
    .replace(/[òóôõö]/g, 'o')
    .replace(/[ùúûü]/g, 'u')
    .replace(/[ç]/g, 'c')
    .replace(/\s+/g, ' ')
    .trim();
}
