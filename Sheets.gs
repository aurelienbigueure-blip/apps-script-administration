// Ajoute une ligne de document dans l’onglet 01_Documents.
function appendDocumentRow_(sheet, row) {
  if (!sheet) {
    Logger.log("appendDocumentRow_ ignoré : sheet vide");
    return;
  }

  if (!row || row.length === 0) {
    Logger.log("appendDocumentRow_ ignoré : ligne vide");
    return;
  }

  var hasContent = row.some(function(cell) {
    return cell !== null && cell !== undefined && String(cell).trim() !== "";
  });

  if (!hasContent) {
    Logger.log("appendDocumentRow_ ignoré : ligne sans contenu");
    return;
  }

  sheet.appendRow(row);
}

// Ajoute une ligne d’action dans l’onglet 02_Actions.
function appendActionRow_(sheet, row) {
  if (!sheet) {
    Logger.log("appendActionRow_ ignoré : sheet vide");
    return;
  }

  if (!row || row.length === 0) {
    Logger.log("appendActionRow_ ignoré : ligne vide");
    return;
  }

  var hasContent = row.some(function(cell) {
    return cell !== null && cell !== undefined && String(cell).trim() !== "";
  });

  if (!hasContent) {
    Logger.log("appendActionRow_ ignoré : ligne sans contenu");
    return;
  }

  sheet.appendRow(row);
}

// Retourne/crée un label Gmail.
function getOrCreateLabel_(labelName) {
  var label = GmailApp.getUserLabelByName(labelName);
  if (!label) {
    label = GmailApp.createLabel(labelName);
  }
  return label;
}

// Vérifie si un thread contient déjà un label.
function threadHasLabel_(thread, label) {
  var labels = thread.getLabels();
  for (var i = 0; i < labels.length; i++) {
    if (labels[i].getName() === label.getName()) {
      return true;
    }
  }
  return false;
}

// Retourne/crée la feuille des paramètres.
function getOrCreateParamSheet_() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName('00_Parametres');
  if (!sheet) {
    sheet = ss.insertSheet('00_Parametres', 0);
    sheet.appendRow(['Paramètre', 'Valeur']);
  }
  return sheet;
}

// Lit la date de dernier passage.
function getLastRunDatetime_(paramSheet) {
  var data = paramSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === 'LAST_RUN_DATETIME') {
      var val = data[i][1];
      if (val) {
        return new Date(val);
      }
    }
  }
  return null;
}

// Met à jour la date de dernier passage.
function updateLastRunDatetime_(paramSheet, date) {
  var data = paramSheet.getDataRange().getValues();
  var updated = false;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === 'LAST_RUN_DATETIME') {
      paramSheet.getRange(i + 1, 2).setValue(date);
      updated = true;
      break;
    }
  }
  if (!updated) {
    paramSheet.appendRow(['LAST_RUN_DATETIME', date]);
  }
}

// Retourne/crée la feuille 01_Documents.
function getOrCreateDocSheet_() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName('01_Documents');
  if (!sheet) {
    sheet = ss.insertSheet('01_Documents');
    sheet.appendRow([
      'Date Mail', 'Expéditeur', 'Objet', 'Nom PJ', 'Entité', 'Activité', 'Catégorie', 'Dossier cible', 'Statut', 'Confiance', 'Règle', 'Action recommandée', 'ID Message', 'ID Conversation'
    ]);
  }
  return sheet;
}

// Retourne/crée la feuille 02_Actions.
function getOrCreateActionSheet_() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName('02_Actions');
  if (!sheet) {
    sheet = ss.insertSheet('02_Actions');
    sheet.appendRow([
      'Date création', 'Source', 'Entité', 'Activité', 'Catégorie', 'Action', 'Priorité', 'ID Message', 'ID Conversation', 'Statut'
    ]);
  }
  return sheet;
}
