function setupMiseEnFormeComptaAdmin() {
  var ss = getComptaSpreadsheetMiseEnForme_();

  formatOngletDocuments_(ss);
  formatOngletActions_(ss);
  formatOngletParametres_(ss);
  formatOngletsBanque_(ss);
  rebuildSyntheseComptaAdmin_();
  auditColonnesComptaAdmin();

  SpreadsheetApp.flush();
  Logger.log("Mise en forme + audit terminés.");
}

function auditColonnesComptaAdmin() {
  var ss = getComptaSpreadsheetMiseEnForme_();
  var audit = getOrCreateSheetMiseEnForme_(ss, "11_Audit_Colonnes");

  audit.clear();

  var headers = [
    "Date_audit",
    "Onglet",
    "Ligne",
    "Niveau",
    "Problème",
    "Détail",
    "Action_recommandée"
  ];

  audit.getRange(1, 1, 1, headers.length).setValues([headers]);

  var rows = [];
  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

  rows = rows.concat(auditDocuments_(ss, now));
  rows = rows.concat(auditActions_(ss, now));

  if (rows.length > 0) {
    audit.getRange(2, 1, rows.length, headers.length).setValues(rows);
  } else {
    audit.getRange(2, 1, 1, headers.length).setValues([[
      now,
      "GLOBAL",
      "",
      "OK",
      "Aucun problème bloquant détecté",
      "Les colonnes principales semblent cohérentes.",
      "Continuer les tests quotidiens."
    ]]);
  }

  formatAuditSheet_(audit);
  Logger.log("Audit colonnes terminé. Problèmes détectés : " + rows.length);
}

function rebuildSyntheseComptaAdmin_() {
  var ss = getComptaSpreadsheetMiseEnForme_();
  var sheet = getOrCreateSheetMiseEnForme_(ss, "04_Synthese");

  sheet.clear();

  var values = [
    ["Pilotage Compta Admin", "", "", "", "", ""],
    ["Dernière mise à jour", new Date(), "", "", "", ""],
    ["", "", "", "", "", ""],
    ["Indicateur", "Valeur", "Commentaire", "", "", ""],
    ["Documents détectés", '=COUNTA(\'01_Documents\'!A2:A)', "Nombre de lignes dans 01_Documents", "", "", ""],
    ["Actions ouvertes", '=COUNTIF(\'02_Actions\'!L:L,"Ouvert")', "Actions à traiter", "", "", ""],
    ["Documents à vérifier", '=COUNTIF(\'01_Documents\'!Q:Q,"À vérifier")', "Documents ambigus", "", "", ""],
    ["Documents à extraire", '=COUNTIF(\'01_Documents\'!Q:Q,"À extraire")', "Pièces jointes à classer plus tard", "", "", ""],
    ["Documents détectés", '=COUNTIF(\'01_Documents\'!Q:Q,"Détecté")', "Mails utiles sans extraction immédiate", "", "", ""],
    ["Actions en attente", '=COUNTIF(\'02_Actions\'!L:L,"En attente")', "Actions à suivre", "", "", ""],
    ["", "", "", "", "", ""],
    ["Répartition par entité", "", "", "", "", ""],
    ["SAS", '=COUNTIF(\'01_Documents\'!L:L,"SAS")', "", "", "", ""],
    ["Perso_Locations", '=COUNTIF(\'01_Documents\'!L:L,"Perso_Locations")', "", "", "", ""],
    ["Perso_Vie", '=COUNTIF(\'01_Documents\'!L:L,"Perso_Vie")', "", "", "", ""],
    ["À vérifier", '=COUNTIF(\'01_Documents\'!L:L,"À vérifier")', "", "", "", ""],
    ["", "", "", "", "", ""],
    ["Routine quotidienne", "", "", "", "", ""],
    ["1", "Ouvrir 01_Documents", "Contrôler les nouveaux mails détectés", "", "", ""],
    ["2", "Ouvrir 02_Actions", "Traiter les actions ouvertes", "", "", ""],
    ["3", "Repérer les faux positifs", "Les signaler pour améliorer les règles", "", "", ""],
    ["4", "Garder MODE_TEST à true", "Pas d’extraction Drive tant que la détection n’est pas stable", "", "", ""]
  ];

  sheet.getRange(1, 1, values.length, 6).setValues(values);

  sheet.getRange("A1:F1").merge();
  sheet.getRange("A1").setFontSize(18).setFontWeight("bold").setFontColor("#1F2937");
  sheet.getRange("A1:F1").setBackground("#EAF2F8");
  sheet.getRange("A4:C4").setFontWeight("bold").setBackground("#D9EAF7");
  sheet.getRange("A12:C12").setFontWeight("bold").setBackground("#D9EAF7");
  sheet.getRange("A18:C18").setFontWeight("bold").setBackground("#D9EAF7");

  sheet.setFrozenRows(4);
  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 220);
  sheet.setColumnWidth(3, 420);
  sheet.setColumnWidth(4, 40);
  sheet.setColumnWidth(5, 40);
  sheet.setColumnWidth(6, 40);

  sheet.getRange("A1:F22").setVerticalAlignment("middle");
  sheet.getRange("A1:F22").setWrap(true);
  sheet.getRange("B2").setNumberFormat("yyyy-mm-dd hh:mm");
}

function formatOngletDocuments_(ss) {
  var sheet = ss.getSheetByName("01_Documents");
  if (!sheet) return;

  var lastCol = Math.max(sheet.getLastColumn(), 27);
  var lastRow = Math.max(sheet.getLastRow(), 2);

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, lastCol)
    .setFontWeight("bold")
    .setFontColor("#1F2937")
    .setBackground("#D9EAF7")
    .setVerticalAlignment("middle")
    .setWrap(true);

  sheet.getRange(1, 1, lastRow, lastCol).setFontFamily("Arial").setFontSize(10);
  sheet.getRange(2, 1, Math.max(lastRow - 1, 1), lastCol).setVerticalAlignment("top");

  safeCreateFilter_(sheet);

  var widths = {
    1: 160,
    2: 145,
    3: 145,
    4: 90,
    5: 240,
    6: 240,
    7: 320,
    8: 420,
    9: 260,
    10: 150,
    11: 120,
    12: 130,
    13: 160,
    14: 170,
    15: 300,
    16: 180,
    17: 130,
    18: 110,
    19: 180,
    20: 180,
    21: 320,
    22: 220,
    23: 220,
    24: 150,
    25: 130,
    26: 250,
    27: 145
  };

  applyWidths_(sheet, widths);

  sheet.getRange(2, 8, Math.max(lastRow - 1, 1), 1).setWrap(true);
  sheet.getRange(2, 21, Math.max(lastRow - 1, 1), 1).setWrap(true);

  applyValidation_(sheet, "Q2:Q", ["À extraire", "Détecté", "À vérifier", "Exclu", "Traité"]);
  applyValidation_(sheet, "R2:R", ["Forte", "Moyenne", "Faible"]);
  applyValidation_(sheet, "X2:X", ["Oui", "Non"]);
  applyValidation_(sheet, "Y2:Y", ["Oui", "Non"]);

  applyConditionalFormattingDocuments_(sheet);
  highlightSuspectRowsDocuments_(sheet);
}

function formatOngletActions_(ss) {
  var sheet = ss.getSheetByName("02_Actions");
  if (!sheet) return;

  var lastCol = Math.max(sheet.getLastColumn(), 20);
  var lastRow = Math.max(sheet.getLastRow(), 2);

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, lastCol)
    .setFontWeight("bold")
    .setFontColor("#1F2937")
    .setBackground("#D9EAF7")
    .setVerticalAlignment("middle")
    .setWrap(true);

  sheet.getRange(1, 1, lastRow, lastCol).setFontFamily("Arial").setFontSize(10);
  sheet.getRange(2, 1, Math.max(lastRow - 1, 1), lastCol).setVerticalAlignment("top");

  safeCreateFilter_(sheet);

  var widths = {
    1: 220,
    2: 145,
    3: 90,
    4: 170,
    5: 150,
    6: 160,
    7: 170,
    8: 420,
    9: 100,
    10: 120,
    11: 120,
    12: 120,
    13: 220,
    14: 220,
    15: 100,
    16: 160,
    17: 110,
    18: 160,
    19: 280,
    20: 145
  };

  applyWidths_(sheet, widths);

  sheet.getRange(2, 8, Math.max(lastRow - 1, 1), 1).setWrap(true);
  sheet.getRange(2, 19, Math.max(lastRow - 1, 1), 1).setWrap(true);

  applyValidation_(sheet, "I2:I", ["Haute", "Moyenne", "Faible"]);
  applyValidation_(sheet, "L2:L", ["Ouvert", "Traité", "À ignorer", "En attente"]);
  applyValidation_(sheet, "O2:O", ["Oui", "Non"]);
  applyValidation_(sheet, "Q2:Q", ["Oui", "Non"]);

  applyConditionalFormattingActions_(sheet);
  highlightSuspectRowsActions_(sheet);
}

function formatOngletParametres_(ss) {
  var sheet = ss.getSheetByName("00_Parametres");
  if (!sheet) return;

  var lastCol = Math.max(sheet.getLastColumn(), 2);
  var lastRow = Math.max(sheet.getLastRow(), 2);

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, lastCol)
    .setFontWeight("bold")
    .setBackground("#E2F0D9")
    .setFontColor("#1F2937");

  sheet.setColumnWidth(1, 260);
  sheet.setColumnWidth(2, 420);
  sheet.getRange(1, 1, lastRow, lastCol).setWrap(true);
}

function formatOngletsBanque_(ss) {
  var names = [
    "07_Banque_SAS_CA_Import",
    "08_Banques_Perso_Import",
    "09_Banque_Normalisee",
    "10_Rapprochements"
  ];

  names.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) return;

    var lastCol = Math.max(sheet.getLastColumn(), 10);
    var lastRow = Math.max(sheet.getLastRow(), 2);

    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, lastCol)
      .setFontWeight("bold")
      .setBackground("#FCE5CD")
      .setFontColor("#1F2937");

    safeCreateFilter_(sheet);
    sheet.getRange(1, 1, lastRow, lastCol).setFontFamily("Arial").setFontSize(10);
    sheet.autoResizeColumns(1, Math.min(lastCol, 10));
  });
}

function auditDocuments_(ss, now) {
  var sheet = ss.getSheetByName("01_Documents");
  if (!sheet) return [[now, "01_Documents", "", "ERREUR", "Onglet manquant", "L’onglet 01_Documents est introuvable.", "Recréer l’onglet ou relancer le script d’initialisation."]];

  var data = sheet.getDataRange().getValues();
  var rows = [];

  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    var line = i + 1;
    if (rowIsEmpty_(r)) continue;

    var documentId = r[0];
    var dateDetection = r[1];
    var dateMail = r[2];
    var expediteur = r[4];
    var objet = r[6];
    var statut = r[16];
    var regle = r[18];
    var lienGmail = r[21];

    if (looksLikeDateText_(documentId)) {
      rows.push([now, "01_Documents", line, "BLOQUANT", "Colonnes probablement décalées", "Document_ID contient une date. Cela indique que buildDocumentRow_ écrit encore dans un ancien format.", "Corriger buildDocumentRow_ pour retourner exactement 27 valeurs dans le bon ordre."]);
    }

    if (!documentId) {
      rows.push([now, "01_Documents", line, "ALERTE", "Document_ID vide", "La colonne Document_ID n’est pas remplie.", "Vérifier buildDocumentRow_ dans Sheets.gs ou Main.gs."]);
    }

    if (!dateDetection) {
      rows.push([now, "01_Documents", line, "ALERTE", "Date_detection vide", "La date de détection n’est pas remplie.", "Vérifier buildDocumentRow_."]);
    }

    if (!dateMail) {
      rows.push([now, "01_Documents", line, "ALERTE", "Date_mail vide", "La date du mail n’est pas remplie.", "Vérifier processMessage_ et buildDocumentRow_."]);
    }

    if (!expediteur) {
      rows.push([now, "01_Documents", line, "ALERTE", "Expéditeur vide", "La colonne Expéditeur n’est pas remplie.", "Vérifier le mapping des colonnes."]);
    }

    if (!objet) {
      rows.push([now, "01_Documents", line, "INFO", "Objet vide", "Certains mails peuvent ne pas avoir d’objet.", "Pas bloquant, à surveiller."]);
    }

    if (!statut) {
      rows.push([now, "01_Documents", line, "BLOQUANT", "Statut_document vide", "La colonne Statut_document n’est pas remplie.", "Corriger buildDocumentRow_ ou la classification."]);
    }

    if (statut && ["À extraire", "Détecté", "À vérifier", "Exclu", "Traité"].indexOf(statut) === -1) {
      rows.push([now, "01_Documents", line, "ALERTE", "Statut_document non reconnu", String(statut), "Utiliser uniquement les statuts validés."]);
    }

    if (!regle) {
      rows.push([now, "01_Documents", line, "ALERTE", "Règle_appliquée vide", "La règle appliquée n’est pas renseignée.", "Vérifier classifyEmail_ et buildDocumentRow_."]);
    }

    if (!lienGmail) {
      rows.push([now, "01_Documents", line, "ALERTE", "Lien_Gmail vide", "Le lien Gmail n’est pas renseigné.", "Vérifier buildDocumentRow_."]);
    }
  }

  return rows;
}

function auditActions_(ss, now) {
  var sheet = ss.getSheetByName("02_Actions");
  if (!sheet) return [[now, "02_Actions", "", "ERREUR", "Onglet manquant", "L’onglet 02_Actions est introuvable.", "Recréer l’onglet ou relancer le script d’initialisation."]];

  var data = sheet.getDataRange().getValues();
  var rows = [];

  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    var line = i + 1;
    if (rowIsEmpty_(r)) continue;

    var actionId = r[0];
    var dateCreation = r[1];
    var source = r[2];
    var documentId = r[3];
    var action = r[7];
    var priorite = r[8];
    var statut = r[11];
    var lienGmail = r[12];

    if (looksLikeDateText_(actionId)) {
      rows.push([now, "02_Actions", line, "BLOQUANT", "Colonnes probablement décalées", "Action_ID contient une date. Cela indique que buildActionRow_ écrit encore dans un ancien format.", "Corriger buildActionRow_ pour retourner exactement 20 valeurs dans le bon ordre."]);
    }

    if (!actionId) {
      rows.push([now, "02_Actions", line, "ALERTE", "Action_ID vide", "La colonne Action_ID n’est pas remplie.", "Vérifier buildActionRow_."]);
    }

    if (!dateCreation) {
      rows.push([now, "02_Actions", line, "ALERTE", "Date_creation vide", "La date de création de l’action n’est pas remplie.", "Vérifier buildActionRow_."]);
    }

    if (!source) {
      rows.push([now, "02_Actions", line, "ALERTE", "Source vide", "La source de l’action n’est pas remplie.", "Mettre Gmail par défaut dans buildActionRow_."]);
    }

    if (!documentId) {
      rows.push([now, "02_Actions", line, "ALERTE", "Document_ID vide", "L’action n’est pas rattachée à un document.", "Vérifier le passage de documentId dans buildActionRow_."]);
    }

    if (!action) {
      rows.push([now, "02_Actions", line, "BLOQUANT", "Action vide", "La colonne Action n’est pas remplie.", "Vérifier classification.action et buildActionRow_."]);
    }

    if (!priorite) {
      rows.push([now, "02_Actions", line, "INFO", "Priorité vide", "La priorité n’est pas renseignée.", "Mettre Moyenne par défaut."]);
    }

    if (!statut) {
      rows.push([now, "02_Actions", line, "BLOQUANT", "Statut_action vide", "Le statut de l’action n’est pas renseigné.", "Mettre Ouvert par défaut dans buildActionRow_."]);
    }

    if (statut && ["Ouvert", "Traité", "À ignorer", "En attente"].indexOf(statut) === -1) {
      rows.push([now, "02_Actions", line, "ALERTE", "Statut_action non reconnu", String(statut), "Utiliser uniquement les statuts validés."]);
    }

    if (!lienGmail) {
      rows.push([now, "02_Actions", line, "ALERTE", "Lien_Gmail vide", "Le lien Gmail n’est pas renseigné.", "Vérifier buildActionRow_."]);
    }
  }

  return rows;
}

function highlightSuspectRowsDocuments_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var data = sheet.getRange(2, 1, lastRow - 1, 27).getValues();
  var backgrounds = [];

  for (var i = 0; i < data.length; i++) {
    var r = data[i];
    var suspect = looksLikeDateText_(r[0]) || !r[16] || !r[21];
    var color = suspect ? "#FFF2CC" : "#FFFFFF";
    var rowColors = [];
    for (var c = 0; c < 27; c++) rowColors.push(color);
    backgrounds.push(rowColors);
  }

  sheet.getRange(2, 1, data.length, 27).setBackgrounds(backgrounds);
}

function highlightSuspectRowsActions_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var data = sheet.getRange(2, 1, lastRow - 1, 20).getValues();
  var backgrounds = [];

  for (var i = 0; i < data.length; i++) {
    var r = data[i];
    var suspect = looksLikeDateText_(r[0]) || !r[7] || !r[11] || !r[12];
    var color = suspect ? "#FFF2CC" : "#FFFFFF";
    var rowColors = [];
    for (var c = 0; c < 20; c++) rowColors.push(color);
    backgrounds.push(rowColors);
  }

  sheet.getRange(2, 1, data.length, 20).setBackgrounds(backgrounds);
}

function applyConditionalFormattingDocuments_(sheet) {
  var rules = [];

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("À vérifier")
    .setBackground("#FFF2CC")
    .setRanges([sheet.getRange("Q2:Q")])
    .build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("À extraire")
    .setBackground("#D9EAF7")
    .setRanges([sheet.getRange("Q2:Q")])
    .build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Détecté")
    .setBackground("#E2F0D9")
    .setRanges([sheet.getRange("Q2:Q")])
    .build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Traité")
    .setBackground("#D9EAD3")
    .setRanges([sheet.getRange("Q2:Q")])
    .build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Exclu")
    .setBackground("#E7E6E6")
    .setRanges([sheet.getRange("Q2:Q")])
    .build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Forte")
    .setBackground("#D9EAD3")
    .setRanges([sheet.getRange("R2:R")])
    .build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Moyenne")
    .setBackground("#FCE5CD")
    .setRanges([sheet.getRange("R2:R")])
    .build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Faible")
    .setBackground("#F4CCCC")
    .setRanges([sheet.getRange("R2:R")])
    .build());

  sheet.setConditionalFormatRules(rules);
}

function applyConditionalFormattingActions_(sheet) {
  var rules = [];

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Ouvert")
    .setBackground("#FFF2CC")
    .setRanges([sheet.getRange("L2:L")])
    .build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Traité")
    .setBackground("#D9EAD3")
    .setRanges([sheet.getRange("L2:L")])
    .build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("À ignorer")
    .setBackground("#E7E6E6")
    .setRanges([sheet.getRange("L2:L")])
    .build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("En attente")
    .setBackground("#FCE5CD")
    .setRanges([sheet.getRange("L2:L")])
    .build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Haute")
    .setBackground("#F4CCCC")
    .setRanges([sheet.getRange("I2:I")])
    .build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Moyenne")
    .setBackground("#FCE5CD")
    .setRanges([sheet.getRange("I2:I")])
    .build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Faible")
    .setBackground("#E2F0D9")
    .setRanges([sheet.getRange("I2:I")])
    .build());

  sheet.setConditionalFormatRules(rules);
}

function formatAuditSheet_(sheet) {
  var lastRow = Math.max(sheet.getLastRow(), 2);
  var lastCol = Math.max(sheet.getLastColumn(), 7);

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, lastCol)
    .setFontWeight("bold")
    .setBackground("#F4CCCC")
    .setFontColor("#1F2937")
    .setWrap(true);

  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 180);
  sheet.setColumnWidth(3, 70);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 240);
  sheet.setColumnWidth(6, 420);
  sheet.setColumnWidth(7, 420);

  sheet.getRange(1, 1, lastRow, lastCol).setFontFamily("Arial").setFontSize(10).setWrap(true);
  safeCreateFilter_(sheet);

  var rules = [];

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("BLOQUANT")
    .setBackground("#F4CCCC")
    .setRanges([sheet.getRange("D2:D")])
    .build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("ALERTE")
    .setBackground("#FFF2CC")
    .setRanges([sheet.getRange("D2:D")])
    .build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("OK")
    .setBackground("#D9EAD3")
    .setRanges([sheet.getRange("D2:D")])
    .build());

  sheet.setConditionalFormatRules(rules);
}

function applyValidation_(sheet, rangeA1, values) {
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(true)
    .build();

  sheet.getRange(rangeA1).setDataValidation(rule);
}

function safeCreateFilter_(sheet) {
  try {
    var filter = sheet.getFilter();
    if (!filter) {
      var range = sheet.getDataRange();
      if (range.getNumRows() > 0 && range.getNumColumns() > 0) {
        range.createFilter();
      }
    }
  } catch (e) {
    Logger.log("Filtre non créé sur " + sheet.getName() + " : " + e);
  }
}

function applyWidths_(sheet, widths) {
  Object.keys(widths).forEach(function(col) {
    sheet.setColumnWidth(Number(col), widths[col]);
  });
}

function rowIsEmpty_(row) {
  return row.every(function(cell) {
    return cell === "" || cell === null;
  });
}

function looksLikeDateText_(value) {
  if (!value) return false;
  var text = String(value);
  return /^\d{4}-\d{2}-\d{2}/.test(text);
}

function getOrCreateSheetMiseEnForme_(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

function getComptaSpreadsheetMiseEnForme_() {
  var id = "";

  try {
    if (typeof Config !== "undefined" && Config.SPREADSHEET_ID) {
      id = Config.SPREADSHEET_ID;
    }
  } catch (e) {}

  try {
    if (!id && typeof SPREADSHEET_ID !== "undefined") {
      id = SPREADSHEET_ID;
    }
  } catch (e2) {}

  if (!id) {
    id = "1b6xwGgOi7MVl2dJR1gQ4PcI0uM2rJQKp-abph5_GPQA";
  }

  return SpreadsheetApp.openById(id);
}

function setupMiseEnFormeLisibleComptaAdmin() {
  setupMiseEnFormeRapideComptaAdmin();
}

function setupMiseEnFormeRapideComptaAdmin() {
  Logger.log('Début setup rapide');
  formatVueActionsSimple_();
  Logger.log('Format 02_Actions terminé');
  formatDomainesSuivi_();
  Logger.log('Format 12_Domaines_Suivi terminé');
  creerOuMajModeEmploi_();
  Logger.log('Mode emploi terminé');
  SpreadsheetApp.flush();
  Logger.log('Setup rapide terminé');
}

function formatVueActionsSimple_() {
  var ss = getComptaSpreadsheetMiseEnForme_();
  var sheet = ss.getSheetByName('02_Actions');
  if (!sheet) return;

  var lastRow = Math.min(Math.max(sheet.getLastRow(), 2), 300);
  var lastCol = Math.min(Math.max(sheet.getLastColumn(), 20), 26);
  var maxColumns = sheet.getMaxColumns();
  var columnsToFormat = Math.min(26, maxColumns);
  var visibleCols = [5, 8, 9, 12, 13, 19, 20];

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, columnsToFormat)
    .setBackground('#D9EAF7')
    .setFontWeight('bold')
    .setFontColor('#1F2937')
    .setWrap(true);
  sheet.getRange(1, 1, lastRow, columnsToFormat)
    .setFontFamily('Arial')
    .setFontSize(10);

  for (var c = 1; c <= Math.min(lastCol, maxColumns); c++) {
    sheet.showColumns(c);
  }
  for (var h = 1; h <= Math.min(26, maxColumns); h++) {
    if (visibleCols.indexOf(h) === -1) {
      sheet.hideColumns(h);
    }
  }

  sheet.setColumnWidth(5, 220);   // Domaine
  sheet.setColumnWidth(8, 460);   // Action
  sheet.setColumnWidth(9, 120);   // Priorité
  sheet.setColumnWidth(12, 150);  // Statut_action
  sheet.setColumnWidth(13, 180);  // Lien_Gmail
  sheet.setColumnWidth(19, 360);  // Commentaire
  sheet.setColumnWidth(20, 150);  // Dernière_maj

  if (lastRow > 1) {
    sheet.getRange(2, 8, lastRow - 1, 1).setWrap(true);   // Action
    sheet.getRange(2, 19, lastRow - 1, 1).setWrap(true);  // Commentaire
    sheet.getRange(2, 12, lastRow - 1, 1).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['Ouvert', 'Traité', 'À ignorer', 'En attente'], true)
        .setAllowInvalid(true)
        .build()
    );
  }

  var rules = [];
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Ouvert').setBackground('#FFF2CC').setRanges([sheet.getRange('L2:L300')]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Traité').setBackground('#D9EAD3').setRanges([sheet.getRange('L2:L300')]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('À ignorer').setBackground('#E7E6E6').setRanges([sheet.getRange('L2:L300')]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('En attente').setBackground('#FCE5CD').setRanges([sheet.getRange('L2:L300')]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Haute').setBackground('#F4CCCC').setRanges([sheet.getRange('I2:I300')]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Moyenne').setBackground('#FCE5CD').setRanges([sheet.getRange('I2:I300')]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Faible').setBackground('#D9EAD3').setRanges([sheet.getRange('I2:I300')]).build());
  sheet.setConditionalFormatRules(rules);
}

function creerOuMajModeEmploi_() {
  var ss = getComptaSpreadsheetMiseEnForme_();
  var sheet = ss.getSheetByName('13_Mode_Emploi');
  if (!sheet) {
    sheet = ss.insertSheet('13_Mode_Emploi');
  }

  var rows = [
    ['Section', 'Contenu', '', ''],
    ['Objectif', 'Piloter rapidement les actions compta/admin issues de Gmail.', '', ''],
    ['02_Actions', 'Vue quotidienne : Domaine, Action, Priorité, Statut_action, Lien_Gmail, Commentaire, Dernière_maj.', '', ''],
    ['Statut_action', 'Décision finale utilisateur. Ne pas modifier automatiquement.', '', ''],
    ['Commentaire', 'Note manuelle courte après lecture.', '', ''],
    ['Tri manuel', 'À ignorer renforce les exclusions ; Traité renforce les signaux à conserver.', '', ''],
    ['IA', 'Utiliser testOpenAIAnalyseLigneActive sur une ligne sélectionnée si besoin.', '', ''],
    ['Mise en forme', 'setupMiseEnFormeRapideComptaAdmin applique uniquement une mise en forme légère.', '', '']
  ];

  sheet.getRange('A1:D30').clearContent().clearFormat();
  sheet.getRange(1, 1, rows.length, 4).setValues(rows);
  sheet.setFrozenRows(1);
  sheet.getRange('A1:D1').setBackground('#D9EAF7').setFontWeight('bold').setFontColor('#1F2937');
  sheet.getRange(1, 1, rows.length, 4).setWrap(true).setVerticalAlignment('middle').setFontFamily('Arial').setFontSize(10);
  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 700);
}

function reconstruireExplicationAction_() {
  assurerColonnesAnalyseIAActions_();
}

function assurerColonnesAnalyseIAActions_() {
  var ss = getComptaSpreadsheetMiseEnForme_();
  var sheet = ss.getSheetByName('02_Actions');
  if (!sheet) return;

  var targetHeaders = ['Analyse_IA', 'Décision_IA', 'Catégorie_IA', 'Action_recommandée_IA', 'Règle_proposée_IA', 'Date_analyse_IA'];
  var startCol = 21;
  var neededLastCol = startCol + targetHeaders.length - 1;
  while (sheet.getLastColumn() < neededLastCol) {
    sheet.insertColumnAfter(sheet.getLastColumn());
  }

  for (var i = 0; i < targetHeaders.length; i++) {
    var col = startCol + i;
    if (!sheet.getRange(1, col).getValue()) {
      sheet.getRange(1, col).setValue(targetHeaders[i]);
    } else if (col === 21) {
      sheet.getRange(1, col).setValue('Analyse_IA');
    }
  }
}



function formatDocumentsLisible_() {
  var ss = getComptaSpreadsheetMiseEnForme_();
  var sheet = ss.getSheetByName('01_Documents');
  if (!sheet) return;

  var lastRow = Math.max(sheet.getLastRow(), 2);
  var lastCol = Math.max(sheet.getLastColumn(), 27);

  sheet.setFrozenRows(1);
  safeCreateFilter_(sheet);
  sheet.getRange(1, 1, 1, lastCol).setBackground('#D9EAF7').setFontWeight('bold').setFontColor('#1F2937').setWrap(true);
  sheet.getRange(1, 1, lastRow, lastCol).setFontFamily('Arial').setFontSize(10);

  for (var c = 1; c <= lastCol; c++) {
    sheet.showColumns(c);
  }
  [1, 2, 4, 6, 8, 10, 11, 16, 20, 24, 25, 26].forEach(function(col) {
    if (col <= lastCol) sheet.hideColumns(col);
  });

  sheet.setColumnWidth(3, 160);   // Date_mail
  sheet.setColumnWidth(5, 220);   // Expéditeur
  sheet.setColumnWidth(7, 420);   // Objet
  sheet.setColumnWidth(9, 320);   // Nom_piece_jointe
  sheet.setColumnWidth(12, 130);  // Entité
  sheet.setColumnWidth(13, 180);  // Activité
  sheet.setColumnWidth(14, 220);  // Catégorie
  sheet.setColumnWidth(15, 360);  // Dossier_cible
  sheet.setColumnWidth(17, 140);  // Statut_document
  sheet.setColumnWidth(18, 120);  // Confiance
  sheet.setColumnWidth(19, 200);  // Règle_appliquée
  sheet.setColumnWidth(21, 360);  // Action_recommandée
  sheet.setColumnWidth(22, 220);  // Lien_Gmail
  sheet.setColumnWidth(23, 220);  // Lien_Drive
  sheet.setColumnWidth(27, 170);  // Dernière_maj

  sheet.getRange(2, 7, lastRow - 1, 1).setWrap(true);   // Objet
  sheet.getRange(2, 9, lastRow - 1, 1).setWrap(true);   // PJ
  sheet.getRange(2, 15, lastRow - 1, 1).setWrap(true);  // Dossier
  sheet.getRange(2, 21, lastRow - 1, 1).setWrap(true);  // Action reco

  applyConditionalFormattingDocuments_(sheet);
}

function formatSyntheseLisible_() {
  rebuildSyntheseComptaAdmin_();
  var ss = getComptaSpreadsheetMiseEnForme_();
  var sheet = ss.getSheetByName('04_Synthese');
  if (!sheet) return;

  var routineStart = 24;
  var routine = [
    ['Routine IA prioritaire', '', '', '', '', ''],
    ['1', 'Regarder 04_Synthese', 'Identifier les volumes et urgences', '', '', ''],
    ['2', 'Ouvrir 02_Actions', 'Lire Analyse_IA (U) et ouvrir Lien_Gmail si nécessaire', '', '', ''],
    ['3', 'Mettre Statut_action', 'Décider Traité / En attente / À ignorer', '', '', ''],
    ['4', 'Renseigner Commentaire', 'Ajouter une note manuelle courte si utile', '', '', '']
  ];
  sheet.getRange(routineStart, 1, routine.length, 6).setValues(routine);
  sheet.getRange('A24:C24').setFontWeight('bold').setBackground('#D9EAF7');
  sheet.getRange('A24:F28').setWrap(true).setVerticalAlignment('middle');
}

function formatDomainesSuivi_() {
  var ss = getComptaSpreadsheetMiseEnForme_();
  var sheet = ss.getSheetByName('12_Domaines_Suivi');
  if (!sheet) return;

  var lastRow = Math.min(Math.max(sheet.getLastRow(), 2), 300);
  var lastCol = Math.min(Math.max(sheet.getLastColumn(), 8), 8);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, lastCol)
    .setBackground('#E6D9F2')
    .setFontWeight('bold')
    .setFontColor('#1F2937')
    .setWrap(true);
  sheet.getRange(1, 1, lastRow, lastCol)
    .setFontFamily('Arial')
    .setFontSize(10)
    .setWrap(true);

  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 260);
  sheet.setColumnWidth(3, 140);
  sheet.setColumnWidth(4, 180);
  sheet.setColumnWidth(5, 320);
}

function formatSynthesesIA_() {
  var ss = getComptaSpreadsheetMiseEnForme_();
  var sheet = ss.getSheetByName('05_Syntheses_IA');
  if (!sheet) return;

  var lastRow = Math.max(sheet.getLastRow(), 2);
  var lastCol = Math.max(sheet.getLastColumn(), 8);
  sheet.setFrozenRows(1);
  safeCreateFilter_(sheet);
  sheet.getRange(1, 1, 1, lastCol).setBackground('#D9EAF7').setFontWeight('bold').setFontColor('#1F2937').setWrap(true);
  sheet.getRange(1, 1, lastRow, lastCol).setFontFamily('Arial').setFontSize(10).setWrap(true);
  sheet.autoResizeColumns(1, Math.min(lastCol, 12));
}

function formatAuditColonnes_() {
  var ss = getComptaSpreadsheetMiseEnForme_();
  var sheet = ss.getSheetByName('11_Audit_Colonnes');
  if (!sheet) return;
  formatAuditSheet_(sheet);
}

function appliquerNavigationEtCouleurs_() {
  var ss = getComptaSpreadsheetMiseEnForme_();
  var tabColors = {
    '01_Documents': '#9FC5E8',
    '02_Actions': '#F9CB9C',
    '04_Synthese': '#B6D7A8',
    '05_Syntheses_IA': '#CFE2F3',
    '11_Audit_Colonnes': '#F4CCCC',
    '12_Domaines_Suivi': '#D9D2E9'
  };
  Object.keys(tabColors).forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (sh) sh.setTabColor(tabColors[name]);
  });
}

function reconstruireVisionIAActions_() {
  var ss = getComptaSpreadsheetMiseEnForme_();
  var sheet = ss.getSheetByName('02_Actions');
  if (!sheet) return;

  while (sheet.getLastColumn() < 21) {
    sheet.insertColumnAfter(sheet.getLastColumn());
  }
  sheet.getRange(1, 21).setValue('Analyse_IA');
  assurerColonnesAnalyseIAActions_();
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Compta Admin')
    .addItem('Mettre en forme le Sheet', 'setupMiseEnFormeLisibleComptaAdmin')
    .addItem('Reconstruire vision IA Actions', 'reconstruireVisionIAActions_')
    .addSeparator()
    .addItem('Configurer clé OpenAI', 'setOpenAIKey')
    .addItem('Tester IA sur 1 action ouverte', 'testOpenAIAnalyseUneAction')
    .addItem('Analyser 10 actions ouvertes (IA)', 'analyserActionsOuvertesAvecIA')
    .addItem('Auditer colonnes', 'auditColonnesComptaAdmin')
    .addItem('Recréer synthèse', 'rebuildSyntheseComptaAdmin_')
    .addToUi();
}

function setupPerformanceComptaAdmin() {
  Logger.log('setupPerformanceComptaAdmin : début');
  var ss = getComptaSpreadsheetMiseEnForme_();
  var sheet = ss.getSheetByName('02_Actions');
  if (!sheet) throw new Error('Onglet introuvable : 02_Actions');

  Logger.log('setupPerformanceComptaAdmin : limitation visuelle à 300 lignes et 26 colonnes');
  var maxRows = 300;
  var maxCols = 26;
  var rowsToFormat = Math.min(Math.max(sheet.getLastRow(), 1), maxRows);
  var colsToFormat = Math.min(Math.max(sheet.getLastColumn(), maxCols), maxCols);

  Logger.log('setupPerformanceComptaAdmin : suppression des règles conditionnelles lourdes');
  sheet.setConditionalFormatRules([]);
  var lightRules = [];
  lightRules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Ouvert').setBackground('#FFF2CC').setRanges([sheet.getRange('L2:L300')]).build());
  lightRules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Traité').setBackground('#D9EAD3').setRanges([sheet.getRange('L2:L300')]).build());
  lightRules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('À ignorer').setBackground('#E7E6E6').setRanges([sheet.getRange('L2:L300')]).build());
  sheet.setConditionalFormatRules(lightRules);

  Logger.log('setupPerformanceComptaAdmin : suppression/recréation du filtre A1:Z300');
  try {
    var filter = sheet.getFilter();
    if (filter) filter.remove();
  } catch (e) {
    Logger.log('setupPerformanceComptaAdmin : filtre existant non supprimé : ' + e);
  }
  sheet.getRange(1, 1, maxRows, maxCols).createFilter();

  Logger.log('setupPerformanceComptaAdmin : gel ligne 1 uniquement');
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(0);

  Logger.log('setupPerformanceComptaAdmin : mise en forme légère bornée');
  sheet.getRange(1, 1, 1, colsToFormat)
    .setBackground('#D9EAF7')
    .setFontWeight('bold')
    .setFontColor('#1F2937')
    .setWrap(true);
  sheet.getRange(1, 1, rowsToFormat, colsToFormat)
    .setFontFamily('Arial')
    .setFontSize(10)
    .setVerticalAlignment('top');

  if (rowsToFormat > 1) {
    sheet.getRange(2, 8, rowsToFormat - 1, 1).setWrap(true);
    sheet.getRange(2, 19, rowsToFormat - 1, 1).setWrap(true);
    sheet.getRange(2, 24, rowsToFormat - 1, 1).setWrap(true);
    sheet.getRange(2, 26, rowsToFormat - 1, 1).setWrap(true);
  }

  Logger.log('setupPerformanceComptaAdmin : masquage des colonnes non utiles');
  var maxColumns = sheet.getMaxColumns();
  var visibleCols = [5, 8, 9, 12, 13, 19, 21, 22, 23, 24, 25, 26];
  for (var c = 1; c <= maxColumns; c++) {
    sheet.showColumns(c);
  }
  for (var h = 1; h <= maxColumns; h++) {
    if (visibleCols.indexOf(h) === -1) sheet.hideColumns(h);
  }

  Logger.log('setupPerformanceComptaAdmin : largeurs colonnes visibles');
  var widths = {
    5: 180,
    8: 420,
    9: 100,
    12: 130,
    13: 180,
    19: 320,
    21: 360,
    22: 240,
    23: 170,
    24: 360,
    25: 360,
    26: 420
  };
  applyWidths_(sheet, widths);

  SpreadsheetApp.flush();
  Logger.log('setupPerformanceComptaAdmin : fin');
}
