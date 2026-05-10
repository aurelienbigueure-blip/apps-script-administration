function enrichirActionsAvecCommentairesIA() {
  const ss = getComptaSpreadsheetForAnalyseIA_();

  const sheetActions = ss.getSheetByName("02_Actions");
  const sheetDocs = ss.getSheetByName("01_Documents");

  if (!sheetActions) {
    throw new Error("Onglet introuvable : 02_Actions");
  }
  if (!sheetDocs) {
    throw new Error("Onglet introuvable : 01_Documents");
  }

  const lastRowActions = sheetActions.getLastRow();
  if (lastRowActions < 2) {
    Logger.log("Aucune action à enrichir.");
    return;
  }

  while (sheetActions.getLastColumn() < 21) {
    sheetActions.insertColumnAfter(sheetActions.getLastColumn());
  }
  sheetActions.getRange(1, 21).setValue("Explication_action");

  const headersActions = sheetActions
    .getRange(1, 1, 1, Math.max(sheetActions.getLastColumn(), 21))
    .getValues()[0];

  const colActionDocumentId = headersActions.indexOf("Document_ID") + 1;
  const colExplicationAction = 21;

  if (colActionDocumentId < 1) {
    throw new Error("Colonne Document_ID introuvable dans 02_Actions");
  }

  const docsValues = sheetDocs
    .getRange(1, 1, sheetDocs.getLastRow(), sheetDocs.getLastColumn())
    .getValues();

  const headersDocs = docsValues[0];

  const colDocId = headersDocs.indexOf("Document_ID");
  const colExpediteur = headersDocs.indexOf("Expéditeur");
  const colObjet = headersDocs.indexOf("Objet");
  const colExtrait = headersDocs.indexOf("Extrait_mail");
  const colPJ = headersDocs.indexOf("Nom_piece_jointe");

  if (colDocId === -1) {
    throw new Error("Colonne Document_ID introuvable dans 01_Documents");
  }

  const docsById = {};
  for (let i = 1; i < docsValues.length; i++) {
    const row = docsValues[i];
    const documentId = row[colDocId];
    if (documentId) {
      docsById[String(documentId)] = row;
    }
  }

  const actionsValues = sheetActions
    .getRange(2, 1, lastRowActions - 1, sheetActions.getLastColumn())
    .getValues();

  let nbEnrichies = 0;

  for (let i = 0; i < actionsValues.length; i++) {
    const actionRow = actionsValues[i];
    const documentId = actionRow[colActionDocumentId - 1];

    if (!documentId) continue;

    const docRow = docsById[String(documentId)];
    if (!docRow) {
      sheetActions
        .getRange(i + 2, colExplicationAction)
        .setValue("Mail non reconnu : ouvrir le lien Gmail, décider s’il faut traiter, ignorer ou créer une règle.");
      nbEnrichies++;
      continue;
    }

    const explication = analyserMailPourActionIA_({
      expediteur: safeValue(docRow[colExpediteur]),
      objet: safeValue(docRow[colObjet]),
      extrait: safeValue(docRow[colExtrait]),
      pieceJointe: safeValue(docRow[colPJ])
    });

    sheetActions.getRange(i + 2, colExplicationAction).setValue(explication);
    nbEnrichies++;
  }

  formaterColonnesIAActions_();
  Logger.log("Enrichissement terminé. Lignes enrichies : " + nbEnrichies);
}

function safeValue(value) {
  if (value === null || value === undefined) return "";
  return String(value);
}

function analyserMailPourActionIA_(mail) {
  const texte = [mail.expediteur, mail.objet, mail.extrait, mail.pieceJointe]
    .join(" ")
    .toLowerCase();

  if (texte.includes("msa") || texte.includes("notification.m")) {
    return "Mail MSA ou notification administrative : vérifier le contenu, puis conserver ou traiter si cela concerne une démarche personnelle.";
  }

  if (texte.includes("people-doc") || texte.includes("peopledoc")) {
    return "Mail PeopleDoc : document RH/personnel probable. À conserver si bulletin, attestation ou document salarié.";
  }

  if (
    texte.includes("groupama-epargne") ||
    texte.includes("épargne") ||
    texte.includes("epargne") ||
    texte.includes("intéressement") ||
    texte.includes("interessement")
  ) {
    return "Mail Groupama Épargne : sujet épargne salariale / intéressement. À traiter ou conserver.";
  }

  if (texte.includes("chimirec")) {
    return "Mail Chimirec : fournisseur potentiel. Vérifier s’il s’agit d’une facture, d’un contrat ou d’une simple notification.";
  }

  if (
    texte.includes("creditmutuel") &&
    (texte.includes("candidature") || texte.includes("emploi") || texte.includes("offre"))
  ) {
    return "Mail emploi ou candidature : faux positif hors compta/admin. Passer Statut_action à À ignorer.";
  }

  if (texte.includes("google.com")) {
    return "Notification Google ou technique : à ignorer sauf sécurité compte ou erreur Apps Script à corriger.";
  }

  if (
    texte.includes("publicit") ||
    texte.includes("newsletter") ||
    texte.includes("marketing") ||
    texte.includes("promotion")
  ) {
    return "Mail promotionnel ou newsletter : faux positif. Passer Statut_action à À ignorer.";
  }

  if (
    texte.includes("groupama.com") ||
    texte.includes("groupama-ne") ||
    texte.includes("groupama-ra") ||
    texte.includes("comet")
  ) {
    return "Mail professionnel transféré : hors périmètre compta/admin personnel. Passer à À ignorer sauf exception.";
  }

  return "Mail non reconnu : ouvrir le lien Gmail, décider s’il faut traiter, ignorer ou créer une règle.";
}

function formaterColonnesIAActions_() {
  const ss = getComptaSpreadsheetForAnalyseIA_();
  const sheet = ss.getSheetByName("02_Actions");
  if (!sheet) return;

  const headers = sheet
    .getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 21))
    .getValues()[0];

  const colExplication = headers.indexOf("Explication_action") + 1;
  if (colExplication > 0) {
    sheet.setColumnWidth(colExplication, 460);
    sheet.getRange(1, colExplication).setFontWeight("bold").setBackground("#DDEBF7");

    if (sheet.getLastRow() > 1) {
      sheet
        .getRange(2, colExplication, sheet.getLastRow() - 1, 1)
        .setWrap(true)
        .setBackground("#EAF2F8");
    }
  }
}

function getComptaSpreadsheetForAnalyseIA_() {
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
