// Point d’entrée principal du traitement Gmail vers Sheets/Drive.
function runComptaMailProcessing() {
  var params = getOrCreateParamSheet_();
  var lastRun = getLastRunDatetime_(params);
  var now = new Date();

  var query = buildGmailQuery_(lastRun);
  Logger.log("Requête Gmail utilisée : " + query);

  var threads = GmailApp.search(query, 0, PROCESSING_CONFIG.MAX_THREADS);

  var docSheet = getOrCreateDocSheet_();
  var actionSheet = getOrCreateActionSheet_();

  var treatedLabel = getOrCreateLabel_(PROCESSING_CONFIG.LABEL_TREATED);
  var verifyLabel = getOrCreateLabel_(PROCESSING_CONFIG.LABEL_TO_VERIFY);
  var excludedLabel = getOrCreateLabel_(PROCESSING_CONFIG.LABEL_EXCLUDED);

  var processedCount = 0;
  var documentRowsCount = 0;
  var actionRowsCount = 0;
  var excludedCount = 0;
  var errorCount = 0;

  threads.forEach(function(thread) {
    try {
      if (
        threadHasLabel_(thread, treatedLabel) ||
        threadHasLabel_(thread, verifyLabel) ||
        threadHasLabel_(thread, excludedLabel)
      ) {
        return;
      }

      processedCount++;

      var messages = thread.getMessages();
      var threadFinalStatus = "Traité";

      messages.forEach(function(message) {
        try {
          var result = processMessage_(message);

          if (!result) {
            Logger.log("Message ignoré : résultat vide");
            return;
          }

          if (result.status === "Exclu") {
            excludedCount++;
            threadFinalStatus = "Exclu";
            return;
          }

          if (result.documentRow && result.documentRow.length > 0) {
            appendDocumentRow_(docSheet, result.documentRow);
            documentRowsCount++;
          } else {
            Logger.log("Aucune ligne document à écrire pour le message : " + message.getId());
          }

          if (result.actionRow && result.actionRow.length > 0) {
            appendActionRow_(actionSheet, result.actionRow);
            actionRowsCount++;
          }

          if (result.status === "À vérifier") {
            threadFinalStatus = "À vérifier";
          }

        } catch (messageError) {
          errorCount++;
          Logger.log("Erreur message " + message.getId() + " : " + messageError);
        }
      });

      if (threadFinalStatus === "Exclu") {
        thread.addLabel(excludedLabel);
      } else if (threadFinalStatus === "À vérifier") {
        thread.addLabel(verifyLabel);
      } else {
        thread.addLabel(treatedLabel);
      }

    } catch (threadError) {
      errorCount++;
      Logger.log("Erreur thread : " + threadError);
    }
  });

  updateLastRunDatetime_(params, now);

  Logger.log("Traitement terminé");
  Logger.log("Conversations analysées : " + processedCount);
  Logger.log("Documents écrits : " + documentRowsCount);
  Logger.log("Actions écrites : " + actionRowsCount);
  Logger.log("Mails exclus : " + excludedCount);
  Logger.log("Erreurs : " + errorCount);
}

// Construit la requête Gmail de recherche des nouveaux emails.
function buildGmailQuery_(lastRunDate) {
  var parts = [];
  parts.push('-label:' + quoteGmailLabel_(PROCESSING_CONFIG.LABEL_TREATED));
  parts.push('-label:' + quoteGmailLabel_(PROCESSING_CONFIG.LABEL_TO_VERIFY));
  parts.push('-label:' + quoteGmailLabel_(PROCESSING_CONFIG.LABEL_EXCLUDED));

  if (lastRunDate) {
    var formatted = Utilities.formatDate(lastRunDate, Session.getScriptTimeZone(), 'yyyy/MM/dd');
    parts.push('after:' + formatted);
  } else {
    parts.push('newer_than:' + PROCESSING_CONFIG.LOOKBACK_DAYS + 'd');
  }

  return parts.join(' ');
}

// Traite un email unique et prépare les lignes à écrire.
function processMessage_(message) {
  var from = message.getFrom() || '';
  var to = (message.getTo() || '') + '\n' + (message.getCc() || '') + '\n' + (message.getBcc() || '');
  var subject = message.getSubject() || '';
  var date = message.getDate();
  var attachments = message.getAttachments({ includeInlineImages: false, includeAttachments: true });
  var bodyPreview = (message.getPlainBody() || '').substring(0, 1000);

  var classification = classifyEmail_(from, to, subject, bodyPreview);

  if (classification.status === 'Exclu') {
    return { status: 'Exclu' };
  }

  var relevantAttachments = [];
  attachments.forEach(function(att) {
    var filename = att.getName();
    var ext = getExtension_(filename);
    var allowedExt = ['pdf', 'doc', 'docx', 'xls', 'xlsx', 'csv', 'jpg', 'jpeg', 'png'];
    if (allowedExt.indexOf(ext) !== -1) {
      relevantAttachments.push(att);
    }
  });

  if (relevantAttachments.length === 0 && classification.activity === '' && classification.category === '') {
    classification.status = 'À vérifier';
    classification.action = 'Vérifier le classement de cet e-mail (pas de pièce jointe comptable détectée)';
  }

  var documentRow = buildDocumentRow_(message, {
    from: from,
    to: to,
    subject: subject,
    bodyPreview: bodyPreview,
    dateMail: date,
    relevantAttachments: relevantAttachments,
    classification: classification
  });

  var actionRow = null;
  if (classification.status === 'À vérifier' || classification.action) {
    actionRow = buildActionRow_(message, classification, documentRow[0]);
  }

  if (!PROCESSING_CONFIG.MODE_TEST && classification.status !== 'Exclu' && relevantAttachments.length > 0) {
    saveAttachmentsToDrive_(relevantAttachments, date, from, subject, classification.dossier);
  }

  return {
    documentRow: documentRow,
    actionRow: actionRow,
    status: classification.status
  };
}

// Construit une ligne 01_Documents au format 27 colonnes attendu.
function buildDocumentRow_(message, payload) {
  var now = new Date();
  var nowText = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  var dateMailText = Utilities.formatDate(payload.dateMail, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  var attachmentNames = payload.relevantAttachments.map(function(a) { return a.getName(); }).join('\n');
  var attachmentTypes = payload.relevantAttachments.map(function(a) { return getExtension_(a.getName()); }).join('\n');
  var attachmentSizes = payload.relevantAttachments.map(function(a) { return String(a.getBytes().length); }).join('\n');
  var gmailLink = 'https://mail.google.com/mail/u/0/#all/' + message.getId();
  var decisionIA = payload.classification.status === 'À vérifier' ? 'À valider' : 'Automatique';

  return [
    message.getId(),                                   // Document_ID
    nowText,                                           // Date_detection
    dateMailText,                                      // Date_mail
    inferMailDirection_(payload.from),                 // Sens_mail
    payload.from,                                      // Expéditeur
    payload.to,                                        // Destinataires
    payload.subject,                                   // Objet
    payload.bodyPreview,                               // Extrait_mail
    attachmentNames,                                   // Nom_piece_jointe
    attachmentTypes,                                   // Type_piece_jointe
    attachmentSizes,                                   // Taille_piece_jointe
    payload.classification.entity,                     // Entité
    payload.classification.activity,                   // Activité
    payload.classification.category,                   // Catégorie
    payload.classification.dossier,                    // Dossier_cible
    payload.classification.dossier,                    // Dossier_existant_probable
    payload.classification.status,                     // Statut_document
    payload.classification.confiance,                  // Confiance
    payload.classification.rule,                       // Règle_appliquée
    decisionIA,                                        // Décision_IA
    payload.classification.action || '',               // Action_recommandée
    gmailLink,                                         // Lien_Gmail
    '',                                                // Lien_Drive
    'Non',                                             // Déjà_transmis_comptable
    'Non',                                             // Doublon_probable
    '',                                                // Commentaire
    nowText                                            // Dernière_maj
  ];
}

// Construit une ligne 02_Actions au format 20 colonnes attendu.
function buildActionRow_(message, classification, documentId) {
  var now = new Date();
  var nowText = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  var gmailLink = 'https://mail.google.com/mail/u/0/#all/' + message.getId();

  return [
    Utilities.getUuid(),                               // Action_ID
    nowText,                                           // Date_creation
    'Gmail',                                           // Source
    documentId || message.getId(),                     // Document_ID
    extractDomainFromFrom_(message.getFrom() || ''),   // Domaine
    classification.activity || '',                     // Activité
    classification.category || '',                     // Type_action
    classification.action || 'Vérifier',               // Action
    'Moyenne',                                         // Priorité
    '',                                                // Échéance
    'Aurélien',                                        // Responsable
    'Ouvert',                                          // Statut_action
    gmailLink,                                         // Lien_Gmail
    '',                                                // Lien_Drive
    'Non',                                             // Créer_Task
    '',                                                // Task_ID
    'Non',                                             // Créer_Calendar
    '',                                                // Calendar_Event_ID
    '',                                                // Commentaire
    nowText                                            // Dernière_maj
  ];
}

// Détermine le sens du mail (reçu / envoyé) depuis l'expéditeur.
function inferMailDirection_(from) {
  var me = '';
  try {
    me = Session.getActiveUser().getEmail();
  } catch (e) {}
  if (me && from && from.toLowerCase().indexOf(me.toLowerCase()) !== -1) {
    return 'Envoyé';
  }
  return 'Reçu';
}

// Extrait le domaine depuis le champ expéditeur.
function extractDomainFromFrom_(from) {
  var source = extractEmailOrName_(from || '');
  var match = source.match(/@([^>\\s]+)/);
  return match && match[1] ? match[1].toLowerCase() : '';
}
