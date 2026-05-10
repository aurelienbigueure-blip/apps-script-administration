// Détermine si un email est clairement à exclure.
function isClearlyExcluded_(from, to, subject, bodyPreview) {
  var lowFrom = (from || '').toLowerCase();
  var lowTo = (to || '').toLowerCase();
  var lowSub = (subject || '').toLowerCase();
  var lowBody = (bodyPreview || '').toLowerCase();
  var haystack = [lowFrom, lowTo, lowSub, lowBody].join(' ');

  var neverExcludePatterns = [
    'cerfrance',
    'alliancecomtoise.cerfrance.fr',
    'enedis',
    'wecasa',
    'airbnb',
    'banque',
    'credit agricole',
    'crédit agricole',
    'societe generale',
    'société générale',
    'bnp',
    'lcl',
    'caisse d\'epargne',
    'caisse d\'épargne',
    'tva',
    'urssaf',
    'impôt',
    'impot',
    'facture',
    'invoice',
    'justificatif',
    'reçu',
    'recu'
  ];
  for (var i = 0; i < neverExcludePatterns.length; i++) {
    if (haystack.indexOf(neverExcludePatterns[i]) !== -1) {
      return { excluded: false, rule: '' };
    }
  }

  var exclusionRules = [
    { rule: 'EXCL_NEWSLETTER', patterns: ['newsletter', 'unsubscribe', 'se désabonner', 'se desabonner', 'désinscription', 'desinscription'] },
    { rule: 'EXCL_PUBLICITE', patterns: ['publicité', 'publicite', 'promo', 'promotion', 'offre spéciale', 'offre speciale', 'soldes', 'black friday'] },
    { rule: 'EXCL_INDEED_JOB_ALERT', patterns: ['indeed', 'job alert', 'alerte emploi', 'nouveaux emplois', 'votre alerte emploi'] },
    { rule: 'EXCL_GOOGLE_NOTIFICATION', patterns: ['google alerts', 'google account', 'security alert', 'alerte de sécurité', 'alerte de securite', 'google no-reply', 'no-reply@google.com', 'accounts.google.com'] },
    { rule: 'EXCL_CALENDAR_INVITE', patterns: ['invite.ics', 'invitation calendrier', 'calendar invite', 'google calendar', 'microsoft teams', 'teams meeting', 'join the meeting', 'réunion teams', 'reunion teams'] }
  ];

  for (var r = 0; r < exclusionRules.length; r++) {
    var rule = exclusionRules[r];
    for (var p = 0; p < rule.patterns.length; p++) {
      if (haystack.indexOf(rule.patterns[p]) !== -1) {
        return { excluded: true, rule: rule.rule };
      }
    }
  }

  return { excluded: false, rule: '' };
}

// Classe un email selon des règles métier simples.
function classifyEmail_(from, to, subject, bodyPreview) {
  var lowFrom = from.toLowerCase();
  var lowSub = subject.toLowerCase();
  var lowBody = bodyPreview.toLowerCase();

  var classification = {
    entity: '',
    activity: '',
    category: '',
    dossier: '',
    status: 'À extraire',
    confiance: 'Moyenne',
    rule: '',
    action: ''
  };

  var exclusionResult = isClearlyExcluded_(from, to, subject, bodyPreview);
  if (exclusionResult && exclusionResult.excluded) {
    classification.status = 'Exclu';
    classification.rule = exclusionResult.rule || 'EXCLUSION_RULE';
    return classification;
  }

  if (lowFrom.indexOf('enedis-noreply') !== -1 && lowSub.indexOf('facture') !== -1) {
    classification.entity = 'SAS';
    classification.activity = 'Solaire';
    classification.category = 'Facture';
    classification.dossier = buildSasDossierPath_('01_Solaire/Enedis_Factures');
    classification.rule = 'ENEDIS_FACTURE';
    classification.confiance = 'Forte';
    return classification;
  }
  if (lowFrom.indexOf('afc-agence prod') !== -1 || lowSub.indexOf('publication semestrielle') !== -1) {
    classification.entity = 'SAS';
    classification.activity = 'Solaire';
    classification.category = 'Données de comptage';
    classification.dossier = buildSasDossierPath_('01_Solaire/Enedis_Comptage');
    classification.rule = 'ENEDIS_COMPTAGE';
    classification.confiance = 'Forte';
    return classification;
  }
  if (lowFrom.indexOf('isolarcloud') !== -1) {
    classification.entity = 'SAS';
    classification.activity = 'Solaire';
    classification.category = 'Contrat';
    classification.dossier = buildSasDossierPath_('01_Solaire/iSolarCloud');
    classification.rule = 'ISOLARCLOUD';
    classification.confiance = 'Moyenne';
    return classification;
  }
  if (lowFrom.indexOf('oa-solaire') !== -1 || lowSub.indexOf('obligation d\'achat') !== -1) {
    classification.entity = 'SAS';
    classification.activity = 'Solaire';
    classification.category = 'OA';
    classification.dossier = buildSasDossierPath_('01_Solaire/EDF_OA');
    classification.rule = 'OA_SOLAIRE';
    classification.confiance = 'Moyenne';
    return classification;
  }
  if (lowFrom.indexOf('bfc') !== -1 && lowSub.indexOf('solaire') !== -1) {
    classification.entity = 'SAS';
    classification.activity = 'Solaire';
    classification.category = 'BFC';
    classification.dossier = buildSasDossierPath_('01_Solaire/BFC_Solaire');
    classification.rule = 'BFC_SOLAIRE';
    classification.confiance = 'Moyenne';
    return classification;
  }

  if (lowFrom.indexOf('wecasa') !== -1) {
    classification.entity = 'Perso_Locations';
    classification.activity = 'Airbnb_Aubervilliers';
    classification.category = 'Frais_Ménage';
    classification.dossier = buildPersoLocationDossierPath_('01_Airbnb_Aubervilliers/Frais_Ménage_Wecasa');
    classification.rule = 'WECASA_AUBERVILLIERS';
    classification.confiance = 'Forte';
    return classification;
  }
  if (lowFrom.indexOf('airbnb') !== -1 || lowSub.indexOf('airbnb') !== -1) {
    if (lowSub.indexOf('aubervilliers') !== -1) {
      classification.entity = 'Perso_Locations';
      classification.activity = 'Airbnb_Aubervilliers';
      classification.category = 'Réservation';
      classification.dossier = buildPersoLocationDossierPath_('01_Airbnb_Aubervilliers/Airbnb_Réservations');
    } else if (lowSub.indexOf('annoire') !== -1) {
      classification.entity = 'Perso_Locations';
      classification.activity = 'Airbnb_Annoire';
      classification.category = 'Réservation';
      classification.dossier = buildPersoLocationDossierPath_('02_Airbnb_Annoire/Airbnb_Réservations');
    } else {
      classification.entity = 'Perso_Locations';
      classification.activity = 'Airbnb';
      classification.category = 'Réservation';
      classification.dossier = buildPersoLocationDossierPath_('01_Airbnb_Aubervilliers/Airbnb_Réservations');
      classification.status = 'À vérifier';
      classification.action = 'Confirmer si Airbnb concerne Aubervilliers ou Annoire';
      classification.confiance = 'Faible';
    }
    classification.rule = 'AIRBNB';
    return classification;
  }

  var tvaKeywords = ['tva', 'taxe', 'urssaf', 'impôt', 'impot', 'cotisation'];
  for (var j = 0; j < tvaKeywords.length; j++) {
    if (lowSub.indexOf(tvaKeywords[j]) !== -1) {
      classification.entity = 'SAS';
      classification.activity = 'Fiscalité';
      classification.category = 'TVA/Impôts/Charges';
      classification.dossier = buildSasDossierPath_('03_TVA_Impots');
      classification.rule = 'FISCALITE';
      classification.confiance = 'Moyenne';
      return classification;
    }
  }

  if (lowSub.indexOf('facture') !== -1 || lowSub.indexOf('invoice') !== -1) {
    var fournisseurs = ['imc', 'fournisseur', 'sas bigueure brothers', 'earl des acacias'];
    for (var k = 0; k < fournisseurs.length; k++) {
      if (lowFrom.indexOf(fournisseurs[k]) !== -1) {
        classification.entity = 'SAS';
        classification.activity = 'Fournisseurs';
        classification.category = 'Facture';
        classification.dossier = buildSasDossierPath_('05_Factures_Fournisseurs');
        classification.rule = 'FACTURE_FOURNISSEUR';
        classification.confiance = 'Moyenne';
        return classification;
      }
    }
  }

  if (lowFrom.indexOf('alliancecomtoise.cerfrance.fr') !== -1 || lowFrom.indexOf('vbardoux') !== -1 || lowFrom.indexOf('sligey') !== -1) {
    classification.entity = 'SAS';
    classification.activity = 'Comptable';
    classification.category = 'Échanges comptables';
    classification.dossier = buildSasDossierPath_('06_Bilan_Comptable');
    classification.rule = 'CERFRANCE_COMPTABLE';
    classification.status = 'À vérifier';
    classification.action = 'Vérifier l’objet et classer correctement les pièces jointes (échanges comptables)';
    classification.confiance = 'Moyenne';
    return classification;
  }

  var persoPatterns = ['licitor', 'congé payé', 'rh', 'nomination'];
  for (var m = 0; m < persoPatterns.length; m++) {
    if (lowSub.indexOf(persoPatterns[m]) !== -1 || lowFrom.indexOf(persoPatterns[m]) !== -1) {
      classification.entity = 'Perso_Vie';
      classification.activity = 'Administratif';
      classification.category = 'Notification';
      classification.dossier = buildPersoVieDossierPath_('99_A_Verifier');
      classification.rule = 'PERSO_NOTIF';
      classification.status = 'À vérifier';
      classification.action = 'Vérifier si ce document concerne la vie personnelle';
      classification.confiance = 'Faible';
      return classification;
    }
  }

  classification.entity = '';
  classification.activity = '';
  classification.category = '';
  classification.dossier = '';
  classification.status = 'À vérifier';
  classification.confiance = 'Faible';
  classification.rule = 'DEFAULT';
  classification.action = 'Examiner cet e-mail manuellement pour déterminer son classement.';
  return classification;
}

// Construit le chemin dossier SAS.
function buildSasDossierPath_(subPath) {
  var year = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy');
  return '01_SAS_BIGUEURE_BROTHERS/' + year + '/' + subPath;
}

// Construit le chemin dossier locations perso.
function buildPersoLocationDossierPath_(subPath) {
  var year = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy');
  return '02_PERSO_LOCATIONS/' + year + '/' + subPath;
}

// Construit le chemin dossier vie perso.
function buildPersoVieDossierPath_(subPath) {
  var year = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy');
  return '03_PERSO_VIE/' + year + '/' + subPath;
}
