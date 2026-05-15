// Détermine si un email est clairement à exclure.
function isClearlyExcluded_(from, to, subject, bodyPreview) {
  var lowFrom = (from || '').toLowerCase();
  var lowTo = (to || '').toLowerCase();
  var lowSub = (subject || '').toLowerCase();
  var lowBody = (bodyPreview || '').toLowerCase();
  var haystack = [lowFrom, lowTo, lowSub, lowBody].join(' ');

  var auroreForwardingSignals = ['aurorebigus', 'aurorebigus.com', 'orebigus', 'rebigus', 'rébigus'];
  for (var a = 0; a < auroreForwardingSignals.length; a++) {
    if (haystack.indexOf(auroreForwardingSignals[a]) !== -1) {
      return { excluded: false, rule: '' };
    }
  }

  var strongAccountingSignals = [
    'cerfrance',
    'alliancecomtoise.cerfrance.fr',
    'enedis',
    'edf',
    'msa',
    'wecasa',
    'airbnb',
    'tva',
    'urssaf',
    'relevé bancaire',
    'releve bancaire',
    'attestation',
    'tiers payant',
    'impôt',
    'impot',
    'facture',
    'invoice',
    'justificatif',
    'reçu',
    'recu',
    'sas bigueure brothers',
    'crédit agricole vrai document bancaire',
    'credit agricole vrai document bancaire',
    'banque postale vrai document bancaire',
    'annoire',
    'photovoltaïque',
    'photovoltaique',
    'consuel',
    'groupama-epargne',
    'groupama-epargne.fr',
    'intéressement',
    'interessement',
    'épargne salariale',
    'epargne salariale',
    'people-doc.com',
    'chimirec.fr',
    'contrat',
    'document rh',
    'attestation employeur',
    'devis',
    'commande utile',
    'réservation',
    'reservation',
    'booking confirmation',
    'pulse',
    'colissimo',
    'comptable',
    'google apps script',
    'apps script',
    'script function',
    'summary of failures',
    'projet ia',
    'rdv santé',
    'rdv sante',
    'kiné',
    'kine'
  ];
  for (var s = 0; s < strongAccountingSignals.length; s++) {
    if (haystack.indexOf(strongAccountingSignals[s]) !== -1) {
      return { excluded: false, rule: '' };
    }
  }

  var preserveSignals = ['facture', 'invoice', 'reçu', 'recu', 'devis', 'contrat', 'justificatif', 'commande utile', 'réservation', 'reservation', 'booking confirmation'];
  var marketingSignals = ['publicité', 'publicite', 'newsletter', 'promotion', 'promo', 'avis produit', "demande d'avis", 'votre commande a-t-elle répondu à vos attentes', 'votre commande a-t-elle repondu a vos attentes', 'recommandation'];
  var jobSignals = ["offre d'emploi", 'candidature', 'recrutement', 'job'];
  var excludedIfMarketing = [
    { rule: 'EXCL_DJI_MARKETING', pattern: 'e.dji.com', signals: marketingSignals },
    { rule: 'EXCL_HOTMAIL_MARKETING_JOB', pattern: 'hotmail.com', signals: marketingSignals.concat(jobSignals) },
    { rule: 'EXCL_NEW_MARKETING', pattern: 'new', signals: ['publicité', 'publicite', 'newsletter', 'promotion', 'promo'] }
  ];
  for (var m = 0; m < excludedIfMarketing.length; m++) {
    var conditionalRule = excludedIfMarketing[m];
    if (haystack.indexOf(conditionalRule.pattern) !== -1 && containsAnyClassification_(haystack, conditionalRule.signals)) {
      return { excluded: true, rule: conditionalRule.rule };
    }
  }

  if (haystack.indexOf('amazon.fr') !== -1 && containsAnyClassification_(haystack, marketingSignals) && !containsAnyClassification_(haystack, preserveSignals)) {
    return { excluded: true, rule: 'EXCL_AMAZON_MARKETING_AVIS' };
  }

  var exclusionRules = [
    {
      rule: 'EXCL_PRO_GROUPAMA',
      patterns: ['groupama.com', 'groupama-ne.fr', 'groupama-ra.fr', 'groupama.com transféré vers gmail', 'groupama-ne transféré vers gmail', 'groupama-ra transféré vers gmail'],
      exceptions: ['groupama-epargne', 'groupama-epargne.fr', 'intéressement', 'interessement', 'épargne salariale', 'epargne salariale']
    },
    {
      rule: 'EXCL_MARKETING_TRAVEL',
      patterns: ['email.minorhotel', 'minorhotel'],
      exceptions: ['facture', 'invoice', 'reçu', 'recu', 'réservation', 'reservation']
    },
    { rule: 'EXCL_NEWSLETTER', patterns: ['newsletter', 'unsubscribe', 'se désabonner', 'se desabonner', 'désinscription', 'desinscription'] },
    { rule: 'EXCL_PUBLICITE', patterns: ['publicité', 'publicite', 'promo', 'promotion', 'offre spéciale', 'offre speciale', 'soldes', 'black friday', 'avis produit', 'demande d\'avis', 'votre commande a-t-elle répondu à vos attentes', 'votre commande a-t-elle repondu a vos attentes'] },
    { rule: 'EXCL_INDEED_JOB_ALERT', patterns: ['indeed', 'job alert', 'alerte emploi', 'alerte emploi indeed', 'nouveaux emplois', 'votre alerte emploi', 'donotreply@jobalert.indeed.com', 'offres de "', 'recherche un/e', 'indeed.com'] },
    {
      rule: 'EXCL_MARKETING_BRANDS',
      patterns: ['github', 'noreply@github.com', 'mermaid.ai', 'hello@mermaid.ai', 'arlettie', 'annonce@amazon.fr', 'business.amazon.fr', 'instagram', 'decantalo', 'patchplants', 'leboncoin newsletters', 'info@news.leboncoin.fr', 'no.reply@leboncoin.fr', 'thecoolrepublic', 'fnac photo', 'printemps', 'filovent', 'lovable.dev', 'baobabcollection', 'chatgpt task update', 'noreply@tm.openai.com', 'bonsoirs', 'eurostar promo', 'no-reply@e.eurostar.com', 'code41', 'code41.com', 'team@code41.com', 'chrono manufacture', 'nb24', 'claudio d\'amore', 'edenredplus.com', 'noreply@edenredplus.com', 'ucon acrobatics', 'programme force sculpt', 'rappel et motivation', 'email.boulanger.com', 'mails.interflora.fr', 'farfetch@email.farfetch.com', 'email.farfetch.com', 'hello@noo.ma', 'noo.ma', 'gimber.com', 'figma.com', 'actu.fdj.fr', 'mail.leroymerlin.fr', 'leroymerlin', 'artsper.com', 'bolia.com', 'thebicestercollection.com', 'lavallee', 'pointp-et-moi.fr', 'pointp', 'lexology.com', 'afje newsstand', 'pinterest', 'sauvons l\'europe', 'sauvonsleurope', 'sncf-connect.com', 'sncf connect', 'flying blue', 'info-flyingblue.com', 'netflix', 'members.netflix.com', 'maddyness', 'maddynews', 'cryptoast', 'contact@cryptoast.fr', 'adidas', 'fr-news.adidas.com', 'cegos', 'agro-media.fr', 'koanna.com', 'candor.be', 'blog-ge', 'blog-gestion-de-projet', 'afa-arbitrage.com', 'apoticari.com', 'ona.com', 'email.play', 'dataiku.com', 'email.experteer.com', 'france.marionnaud.pari', 'marionnaud', 'altermundi.com', 'merci-merci.link', 'cfoc.fr', 'fnac.com', 'theatreonline.com', 'actu-monclub.totalenergie', 'brick', 'crm.'],
      exceptions: ['facture', 'invoice', 'reçu', 'recu', 'devis', 'contrat', 'justificatif', 'commande utile', 'réservation', 'reservation', 'booking confirmation']
    },
    {
      rule: 'EXCL_ALERTS_IMMO_EMPLOI',
      patterns: ['alertes.seloger.com', 'seloger', 'bienici', 'nouvelle annonce', 'nouvelles annonces', 'baisse de prix', 'dernières exclusivités', 'dernieres exclusivites', 'annonce immobilière', 'annonce immobiliere', 'profils.org', 'groupecreditagricole@profils.org', 'offre correspond à vos recherches', 'alerte emploi', 'candidature', 'poste de', 'recrutement', 'noreply2.job', 'offre d\'emploi', 'job'],
      exceptions: ['contrat', 'bulletin', 'document rh', 'attestation employeur']
    },
    {
      rule: 'EXCL_PRO_NOTIFICATIONS',
      patterns: ['engage.mail.micro', 'viva engage', 'microsoft viva', 'yammer'],
      exceptions: ['facture', 'contrat']
    },
    { rule: 'EXCL_GOOGLE_NOTIFICATION', patterns: ['google alerts', 'google account', 'security alert', 'alerte de sécurité', 'alerte de securite', 'google no-reply', 'no-reply@google.com', 'accounts.google.com'] },
    { rule: 'EXCL_CALENDAR_INVITE', patterns: ['invite.ics', 'invitation calendrier', 'calendar invite', 'google calendar', 'microsoft teams', 'teams meeting', 'join the meeting', 'réunion teams', 'reunion teams'] }
  ];

  for (var r = 0; r < exclusionRules.length; r++) {
    var rule = exclusionRules[r];
    var matched = false;
    for (var p = 0; p < rule.patterns.length; p++) {
      if (haystack.indexOf(rule.patterns[p]) !== -1) {
        matched = true;
        break;
      }
    }
    if (!matched) {
      continue;
    }

    if (rule.exceptions && rule.exceptions.length > 0) {
      for (var e = 0; e < rule.exceptions.length; e++) {
        if (haystack.indexOf(rule.exceptions[e]) !== -1) {
          return { excluded: false, rule: '' };
        }
      }
    }

    return { excluded: true, rule: rule.rule };
  }

  return { excluded: false, rule: '' };
}

function containsAnyClassification_(text, patterns) {
  for (var i = 0; i < patterns.length; i++) {
    if (text.indexOf(patterns[i]) !== -1) {
      return true;
    }
  }
  return false;
}

// Classe un email selon des règles métier simples.
function classifyEmail_(from, to, subject, bodyPreview) {
  var classification = {
    entity: '',
    activity: '',
    category: '',
    dossier: '',
    status: 'À extraire',
    confiance: 'Moyenne',
    rule: '',
    action: '',
    createAction: true
  };

  var exclusionResult = isClearlyExcluded_(from, to, subject, bodyPreview);
  if (exclusionResult && exclusionResult.excluded) {
    classification.status = 'Exclu';
    classification.confiance = 'Forte';
    classification.rule = exclusionResult.rule || 'EXCLUSION_RULE';
    classification.action = '';
    classification.createAction = false;
    return classification;
  }

  var lowFrom = from.toLowerCase();
  var lowSub = subject.toLowerCase();
  var lowBody = bodyPreview.toLowerCase();

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
