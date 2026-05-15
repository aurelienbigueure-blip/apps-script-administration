// Nettoie un texte pour l’utiliser dans des noms de fichiers.
function cleanText_(value) {
  if (!value) return '';
  return String(value)
    .replace(/[\\/:*?"<>|#%{}~&]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

// Extrait l’email (ou nom) de l’expéditeur.
function extractEmailOrName_(from) {
  if (!from) return '';
  var emailMatch = from.match(/<([^>]+)>/);
  if (emailMatch && emailMatch[1]) {
    return emailMatch[1];
  }
  return from;
}

// Retourne l’extension en minuscules.
function getExtension_(name) {
  if (!name || name.indexOf('.') === -1) {
    return '';
  }
  return name.split('.').pop().toLowerCase();
}

// Entoure un label Gmail de guillemets.
function quoteGmailLabel_(label) {
  return '"' + label + '"';
}
