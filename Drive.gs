// Sauvegarde des pièces jointes dans l’arborescence Drive cible.
function saveAttachmentsToDrive_(attachments, date, from, subject, dossierPath) {
  if (!dossierPath) {
    return;
  }
  var rootFolder = getOrCreateFolder_(DriveApp.getRootFolder(), '00_Automatisation_Compta_Admin');
  var yearFolder = getOrCreateFolder_(rootFolder, dossierPath.split('/')[0]);
  var subFolders = dossierPath.split('/').slice(1);
  var currentFolder = yearFolder;
  subFolders.forEach(function(name) {
    currentFolder = getOrCreateFolder_(currentFolder, name);
  });
  attachments.forEach(function(att) {
    var cleanName = buildCleanFileName_(date, from, subject, att.getName());
    if (!fileAlreadyExists_(currentFolder, cleanName)) {
      currentFolder.createFile(att.copyBlob()).setName(cleanName);
    }
  });
}

// Construit un nom de fichier propre et lisible.
function buildCleanFileName_(date, from, subject, originalName) {
  var formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var cleanSender = cleanText_(extractEmailOrName_(from)).substring(0, 40);
  var cleanSubject = cleanText_(subject).substring(0, 80);
  var cleanOriginalName = cleanText_(originalName).substring(0, 80);
  return [formattedDate, cleanSender, cleanSubject, cleanOriginalName].filter(Boolean).join(' - ');
}

// Vérifie l’existence d’un fichier homonyme dans un dossier.
function fileAlreadyExists_(folder, fileName) {
  var files = folder.getFilesByName(fileName);
  return files.hasNext();
}

// Retourne/crée un sous-dossier Drive.
function getOrCreateFolder_(parent, name) {
  var folders = parent.getFoldersByName(name);
  if (folders.hasNext()) {
    return folders.next();
  }
  return parent.createFolder(name);
}
