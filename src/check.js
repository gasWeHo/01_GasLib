// Prüfung, ob Ablagesystem in Google Drive ok, Rückgabeobjekt ret
function checkAblageRechnung() {
  const root = "Meine Ablage/prog/gas/01_GasLib/"
  let iDr = drInfo(root + "Vorlagen", "VL_Rechnung"); // Vorlage vorhanden?
  if (iDr.fileEx)
    prop.setProperty('idVorlage', iDr.fileId);
  else
    return false;
  iDr = drInfo(root + "Rechnungen_DOCS");             // Ordner für docs vorhanden?
  if (iDr.folEx)
    prop.setProperty('idFolDocs', iDr.folId);
  else
    return false;
  iDr = drInfo(root + "Rechnungen_PDF");             // Ordner für PDFs vorhanden?
  if (iDr.folEx)
    prop.setProperty('idFolPdf', iDr.folId);
  else
    return false;
  return true;
}

function test() { 
  if (prop.getProperty('ablageOK') == 'true'){
    let file = DriveApp.getFileById(prop.getProperty('idVorlage'));
    Logger.log("Datei-Name = " + file.getName());
    let fol1 = DriveApp.getFolderById(prop.getProperty('idFolDocs'));
    Logger.log("Docs-Name = " + fol1.getName());
    let fol2 = DriveApp.getFolderById(prop.getProperty('idFolPdf'));
    Logger.log("Pdf-Name = " + fol2.getName());
  }
}