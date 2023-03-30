// Rechnung erzeugen und ggf. per E-Mail versenden
function reCreate() { 
  const nameRe = "Rechnungen";
  const nameEin = "Einstellungen";
  const nameAdr = "Adressen";
  const colPos = 11;                                // ab Spalte 10 beginnen die Rechnungspositionen
  const colOff = 6;                                 // Offset zur nächsten Bestellposition
  if (!uiMsgBox("Tatsächlich Rechnung erzeugen?"))  // Sicherheitsabfrage, es kann noch abgebrochen werden
    return;
  if (reProp.getProperty('reAblageOK') != 'true'){  // Ablage in Drive ok?
    SpreadsheetApp.getUi().alert("Bitte Ablage prüfen, die Rechnung wurde nicht erzeugt!");
    return;
  }
  try {
    let sActive = SpreadsheetApp.getActiveSheet();
    if (sActive.getName() != nameRe){               // Sheet Rechnungen geöffnet
      SpreadsheetApp.getUi().alert(`Eine Rechnung kann nur im Sheet ${nameRe} erzeugt werden!`);
      return;
    }
    let aRow = sActive.getActiveCell().getRow();
    // aktive Rechnungszeile ausgefüllt?
    if (aRow < 2 || sActive.getRange(aRow, 2).isBlank() || sActive.getRange(aRow, 3).isBlank() || 
                    sActive.getRange(aRow, 4).isBlank() || sActive.getRange(aRow, 7).isBlank()){
      SpreadsheetApp.getUi().alert("Aktive Rechnungszeile erst komplett ausfüllen!");
      return;
    }
    let reDatum = Utilities.formatDate(sActive.getRange(aRow, 7).getValue(), Session.getScriptTimeZone(), "dd.MM.yyyy");
    let reName = sActive.getRange(aRow, 2).getValue().toString() + '_' + sActive.getRange(aRow, 3).getValue() + 
                 '_' + reDatum;
    let fileVorlage = DriveApp.getFileById(reProp.getProperty('reIdVorlage'));
    let folDocs = DriveApp.getFolderById(reProp.getProperty('reIdFolDocs'));
    let folPdf = DriveApp.getFolderById(reProp.getProperty('reIdFolPdf'));
    let sEin = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nameEin);
    let sAdr = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nameAdr);
    // falls das Rechnungsfile schon besteht -> in den Papierkorb
    let files = folDocs.getFilesByName(reName);
    let file;
    while (files.hasNext()){
      file = files.next();
      file.setTrashed(true);                // Docs in den Papierkorb
      break;
    }
    files = folPdf.getFilesByName(reName);
    while (files.hasNext()){
      file = files.next();
      file.setTrashed(true);                // Pdf in den Papierkorb
      break;
    }
    // Vorlage kopieren und unter dem Rechnungsnamen im Docs-Ordner ablegen, return vom Typ File
    let rDatei = fileVorlage.makeCopy(reName, folDocs);
    // Zugriff auf das neu erstellte Dokument
    let re = DocumentApp.openById(rDatei.getId());
    // ab hier wird die zuvor kopierte Vorlage ausgefüllt
    re.getBody().replaceText("{absVN}", sAdr.getRange(2, 7).getValue());
    re.getBody().replaceText("{absNN}", sAdr.getRange(2, 8).getValue());
    re.getBody().replaceText("{absUnt}", sAdr.getRange(2, 2).getValue());
    re.getBody().replaceText("{absStrasse}", sAdr.getRange(2, 11).getValue());
    re.getBody().replaceText("{absOrt}", sAdr.getRange(2, 10).getValue());
    





    re.saveAndClose();
    Logger.log("reName = " + reName);
  }
  catch(error){
    SpreadsheetApp.getUi().alert("catch1 in reCreate = " + error);
    return;
  }




  

  Logger.log("Ende Rechnung");



}

// Prüfung, ob Ablagesystem in Google Drive ok, Rückgabeobjekt ret
function reCheckAblage() {
  const root = "Meine Ablage/prog/gas/01_GasLib/"
  let iDr = drInfo(root + "Vorlagen", "VL_Rechnung"); // Vorlage vorhanden?
  if (iDr.fileEx)
    reProp.setProperty('reIdVorlage', iDr.fileId);
  else
    return false;
  iDr = drInfo(root + "Rechnungen_DOCS");             // Ordner für docs vorhanden?
  if (iDr.folEx)
    reProp.setProperty('reIdFolDocs', iDr.folId);
  else
    return false;
  iDr = drInfo(root + "Rechnungen_PDF");             // Ordner für PDFs vorhanden?
  if (iDr.folEx)
    reProp.setProperty('reIdFolPdf', iDr.folId);
  else
    return false;
  return true;
}


