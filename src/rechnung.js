// Rechnung erzeugen und ggf. per E-Mail versenden
function reCreate() {
  const nameRe = "Rechnungen";
  const nameEin = "Einstellungen";
  const nameAdr = "Adressen";
  const dh = "Damen und Herren";
  const herr = "Herr";
  const frau = "Frau"
  const zHd = "z.Hd. ";
  const colPos = 11;                                // ab Spalte 10 beginnen die Rechnungspositionen
  const colOff = 6;                                 // Offset zur nächsten Bestellposition
  const statusVersendet = "versendet";              // direkt nach Rechnungserzeugung

  try {
    if (!uiMsgBox("Tatsächlich Rechnung erzeugen?"))  // Sicherheitsabfrage, es kann noch abgebrochen werden
      return;
    if (reProp.getProperty('reAblageOK') != 'true') {  // Ablage in Drive ok?
      SpreadsheetApp.getUi().alert("Bitte Ablage prüfen, die Rechnung wurde nicht erzeugt!");
      return;
    }
    let sActive = SpreadsheetApp.getActiveSheet();
    if (sActive.getName() != nameRe) {               // Sheet Rechnungen geöffnet
      SpreadsheetApp.getUi().alert(`Eine Rechnung kann nur im Sheet ${nameRe} erzeugt werden!`);
      return;
    }
    let aRow = sActive.getActiveCell().getRow();
    // aktive Rechnungszeile ausgefüllt?
    if (aRow < 2 || sActive.getRange(aRow, 2).isBlank() || sActive.getRange(aRow, 3).isBlank() ||
      sActive.getRange(aRow, 4).isBlank() || sActive.getRange(aRow, 7).isBlank()) {
      SpreadsheetApp.getUi().alert("Aktive Rechnungszeile erst komplett ausfüllen!");
      return;
    }
    let sEin = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nameEin);
    let sAdr = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nameAdr);
    // Ermitteln der Zeilen-Nr des Rechnungsempfängers im Sheet Adressen
    let adrRow = 2 + shGetRowUsingfindIndex(sActive.getRange(aRow, 3).getValue(), sAdr.getRange(2, 2, sAdr.getLastRow(), 1));
    if (adrRow <= 1) {   // der ausgewählte Rechnungsempfänger wurde nicht gefunden
      SpreadsheetApp.getUi().alert("Der Rechnungsempfänger wurde im Sheet Adressen nicht gefunden, es wird keine Rechnung erzeugt!");
      return;
    }
    let reDatum = Utilities.formatDate(sActive.getRange(aRow, 7).getValue(), Session.getScriptTimeZone(), "dd.MM.yyyy");
    let reZiel = Utilities.formatDate(sActive.getRange(aRow, 8).getValue(), Session.getScriptTimeZone(), "dd.MM.yyyy");
    let reNr = sActive.getRange(aRow, 2).getValue().toString();
    let reName = reNr + '_' + sActive.getRange(aRow, 3).getValue() +
      '_' + reDatum;
    let fileVorlage = DriveApp.getFileById(reProp.getProperty('reIdVorlage'));
    let folDocs = DriveApp.getFolderById(reProp.getProperty('reIdFolDocs'));
    let folPdf = DriveApp.getFolderById(reProp.getProperty('reIdFolPdf'));
    // falls das Rechnungsfile schon besteht -> in den Papierkorb
    let files = folDocs.getFilesByName(reName);
    let file;
    while (files.hasNext()) {
      file = files.next();
      file.setTrashed(true);                // Docs in den Papierkorb
      break;
    }
    files = folPdf.getFilesByName(reName);
    while (files.hasNext()) {
      file = files.next();
      file.setTrashed(true);                // Pdf in den Papierkorb
      break;
    }
    // Vorlage kopieren und unter dem Rechnungsnamen im Docs-Ordner ablegen, return vom Typ File
    let rDatei = fileVorlage.makeCopy(reName, folDocs);
    // Zugriff auf das neu erstellte Dokument
    let re = DocumentApp.openById(rDatei.getId());
    // ab hier wird die zuvor kopierte Vorlage ausgefüllt
    let absVN = sAdr.getRange(2, 7).getValue();
    let absNN = sAdr.getRange(2, 8).getValue();
    let reBetreff = sActive.getRange(aRow, 4).getValue();
    re.getBody().replaceText("{absVN}", absVN);
    re.getBody().replaceText("{absNN}", absNN);
    re.getBody().replaceText("{absUnt}", sAdr.getRange(2, 2).getValue());
    re.getBody().replaceText("{absStrasse}", sAdr.getRange(2, 11).getValue());
    re.getBody().replaceText("{absOrt}", sAdr.getRange(2, 9).getValue() + " " + sAdr.getRange(2, 10).getValue());
    re.getBody().replaceText("{absTel}", sAdr.getRange(2, 12).getValue());
    re.getBody().replaceText("{absMail}", sAdr.getRange(2, 13).getValue());
    re.getBody().replaceText("{reNr}", reNr);
    re.getBody().replaceText("{reDat}", reDatum);
    re.getBody().replaceText("{reZiel}", reZiel);
    re.getBody().replaceText("{reBetreff}", reBetreff);
    re.getBody().replaceText("{empUnt}", sActive.getRange(aRow, 3).getValue());
    re.getBody().replaceText("{klMwst}", (100 * sEin.getRange(2, 2).getValue()).toFixed(0) + " %");
    re.getBody().replaceText("{grMwst}", (100 * sEin.getRange(3, 2).getValue()).toFixed(0) + " %");

    re.getBody().replaceText("{empStrasse}", sAdr.getRange(adrRow, 11).getValue());
    re.getBody().replaceText("{empOrt}", sAdr.getRange(adrRow, 9).getValue() + " " + sAdr.getRange(adrRow, 10).getValue());
    re.getBody().replaceText("{kuNr}", sAdr.getRange(adrRow, 4).getValue());
    let empAnr1 = "";                       // Anrede 1
    let empAnr2 = "";                       // Anrede 2
    let titel = " ";                        // Titel
    let anr = sAdr.getRange(adrRow, 5).getValue();
    let nn = sAdr.getRange(adrRow, 8).getValue();
    if (!sAdr.getRange(adrRow, 6).isBlank() && anr != dh) {
      titel += sAdr.getRange(adrRow, 6).getValue();
      titel += " ";
    }
    if (anr == frau) {
      empAnr1 = zHd + frau + titel + nn;
      empAnr2 = " " + frau + titel + nn;
    }
    else if (anr == herr) {
      empAnr1 = zHd + herr + "n" + titel + nn;
      empAnr2 = "r " + herr + titel + nn;
    }
    else {
      empAnr2 = " " + dh;
    }
    re.getBody().replaceText("{empAnr1}", empAnr1);
    re.getBody().replaceText("{empAnr2}", empAnr2);

    re.getFooter().replaceText("{finAmt}", sAdr.getRange(2, 16).getValue());
    re.getFooter().replaceText("{steuerNr}", sAdr.getRange(2, 17).getValue());
    re.getFooter().replaceText("{bank}", sAdr.getRange(2, 18).getValue());
    re.getFooter().replaceText("{iban}", sAdr.getRange(2, 19).getValue());
    re.getFooter().replaceText("{bic}", sAdr.getRange(2, 20).getValue());

    // ab hier werden die Rechnungspositionen eingetragen 
    let sumNetto = 0;
    let klMwstSum = 0;
    let grMwstSum = 0;
    let menge = 0;
    let einzelpreis = 0;
    let mwst = 0;
    let gesamtbetrag = 0;
    let gesamtpreis = 0;
    let cPos = colPos;
    let anzPos = 0;
    let klMwst = sEin.getRange(2, 2).getValue();
    let grMwst = sEin.getRange(3, 2).getValue();
    let pCurr = {
      style: "currency",
      currency: sEin.getRange(8, 1).getValue()                    // EUR
    };
    let lk = sEin.getRange(9, 1).getValue();         // Länderkennzeichen  de-DE
    let tbl = re.getBody().getTables()[2];           // Liste aller Tables im Body des Rechnungsdokumentes
    while (!sActive.getRange(aRow, cPos + 5).isBlank()) {
      menge = parseFloat(sActive.getRange(aRow, cPos + 1).getValue()).toFixed(2);
      einzelpreis = parseFloat(sActive.getRange(aRow, cPos + 4).getValue()).toFixed(2);
      mwst = parseFloat(sActive.getRange(aRow, cPos + 3).getValue()).toFixed(2);
      gesamtpreis = menge * einzelpreis;
      tbl.getCell(anzPos + 1, 0).setText(sActive.getRange(aRow, cPos).getValue());               // Positions-Nummer
      tbl.getCell(anzPos + 1, 1).setText(sActive.getRange(aRow, cPos + 1).getValue());             // Menge
      tbl.getCell(anzPos + 1, 2).setText(sActive.getRange(aRow, cPos + 2).getValue());             // Einheit
      tbl.getCell(anzPos + 1, 3).setText(sActive.getRange(aRow, cPos + 5).getValue());             // Leistung
      tbl.getCell(anzPos + 1, 4).setText(new Intl.NumberFormat(lk, pCurr).format(einzelpreis));  // Einzelpreis
      tbl.getCell(anzPos + 1, 5).setText((100 * mwst).toFixed(0) + " %");                          // MWST in %
      tbl.getCell(anzPos + 1, 6).setText(new Intl.NumberFormat(lk, pCurr).format(gesamtpreis));  // Gesamtpreis
      if (mwst == klMwst) {
        klMwstSum += mwst * gesamtpreis;
      }
      if (mwst == grMwst) {
        grMwstSum += mwst * gesamtpreis;
      }
      sumNetto += gesamtpreis;

      anzPos++;
      cPos += colOff;
    }
    gesamtbetrag = sumNetto + klMwstSum + grMwstSum;
    tbl.getCell(7, 6).setText(new Intl.NumberFormat(lk, pCurr).format(sumNetto));       // Summe Netto
    tbl.getCell(8, 6).setText(new Intl.NumberFormat(lk, pCurr).format(klMwstSum));      // Summe kleine MWST
    tbl.getCell(9, 6).setText(new Intl.NumberFormat(lk, pCurr).format(grMwstSum));      // Summe große MWST
    tbl.getCell(11, 6).setText(new Intl.NumberFormat(lk, pCurr).format(gesamtbetrag));  // Summe Brutto  
    // Rechnungsdatei schliessen
    re.saveAndClose();
    // pdf erzeugen
    let pdfBlob = re.getAs("application/pdf");
    pdfBlob.setName(reName);
    let filePdf = folPdf.createFile(pdfBlob);
    // Rechnung erstellt und versendet
    sActive.getRange(aRow, 6).setValue(statusVersendet);
    // pdf ggf. als E-Mail versenden
    if (sActive.getRange(aRow, 5).getValue()) {
      let mailFrom = absVN + " " + absNN;
      let mailCc = "";
      let mailBcc = "";
      let mailTo = sAdr.getRange(adrRow, 13).getValue();
      let subject = `Rechnung: ${reName} - ${reBetreff}`;
      let content = `Sehr geehrte${empAnr2},\n\nin der beiliegenden Mail-Anlage finden Sie Ihre Rechnung.\n\n\nMit freundlichen Grüßen\n\n${absVN} ${absNN}`;
      MailApp.sendEmail(mailTo, subject, content, {
        name: mailFrom,
        attachments: [filePdf],
        cc: mailCc,
        bcc: mailBcc
      });
    }
  }
  catch (error) {
    SpreadsheetApp.getUi().alert("catch1 in reCreate = " + error);
    return;
  }
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

// passt automatisch den Rechnungsstatus an
function reStatus(row, col, value, sName) {
  const nameRe = "Rechnungen";
  const colStatus = 6;
  const colBezahltAm = 9;
  const statusVersendet = "versendet";              // direkt nach Rechnungserzeugung
  const statusAbgeschlossen = "abgeschlossen";      // sobald Bezahlt am eingegeben wurde
  if (sName != nameRe || col != colBezahltAm)       // erfolgte die Bezahlt-Änderung im Sheet Rechnungen?
    return;
  let sRech = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nameRe);
  if (sRech.getRange(row, colStatus).getValue() != statusVersendet)  // nur abschliessen, wenn Status auf versendet steht
    return;
  sRech.getRange(row, colStatus).setValue(statusAbgeschlossen);
}

