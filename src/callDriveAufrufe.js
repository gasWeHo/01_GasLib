function old1_drAufruf() {
  //let info = drFiles("Meine Ablage/prog/gas/02_Drive", "Drive");
  //info.forEach((val) => {Logger.log(JSON.stringify(val))});

  let rootInfo = JSON.stringify(drRoot());      // https://drive.google.com/drive/my-drive
  //Logger.log("Rootinfo = " + rootInfo);
  let info = drInfo("Meine Ablage/prog/gas/02_Drive", "Drive");
  //Logger.log("drInfo = " + JSON.stringify(info));
  let iDr = drCreateFolder("Meine Ablage/prog/gas/02_Drive", "Pedro");
  if (iDr > 0)
    Logger.log("ging gut mit dem createn");
  else
    Logger.log("Fehlercode = " + iDr);

  let iD = drTrashFolder("Meine Ablage/prog/gas/02_Drive", "Pedro");
  if (iD > 0)
    Logger.log("ging gut mit dem removen");
  else
    Logger.log("Fehlercode = " + iD);
}

function old2_drAufruf() {
  let r = drCopyFolder("Meine Ablage/prog/gas/02_Drive", "Pedro", "Meine Ablage/prog/input");
  if (r > 0)
    Logger.log("Kopieren war wohl erfolgreich");
  if (r == -1)
    Logger.log("Fehlende Aufrufparameter");
  if (r == -2)
    Logger.log("Quell-Folder existiert nicht");
  if (r == -3)
    Logger.log("Ziel Parent-Folder existiert nicht");
  if (r == -4)
    Logger.log("Ziel-Folder existiert schon");
  if (r == -5)
    Logger.log("Ziel-Folder konnte nicht angelegt werden");
  if (r == -6)
    Logger.log("something wrong");
}

function old3_drAufruf() {
  let i = drFileCopy("Meine Ablage/prog/gas/02_Drive", "Drive", "Meine Ablage/prog/input", "Drive2");
  if (i == 1)
    Logger.log("kopieren war erfolgreich")
  else
    Logger.log("kopieren war NICHT erfolgreich")

  i= drFileTrash("Meine Ablage/prog/input", "Drive2");
  if (i == 1)
    Logger.log("löschen war erfolgreich")
  else
    Logger.log("löschen war NICHT möglich")
}

function old4_drAufruf() {
  let i = drPdfCreate("Meine Ablage/prog/input", "Code", "Meine Ablage", "pdfCode");
  if (i == 1)
    Logger.log("Pdf erzeugen war erfolgreich")
  else
    Logger.log("Pdf erzeugen war NICHT erfolgreich")
}

function old5_drAufruf() {
  let i = drCreateFile("Meine Ablage/prog/input", "Werner_create", "Slides");
  if (i == 1)
    Logger.log("File erzeugen war erfolgreich")
  else
    Logger.log("File erzeugen war NICHT erfolgreich")
}

function old6_drAufruf() {
  let i = drCreateShortCut("Meine Ablage/prog/gas/02_Drive/Pedro", "Meine Ablage/prog/input", "scPedro", "");
  if (i == 1){
    Logger.log("Shortcut erzeugen war erfolgreich");
  }
  else
    Logger.log("File erzeugen war NICHT erfolgreich, Fehlercode = " + i);
}
