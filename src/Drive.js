function drAufruf() {
  let i = drGetFileCount("Meine Ablage/prog/gas/01_GasLib/Pedro/Hans");
  Logger.log("i = " + i);
}

function drCreateShortCut(folNameFrom="", folNameTo="", shortcutName="", fileNameFrom="") {
  // Shortcut erzeugen, erfordert die Drive API unter Dienste!!!
  if (folNameFrom == "" || folNameTo == "" || shortcutName == "")
    return (-1);              // fehlende Funktionsparameter
  let iDrD = drInfo(folNameTo, shortcutName);
  if (!iDrD.folEx || iDrD.fileEx)
    return (-2);              // Abbruch Ziel-Folder existiert nicht oder Shortcut existiert bereits
  let iDrS = drInfo(folNameFrom, fileNameFrom);
  if (!iDrS.folEx)
    return (-3);              // Abbruch Quell-Folder existiert nicht
  if (!iDrS.fileEx && fileNameFrom)
    return (-4);              // Abbruch Quell-File existiert nicht, obwohl Parameter fileNameFrom angegeben
  let idSource;
  if (!fileNameFrom)
    idSource = iDrS.folId;    // wenn keine File-Angabe ist die Folder-ID die Quelle
  else
    idSource = iDrS.fileId;   // andernfalls die File-ID
  const resource = {
    shortcutDetails: { targetId: idSource },
    title: shortcutName,
    mimeType: "application/vnd.google-apps.shortcut",
  };
  resource.parents = [{id: iDrD.folId}];
  const shortcut = Drive.Files.insert(resource);
  return (1);                 // Shortcut erzeugen erfolgreich
}

function drCreateFile2(folNameTo="", fileNameTo="", mime="") {
  // leeres Dokument erzeugen, erfordert die Drive API unter Dienste!!!
  // einiges flexibler, da der Datei-Typ über den mime-Parameter vorgegeben wird
  if (folNameTo == "" || fileNameTo == "" || mime == "")
    return (-1);              // fehlende Funktionsparameter
  let iDrD = drInfo(folNameTo, fileNameTo);
  if (!iDrD.folEx || iDrD.fileEx)
    return (-2);              // Abbruch Ziel-Folder existiert nicht oder Ziel-File existiert
  Drive.Files.insert({mimeType: mime, title: fileNameTo, parents: [{id: iDrD.folId}]});
  return (1);                 // Dokument erzeugen erfolgreich
}

function drCreateFile(folNameTo="", fileNameTo="", fileTyp="") {
  // leeres Dokument erzeugen, erfordert die standardmäßige DriveApp,  die Drive API unter Dienste ist nicht erforderlich!
  if (folNameTo == "" || fileNameTo == "" || fileTyp == "")
    return (-1);              // fehlende Funktionsparameter
  let iDrD = drInfo(folNameTo, fileNameTo);
  if (!iDrD.folEx || iDrD.fileEx)
    return (-2);              // Abbruch Ziel-Folder existiert nicht oder Ziel-File existiert
  let doc;
  fileTyp = fileTyp.toLowerCase();
  switch (fileTyp){
    case "sheets":  
      doc = SpreadsheetApp.create("tempFile");
      break;
    case "slides":
      doc = SlidesApp.create("tempFile");
      break;
    case "forms":
      doc = FormApp.create("tempFile");
      break;  
    default:
      doc = DocumentApp.create("tempFile");
  }
  let id = doc.getId();
  let file = DriveApp.getFileById(id);
  file.moveTo(iDrD.folder);
  file.setName(fileNameTo);
  return (1);                 // Dokument erzeugen erfolgreich
}

function drPdfCreate(folNameFrom="", fileNameFrom="", folNameTo="", fileNameTo="") {
  if (folNameFrom == "" || fileNameFrom == "" || folNameTo == "" || fileNameTo == "")
    return (-1);              // fehlende Funktionsparameter
  let iDrS = drInfo(folNameFrom, fileNameFrom);
  if (!iDrS.folEx || !iDrS.fileEx) {
    return (-2);              // Abbruch Quell-Folder oder Quell-File existiert nicht
  }
  let iDrD = drInfo(folNameTo, fileNameTo);
  if (!iDrD.folEx || iDrD.fileEx)
    return (-3);              // Abbruch Ziel-Folder existiert nicht oder Ziel-File existiert
  let sFile = iDrS.file;
  let blob = sFile.getAs('application/pdf');
  iDrD.folder.createFile(blob).setName(fileNameTo);
  return (1);                 // erfolgreich PDF erzeugt
}

function drFileCopy(folNameFrom="", fileNameFrom="", folNameTo="", fileNameTo="") {
  if (folNameFrom == "" || fileNameFrom == "" || folNameTo == "" || fileNameTo == "")
    return (-1);              // fehlende Funktionsparameter
  let iDrS = drInfo(folNameFrom, fileNameFrom);
  if (!iDrS.folEx || !iDrS.fileEx) {
    return (-2);              // Abbruch Quell-Folder oder Quell-File existiert nicht
  }
  let iDrD = drInfo(folNameTo, fileNameTo);
  if (!iDrD.folEx || iDrD.fileEx)
    return (-3);              // Abbruch Ziel-Folder existiert nicht oder Ziel-File existiert
  let file = iDrS.file;
  file.makeCopy(fileNameTo, iDrD.folder);
  return (1);                 // erfolgreich kopiert
}

function drFileTrash(folNameFrom="", fileNameFrom="") {
  if (folNameFrom == "" || fileNameFrom == "")
    return (-1);              // fehlende Funktionsparameter
  let iDrS = drInfo(folNameFrom, fileNameFrom);
  if (!iDrS.folEx || !iDrS.fileEx) {
    return (-2);              // Abbruch Quell-Folder oder Quell-File existiert nicht
  }
  let file = iDrS.file;
  file.setTrashed(true);
  return (1);                 // erfolgreich gelöscht
}

function drCopyFolder(folNameFrom="", folName="", folNameTo="") {
  if (folNameFrom == "" || folName == "" || folNameTo == "")
    return (-1);              // fehlende Funktionsparameter
  let iDrS = drInfo(folNameFrom + '/' + folName);
  if (!iDrS.folEx) {
    return (-2);              // Abbruch Quell-Folder existiert nicht
  }
  let iDrD = drInfo(folNameTo);
  if (!iDrD.folEx)
    return (-3);              // Abbruch Ziel Parent-Folder existiert nicht
  iDrD = drInfo(folNameTo + '/' + folName);
  if (iDrD.folEx)
    return (-4);              // Abbruch Ziel-Folder existiert schon, es wird nicht kopiert
  iDrD = drCreateFolder(folNameTo, folName);
  if (iDrD < 1)
    return (-5);              // Abbruch Ziel-Folder konnte nicht angelegt werden
  iDrD = drInfo(folNameTo + '/' + folName);
  if (iDrD.folEx) {
    drCopyFol(iDrS.folder, iDrD.folder);
    return (1);               // erfolgreich kopiert
  }
  else {
    return (-6);              // something wrong
  }
}

function drFiles() {
  // Info über alle Dateien auf Google Drive mit Ordner, Dateiname, Typ, ID, URL. Rückgabewert als Array, Elemente als JSON-Objekt
  let info = {}
  let nameFileFolder;
  let file;
  let parents;
  let parent;
  let arr= [];
  let files = DriveApp.getFiles();
  while (files.hasNext()) {
    nameFileFolder = "";
    file = files.next();
    info.fileName = file.getName();
    info.fileTyp = file.getMimeType();
    info.fileId = file.getId();
    info.fileUrl = file.getUrl();
    parents = file.getParents();
    while (parents.hasNext()){
      parent = parents.next();
      if (nameFileFolder == "") {
        info.folderId = parent.getId();
        info.folderUrl = parent.getUrl();
      }
      nameFileFolder = parent.getName() + "/" + nameFileFolder;
      parents = parent.getParents();
    }
    nameFileFolder = nameFileFolder.substring(0, nameFileFolder.length-1);

    info.folderName = nameFileFolder;
    arr.push(info);
    info ={};
  }
  return arr;
}

function drRoot() {
  // Info über Rootordner -> Name, ID, URL. Rückgabewert als JSON-Objekt
  let folder = DriveApp.getRootFolder();
  let info = {};
  info.folName = folder.getName();
  info.folId = folder.getId();
  info.folUrl = folder.getUrl();
  info.folder = DriveApp.getFolderById(info.folId);
  return info;
}

function drInfo(folderName="", fileName="") {
  // Überprüft, ob Ordner und/oder File vorhanden und gibt Info dazu als JSON Objekt zurück
  let info = {};
  info.folEx = false;          // Folder Exist
  info.fileEx = false;         // File Exist
  if (folderName == "")        // Abbruch, wenn kein Folder angegeben wird
    return info;
  let i=1;
  let folders;
  let folder = DriveApp.getRootFolder();      // default Root-Folder
  let arr = folderName.split('/');
  if (arr.length > 1) {
    folders = folder.getFoldersByName(arr[i]);
    while (folders.hasNext()) {
      folder = folders.next();
      i++;
      if (i >= arr.length)
        break;
      folders = folder.getFoldersByName(arr[i]);
    }
    if (i < arr.length)
      return info;                            // Zielordner nicht erreicht
  }
  info.folName = folder.getName();
  info.folId = folder.getId();
  info.folUrl = folder.getUrl();
  info.folder = folder;
  info.folEx = true;
  if (fileName != "") {                       // Überprüfung Filename
    let files = folder.getFiles();
    while (files.hasNext()) {
      let file =files.next();
        if (fileName == file.getName()) {
          info.fileName = file.getName();
          info.fileId = file.getId();
          info.fileUrl = file.getUrl();
          info.fileTyp = file.getMimeType();
          info.file = file;
          info.fileEx = true;
          break;
        }
    }
  }
  return info;
}

function drCreateFolder(folNameFrom="", folName="") {
  // legt Folder 'folName' im Folder 'folNameFrom' an, Rückgabe ist ein Fehlercode
  let iDr = drInfo(folNameFrom + '/' + folName);
  if (iDr.folEx) {
    return (-1);             // Abbruch folName existiert bereits
  }
  iDr = drInfo(folNameFrom);
  if (!iDr.folEx) {
    return (-2);             // Abbruch folNameFrom existiert nicht
  }
  iDr.folder.createFolder(folName);
  return (1);
}

function drTrashFolder(folNameFrom="", folName="") {
  // löscht Folder 'folName' im Folder 'folNameFrom' mit allen Dateien und Unterfoldern, Rückgabe ist ein Fehlercode, landet im Papierkorb
  let iDr = drInfo(folNameFrom);
  if (!iDr.folEx) {
    return (-2);             // Abbruch folNameFrom existiert nicht
  }
  iDr = drInfo(folNameFrom + '/' + folName);
  if (!iDr.folEx) {
    return (-1);             // Abbruch folName existiert nicht
  }
  iDr.folder.setTrashed(true);
  return (1);
}

function drCopyFol(fromFolder, toFolder) {
  // copy files
  let files = fromFolder.getFiles()
  while (files.hasNext()) {
    let file = files.next();
    let newFile = file.makeCopy(toFolder)
    newFile.setName(file.getName())
  }

  // copy folders
  let folders = fromFolder.getFolders()
  while (folders.hasNext()) {
    let folder = folders.next()
    let newFolder = toFolder.createFolder(folder.getName())
    drCopyFol(folder, newFolder)
  }
}

// zählt die Anzahl der Dateien in einem Verzeichnis
function drGetFileCount(folNameFrom=""){
  if (folNameFrom == "")
    return (-1);              // fehlende Funktionsparameter
  let iDrS = drInfo(folNameFrom);
  if (!iDrS.folEx)
    return (-2);              // Abbruch Folder existiert nicht
  let files = iDrS.folder.getFiles();
  let count = 0;
  while (files.hasNext()) {
    files.next();
    count++; 
  }
  return count;
}





