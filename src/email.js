function emAufruf() {
  let i = emAnhang("Meine Ablage/prog/gas/01_GasLib/Pedro/Hans", "Pedro_Tabelle", "Werner2 Hofmann",
                    "werner.hofmann24@gmail.com", "Subject der Mail", "My content of these e-mail", "werner.hofmann16@gmail.com", "werner.hofmann@siemens.com");
  Logger.log("Returncode = " + i);
}

function emGetEmailAddresses(contactGroup){
  // gibt alle Mail-Adressen einer Kontaktgruppe als String zurück, erfordert die People API unter Dienste!!!
  try {
    const people = People.ContactGroups.list();         // Liste aller Kontaktgruppen
    let i;
    let resName = "";
    for(i=0; i<people.contactGroups.length; i++){       // die gewünschte Kontaktgruppe idenifizieren
      if(people.contactGroups[i].name === contactGroup  &&  people.contactGroups[i].groupType === "USER_CONTACT_GROUP"){
        resName = people.contactGroups[i].resourceName;
        break;
      }
    }
    if(resName != ""){
      group = People.ContactGroups.get(resName, {maxMembers: 25000});   // group-Objekt der gewünschten Gruppe
      let group_contacts = People.People.getBatchGet({                  // Kontakte der Gruppenmitglieder
      resourceNames: group.memberResourceNames,
      personFields: "emailAddresses"});
      let emailAdr = group_contacts.responses.map(x => {
        let emailObjects = x.person.emailAddresses;
        if (emailObjects != null) {
          return emailObjects.map(eo => eo.value);}                     // Array der E-Mailadressen der Gruppenmitglieder
        });
      return (emailAdr.toString());                                     // Stringrückgabe  "em1,em2,em3,...
    }
    else{
      return ("");
    }
  }
  catch (err) {
    return ("");
  }
}

function emDownloadUrl(folNameFrom="", fileNameFrom="", url=" ", mailFrom="", mailTo="", subject="", content="", access="", mailCc="", mailBcc="") {
  if(!folNameFrom || !fileNameFrom || !url || !mailFrom || !mailTo)
    return (-1);             // fehlende Funktionsparameter
  let iDrS = drInfo(folNameFrom, fileNameFrom);
  if (!iDrS.folEx || !iDrS.fileEx)
    return (-2);              // Abbruch Folder oder Datei existiert nicht
  let link = "https://docs.google.com/spreadsheets/d/" + iDrS.fileId + "/export?format=pdf";
  if (access == "viewer"){
    let file = DriveApp.getFileById(iDrS.fileId);
    //file.addViewer(mailTo);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  }
  MailApp.sendEmail(mailTo, subject, content + link, {
    name: mailFrom,
    cc: mailCc,
    bcc: mailBcc});
  return (1);                 // E-mail erfolgreich versendet
}

function emAnhang(folNameFrom="", fileNameFrom="", mailFrom ="", mailTo="", subject="", content="", mailCc="", mailBcc="") {
  if(!folNameFrom || !fileNameFrom || !mailFrom || !mailTo)
    return (-1);             // fehlende Funktionsparameter
  let iDrS = drInfo(folNameFrom, fileNameFrom);
  if (!iDrS.folEx || !iDrS.fileEx)
    return (-2);              // Abbruch Folder oder Datei existiert nicht
  MailApp.sendEmail(mailTo, subject, content, {
    name: mailFrom,
    attachments: [iDrS.file],
    cc: mailCc,
    bcc: mailBcc});
  return (1);                 // E-mail erfolgreich versendet
}