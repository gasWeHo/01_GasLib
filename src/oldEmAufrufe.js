function old1_emAufruf() {
    let i = emDownloadUrl("Meine Ablage/prog/gas/01_GasLib/Pedro/Hans", "Pedro_Tabelle", " ", "Werner Hofmann",
                    "werner.hofmann16@gmail.com", "Subject der Mail", "My content of these e-mail:  ", "viewer");
  Logger.log("Returncode = " + i);
}

function old2_emAufruf() {
  //let i = emAnhang("Meine Ablage/prog/gas/01_GasLib/Pedro/Hans", "Pedro_Tabelle", Session.getActiveUser().getEmail(),
  //                  "werner.hofmann16@gmail.com", "Subject der Mail", "My content of these e-mail");
  let i = emAnhang("Meine Ablage/prog/gas/01_GasLib/Pedro/Hans", "Pedro_Tabelle", "Werner Hofmann",
                    "werner.hofmann16@gmail.com", "Subject der Mail", "My content of these e-mail");
  Logger.log("Returncode = " + i);
}

function old3_emAufruf() {
  let mailTo = emGetEmailAddresses("myContacts");
  if(mailTo != "")
    MailApp.sendEmail(mailTo, "heute mein Betreff", "und hier ist der Body der e-mail");
}
