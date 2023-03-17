function old1_calAufruf() {
  let mailTo = emGetEmailAddresses("myContacts");
  let start = new Date('March 16, 2023 14:00:00 UTC');
  let end = new Date('March 16, 2023 15:30:00 UTC');
  let i = calCreateEvent("Besprechungstitel", start, end, "zusätzliche Event-Beschreibung", "Raum R4.55", mailTo, true);
  Logger.log("Returncode = " + i);
}
