// Test 
function calAufruf() {
  let mailTo = emGetEmailAddresses("myContacts");
  let start = new Date('March 16, 2023 14:00:00 UTC');
  let end = new Date('March 16, 2023 15:30:00 UTC');
  let i = calCreateEvent("Besprechungstitel", start, end, "zusätzliche Event-Beschreibung", "Raum R4.55", mailTo, true);
  Logger.log("Returncode = " + i);
}
// fügt im Google Kalender ein Event ein
function calCreateEvent(title, start, end, desc = "", loc = "", guests = "", sendInvites = false) {
  let email = Session.getEffectiveUser().getEmail();      // Mail-Adresse des Users, der das GAS-Script gerade ausführt
  let cal = CalendarApp.getCalendarsByName(email)[0];
  if (cal != null) {
    cal.createEvent(title, start, end, {
      description: desc,
      location: loc,
      guests: guests,
      sendInvites: sendInvites
    });
    return (1);
  }
  else {
    return (-1);
  }
}

