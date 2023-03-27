function old1_shAufruf() {
  let range = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("A1:A");
  let find = 25;
  let i = 1 + shGetRowUsingfindIndex(find, range);
  if (i)
    Logger.log("i = " + i);
  else
    Logger.log(`Der Wert ${find} wurde nicht gefunden!`);
}
