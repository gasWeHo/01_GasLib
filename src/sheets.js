function shAufruf() {
  let range = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("A1:A");
  let find = 25;
  let i = 1 + shGetRowUsingfindIndex(find, range);
  if (i)
    Logger.log("i = " + i);
  else
    Logger.log(`Der Wert ${find} wurde nicht gefunden!`);
}

// gibt den ersten Index in einem Array zurück, sofern term im range gefunden wird, anderfalls -1 
function shGetRowUsingfindIndex(term, range){
  let data = range.getValues();
  let row = 0;
  row = data.findIndex(users => {return users[0] == term});  
  return row;
}

