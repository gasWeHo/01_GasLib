let reProp = PropertiesService.getScriptProperties();

// Trigger, sobald die Spreadsheet Datei geladen wurde
function onLoad() {
  if (reCheckAblage())
    reProp.setProperty('reAblageOK', 'true');
  else
    reProp.setProperty('reAblageOK', 'false');
}

// Trigger, sobald eine Zelle in der Spreadsheet Datei geändert wurde
function onEdit(e) {
  try {
    let row = e.range.getRow();
    let col = e.range.getColumn();
    let value = e.value;
    let sName = SpreadsheetApp.getActiveSheet().getName();
    reStatus(row, col, value, sName);
  }
  catch { ; }
}


