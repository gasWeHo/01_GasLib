// Sicherheitsabfrage
function uiMsgBox(strAbfrage){
  let ui = SpreadsheetApp.getUi();
  let result = ui.alert(
     'Bitte bestätigen',
     strAbfrage,
     ui.ButtonSet.YES_NO);
  // user response.
  if (result == ui.Button.YES) {
    return true;
    //ui.alert('Confirmation received.');
  } else {
    // User clicked "No" or X in the title bar.
    return false;
    //ui.alert('Permission denied.');
  }
}
