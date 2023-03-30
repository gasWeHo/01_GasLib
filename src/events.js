let reProp = PropertiesService.getScriptProperties();

function onLoad() {
  if (reCheckAblage())
    reProp.setProperty('reAblageOK', 'true');
  else
    reProp.setProperty('reAblageOK', 'false');
}


