let prop = PropertiesService.getScriptProperties();

function onLoad() {
  if (checkAblageRechnung())
    prop.setProperty('ablageOK', 'true');
  else
    prop.setProperty('ablageOK', 'false');
}
