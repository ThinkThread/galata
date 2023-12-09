function setLastUpdate(timeZone: string) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const lastUpdate = Utilities.formatDate(new Date(), timeZone, "yyyy/MM/dd");
  scriptProperties.setProperty(LAST_UPDATE_PROPERTY, lastUpdate);
}

function getLastUpdate(): string | null {
  const scriptProperties = PropertiesService.getScriptProperties();
  return scriptProperties.getProperty(LAST_UPDATE_PROPERTY);
}
