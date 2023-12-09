function onOpen() {
  const ui = SpreadsheetApp.getUi();

  const pivotMenu = ui
    .createMenu("Pivots")
    .addItem("ALL", "addAllPivotSheets")
    .addSeparator()
    .addItem("Email Count", "addEmailPivotSheet")
    .addItem("Domain Count", "addDomainPivotSheet")
    .addItem("Busiest Hours", "addBusiestHoursPivotSheet");

  const advancedMenu = ui
    .createMenu("Advanced")
    .addItem("Init Inbox", "initInboxSheet")
    .addItem("Init Actions", "initActionsSheet")
    .addItem("Init Log", "initLogSheet")
    .addItem("Init Triggers", "initTrigger");

  ui.createMenu("Galata")
    .addItem("Update Inbox", "updateInboxSheet")
    .addItem("Execute Actions", "executeActions")
    .addSeparator()
    .addSubMenu(pivotMenu)
    .addSubMenu(advancedMenu)
    .addToUi();
}

function onInstall() {
  initInboxSheet();
  initActionsSheet();
  initLogSheet();
  addAllPivotSheets();
  initTrigger();
}

function initTrigger() {
  const existingTriggers = ScriptApp.getProjectTriggers();
  for (const element of existingTriggers) {
    ScriptApp.deleteTrigger(element);
  }

  ScriptApp.newTrigger("initInboxSheet").timeBased().everyDays(1).create();
  ScriptApp.newTrigger("updateInboxSheet").timeBased().everyHours(1).create();
  ScriptApp.newTrigger("executeActions").timeBased().everyHours(1).create();
}
