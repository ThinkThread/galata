function onOpen() {
  const ui = SpreadsheetApp.getUi();

  const pivotMenu = ui.createMenu("Pivots")
    .addItem("ALL", "addAllPivotSheets")
    .addSeparator()
    .addItem("Email Count", "addEmailPivotSheet")
    .addItem("Domain Count", "addDomainPivotSheet")
    .addItem("Busiest Hours", "addBusiestHoursPivotSheet");

  const advancedMenu = ui.createMenu("Advanced")
    .addItem("Init Actions", "initActionsSheet");

  ui.createMenu("Galata")
    .addItem("Update Inbox", "updateInboxSheet")
    .addSeparator()
    .addSubMenu(pivotMenu)
    .addSubMenu(advancedMenu)
    .addToUi();
}

function onInstall() {
  onOpen();
  updateInboxSheet();
  initActionsSheet();
  addAllPivotSheets();
  createTrigger();
}

function createTrigger() {
  const existingTriggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < existingTriggers.length; i++) {
    if (existingTriggers[i].getHandlerFunction() === "updateInboxSheet") {
      ScriptApp.deleteTrigger(existingTriggers[i]);
    }
  }

  ScriptApp.newTrigger("updateInboxSheet").timeBased().everyHours(1).create();
}

function getCleanSheet(sheetName: string) {
  const doc = SpreadsheetApp.getActive();

  const sheet = doc.getSheetByName(sheetName);
  if (sheet == null) {
    return doc.insertSheet(sheetName);
  }

  const filter = sheet.getFilter();
  if (filter != null) {
    filter.remove();
  }

  sheet.clear();
  return sheet;
}

function getInboxSheetContent() {
  const doc = SpreadsheetApp.getActive();
  const sheet = doc.getSheetByName(Sheet.INBOX);
  if (sheet == null) {
    throw new Error("Email sheet not found");
  }
  return sheet;
}

function initActionsSheet() {
  const sheet = getCleanSheet(Sheet.ACTIONS);
  sheet.setFrozenRows(1);

  const data: any[] = [
    ["Target", "Type", "Action"],
    ["email.com", "Domain", "Archive"],
    ["admin@email.com", "Email", "Delete"]
  ];

  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);

  const typeValidationRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(TARGET_TYPES, true)
    .build();
  sheet.getRange("B2:B").setDataValidation(typeValidationRule);

  const actionValidationRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(ACTIONS, true)
    .build();
  sheet.getRange("C2:C").setDataValidation(actionValidationRule);
}

function updateInboxSheet() {
  const doc = SpreadsheetApp.getActive();
  const sheet = getCleanSheet(Sheet.INBOX);
  sheet.setFrozenRows(1);
  const threads = GmailApp.search("label:inbox");
  const messages = GmailApp.getMessagesForThreads(threads);
  const data: any[] = [
    ["Email", "Email Domain", "Date", "Subject", "Weekday", "Hour"],
  ];
  const timeZone = doc.getSpreadsheetTimeZone();

  messages.forEach((thread) => {
    const message = thread[0];
    const emailDetails = extractEmailDetails(message, timeZone);
    data.push(emailDetails);
  });

  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);

  const range = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn());
  range.createFilter();
}

function extractEmailDetails(
  message: GoogleAppsScript.Gmail.GmailMessage,
  timeZone: string
) {
  const sender = message.getFrom();
  const match = sender.match(/<([^>]+)>/);
  const email = match ? match[1] : sender.replace(/[\s"]/g, "");
  const domain = email.substring(email.indexOf("@") + 1);
  const date = message.getDate();
  const subject = message.getSubject();
  const weekday = Utilities.formatDate(date, timeZone, "EEE");
  const hour = Utilities.formatDate(date, timeZone, "H");
  return [email, domain, date, subject, weekday, hour];
}

function addAllPivotSheets() {
  addEmailPivotSheet();
  addDomainPivotSheet();
  addBusiestHoursPivotSheet();
}

function addEmailPivotSheet() {
  const sheet = getInboxSheetContent();
  const pivotSheet = getCleanSheet(Sheet.EMAIL_PIVOT);
  pivotSheet.setFrozenRows(1);

  const pivotTable = pivotSheet
    .getRange("A1")
    .createPivotTable(sheet.getRange("A1:F"));

  const pivotGroup = pivotTable.addRowGroup(1);
  const pivotValue = pivotTable.addPivotValue(
    1,
    SpreadsheetApp.PivotTableSummarizeFunction.COUNTA
  );
  pivotGroup.sortBy(pivotValue, []);
  pivotGroup.sortDescending();
  const criteria = SpreadsheetApp.newFilterCriteria()
    .whenCellNotEmpty()
    .build();
  pivotTable.addFilter(1, criteria);
  pivotValue.setDisplayName("Count");
}

function addDomainPivotSheet() {
  const sheet = getInboxSheetContent();
  const pivotSheet = getCleanSheet(Sheet.DOMAIN_PIVOT);
  pivotSheet.setFrozenRows(1);

  const pivotTable = pivotSheet
    .getRange("A1")
    .createPivotTable(sheet.getRange("A1:F"));

  const pivotGroup = pivotTable.addRowGroup(2);
  const pivotValue = pivotTable.addPivotValue(
    2,
    SpreadsheetApp.PivotTableSummarizeFunction.COUNTA
  );
  pivotGroup.sortBy(pivotValue, []);
  pivotGroup.sortDescending();
  const criteria = SpreadsheetApp.newFilterCriteria()
    .whenCellNotEmpty()
    .build();
  pivotTable.addFilter(2, criteria);
  pivotValue.setDisplayName("Count");
}

function addBusiestHoursPivotSheet() {
  const pivotSheet = getCleanSheet(Sheet.HOURS_PIVOT);
  pivotSheet.setFrozenRows(1);
  pivotSheet.setFrozenColumns(1);

  const headerRange = pivotSheet.getRange(1, 2, 1, WEEKDAYS.length);
  headerRange.setValues([WEEKDAYS]);
  headerRange.setFontWeight("bold");

  const hoursLabels = new Array(24).fill("").map((_, i) => [i]);
  const hoursRange = pivotSheet.getRange(2, 1, 24, 1);
  hoursRange.setValues(hoursLabels);
  hoursRange.setFontWeight("bold");

  for (let hour = 0; hour < 24; hour++) {
    for (let dayIndex = 0; dayIndex < WEEKDAYS.length; dayIndex++) {
      let formula = `=COUNTIFS(Inbox!E:E, "${WEEKDAYS[dayIndex]}", Inbox!F:F, ${hour})`;
      pivotSheet.getRange(hour + 2, dayIndex + 2).setFormula(formula);
    }
  }

  const dataRange = pivotSheet.getRange(2, 2, 24, 7);
  const rules = pivotSheet.getConditionalFormatRules();

  const colorScaleRule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpoint(Color.GREEN)
    .setGradientMidpointWithValue("white", SpreadsheetApp.InterpolationType.PERCENT, "50")
    .setGradientMaxpoint(Color.RED)
    .setRanges([dataRange])
    .build();
  rules.push(colorScaleRule);

  pivotSheet.setConditionalFormatRules(rules);
}