import {
  ACTIONS,
  EnumAction,
  EnumColor,
  EnumSheet,
  EnumTargetType,
  TARGET_TYPES,
  WEEKDAYS,
  LAST_UPDATE_PROPERTY,
} from "./config";

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
  onOpen();
  initInboxSheet();
  initActionsSheet();
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

function getCleanSheet(name: string) {
  const doc = SpreadsheetApp.getActive();

  const sheet = doc.getSheetByName(name);
  if (sheet == null) {
    return doc.insertSheet(name);
  }

  const filter = sheet.getFilter();
  if (filter != null) {
    filter.remove();
  }

  sheet.clear();
  return sheet;
}

function getSheet(name: string) {
  const doc = SpreadsheetApp.getActive();
  const sheet = doc.getSheetByName(name);
  if (sheet == null) {
    throw new Error(`${name} sheet not found`);
  }
  return sheet;
}

function getInboxSheet() {
  return getSheet(EnumSheet.INBOX);
}

function getActionsSheet() {
  return getSheet(EnumSheet.ACTIONS);
}

function initActionsSheet() {
  const sheet = getCleanSheet(EnumSheet.ACTIONS);
  sheet.setFrozenRows(1);

  const data: any[] = [
    ["Target", "Type", "Action"],
    ["email.com", "Domain", "Archive"],
    ["admin@email.com", "Email", "Delete"],
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

function getAllEmailsWithQuery(query: string, timeZone: string) {
  const data: any[] = [];

  let start = 0;
  const maxThreadsPerBatch = 100;
  let threads;

  do {
    threads = GmailApp.search(query, start, maxThreadsPerBatch);

    let messages = GmailApp.getMessagesForThreads(threads);

    messages.forEach((thread) => {
      thread.forEach((message) => {
        const emailDetails = extractEmailDetails(message, timeZone);
        data.push(emailDetails);
      });
    });

    start += maxThreadsPerBatch;
  } while (threads.length === maxThreadsPerBatch);

  return data;
}

function initInboxSheet() {
  const doc = SpreadsheetApp.getActive();
  const timeZone = doc.getSpreadsheetTimeZone();

  const dataHeader = [
    "Thread Id",
    "Mail Id",
    "Email",
    "Email Domain",
    "Date",
    "Subject",
    "Weekday",
    "Hour",
  ];
  const emailsData = getAllEmailsWithQuery("in:inbox", timeZone);
  const data = [dataHeader, ...emailsData];

  const sheet = getCleanSheet(EnumSheet.INBOX);
  sheet.setFrozenRows(1);

  sheet.getRange(1, 1, data.length, dataHeader.length).setValues(data);
  const range = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn());
  range.createFilter();

  setLastUpdate(timeZone);
}

function updateInboxSheet() {
  const doc = SpreadsheetApp.getActive();
  const timeZone = doc.getSpreadsheetTimeZone();
  const lastUpdate = getLastUpdate();

  if (lastUpdate == null) {
    return initInboxSheet();
  }

  const query = `in:inbox after:${lastUpdate}`;
  const existingEmailIds = getExistingEmailIds();
  const allData = getAllEmailsWithQuery(query, timeZone);
  if (allData.length === 0) {
    return;
  }

  const data = allData.filter((email) => !existingEmailIds.includes(email[0]));
  if (data.length === 0) {
    return;
  }

  const sheet = getInboxSheet();
  const numRows = sheet.getLastRow();
  sheet.getRange(numRows + 1, 1, data.length, data[0].length).setValues(data);

  setLastUpdate(timeZone);
}

function executeActions() {
  const actionsSheet = getActionsSheet();
  const inboxSheet = getInboxSheet();
  const actionData = actionsSheet.getDataRange().getValues();
  const inboxData = inboxSheet.getDataRange().getValues();

  if (actionData.length === 0) {
    return;
  }

  const inbox = getInboxValues();
  const threadsForAction = new Map<string, string>();

  for (const row of actionData) {
    const target = row[0];
    const type = row[1];
    const action = row[2];

    if (type === EnumTargetType.DOMAIN) {
      const domainThreadIds = inbox
        .filter((mail) => mail[3] === target)
        .map((mail) => mail[0]);
      domainThreadIds.forEach((emailId) =>
        threadsForAction.set(emailId, action)
      );
    } else if (type === EnumTargetType.EMAIL) {
      const threadIds = inbox
        .filter((mail) => mail[2] === target)
        .map((mail) => mail[0]);
      threadIds.forEach((threadId) => threadsForAction.set(threadId, action));
    }
  }

  const threadArray = Array.from(threadsForAction.entries());

  for (const [threadId, action] of threadArray) {
    const thread = GmailApp.getThreadById(threadId);
    if (action === EnumAction.ARCHIVE) {
      thread.moveToArchive();
    } else if (action === EnumAction.DELETE) {
      thread.moveToTrash();
    } else if (action === EnumAction.SPAM) {
      thread.moveToSpam();
    }
  }

  const rowsToDelete: number[] = [];
  for (let i = inboxData.length - 1; i >= 0; i--) {
    const threadId = inboxData[i][0];

    if (threadsForAction.has(threadId)) {
      rowsToDelete.push(i + 1);
    }
  }

  for (const rowIndex of rowsToDelete) {
    inboxSheet.deleteRow(rowIndex);
  }
}

function getExistingEmailIds() {
  const sheet = getInboxSheet();
  const numRows = sheet.getLastRow();
  const emailIds = sheet.getRange(2, 1, numRows - 1, 1).getValues();
  return emailIds.flat();
}

function getInboxValues() {
  const sheet = getInboxSheet();
  const numRows = sheet.getLastRow();
  const data = sheet.getRange(2, 1, numRows - 1, 7).getValues();
  return data;
}

function extractEmailDetails(
  message: GoogleAppsScript.Gmail.GmailMessage,
  timeZone: string
) {
  const threadId = message.getThread().getId();
  const mailId = message.getId();
  const sender = message.getFrom();
  const match = sender.match(/<([^>]+)>/);
  const email = match ? match[1] : sender.replace(/[\s"]/g, "");
  const domain = email.substring(email.indexOf("@") + 1);
  const date = message.getDate();
  const subject = message.getSubject();
  const weekday = Utilities.formatDate(date, timeZone, "EEE");
  const hour = Utilities.formatDate(date, timeZone, "H");
  return [threadId, mailId, email, domain, date, subject, weekday, hour];
}

function addAllPivotSheets() {
  addEmailPivotSheet();
  addDomainPivotSheet();
  addBusiestHoursPivotSheet();
}

function addEmailPivotSheet() {
  const sheet = getInboxSheet();
  const pivotSheet = getCleanSheet(EnumSheet.EMAIL_PIVOT);
  pivotSheet.setFrozenRows(1);

  const pivotTable = pivotSheet
    .getRange("A1")
    .createPivotTable(sheet.getRange("A1:H"));

  const pivotGroup = pivotTable.addRowGroup(3);
  const pivotValue = pivotTable.addPivotValue(
    3,
    SpreadsheetApp.PivotTableSummarizeFunction.COUNTA
  );
  pivotGroup.sortBy(pivotValue, []);
  pivotGroup.sortDescending();
  const criteria = SpreadsheetApp.newFilterCriteria()
    .whenCellNotEmpty()
    .build();
  pivotTable.addFilter(3, criteria);
  pivotValue.setDisplayName("Count");
}

function addDomainPivotSheet() {
  const sheet = getInboxSheet();
  const pivotSheet = getCleanSheet(EnumSheet.DOMAIN_PIVOT);
  pivotSheet.setFrozenRows(1);

  const pivotTable = pivotSheet
    .getRange("A1")
    .createPivotTable(sheet.getRange("A1:H"));

  const pivotGroup = pivotTable.addRowGroup(4);
  const pivotValue = pivotTable.addPivotValue(
    4,
    SpreadsheetApp.PivotTableSummarizeFunction.COUNTA
  );
  pivotGroup.sortBy(pivotValue, []);
  pivotGroup.sortDescending();
  const criteria = SpreadsheetApp.newFilterCriteria()
    .whenCellNotEmpty()
    .build();
  pivotTable.addFilter(4, criteria);
  pivotValue.setDisplayName("Count");
}

function addBusiestHoursPivotSheet() {
  const pivotSheet = getCleanSheet(EnumSheet.HOURS_PIVOT);
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
      let formula = `=COUNTIFS(Inbox!G2:G, "${WEEKDAYS[dayIndex]}", Inbox!H2:H, ${hour})`;
      pivotSheet.getRange(hour + 2, dayIndex + 2).setFormula(formula);
    }
  }

  const dataRange = pivotSheet.getRange(2, 2, 24, 7);
  const rules = pivotSheet.getConditionalFormatRules();

  const colorScaleRule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpoint(EnumColor.GREEN)
    .setGradientMidpointWithValue(
      "white",
      SpreadsheetApp.InterpolationType.PERCENT,
      "50"
    )
    .setGradientMaxpoint(EnumColor.RED)
    .setRanges([dataRange])
    .build();
  rules.push(colorScaleRule);

  pivotSheet.setConditionalFormatRules(rules);
}

function setLastUpdate(timeZone: string) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const lastUpdate = Utilities.formatDate(new Date(), timeZone, "yyyy/MM/dd");
  scriptProperties.setProperty(LAST_UPDATE_PROPERTY, lastUpdate);
}

function getLastUpdate(): string | null {
  const scriptProperties = PropertiesService.getScriptProperties();
  return scriptProperties.getProperty(LAST_UPDATE_PROPERTY);
}
