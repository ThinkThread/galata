import { ACTIONS, TARGET_TYPES, WEEKDAYS } from "./constants";
import { EnumAction, EnumColor, EnumSheet, EnumTargetType } from "./enums";
import { getAllEmailsWithQuery } from "./gmail";
import { getLastUpdate, setLastUpdate } from "./props";

function getSheet(name: string) {
  const doc = SpreadsheetApp.getActive();
  const sheet = doc.getSheetByName(name);
  if (sheet == null) {
    throw new Error(`${name} sheet not found`);
  }
  return sheet;
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
  const emails = getAllEmailsWithQuery("in:inbox", timeZone);
  const emailsData = emails.map((email) => extractEmailDetails(email, timeZone));
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
  const emails = getAllEmailsWithQuery(query, timeZone);
  if (emails.length === 0) {
    return;
  }

  const newEmails = emails.filter((email) => !existingEmailIds.includes(email.getId()));
  if (newEmails.length === 0) {
    return;
  }

  const emailsData = newEmails.map((email) => extractEmailDetails(email, timeZone));

  const sheet = getSheet(EnumSheet.INBOX);
  const numRows = sheet.getLastRow();
  sheet
    .getRange(numRows + 1, 1, emailsData.length, emailsData[0].length)
    .setValues(emailsData);

  setLastUpdate(timeZone);
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

function initLogSheet() {
  const sheet = getCleanSheet(EnumSheet.LOG);
  sheet.setFrozenRows(1);

  const dataHeader = [
    "Action Date",
    "Action Target",
    "Action Type",
    "Action",
    "Thread Id",
    "Mail Id",
    "Email",
    "Email Domain",
    "Date",
    "Subject",
    "Weekday",
    "Hour",
  ];
  const data: any[] = [dataHeader];

  sheet.getRange(1, 1, data.length, dataHeader.length).setValues(data);
  const range = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn());
  range.createFilter();
}

function executeActions() {
  const actionsSheet = getSheet(EnumSheet.ACTIONS);
  const inboxSheet = getSheet(EnumSheet.INBOX);
  const actionData = actionsSheet.getDataRange().getValues();
  const inboxData = inboxSheet.getDataRange().getValues();

  if (actionData.length === 0) {
    return;
  }

  const inbox = getInboxValues();
  const threadsForAction = new Map<string, string>();
  const rowsToLog: any[] = [];

  for (const row of actionData) {
    const target = row[0];
    const type = row[1];
    const action = row[2];

    if (type === EnumTargetType.DOMAIN) {
      const domainThreads = inbox.filter((mail) => mail[3] === target);
      domainThreads.forEach((thread) => {
        threadsForAction.set(thread[0], action);
        rowsToLog.push([
          new Date(),
          target,
          type,
          action,
          ...thread,
        ]);
      });
    } else if (type === EnumTargetType.EMAIL) {
      const threads = inbox.filter((mail) => mail[2] === target);
      threads.forEach((thread) => {
        threadsForAction.set(thread[0], action);
        rowsToLog.push([
          new Date(),
          target,
          type,
          action,
          ...thread,
        ]);
      });
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

  const logSheet = getSheet(EnumSheet.LOG);
  const numRows = logSheet.getLastRow();
  logSheet
    .getRange(numRows + 1, 1, rowsToLog.length, rowsToLog[0].length)
    .setValues(rowsToLog);
}

function addAllPivotSheets() {
  addEmailPivotSheet();
  addDomainPivotSheet();
  addBusiestHoursPivotSheet();
}

function addEmailPivotSheet() {
  const sheet = getSheet(EnumSheet.INBOX);
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
  const sheet = getSheet(EnumSheet.INBOX);
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

function getExistingEmailIds() {
  const sheet = getSheet(EnumSheet.INBOX);
  const numRows = sheet.getLastRow();
  const emailIds = sheet.getRange(2, 1, numRows - 1, 1).getValues();
  return emailIds.flat();
}

function getInboxValues() {
  const sheet = getSheet(EnumSheet.INBOX);
  const numRows = sheet.getLastRow();
  const data = sheet.getRange(2, 1, numRows - 1, 8).getValues();
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

export {
  initInboxSheet,
  updateInboxSheet,
  initActionsSheet,
  initLogSheet,
  executeActions,
  addAllPivotSheets,
};
