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

  const query = "in:inbox";
  const startedAt = Date.now();
  const progress = new ProgressReporter(PROGRESS_TITLE_LIST, 0);

  const emails = fetchMessages(query, progress);
  const emailsData = extractAllEmailDetails(emails, timeZone, progress);
  const data = [INBOX_HEADER, ...emailsData];

  const sheet = getCleanSheet(EnumSheet.INBOX);
  sheet.setFrozenRows(1);

  sheet.getRange(1, 1, data.length, INBOX_COLUMN_COUNT).setValues(data);
  sheet.getRange(1, 1, data.length, INBOX_COLUMN_COUNT).createFilter();

  setLastUpdate(timeZone);

  progress.finish(
    `${formatCount(emailsData.length)} mails · ${formatDuration(
      Date.now() - startedAt
    )}`
  );
}

function updateInboxSheet() {
  const doc = SpreadsheetApp.getActive();
  const timeZone = doc.getSpreadsheetTimeZone();
  const lastUpdate = getLastUpdate();

  if (lastUpdate == null) {
    return initInboxSheet();
  }

  const query = `in:inbox after:${lastUpdate}`;
  const startedAt = Date.now();
  const progress = new ProgressReporter(PROGRESS_TITLE_LIST, 0);

  const emails = fetchMessages(query, progress);
  if (emails.length === 0) {
    progress.finish("No new mail");
    return;
  }

  // Only read the sheet once we know there is something to deduplicate against.
  const existingMailIds = getExistingMailIds();
  const newEmails = emails.filter((email) => !existingMailIds.has(email.id));
  if (newEmails.length === 0) {
    progress.finish("No new mail");
    return;
  }

  const emailsData = extractAllEmailDetails(newEmails, timeZone, progress);

  const sheet = getSheet(EnumSheet.INBOX);
  const numRows = sheet.getLastRow();
  sheet
    .getRange(numRows + 1, 1, emailsData.length, INBOX_COLUMN_COUNT)
    .setValues(emailsData);

  setLastUpdate(timeZone);

  progress.finish(
    `${formatCount(emailsData.length)} new mails · ${formatDuration(
      Date.now() - startedAt
    )}`
  );
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

  sheet.getRange(1, 1, 1, LOG_HEADER.length).setValues([LOG_HEADER]);
  sheet.getRange(1, 1, 1, LOG_HEADER.length).createFilter();
}

function executeActions() {
  const actionsSheet = getSheet(EnumSheet.ACTIONS);
  const inboxSheet = getSheet(EnumSheet.INBOX);
  const startedAt = Date.now();

  const actionData = actionsSheet.getDataRange().getValues();
  if (actionData.length <= 1) {
    return;
  }

  const inboxData = readInboxData(inboxSheet);
  if (inboxData.length === 0) {
    return;
  }

  // Index the inbox once so every action row is a map lookup instead of a
  // full scan: O(actions + mails) rather than O(actions * mails).
  const rowsByEmail = new Map<string, number[]>();
  const rowsByDomain = new Map<string, number[]>();
  for (let i = 0; i < inboxData.length; i++) {
    addToIndex(rowsByEmail, inboxData[i][EnumInboxColumn.EMAIL], i);
    addToIndex(rowsByDomain, inboxData[i][EnumInboxColumn.EMAIL_DOMAIN], i);
  }

  const threadActions = new Map<string, string>();
  const rowsToLog: any[][] = [];
  const actionDate = new Date();

  for (let i = 1; i < actionData.length; i++) {
    const target = actionData[i][0];
    const type = actionData[i][1];
    const action = actionData[i][2];

    if (ACTIONS.indexOf(action) === -1) {
      continue;
    }

    let matchedRows: number[] | undefined;
    if (type === EnumTargetType.DOMAIN) {
      matchedRows = rowsByDomain.get(target);
    } else if (type === EnumTargetType.EMAIL) {
      matchedRows = rowsByEmail.get(target);
    }
    if (matchedRows == null) {
      continue;
    }

    for (const rowIndex of matchedRows) {
      const row = inboxData[rowIndex];
      threadActions.set(row[EnumInboxColumn.THREAD_ID], action);
      rowsToLog.push([
        actionDate,
        target,
        type,
        action,
        row[EnumInboxColumn.THREAD_ID],
        row[EnumInboxColumn.MAIL_ID],
        row[EnumInboxColumn.EMAIL],
        row[EnumInboxColumn.EMAIL_DOMAIN],
        row[EnumInboxColumn.DATE],
        row[EnumInboxColumn.SUBJECT],
      ]);
    }
  }

  if (threadActions.size === 0) {
    return;
  }

  const progress = new ProgressReporter(
    PROGRESS_TITLE_ACTIONS,
    threadActions.size
  );
  applyThreadActions(threadActions, progress);

  const remainingRows = inboxData.filter(
    (row) => !threadActions.has(row[EnumInboxColumn.THREAD_ID])
  );
  writeInboxRows(inboxSheet, remainingRows, inboxData.length);

  if (rowsToLog.length > 0) {
    const logSheet = getSheet(EnumSheet.LOG);
    const numRows = logSheet.getLastRow();
    logSheet
      .getRange(numRows + 1, 1, rowsToLog.length, LOG_HEADER.length)
      .setValues(rowsToLog);
  }

  progress.finish(
    `${formatCount(threadActions.size)} threads · ${formatDuration(
      Date.now() - startedAt
    )}`
  );
}

/**
 * Applies every pending action with the batch Gmail endpoints: one call per
 * (action, 100 threads) instead of one call per thread.
 */
function applyThreadActions(
  threadActions: Map<string, string>,
  progress?: ProgressReporter
) {
  const threadsByAction = new Map<
    string,
    GoogleAppsScript.Gmail.GmailThread[]
  >();

  // Gmail has no batch thread lookup, so this loop is the slow half of the
  // work and the only place worth reporting from.
  const entries = Array.from(threadActions.entries());
  for (const entry of entries) {
    const thread = GmailApp.getThreadById(entry[0]);
    if (progress != null) {
      progress.advance(1);
    }
    if (thread == null) {
      continue;
    }
    const bucket = threadsByAction.get(entry[1]);
    if (bucket == null) {
      threadsByAction.set(entry[1], [thread]);
    } else {
      bucket.push(thread);
    }
  }

  const groups = Array.from(threadsByAction.entries());
  for (const group of groups) {
    const action = group[0];
    const threads = group[1];

    for (let i = 0; i < threads.length; i += GMAIL_BATCH_SIZE) {
      const batch = threads.slice(i, i + GMAIL_BATCH_SIZE);
      if (action === EnumAction.ARCHIVE) {
        GmailApp.moveThreadsToArchive(batch);
      } else if (action === EnumAction.DELETE) {
        GmailApp.moveThreadsToTrash(batch);
      } else if (action === EnumAction.SPAM) {
        GmailApp.moveThreadsToSpam(batch);
      }
    }
  }
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

  // Build the whole 24x7 grid in memory and write it in a single call.
  const formulas: string[][] = [];
  for (let hour = 0; hour < 24; hour++) {
    const row: string[] = [];
    for (const weekday of WEEKDAYS) {
      row.push(
        `=COUNTIFS(${EnumSheet.INBOX}!G2:G, "${weekday}", ${EnumSheet.INBOX}!H2:H, ${hour})`
      );
    }
    formulas.push(row);
  }

  const dataRange = pivotSheet.getRange(2, 2, 24, WEEKDAYS.length);
  dataRange.setFormulas(formulas);

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

function readInboxData(sheet?: GoogleAppsScript.Spreadsheet.Sheet): any[][] {
  const inboxSheet = sheet != null ? sheet : getSheet(EnumSheet.INBOX);
  const lastRow = inboxSheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }
  return inboxSheet
    .getRange(2, 1, lastRow - 1, INBOX_COLUMN_COUNT)
    .getValues();
}

/**
 * Overwrites the inbox body with `rows` and blanks whatever is left over.
 * Two range calls regardless of how many rows disappeared, where deleting the
 * rows one by one costs one call each.
 */
function writeInboxRows(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  rows: any[][],
  previousRowCount: number
) {
  if (rows.length === previousRowCount) {
    return;
  }

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, INBOX_COLUMN_COUNT).setValues(rows);
  }

  const staleRowCount = previousRowCount - rows.length;
  if (staleRowCount > 0) {
    sheet
      .getRange(2 + rows.length, 1, staleRowCount, INBOX_COLUMN_COUNT)
      .clearContent();
  }
}

function addToIndex(index: Map<string, number[]>, key: any, rowIndex: number) {
  if (key == null || key === "") {
    return;
  }
  const bucket = index.get(key);
  if (bucket == null) {
    index.set(key, [rowIndex]);
  } else {
    bucket.push(rowIndex);
  }
}

function getExistingMailIds(): Set<string> {
  const sheet = getSheet(EnumSheet.INBOX);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return new Set<string>();
  }

  const mailIds = sheet
    .getRange(2, EnumInboxColumn.MAIL_ID + 1, lastRow - 1, 1)
    .getValues();

  const ids = new Set<string>();
  for (const row of mailIds) {
    ids.add(String(row[0]));
  }
  return ids;
}

function getInboxValues(): IInboxRow[] {
  return readInboxData().map((row) => ({
    threadId: row[EnumInboxColumn.THREAD_ID],
    mailId: row[EnumInboxColumn.MAIL_ID],
    email: row[EnumInboxColumn.EMAIL],
    emailDomain: row[EnumInboxColumn.EMAIL_DOMAIN],
    date: row[EnumInboxColumn.DATE],
    subject: row[EnumInboxColumn.SUBJECT],
    weekday: row[EnumInboxColumn.WEEKDAY],
    hour: row[EnumInboxColumn.HOUR],
  }));
}

/**
 * Second phase of a run: the fetched messages are turned into sheet rows.
 * Reported separately from the scan because the message count is known exactly
 * here, so the bar stops depending on the Gmail estimate.
 */
function extractAllEmailDetails(
  emails: IMessageMetadata[],
  timeZone: string,
  progress?: ProgressReporter
): any[][] {
  if (progress != null) {
    progress.startPhase(PROGRESS_TITLE_PROCESS, emails.length);
  }

  const rows: any[][] = [];
  for (const email of emails) {
    rows.push(extractEmailDetails(email, timeZone));
    if (progress != null) {
      progress.advance(1);
    }
  }
  return rows;
}

function extractEmailDetails(message: IMessageMetadata, timeZone: string) {
  const match = message.from.match(SENDER_EMAIL_PATTERN);
  const email = match ? match[1] : message.from.replace(/[\s"]/g, "");
  const domain = email.substring(email.indexOf("@") + 1);

  const formatted = Utilities.formatDate(
    message.date,
    timeZone,
    WEEKDAY_HOUR_FORMAT
  );
  const separatorIndex = formatted.indexOf("|");
  const weekday = formatted.substring(0, separatorIndex);
  const hour = formatted.substring(separatorIndex + 1);

  return [
    message.threadId,
    message.id,
    email,
    domain,
    message.date,
    message.subject,
    weekday,
    hour,
  ];
}
