const WEEKDAYS = [
  EnumWeekday.MONDAY,
  EnumWeekday.TUESDAY,
  EnumWeekday.WEDNESDAY,
  EnumWeekday.THURSDAY,
  EnumWeekday.FRIDAY,
  EnumWeekday.SATURDAY,
  EnumWeekday.SUNDAY,
];

const ACTIONS = [EnumAction.ARCHIVE, EnumAction.DELETE, EnumAction.SPAM];

const TARGET_TYPES = [EnumTargetType.DOMAIN, EnumTargetType.EMAIL];

const LAST_UPDATE_PROPERTY = "LAST_UPDATE";

const INBOX_HEADER = [
  "Thread Id",
  "Mail Id",
  "Email",
  "Email Domain",
  "Date",
  "Subject",
  "Weekday",
  "Hour",
];

const INBOX_COLUMN_COUNT = INBOX_HEADER.length;

const LOG_HEADER = [
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
];

// GmailApp's batch move operations cap out at 100 threads per call.
const GMAIL_BATCH_SIZE = 100;

const GMAIL_API_BASE = "https://gmail.googleapis.com/gmail/v1/users/me";

// Gmail API's own maximum for messages.list.
const GMAIL_LIST_PAGE_SIZE = 500;

// How many message reads UrlFetchApp.fetchAll runs in parallel. Higher chunks
// finish sooner but make a single failure cost more work to retry.
const GMAIL_FETCH_CHUNK_SIZE = 100;

// Ask for the three headers the sheet needs and nothing else: the response
// drops from tens of kilobytes per message to a few hundred bytes.
const GMAIL_METADATA_QUERY =
  "format=metadata" +
  "&metadataHeaders=From" +
  "&metadataHeaders=Subject" +
  "&fields=id,threadId,internalDate,payload/headers";

const SENDER_EMAIL_PATTERN = /<([^>]+)>/;

// "Mon|9" - a single formatDate call instead of one per field.
const WEEKDAY_HOUR_FORMAT = "EEE'|'H";

// Every toast is a round trip to the spreadsheet, so renders are throttled by
// elapsed time: a fast phase cannot spam the UI and a slow phase still
// refreshes on a predictable cadence.
const PROGRESS_INTERVAL_MS = 2000;

// Comfortably longer than the interval, so the toast never blinks out between
// two updates, but short enough to clear itself if the script dies mid-run.
const PROGRESS_TOAST_SECONDS = 15;
const PROGRESS_FINAL_TOAST_SECONDS = 8;

const PROGRESS_BAR_WIDTH = 16;
const PROGRESS_BAR_FILLED = "█";
const PROGRESS_BAR_EMPTY = "░";

const PROGRESS_TITLE_LIST = "Galata · Listing inbox";
const PROGRESS_TITLE_SCAN = "Galata · Fetching mails";
const PROGRESS_TITLE_PROCESS = "Galata · Processing mails";
const PROGRESS_TITLE_ACTIONS = "Galata · Applying actions";
const PROGRESS_TITLE_DONE = "Galata · Done";
