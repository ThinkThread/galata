/** What messages.list returns: enough to identify a message, nothing more. */
interface IMessageRef {
  id: string;
  threadId: string;
}

/** A message hydrated with just the headers the Inbox sheet needs. */
interface IMessageMetadata {
  id: string;
  threadId: string;
  from: string;
  subject: string;
  date: Date;
}

interface IInboxRow {
  threadId: string;
  mailId: string;
  email: string;
  emailDomain: string;
  date: GoogleAppsScript.Base.Date;
  subject: string;
  weekday: string;
  hour: string;
}
