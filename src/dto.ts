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