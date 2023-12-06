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

export { getAllEmailsWithQuery };
