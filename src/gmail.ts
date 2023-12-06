function getAllEmailsWithQuery(query: string, timeZone: string) {
  const data: GoogleAppsScript.Gmail.GmailMessage[] = [];

  let start = 0;
  const maxThreadsPerBatch = 100;
  let threads;

  do {
    threads = GmailApp.search(query, start, maxThreadsPerBatch);

    let messages = GmailApp.getMessagesForThreads(threads);

    messages.forEach((thread) => {
      thread.forEach((message) => {
        data.push(message);
      });
    });

    start += maxThreadsPerBatch;
  } while (threads.length === maxThreadsPerBatch);

  return data;
}

export { getAllEmailsWithQuery };
