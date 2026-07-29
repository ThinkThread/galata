/**
 * Reads the mails matching a Gmail query.
 *
 * Deliberately bypasses GmailApp: its batch reader walks every thread and every
 * message serially, which is what made a 1.000 mail scan take minutes. The REST
 * API instead hands back every message id in two calls, and the per-message
 * metadata is then fetched with UrlFetchApp.fetchAll, which runs the requests
 * in parallel.
 */
function fetchMessages(
  query: string,
  progress?: ProgressReporter
): IMessageMetadata[] {
  const listStartedAt = Date.now();
  const refs = listMessageRefs(query);
  console.info(
    `Listed ${refs.length} messages in ${Date.now() - listStartedAt}ms`
  );

  if (refs.length === 0) {
    return [];
  }

  if (progress != null) {
    // The exact total is known now, so the bar never needs an estimate.
    progress.startPhase(PROGRESS_TITLE_SCAN, refs.length);
  }

  const fetchStartedAt = Date.now();
  const messages = fetchMessagesMetadata(refs, progress);
  console.info(
    `Fetched ${messages.length} messages in ${Date.now() - fetchStartedAt}ms`
  );

  return messages;
}

/**
 * Every message id matching the query. messages.list returns threadId along
 * with each id, so the thread never has to be looked up separately.
 */
function listMessageRefs(query: string): IMessageRef[] {
  const refs: IMessageRef[] = [];
  let pageToken: string | undefined;

  do {
    const optionalArgs: any = {
      q: query,
      maxResults: GMAIL_LIST_PAGE_SIZE,
      fields: "messages(id,threadId),nextPageToken",
    };
    if (pageToken != null) {
      optionalArgs.pageToken = pageToken;
    }

    const response = Gmail.Users!.Messages!.list("me", optionalArgs);
    const messages = response.messages;

    if (messages != null) {
      for (const message of messages) {
        refs.push({ id: message.id!, threadId: message.threadId! });
      }
    }

    pageToken = response.nextPageToken;
  } while (pageToken != null);

  return refs;
}

function fetchMessagesMetadata(
  refs: IMessageRef[],
  progress?: ProgressReporter
): IMessageMetadata[] {
  const authHeader = { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` };
  const messages: IMessageMetadata[] = [];

  for (let i = 0; i < refs.length; i += GMAIL_FETCH_CHUNK_SIZE) {
    const chunk = refs.slice(i, i + GMAIL_FETCH_CHUNK_SIZE);

    const requests = chunk.map((ref) => ({
      url: `${GMAIL_API_BASE}/messages/${ref.id}?${GMAIL_METADATA_QUERY}`,
      headers: authHeader,
      muteHttpExceptions: true,
    }));

    // The whole chunk goes out at once instead of one request at a time.
    const responses = UrlFetchApp.fetchAll(requests);

    for (let j = 0; j < responses.length; j++) {
      const message = parseMessageMetadata(responses[j], chunk[j]);
      if (message != null) {
        messages.push(message);
      }
    }

    if (progress != null) {
      progress.advance(chunk.length);
    }
  }

  return messages;
}

function parseMessageMetadata(
  response: GoogleAppsScript.URL_Fetch.HTTPResponse,
  ref: IMessageRef
): IMessageMetadata | null {
  const status = response.getResponseCode();
  if (status !== 200) {
    // A single unreadable message must not sink the whole run.
    console.warn(`Message ${ref.id} skipped with status ${status}`);
    return null;
  }

  const body = JSON.parse(response.getContentText());
  const headers = body.payload != null ? body.payload.headers : null;

  return {
    id: body.id,
    threadId: body.threadId,
    from: findHeader(headers, "From"),
    subject: findHeader(headers, "Subject"),
    // internalDate is epoch milliseconds as a string.
    date: new Date(Number(body.internalDate)),
  };
}

function findHeader(headers: any[] | null, name: string): string {
  if (headers == null) {
    return "";
  }
  for (const header of headers) {
    if (header.name === name) {
      return header.value == null ? "" : header.value;
    }
  }
  return "";
}
