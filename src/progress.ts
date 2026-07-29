/**
 * Toast based progress indicator.
 *
 * A single reporter is carried through a whole run and re-pointed at each
 * phase, so the user sees one moving bar instead of a queue of notifications.
 * Renders are throttled by elapsed time rather than by item count: `advance()`
 * itself only costs a `Date.now()`, and the actual toast - which is a round
 * trip to the spreadsheet - happens at most once per PROGRESS_INTERVAL_MS.
 */
class ProgressReporter {
  private title: string;
  private total: number;
  private done: number;
  private startedAt: number;
  private lastRenderAt: number;

  constructor(title: string, total: number) {
    this.title = title;
    this.total = total > 0 ? total : 0;
    this.done = 0;
    this.startedAt = Date.now();
    this.lastRenderAt = 0;
    this.render(true);
  }

  /** Re-points the reporter at the next phase and restarts its clock. */
  startPhase(title: string, total: number) {
    this.title = title;
    this.total = total > 0 ? total : 0;
    this.done = 0;
    this.startedAt = Date.now();
    this.lastRenderAt = 0;
    this.render(true);
  }

  advance(count: number) {
    this.done += count;
    if (this.total > 0 && this.done > this.total) {
      // The Gmail estimate came in low. Grow the denominator instead of
      // showing a bar past 100%.
      this.total = this.done;
    }
    this.render(false);
  }

  finish(message: string) {
    showToast(message, PROGRESS_TITLE_DONE, PROGRESS_FINAL_TOAST_SECONDS);
  }

  private render(force: boolean) {
    const now = Date.now();
    if (!force && now - this.lastRenderAt < PROGRESS_INTERVAL_MS) {
      return;
    }
    this.lastRenderAt = now;
    showToast(this.buildMessage(now), this.title, PROGRESS_TOAST_SECONDS);
  }

  private buildMessage(now: number): string {
    const elapsed = now - this.startedAt;

    // No usable estimate (Advanced Gmail Service off, or an empty result):
    // report honestly with a live counter instead of a fake percentage.
    if (this.total === 0) {
      return `${formatCount(this.done)} · ${formatDuration(elapsed)}`;
    }

    const ratio = this.done / this.total;
    const percent = Math.floor(ratio * 100);
    const remaining =
      this.done > 0 ? (elapsed / this.done) * (this.total - this.done) : 0;

    return (
      `${renderProgressBar(ratio)}  ${percent}%\n` +
      `${formatCount(this.done)} / ${formatCount(this.total)}\n` +
      `Elapsed ${formatDuration(elapsed)} · Left ~${formatDuration(remaining)}`
    );
  }
}

function showToast(message: string, title: string, seconds: number) {
  try {
    SpreadsheetApp.getActive().toast(message, title, seconds);
  } catch (error) {
    // Running from a time-driven trigger: there is no spreadsheet UI to talk
    // to. Progress is cosmetic, so never let it break the actual work.
    console.info(`Progress toast skipped: ${error}`);
  }
}

function renderProgressBar(ratio: number): string {
  const clamped = Math.max(0, Math.min(1, ratio));
  const filled = Math.round(clamped * PROGRESS_BAR_WIDTH);
  return (
    PROGRESS_BAR_FILLED.repeat(filled) +
    PROGRESS_BAR_EMPTY.repeat(PROGRESS_BAR_WIDTH - filled)
  );
}

function formatDuration(milliseconds: number): string {
  const totalSeconds = Math.max(0, Math.round(milliseconds / 1000));
  const minutes = Math.floor(totalSeconds / 60);
  const seconds = totalSeconds % 60;
  return `${minutes}:${seconds < 10 ? "0" : ""}${seconds}`;
}

function formatCount(value: number): string {
  return String(value).replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}
