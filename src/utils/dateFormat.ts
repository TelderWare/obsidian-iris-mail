/**
 * Format a YYYY-MM-DD date (or YYYY-MM-DD/YYYY-MM-DD range) with relative labels.
 * e.g. "Today", "Tomorrow", "In 3d", "Yesterday", "2d ago", "13 Mar", "13 Mar – 15 Mar"
 */
export function formatItemDate(raw: string): string {
  const parts = raw.split("/");
  if (parts.length === 2 && parts[0] && parts[1]) {
    return `${formatSingleItemDate(parts[0])} – ${formatSingleItemDate(parts[1])}`;
  }
  return formatSingleItemDate(raw);
}

function formatSingleItemDate(dateStr: string): string {
  const m = dateStr.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return dateStr;

  const target = new Date(parseInt(m[1]), parseInt(m[2]) - 1, parseInt(m[3]));
  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const diff = Math.round((target.getTime() - today.getTime()) / 86400000);

  if (diff === 0) return "Today";
  if (diff === 1) return "Tomorrow";
  if (diff === -1) return "Yesterday";
  if (diff > 1 && diff <= 6) return `In ${diff}d`;
  if (diff < -1 && diff >= -6) return `${-diff}d ago`;
  return target.toLocaleDateString(undefined, { day: "numeric", month: "short" });
}

const MONTH_MAP: Record<string, number> = {
  january: 0, february: 1, march: 2, april: 3, may: 4, june: 5,
  july: 6, august: 7, september: 8, october: 9, november: 10, december: 11,
  jan: 0, feb: 1, mar: 2, apr: 3, jun: 5, jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11,
};

function fmtDate(d: Date): string {
  const day = d.getDate();
  const month = d.toLocaleDateString("en-GB", { month: "long" });
  const year = d.getFullYear();
  return `${day} ${month} ${year}`;
}

/**
 * Expand vague date phrases in email text so the LLM doesn't have to do date math.
 * e.g. "w/c 20 April" → "week of 20 April 2026 (20 April 2026 to 26 April 2026)"
 *      "week commencing 3 March" → "week of 3 March 2026 (3 March 2026 to 9 March 2026)"
 */
export function expandDatePhrases(text: string, emailYear: number): string {
  // Match: "w/c", "week commencing", "week beginning", "week starting", "week of"
  // followed by an optional "the", then a day number and month name
  const pattern = /\b(?:w\/c|week\s+(?:commencing|beginning|starting|of))\s+(?:the\s+)?(\d{1,2})(?:st|nd|rd|th)?\s+(january|february|march|april|may|june|july|august|september|october|november|december|jan|feb|mar|apr|jun|jul|aug|sep|oct|nov|dec)\b/gi;

  return text.replace(pattern, (match, dayStr, monthStr) => {
    const day = parseInt(dayStr);
    const monthIdx = MONTH_MAP[monthStr.toLowerCase()];
    if (monthIdx === undefined || day < 1 || day > 31) return match;

    const start = new Date(emailYear, monthIdx, day);
    const end = new Date(start);
    end.setDate(end.getDate() + 6);

    return `${match} (${fmtDate(start)} to ${fmtDate(end)})`;
  });
}

export function formatRelativeDate(isoString: string): string {
  const date = new Date(isoString);
  const now = new Date();
  const diffMs = now.getTime() - date.getTime();
  const diffMins = Math.floor(diffMs / 60000);

  if (diffMins < 1) return "Just now";
  if (diffMins < 60) return `${diffMins}m ago`;

  // Use calendar days to avoid midnight-crossing bugs
  const startOfToday = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const startOfDate = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  const calendarDays = Math.round(
    (startOfToday.getTime() - startOfDate.getTime()) / 86400000,
  );

  if (calendarDays === 0) {
    return date.toLocaleTimeString(undefined, {
      hour: "numeric",
      minute: "2-digit",
    });
  }
  if (calendarDays === 1) return "Yesterday";
  if (calendarDays < 7) return `${calendarDays}d ago`;
  return date.toLocaleDateString(undefined, {
    month: "short",
    day: "numeric",
  });
}
