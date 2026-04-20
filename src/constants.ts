import type { IrisMailSettings } from "./types";

export const VIEW_TYPE_IRIS_MAIL = "iris-mail-view";
export const ICON_NAME = "mail";

export const GRAPH_SCOPES = ["Mail.ReadWrite", "User.Read"];
export const MSAL_AUTHORITY_DEFAULT = "https://login.microsoftonline.com/common";

/** localStorage key prefix for IMAP app passwords (one entry per account id). */
export const IMAP_PASSWORD_STORAGE_KEY = "iris-mail-imap-password";

export const DEFAULT_SETTINGS: IrisMailSettings = {
  accounts: [],
  redirectPort: 3847,
  refreshIntervalMinutes: 5,
  pageSize: 25,
  showReadEmails: true,
  badgeCount: "unread",
  badgePosition: "bottom-left",
  enableClaudeProcessing: false,
  anthropicApiKey: "",
  claudeModel: "claude-haiku-4-5-20251001",
  claudeSystemPrompt: "",
  tagCategories: "",
  tagIcons: {},
  tagDescriptions: {},
  enableAutoTagging: true,
  tagClassifyPrompt: "",
  tagPromptVersions: {},
  prefetchLimit: 10,
  resolveForwardedSender: false,
  eventNoteFolderPath: "Events",
  taskNoteFolderPath: "Tasks",
  itemDetectionPrompt: "",
  senderRules: {},
  viewMode: "messages",
  sortNewestFirst: true,
  filterUnreadOnly: true,
  debugLogging: false,
};

export const NICKNAME_PROMPT =
  "Convert a raw email sender into a clean, natural name. The sender may be a person or an organization.\n" +
  "Input is two lines: the raw display name, then the email address.\n" +
  "Fix casing, reorder surname-first formats, and strip codes or parenthetical tags.\n" +
  "Use the email address as a hint when the display name is ambiguous or missing.\n" +
  "If the input is already clean, return it unchanged.\n" +
  "Return only the name, nothing else.";

export const NICKNAME_BATCH_PROMPT =
  "Convert each raw email sender into a clean, natural name. Senders may be people or organizations.\n" +
  "Input is a numbered list. Each line is: N. <raw display name> | <email address>\n" +
  "Fix casing, reorder surname-first formats, and strip codes or parenthetical tags.\n" +
  "Use the email address as a hint when the display name is ambiguous or missing.\n" +
  "Output one line per input, in the same order, formatted exactly: N. <clean name>\n" +
  "Output only the numbered list, nothing else.";

export const DEFAULT_CLAUDE_PROMPT =
  "You are an information extraction engine. Input is raw HTML email. Output is clean Markdown.\n\n" +
  "Extract only what the recipient doesn't already know and couldn't trivially infer. " +
  "Strip boilerplate, obvious warnings, footers, disclaimers, and pleasantries. " +
  "Prefer fewer, denser lines. " +
  "Email metadata (sender, recipient, date, subject) is shown separately — do not include it.\n\n" +
  "Output a flat bullet-point list. " +
  "Preserve names, numbers, dates exactly. " +
  "Convert URLs to markdown links with short descriptive text, e.g. [View invoice](https://...). " +
  "If no substantive content, return only: No substantive content.";

export const MESSAGE_LIST_SELECT = [
  "id",
  "subject",
  "bodyPreview",
  "sender",
  "from",
  "receivedDateTime",
  "isRead",
  "isDraft",
  "hasAttachments",
  "importance",
  "flag",
].join(",");

export const WELL_KNOWN_FOLDERS = [
  "Inbox",
  "Drafts",
  "Sent Items",
  "Deleted Items",
  "Junk Email",
  "Archive",
];

/** Curated Lucide icon pool. Used as the fallback seed when Claude is unavailable
 *  and as the candidate list when Claude picks an icon for a new tag. */
export const TAG_ICON_POOL = [
  "bookmark", "flag", "star", "zap", "briefcase", "folder", "heart", "flame",
  "shield", "circle-dot", "hash", "gem", "banknote", "credit-card", "wallet",
  "receipt", "graduation-cap", "book", "book-open", "library", "school",
  "users", "user", "building", "building-2", "home", "hospital", "stethoscope",
  "pill", "plane", "car", "train", "ship", "map", "map-pin", "globe",
  "calendar", "clock", "bell", "megaphone", "newspaper", "mail",
  "message-square", "phone", "video", "camera", "image", "music",
  "shopping-cart", "shopping-bag", "package", "truck", "utensils",
  "coffee", "gift", "wrench", "cog", "laptop", "code", "cpu", "database",
  "server", "cloud", "leaf", "tree-pine", "paw-print", "dumbbell",
  "trophy", "target", "lightbulb", "key", "lock",
];

export const TAG_CLASSIFY_PROMPT =
  "You are an email tag classifier. You are given a single tag (with an optional definition) and an email. " +
  "Decide whether the tag applies to this email.\n\n" +
  "Rules:\n" +
  "- Answer with a single word: yes or no\n" +
  "- Consider the tag name and definition carefully\n" +
  "- Be conservative — answer yes only if you are confident the tag applies";

export const ITEM_DETECTION_PROMPT =
  "You scan emails for calendar events and actionable tasks.\n" +
  "Analyze the ENTIRE email body. An email may contain multiple events and/or tasks, or none at all.\n\n" +
  "Return ONLY valid JSON: an array of objects. Each object has:\n" +
  '- "type": "event" or "task"\n' +
  '- "title": short descriptive title\n' +
  '- "date": "YYYY-MM-DD" or "YYYY-MM-DD/YYYY-MM-DD" for a date range (for events, required if determinable)\n' +
  '- "time": "HH:MM" (for events, if mentioned)\n' +
  '- "location": string (for events, if mentioned)\n' +
  '- "dueDate": "YYYY-MM-DD" or "YYYY-MM-DD/YYYY-MM-DD" for a date range (for tasks, if determinable)\n' +
  '- "description": 1-2 sentence summary\n' +
  '- "sourceText": the exact excerpt from the email body that this item was detected from (copy verbatim, keep it short — just the key phrase or sentence, not entire paragraphs)\n\n' +
  "Use the email date to resolve relative references (tomorrow, next week, etc.).\n" +
  "Use date ranges when the email describes a span of time (e.g. \"week commencing 13 April\" → \"2026-04-13/2026-04-17\", \"between 1 June and 30 June\" → \"2026-06-01/2026-06-30\").\n" +
  "Omit fields that don't apply (no time on tasks, no dueDate on events).\n" +
  "If the email contains no events or tasks, return an empty array: []\n" +
  "Be conservative — only extract clear, actionable items. Do not invent items.\n" +
  "IMPORTANT: Only extract tasks that are actionable by the recipient (the user reading the email). " +
  "Do NOT extract tasks assigned to or owned by other people mentioned in the email.";

export const CACHE_STORAGE_KEY = "iris-mail-msal-cache";

/** Parse comma-separated tag categories string into a trimmed, non-empty array. */
export function parseTagCategories(raw: string): string[] {
  if (!raw) return [];
  return raw.split(",").map((s) => s.trim()).filter(Boolean);
}

/** Current prompt version for a tag — defaults to 1 so legacy entries match. */
export function getTagVersion(versions: Record<string, number> | undefined, tag: string): number {
  return versions?.[tag] ?? 1;
}

/** Increment a tag's version in-place and return the new value. */
export function bumpTagVersion(versions: Record<string, number>, tag: string): number {
  const next = getTagVersion(versions, tag) + 1;
  versions[tag] = next;
  return next;
}
