import type { IrisMailSettings } from "./types";

export const VIEW_TYPE_IRIS_MAIL = "iris-mail-view";
export const ICON_NAME = "mail";

export const GRAPH_SCOPES = ["Mail.ReadWrite", "User.Read"];
export const MSAL_AUTHORITY_DEFAULT = "https://login.microsoftonline.com/common";

export const DEFAULT_SETTINGS: IrisMailSettings = {
  clientId: "",
  authority: MSAL_AUTHORITY_DEFAULT,
  authMethod: "auth-code",
  redirectPort: 3847,
  refreshIntervalMinutes: 5,
  saveFolderPath: "Emails",
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
  enableAutoTagging: true,
  tagClassifyPrompt: "",
  tagPromptVersion: 1,
  importanceClassifyPrompt: "",
  importancePromptVersion: 1,
  prefetchLimit: 10,
  resolveForwardedSender: false,
  enableAutoItemDetection: true,
  eventNoteFolderPath: "Events",
  taskNoteFolderPath: "Tasks",
  itemDetectionPrompt: "",
  viewMode: "senders",
  sortNewestFirst: true,
  filterUnreadOnly: true,
  filterHideNoise: true,
  filterImportantOnly: false,
  debugLogging: false,
};

export const IMPORTANCE_CLASSIFY_PROMPT =
  "Classify this email as: important, routine, or noise.\n" +
  "important — requires attention, action, or contains personally relevant information\n" +
  "routine — informational, no action needed\n" +
  "noise — marketing, automated, newsletters, or zero-value content\n" +
  "Return only the single word.";

export const NICKNAME_PROMPT =
  "Convert the raw email display name into a clean, natural person name.\n" +
  "Fix casing, reorder surname-first formats, and strip codes or parenthetical tags.\n" +
  "If the input is already a clean name, return it unchanged.\n" +
  "Return only the name, nothing else.";

export const DEFAULT_CLAUDE_PROMPT =
  "You are an information extraction engine. Input is raw HTML email. Output is clean Markdown.\n\n" +
  "Extract only what the recipient doesn't already know and couldn't trivially infer. " +
  "Strip boilerplate, obvious warnings, footers, disclaimers, and pleasantries. " +
  "Prefer fewer, denser lines. " +
  "Email metadata (sender, recipient, date, subject) is shown separately — do not include it.\n\n" +
  "Output a flat bullet-point list. " +
  "Preserve names, numbers, dates exactly. " +
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
  "conversationId",
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

export const TAG_ICON_CYCLE = [
  "bookmark",
  "flag",
  "star",
  "zap",
  "briefcase",
  "folder",
  "heart",
  "flame",
  "shield",
  "circle-dot",
  "hash",
  "gem",
];

export const TAG_CLASSIFY_PROMPT =
  "You are an email tag classifier. Given an email and a list of tag categories, " +
  "return the tags that apply to this email.\n\n" +
  "Rules:\n" +
  "- Return ONLY a JSON array of strings, e.g. [\"Finance\", \"Projects\"]\n" +
  "- Only use tags from the provided list\n" +
  "- Return an empty array [] if no tags clearly apply\n" +
  "- Assign 1-3 tags maximum\n" +
  "- Be conservative — only assign tags you are confident about";

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
