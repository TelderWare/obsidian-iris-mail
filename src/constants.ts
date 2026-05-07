import type { Box, IrisMailSettings } from "./types";

/** Seeded box set used for new installs and when migrating old settings. */
export const DEFAULT_BOXES: Box[] = [
  { id: "in", name: "In", icon: "inbox", builtin: "in" },
  { id: "read", name: "Read", icon: "mail-open", builtin: "read" },
  { id: "todo", name: "To-do", icon: "check-square", builtin: "todo", saved: true },
  { id: "junk", name: "Junk", icon: "shield-x", builtin: "junk" },
  { id: "secretary", name: "Secretary", icon: "brain-circuit", builtin: "secretary" },
];

/** Return a fresh copy of the default boxes (arrays are per-install mutable). */
export function freshDefaultBoxes(): Box[] {
  return DEFAULT_BOXES.map((b) => ({ ...b }));
}

export const VIEW_TYPE_IRIS_MAIL = "iris-mail-view";
export const ICON_NAME = "mail";

/** Class applied to the view container when the user pins single-pane layout via `f`. */
export const COMPACT_MODE_CLASS = "iris-compact";

export const GRAPH_SCOPES = ["Mail.ReadWrite", "User.Read"];
export const MSAL_AUTHORITY_DEFAULT = "https://login.microsoftonline.com/common";

/** localStorage key prefix for IMAP app passwords (one entry per account id). */
export const IMAP_PASSWORD_STORAGE_KEY = "iris-mail-imap-password";

export const DEFAULT_SETTINGS: IrisMailSettings = {
  accounts: [],
  redirectPort: 3847,
  refreshIntervalMinutes: 5,
  pageSize: 25,
  initialSyncLookbackDays: 30,
  showReadEmails: true,
  badgeCount: true,
  badgePosition: "bottom-left",
  enableClaudeProcessing: false,
  anthropicApiKey: "",
  claudeModel: "claude-haiku-4-5-20251001",
  claudeSystemPrompt: "",
  tagCategories: "",
  tagIcons: {},
  tagColors: {},
  tagHiddenInList: {},
  tagContradictions: {},
  tagPrecludes: {},
  tagDescriptions: {},
  enableAutoTagging: true,
  autoTagLimit: 10,
  tagClassifyPrompt: "",
  tagPromptVersions: {},
  prefetchLimit: 10,
  resolveForwardedSender: false,
  eventNoteFolderPath: "Events",
  taskNoteFolderPath: "Tasks",
  senderRules: {},
  sortNewestFirst: true,
  boxes: freshDefaultBoxes(),
  selectedBoxId: "in",
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
  "You are an email tag classifier. You are given a tag definition and an email. " +
  "Decide whether the email matches the definition.\n\n" +
  "Rules:\n" +
  "- Answer with a single word: yes or no\n" +
  "- If the definition is empty or ambiguous, answer no\n" +
  "- Be conservative — answer yes only if you are confident the definition applies";

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

/**
 * Overwrite `tag`'s contradictions with `nextList` and keep the store symmetric.
 * Adds `tag` to every new partner's list; removes `tag` from every dropped partner.
 * Mutates `map` in place.
 */
export function setTagContradictions(
  map: Record<string, string[]>,
  tag: string,
  nextList: string[],
): void {
  const prev = new Set(map[tag] || []);
  const next = new Set(nextList);
  if (next.size === 0) {
    delete map[tag];
  } else {
    map[tag] = Array.from(next);
  }
  // Partners newly added: add `tag` to their lists.
  for (const partner of next) {
    if (prev.has(partner)) continue;
    const partnerList = new Set(map[partner] || []);
    partnerList.add(tag);
    map[partner] = Array.from(partnerList);
  }
  // Partners dropped: remove `tag` from their lists.
  for (const partner of prev) {
    if (next.has(partner)) continue;
    const partnerList = (map[partner] || []).filter((t) => t !== tag);
    if (partnerList.length === 0) {
      delete map[partner];
    } else {
      map[partner] = partnerList;
    }
  }
}

/** Remove a tag entirely from the contradictions map (both keys and referenced lists). */
export function removeTagFromContradictions(
  map: Record<string, string[]>,
  tag: string,
): void {
  setTagContradictions(map, tag, []);
}

/** Overwrite `tag`'s own precludes list. Directional — does not touch any other tag's entry. */
export function setTagPrecludesList(
  map: Record<string, string[]>,
  tag: string,
  nextList: string[],
): void {
  const deduped = Array.from(new Set(nextList));
  if (deduped.length === 0) {
    delete map[tag];
  } else {
    map[tag] = deduped;
  }
}

/**
 * Return the set of tag names whose precludes list contains `tag` — i.e., tags that
 * preclude this one. Derived view; not stored.
 */
export function getPrecludedBy(
  map: Record<string, string[]>,
  tag: string,
): string[] {
  const result: string[] = [];
  for (const [other, list] of Object.entries(map)) {
    if (other === tag) continue;
    if ((list || []).includes(tag)) result.push(other);
  }
  return result;
}

/**
 * Update which tags preclude `tag` by editing each other tag's precludes list.
 * Adds `tag` to newly-selected tags' lists; removes from unselected ones.
 * Mutates `map` in place.
 */
export function setPrecludedByFor(
  map: Record<string, string[]>,
  tag: string,
  nextList: string[],
): void {
  const prev = new Set(getPrecludedBy(map, tag));
  const next = new Set(nextList);
  for (const other of next) {
    if (prev.has(other)) continue;
    const list = new Set(map[other] || []);
    list.add(tag);
    map[other] = Array.from(list);
  }
  for (const other of prev) {
    if (next.has(other)) continue;
    const list = (map[other] || []).filter((t) => t !== tag);
    if (list.length === 0) {
      delete map[other];
    } else {
      map[other] = list;
    }
  }
}

/** Remove a tag entirely from a precludes map (both its own entry and references to it). */
export function removeTagFromPrecludes(
  map: Record<string, string[]>,
  tag: string,
): void {
  delete map[tag];
  for (const other of Object.keys(map)) {
    const list = (map[other] || []).filter((t) => t !== tag);
    if (list.length === 0) {
      delete map[other];
    } else {
      map[other] = list;
    }
  }
}
