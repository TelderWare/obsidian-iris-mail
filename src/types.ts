import type * as MicrosoftGraph from "@microsoft/microsoft-graph-types";

// Re-export Graph types we use frequently
export type MailFolder = MicrosoftGraph.MailFolder;
export type Recipient = MicrosoftGraph.Recipient;
export type ItemBody = MicrosoftGraph.ItemBody;

/**
 * Plugin-level message type. The `id` is composite (`{accountId}:{nativeId}`)
 * once it has flowed through MailDispatcher. The leading-underscore fields
 * are transient — they're attached at fan-out time so the UI can show which
 * account a message belongs to and are NOT persisted.
 */
export type Message = MicrosoftGraph.Message & {
  _accountId?: string;
  _accountLabel?: string;
};

// Plugin settings
export type MailProvider = "outlook" | "imap";
export type AuthMethod = "auth-code" | "device-code";
export type BadgePosition = "top-right" | "top-left" | "bottom-right" | "bottom-left" | "off";

/**
 * One configured mail account. Multiple accounts run simultaneously and the
 * inbox view shows their messages merged together. The `id` is a stable
 * opaque token used to namespace tokens, messages, and cache entries.
 */
export interface Account {
  id: string;
  label: string;
  provider: MailProvider;
  /** Whether this account contributes to the unified inbox. */
  enabled: boolean;
  // Outlook
  clientId?: string;
  authority?: string;
  authMethod?: AuthMethod;
  // IMAP (generic)
  imapHost?: string;
  imapPort?: number;
  imapSecure?: boolean;
  imapEmail?: string;
  /** Optional preset key used to pre-fill host/port in the UI. */
  imapPreset?: string;
}

export type BoxBuiltin = "in" | "read" | "todo" | "junk" | "secretary";

/**
 * A named view over the message list. Built-in boxes have a fixed predicate
 * keyed off per-message state (isRead / isTodo / isJunk / classifier in-flight).
 * User boxes are predicated on `tags`: a message is in the box if it carries
 * any of the listed tags. Built-in boxes may also carry extra `tags` to widen
 * their predicate — e.g. tagging a message `task` can feed it into the built-in
 * to-do box.
 */
export interface Box {
  id: string;
  name: string;
  icon: string;
  color?: string;
  builtin?: BoxBuiltin;
  tags?: string[];
  /** When true, the box is omitted from the box strip. Restored via the "+" menu. */
  hidden?: boolean;
  /**
   * When true, messages matching this box's predicate have their envelopes
   * persisted in the cache so they stay visible even after they age out of
   * the server's sync window. Meaningful only for boxes whose predicate is
   * driven by locally-tracked state (todo, junk, user tag boxes); silently
   * ignored for in / read / secretary.
   */
  saved?: boolean;
}

export interface IrisMailSettings {
  accounts: Account[];
  redirectPort: number;
  refreshIntervalMinutes: number;
  pageSize: number;
  /** Sync window in days — messages older than this are not fetched. 0 = unlimited. */
  initialSyncLookbackDays: number;
  showReadEmails: boolean;
  /** Whether the ribbon badge shows the In-box (unread) count. */
  badgeCount: boolean;
  /** Where the ribbon badge is positioned. */
  badgePosition: BadgePosition;
  enableClaudeProcessing: boolean;
  anthropicApiKey: string;
  claudeModel: string;
  claudeSystemPrompt: string;
  /** Comma-separated user-defined tag categories. */
  tagCategories: string;
  /** Map of tag name → Lucide icon name. */
  tagIcons: Record<string, string>;
  /** Map of tag name → hex color applied to the tag icon and viewer badge. Unset = default theme color. */
  tagColors: Record<string, string>;
  /** Map of tag name → true when the tag badge should not appear on message list rows. */
  tagHiddenInList: Record<string, boolean>;
  /** Map of tag name → list of tag names it contradicts (mutually exclusive). Always stored symmetrically. */
  tagContradictions: Record<string, string[]>;
  /** Map of tag name → list of tag names it precludes (directional: if this tag fires, skip those). */
  tagPrecludes: Record<string, string[]>;
  /** Map of tag name → plain-English definition used for yes/no classification. */
  tagDescriptions: Record<string, string>;
  /** Whether auto-tagging via Claude is enabled. */
  enableAutoTagging: boolean;
  /** Max unread messages to auto-tag per inbox load. 0 = disabled, -1 = all. */
  autoTagLimit: number;
  /** Custom tag classification prompt (overrides default). */
  tagClassifyPrompt: string;
  /** Per-tag prompt version — incremented when the tag's definition or the meta-prompt changes. */
  tagPromptVersions: Record<string, number>;
  /** Max messages to prefetch Claude extraction for in background. 0 = disabled, -1 = all. */
  prefetchLimit: number;
  /** Show the original sender of forwarded emails instead of the forwarder. */
  resolveForwardedSender: boolean;
  /** Folder for auto-created event notes. */
  eventNoteFolderPath: string;
  /** Folder for auto-created task notes. */
  taskNoteFolderPath: string;
  /** Per-sender automation rules, keyed by lowercased email address. */
  senderRules: Record<string, SenderRule>;
  // Persisted view state
  sortNewestFirst: boolean;
  /** Ordered list of boxes shown in the box strip. */
  boxes: Box[];
  /** Currently selected box id. */
  selectedBoxId: string;
  /** Enable debug logging to console. */
  debugLogging: boolean;
}

/** Automation rule applied to every incoming message from a given sender. */
export interface SenderRule {
  /** Move new messages to the provider's trash folder on arrival. */
  autoBin?: boolean;
  /** Apply this tag to new messages on arrival. Empty string = none. */
  autoTag?: string;
}

// Internal UI state
export interface MessageListState {
  messages: Message[];
  nextLink: string | null;
  isLoading: boolean;
  searchQuery: string;
}

export interface FolderState {
  folders: MailFolder[];
  selectedFolderId: string | null;
}

