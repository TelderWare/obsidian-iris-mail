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
export type BadgeCountMode = "off" | "unread" | "total";
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

export interface IrisMailSettings {
  accounts: Account[];
  redirectPort: number;
  refreshIntervalMinutes: number;
  pageSize: number;
  showReadEmails: boolean;
  /** What the ribbon badge displays. */
  badgeCount: BadgeCountMode;
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
  /** Custom prompt for event/task extraction (overrides default). */
  itemDetectionPrompt: string;
  /** Per-sender automation rules, keyed by lowercased email address. */
  senderRules: Record<string, SenderRule>;
  // Persisted view state
  viewMode: "messages" | "senders";
  sortNewestFirst: boolean;
  filterUnreadOnly: boolean;
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

// Sender grouping
export interface SenderGroup {
  /** Stable key used to identify this group (may differ from address for via senders). */
  groupKey: string;
  address: string;
  name: string;
  messages: Message[];
  latestMessage: Message;
  unreadCount: number;
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

