import type * as MicrosoftGraph from "@microsoft/microsoft-graph-types";

// Re-export Graph types we use frequently
export type MailFolder = MicrosoftGraph.MailFolder;
export type Message = MicrosoftGraph.Message;
export type Recipient = MicrosoftGraph.Recipient;
export type ItemBody = MicrosoftGraph.ItemBody;

// Graph API list response with pagination
export interface GraphPagedResponse<T> {
  "@odata.context"?: string;
  "@odata.nextLink"?: string;
  value: T[];
}

// Plugin settings
export type AuthMethod = "auth-code" | "device-code";
export type BadgeCountMode = "off" | "unread" | "total";
export type BadgePosition = "top-right" | "top-left" | "bottom-right" | "bottom-left" | "off";

export interface IrisMailSettings {
  clientId: string;
  authority: string;
  authMethod: AuthMethod;
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
  /** Map of tag name → plain-English definition used for yes/no classification. */
  tagDescriptions: Record<string, string>;
  /** Whether auto-tagging via Claude is enabled. */
  enableAutoTagging: boolean;
  /** Custom tag classification prompt (overrides default). */
  tagClassifyPrompt: string;
  /** Per-tag prompt version — incremented when the tag's definition or the meta-prompt changes. */
  tagPromptVersions: Record<string, number>;
  /** Max messages to prefetch Claude extraction for in background. 0 = disabled, -1 = all. */
  prefetchLimit: number;
  /** Show the original sender of forwarded emails instead of the forwarder. */
  resolveForwardedSender: boolean;
  /** Enable automatic event/task detection in emails. */
  enableAutoItemDetection: boolean;
  /** Folder for auto-created event notes. */
  eventNoteFolderPath: string;
  /** Folder for auto-created task notes. */
  taskNoteFolderPath: string;
  /** Custom prompt for event/task extraction (overrides default). */
  itemDetectionPrompt: string;
  // Persisted view state
  viewMode: "messages" | "senders";
  sortNewestFirst: boolean;
  filterUnreadOnly: boolean;
  /** Enable debug logging to console. */
  debugLogging: boolean;
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

export type AuthState = "signed-out" | "signing-in" | "signed-in" | "error";
