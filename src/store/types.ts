export interface BodyCacheEntry {
  messageId: string;
  conversationId: string;
  subject: string;
  from: string;
  receivedDateTime: string;
  bodyHtml: string;
  strippedHtml: string;
  /** Original sender extracted from a forwarded email body, if detected. */
  originalSender?: { name?: string; address?: string };
  cachedAt: number;
}

export interface ProcessedCacheEntry {
  messageId: string;
  promptHash: string;
  processedMarkdown: string;
  vaultPath: string;
  processedAt: number;
}

export type ImportanceClass = "important" | "routine" | "noise";

export interface ClassificationCacheEntry {
  messageId: string;
  classification: ImportanceClass;
  /** "manual" if user-assigned, "auto" if classifier-predicted. */
  source: "auto" | "manual";
  /** Importance prompt version that produced this classification (auto only). */
  promptVersion?: number;
  classifiedAt: number;
}

export interface NicknameCacheEntry {
  address: string;
  nickname: string;
  generatedAt: number;
}

export interface TagCacheEntry {
  messageId: string;
  tag: string;
  /** "manual" if user-assigned, "auto" if classifier-predicted. */
  source: "manual" | "auto";
  /** Tag prompt version that produced this tag (auto only). */
  promptVersion?: number;
  taggedAt: number;
}

export type DetectedItemStatus = "pending" | "accepted" | "dismissed";

export interface DetectedItemEntry {
  /** Unique ID for this detected item (messageId + index). */
  itemId: string;
  messageId: string;
  type: "event" | "task";
  title: string;
  date?: string;
  time?: string;
  location?: string;
  dueDate?: string;
  priority?: "high" | "medium" | "low";
  description: string;
  /** Verbatim excerpt from the email body this item was detected from. */
  sourceText?: string;
  status: DetectedItemStatus;
  /** Vault path if accepted and saved. */
  vaultPath?: string;
  detectedAt: number;
  resolvedAt?: number;
}

export interface EmailStoreIndex {
  version: 1;
  bodies: Record<string, BodyCacheEntry>;
  processed: Record<string, ProcessedCacheEntry>;
  classifications: Record<string, ClassificationCacheEntry>;
  nicknames: Record<string, NicknameCacheEntry>;
  /** Message IDs marked as read → timestamp when marked. */
  readMessages: Record<string, number>;
  /** User-assigned and auto-predicted tags keyed by messageId (array per message). */
  tags: Record<string, TagCacheEntry[]>;
  /** Auto-detected events and tasks keyed by messageId. */
  detectedItems: Record<string, DetectedItemEntry[]>;
  /** Message IDs that have been scanned for items -> timestamp when scanned. */
  itemsScanned: Record<string, number>;
  /** Addresses whose nicknames were explicitly deleted by the user. */
  deletedNicknames: Record<string, number>;
}
