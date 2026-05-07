export interface BodyCacheEntry {
  messageId: string;
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

export interface MessageListCacheEntry {
  /** Raw Graph Message objects returned by the last successful list fetch. */
  messages: unknown[];
  nextLink: string | null;
  cachedAt: number;
}

/**
 * All per-message metadata, keyed on messageId under `EmailStoreIndex.messages`.
 * Every field is optional — an entry exists only for messages that have at
 * least one attribute set. Omit an entry entirely once it becomes empty so the
 * index doesn't accumulate tombstones.
 */
export interface MessageMetadata {
  /** Timestamp when the message was marked read locally. */
  readAt?: number;
  /** Timestamp when the message was flagged as a to-do. */
  todoAt?: number;
  /** Timestamp when the message was flagged as junk. */
  junkAt?: number;
  /** Timestamp when the message was pinned. Pinned messages are always
   *  injected into the inbox list regardless of sync window or read filter. */
  pinnedAt?: number;
  /** User-assigned and auto-predicted tags (non-empty array when present). */
  tags?: TagCacheEntry[];
}

export interface EmailStoreIndex {
  /** v1 = single-account; v2 = composite ids; v3 = consolidated per-message metadata. */
  version: 3;
  bodies: Record<string, BodyCacheEntry>;
  /** Last successful message list per `${folderId}:${showRead ? "all" : "unread"}`. */
  messageLists: Record<string, MessageListCacheEntry>;
  processed: Record<string, ProcessedCacheEntry>;
  nicknames: Record<string, NicknameCacheEntry>;
  /** Addresses whose nicknames were explicitly deleted by the user. */
  deletedNicknames: Record<string, number>;
  /** All per-message metadata: read/todo/junk flags, tags, detected items, scan state. */
  messages: Record<string, MessageMetadata>;
  /**
   * Cached Message envelopes (from/subject/receivedDateTime/bodyPreview/etc.)
   * for messages that should remain visible beyond the server's sync window.
   * A message's envelope is persisted when it's pinned OR it matches the
   * predicate of a box whose `saved` flag is true. Kept optional for
   * compatibility with older v3 caches — defaulted to `{}` by the store.
   */
  persistedEnvelopes?: Record<string, unknown>;
}
