import type { App, Plugin } from "obsidian";
import { stripQuotedContent } from "../utils/stripQuotedContent";
import { extractForwardedSender } from "../utils/extractForwardedSender";
import { logger } from "../utils/logger";
import type { Box, Message } from "../types";
import type {
  BodyCacheEntry,
  ProcessedCacheEntry,
  TagCacheEntry,
  EmailStoreIndex,
  MessageListCacheEntry,
  MessageMetadata,
} from "./types";

/** Legacy paths migrated into data.json on first load. */
const LEGACY_STORE_DIR = ".obsidian/plugins/iris-mail/cache";
const LEGACY_INDEX_PATH = `${LEGACY_STORE_DIR}/index.json`;
const LEGACY_BODIES_PATH = `${LEGACY_STORE_DIR}/bodies.json`;
const LEGACY_SINGLE_STORE_PATH = ".obsidian/plugins/iris-mail/email-store.json";

/** Key under which the cache is stored in the plugin's data.json. */
const DATA_CACHE_KEY = "__cache";

const SAVE_DEBOUNCE_MS = 2000;

function emptyIndex(): EmailStoreIndex {
  return {
    version: 3,
    bodies: {},
    messageLists: {},
    processed: {},
    nicknames: {},
    deletedNicknames: {},
    messages: {},
    persistedEnvelopes: {},
  };
}

/** Rewrite an existing v1 (single-account, native ids) blob into v2 shape. */
function migrateV1ToV2(legacy: Record<string, unknown>, accountId: string): Record<string, unknown> {
  const prefix = (id: string) => `${accountId}:${id}`;
  const remap = <T>(obj: Record<string, T> | undefined): Record<string, T> => {
    const out: Record<string, T> = {};
    for (const [k, v] of Object.entries(obj ?? {})) out[prefix(k)] = v;
    return out;
  };
  const remapBodies = (obj: Record<string, BodyCacheEntry> | undefined): Record<string, BodyCacheEntry> => {
    const out: Record<string, BodyCacheEntry> = {};
    for (const [k, v] of Object.entries(obj ?? {})) {
      const newKey = prefix(k);
      out[newKey] = { ...v, messageId: newKey };
    }
    return out;
  };
  const remapTags = (obj: Record<string, TagCacheEntry[]> | undefined): Record<string, TagCacheEntry[]> => {
    const out: Record<string, TagCacheEntry[]> = {};
    for (const [k, arr] of Object.entries(obj ?? {})) {
      const newKey = prefix(k);
      out[newKey] = arr.map((e) => ({ ...e, messageId: newKey }));
    }
    return out;
  };
  return {
    version: 2,
    bodies: remapBodies(legacy.bodies as Record<string, BodyCacheEntry>),
    // Drop messageLists — old folder ids are no longer meaningful in unified mode.
    messageLists: {},
    processed: remap(legacy.processed as Record<string, ProcessedCacheEntry>),
    nicknames: (legacy.nicknames as Record<string, unknown>) ?? {},
    readMessages: remap(legacy.readMessages as Record<string, number>),
    todoMessages: remap(legacy.todoMessages as Record<string, number>),
    junkMessages: remap(legacy.junkMessages as Record<string, number>),
    tags: remapTags(legacy.tags as Record<string, TagCacheEntry[]>),
    deletedNicknames: (legacy.deletedNicknames as Record<string, number>) ?? {},
  };
}

/**
 * Fold the v2 per-message caches (readMessages, todoMessages, junkMessages,
 * tags) into a single `messages` record keyed by messageId. Legacy
 * `itemsScanned` and `detectedItems` records are dropped silently — they
 * belonged to the removed auto-detection feature.
 */
function migrateV2ToV3(v2: Record<string, unknown>): EmailStoreIndex {
  const messages: Record<string, MessageMetadata> = {};

  const touch = (id: string): MessageMetadata => (messages[id] ??= {});

  const readMessages = (v2.readMessages as Record<string, number>) ?? {};
  for (const [id, ts] of Object.entries(readMessages)) touch(id).readAt = ts;

  const todoMessages = (v2.todoMessages as Record<string, number>) ?? {};
  for (const [id, ts] of Object.entries(todoMessages)) touch(id).todoAt = ts;

  const junkMessages = (v2.junkMessages as Record<string, number>) ?? {};
  for (const [id, ts] of Object.entries(junkMessages)) touch(id).junkAt = ts;

  const tags = (v2.tags as Record<string, TagCacheEntry[] | TagCacheEntry>) ?? {};
  for (const [id, value] of Object.entries(tags)) {
    // v2 briefly allowed a single entry under this key; normalize to array.
    const arr = Array.isArray(value) ? value : [value as TagCacheEntry];
    if (arr.length > 0) touch(id).tags = arr;
  }

  return {
    version: 3,
    bodies: (v2.bodies as EmailStoreIndex["bodies"]) ?? {},
    messageLists: (v2.messageLists as EmailStoreIndex["messageLists"]) ?? {},
    processed: (v2.processed as EmailStoreIndex["processed"]) ?? {},
    nicknames: (v2.nicknames as EmailStoreIndex["nicknames"]) ?? {},
    deletedNicknames: (v2.deletedNicknames as EmailStoreIndex["deletedNicknames"]) ?? {},
    messages,
  };
}

export class EmailStore {
  private app: App;
  private plugin: Plugin;
  private index: EmailStoreIndex = emptyIndex();
  private dirty = false;
  private bodiesDirty = false;
  private saveTimeout: number | null = null;
  private processedListeners = new Set<(messageId: string) => void>();

  constructor(plugin: Plugin) {
    this.plugin = plugin;
    this.app = plugin.app;
  }

  /** Subscribe to processed-cache writes. Fired after a summary is stored
   *  for a message. Used by the plugin to nudge Iris Tasks once a flagged
   *  to-do becomes eligible (i.e. its summary lands). */
  onProcessedChanged(cb: (messageId: string) => void): () => void {
    this.processedListeners.add(cb);
    return () => {
      this.processedListeners.delete(cb);
    };
  }

  // ── Lifecycle ──────────────────────────────────────────────

  async load(): Promise<void> {
    try {
      const data = (await this.plugin.loadData()) ?? {};
      const cached = data[DATA_CACHE_KEY];
      if (cached?.version === 3) {
        this.index = { ...emptyIndex(), ...cached } as EmailStoreIndex;
        // Carry forward envelopes from the short-lived `pinnedEnvelopes` field
        // (same shape, renamed as the feature generalized).
        const legacy = (cached as Record<string, unknown>).pinnedEnvelopes as
          | Record<string, unknown>
          | undefined;
        if (legacy) {
          const current = this.index.persistedEnvelopes ?? (this.index.persistedEnvelopes = {});
          for (const [id, env] of Object.entries(legacy)) {
            if (!(id in current)) current[id] = env;
          }
          delete (this.index as unknown as Record<string, unknown>).pinnedEnvelopes;
          this.scheduleSave();
        }
        return;
      }
      if (cached?.version === 2) {
        this.index = migrateV2ToV3(cached as Record<string, unknown>);
        await this.save();
        logger.info("Store", "Migrated v2 cache to v3 (consolidated per-message metadata)");
        return;
      }
      if (cached?.version === 1) {
        const accountId = this.firstAccountId();
        if (accountId) {
          const v2 = migrateV1ToV2(cached as Record<string, unknown>, accountId);
          this.index = migrateV2ToV3(v2);
          await this.save();
          logger.info("Store", `Migrated v1 cache through v2→v3 under account ${accountId}`);
        } else {
          this.index = emptyIndex();
          await this.save();
          logger.warn("Store", "v1 cache dropped: no account exists to migrate it under");
        }
        return;
      }

      // Migrate legacy file-based cache into data.json (then composite-migrate).
      const legacyV1 = await this.readLegacyFileCache();
      if (legacyV1) {
        const accountId = this.firstAccountId();
        if (accountId) {
          const v2 = migrateV1ToV2(legacyV1, accountId);
          this.index = migrateV2ToV3(v2);
        } else {
          this.index = emptyIndex();
        }
        await this.save();
      }
    } catch (err) {
      logger.warn("Store", "Failed to load cache, starting fresh", err);
      this.index = emptyIndex();
    }
  }

  private firstAccountId(): string | undefined {
    const settings = (this.plugin as unknown as { settings?: { accounts?: Array<{ id: string }> } }).settings;
    return settings?.accounts?.[0]?.id;
  }

  /** Read and remove legacy on-disk v1 cache files. Returns a v1 blob or null. */
  private async readLegacyFileCache(): Promise<Record<string, unknown> | null> {
    const adapter = this.app.vault.adapter;
    try {
      if (await adapter.exists(LEGACY_SINGLE_STORE_PATH)) {
        const raw = await adapter.read(LEGACY_SINGLE_STORE_PATH);
        const parsed = JSON.parse(raw);
        if (parsed?.version === 1) {
          await adapter.remove(LEGACY_SINGLE_STORE_PATH);
          logger.info("Store", "Read legacy email-store.json for migration");
          return parsed;
        }
      }
      if (await adapter.exists(LEGACY_INDEX_PATH)) {
        const raw = await adapter.read(LEGACY_INDEX_PATH);
        const parsed = JSON.parse(raw);
        if (parsed?.version === 1) {
          let bodies: Record<string, BodyCacheEntry> = {};
          if (await adapter.exists(LEGACY_BODIES_PATH)) {
            try {
              bodies = JSON.parse(await adapter.read(LEGACY_BODIES_PATH));
            } catch (err) {
              logger.warn("Store", "Failed to read legacy bodies.json", err);
            }
          }
          try {
            await adapter.remove(LEGACY_INDEX_PATH);
            if (await adapter.exists(LEGACY_BODIES_PATH)) await adapter.remove(LEGACY_BODIES_PATH);
            if (await adapter.exists(LEGACY_STORE_DIR)) await adapter.rmdir(LEGACY_STORE_DIR, true);
          } catch (err) {
            logger.warn("Store", "Failed to clean up legacy cache files", err);
          }
          logger.info("Store", "Read legacy split cache files for migration");
          return { ...parsed, bodies };
        }
      }
    } catch (err) {
      logger.warn("Store", "Legacy cache migration failed", err);
    }
    return null;
  }

  async save(): Promise<void> {
    try {
      // Read-modify-write to avoid clobbering plugin settings that share data.json
      const data = (await this.plugin.loadData()) ?? {};
      data[DATA_CACHE_KEY] = this.index;
      await this.plugin.saveData(data);
      this.dirty = false;
      this.bodiesDirty = false;
    } catch (err) {
      logger.warn("Store", "Failed to save cache", err);
    }
  }

  async flush(): Promise<void> {
    if (this.saveTimeout !== null) {
      window.clearTimeout(this.saveTimeout);
      this.saveTimeout = null;
    }
    if (this.dirty) {
      await this.save();
    }
  }

  // ── Per-message metadata helpers ──────────────────────────

  private getMeta(messageId: string): MessageMetadata | undefined {
    return this.index.messages[messageId];
  }

  /** Mutate (or create) the metadata entry for a message, then save. */
  private mutateMeta(messageId: string, fn: (meta: MessageMetadata) => void): void {
    const meta = this.index.messages[messageId] ?? {};
    fn(meta);
    if (isEmptyMeta(meta)) {
      delete this.index.messages[messageId];
    } else {
      this.index.messages[messageId] = meta;
    }
    this.scheduleSave();
  }

  // ── Body cache ─────────────────────────────────────────────

  hasBody(messageId: string): boolean {
    return messageId in this.index.bodies;
  }

  getBody(messageId: string): BodyCacheEntry | undefined {
    return this.index.bodies[messageId];
  }

  setBody(msg: Message, bodyHtml: string): BodyCacheEntry {
    const stripped = stripQuotedContent(bodyHtml);
    const subject = msg.subject || "";
    const isForward = /^(?:fw|fwd)\s*:/i.test(subject);
    const originalSender = isForward
      ? extractForwardedSender(bodyHtml) ?? undefined
      : undefined;
    const entry: BodyCacheEntry = {
      messageId: msg.id!,
      subject,
      from: msg.from?.emailAddress?.address || "",
      receivedDateTime: msg.receivedDateTime || "",
      bodyHtml,
      strippedHtml: stripped,
      originalSender,
      cachedAt: Date.now(),
    };
    this.index.bodies[msg.id!] = entry;
    this.bodiesDirty = true;
    this.scheduleSave();
    return entry;
  }

  // ── Message list cache ─────────────────────────────────────

  private listKey(folderId: string, showRead: boolean): string {
    return `${folderId}:${showRead ? "all" : "unread"}`;
  }

  getMessageList(folderId: string, showRead: boolean): MessageListCacheEntry | undefined {
    return this.index.messageLists[this.listKey(folderId, showRead)];
  }

  setMessageList(folderId: string, showRead: boolean, messages: Message[], nextLink: string | null): void {
    this.index.messageLists[this.listKey(folderId, showRead)] = {
      messages,
      nextLink,
      cachedAt: Date.now(),
    };
    this.scheduleSave();
  }

  // ── Processed cache ────────────────────────────────────────

  hasProcessed(messageId: string, currentPromptHash?: string): boolean {
    const entry = this.index.processed[messageId];
    if (!entry) return false;
    if (currentPromptHash !== undefined) return entry.promptHash === currentPromptHash;
    return true;
  }

  getProcessed(messageId: string): ProcessedCacheEntry | undefined {
    return this.index.processed[messageId];
  }

  clearProcessed(messageId: string): void {
    delete this.index.processed[messageId];
    this.scheduleSave();
  }

  async setProcessed(
    messageId: string,
    markdown: string,
    promptHash: string,
  ): Promise<ProcessedCacheEntry> {
    // Stored in data.json only; no vault file is created.
    const entry: ProcessedCacheEntry = {
      messageId,
      promptHash,
      processedMarkdown: markdown,
      vaultPath: "",
      processedAt: Date.now(),
    };
    this.index.processed[messageId] = entry;
    this.scheduleSave();
    for (const cb of Array.from(this.processedListeners)) {
      try { cb(messageId); } catch (err) { logger.warn("Store", "processedChanged listener failed", err); }
    }
    return entry;
  }

  // ── Nickname cache ────────────────────────────────────────

  getNickname(address: string): string | undefined {
    return this.index.nicknames[address.toLowerCase()]?.nickname;
  }

  setNickname(address: string, nickname: string): void {
    const key = address.toLowerCase();
    this.index.nicknames[key] = {
      address: key,
      nickname,
      generatedAt: Date.now(),
    };
    delete this.index.deletedNicknames[key];
    this.scheduleSave();
  }

  deleteNickname(address: string): void {
    const key = address.toLowerCase();
    delete this.index.nicknames[key];
    this.index.deletedNicknames[key] = Date.now();
    this.scheduleSave();
  }

  isNicknameDeleted(address: string): boolean {
    return address.toLowerCase() in this.index.deletedNicknames;
  }

  getAllNicknames(): Map<string, string> {
    const map = new Map<string, string>();
    for (const [addr, entry] of Object.entries(this.index.nicknames)) {
      map.set(addr, entry.nickname);
    }
    return map;
  }

  // ── Read state ────────────────────────────────────────────

  isMarkedRead(messageId: string): boolean {
    return this.getMeta(messageId)?.readAt !== undefined;
  }

  markRead(messageId: string): void {
    if (this.isMarkedRead(messageId)) return;
    this.mutateMeta(messageId, (m) => { m.readAt = Date.now(); });
  }

  markUnread(messageId: string): void {
    if (!this.isMarkedRead(messageId)) return;
    this.mutateMeta(messageId, (m) => { delete m.readAt; });
  }

  /** Return all message IDs that were marked read locally. */
  getLocallyReadIds(): string[] {
    const out: string[] = [];
    for (const [id, meta] of Object.entries(this.index.messages)) {
      if (meta.readAt !== undefined) out.push(id);
    }
    return out;
  }

  /** Remove a message ID from the local read set (after syncing to server). */
  clearLocalRead(messageId: string): void {
    this.markUnread(messageId);
  }

  /** Apply cached read state to a list of messages (mutates in place). */
  applyReadState(messages: Message[]): void {
    for (const msg of messages) {
      if (msg.id && this.isMarkedRead(msg.id)) {
        msg.isRead = true;
      }
    }
  }

  // ── To-do state ────────────────────────────────────────────

  isMarkedTodo(messageId: string): boolean {
    return this.getMeta(messageId)?.todoAt !== undefined;
  }

  getTodoAt(messageId: string): number | undefined {
    return this.getMeta(messageId)?.todoAt;
  }

  markTodo(messageId: string): void {
    if (this.isMarkedTodo(messageId)) return;
    this.mutateMeta(messageId, (m) => { m.todoAt = Date.now(); });
  }

  unmarkTodo(messageId: string): void {
    if (!this.isMarkedTodo(messageId)) return;
    this.mutateMeta(messageId, (m) => { delete m.todoAt; });
  }

  getAllTodoIds(): Set<string> {
    const out = new Set<string>();
    for (const [id, meta] of Object.entries(this.index.messages)) {
      if (meta.todoAt !== undefined) out.add(id);
    }
    return out;
  }

  // ── Junk state ─────────────────────────────────────────────

  isMarkedJunk(messageId: string): boolean {
    return this.getMeta(messageId)?.junkAt !== undefined;
  }

  markJunk(messageId: string): void {
    if (this.isMarkedJunk(messageId)) return;
    this.mutateMeta(messageId, (m) => { m.junkAt = Date.now(); });
  }

  unmarkJunk(messageId: string): void {
    if (!this.isMarkedJunk(messageId)) return;
    this.mutateMeta(messageId, (m) => { delete m.junkAt; });
  }

  getAllJunkIds(): Set<string> {
    const out = new Set<string>();
    for (const [id, meta] of Object.entries(this.index.messages)) {
      if (meta.junkAt !== undefined) out.add(id);
    }
    return out;
  }

  // ── Pinned state ──────────────────────────────────────────

  isMarkedPinned(messageId: string): boolean {
    return this.getMeta(messageId)?.pinnedAt !== undefined;
  }

  /**
   * Pin a message. The envelope is copied into `persistedEnvelopes` so the
   * message can be re-injected into the inbox even when it falls outside the
   * server's current window. Repeat calls update the envelope snapshot with
   * whatever fresh data the caller has.
   */
  markPinned(messageId: string, envelope: Message): void {
    this.mutateMeta(messageId, (m) => {
      if (m.pinnedAt === undefined) m.pinnedAt = Date.now();
    });
    this.upsertEnvelope(messageId, envelope);
  }

  unmarkPinned(messageId: string): void {
    this.mutateMeta(messageId, (m) => { delete m.pinnedAt; });
    // Leave the envelope in place — mergePersistedMessages will prune it on
    // the next load if it no longer matches any saved box either.
  }

  getAllPinnedIds(): Set<string> {
    const out = new Set<string>();
    for (const [id, meta] of Object.entries(this.index.messages)) {
      if (meta.pinnedAt !== undefined) out.add(id);
    }
    return out;
  }

  // ── Persisted envelopes (pinned + saved-box messages) ────

  /** Read-only envelope lookup for persisted messages (pinned or saved-box). */
  getPersistedEnvelope(messageId: string): Message | undefined {
    return this.index.persistedEnvelopes?.[messageId] as Message | undefined;
  }

  private upsertEnvelope(messageId: string, envelope: Message): void {
    const envelopes =
      this.index.persistedEnvelopes ?? (this.index.persistedEnvelopes = {});
    envelopes[messageId] = snapshotEnvelope(envelope);
    this.scheduleSave();
  }

  /**
   * True when `messageId` still belongs in `box` according to locally-tracked
   * state (todo/junk flags, tags). Evaluated without needing the full Message
   * so it can run on messages that fall outside the server window. In/Read/
   * Secretary return false — they're dynamic and can't be "saved".
   */
  savedBoxContains(messageId: string, box: Box): boolean {
    const boxTags = box.tags ?? [];
    const tags = this.getTags(messageId);
    const hasBoxTag =
      boxTags.length > 0 &&
      tags.some((t) => boxTags.includes(t.tag));
    switch (box.builtin) {
      case "todo":
        return this.isMarkedTodo(messageId) || hasBoxTag;
      case "junk":
        return this.isMarkedJunk(messageId) || hasBoxTag;
      case undefined:
        return hasBoxTag;
      default:
        return false;
    }
  }

  /**
   * Return `fetched` augmented with any pinned or saved-box messages the
   * server didn't return, and refresh cached envelopes from whatever fresh
   * data came in. Envelopes that no longer match any reason to be kept
   * (pin removed, todo cleared, tag lost, saved flag turned off) are pruned.
   *
   * The result is re-sorted by date so callers can use it directly.
   */
  mergePersistedMessages(fetched: Message[], savedBoxes: readonly Box[]): Message[] {
    const envelopes =
      this.index.persistedEnvelopes ?? (this.index.persistedEnvelopes = {});
    const byId = new Map<string, Message>();
    for (const m of fetched) if (m.id) byId.set(m.id, m);

    const pinned = this.getAllPinnedIds();
    const shouldKeep = (id: string): boolean => {
      if (pinned.has(id)) return true;
      for (const box of savedBoxes) {
        if (this.savedBoxContains(id, box)) return true;
      }
      return false;
    };

    let changed = false;

    // Snapshot envelopes for everything in the fresh list that deserves to
    // be kept, so the cache tracks the latest subject / read state / etc.
    for (const [id, msg] of byId) {
      if (shouldKeep(id)) {
        envelopes[id] = snapshotEnvelope(msg);
        changed = true;
      }
    }

    // Reinject envelopes the server didn't return; prune ones that no longer
    // have a reason to stick around.
    for (const id of Object.keys(envelopes)) {
      if (byId.has(id)) continue;
      if (shouldKeep(id)) {
        byId.set(id, envelopes[id] as Message);
      } else {
        delete envelopes[id];
        changed = true;
      }
    }

    if (changed) this.scheduleSave();

    return Array.from(byId.values()).sort((a, b) => {
      const ad = a.receivedDateTime ?? "";
      const bd = b.receivedDateTime ?? "";
      return bd.localeCompare(ad);
    });
  }

  // ── Tags ──────────────────────────────────────────────────

  getTags(messageId: string): TagCacheEntry[] {
    return this.getMeta(messageId)?.tags ?? [];
  }

  /** Add or replace a single tag on a message (preserves other tags). */
  setTag(messageId: string, tag: string, source: "manual" | "auto", promptVersion?: number): void {
    this.mutateMeta(messageId, (m) => {
      const arr = m.tags ?? [];
      const idx = arr.findIndex((e) => e.tag === tag);
      const entry: TagCacheEntry = { messageId, tag, source, promptVersion, taggedAt: Date.now() };
      if (idx >= 0) arr[idx] = entry;
      else arr.push(entry);
      m.tags = arr;
    });
  }

  /** Remove a specific tag, or all tags if tag is omitted. */
  removeTag(messageId: string, tag?: string): void {
    this.mutateMeta(messageId, (m) => {
      if (!m.tags) return;
      if (!tag) {
        delete m.tags;
        return;
      }
      const filtered = m.tags.filter((e) => e.tag !== tag);
      if (filtered.length === 0) delete m.tags;
      else m.tags = filtered;
    });
  }

  getAllTags(): Map<string, TagCacheEntry[]> {
    const map = new Map<string, TagCacheEntry[]>();
    for (const [id, meta] of Object.entries(this.index.messages)) {
      if (meta.tags && meta.tags.length > 0) map.set(id, meta.tags);
    }
    return map;
  }

  // ── Prompt hashing ─────────────────────────────────────────

  /**
   * SHA-256-based prompt hash. Returns a 12-char hex string.
   * Falls back to a safe DJB2 implementation if crypto.subtle is unavailable.
   */
  static hashPrompt(prompt: string): string {
    // Synchronous fallback using DJB2 with Math.abs to avoid dash issues
    let hash = 5381;
    for (let i = 0; i < prompt.length; i++) {
      hash = ((hash << 5) + hash + prompt.charCodeAt(i)) | 0;
    }
    return Math.abs(hash).toString(36).padStart(7, "0");
  }

  // ── Internal ───────────────────────────────────────────────

  private scheduleSave(): void {
    this.dirty = true;
    if (this.saveTimeout !== null) return;
    this.saveTimeout = window.setTimeout(() => {
      this.saveTimeout = null;
      if (this.dirty) {
        void this.save();
      }
    }, SAVE_DEBOUNCE_MS);
  }
}

/** True when a MessageMetadata has no attributes set and can be dropped from the index. */
function isEmptyMeta(m: MessageMetadata): boolean {
  if (m.readAt !== undefined) return false;
  if (m.todoAt !== undefined) return false;
  if (m.junkAt !== undefined) return false;
  if (m.pinnedAt !== undefined) return false;
  if (m.tags && m.tags.length > 0) return false;
  return true;
}

/**
 * Build a minimal persistable copy of a Message. Drops transient fan-out
 * fields (`_accountId` / `_accountLabel` are re-attached at dispatch time)
 * and leaves heavy HTML bodies in the body cache where they already live.
 */
function snapshotEnvelope(msg: Message): Message {
  const snap: Message = {
    id: msg.id,
    subject: msg.subject,
    bodyPreview: msg.bodyPreview,
    receivedDateTime: msg.receivedDateTime,
    sentDateTime: msg.sentDateTime,
    isRead: msg.isRead,
    hasAttachments: msg.hasAttachments,
    from: msg.from,
    sender: msg.sender,
    toRecipients: msg.toRecipients,
    ccRecipients: msg.ccRecipients,
    conversationId: msg.conversationId,
    internetMessageId: msg.internetMessageId,
  };
  // Preserve the fan-out tags — they let the inbox still route pinned
  // messages to the right account column after a reload.
  if (msg._accountId) snap._accountId = msg._accountId;
  if (msg._accountLabel) snap._accountLabel = msg._accountLabel;
  return snap;
}
