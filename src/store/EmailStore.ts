import type { App, Plugin } from "obsidian";
import { stripQuotedContent } from "../utils/stripQuotedContent";
import { extractForwardedSender } from "../utils/extractForwardedSender";
import { logger } from "../utils/logger";
import type { Message } from "../types";
import type {
  BodyCacheEntry,
  ProcessedCacheEntry,
  TagCacheEntry,
  DetectedItemEntry,
  DetectedItemStatus,
  EmailStoreIndex,
  MessageListCacheEntry,
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
  return { version: 1, bodies: {}, messageLists: {}, processed: {}, nicknames: {}, readMessages: {}, tags: {}, detectedItems: {}, itemsScanned: {}, deletedNicknames: {} };
}

export class EmailStore {
  private app: App;
  private plugin: Plugin;
  private index: EmailStoreIndex = emptyIndex();
  private dirty = false;
  private bodiesDirty = false;
  private saveTimeout: number | null = null;

  constructor(plugin: Plugin) {
    this.plugin = plugin;
    this.app = plugin.app;
  }

  // ── Lifecycle ──────────────────────────────────────────────

  async load(): Promise<void> {
    try {
      const data = (await this.plugin.loadData()) ?? {};
      const cached = data[DATA_CACHE_KEY];
      if (cached?.version === 1) {
        this.index = { ...emptyIndex(), ...cached } as EmailStoreIndex;
        this.migrateTags();
        return;
      }

      // Migrate legacy cache files into data.json
      const migrated = await this.migrateLegacyCache();
      if (migrated) {
        this.migrateTags();
        await this.save();
      }
    } catch (err) {
      logger.warn("Store", "Failed to load cache, starting fresh", err);
      this.index = emptyIndex();
    }
  }

  private async migrateLegacyCache(): Promise<boolean> {
    const adapter = this.app.vault.adapter;
    try {
      if (await adapter.exists(LEGACY_SINGLE_STORE_PATH)) {
        const raw = await adapter.read(LEGACY_SINGLE_STORE_PATH);
        const parsed = JSON.parse(raw);
        if (parsed?.version === 1) {
          this.index = { ...emptyIndex(), ...parsed } as EmailStoreIndex;
          await adapter.remove(LEGACY_SINGLE_STORE_PATH);
          logger.info("Store", "Migrated legacy email-store.json into data.json");
          return true;
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
          this.index = { ...emptyIndex(), ...parsed, bodies } as EmailStoreIndex;
          try {
            await adapter.remove(LEGACY_INDEX_PATH);
            if (await adapter.exists(LEGACY_BODIES_PATH)) await adapter.remove(LEGACY_BODIES_PATH);
            if (await adapter.exists(LEGACY_STORE_DIR)) await adapter.rmdir(LEGACY_STORE_DIR, true);
          } catch (err) {
            logger.warn("Store", "Failed to clean up legacy cache files", err);
          }
          logger.info("Store", "Migrated legacy split cache files into data.json");
          return true;
        }
      }
    } catch (err) {
      logger.warn("Store", "Legacy cache migration failed", err);
    }
    return false;
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

  private migrateTags(): void {
    for (const [id, val] of Object.entries(this.index.tags)) {
      if (val && !Array.isArray(val) && (val as TagCacheEntry).tag) {
        (this.index.tags as Record<string, unknown>)[id] = [val];
      }
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

  // ── Read state cache ─────────────────────────────────────────

  isMarkedRead(messageId: string): boolean {
    return messageId in this.index.readMessages;
  }

  markRead(messageId: string): void {
    if (!this.index.readMessages[messageId]) {
      this.index.readMessages[messageId] = Date.now();
      this.scheduleSave();
    }
  }

  markUnread(messageId: string): void {
    if (this.index.readMessages[messageId]) {
      delete this.index.readMessages[messageId];
      this.scheduleSave();
    }
  }

  /** Return all message IDs that were marked read locally. */
  getLocallyReadIds(): string[] {
    return Object.keys(this.index.readMessages);
  }

  /** Remove a message ID from the local read set (after syncing to server). */
  clearLocalRead(messageId: string): void {
    if (this.index.readMessages[messageId]) {
      delete this.index.readMessages[messageId];
      this.scheduleSave();
    }
  }

  /** Apply cached read state to a list of messages (mutates in place). */
  applyReadState(messages: Message[]): void {
    for (const msg of messages) {
      if (msg.id && this.isMarkedRead(msg.id)) {
        msg.isRead = true;
      }
    }
  }

  // ── Tag cache ────────────────────────────────────────────────

  getTags(messageId: string): TagCacheEntry[] {
    return this.index.tags[messageId] || [];
  }

  /** Add or replace a single tag on a message (preserves other tags). */
  setTag(messageId: string, tag: string, source: "manual" | "auto", promptVersion?: number): void {
    const arr = this.index.tags[messageId] || [];
    const idx = arr.findIndex((e) => e.tag === tag);
    const entry: TagCacheEntry = { messageId, tag, source, promptVersion, taggedAt: Date.now() };
    if (idx >= 0) {
      arr[idx] = entry;
    } else {
      arr.push(entry);
    }
    this.index.tags[messageId] = arr;
    this.scheduleSave();
  }

  /** Remove a specific tag, or all tags if tag is omitted. */
  removeTag(messageId: string, tag?: string): void {
    if (!tag) {
      delete this.index.tags[messageId];
    } else {
      const arr = this.index.tags[messageId];
      if (!arr) return;
      const filtered = arr.filter((e) => e.tag !== tag);
      if (filtered.length === 0) {
        delete this.index.tags[messageId];
      } else {
        this.index.tags[messageId] = filtered;
      }
    }
    this.scheduleSave();
  }

  getAllTags(): Map<string, TagCacheEntry[]> {
    const map = new Map<string, TagCacheEntry[]>();
    for (const [id, entries] of Object.entries(this.index.tags)) {
      map.set(id, entries);
    }
    return map;
  }

  // ── Detected items cache ─────────────────────────────────────

  hasItemsScan(messageId: string): boolean {
    return messageId in this.index.itemsScanned;
  }

  setItemsScanned(messageId: string): void {
    this.index.itemsScanned[messageId] = Date.now();
    this.scheduleSave();
  }

  clearItemsScan(messageId: string): void {
    delete this.index.itemsScanned[messageId];
    delete this.index.detectedItems[messageId];
    this.scheduleSave();
  }

  getDetectedItems(messageId: string): DetectedItemEntry[] {
    return this.index.detectedItems[messageId] || [];
  }

  setDetectedItems(messageId: string, items: DetectedItemEntry[]): void {
    this.index.detectedItems[messageId] = items;
    this.scheduleSave();
  }

  updateDetectedItemStatus(messageId: string, itemId: string, status: DetectedItemStatus, vaultPath?: string): void {
    const items = this.index.detectedItems[messageId];
    if (!items) return;
    const item = items.find((i) => i.itemId === itemId);
    if (!item) return;
    item.status = status;
    item.resolvedAt = Date.now();
    if (vaultPath) item.vaultPath = vaultPath;
    this.scheduleSave();
  }

  getAllPendingItems(): Map<string, DetectedItemEntry[]> {
    const map = new Map<string, DetectedItemEntry[]>();
    for (const [msgId, items] of Object.entries(this.index.detectedItems)) {
      const pending = items.filter((i) => i.status === "pending");
      if (pending.length > 0) map.set(msgId, pending);
    }
    return map;
  }

  getPendingItemCount(): number {
    let count = 0;
    for (const items of Object.values(this.index.detectedItems)) {
      count += items.filter((i) => i.status === "pending").length;
    }
    return count;
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
