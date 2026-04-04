import type { App } from "obsidian";
import { stripQuotedContent } from "../utils/stripQuotedContent";
import { extractForwardedSender } from "../utils/extractForwardedSender";
import { logger } from "../utils/logger";
import type { Message } from "../types";
import type {
  BodyCacheEntry,
  ProcessedCacheEntry,
  ClassificationCacheEntry,
  ImportanceClass,
  TagCacheEntry,
  DetectedItemEntry,
  DetectedItemStatus,
  EmailStoreIndex,
} from "./types";

const STORE_DIR = ".obsidian/plugins/iris-mail/cache";
const INDEX_PATH = `${STORE_DIR}/index.json`;
const BODIES_PATH = `${STORE_DIR}/bodies.json`;
/** Legacy single-file path — migrated on first load. */
const LEGACY_STORE_PATH = ".obsidian/plugins/iris-mail/email-store.json";

const SAVE_DEBOUNCE_MS = 2000;

function emptyIndex(): EmailStoreIndex {
  return { version: 1, bodies: {}, processed: {}, classifications: {}, nicknames: {}, readMessages: {}, tags: {}, detectedItems: {}, itemsScanned: {}, deletedNicknames: {} };
}

export class EmailStore {
  private app: App;
  private index: EmailStoreIndex = emptyIndex();
  private dirty = false;
  private bodiesDirty = false;
  private saveTimeout: number | null = null;

  constructor(app: App) {
    this.app = app;
  }

  // ── Lifecycle ──────────────────────────────────────────────

  async load(): Promise<void> {
    try {
      // Migrate from legacy single-file format
      if (await this.app.vault.adapter.exists(LEGACY_STORE_PATH)) {
        const raw = await this.app.vault.adapter.read(LEGACY_STORE_PATH);
        const parsed = JSON.parse(raw);
        if (parsed?.version === 1) {
          const defaults = emptyIndex();
          this.index = { ...defaults, ...parsed } as EmailStoreIndex;
          this.migrateTags();
          // Write split files and remove legacy
          await this.ensureCacheDir();
          await this.save();
          await this.app.vault.adapter.remove(LEGACY_STORE_PATH);
          logger.info("Store", "Migrated legacy email-store.json to split cache files");
          return;
        }
      }

      // Load split cache files
      if (await this.app.vault.adapter.exists(INDEX_PATH)) {
        const raw = await this.app.vault.adapter.read(INDEX_PATH);
        const parsed = JSON.parse(raw);
        if (parsed?.version === 1) {
          const defaults = emptyIndex();
          // Bodies are stored separately
          const bodies = await this.loadBodies();
          this.index = { ...defaults, ...parsed, bodies } as EmailStoreIndex;
          this.migrateTags();
        } else {
          logger.warn("Store", "Cache index version mismatch, resetting");
          this.index = emptyIndex();
        }
      }
    } catch (err) {
      logger.warn("Store", "Failed to load cache index, starting fresh", err);
      this.index = emptyIndex();
    }
  }

  async save(): Promise<void> {
    try {
      await this.ensureCacheDir();
      // Write index (everything except bodies) and bodies separately
      const { bodies, ...indexWithoutBodies } = this.index;
      await this.app.vault.adapter.write(INDEX_PATH, JSON.stringify(indexWithoutBodies));
      if (this.bodiesDirty) {
        await this.app.vault.adapter.write(BODIES_PATH, JSON.stringify(bodies));
        this.bodiesDirty = false;
      }
      this.dirty = false;
    } catch (err) {
      logger.warn("Store", "Failed to save cache index", err);
    }
  }

  private async loadBodies(): Promise<Record<string, BodyCacheEntry>> {
    try {
      if (await this.app.vault.adapter.exists(BODIES_PATH)) {
        return JSON.parse(await this.app.vault.adapter.read(BODIES_PATH));
      }
    } catch (err) {
      logger.warn("Store", "Failed to load bodies cache", err);
    }
    return {};
  }

  private migrateTags(): void {
    for (const [id, val] of Object.entries(this.index.tags)) {
      if (val && !Array.isArray(val) && (val as TagCacheEntry).tag) {
        (this.index.tags as Record<string, unknown>)[id] = [val];
      }
    }
  }

  private async ensureCacheDir(): Promise<void> {
    if (!(await this.app.vault.adapter.exists(STORE_DIR))) {
      await this.app.vault.adapter.mkdir(STORE_DIR);
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
      conversationId: msg.conversationId || "",
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

  getConversationBodies(conversationId: string): BodyCacheEntry[] {
    return Object.values(this.index.bodies)
      .filter((e) => e.conversationId === conversationId)
      .sort(
        (a, b) =>
          new Date(a.receivedDateTime).getTime() -
          new Date(b.receivedDateTime).getTime(),
      );
  }

  hasFullConversation(messageIds: string[]): boolean {
    return messageIds.every((id) => this.hasBody(id));
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
    fileContent: string,
    promptHash: string,
    saveFolderPath: string,
    msg: Message,
  ): Promise<ProcessedCacheEntry> {
    const vaultPath = this.buildVaultPath(saveFolderPath, msg);
    await this.ensureFolder(saveFolderPath);

    // Write the vault file with frontmatter
    const existing = this.app.vault.getAbstractFileByPath(vaultPath);
    if (existing) {
      await this.app.vault.modify(existing as never, fileContent);
    } else {
      await this.app.vault.create(vaultPath, fileContent);
    }

    // Cache the raw markdown (no frontmatter) for viewer display
    const entry: ProcessedCacheEntry = {
      messageId,
      promptHash,
      processedMarkdown: markdown,
      vaultPath,
      processedAt: Date.now(),
    };
    this.index.processed[messageId] = entry;
    this.scheduleSave();
    return entry;
  }

  // ── Classification cache ──────────────────────────────────

  getClassification(messageId: string): ImportanceClass | undefined {
    return this.index.classifications[messageId]?.classification;
  }

  setClassification(messageId: string, classification: ImportanceClass, source: "auto" | "manual" = "auto", promptVersion?: number): void {
    this.index.classifications[messageId] = {
      messageId,
      classification,
      source,
      promptVersion,
      classifiedAt: Date.now(),
    };
    this.scheduleSave();
  }

  /** Return all classification data in a single pass: classes, sources, and versions. */
  getAllClassificationData(): {
    classes: Map<string, ImportanceClass>;
    sources: Map<string, "auto" | "manual">;
    versions: Map<string, number | undefined>;
  } {
    const classes = new Map<string, ImportanceClass>();
    const sources = new Map<string, "auto" | "manual">();
    const versions = new Map<string, number | undefined>();
    for (const [id, entry] of Object.entries(this.index.classifications)) {
      classes.set(id, entry.classification);
      sources.set(id, entry.source || "auto");
      versions.set(id, entry.promptVersion);
    }
    return { classes, sources, versions };
  }

  getClassificationSource(messageId: string): "auto" | "manual" {
    return this.index.classifications[messageId]?.source || "auto";
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

  private buildVaultPath(folderPath: string, msg: Message): string {
    const date = msg.receivedDateTime
      ? new Date(msg.receivedDateTime).toISOString().split("T")[0]
      : "unknown";
    const subject = (msg.subject || "no-subject")
      .replace(/^(?:re|fw|fwd)\s*:\s*/i, "")  // strip reply/forward prefix
      .replace(/[\\/:*?"<>|]/g, "")            // remove illegal chars (don't replace with -)
      .replace(/\s+/g, " ")
      .trim()
      .slice(0, 50);
    // Short hash of message ID to avoid collisions
    const idHash = EmailStore.hashPrompt(msg.id || "").slice(0, 5);
    return `${folderPath}/${date} ${subject} ${idHash}.md`;
  }

  private async ensureFolder(path: string): Promise<void> {
    if (!(await this.app.vault.adapter.exists(path))) {
      await this.app.vault.createFolder(path);
    }
  }

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
