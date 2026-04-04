import { Notice } from "obsidian";
import { classifyEmailImportance, classifyEmailTags } from "../utils/claudeApi";
import { ConcurrencyPool } from "../utils/concurrency";
import { logger } from "../utils/logger";
import type { EmailStore } from "../store/EmailStore";
import { parseTagCategories } from "../constants";
import type { ImportanceClass, TagCacheEntry } from "../store/types";
import type { Message, IrisMailSettings } from "../types";

const CLASSIFY_CONCURRENCY = 4;

export class EmailClassifier {
  private classificationCache = new Map<string, ImportanceClass>();
  private classificationSourceCache = new Map<string, "auto" | "manual">();
  private classificationVersionCache = new Map<string, number | undefined>();
  private tagCache = new Map<string, TagCacheEntry[]>();
  private pool = new ConcurrencyPool(CLASSIFY_CONCURRENCY);

  constructor(
    private store: EmailStore,
    private getSettings: () => IrisMailSettings,
  ) {}

  get classifications(): Map<string, ImportanceClass> {
    return this.classificationCache;
  }

  get classificationSources(): Map<string, "auto" | "manual"> {
    return this.classificationSourceCache;
  }

  get classificationVersions(): Map<string, number | undefined> {
    return this.classificationVersionCache;
  }

  get tags(): Map<string, TagCacheEntry[]> {
    return this.tagCache;
  }

  /** Reload all caches from the persistent store. */
  reloadCaches(): void {
    const classData = this.store.getAllClassificationData();
    this.classificationCache = classData.classes;
    this.classificationSourceCache = classData.sources;
    this.classificationVersionCache = classData.versions;
    this.tagCache = this.store.getAllTags();
  }

  /** Classify a single message and update all caches. */
  async classifyAndCache(msgId: string, content: string): Promise<ImportanceClass> {
    const s = this.getSettings();
    const result = await classifyEmailImportance(
      s.anthropicApiKey, s.claudeModel,
      s.importanceClassifyPrompt || DEFAULT_IMPORTANCE_PROMPT,
      content,
    );
    const ver = s.importancePromptVersion;
    this.classificationCache.set(msgId, result);
    this.classificationSourceCache.set(msgId, "auto");
    this.classificationVersionCache.set(msgId, ver);
    this.store.setClassification(msgId, result, "auto", ver);
    return result;
  }

  /**
   * Classify all unclassified messages in parallel with a concurrency limit.
   * Throws on complete failure, individual message failures are logged.
   */
  async classifyAllMessages(
    messages: Message[],
    onProgress?: () => void,
  ): Promise<void> {
    const s = this.getSettings();
    if (!s.enableClaudeProcessing || !s.anthropicApiKey) return;

    const unclassified = messages.filter(
      (m) => m.id && !this.classificationCache.has(m.id),
    );
    if (unclassified.length === 0) return;

    logger.debug("Classifier", `Classifying ${unclassified.length} messages (concurrency=${CLASSIFY_CONCURRENCY})`);

    const promises = unclassified.map((msg) =>
      this.pool.run(async () => {
        const content = [msg.subject, msg.bodyPreview].filter(Boolean).join("\n");
        if (!content) return;
        try {
          await this.classifyAndCache(msg.id!, content);
          onProgress?.();
        } catch (err) {
          logger.warn("Classifier", `Classification failed for ${msg.id}`, err);
        }
      }),
    );

    await Promise.all(promises);
    logger.debug("Classifier", "Classification batch complete");
  }

  /** Auto-tag untagged messages using Claude API with concurrency. */
  async autoTagAllMessages(
    messages: Message[],
    onProgress?: () => void,
  ): Promise<void> {
    const s = this.getSettings();
    if (!s.enableAutoTagging || !s.enableClaudeProcessing || !s.anthropicApiKey) return;

    const categories = this.getTagCategories();
    if (categories.length === 0) return;

    const untagged = messages.filter(
      (m) => m.id && !this.tagCache.has(m.id),
    );
    if (untagged.length === 0) return;

    logger.debug("Classifier", `Tagging ${untagged.length} messages`);

    const promises = untagged.map((msg) =>
      this.pool.run(async () => {
        const content = [msg.subject, msg.bodyPreview].filter(Boolean).join("\n");
        if (!content || this.tagCache.has(msg.id!)) return;

        try {
          const tags = await classifyEmailTags(
            s.anthropicApiKey, s.claudeModel,
            s.tagClassifyPrompt || DEFAULT_TAG_PROMPT,
            content, categories,
          );

          if (tags.length > 0 && !this.tagCache.has(msg.id!)) {
            const tagVer = s.tagPromptVersion;
            const entries: TagCacheEntry[] = tags.map((tag) => ({
              messageId: msg.id!,
              tag,
              source: "auto" as const,
              promptVersion: tagVer,
              taggedAt: Date.now(),
            }));
            this.tagCache.set(msg.id!, entries);
            for (const tag of tags) {
              this.store.setTag(msg.id!, tag, "auto", tagVer);
            }
            onProgress?.();
          }
        } catch (err) {
          logger.warn("Classifier", `Tag classification failed for ${msg.id}`, err);
        }
      }),
    );

    await Promise.all(promises);
  }

  /** Set manual classification, returns the old classification. */
  setManualClassification(msgId: string, importance: ImportanceClass): ImportanceClass | undefined {
    const old = this.classificationCache.get(msgId);
    this.classificationCache.set(msgId, importance);
    this.classificationSourceCache.set(msgId, "manual");
    this.classificationVersionCache.delete(msgId);
    this.store.setClassification(msgId, importance, "manual");
    return old;
  }

  /** Remove all auto-assigned classifications, keeping manual ones. */
  clearAutoClassifications(): void {
    for (const [id] of this.classificationCache) {
      if (this.classificationSourceCache.get(id) === "manual") continue;
      this.classificationCache.delete(id);
      this.classificationSourceCache.delete(id);
    }
  }

  /** Remove all auto-assigned tags, keeping manual ones. */
  clearAutoTags(): void {
    for (const [id, entries] of this.tagCache) {
      const manual = entries.filter((e) => e.source === "manual");
      if (manual.length === 0) {
        this.tagCache.delete(id);
        this.store.removeTag(id);
      } else {
        this.tagCache.set(id, manual);
        for (const e of entries) {
          if (e.source === "auto") this.store.removeTag(id, e.tag);
        }
      }
    }
  }

  private getTagCategories(): string[] {
    return parseTagCategories(this.getSettings().tagCategories);
  }
}

// Re-import prompts to avoid circular dependency with constants
const DEFAULT_IMPORTANCE_PROMPT =
  "Classify this email as: important, routine, or noise.\n" +
  "important — requires attention, action, or contains personally relevant information\n" +
  "routine — informational, no action needed\n" +
  "noise — marketing, automated, newsletters, or zero-value content\n" +
  "Return only the single word.";

const DEFAULT_TAG_PROMPT =
  "You are an email tag classifier. Given an email and a list of tag categories, " +
  "return the tags that apply to this email.\n\n" +
  "Rules:\n" +
  "- Return ONLY a JSON array of strings, e.g. [\"Finance\", \"Projects\"]\n" +
  "- Only use tags from the provided list\n" +
  "- Return an empty array [] if no tags clearly apply\n" +
  "- Assign 1-3 tags maximum\n" +
  "- Be conservative — only assign tags you are confident about";
