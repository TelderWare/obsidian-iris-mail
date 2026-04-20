import { classifyEmailTagsYesNo, hasClaudeAccess, type TagCandidate } from "../utils/claudeApi";
import { ConcurrencyPool } from "../utils/concurrency";
import { logger } from "../utils/logger";
import type { EmailStore } from "../store/EmailStore";
import { parseTagCategories, getTagVersion } from "../constants";
import type { TagCacheEntry } from "../store/types";
import type { Message, IrisMailSettings } from "../types";

const CLASSIFY_CONCURRENCY = 4;

export class EmailClassifier {
  private tagCache = new Map<string, TagCacheEntry[]>();
  private pool = new ConcurrencyPool(CLASSIFY_CONCURRENCY);

  constructor(
    private store: EmailStore,
    private getSettings: () => IrisMailSettings,
  ) {}

  get tags(): Map<string, TagCacheEntry[]> {
    return this.tagCache;
  }

  /** Reload all caches from the persistent store. */
  reloadCaches(): void {
    this.tagCache = this.store.getAllTags();
  }

  /** Auto-tag untagged messages using Claude API with concurrency. */
  async autoTagAllMessages(
    messages: Message[],
    onProgress?: () => void,
  ): Promise<void> {
    const s = this.getSettings();
    if (!s.enableAutoTagging || !s.enableClaudeProcessing || !hasClaudeAccess(s.anthropicApiKey)) return;

    const limit = s.autoTagLimit ?? 10;
    if (limit === 0) return;

    const candidates = this.getTagCandidates();
    if (candidates.length === 0) return;

    const untagged = messages.filter(
      (m) => m.id && !m.isRead && !this.tagCache.has(m.id),
    );
    const queue = limit === -1 ? untagged : untagged.slice(0, limit);
    if (queue.length === 0) return;

    logger.debug("Classifier", `Tagging ${queue.length} of ${untagged.length} unread untagged messages`);

    const promises = queue.map((msg) =>
      this.pool.run(async () => {
        const content = this.getBestContent(msg);
        if (!content || this.tagCache.has(msg.id!)) return;

        try {
          const tags = await classifyEmailTagsYesNo(
            s.anthropicApiKey, s.claudeModel,
            s.tagClassifyPrompt || DEFAULT_TAG_PROMPT,
            content, candidates,
            s.tagContradictions || {},
            s.tagPrecludes || {},
          );

          if (tags.length > 0 && !this.tagCache.has(msg.id!)) {
            const entries: TagCacheEntry[] = tags.map((tag) => ({
              messageId: msg.id!,
              tag,
              source: "auto" as const,
              promptVersion: getTagVersion(s.tagPromptVersions, tag),
              taggedAt: Date.now(),
            }));
            this.tagCache.set(msg.id!, entries);
            for (const tag of tags) {
              this.store.setTag(msg.id!, tag, "auto", getTagVersion(s.tagPromptVersions, tag));
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

  /**
   * Pick the richest available text for classification, in order of preference:
   * (1) Claude-extracted processed markdown — densest, already stripped of boilerplate;
   * (2) full stripped HTML body — good signal but noisier;
   * (3) subject + bodyPreview — last-resort shallow preview.
   */
  private getBestContent(msg: Message): string {
    if (!msg.id) return [msg.subject, msg.bodyPreview].filter(Boolean).join("\n");
    const processed = this.store.getProcessed(msg.id);
    if (processed?.processedMarkdown) {
      return [msg.subject, processed.processedMarkdown].filter(Boolean).join("\n");
    }
    const body = this.store.getBody(msg.id);
    if (body?.strippedHtml) {
      return [msg.subject, body.strippedHtml].filter(Boolean).join("\n");
    }
    return [msg.subject, msg.bodyPreview].filter(Boolean).join("\n");
  }

  private getTagCandidates(): TagCandidate[] {
    const s = this.getSettings();
    const names = parseTagCategories(s.tagCategories);
    const descriptions = s.tagDescriptions || {};
    return names.map((name) => ({ name, description: descriptions[name] || "" }));
  }
}

const DEFAULT_TAG_PROMPT =
  "You are an email tag classifier. Given an email and a list of tag categories, " +
  "return the tags that apply to this email.\n\n" +
  "Rules:\n" +
  "- Return ONLY a JSON array of strings, e.g. [\"Finance\", \"Projects\"]\n" +
  "- Only use tags from the provided list\n" +
  "- Return an empty array [] if no tags clearly apply\n" +
  "- Assign 1-3 tags maximum\n" +
  "- Be conservative — only assign tags you are confident about";
