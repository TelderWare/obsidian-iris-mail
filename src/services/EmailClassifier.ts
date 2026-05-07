import { classifyEmailTagsYesNo, hasClaudeAccess, type TagCandidate } from "../utils/claudeApi";
import { ConcurrencyPool } from "../utils/concurrency";
import { logger } from "../utils/logger";
import type { EmailStore } from "../store/EmailStore";
import { parseTagCategories, getTagVersion } from "../constants";
import type { Message, IrisMailSettings } from "../types";

const CLASSIFY_CONCURRENCY = 4;

export class EmailClassifier {
  private pool = new ConcurrencyPool(CLASSIFY_CONCURRENCY);
  private inFlight = new Set<string>();
  private inFlightListeners = new Set<() => void>();

  constructor(
    private store: EmailStore,
    private getSettings: () => IrisMailSettings,
  ) {}

  /** Messages currently being processed by the auto-tagger. */
  getInFlightIds(): Set<string> {
    return new Set(this.inFlight);
  }

  /** Subscribe to in-flight set changes. Returns an unsubscribe fn. */
  onInFlightChange(cb: () => void): () => void {
    this.inFlightListeners.add(cb);
    return () => { this.inFlightListeners.delete(cb); };
  }

  private notifyInFlight(): void {
    for (const cb of this.inFlightListeners) {
      try { cb(); } catch (err) { logger.warn("Classifier", "in-flight listener threw", err); }
    }
  }

  /** Auto-tag untagged messages with all user-defined tags. */
  async autoTagAllMessages(
    messages: Message[],
    onProgress?: () => void,
  ): Promise<void> {
    await this.autoTagMessages(messages, this.getTagCandidates(), onProgress);
  }

  /**
   * Classify `messages` against `candidates` (a subset of the user's tags) and
   * merge any matches into the store. Used to run multi-stage passes — e.g. a
   * junk-tag pass before a rest-of-tags pass — without short-circuiting on
   * messages that already carry tags from a prior pass.
   */
  async autoTagMessages(
    messages: Message[],
    candidates: TagCandidate[],
    onProgress?: () => void,
  ): Promise<void> {
    const s = this.getSettings();
    if (!s.enableAutoTagging || !s.enableClaudeProcessing || !hasClaudeAccess(s.anthropicApiKey)) return;

    const limit = s.autoTagLimit ?? 10;
    if (limit === 0) return;
    if (candidates.length === 0) return;

    // Eligible = unread messages missing at least one of the candidate tags.
    const eligible = messages.filter((m) => {
      if (!m.id || m.isRead) return false;
      const existing = new Set(this.store.getTags(m.id).map((e) => e.tag));
      return candidates.some((c) => !existing.has(c.name));
    });
    const queue = limit === -1 ? eligible : eligible.slice(0, limit);
    if (queue.length === 0) return;

    logger.debug("Classifier", `Tagging ${queue.length} of ${eligible.length} eligible messages against ${candidates.length} candidate(s)`);

    const promises = queue.map((msg) =>
      this.pool.run(async () => {
        const content = this.getBestContent(msg);
        if (!content) return;

        const existingNames = new Set(this.store.getTags(msg.id!).map((e) => e.tag));
        const toAsk = candidates.filter((c) => !existingNames.has(c.name));
        if (toAsk.length === 0) return;

        this.inFlight.add(msg.id!);
        this.notifyInFlight();
        try {
          const tags = await classifyEmailTagsYesNo(
            s.anthropicApiKey, s.claudeModel,
            s.tagClassifyPrompt || DEFAULT_TAG_PROMPT,
            content, toAsk,
            s.tagContradictions || {},
            s.tagPrecludes || {},
          );

          const currentNames = new Set(this.store.getTags(msg.id!).map((e) => e.tag));
          const fresh = tags.filter((t) => !currentNames.has(t));
          if (fresh.length > 0) {
            for (const tag of fresh) {
              this.store.setTag(msg.id!, tag, "auto", getTagVersion(s.tagPromptVersions, tag));
            }
            onProgress?.();
          }
        } catch (err) {
          logger.warn("Classifier", `Tag classification failed for ${msg.id}`, err);
        } finally {
          this.inFlight.delete(msg.id!);
          this.notifyInFlight();
        }
      }),
    );

    await Promise.all(promises);
  }

  /** Remove all auto-assigned tags, keeping manual ones. */
  clearAutoTags(): void {
    for (const [id, entries] of this.store.getAllTags()) {
      for (const e of entries) {
        if (e.source === "auto") this.store.removeTag(id, e.tag);
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
