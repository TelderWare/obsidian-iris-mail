import { Menu, setIcon } from "obsidian";
import type { Message } from "../../types";
import type { TagCacheEntry } from "../../store/types";
import { formatRelativeDate } from "../../utils/dateFormat";
import { getEnvelopeSender } from "../../utils/envelopeSender";

/** Strip forwarding prefixes (FW: / Fwd:) from subject lines. */
function cleanSubject(raw: string): string {
  return raw.replace(/^(?:fw|fwd)\s*:\s*/i, "").trim();
}

/** Strip leading underscore/dash separator lines and whitespace from body previews. */
function cleanPreview(raw: string): string {
  return raw.replace(/^[\s_\-=*]+/, "").trim();
}

/**
 * Strip forwarded-email header block from a body preview.
 * Previews from forwarded messages often begin with a run of
 * "From: … Sent: … To: … Subject: …" tokens.  Remove them so
 * only the real body text remains.
 */
function stripForwardedHeaders(preview: string): string {
  const headerRe = /^(from|to|sent|date|subject|cc|bcc)\s*:/i;
  const words = preview.split(/\s+/);
  let i = 0;
  while (i < words.length) {
    // Check if the current position starts a header label
    const rest = words.slice(i).join(" ");
    if (headerRe.test(rest)) {
      // Skip past this header's value until the next header keyword or end
      i++; // skip the label itself
      while (i < words.length) {
        const upcoming = words.slice(i).join(" ");
        if (headerRe.test(upcoming)) break;
        i++;
      }
    } else {
      break;
    }
  }
  return words.slice(i).join(" ").trim();
}

/**
 * Replace `Name <email@domain>` patterns in preview text with resolved
 * nicknames and strip bare `<email@domain>` occurrences.
 */
function resolvePreviewNames(
  preview: string,
  resolveName: NameResolver,
): string {
  // Replace "Display Name <email@domain>" with resolved name
  let result = preview.replace(
    /([^<,;]+?)\s*<([^@>]+@[^>]+)>/g,
    (_match, rawName: string, addr: string) => {
      return resolveName(addr.trim(), rawName.trim());
    },
  );
  // Strip any remaining bare <email@domain> patterns
  result = result.replace(/<[^@>]+@[^>]+>/g, "");
  return result.replace(/\s{2,}/g, " ").trim();
}

interface MessageListCallbacks {
  onMessageSelect: (msg: Message) => void;
  onLoadMore: () => void;
  onMultiSelect: (selectedIds: Set<string>) => void;
  onEditNickname: (address: string, rawName: string) => void;
  /** Open the sender-rule editor (auto-bin, auto-tag, ...) for an address. */
  onEditSenderRule: (address: string, rawName: string) => void;
  /** Add a sender rule that auto-bins all future messages from this address. */
  onMarkSenderAsJunk: (address: string, rawName: string) => void;
}

export type NameResolver = (address: string, rawName: string) => string;

export interface EffectiveSender {
  address: string;
  name: string;
  viaAddress?: string;
  viaName?: string;
}

export type EffectiveSenderResolver = (msg: Message) => EffectiveSender;

/** Maximum items to render per page before showing "Show more". */
const PAGE_SIZE = 100;

/** Render the shared "All caught up" empty state. */
function renderEmptyState(parent: HTMLElement, iconName = "inbox", title = "No messages"): void {
  const empty = parent.createDiv({ cls: "iris-empty-state" });
  const icon = empty.createDiv({ cls: "iris-empty-icon" });
  setIcon(icon, iconName);
  empty.createDiv({ cls: "iris-empty-title", text: title });
}


export class MessageList {
  private containerEl: HTMLElement;
  private callbacks: MessageListCallbacks;
  private resolveName: NameResolver;
  private resolveEffectiveSender: EffectiveSenderResolver | null;
  private selectedMessageId: string | null = null;
  /** Row refs keyed by messageId (or senderGroupKey for sender list) for surgical DOM updates. */
  private rowRefs = new Map<string, { el: HTMLElement; messages: Message[]; tagSlot?: HTMLElement }>();
  private tagCache = new Map<string, TagCacheEntry[]>();
  private tagIcons = new Map<string, string>();
  private tagColors = new Map<string, string>();
  private tagDescriptions = new Map<string, string>();
  private hiddenListTags = new Set<string>();
  private pinnedIds = new Set<string>();

  // Multi-selection state
  private selectedIds: Set<string> = new Set();
  private anchorId: string | null = null;
  private renderedOrder: string[] = [];
  /** Incremented on every render — awaits check this to drop superseded renders. */
  private renderToken = 0;

  // Pagination state for large lists
  private renderedPageCount = 1;

  constructor(
    containerEl: HTMLElement,
    callbacks: MessageListCallbacks,
    resolveName: NameResolver = (_a, r) => r,
    resolveEffectiveSender: EffectiveSenderResolver | null = null,
  ) {
    this.containerEl = containerEl;
    this.callbacks = callbacks;
    this.resolveName = resolveName;
    this.resolveEffectiveSender = resolveEffectiveSender;
  }

  setEffectiveSenderResolver(resolver: EffectiveSenderResolver | null): void {
    this.resolveEffectiveSender = resolver;
  }

  setTagCache(cache: Map<string, TagCacheEntry[]>): void {
    this.tagCache = cache;
  }

  setTagIcons(icons: Map<string, string>): void {
    this.tagIcons = icons;
  }

  setTagColors(colors: Map<string, string>): void {
    this.tagColors = colors;
  }

  setTagDescriptions(descriptions: Map<string, string>): void {
    this.tagDescriptions = descriptions;
  }

  setHiddenListTags(hidden: Set<string>): void {
    this.hiddenListTags = hidden;
  }

  setPinnedIds(ids: Set<string>): void {
    this.pinnedIds = ids;
  }

  /** Fill the provided top-left slot with a tag badge for the message (if any non-hidden tags). */
  private renderTagSlot(slot: HTMLElement, msg: Message): void {
    slot.empty();
    if (!msg.id) return;
    const entries = this.tagCache.get(msg.id);
    if (!entries || entries.length === 0) return;
    const visible = entries.filter((e) => !this.hiddenListTags.has(e.tag));
    if (visible.length === 0) return;
    const first = visible[0];
    const iconName = this.tagIcons.get(first.tag) || "tag";
    const tooltip = visible
      .map((e) => {
        const desc = this.tagDescriptions.get(e.tag);
        return desc ? `${e.tag} — ${desc}` : e.tag;
      })
      .join("\n");
    const badge = slot.createSpan({
      cls: "iris-msg-tag-badge",
      attr: { title: tooltip },
    });
    const color = this.tagColors.get(first.tag);
    if (color) badge.style.color = color;
    setIcon(badge, iconName);
  }

  /** Update tag badges on existing rows without rebuilding the list. */
  refreshTagBadges(): void {
    for (const ref of this.rowRefs.values()) {
      if (!ref.tagSlot) continue;
      const msg = ref.messages[0];
      if (!msg) continue;
      this.renderTagSlot(ref.tagSlot, msg);
    }
  }

  /** Create a name span with a right-click context menu to edit the nickname. */
  private createAddressSpan(
    parent: HTMLElement,
    address: string,
    rawName: string,
    cls?: string,
  ): HTMLSpanElement {
    const span = parent.createSpan({
      text: this.resolveName(address, rawName),
      cls,
      attr: address ? { title: address } : {},
    });
    if (address) {
      span.addEventListener("contextmenu", (evt) => {
        evt.preventDefault();
        evt.stopPropagation();

        const menu = new Menu();
        menu.addItem((item) =>
          item
            .setTitle("Edit nickname…")
            .setIcon("pencil")
            .onClick(() => this.callbacks.onEditNickname(address, rawName)),
        );
        menu.addItem((item) =>
          item
            .setTitle("Create rule…")
            .setIcon("filter")
            .onClick(() => this.callbacks.onEditSenderRule(address, rawName)),
        );
        menu.addItem((item) =>
          item
            .setTitle("Block")
            .setIcon("trash-2")
            .onClick(() => this.callbacks.onMarkSenderAsJunk(address, rawName)),
        );
        menu.showAtMouseEvent(evt);
      });
    }
    return span;
  }

  /** Render sender name into a container, appending a "via" label when applicable. */
  private renderSenderName(
    container: HTMLElement,
    msg: Message,
    fallbackAddr?: string,
    fallbackName?: string,
  ): void {
    if (this.resolveEffectiveSender) {
      const eff = this.resolveEffectiveSender(msg);
      this.createAddressSpan(container, eff.address, eff.name);
      if (eff.viaName) {
        const viaSpan = this.createAddressSpan(container, eff.viaAddress || "", eff.viaName, "iris-msg-via");
        viaSpan.textContent = ` via ${viaSpan.textContent}`;
      }
    } else {
      const envelope = getEnvelopeSender(msg);
      const addr = fallbackAddr ?? envelope.address;
      const raw = fallbackName ?? envelope.name;
      this.createAddressSpan(container, addr, raw);
    }
  }

  // --- Multi-selection helpers ---

  /** Compute range between anchor and target, populate selectedIds. */
  private selectRange(targetId: string): void {
    if (!this.anchorId || this.renderedOrder.length === 0) {
      this.selectedIds.clear();
      this.selectedIds.add(targetId);
      this.anchorId = targetId;
      return;
    }

    const anchorIdx = this.renderedOrder.indexOf(this.anchorId);
    const targetIdx = this.renderedOrder.indexOf(targetId);

    if (anchorIdx === -1 || targetIdx === -1) {
      this.selectedIds.clear();
      this.selectedIds.add(targetId);
      this.anchorId = targetId;
      return;
    }

    const start = Math.min(anchorIdx, targetIdx);
    const end = Math.max(anchorIdx, targetIdx);

    this.selectedIds.clear();
    for (let i = start; i <= end; i++) {
      this.selectedIds.add(this.renderedOrder[i]);
    }
    // Anchor stays the same on shift-click
  }

  /** Toggle is-selected class on existing DOM rows without re-rendering. */
  private syncSelectionClasses(): void {
    for (const [id, ref] of this.rowRefs) {
      const isSelected = this.selectedIds.has(id);
      ref.el.toggleClass("is-selected", isSelected);
    }
  }

  /** Clear multi-selection state and update DOM. */
  clearMultiSelection(): void {
    this.selectedIds.clear();
    this.anchorId = null;
    this.syncSelectionClasses();
  }

  /** Select every currently rendered row. Returns the resulting selection. */
  selectAll(): Set<string> {
    this.selectedIds.clear();
    for (const id of this.renderedOrder) {
      if (id) this.selectedIds.add(id);
    }
    if (this.renderedOrder.length > 0) {
      this.anchorId = this.renderedOrder[0];
    }
    this.syncSelectionClasses();
    return new Set(this.selectedIds);
  }

  /**
   * Update the highlighted single-selection row (used when the viewer is
   * driven programmatically, e.g. auto-advance to next unread) so the matching
   * list row shows the `is-selected` class.
   */
  setSelectedMessageId(id: string | null): void {
    this.selectedMessageId = id;
    this.selectedIds.clear();
    if (id) {
      this.selectedIds.add(id);
      this.anchorId = id;
    }
    this.syncSelectionClasses();
  }

  /** Return a copy of the current selected IDs. */
  getSelectedIds(): Set<string> {
    return new Set(this.selectedIds);
  }

  /** Return the IDs of every rendered row, in visual order. */
  getRenderedOrder(): readonly string[] {
    return this.renderedOrder;
  }

  private async transitionTo(
    _keepIds: Set<string>,
    rebuild: () => void,
  ): Promise<void> {
    rebuild();
  }

  /** Flat message list: top-level messages view (no grouping, no drill-down). */
  async renderFlatMessages(
    messages: Message[],
    hasMore: boolean,
    empty?: { icon?: string; title?: string },
  ): Promise<void> {
    const keepIds = new Set(messages.map((m) => m.id || ""));
    const token = ++this.renderToken;
    await this.transitionTo(keepIds, () => {
      if (token !== this.renderToken) return;
      this.rebuildFlatMessages(messages, hasMore, empty);
    });
  }

  private rebuildFlatMessages(
    messages: Message[],
    hasMore: boolean,
    empty?: { icon?: string; title?: string },
  ): void {
    this.containerEl.empty();
    this.rowRefs.clear();
    this.renderedOrder = [];
    this.selectedIds.clear();
    this.anchorId = null;
    this.renderedPageCount = 1;

    if (messages.length === 0) {
      renderEmptyState(this.containerEl, empty?.icon, empty?.title);
      return;
    }

    const listEl = this.containerEl.createDiv({
      cls: "iris-msg-list-inner iris-conv-drilldown",
    });

    const visibleMessages = messages.slice(0, PAGE_SIZE * this.renderedPageCount);
    const hasMoreLocal = visibleMessages.length < messages.length;

    for (const msg of visibleMessages) {
      const msgId = msg.id || "";

      const isPinned = msg.id ? this.pinnedIds.has(msg.id) : false;
      const row = listEl.createDiv({
        cls:
          "iris-msg-row" +
          (!msg.isRead ? " is-unread" : "") +
          (isPinned ? " is-pinned" : "") +
          (this.selectedIds.has(msgId) || msg.id === this.selectedMessageId
            ? " is-selected"
            : ""),
      });

      // Top-left slot: tag badge (when message has tags)
      const tagSlot = row.createDiv({ cls: "iris-msg-slot-tag" });
      this.renderTagSlot(tagSlot, msg);

      // Row 1: sender name — prefixed with a pin icon when pinned.
      const nameEl = row.createDiv({ cls: "iris-msg-sender-name" });
      if (isPinned) {
        const pinEl = nameEl.createSpan({
          cls: "iris-msg-pin-indicator",
          attr: { title: "Pinned" },
        });
        setIcon(pinEl, "pin");
      }
      this.renderSenderName(nameEl, msg);

      // Row 2: [attachment slot] [subject · date (on account)]
      const clipSlot = row.createDiv({ cls: "iris-msg-slot-attachment" });
      if (msg.hasAttachments) {
        const clip = clipSlot.createSpan({ cls: "iris-msg-attachment" });
        setIcon(clip, "paperclip");
      }

      const parts: string[] = [];
      const subj = cleanSubject(msg.subject || "") || "(no subject)";
      parts.push(subj);
      if (msg.receivedDateTime) {
        const date = formatRelativeDate(msg.receivedDateTime);
        parts.push(msg._accountLabel ? `${date} on ${msg._accountLabel}` : date);
      } else if (msg._accountLabel) {
        parts.push(`on ${msg._accountLabel}`);
      }

      row.createDiv({
        cls: "iris-msg-summary",
        text: parts.join(" · "),
      });

      this.rowRefs.set(msgId, { el: row, messages: [msg], tagSlot });
      this.renderedOrder.push(msgId);

      row.addEventListener("click", (e: MouseEvent) => {
        if (e.shiftKey && this.anchorId) {
          e.preventDefault();
          this.selectRange(msgId);
          this.syncSelectionClasses();
          this.callbacks.onMultiSelect(new Set(this.selectedIds));
        } else {
          this.selectedIds.clear();
          this.selectedIds.add(msgId);
          this.anchorId = msgId;
          this.selectedMessageId = msg.id || null;
          void this.renderFlatMessages(messages, hasMore, empty);
          this.callbacks.onMessageSelect(msg);
        }
      });
    }

    if (hasMoreLocal || hasMore) {
      this.attachLoadMoreSentinel(messages, hasMore, empty, hasMoreLocal);
    }
  }

  /**
   * Infinite-scroll sentinel: when it scrolls into view, reveal the next
   * local page (cheap — already in memory) or fetch the next page from the
   * server. The observer is disconnected as soon as it fires so the same
   * sentinel can't double-trigger; the next render creates a fresh one.
   */
  private loadMoreObserver: IntersectionObserver | null = null;

  private attachLoadMoreSentinel(
    messages: Message[],
    hasMore: boolean,
    empty: { icon?: string; title?: string } | undefined,
    hasMoreLocal: boolean,
  ): void {
    const sentinel = this.containerEl.createDiv({ cls: "iris-load-more-sentinel" });
    sentinel.createDiv({ cls: "iris-loading-spinner iris-loading-spinner--small" });

    this.loadMoreObserver?.disconnect();
    const observer = new IntersectionObserver((entries) => {
      if (!entries.some((e) => e.isIntersecting)) return;
      observer.disconnect();
      this.loadMoreObserver = null;

      if (hasMoreLocal) {
        this.renderedPageCount++;
        void this.renderFlatMessages(messages, hasMore, empty);
      } else if (hasMore) {
        this.callbacks.onLoadMore();
      }
    }, { root: this.containerEl, rootMargin: "200px 0px" });
    this.loadMoreObserver = observer;
    observer.observe(sentinel);
  }

  renderLoggedOut(onOpenSettings: () => void): void {
    this.containerEl.empty();
    this.rowRefs.clear();
    this.renderedOrder = [];

    const empty = this.containerEl.createDiv({ cls: "iris-empty-state" });
    const icon = empty.createDiv({ cls: "iris-empty-icon" });
    setIcon(icon, "log-out");
    empty.createDiv({ cls: "iris-empty-title", text: "No accounts signed in" });

    const btnGroup = empty.createDiv({ cls: "iris-sign-in-buttons" });
    const settingsBtn = btnGroup.createEl("button", {
      text: "Open Iris Mail settings",
      cls: "mod-cta",
    });
    settingsBtn.addEventListener("click", onOpenSettings);
  }

  showLoading(): void {
    // Remove any previous overlay, but keep existing messages visible
    this.containerEl.querySelector(".iris-loading-overlay")?.remove();
    const overlay = this.containerEl.createDiv({ cls: "iris-loading-overlay" });
    overlay.createDiv({ cls: "iris-loading-spinner iris-loading-spinner--small" });
  }
}
