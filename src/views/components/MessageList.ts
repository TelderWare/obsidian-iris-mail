import { setIcon } from "obsidian";
import type { Message, SenderGroup } from "../../types";
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
  onSenderSelect: (sender: SenderGroup) => void;
  onBack: () => void;
  onLoadMore: () => void;
  onMultiSelect: (selectedIds: Set<string>) => void;
  onEditNickname: (address: string, rawName: string) => void;
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

/** True when the user has opted into reduced motion at the OS level. */
function prefersReducedMotion(): boolean {
  return typeof window !== "undefined"
    && window.matchMedia?.("(prefers-reduced-motion: reduce)").matches === true;
}

/** Render the shared "All caught up" empty state. */
function renderEmptyState(parent: HTMLElement): void {
  const empty = parent.createDiv({ cls: "iris-empty-state" });
  const icon = empty.createDiv({ cls: "iris-empty-icon" });
  setIcon(icon, "inbox");
  empty.createDiv({ cls: "iris-empty-title", text: "No messages" });
}

/** Render a "Show more" local pagination button. */
function renderShowMoreButton(
  parent: HTMLElement,
  remainingCount: number,
  onClick: () => void,
): void {
  const el = parent.createDiv({ cls: "iris-load-more" });
  const btn = el.createEl("button", {
    cls: "iris-header-icon clickable-icon",
    attr: { "aria-label": `Show more (${remainingCount} remaining)` },
  });
  setIcon(btn, "chevrons-down");
  btn.addEventListener("click", onClick);
}

/** Render a "Load more from server" button. */
function renderLoadMoreButton(
  parent: HTMLElement,
  onClick: () => void,
): void {
  const el = parent.createDiv({ cls: "iris-load-more" });
  const btn = el.createEl("button", {
    cls: "iris-header-icon clickable-icon",
    attr: { "aria-label": "Load more from server" },
  });
  setIcon(btn, "chevrons-down");
  btn.addEventListener("click", onClick);
}

export class MessageList {
  private containerEl: HTMLElement;
  private callbacks: MessageListCallbacks;
  private resolveName: NameResolver;
  private resolveEffectiveSender: EffectiveSenderResolver | null;
  private selectedMessageId: string | null = null;
  /** Row refs keyed by messageId (or senderGroupKey for sender list) for surgical DOM updates. */
  private rowRefs = new Map<string, { el: HTMLElement; messages: Message[] }>();

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
        this.callbacks.onEditNickname(address, rawName);
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

  /**
   * FLIP-style transition: snapshot current row positions, run the rebuild,
   * then (in parallel) collapse ghost clones of disappeared rows and translate
   * surviving rows from their old positions to their new ones. Everything
   * animates synchronously over a single 220ms window.
   */
  private async transitionTo(
    keepIds: Set<string>,
    rebuild: () => void,
  ): Promise<void> {
    if (prefersReducedMotion()) {
      rebuild();
      return;
    }
    const DURATION = 220;

    // Snapshot old positions and disappearing row visuals before the rebuild.
    const oldRects = new Map<string, DOMRect>();
    const ghosts: Array<{ ghost: HTMLElement; rect: DOMRect; height: number }> = [];
    for (const [id, ref] of this.rowRefs) {
      const rect = ref.el.getBoundingClientRect();
      oldRects.set(id, rect);
      if (!keepIds.has(id)) {
        const ghost = ref.el.cloneNode(true) as HTMLElement;
        ghosts.push({ ghost, rect, height: rect.height });
      }
    }

    if (oldRects.size === 0) {
      rebuild();
      return;
    }

    rebuild();

    const container = this.containerEl;
    const containerStyle = window.getComputedStyle(container);
    if (containerStyle.position === "static") {
      container.style.position = "relative";
    }
    const containerRect = container.getBoundingClientRect();

    // Overlay ghosts at their pre-rebuild positions and collapse them.
    const animations: Promise<void>[] = [];
    for (const { ghost, rect, height } of ghosts) {
      ghost.removeClass("is-selected");
      ghost.style.position = "absolute";
      ghost.style.top = `${rect.top - containerRect.top + container.scrollTop}px`;
      ghost.style.left = `${rect.left - containerRect.left}px`;
      ghost.style.width = `${rect.width}px`;
      ghost.style.height = `${height}px`;
      ghost.style.margin = "0";
      ghost.style.pointerEvents = "none";
      ghost.style.zIndex = "1";
      container.appendChild(ghost);

      void ghost.offsetHeight;

      ghost.addClass("is-collapsing");
      requestAnimationFrame(() => {
        ghost.style.height = "0px";
      });

      animations.push(new Promise((resolve) => {
        const cleanup = () => { ghost.remove(); resolve(); };
        setTimeout(cleanup, DURATION + 40);
      }));
    }

    // FLIP the surviving rows from their old positions to their new ones.
    for (const [id, ref] of this.rowRefs) {
      const oldRect = oldRects.get(id);
      if (!oldRect) continue;
      const newRect = ref.el.getBoundingClientRect();
      const dy = oldRect.top - newRect.top;
      if (Math.abs(dy) < 0.5) continue;

      const el = ref.el;
      el.style.transition = "none";
      el.style.transform = `translateY(${dy}px)`;
      void el.offsetHeight;
      el.style.transition = `transform ${DURATION}ms ease`;
      requestAnimationFrame(() => {
        el.style.transform = "";
      });

      animations.push(new Promise((resolve) => {
        setTimeout(() => {
          el.style.transition = "";
          el.style.transform = "";
          resolve();
        }, DURATION + 40);
      }));
    }

    await Promise.all(animations);
  }

  /** Flat message list: top-level messages view (no grouping, no drill-down). */
  async renderFlatMessages(
    messages: Message[],
    hasMore: boolean,
  ): Promise<void> {
    const keepIds = new Set(messages.map((m) => m.id || ""));
    const token = ++this.renderToken;
    await this.transitionTo(keepIds, () => {
      if (token !== this.renderToken) return;
      this.rebuildFlatMessages(messages, hasMore);
    });
  }

  private rebuildFlatMessages(messages: Message[], hasMore: boolean): void {
    this.containerEl.empty();
    this.rowRefs.clear();
    this.renderedOrder = [];
    this.selectedIds.clear();
    this.anchorId = null;
    this.renderedPageCount = 1;

    if (messages.length === 0) {
      renderEmptyState(this.containerEl);
      return;
    }

    const listEl = this.containerEl.createDiv({
      cls: "iris-msg-list-inner iris-conv-drilldown",
    });

    const visibleMessages = messages.slice(0, PAGE_SIZE * this.renderedPageCount);
    const hasMoreLocal = visibleMessages.length < messages.length;

    for (const msg of visibleMessages) {
      const msgId = msg.id || "";

      const row = listEl.createDiv({
        cls:
          "iris-msg-row" +
          (!msg.isRead ? " is-unread" : "") +
          (this.selectedIds.has(msgId) || msg.id === this.selectedMessageId
            ? " is-selected"
            : ""),
      });

      // Row 1: sender name
      const nameEl = row.createDiv({ cls: "iris-msg-sender-name" });
      this.renderSenderName(nameEl, msg);

      // Row 2: [attachment slot] [subject · date]
      const clipSlot = row.createDiv({ cls: "iris-msg-slot-attachment" });
      if (msg.hasAttachments) {
        const clip = clipSlot.createSpan({ cls: "iris-msg-attachment" });
        setIcon(clip, "paperclip");
      }

      const parts: string[] = [];
      const subj = cleanSubject(msg.subject || "") || "(no subject)";
      parts.push(subj);
      if (msg.receivedDateTime) parts.push(formatRelativeDate(msg.receivedDateTime));

      row.createDiv({
        cls: "iris-msg-summary",
        text: parts.join(" · "),
      });

      this.rowRefs.set(msgId, { el: row, messages: [msg] });
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
          void this.renderFlatMessages(messages, hasMore);
          this.callbacks.onMessageSelect(msg);
        }
      });
    }

    if (hasMoreLocal) {
      renderShowMoreButton(
        this.containerEl,
        messages.length - visibleMessages.length,
        () => { this.renderedPageCount++; void this.renderFlatMessages(messages, hasMore); },
      );
    }

    if (hasMore) {
      renderLoadMoreButton(this.containerEl, () => this.callbacks.onLoadMore());
    }
  }

  /** Sender drilldown: messages from a single sender, with back header. */
  async renderSenderMessages(
    senderName: string,
    messages: Message[],
  ): Promise<void> {
    const keepIds = new Set(messages.map((m) => m.id || ""));
    const token = ++this.renderToken;
    await this.transitionTo(keepIds, () => {
      if (token !== this.renderToken) return;
      this.rebuildSenderMessages(senderName, messages);
    });
  }

  private rebuildSenderMessages(senderName: string, messages: Message[]): void {
    this.containerEl.empty();
    this.rowRefs.clear();
    this.renderedOrder = [];

    // Back header
    const header = this.containerEl.createDiv({
      cls: "iris-conv-header",
    });
    const backBtn = header.createEl("button", {
      cls: "iris-conv-back clickable-icon",
      attr: { "aria-label": "Back" },
    });
    setIcon(backBtn, "arrow-left");
    backBtn.addEventListener("click", () => this.callbacks.onBack());

    header.createSpan({
      cls: "iris-conv-title",
      text: senderName || "(unknown sender)",
    });

    const listEl = this.containerEl.createDiv({
      cls: "iris-msg-list-inner iris-sender-drilldown",
    });

    for (const msg of messages) {
      const msgId = msg.id || "";

      const row = listEl.createDiv({
        cls:
          "iris-msg-row" +
          (!msg.isRead ? " is-unread" : "") +
          (this.selectedIds.has(msgId) || msg.id === this.selectedMessageId
            ? " is-selected"
            : ""),
      });

      row.createDiv({
        cls: "iris-msg-subject",
        text: cleanSubject(msg.subject || "") || "(no subject)",
      });

      const clipSlot = row.createDiv({ cls: "iris-msg-slot-attachment" });
      if (msg.hasAttachments) {
        const clip = clipSlot.createSpan({ cls: "iris-msg-attachment" });
        setIcon(clip, "paperclip");
      }

      row.createDiv({
        cls: "iris-msg-date",
        text: msg.receivedDateTime
          ? formatRelativeDate(msg.receivedDateTime)
          : "",
      });

      this.rowRefs.set(msgId, { el: row, messages: [msg] });
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
          void this.renderSenderMessages(senderName, messages);
          this.callbacks.onMessageSelect(msg);
        }
      });
    }
  }

  /** Sender list: homepage in sender view mode. */
  async renderSenders(
    senders: SenderGroup[],
    hasMore: boolean,
    msgFilter?: (m: Message) => boolean,
  ): Promise<void> {
    const keepIds = new Set(senders.map((s) => s.groupKey));
    const token = ++this.renderToken;
    await this.transitionTo(keepIds, () => {
      if (token !== this.renderToken) return;
      this.rebuildSenders(senders, hasMore, msgFilter);
    });
  }

  private rebuildSenders(
    senders: SenderGroup[],
    hasMore: boolean,
    msgFilter?: (m: Message) => boolean,
  ): void {
    this.containerEl.empty();
    this.selectedMessageId = null;
    this.rowRefs.clear();
    this.selectedIds.clear();
    this.anchorId = null;
    this.renderedOrder = [];
    this.renderedPageCount = 1;

    if (senders.length === 0) {
      renderEmptyState(this.containerEl);
      return;
    }

    const listEl = this.containerEl.createDiv({
      cls: "iris-msg-list-inner iris-sender-list",
    });

    const visibleSenders = senders.slice(0, PAGE_SIZE * this.renderedPageCount);
    const hasMoreLocal = visibleSenders.length < senders.length;

    for (const sender of visibleSenders) {
      const visibleMessages = msgFilter
        ? sender.messages.filter(msgFilter)
        : sender.messages;
      const isUnread = sender.unreadCount > 0;
      const latest = sender.latestMessage;

      const row = listEl.createDiv({
        cls: "iris-msg-row" + (isUnread ? " is-unread" : ""),
      });

      const nameEl = row.createDiv({ cls: "iris-msg-sender-name" });
      this.createAddressSpan(nameEl, sender.address, sender.name || sender.address);

      // Row 2: [count slot] [date]
      const countN = visibleMessages.length;
      const countSlot = row.createDiv({ cls: "iris-msg-slot-count" });
      if (countN > 1) {
        countSlot.createSpan({
          cls: "iris-msg-count",
          text: String(countN),
        });
      }

      const dateStr = latest.receivedDateTime
        ? formatRelativeDate(latest.receivedDateTime)
        : "";
      row.createDiv({
        cls: "iris-msg-summary",
        text: dateStr,
      });

      this.rowRefs.set(sender.groupKey, { el: row, messages: sender.messages });
      this.renderedOrder.push(sender.groupKey);

      row.addEventListener("click", () => {
        this.callbacks.onSenderSelect(sender);
      });
    }

    if (hasMoreLocal) {
      renderShowMoreButton(
        this.containerEl,
        senders.length - visibleSenders.length,
        () => { this.renderedPageCount++; void this.renderSenders(senders, hasMore, msgFilter); },
      );
    }

    if (hasMore) {
      renderLoadMoreButton(this.containerEl, () => this.callbacks.onLoadMore());
    }
  }

  renderLoggedOut(
    onBrowserSignIn: () => void,
    onDeviceCodeSignIn: () => void,
  ): void {
    this.containerEl.empty();
    this.rowRefs.clear();
    this.renderedOrder = [];

    const empty = this.containerEl.createDiv({ cls: "iris-empty-state" });
    const icon = empty.createDiv({ cls: "iris-empty-icon" });
    setIcon(icon, "log-out");
    empty.createDiv({ cls: "iris-empty-title", text: "Signed out" });

    const btnGroup = empty.createDiv({ cls: "iris-sign-in-buttons" });

    const browserBtn = btnGroup.createEl("button", {
      text: "Sign in with browser",
      cls: "mod-cta",
    });
    browserBtn.addEventListener("click", onBrowserSignIn);

    const deviceBtn = btnGroup.createEl("button", {
      text: "Sign in with device code",
    });
    deviceBtn.addEventListener("click", onDeviceCodeSignIn);
  }

  showLoading(): void {
    // Remove any previous overlay, but keep existing messages visible
    this.containerEl.querySelector(".iris-loading-overlay")?.remove();
    const overlay = this.containerEl.createDiv({ cls: "iris-loading-overlay" });
    overlay.createDiv({ cls: "iris-loading-spinner iris-loading-spinner--small" });
  }
}
