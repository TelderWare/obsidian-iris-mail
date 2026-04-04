import { setIcon } from "obsidian";
import type { ConversationGroup, Message, SenderGroup } from "../../types";
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
  onConversationSelect: (conversation: ConversationGroup) => void;
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

/** Render the shared "All caught up" empty state. */
function renderEmptyState(parent: HTMLElement): void {
  const empty = parent.createDiv({ cls: "iris-empty-state" });
  const icon = empty.createDiv({ cls: "iris-empty-icon" });
  setIcon(icon, "inbox");
  empty.createDiv({ cls: "iris-empty-title", text: "All caught up" });
  empty.createDiv({ cls: "iris-empty-desc", text: "No messages to show right now." });
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
  private selectedConversationId: string | null = null;
  private selectedMessageId: string | null = null;
  private classifications = new Map<string, "important" | "routine" | "noise">();
  /** Row refs keyed by conversationId or messageId for surgical DOM updates. */
  private rowRefs = new Map<string, { el: HTMLElement; messages: Message[] }>();

  // Multi-selection state
  private selectedIds: Set<string> = new Set();
  private anchorId: string | null = null;
  private renderedOrder: string[] = [];

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

  setClassifications(map: Map<string, "important" | "routine" | "noise">): void {
    this.classifications = map;
  }

  /** Add or remove `!` indicators on existing rows without re-rendering. */
  updateImportanceIndicators(): void {
    for (const [, ref] of this.rowRefs) {
      const hasImportant = ref.messages.some(
        (m) => this.classifications.get(m.id || "") === "important",
      );
      const existing = ref.el.querySelector(".iris-msg-importance");
      if (hasImportant && !existing) {
        const span = document.createElement("span");
        span.className = "iris-msg-importance";
        setIcon(span, "circle-alert");
        // Always target badges container first; fall back to meta for hideSender rows
        const badges = ref.el.querySelector(".iris-msg-badges");
        if (badges) {
          badges.appendChild(span);
        } else {
          const meta = ref.el.querySelector(".iris-msg-meta");
          const date = meta?.querySelector(".iris-msg-date");
          if (date) {
            meta!.insertBefore(span, date);
          } else {
            ref.el.appendChild(span);
          }
        }
      } else if (!hasImportant && existing) {
        existing.remove();
      }
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

  /** Return a copy of the current selected IDs. */
  getSelectedIds(): Set<string> {
    return new Set(this.selectedIds);
  }

  /** Conversation list: homepage in conversation view mode. */
  renderConversations(
    conversations: ConversationGroup[],
    hasMore: boolean,
  ): void {
    this.containerEl.empty();
    this.selectedMessageId = null;
    this.rowRefs.clear();
    this.renderedOrder = [];
    this.selectedIds.clear();
    this.anchorId = null;
    this.renderedPageCount = 1;

    if (conversations.length === 0) {
      renderEmptyState(this.containerEl);
      return;
    }

    const listEl = this.containerEl.createDiv({
      cls: "iris-msg-list-inner iris-conv-list",
    });

    // Paginate: only render first PAGE_SIZE items
    const visibleConversations = conversations.slice(0, PAGE_SIZE * this.renderedPageCount);
    const hasMoreLocal = visibleConversations.length < conversations.length;

    for (const conv of visibleConversations) {
      const latest = conv.latestMessage;
      const isUnread = conv.unreadCount > 0;

      const row = listEl.createDiv({
        cls:
          "iris-msg-row" +
          (isUnread ? " is-unread" : "") +
          (conv.conversationId === this.selectedConversationId
            ? " is-selected"
            : ""),
      });

      // Conversation list row 1: [importance] [subject]
      const impSlot = row.createDiv({ cls: "iris-msg-slot-importance" });
      if (conv.messages.some((m) => this.classifications.get(m.id || "") === "important")) {
        const imp = impSlot.createSpan({ cls: "iris-msg-importance" });
        setIcon(imp, "circle-alert");
      }

      row.createDiv({
        cls: "iris-msg-subject",
        text: cleanSubject(conv.subject || "") || "(no subject)",
      });

      // Conversation list row 2: [count] [sender · date]
      const countSlot = row.createDiv({ cls: "iris-msg-slot-count" });
      if (conv.messages.length > 1) {
        countSlot.createSpan({
          cls: "iris-msg-count",
          text: String(conv.messages.length),
        });
      }

      const parts: string[] = [];
      // Collect unique effective (original) senders across the conversation
      const seenAddrs = new Set<string>();
      const senderNames: string[] = [];
      for (const m of conv.messages) {
        let addr: string;
        let name: string;
        if (this.resolveEffectiveSender) {
          const eff = this.resolveEffectiveSender(m);
          addr = (eff.address || "").toLowerCase();
          name = this.resolveName(eff.address, eff.name);
        } else {
          const envelope = getEnvelopeSender(m);
          addr = envelope.address.toLowerCase();
          name = this.resolveName(envelope.address, envelope.name);
        }
        if (addr && !seenAddrs.has(addr)) {
          seenAddrs.add(addr);
          if (name) senderNames.push(name);
        }
      }
      if (senderNames.length > 0) parts.push(senderNames.join(", "));
      if (latest.receivedDateTime) parts.push(formatRelativeDate(latest.receivedDateTime));

      row.createDiv({
        cls: "iris-msg-summary",
        text: parts.join(" · "),
      });

      this.rowRefs.set(conv.conversationId, { el: row, messages: conv.messages });
      this.renderedOrder.push(conv.conversationId);

      row.addEventListener("click", (e: MouseEvent) => {
        if (e.shiftKey && this.anchorId) {
          e.preventDefault();
          this.selectRange(conv.conversationId);
          this.syncSelectionClasses();
          this.callbacks.onMultiSelect(new Set(this.selectedIds));
        } else {
          this.selectedIds.clear();
          this.selectedIds.add(conv.conversationId);
          this.anchorId = conv.conversationId;
          this.selectedConversationId = conv.conversationId;
          this.callbacks.onConversationSelect(conv);
        }
      });
    }

    if (hasMoreLocal) {
      renderShowMoreButton(
        this.containerEl,
        conversations.length - visibleConversations.length,
        () => { this.renderedPageCount++; this.renderConversations(conversations, hasMore); },
      );
    }

    if (hasMore) {
      renderLoadMoreButton(this.containerEl, () => this.callbacks.onLoadMore());
    }
  }

  /**
   * Conversation drilldown (hideSender=false): messages within a conversation thread.
   * Sender drilldown (hideSender=true): messages from a single sender.
   */
  renderConversationMessages(
    subject: string,
    messages: Message[],
    hideSender = false,
  ): void {
    this.containerEl.empty();
    this.rowRefs.clear();
    this.renderedOrder = [];

    // Back header
    const header = this.containerEl.createDiv({
      cls: "iris-conv-header",
    });
    const backBtn = header.createEl("button", {
      cls: "iris-conv-back clickable-icon",
      attr: { "aria-label": "Back to conversations" },
    });
    setIcon(backBtn, "arrow-left");
    backBtn.addEventListener("click", () => this.callbacks.onBack());

    header.createSpan({
      cls: "iris-conv-title",
      text: subject || "(no subject)",
    });

    // Message rows
    const listEl = this.containerEl.createDiv({
      cls: "iris-msg-list-inner" + (hideSender ? " iris-sender-drilldown" : " iris-conv-drilldown"),
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

      // Row 1: [importance slot] [sender name or subject]
      const impSlot = row.createDiv({ cls: "iris-msg-slot-importance" });
      if (this.classifications.get(msgId) === "important") {
        const imp = impSlot.createSpan({ cls: "iris-msg-importance" });
        setIcon(imp, "circle-alert");
      }

      if (hideSender) {
        // Sender drilldown: row 1 is subject
        row.createDiv({
          cls: "iris-msg-subject",
          text: cleanSubject(msg.subject || "") || "(no subject)",
        });
      } else {
        // Conversation drilldown: row 1 is sender name
        const nameEl = row.createDiv({ cls: "iris-msg-sender-name" });
        this.renderSenderName(nameEl, msg);
      }

      // Row 2: [attachment slot] [bottom text]
      const clipSlot = row.createDiv({ cls: "iris-msg-slot-attachment" });
      if (msg.hasAttachments) {
        const clip = clipSlot.createSpan({ cls: "iris-msg-attachment" });
        setIcon(clip, "paperclip");
      }

      if (hideSender) {
        // Sender drilldown: row 2 is date
        row.createDiv({
          cls: "iris-msg-date",
          text: msg.receivedDateTime
            ? formatRelativeDate(msg.receivedDateTime)
            : "",
        });
      } else {
        // Conversation drilldown: row 2 is subject · date
        const parts: string[] = [];
        const subj = cleanSubject(msg.subject || "") || "(no subject)";
        parts.push(subj);
        if (msg.receivedDateTime) parts.push(formatRelativeDate(msg.receivedDateTime));

        row.createDiv({
          cls: "iris-msg-summary",
          text: parts.join(" · "),
        });
      }

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
          this.renderConversationMessages(subject, messages, hideSender);
          this.callbacks.onMessageSelect(msg);
        }
      });
    }
  }

  /** Sender list: homepage in sender view mode. */
  renderSenders(
    senders: SenderGroup[],
    hasMore: boolean,
    msgFilter?: (m: Message) => boolean,
  ): void {
    this.containerEl.empty();
    this.selectedMessageId = null;
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

      // Row 1: [importance slot] [sender name]
      const impSlot = row.createDiv({ cls: "iris-msg-slot-importance" });
      if (visibleMessages.some((m) => this.classifications.get(m.id || "") === "important")) {
        const imp = impSlot.createSpan({ cls: "iris-msg-importance" });
        setIcon(imp, "circle-alert");
      }

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

      row.addEventListener("click", () => {
        this.callbacks.onSenderSelect(sender);
      });
    }

    if (hasMoreLocal) {
      renderShowMoreButton(
        this.containerEl,
        senders.length - visibleSenders.length,
        () => { this.renderedPageCount++; this.renderSenders(senders, hasMore, msgFilter); },
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
    empty.createDiv({ cls: "iris-empty-title", text: "Session expired" });
    empty.createDiv({
      cls: "iris-empty-desc",
      text: "Sign in to view your messages.",
    });

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
