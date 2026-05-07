import { App, MarkdownRenderer, Menu, sanitizeHTMLToDom, setIcon } from "obsidian";
import type { Message } from "../../types";
import type { TagCacheEntry } from "../../store/types";
import type { NameResolver, EffectiveSenderResolver } from "./MessageList";
import type { NoteType } from "../../utils/claudeApi";
import { formatRelativeDate } from "../../utils/dateFormat";

interface MessageViewerCallbacks {
  onMarkAsRead: (msg: Message) => void;
  onMarkAsUnread: (msg: Message) => void;
  /** Toggle the message's client-side to-do flag (moves it to/from the To-do box). */
  onToggleTodo: (msg: Message) => void;
  /** Toggle the message's pin state. Pinned messages are always loaded. */
  onTogglePin: (msg: Message) => void;
  onTagChange: (msg: Message, tag: string | null) => void;
  /** Re-tag a message with the current prompt (replaces obsolete auto-tags). */
  onRetagMessage: (msg: Message) => void;
  /** Batch mark selected messages as read. */
  onBatchMarkAsRead: (ids: Set<string>) => void;
  /** Batch mark selected messages as unread. */
  onBatchMarkAsUnread: (ids: Set<string>) => void;
  /** Move a single message to the provider's trash folder. */
  onDeleteMessage: (msg: Message) => void;
  /** Batch-delete selected messages. */
  onBatchDelete: (ids: Set<string>) => void;
  /** Batch assign a tag to selected messages. */
  onBatchTag: (ids: Set<string>, tag: string) => void;
  /** Bulk deny tag: remove tag from all selected, merge into formula, refine prompt. */
  onBulkDenyTag: (ids: Set<string>, tag: string) => void;
  /** Create an Event or Task note from selected text in the email body. */
  onCreateNoteFromSelection: (selectedText: string, noteType: NoteType, msg: Message) => void;
  /** Re-process the current message with the current prompt (replaces stale cached result). */
  onReprocessMessage: (msg: Message) => void;
  /** Open the nickname editor for an address. */
  onEditNickname: (address: string, rawName: string) => void;
  /** Add a sender rule that auto-bins all future messages from this address. */
  onMarkSenderAsJunk: (address: string, rawName: string) => void;
  /** Dismiss the viewer (used by compact back button). */
  onDismiss?: () => void;
}

export class MessageViewer {
  private containerEl: HTMLElement;
  private app: App;
  private callbacks: MessageViewerCallbacks;
  private resolveName: NameResolver;
  private currentMessageId: string | null = null;
  private currentHtml: string = "";
  private processedMarkdown: string | null = null;
  private showingProcessed = false;
  private currentMsg: Message | null = null;
  private currentTagPromptVersions: Record<string, number> = {};
  private tagCategories: string[] = [];
  private tagIcons = new Map<string, string>();
  private tagColors = new Map<string, string>();
  private tagDescriptions = new Map<string, string>();
  private tagCache = new Map<string, TagCacheEntry[]>();
  private todoIds = new Set<string>();
  private junkIds = new Set<string>();
  private pinnedIds = new Set<string>();
  private resolveEffectiveSender: EffectiveSenderResolver | null = null;
  private processedIsStale = false;

  constructor(
    containerEl: HTMLElement,
    app: App,
    callbacks: MessageViewerCallbacks,
    resolveName: NameResolver = (_a, r) => r,
  ) {
    this.containerEl = containerEl;
    this.app = app;
    this.callbacks = callbacks;
    this.resolveName = resolveName;
    this.clear();
  }

  setEffectiveSenderResolver(resolver: EffectiveSenderResolver | null): void {
    this.resolveEffectiveSender = resolver;
  }

  setTagCategories(categories: string[]): void {
    this.tagCategories = categories;
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

  setTagCache(cache: Map<string, TagCacheEntry[]>): void {
    this.tagCache = cache;
  }

  setTodoIds(ids: Set<string>): void {
    this.todoIds = ids;
    if (this.currentMsg) this.rebuildView();
  }

  setJunkIds(ids: Set<string>): void {
    this.junkIds = ids;
    if (this.currentMsg) this.rebuildView();
  }

  setPinnedIds(ids: Set<string>): void {
    this.pinnedIds = ids;
    if (this.currentMsg) this.rebuildView();
  }

  setPromptVersions(tagVersions: Record<string, number>): void {
    this.currentTagPromptVersions = tagVersions;
  }

  /** Render a single message with its stripped HTML body. */
  render(
    msg: Message,
    strippedHtml: string,
  ): void {
    const sameMessage = msg.id != null && msg.id === this.currentMessageId;

    this.currentMessageId = msg.id || null;
    this.currentMsg = msg;
    this.currentHtml = strippedHtml;
    if (!sameMessage) {
      this.processedMarkdown = null;
      this.showingProcessed = false;
    }

    this.rebuildView();
  }

  /** Re-render the current message without resetting toggle state. */
  refresh(): void {
    if (this.currentMsg) this.rebuildView();
  }

  /** Update with processed markdown. Guards against stale message.
   *  When `stale` is true, the result was cached with an older prompt hash. */
  showProcessedMarkdown(messageId: string, markdown: string, stale = false): void {
    if (messageId !== this.currentMessageId) return;
    this.processedMarkdown = markdown;
    this.processedIsStale = stale;
    this.showingProcessed = true;
    this.rebuildView();
  }

  showProcessingIndicator(): void {
    const existing = this.containerEl.querySelector(
      ".iris-viewer-processing",
    );
    if (existing) return;
    const indicator = createDiv({ cls: "iris-viewer-processing" });
    indicator.createDiv({ cls: "iris-loading-spinner iris-loading-spinner--small" });
    indicator.createSpan({ text: "Extracting key information\u2026" });
    // Insert after the header
    const header = this.containerEl.querySelector(".iris-viewer-header");
    if (header) {
      header.after(indicator);
    }
  }

  hideProcessingIndicator(): void {
    this.containerEl
      .querySelector(".iris-viewer-processing")
      ?.remove();
  }

  clear(): void {
    this.containerEl.empty();
    this.containerEl.style.display = "none";
    this.containerEl.parentElement?.removeClass("has-viewer");
    this.currentMessageId = null;
    this.currentMsg = null;
    this.currentHtml = "";
    this.processedMarkdown = null;
    this.showingProcessed = false;
  }

  /** Show a batch action panel when multiple messages are selected. */
  renderBatchPanel(selectedCount: number, selectedIds: Set<string>): void {
    this.containerEl.empty();
    this.containerEl.style.display = "";
    this.containerEl.parentElement?.addClass("has-viewer");
    this.currentMessageId = null;
    this.currentMsg = null;

    const panel = this.containerEl.createDiv({ cls: "iris-batch-panel" });

    panel.createEl("h3", {
      cls: "iris-batch-title",
      text: `${selectedCount} messages selected`,
    });

    const actions = panel.createDiv({ cls: "iris-batch-actions" });

    const readBtn = actions.createEl("button", {
      cls: "iris-header-icon clickable-icon",
      attr: { "aria-label": "Mark as read" },
    });
    setIcon(readBtn, "mail-open");
    readBtn.addEventListener("click", () => {
      this.callbacks.onBatchMarkAsRead(new Set(selectedIds));
    });

    const unreadBtn = actions.createEl("button", {
      cls: "iris-header-icon clickable-icon",
      attr: { "aria-label": "Mark as unread" },
    });
    setIcon(unreadBtn, "mail");
    unreadBtn.addEventListener("click", () => {
      this.callbacks.onBatchMarkAsUnread(new Set(selectedIds));
    });

    const batchDeleteBtn = actions.createEl("button", {
      cls: "iris-header-icon clickable-icon",
      attr: { "aria-label": "Move to bin" },
    });
    setIcon(batchDeleteBtn, "trash-2");
    batchDeleteBtn.addEventListener("click", () => {
      this.callbacks.onBatchDelete(new Set(selectedIds));
    });

    // Bulk deny tag: pick which tag is wrong on these emails
    if (this.tagCategories.length > 0) {
      const denyTagBtn = actions.createEl("button", {
        cls: "iris-header-icon clickable-icon",
        attr: { "aria-label": "Deny tag" },
      });
      setIcon(denyTagBtn, "tag");
      denyTagBtn.addEventListener("click", () => {
        this.openDropdown(actions, ".iris-tag-dropdown", "iris-tag-dropdown", (dropdown) => {
          dropdown.createDiv({ cls: "iris-tag-option-header", text: "Remove & deny tag:" });
          for (const cat of this.tagCategories) {
            const item = dropdown.createDiv({ cls: "iris-tag-option" });
            const catIcon = item.createSpan({ cls: "iris-tag-option-icon" });
            setIcon(catIcon, this.tagIcons.get(cat) || "tag");
            item.createSpan({ text: cat });
            item.addEventListener("click", () => {
              dropdown.remove();
              this.callbacks.onBulkDenyTag(new Set(selectedIds), cat);
            });
          }
        });
      });

      // Tag (manual assign — keep this)
      const tagBtn = actions.createEl("button", {
        cls: "iris-header-icon clickable-icon",
        attr: { "aria-label": "Tag" },
      });
      setIcon(tagBtn, "tag");
      tagBtn.addEventListener("click", () => {
        this.openDropdown(actions, ".iris-tag-dropdown:not(.iris-deny-dropdown)", "iris-tag-dropdown", (dropdown) => {
          for (const cat of this.tagCategories) {
            const item = dropdown.createDiv({ cls: "iris-tag-option" });
            const catIcon = item.createSpan({ cls: "iris-tag-option-icon" });
            setIcon(catIcon, this.tagIcons.get(cat) || "tag");
            item.createSpan({ text: cat });
            item.addEventListener("click", () => {
              dropdown.remove();
              this.callbacks.onBatchTag(new Set(selectedIds), cat);
            });
          }
        });
      });
    }
  }

  private rebuildView(): void {
    this.containerEl.empty();
    this.containerEl.style.display = "";
    this.containerEl.parentElement?.addClass("has-viewer");

    const msg = this.currentMsg;
    if (!msg) return;

    // Header
    const headerEl = this.containerEl.createDiv({
      cls: "iris-viewer-header",
    });

    // Compact-mode back button (visible only when list is hidden)
    if (this.callbacks.onDismiss) {
      const backBtn = headerEl.createEl("button", {
        cls: "iris-viewer-compact-back clickable-icon",
        attr: { "aria-label": "Back to list" },
      });
      setIcon(backBtn, "arrow-left");
      backBtn.addEventListener("click", () => this.callbacks.onDismiss!());
    }

    // Title line: subject only
    const titleLine = headerEl.createDiv({
      cls: "iris-viewer-title-line",
    });
    titleLine.createEl("h2", {
      cls: "iris-viewer-subject",
      text: (msg.subject || "").replace(/^(?:fw|fwd)\s*:\s*/i, "").trim() || "(no subject)",
    });

    // Tags line
    const tagsLine = headerEl.createDiv({ cls: "iris-viewer-tags-line" });

    if (this.tagCategories.length > 0 && msg.id) {
      const tagEntries = this.tagCache.get(msg.id) || [];
      const activeTags = new Set(tagEntries.map((e) => e.tag));

      if (tagEntries.length > 0) {
        const tagVer = (tag: string) => this.currentTagPromptVersions[tag] ?? 1;
        const hasObsoleteTag = tagEntries.some(
          (e) => e.source === "auto" && e.promptVersion != null
            && e.promptVersion < tagVer(e.tag),
        );

        for (const entry of tagEntries) {
          const isAutoTag = entry.source === "auto";
          const isObsoleteTag = isAutoTag && entry.promptVersion != null
            && entry.promptVersion < tagVer(entry.tag);
          const desc = this.tagDescriptions.get(entry.tag);
          const status = isObsoleteTag
            ? `Auto-tagged v${entry.promptVersion} (obsolete) — right-click to remove`
            : (isAutoTag ? "Auto-tagged" : "Manually tagged") + " — right-click to remove";
          const tagBadge = tagsLine.createSpan({
            cls: `iris-viewer-tag${isAutoTag ? " is-auto" : ""}${isObsoleteTag ? " is-obsolete" : ""}`,
            attr: {
              title: desc ? `${entry.tag} — ${desc}\n\n${status}` : status,
            },
          });
          const tagColor = this.tagColors.get(entry.tag);
          if (tagColor) {
            tagBadge.style.background = tagColor;
            tagBadge.style.color = "#fff";
          }
          const iconName = this.tagIcons.get(entry.tag) || "tag";
          const iconSpan = tagBadge.createSpan({ cls: "iris-viewer-tag-icon" });
          setIcon(iconSpan, iconName);
          tagBadge.createSpan({ cls: "iris-viewer-tag-label", text: entry.tag });
          if (isAutoTag) {
            const autoIcon = tagBadge.createSpan({ cls: "iris-viewer-tag-auto-icon" });
            setIcon(autoIcon, "sparkles");
          }
          tagBadge.addEventListener("contextmenu", (e) => {
            e.preventDefault();
            e.stopPropagation();
            const menu = new Menu();
            menu.addItem((item) =>
              item
                .setTitle(`Remove tag "${entry.tag}"`)
                .setIcon("x")
                .onClick(() => {
                  if (this.currentMsg) this.callbacks.onTagChange(this.currentMsg, entry.tag);
                }),
            );
            menu.showAtMouseEvent(e);
          });
        }

        if (hasObsoleteTag) {
          const retagBtn = tagsLine.createEl("button", {
            cls: "iris-viewer-tag-btn clickable-icon",
            attr: { "aria-label": "Re-tag with current prompt" },
          });
          setIcon(retagBtn, "refresh-cw");
          retagBtn.addEventListener("click", (e) => {
            e.stopPropagation();
            if (this.currentMsg) this.callbacks.onRetagMessage(this.currentMsg);
          });
        }
      }

      const tagBtn = tagsLine.createEl("button", {
        cls: "iris-viewer-tag-btn clickable-icon",
        attr: { "aria-label": "Assign tag" },
      });
      setIcon(tagBtn, "tag");
      tagBtn.addEventListener("click", (e) => {
        e.stopPropagation();
        this.showTagDropdown(tagsLine, msg.id!, activeTags);
      });
    }

    // Meta (from, to, cc, date) with styled labels
    const metaEl = headerEl.createDiv({ cls: "iris-viewer-meta" });
    const fromRow = metaEl.createDiv();
    fromRow.createSpan({ cls: "iris-viewer-meta-label", text: "From" });
    if (this.resolveEffectiveSender) {
      const eff = this.resolveEffectiveSender(msg);
      this.createAddressSpan(fromRow, eff.address, eff.name);
      if (eff.viaName) {
        const viaSpan = this.createAddressSpan(fromRow, eff.viaAddress || "", eff.viaName, "iris-msg-via");
        viaSpan.textContent = ` via ${viaSpan.textContent}`;
      }
    } else {
      const fromAddr = msg.from?.emailAddress;
      const fromAddress = fromAddr?.address || "";
      const fromRaw = fromAddr?.name || fromAddress || "";
      this.createAddressSpan(fromRow, fromAddress, fromRaw);
    }

    if (msg.toRecipients?.length) {
      const toRow = metaEl.createDiv();
      toRow.createSpan({ cls: "iris-viewer-meta-label", text: "To" });
      this.renderRecipients(toRow, msg.toRecipients);
    }

    if (msg.ccRecipients?.length) {
      const ccRow = metaEl.createDiv();
      ccRow.createSpan({ cls: "iris-viewer-meta-label", text: "Cc" });
      this.renderRecipients(ccRow, msg.ccRecipients);
    }

    if (msg.receivedDateTime) {
      const dateRow = metaEl.createDiv();
      dateRow.createSpan({ cls: "iris-viewer-meta-label", text: "Date" });
      dateRow.createSpan({
        text: formatRelativeDate(msg.receivedDateTime),
        attr: { title: new Date(msg.receivedDateTime).toLocaleString() },
      });
    }

    // Actions (mark read/unread + toggle processed) — bottom of header
    const actionsEl = headerEl.createDiv({ cls: "iris-viewer-actions" });

    if (!msg.isRead) {
      const markReadBtn = actionsEl.createEl("button", {
        cls: "iris-header-icon clickable-icon",
        attr: { "aria-label": "Mark as read (e)" },
      });
      setIcon(markReadBtn, "mail-open");
      markReadBtn.addEventListener("click", () => {
        this.callbacks.onMarkAsRead(msg);
      });
    } else {
      const markUnreadBtn = actionsEl.createEl("button", {
        cls: "iris-header-icon clickable-icon",
        attr: { "aria-label": "Mark as unread (e)" },
      });
      setIcon(markUnreadBtn, "mail");
      markUnreadBtn.addEventListener("click", () => {
        this.callbacks.onMarkAsUnread(msg);
      });
    }

    const id = msg.id || "";
    const isTodo = id ? this.todoIds.has(id) : false;
    const todoBtn = actionsEl.createEl("button", {
      cls: "iris-header-icon clickable-icon" + (isTodo ? " is-active" : ""),
      attr: { "aria-label": isTodo ? "Unmark to-do" : "Mark as to-do" },
    });
    setIcon(todoBtn, "check-square");
    todoBtn.addEventListener("click", () => {
      this.callbacks.onToggleTodo(msg);
    });

    const isPinned = id ? this.pinnedIds.has(id) : false;
    const pinBtn = actionsEl.createEl("button", {
      cls: "iris-header-icon clickable-icon" + (isPinned ? " is-active" : ""),
      attr: { "aria-label": isPinned ? "Unpin (p)" : "Pin — always load (p)" },
    });
    setIcon(pinBtn, isPinned ? "pin-off" : "pin");
    pinBtn.addEventListener("click", () => {
      this.callbacks.onTogglePin(msg);
    });

    const deleteBtn = actionsEl.createEl("button", {
      cls: "iris-header-icon clickable-icon",
      attr: { "aria-label": "Move to bin (#)" },
    });
    setIcon(deleteBtn, "trash-2");
    deleteBtn.addEventListener("click", () => {
      this.callbacks.onDeleteMessage(msg);
    });

    if (this.processedMarkdown) {
      const toggleBtn = actionsEl.createEl("button", {
        cls: "iris-header-icon clickable-icon",
        attr: { "aria-label": this.showingProcessed ? "View original" : "View processed" },
      });
      setIcon(toggleBtn, this.showingProcessed ? "code" : "sparkles");
      toggleBtn.addEventListener("click", () => {
        this.showingProcessed = !this.showingProcessed;
        this.rebuildView();
      });

      if (this.processedIsStale) {
        const reprocessBtn = actionsEl.createEl("button", {
          cls: "iris-header-icon clickable-icon",
          attr: { "aria-label": "Reprocess with current prompt" },
        });
        setIcon(reprocessBtn, "refresh-cw");
        reprocessBtn.addEventListener("click", () => {
          this.processedMarkdown = null;
          this.processedIsStale = false;
          this.showingProcessed = false;
          this.callbacks.onReprocessMessage(msg);
        });
      }
    }

    // Attachments — between header and body as a distinct strip
    const attachments = (msg.attachments || []).filter(
      (a) => !a.isInline,
    );
    if (attachments.length > 0) {
      const attachEl = this.containerEl.createDiv({ cls: "iris-viewer-attachments" });
      for (const att of attachments) {
        const size = att.size
          ? att.size < 1024
            ? `${att.size} B`
            : att.size < 1048576
              ? `${(att.size / 1024).toFixed(0)} KB`
              : `${(att.size / 1048576).toFixed(1)} MB`
          : "";
        const label = size ? `${att.name} (${size})` : (att.name || "Attachment");
        attachEl.createDiv({
          cls: "iris-viewer-attachment-item",
          text: label,
        });
      }
    }

    // Body
    const bodyEl = this.containerEl.createDiv({
      cls: "iris-viewer-body",
    });

    // Shared logic: context menu, annotations, and bottom panel.
    // Must run AFTER the body DOM is populated.
    const finishBody = () => {
      // Context menu for creating notes from selected text
      bodyEl.addEventListener("contextmenu", (evt) => {
        const sel = window.getSelection();
        const selectedText = sel?.toString().trim();
        if (!selectedText || !this.currentMsg) return;

        const menu = new Menu();
        menu.addItem((item) => {
          item.setTitle("Create Event Note")
            .setIcon("calendar")
            .onClick(() => {
              this.callbacks.onCreateNoteFromSelection(selectedText, "event", this.currentMsg!);
            });
        });
        menu.addItem((item) => {
          item.setTitle("Create Task Note")
            .setIcon("check-square")
            .onClick(() => {
              this.callbacks.onCreateNoteFromSelection(selectedText, "task", this.currentMsg!);
            });
        });
        menu.showAtMouseEvent(evt);
      });

    };

    if (this.showingProcessed && this.processedMarkdown) {
      bodyEl.addClass("iris-viewer-markdown");
      MarkdownRenderer.render(
        this.app,
        this.processedMarkdown,
        bodyEl,
        "",
        { register: () => {} } as never,
      ).then(finishBody);
    } else {
      this.renderHtmlBody(bodyEl, this.currentHtml);
      finishBody();
    }
  }

  /** Create a dropdown anchored to an element with auto-close on outside click. */
  private openDropdown(
    anchor: HTMLElement,
    removeSelector: string,
    cls: string,
    build: (dropdown: HTMLDivElement) => void,
  ): void {
    anchor.querySelector(removeSelector)?.remove();
    const dropdown = anchor.createDiv({ cls });
    build(dropdown);
    const closeHandler = (e: MouseEvent) => {
      if (!dropdown.contains(e.target as Node)) {
        dropdown.remove();
        document.removeEventListener("click", closeHandler, true);
      }
    };
    setTimeout(() => document.addEventListener("click", closeHandler, true), 0);
  }

  private showTagDropdown(
    anchor: HTMLElement,
    messageId: string,
    activeTags: Set<string>,
  ): void {
    this.openDropdown(anchor, ".iris-tag-dropdown", "iris-tag-dropdown", (dropdown) => {
      if (activeTags.size > 0) {
        const noneItem = dropdown.createDiv({ cls: "iris-tag-option" });
        const noneIcon = noneItem.createSpan({ cls: "iris-tag-option-icon" });
        setIcon(noneIcon, "x");
        noneItem.createSpan({ text: "Remove all" });
        noneItem.addEventListener("click", () => {
          dropdown.remove();
          if (this.currentMsg) this.callbacks.onTagChange(this.currentMsg, null);
        });
      }

      for (const cat of this.tagCategories) {
        const isActive = activeTags.has(cat);
        const item = dropdown.createDiv({
          cls: "iris-tag-option" + (isActive ? " is-selected" : ""),
        });
        const catIcon = item.createSpan({ cls: "iris-tag-option-icon" });
        setIcon(catIcon, this.tagIcons.get(cat) || "tag");
        item.createSpan({ text: cat });
        if (isActive) {
          const check = item.createSpan({ cls: "iris-tag-option-check" });
          setIcon(check, "check");
        }
        item.addEventListener("click", () => {
          if (this.currentMsg) this.callbacks.onTagChange(this.currentMsg, cat);
        });
      }
    });
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
            .setTitle("Block")
            .setIcon("trash-2")
            .onClick(() => this.callbacks.onMarkSenderAsJunk(address, rawName)),
        );
        menu.showAtMouseEvent(evt);
      });
    }
    return span;
  }

  private renderRecipients(
    row: HTMLElement,
    recipients: Message["toRecipients"],
  ): void {
    if (!recipients) return;
    for (let i = 0; i < recipients.length; i++) {
      const r = recipients[i];
      const addr = r.emailAddress?.address || "";
      const rawName = r.emailAddress?.name || addr;
      if (i > 0) row.createSpan({ text: ", " });
      this.createAddressSpan(row, addr, rawName);
    }
  }

  private renderHtmlBody(container: HTMLElement, html: string): void {
    if (!html) {
      container.createDiv({
        text: "(no content)",
        attr: { style: "padding: 16px; color: var(--text-muted);" },
      });
      return;
    }

    const htmlContainer = container.createDiv({
      cls: "iris-viewer-html-content",
    });

    // Sanitize untrusted email HTML via Obsidian's DOM sanitizer (allowlist-based,
    // strips scripts/handlers/javascript: URLs). Regex-based stripping is unsafe.
    const fragment = sanitizeHTMLToDom(html);
    htmlContainer.appendChild(fragment);
  }
}
