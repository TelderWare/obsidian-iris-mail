import { App, MarkdownRenderer, Menu, sanitizeHTMLToDom, setIcon } from "obsidian";
import type { Message } from "../../types";
import type { TagCacheEntry, DetectedItemEntry } from "../../store/types";
import type { NameResolver, EffectiveSenderResolver } from "./MessageList";
import type { NoteType } from "../../utils/claudeApi";
import { formatRelativeDate, formatItemDate } from "../../utils/dateFormat";

interface MessageViewerCallbacks {
  onMarkAsRead: (msg: Message) => void;
  onMarkAsUnread: (msg: Message) => void;
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
  /** Accept a detected item — creates the note. */
  onAcceptDetectedItem: (messageId: string, item: DetectedItemEntry) => void;
  /** Dismiss a detected item. */
  onDismissDetectedItem: (messageId: string, itemId: string) => void;
  /** Update detected item fields before accepting. */
  onUpdateDetectedItem: (messageId: string, itemId: string, updates: Partial<DetectedItemEntry>) => void;
  /** Re-run item detection for the current message. */
  onReloadDetectedItems: (messageId: string) => void;
  /** Re-process the current message with the current prompt (replaces stale cached result). */
  onReprocessMessage: (msg: Message) => void;
  /** Open the nickname editor for an address. */
  onEditNickname: (address: string, rawName: string) => void;
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
  private tagCache = new Map<string, TagCacheEntry[]>();
  private resolveEffectiveSender: EffectiveSenderResolver | null = null;
  private detectedItems: DetectedItemEntry[] = [];
  private showAllDetectedItems = false;
  private editingItemId: string | null = null;
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

  setTagCache(cache: Map<string, TagCacheEntry[]>): void {
    this.tagCache = cache;
  }

  setDetectedItems(items: DetectedItemEntry[]): void {
    this.detectedItems = items;
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
          const tagBadge = tagsLine.createSpan({
            cls: `iris-viewer-tag${isAutoTag ? " is-auto" : ""}${isObsoleteTag ? " is-obsolete" : ""}`,
            attr: {
              title: isObsoleteTag
                ? `Auto-tagged v${entry.promptVersion} (obsolete) — right-click to remove`
                : (isAutoTag ? "Auto-tagged" : "Manually tagged") + " — right-click to remove",
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
        attr: { "aria-label": "Mark as read" },
      });
      setIcon(markReadBtn, "mail-open");
      markReadBtn.addEventListener("click", () => {
        this.callbacks.onMarkAsRead(msg);
      });
    } else {
      const markUnreadBtn = actionsEl.createEl("button", {
        cls: "iris-header-icon clickable-icon",
        attr: { "aria-label": "Mark as unread" },
      });
      setIcon(markUnreadBtn, "mail");
      markUnreadBtn.addEventListener("click", () => {
        this.callbacks.onMarkAsUnread(msg);
      });
    }

    const deleteBtn = actionsEl.createEl("button", {
      cls: "iris-header-icon clickable-icon",
      attr: { "aria-label": "Move to bin" },
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

      // Bottom area: reload button + detected items panel
      if (msg.id) {
        const bottomEl = this.containerEl.createDiv({ cls: "iris-detected-items-bottom" });

        const reloadRow = bottomEl.createDiv({ cls: "iris-detected-items-reload-row" });
        const reloadBtn = reloadRow.createEl("button", {
          cls: "iris-detected-item-btn iris-detected-items-reload clickable-icon",
          attr: { "aria-label": "Detect events & tasks" },
        });
        setIcon(reloadBtn, "refresh-cw");
        reloadBtn.addEventListener("click", (e) => {
          e.stopPropagation();
          this.callbacks.onReloadDetectedItems(msg.id!);
        });

        if (this.detectedItems.length > 0) {
          this.renderDetectedItemsPanel(bottomEl, msg);
        }
      }
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

  private renderDetectedItemsPanel(parentEl: HTMLElement, msg: Message): void {
    const msgId = msg.id;
    if (!msgId) return;

    const visible = this.showAllDetectedItems
      ? this.detectedItems
      : this.detectedItems.filter((i) => i.status === "pending");

    if (visible.length === 0) return;

    const panelEl = parentEl.createDiv({ cls: "iris-detected-items-panel" });

    // Items
    const listEl = panelEl.createDiv({ cls: "iris-detected-items-list" });

    for (const item of visible) {
      const cardEl = listEl.createDiv({
        cls: `iris-detected-item-card is-${item.status}`,
      });

      // Icon
      const iconEl = cardEl.createSpan({ cls: "iris-detected-item-icon" });
      setIcon(iconEl, item.type === "event" ? "calendar" : "check-square");

      // Check if this card is being edited
      const isEditing = this.editingItemId === item.itemId;

      if (isEditing) {
        this.renderEditableCard(cardEl, msgId, item);
      } else {
        this.renderReadonlyCard(cardEl, msgId, item);
      }
    }
  }

  /** Render a card in its normal read-only state with action buttons. */
  private renderReadonlyCard(cardEl: HTMLElement, msgId: string, item: DetectedItemEntry): void {
    const contentEl = cardEl.createDiv({ cls: "iris-detected-item-content" });
    contentEl.createDiv({ cls: "iris-detected-item-title", text: item.title });

    const metaEl = contentEl.createDiv({ cls: "iris-detected-item-meta" });
    if (item.type === "event") {
      if (item.date) metaEl.createSpan({ text: formatItemDate(item.date) });
      if (item.time) metaEl.createSpan({ text: item.time });
      if (item.location) metaEl.createSpan({ text: item.location });
    } else {
      if (item.dueDate) metaEl.createSpan({ text: `Due: ${formatItemDate(item.dueDate)}` });
    }

    if (item.description) {
      contentEl.createDiv({ cls: "iris-detected-item-desc", text: item.description });
    }

    if (item.status === "accepted" && item.vaultPath) {
      const pathEl = contentEl.createDiv({ cls: "iris-detected-item-path" });
      setIcon(pathEl.createSpan(), "check");
      pathEl.createSpan({ text: item.vaultPath });
    }

    // Actions
    if (item.status === "pending") {
      const actionsEl = cardEl.createDiv({ cls: "iris-detected-item-actions" });

      const acceptBtn = actionsEl.createEl("button", {
        cls: "iris-detected-item-btn iris-detected-item-accept clickable-icon",
        attr: { "aria-label": "Accept" },
      });
      setIcon(acceptBtn, "check");
      acceptBtn.addEventListener("click", (e) => {
        e.stopPropagation();
        this.callbacks.onAcceptDetectedItem(msgId, item);
      });

      const editBtn = actionsEl.createEl("button", {
        cls: "iris-detected-item-btn iris-detected-item-edit clickable-icon",
        attr: { "aria-label": "Edit" },
      });
      setIcon(editBtn, "pencil");
      editBtn.addEventListener("click", (e) => {
        e.stopPropagation();
        this.editingItemId = item.itemId;
        this.rebuildView();
      });

      const dismissBtn = actionsEl.createEl("button", {
        cls: "iris-detected-item-btn iris-detected-item-dismiss clickable-icon",
        attr: { "aria-label": "Dismiss" },
      });
      setIcon(dismissBtn, "x");
      dismissBtn.addEventListener("click", (e) => {
        e.stopPropagation();
        this.callbacks.onDismissDetectedItem(msgId, item.itemId);
      });
    }
  }

  /** Render a card with all fields as inline editable inputs. */
  private renderEditableCard(cardEl: HTMLElement, msgId: string, item: DetectedItemEntry): void {
    cardEl.addClass("is-editing");

    const contentEl = cardEl.createDiv({ cls: "iris-detected-item-content" });

    // Row 1: Title + meta fields on a single line
    const topRow = contentEl.createDiv({ cls: "iris-detected-item-edit-row" });

    const titleInput = topRow.createEl("input", {
      cls: "iris-detected-item-inline-input iris-detected-item-inline-title",
      attr: { type: "text", value: item.title, placeholder: "Title" },
    });

    let dateInput: HTMLInputElement | undefined;
    let timeInput: HTMLInputElement | undefined;
    let locationInput: HTMLInputElement | undefined;
    let dueDateInput: HTMLInputElement | undefined;

    if (item.type === "event") {
      dateInput = topRow.createEl("input", {
        cls: "iris-detected-item-inline-input iris-detected-item-inline-meta",
        attr: { type: "text", value: item.date || "", placeholder: "YYYY-MM-DD or range" },
      });
      timeInput = topRow.createEl("input", {
        cls: "iris-detected-item-inline-input iris-detected-item-inline-meta",
        attr: { type: "time", value: item.time || "" },
      });
      // Location on its own row for events (needs more space)
      locationInput = contentEl.createEl("input", {
        cls: "iris-detected-item-inline-input iris-detected-item-inline-location",
        attr: { type: "text", value: item.location || "", placeholder: "Location" },
      });
    } else {
      dueDateInput = topRow.createEl("input", {
        cls: "iris-detected-item-inline-input iris-detected-item-inline-meta",
        attr: { type: "text", value: item.dueDate || "", placeholder: "YYYY-MM-DD or range" },
      });
    }

    // Description - inline editable textarea
    const descInput = contentEl.createEl("textarea", {
      cls: "iris-detected-item-inline-input iris-detected-item-inline-desc",
      attr: { placeholder: "Description", rows: "2" },
    });
    descInput.value = item.description || "";

    // Action buttons row
    const btnRow = contentEl.createDiv({ cls: "iris-detected-item-inline-actions" });

    const collectUpdates = (): Partial<DetectedItemEntry> => {
      const updates: Partial<DetectedItemEntry> = {
        title: titleInput.value.trim(),
        description: descInput.value.trim(),
      };
      if (item.type === "event") {
        updates.date = dateInput?.value || "";
        updates.time = timeInput?.value || "";
        updates.location = locationInput?.value.trim() || "";
      } else {
        updates.dueDate = dueDateInput?.value || "";
      }
      return updates;
    };

    const saveBtn = btnRow.createEl("button", {
      cls: "iris-detected-item-btn iris-detected-item-save clickable-icon",
      attr: { "aria-label": "Save changes" },
    });
    setIcon(saveBtn, "check");
    saveBtn.createSpan({ text: "Save" });
    saveBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      this.callbacks.onUpdateDetectedItem(msgId, item.itemId, collectUpdates());
      this.editingItemId = null;
      this.rebuildView();
    });

    const acceptBtn = btnRow.createEl("button", {
      cls: "iris-detected-item-btn iris-detected-item-accept clickable-icon",
      attr: { "aria-label": "Save & Accept" },
    });
    setIcon(acceptBtn, "check-check");
    acceptBtn.createSpan({ text: "Save & Accept" });
    acceptBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      const updates = collectUpdates();
      this.callbacks.onUpdateDetectedItem(msgId, item.itemId, updates);
      const updated = { ...item, ...updates };
      this.editingItemId = null;
      this.callbacks.onAcceptDetectedItem(msgId, updated as DetectedItemEntry);
    });

    const cancelBtn = btnRow.createEl("button", {
      cls: "iris-detected-item-btn iris-detected-item-cancel clickable-icon",
      attr: { "aria-label": "Cancel" },
    });
    setIcon(cancelBtn, "x");
    cancelBtn.createSpan({ text: "Cancel" });
    cancelBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      this.editingItemId = null;
      this.rebuildView();
    });

    // Auto-focus the title input
    requestAnimationFrame(() => titleInput.focus());
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
        this.callbacks.onEditNickname(address, rawName);
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
