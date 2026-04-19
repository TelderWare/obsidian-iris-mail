import { ItemView, WorkspaceLeaf, Notice, setIcon, Menu } from "obsidian";
import type IrisMailPlugin from "../main";
import {
  VIEW_TYPE_IRIS_MAIL,
  DEFAULT_CLAUDE_PROMPT,
  NICKNAME_PROMPT,
  NICKNAME_BATCH_PROMPT,
  TAG_CLASSIFY_PROMPT,
  TAG_ICON_POOL,
  ITEM_DETECTION_PROMPT,
  parseTagCategories,
  getTagVersion,
  bumpTagVersion,
} from "../constants";
import { MessageList } from "./components/MessageList";
import { MessageViewer } from "./components/MessageViewer";
import { NicknameModal } from "./components/NicknameModal";
import { CreateTagModal } from "./components/CreateTagModal";
import { SearchBar } from "./components/SearchBar";
import { Toolbar } from "./components/Toolbar";
import { processEmailWithClaude, classifyEmailTagsYesNo, refineTagCriteria, generateNickname, generateNicknamesBatch, mergeEmailsToFormula, refineTagCriteriaBulk, extractNoteFromSelection, detectItemsInEmail, hasClaudeAccess, generateTagDescription, pickTagIcon, type TagCandidate } from "../utils/claudeApi";
import type { NoteType, ExtractedNote } from "../utils/claudeApi";
import { htmlToMarkdown } from "../utils/htmlToMarkdown";
import { extractForwardedSender } from "../utils/extractForwardedSender";
import { getEnvelopeSender } from "../utils/envelopeSender";
import { logger } from "../utils/logger";
import { EmailStore } from "../store/EmailStore";
import { EmailClassifier } from "../services/EmailClassifier";
import type { TagCacheEntry, DetectedItemEntry } from "../store/types";
import type {
  Message,
  MessageListState,
  SenderGroup,
  GraphPagedResponse,
} from "../types";

/**
 * Normalize "LastName, FirstName" to "FirstName LastName".
 * Leaves other formats untouched.
 */
function normalizeName(raw: string): string {
  // Strip Outlook delegate suffix, e.g. "Name (via Institution)"
  let name = raw.replace(/\s*\(via\s+[^)]+\)\s*$/i, "").trim();
  const m = name.match(/^([^,]+),\s*(.+)$/);
  if (m) name = `${m[2]} ${m[1]}`;
  return name;
}

export class InboxView extends ItemView {
  private plugin: IrisMailPlugin;
  private messageList!: MessageList;
  private messageViewer!: MessageViewer;
  private searchBar!: SearchBar;
  private toolbar!: Toolbar;
  private compactResizeObserver: ResizeObserver | null = null;

  private inboxFolderId: string | null = null;
  private messageState: MessageListState = {
    messages: [],
    nextLink: null,
    isLoading: false,
    searchQuery: "",
  };

  // View mode: flat messages (default) or senders
  private viewMode!: "messages" | "senders";
  private senderGroups: SenderGroup[] = [];
  private activeSender: SenderGroup | null = null;
  private viewModeToggleBtn!: HTMLButtonElement;
  private sortNewestFirst!: boolean;
  private sortToggleBtn!: HTMLButtonElement;
  private filterWrap!: HTMLDivElement;
  private filterUnreadOnly!: boolean;
  private unreadOptBtn!: HTMLButtonElement;

  private selectedMessageId: string | null = null;
  private lastStrippedHtml: string = "";

  // Extracted classifier handles classification & tagging caches
  private classifier!: EmailClassifier;

  // In-memory caches
  private processedCache = new Map<string, string>();
  private nicknameCache = new Map<string, string>();
  private detectedItemsCache = new Map<string, DetectedItemEntry[]>();

  // Convenience accessors for classifier caches
  private get tagCache() { return this.classifier.tags; }
  private tagWrap!: HTMLDivElement;
  private topBar!: HTMLDivElement;

  // Prefetch state
  private prefetchGeneration = 0;
  private prefetchAllPromise: Promise<void> | null = null;
  private prefetchInflight = new Set<string>();

  constructor(leaf: WorkspaceLeaf, plugin: IrisMailPlugin) {
    super(leaf);
    this.plugin = plugin;
    this.classifier = new EmailClassifier(plugin.store, () => plugin.settings);
    const s = plugin.settings;
    this.viewMode = s.viewMode;
    this.sortNewestFirst = s.sortNewestFirst;
    this.filterUnreadOnly = s.filterUnreadOnly;
  }

  getViewType(): string {
    return VIEW_TYPE_IRIS_MAIL;
  }

  getDisplayText(): string {
    return "Email";
  }

  getIcon(): string {
    return "mail";
  }

  async onOpen(): Promise<void> {
    const container = this.contentEl;
    container.empty();
    container.addClass("iris-mail-container");

    // Hide the top toolbar when vertical space is at a premium.
    const COMPACT_HEIGHT_PX = 700;
    const updateCompact = () => {
      container.toggleClass(
        "iris-compact",
        container.clientHeight > 0 && container.clientHeight < COMPACT_HEIGHT_PX,
      );
    };
    if (this.compactResizeObserver) {
      this.compactResizeObserver.disconnect();
    }
    this.compactResizeObserver = new ResizeObserver(() => updateCompact());
    this.compactResizeObserver.observe(container);
    updateCompact();

    if (
      !this.plugin.authProvider.isSignedIn() &&
      this.plugin.settings.clientId
    ) {
      try {
        await this.plugin.authProvider.initialize(this.plugin.settings);
      } catch {
        // silent init failed
      }
    }

    if (!this.plugin.authProvider.isSignedIn()) {
      this.renderSignInPrompt(container);
      return;
    }

    this.reloadCaches();
    this.renderInboxUI(container);
    await this.loadInbox();
  }

  async onClose(): Promise<void> {
    this.prefetchGeneration++;
    if (this.compactResizeObserver) {
      this.compactResizeObserver.disconnect();
      this.compactResizeObserver = null;
    }
    this.contentEl.empty();
  }

  async refresh(): Promise<void> {
    this.prefetchGeneration++;
    this.processedCache.clear();

    // Re-read persistent caches without tearing down the UI
    this.reloadCaches();

    // Reload data in-place (the existing topbar + message list stay mounted)
    await this.loadInbox();
  }

  private reloadCaches(): void {
    this.classifier.reloadCaches();
    this.nicknameCache = this.plugin.store.getAllNicknames();
  }

  // --- Private: rendering ---

  private renderSignInPrompt(container: HTMLElement): void {
    const prompt = container.createDiv({ cls: "iris-sign-in-prompt" });
    const icon = prompt.createDiv({ cls: "iris-sign-in-icon" });
    setIcon(icon, "mail");
    prompt.createEl("h3", { text: "Iris Mail" });
    prompt.createEl("p", {
      cls: "iris-sign-in-desc",
      text: "Connect your Microsoft account to get started.",
    });

    const btnGroup = prompt.createDiv({ cls: "iris-sign-in-buttons" });

    const browserBtn = btnGroup.createEl("button", {
      text: "Sign in with browser",
      cls: "mod-cta",
    });

    const deviceBtn = btnGroup.createEl("button", {
      text: "Sign in with device code",
    });

    const startLogin = (method: "auth-code" | "device-code") => {
      if (!this.plugin.settings.clientId) {
        this.showClientIdEntry(prompt, btnGroup, method);
        return;
      }
      if (method === "auth-code") this.plugin.handleLoginWithAuthCode();
      else this.plugin.handleLoginWithDeviceCode();
    };

    browserBtn.addEventListener("click", () => startLogin("auth-code"));
    deviceBtn.addEventListener("click", () => startLogin("device-code"));
  }

  private showClientIdEntry(
    prompt: HTMLElement,
    btnGroup: HTMLElement,
    method: "auth-code" | "device-code",
  ): void {
    // Avoid stacking multiple entries on repeat clicks
    const existing = prompt.querySelector(".iris-clientid-entry");
    if (existing) {
      (existing.querySelector("input") as HTMLInputElement | null)?.focus();
      return;
    }
    btnGroup.hide();

    const entry = prompt.createDiv({ cls: "iris-clientid-entry" });
    entry.createEl("label", {
      cls: "iris-clientid-label",
      text: "Azure Application (client) ID",
    });

    const inputWrap = entry.createDiv({ cls: "iris-clientid-input-wrap" });
    const input = inputWrap.createEl("input", {
      cls: "iris-clientid-input",
      attr: {
        type: "text",
        placeholder: "00000000-0000-0000-0000-000000000000",
        spellcheck: "false",
        autocomplete: "off",
      },
    }) as HTMLInputElement;

    const error = entry.createDiv({ cls: "iris-clientid-error" });

    const help = entry.createEl("a", {
      cls: "iris-clientid-help",
      text: "Open Azure portal →",
      href: "https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade",
    });
    help.setAttr("target", "_blank");
    help.setAttr("rel", "noopener");

    const actions = entry.createDiv({ cls: "iris-clientid-actions" });
    const cancelBtn = actions.createEl("button", { text: "Cancel" });
    const submitBtn = actions.createEl("button", {
      text: "Continue",
      cls: "mod-cta",
    });
    submitBtn.disabled = true;

    const guidRe =
      /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

    const validate = (): boolean => {
      const v = input.value.trim();
      const ok = guidRe.test(v);
      submitBtn.disabled = !ok;
      error.setText(v.length > 0 && !ok ? "Must be a valid GUID." : "");
      return ok;
    };

    const submit = async () => {
      if (!validate()) return;
      this.plugin.settings.clientId = input.value.trim();
      await this.plugin.saveSettings();
      entry.remove();
      btnGroup.show();
      if (method === "auth-code") this.plugin.handleLoginWithAuthCode();
      else this.plugin.handleLoginWithDeviceCode();
    };

    input.addEventListener("input", validate);
    input.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        void submit();
      } else if (e.key === "Escape") {
        e.preventDefault();
        entry.remove();
        btnGroup.show();
      }
    });
    cancelBtn.addEventListener("click", () => {
      entry.remove();
      btnGroup.show();
    });
    submitBtn.addEventListener("click", () => void submit());

    setTimeout(() => input.focus(), 0);
  }

  private renderInboxUI(container: HTMLElement): void {
    // Top bar: senders toggle (left) + search & refresh (right)
    const topBar = container.createDiv({ cls: "iris-topbar" });

    // Senders view toggle (off = flat messages; on = grouped by sender)
    this.viewModeToggleBtn = topBar.createEl("button", {
      cls:
        "iris-topbar-btn clickable-icon" +
        (this.viewMode === "senders" ? " is-active" : ""),
      attr: { "aria-label": "Group by sender" },
    });
    setIcon(this.viewModeToggleBtn, "users");
    this.viewModeToggleBtn.addEventListener("click", () =>
      this.handleViewModeToggle(),
    );

    // Sort toggle
    this.sortToggleBtn = topBar.createEl("button", {
      cls: "iris-topbar-btn clickable-icon",
      attr: { "aria-label": this.sortNewestFirst ? "Newest first" : "Oldest first" },
    });
    setIcon(this.sortToggleBtn, this.sortNewestFirst ? "arrow-up" : "arrow-down");
    this.sortToggleBtn.addEventListener("click", () => {
      this.sortNewestFirst = !this.sortNewestFirst;
      setIcon(this.sortToggleBtn, this.sortNewestFirst ? "arrow-up" : "arrow-down");
      this.sortToggleBtn.setAttribute("aria-label", this.sortNewestFirst ? "Newest first" : "Oldest first");
      this.regroupAndSync();
      this.renderCurrentView();
      this.persistViewState();
    });

    // Filter bar: icon + expandable toggle buttons
    this.filterWrap = topBar.createDiv({ cls: "iris-filter-wrap has-options" });
    const filterIcon = this.filterWrap.createEl("button", {
      cls: "iris-topbar-btn clickable-icon",
      attr: { "aria-label": "Filter" },
    });
    setIcon(filterIcon, "list-filter");

    // Expandable filter toggles (revealed on hover)
    this.unreadOptBtn = this.createFilterButton(
      this.filterWrap, "mail", "Unread only",
      () => this.filterUnreadOnly,
      () => { this.filterUnreadOnly = !this.filterUnreadOnly; },
    );

    // Tag bar: existing tag icons + add button
    this.topBar = topBar;
    this.tagWrap = topBar.createDiv({ cls: "iris-tag-wrap" });
    this.rebuildTagWrap();

    const rightControls = topBar.createDiv({ cls: "iris-topbar-right" });

    this.searchBar = new SearchBar(rightControls, {
      onSearch: (query: string) => this.handleSearch(query),
    });

    // AI menu
    const aiMenuBtn = rightControls.createEl("button", {
      cls: "iris-topbar-btn clickable-icon",
      attr: { "aria-label": "AI actions" },
    });
    setIcon(aiMenuBtn, "brain-circuit");
    aiMenuBtn.addEventListener("click", (evt) =>
      this.showAiMenu(evt),
    );

    this.toolbar = new Toolbar(rightControls, {
      onRefresh: () => this.refresh(),
    });

    // Main area: message list + viewer (no sidebar)
    const mainEl = container.createDiv({ cls: "iris-main" });
    const rightPane = mainEl.createDiv({ cls: "iris-right-pane" });

    const nameResolver = (addr: string, raw: string) =>
      this.resolveName(addr, raw);
    const effectiveSenderResolver = this.plugin.settings.resolveForwardedSender
      ? (msg: Message) => this.getEffectiveSender(msg)
      : null;

    const listEl = rightPane.createDiv({ cls: "iris-message-list" });
    this.messageList = new MessageList(listEl, {
      onMessageSelect: (msg: Message) => this.handleMessageSelect(msg),
      onSenderSelect: (sender: SenderGroup) =>
        this.handleSenderSelect(sender),
      onBack: () => this.handleBack(),
      onLoadMore: () => this.handleLoadMore(),
      onMultiSelect: (ids: Set<string>) => this.handleMultiSelect(ids),
      onEditNickname: (addr: string, rawName: string) => this.openNicknameModal(addr, rawName),
    }, nameResolver, effectiveSenderResolver);

    const viewerEl = rightPane.createDiv({ cls: "iris-message-viewer" });
    this.messageViewer = new MessageViewer(viewerEl, this.plugin.app, {
      onMarkAsRead: (msg: Message) => this.handleMarkAsRead(msg),
      onMarkAsUnread: (msg: Message) => this.handleMarkAsUnread(msg),
      onTagChange: (msg: Message, tag: string | null) => this.handleTagChange(msg, tag),
      onRetagMessage: (msg) => this.handleRetagMessage(msg),
      onBatchMarkAsRead: (ids) => this.handleBatchMarkAsRead(ids),
      onBatchMarkAsUnread: (ids) => this.handleBatchMarkAsUnread(ids),
      onBatchTag: (ids, tag) => this.handleBatchTag(ids, tag),
      onBulkDenyTag: (ids, tag) => this.handleBulkDenyTag(ids, tag),
      onCreateNoteFromSelection: (text, noteType, msg) => this.handleCreateNoteFromSelection(text, noteType, msg),
      onAcceptDetectedItem: (messageId, item) => this.handleAcceptDetectedItem(messageId, item),
      onDismissDetectedItem: (messageId, itemId) => this.handleDismissDetectedItem(messageId, itemId),
      onUpdateDetectedItem: (messageId, itemId, updates) => this.handleUpdateDetectedItem(messageId, itemId, updates),
      onReloadDetectedItems: (messageId) => this.handleReloadDetectedItems(messageId),
      onReprocessMessage: (msg) => this.handleReprocessMessage(msg),
      onEditNickname: (addr: string, rawName: string) => this.openNicknameModal(addr, rawName),
      onDismiss: () => { this.messageViewer.clear(); },
    }, nameResolver);
    this.messageViewer.setEffectiveSenderResolver(effectiveSenderResolver);
    this.messageViewer.setTagCategories(this.getTagCategories());
    this.messageViewer.setTagIcons(this.getTagIconMap());
    this.messageViewer.setTagCache(this.tagCache);
    this.messageViewer.setPromptVersions(
      this.plugin.settings.tagPromptVersions || {},
    );
  }

  // --- Private: shared helpers ---

  /** Create a filter-option toggle button wired to applyFilters + persistViewState. */
  private createFilterButton(
    parent: HTMLElement,
    icon: string,
    label: string,
    isActive: () => boolean,
    toggle: () => void,
  ): HTMLButtonElement {
    const btn = parent.createEl("button", {
      cls: "iris-filter-opt clickable-icon" + (isActive() ? " is-active" : ""),
      attr: { "aria-label": label },
    });
    setIcon(btn, icon);
    btn.addEventListener("click", () => {
      toggle();
      btn.toggleClass("is-active", isActive());
      this.applyFilters();
      this.persistViewState();
    });
    return btn;
  }

  /** Recompute sender groupings and update badge after any state change. */
  private regroupAndSync(): void {
    this.senderGroups = this.groupByEffectiveSender(this.messageState.messages);
    this.syncBadge();
  }

  /** Re-render the viewer for the currently selected message. */
  private renderSelectedMessage(msg: Message): void {
    if (!msg.id || this.selectedMessageId !== msg.id) return;
    this.messageViewer.setDetectedItems(this.detectedItemsCache.get(msg.id) || []);
    this.messageViewer.render(msg, this.lastStrippedHtml);
  }

  /** Reset drill-down/selection state and clear viewer. */
  private clearDrillDown(): void {
    this.activeSender = null;
    this.selectedMessageId = null;
    this.messageList.clearMultiSelection();
    this.messageViewer.clear();
  }

  /** Batch mark messages as read or unread. */
  private handleBatchReadState(ids: Set<string>, markAsRead: boolean): void {
    const changed: string[] = [];
    for (const msg of this.messageState.messages) {
      if (msg.id && ids.has(msg.id) && msg.isRead !== markAsRead) {
        msg.isRead = markAsRead;
        if (markAsRead) this.plugin.store.markRead(msg.id);
        else this.plugin.store.markUnread(msg.id);
        changed.push(msg.id);
      }
    }

    // Sync to Graph API in background, rolling back any that fail.
    const api = this.plugin.mailApi;
    for (const id of changed) {
      const call = markAsRead ? api.markAsRead(id) : api.markAsUnread(id);
      void call.catch((err) => this.rollbackReadState(id, !markAsRead, err));
    }

    this.regroupAndSync();
    this.messageList.clearMultiSelection();
    this.messageViewer.clear();
    this.renderCurrentView();
  }

  /**
   * Revert an optimistic read-state change when the Graph API sync fails, so
   * the local view stays consistent with the server. `prevIsRead` is the
   * state to restore (i.e. what it was before the failed update).
   */
  private rollbackReadState(id: string, prevIsRead: boolean, err: unknown): void {
    logger.warn("InboxView", "Failed to sync read state", err);
    const canonical = this.messageState.messages.find((m) => m.id === id);
    if (canonical) canonical.isRead = prevIsRead;
    if (prevIsRead) this.plugin.store.markRead(id);
    else this.plugin.store.markUnread(id);
    this.regroupAndSync();
    this.renderCurrentView();
    const action = prevIsRead ? "unread" : "read";
    new Notice(`Couldn't mark as ${action} on server — reverted.`);
  }

  // --- Private: event handlers ---

  private async loadInbox(): Promise<void> {
    try {
      const folders = await this.plugin.mailApi.listFolders();
      const inbox = folders.find(
        (f) => f.displayName?.toLowerCase() === "inbox",
      );
      if (inbox) {
        this.inboxFolderId = inbox.id || null;
        await this.loadMessages(inbox.id!);
        void this.syncLocalReadStateToServer();
      }
    } catch (err: unknown) {
      if (!this.plugin.authProvider.isSignedIn()) {
        this.renderCurrentView();
      } else {
        const msg = err instanceof Error ? err.message : String(err);
        new Notice(`Failed to load inbox: ${msg}`);
      }
    }
  }

  /**
   * Push locally-tracked read states to the Graph API, then clear them
   * from the local store (the server is now the source of truth).
   */
  private async syncLocalReadStateToServer(): Promise<void> {
    const ids = this.plugin.store.getLocallyReadIds();
    if (ids.length === 0) return;

    logger.info("InboxView", `Syncing ${ids.length} local read states to server`);
    const api = this.plugin.mailApi;

    for (const id of ids) {
      try {
        await api.markAsRead(id);
        this.plugin.store.clearLocalRead(id);
      } catch (err) {
        // Message may have been deleted server-side — just clear it
        logger.warn("InboxView", `Failed to sync read state for ${id}`, err);
        this.plugin.store.clearLocalRead(id);
      }
    }
  }

  private async loadMessages(folderId: string): Promise<void> {
    this.messageState.isLoading = true;
    this.messageList.showLoading();

    const showRead = this.plugin.settings.showReadEmails;
    const searchQuery = this.messageState.searchQuery || undefined;
    try {
      const filter = !showRead ? "isRead eq false" : undefined;

      const response: GraphPagedResponse<Message> =
        await this.plugin.mailApi.listMessages(folderId, {
          top: this.plugin.settings.pageSize,
          search: searchQuery,
          filter,
        });

      this.messageState.messages = response.value;
      this.plugin.store.applyReadState(this.messageState.messages);
      this.messageState.nextLink = response["@odata.nextLink"] || null;
      // Cache non-search list results so we can fall back on next failure.
      if (!searchQuery) {
        this.plugin.store.setMessageList(folderId, showRead, response.value, this.messageState.nextLink);
      }
      this.regroupAndSync();
      this.renderCurrentView();
      this.startBackgroundProcessing();
    } catch (err: unknown) {
      if (!this.plugin.authProvider.isSignedIn()) {
        this.renderCurrentView();
      } else {
        // Fall back to cached list if available (non-search queries only).
        const cached = !searchQuery
          ? this.plugin.store.getMessageList(folderId, showRead)
          : undefined;
        if (cached) {
          this.messageState.messages = cached.messages as Message[];
          this.plugin.store.applyReadState(this.messageState.messages);
          this.messageState.nextLink = cached.nextLink;
          this.regroupAndSync();
          this.renderCurrentView();
          const errMsg = err instanceof Error ? err.message : String(err);
          new Notice(`Offline — showing cached messages (${errMsg})`);
        } else {
          const errMsg = err instanceof Error ? err.message : String(err);
          new Notice(`Failed to load messages: ${errMsg}`);
        }
      }
    } finally {
      this.messageState.isLoading = false;
    }
  }

  private handleMessageSelect(msg: Message): void {
    void this.showMessageInViewer(msg);
  }

  private handleMarkAsRead(msg: Message): void {
    msg.isRead = true;

    // Propagate to the canonical message list so badge/filters see the change
    if (msg.id) {
      const canonical = this.messageState.messages.find((m) => m.id === msg.id);
      if (canonical && canonical !== msg) canonical.isRead = true;

      this.plugin.store.markRead(msg.id);
      const id = msg.id;
      void this.plugin.mailApi.markAsRead(id).catch((err) =>
        this.rollbackReadState(id, false, err));
    }

    this.regroupAndSync();

    // Auto-advance to next unread. The list re-render animates the disappearing row.
    const next = this.findNextUnread(msg.id || null);
    this.renderCurrentView();
    if (next) {
      void this.showMessageInViewer(next);
      return;
    }

    // Nothing unread left; if we were in a sender drill-down that's now empty,
    // fall back to the top-level list.
    if (this.activeSender) {
      const updated = this.senderGroups.find(
        (s) => s.groupKey === this.activeSender!.groupKey,
      );
      const remaining = updated
        ? this.filterSenderMessages(updated.messages)
        : [];
      if (remaining.length === 0) {
        this.handleBack();
        return;
      }
      this.selectedMessageId = null;
      this.messageViewer.clear();
    }
  }

  /**
   * Find the next unread message in the current view (sender drill-down or
   * top-level flat messages list), in the active sort order.
   */
  private findNextUnread(currentId: string | null): Message | null {
    const dir = this.sortNewestFirst ? -1 : 1;
    let src: Message[];
    if (this.activeSender) {
      const updated = this.senderGroups.find(
        (s) => s.groupKey === this.activeSender!.groupKey,
      );
      src = updated?.messages || this.activeSender.messages;
    } else {
      src = this.messageState.messages;
    }
    const list = [...src].sort(
      (a, b) =>
        dir *
        (new Date(a.receivedDateTime || 0).getTime() -
          new Date(b.receivedDateTime || 0).getTime()),
    );
    for (const m of list) {
      if (!m.isRead && m.id !== currentId) return m;
    }
    return null;
  }

  private handleMarkAsUnread(msg: Message): void {
    msg.isRead = false;

    if (msg.id) {
      const canonical = this.messageState.messages.find((m) => m.id === msg.id);
      if (canonical && canonical !== msg) canonical.isRead = false;

      this.plugin.store.markUnread(msg.id);
      const id = msg.id;
      void this.plugin.mailApi.markAsUnread(id).catch((err) =>
        this.rollbackReadState(id, true, err));
    }

    this.regroupAndSync();
    this.renderSelectedMessage(msg);

    // Refresh the list to show the message as unread
    this.renderCurrentView();
  }

  private handleMultiSelect(selectedIds: Set<string>): void {
    if (selectedIds.size <= 1) return;
    this.selectedMessageId = null;
    this.messageViewer.renderBatchPanel(selectedIds.size, selectedIds);
  }

  private handleBatchMarkAsRead(ids: Set<string>): void {
    this.handleBatchReadState(ids, true);
  }

  private handleBatchMarkAsUnread(ids: Set<string>): void {
    this.handleBatchReadState(ids, false);
  }

  private handleBatchTag(ids: Set<string>, tag: string): void {
    const msgIds = this.resolveBatchMessageIds(ids);
    for (const msgId of msgIds) {
      const existing = this.tagCache.get(msgId) || [];
      if (existing.some((e) => e.tag === tag)) continue;
      const entry: TagCacheEntry = {
        messageId: msgId,
        tag,
        source: "manual",
        taggedAt: Date.now(),
      };
      this.tagCache.set(msgId, [...existing, entry]);
      this.plugin.store.setTag(msgId, tag, "manual");
    }
    this.messageViewer.setTagCache(this.tagCache);
    this.messageList.clearMultiSelection();
    this.messageViewer.clear();
  }

  /** Bulk deny tag: remove tag from all selected, merge into formula, refine prompt. */
  private async handleBulkDenyTag(ids: Set<string>, tag: string): Promise<void> {
    const s = this.plugin.settings;
    if (!hasClaudeAccess(s.anthropicApiKey)) return;

    const msgIds = this.resolveBatchMessageIds(ids);
    const contents: string[] = [];

    // Remove the denied tag from each message immediately
    for (const msgId of msgIds) {
      const existing = this.tagCache.get(msgId) || [];
      const updated = existing.filter((e) => e.tag !== tag);
      if (updated.length === 0) {
        this.tagCache.delete(msgId);
      } else {
        this.tagCache.set(msgId, updated);
      }
      this.plugin.store.removeTag(msgId, tag);

      const msg = this.messageState.messages.find((m) => m.id === msgId);
      if (msg) {
        const content = this.getClassifiableContent(msg);
        if (content) contents.push(content);
      }
    }

    this.messageViewer.setTagCache(this.tagCache);
    this.messageList.clearMultiSelection();
    this.messageViewer.clear();

    if (contents.length === 0) return;

    // Merge → refine in background
    try {
      new Notice(`Merging ${contents.length} emails into formula…`);
      const formula = await mergeEmailsToFormula(s.anthropicApiKey, contents);

      const refined = await refineTagCriteriaBulk(
        s.anthropicApiKey,
        tag,
        s.tagDescriptions?.[tag] || "",
        formula,
        "incorrect",
      );
      this.applyRefinedTagCriteria(tag, refined, `Criteria for "${tag}" changed from ${contents.length} emails`);
    } catch (err) {
      logger.warn("InboxView", "Bulk criteria refinement failed", err);
      new Notice("Bulk refinement failed — tags still removed.");
    }
  }

  /** Batch IDs are always message IDs now that the top-level list is flat. */
  private resolveBatchMessageIds(ids: Set<string>): string[] {
    return [...ids];
  }

  private handleBack(): void {
    this.clearDrillDown();
    this.renderCurrentView();
  }

  private async showMessageInViewer(msg: Message): Promise<void> {
    this.selectedMessageId = msg.id || null;
    this.messageList.setSelectedMessageId(msg.id || null);
    const cache = this.plugin.store;

    // Snapshot effective sender before body resolution so we can detect changes
    const effectiveBefore = this.getEffectiveSender(msg).address.toLowerCase();

    // Resolve body: L1 msg object → L2 disk cache → L3 Graph API
    let bodyHtml = msg.body?.content || "";
    let stripped = "";

    if (!bodyHtml && msg.id) {
      // Try disk cache first
      const diskBody = cache.getBody(msg.id);
      if (diskBody) {
        bodyHtml = diskBody.bodyHtml;
        stripped = diskBody.strippedHtml;
        msg.body = { content: bodyHtml, contentType: "html" };
      } else {
        // Fall back to Graph API
        try {
          const fullMsg = await this.plugin.mailApi.getMessageBody(msg.id);
          bodyHtml = fullMsg.body?.content || "";
          msg.body = fullMsg.body;
          // Backfill disk cache
          if (bodyHtml) {
            const entry = cache.setBody(msg, bodyHtml);
            stripped = entry.strippedHtml;
          }
        } catch (err) {
          logger.warn("InboxView", "Failed to fetch body by ID", err);
        }
      }
    } else if (bodyHtml && msg.id) {
      // Body was already on the message object — ensure disk cache is warm
      const diskBody = cache.getBody(msg.id);
      if (diskBody) {
        stripped = diskBody.strippedHtml;
      } else {
        const entry = cache.setBody(msg, bodyHtml);
        stripped = entry.strippedHtml;
      }
    }

    // If the effective sender changed after body resolution (e.g. forwarded
    // sender now extracted), silently regroup so the sender list is correct
    // when the user navigates back.
    const effectiveAfter = this.getEffectiveSender(msg).address.toLowerCase();
    if (effectiveAfter !== effectiveBefore) {
      this.regroupAndSync();
    }

    // Guard against stale render (user clicked something else while fetching)
    if (this.selectedMessageId !== (msg.id || null)) return;

    this.lastStrippedHtml = stripped;
    this.messageViewer.setDetectedItems(this.detectedItemsCache.get(msg.id!) || []);
    this.messageViewer.render(msg, stripped);

    // On-demand item detection: run if never scanned, OR if previously
    // scanned with 0 results (earlier scan may have had empty content).
    // Only fire here when Claude-processed markdown is already cached;
    // otherwise detection will be triggered after processing completes.
    if (msg.id) {
      const scanned = this.plugin.store.hasItemsScan(msg.id);
      const hasItems = (this.detectedItemsCache.get(msg.id)?.length ?? 0) > 0;
      if ((!scanned || !hasItems) && cache.getProcessed(msg.id)) {
        void this.detectItemsOnDemand(msg);
      }
    }

    // --- Processed markdown: L1 memory → L2 disk → Claude API ---
    const msgId = msg.id!;
    const s = this.plugin.settings;
    const effectivePrompt = s.claudeSystemPrompt || DEFAULT_CLAUDE_PROMPT;
    const promptHash = EmailStore.hashPrompt(effectivePrompt);

    // L1: in-memory
    const cachedL1 = this.processedCache.get(msgId);
    if (cachedL1) {
      const stale = !cache.hasProcessed(msgId, promptHash);
      this.messageViewer.showProcessedMarkdown(msgId, cachedL1, stale);
      return;
    }

    // L2: disk cache (serve even if prompt hash is stale — user can reprocess)
    if (cache.hasProcessed(msgId)) {
      const entry = cache.getProcessed(msgId)!;
      const stale = entry.promptHash !== promptHash;
      this.processedCache.set(msgId, entry.processedMarkdown);
      this.messageViewer.showProcessedMarkdown(msgId, entry.processedMarkdown, stale);
      return;
    }

    // If prefetch is currently processing this message, wait for it
    if (this.prefetchInflight.has(msgId)) {
      this.messageViewer.showProcessingIndicator();
      const waitStart = Date.now();
      while (this.prefetchInflight.has(msgId) && Date.now() - waitStart < 10000) {
        await new Promise((r) => setTimeout(r, 200));
      }
      if (this.selectedMessageId !== msgId) return;
      const prefetched = this.processedCache.get(msgId);
      if (prefetched) {
        this.messageViewer.showProcessedMarkdown(msgId, prefetched);
        return;
      }
      // Prefetch didn't finish or failed — fall through to process normally
    }

    // L3: Claude API processing
    if (!s.enableClaudeProcessing) return;
    if (!hasClaudeAccess(s.anthropicApiKey)) {
      logger.warn("InboxView", "Claude processing enabled but no API key set and no iris-router relay available");
      return;
    }
    if (!stripped) return;

    const parsedBody = htmlToMarkdown(stripped);
    if (!parsedBody) return;

    this.messageViewer.showProcessingIndicator();

    if (this.selectedMessageId !== msgId) return;

    // Prepend email context (subject, sender, date) so Claude has full context
    // even when the body is sparse (e.g. meeting invitations, calendar events).
    const parsedContent = this.buildEmailContext(msg) + parsedBody;

    processEmailWithClaude(s.anthropicApiKey, s.claudeModel, effectivePrompt, parsedContent)
      .then(async (markdown) => {
        // Store raw markdown in memory
        this.processedCache.set(msgId, markdown);

        // Persist to data.json
        try {
          await cache.setProcessed(msgId, markdown, promptHash);
        } catch (err) {
          logger.warn("InboxView", "Failed to save processed email", err);
        }

        if (this.selectedMessageId === msgId) {
          this.messageViewer.showProcessedMarkdown(msgId, markdown);
        }

        // Trigger item detection now that processed markdown is available
        const scanned = this.plugin.store.hasItemsScan(msgId);
        const hasItems = (this.detectedItemsCache.get(msgId)?.length ?? 0) > 0;
        if (!scanned || !hasItems) {
          void this.detectItemsOnDemand(msg);
        }
      })
      .catch((err) => {
        const errMsg = err instanceof Error ? err.message : String(err);
        new Notice(`Claude processing failed: ${errMsg}`);
        if (this.selectedMessageId === msgId) {
          this.messageViewer.hideProcessingIndicator();
        }
      });
  }

  private async handleSearch(query: string): Promise<void> {
    this.messageState.searchQuery = query;
    if (!this.inboxFolderId) return;
    await this.loadMessages(this.inboxFolderId);
  }

  private async handleLoadMore(): Promise<void> {
    if (!this.messageState.nextLink) return;

    try {
      const response: GraphPagedResponse<Message> =
        await this.plugin.mailApi.listMessages("", {
          nextLink: this.messageState.nextLink,
        });

      this.plugin.store.applyReadState(response.value);
      this.messageState.messages = [
        ...this.messageState.messages,
        ...response.value,
      ];
      this.messageState.nextLink = response["@odata.nextLink"] || null;

      this.regroupAndSync();
      this.renderCurrentView();
      this.startBackgroundProcessing();
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      new Notice(`Failed to load more messages: ${msg}`);
    }
  }

  // --- Private: AI menu ---

  private showAiMenu(evt: MouseEvent): void {
    const { x, y } = { x: evt.clientX, y: evt.clientY };
    const menu = new Menu();
    menu.addItem((item) =>
      item
        .setTitle("Reset prompts to default")
        .setIcon("rotate-ccw")
        .onClick(() => this.showResetPromptsMenu(x, y)),
    );
    menu.showAtMouseEvent(evt);
  }

  private showResetPromptsMenu(x: number, y: number): void {
    const menu = new Menu();
    menu.addItem((item) =>
      item
        .setTitle("Reset tagging prompt")
        .setIcon("tags")
        .onClick(() => {
          this.plugin.settings.tagClassifyPrompt = "";
          this.bumpAllTagVersions();
          void this.plugin.saveSettings();
          this.syncPromptVersions();
          new Notice("Tagging prompt reset to default");
        }),
    );
    menu.showAtPosition({ x, y });
  }

  // --- Private: view mode ---

  private applyFilters(): void {
    this.clearDrillDown();
    this.renderCurrentView();
    this.syncBadge();
  }

  private persistViewState(): void {
    const s = this.plugin.settings;
    s.viewMode = this.viewMode;
    s.sortNewestFirst = this.sortNewestFirst;
    s.filterUnreadOnly = this.filterUnreadOnly;
    void this.plugin.saveSettings();
  }

  private handleViewModeToggle(): void {
    this.viewMode = this.viewMode === "messages" ? "senders" : "messages";
    this.viewModeToggleBtn.toggleClass("is-active", this.viewMode === "senders");
    this.clearDrillDown();
    this.renderCurrentView();
    this.persistViewState();
  }

  private renderCurrentView(): void {
    if (!this.plugin.authProvider.isSignedIn()) {
      this.messageList.renderLoggedOut(
        () => this.plugin.handleLoginWithAuthCode(),
        () => this.plugin.handleLoginWithDeviceCode(),
      );
      return;
    }

    // If we're drilled into a sender's messages, re-render with current sort/filters
    if (this.activeSender) {
      const updated = this.senderGroups.find(
        (s) => s.groupKey === this.activeSender!.groupKey,
      );
      if (updated) {
        this.activeSender = updated;
        const displayName = this.resolveName(
          updated.address,
          updated.name || updated.address,
        );
        const msgs = this.filterSenderMessages(updated.messages);
        void this.messageList.renderSenderMessages(displayName, msgs);
        return;
      }
    }

    const hasMore = !!this.messageState.nextLink;

    if (this.viewMode === "senders") {
      const passesFilter = (m: Message) => {
        const filtered = this.applyMessageFilters([m]);
        return filtered.length > 0;
      };
      if (this.filterUnreadOnly) {
        const filtered = this.senderGroups.filter((s) =>
          s.messages.some(passesFilter),
        );
        void this.messageList.renderSenders(filtered, hasMore, passesFilter);
      } else {
        void this.messageList.renderSenders(this.senderGroups, hasMore);
      }
      return;
    }

    // Flat messages mode: apply filters and sort.
    const filtered = this.applyMessageFilters(this.messageState.messages);
    const dir = this.sortNewestFirst ? -1 : 1;
    const sorted = [...filtered].sort(
      (a, b) =>
        dir *
        (new Date(a.receivedDateTime || 0).getTime() -
          new Date(b.receivedDateTime || 0).getTime()),
    );
    void this.messageList.renderFlatMessages(sorted, hasMore);
  }

  /**
   * Kick off all background AI processing (tagging, prefetch, detection).
   * Handles errors with user-facing notices instead of swallowing rejections.
   */
  private startBackgroundProcessing(): void {
    void this.generateAllNicknames();

    this.prefetchAllPromise = this.prefetchAllProcessed();
    this.prefetchAllPromise.catch((err) =>
      logger.warn("InboxView", "Background prefetch failed", err),
    );

    // Tag after prefetch so the tagger sees processed markdown
    // for every message within the prefetch window.
    (async () => {
      if (this.prefetchAllPromise) {
        try { await this.prefetchAllPromise; } catch { /* proceed anyway */ }
      }
      await this.classifier.autoTagAllMessages(
        this.messageState.messages,
        () => this.messageViewer.setTagCache(this.tagCache),
      );
    })().catch((err) => logger.warn("InboxView", "Auto-tagging failed", err));

    // When Claude processing is disabled but forwarded-sender resolution is
    // on, bodies are never prefetched by prefetchAllProcessed().  Fetch them
    // here so originalSender gets extracted and the sender list updates.
    const s = this.plugin.settings;
    if (s.resolveForwardedSender && (!s.enableClaudeProcessing || !hasClaudeAccess(s.anthropicApiKey))) {
      void this.prefetchBodiesForSenderResolution();
    }

    void this.autoDetectItems();
  }

  /** Return nickname if available, otherwise normalize "Last, First" to "First Last". */
  private resolveName(address: string, rawName: string): string {
    if (!address) return normalizeName(rawName);
    return this.nicknameCache.get(address.toLowerCase()) || normalizeName(rawName);
  }

  /** Open a modal to edit the nickname for an email address. */
  private openNicknameModal(address: string, rawName: string): void {
    const current = this.nicknameCache.get(address.toLowerCase()) || "";
    const s = this.plugin.settings;
    const canRegenerate = !!(s.enableClaudeProcessing && hasClaudeAccess(s.anthropicApiKey));
    const regenerate = canRegenerate
      ? async () => generateNickname(
          s.anthropicApiKey,
          s.claudeModel,
          NICKNAME_PROMPT,
          rawName,
          address,
        )
      : undefined;
    new NicknameModal(
      this.plugin.app,
      address,
      rawName,
      current,
      (addr, nickname) => {
        const key = addr.toLowerCase();
        if (nickname) {
          this.nicknameCache.set(key, nickname);
          this.plugin.store.setNickname(addr, nickname);
        } else {
          this.nicknameCache.delete(key);
          this.plugin.store.deleteNickname(addr);
        }
        this.regroupAndSync();
        this.renderCurrentView();
        this.messageViewer.refresh();
      },
      (addr) => {
        const key = addr.toLowerCase();
        this.nicknameCache.delete(key);
        this.plugin.store.deleteNickname(addr);
        this.regroupAndSync();
        this.renderCurrentView();
        this.messageViewer.refresh();
      },
      regenerate,
    ).open();
  }

  /**
   * Resolve the effective sender for a message.
   *
   * The "envelope sender" is the address on the message envelope itself
   * (preferring `msg.from` over the Graph API `msg.sender` delegate field).
   * When resolveForwardedSender is enabled and the body cache contains an
   * originalSender extracted from the forwarded body, that original sender
   * is returned instead, with the envelope sender demoted to `via*` fields.
   */
  private getEffectiveSender(msg: Message): {
    address: string;
    name: string;
    viaAddress?: string;
    viaName?: string;
  } {
    const envelope = getEnvelopeSender(msg);

    if (!this.plugin.settings.resolveForwardedSender) {
      return { address: envelope.address, name: envelope.name };
    }

    const cached = this.plugin.store.getBody(msg.id || "");
    if (cached) {
      let original = cached.originalSender;
      // Backfill: extract from bodies cached before this feature existed
      if (!original && /^(?:fw|fwd)\s*:/i.test(cached.subject) && cached.bodyHtml) {
        original = extractForwardedSender(cached.bodyHtml) ?? undefined;
        if (original) {
          cached.originalSender = original;
        }
      }
      if (original?.address) {
        return {
          address: original.address,
          name: original.name || original.address,
          viaAddress: envelope.address,
          viaName: envelope.name,
        };
      }
    }

    return { address: envelope.address, name: envelope.name };
  }

  /** Generate nicknames for all unique senders that don't have one yet. */
  private async generateAllNicknames(): Promise<void> {
    const s = this.plugin.settings;
    if (!s.enableClaudeProcessing || !hasClaudeAccess(s.anthropicApiKey)) return;

    // Collect unique effective-sender addresses + raw names
    const seen = new Map<string, string>();
    const addSeen = (addr: string, rawName: string) => {
      const key = addr.toLowerCase();
      if (!key || this.nicknameCache.has(key) || seen.has(key)) return;
      if (!rawName || rawName === key) return;
      if (this.plugin.store.isNicknameDeleted(key)) return;
      seen.set(key, rawName);
    };
    for (const msg of this.messageState.messages) {
      if (s.resolveForwardedSender) {
        const eff = this.getEffectiveSender(msg);
        // Skip "X via Y" senders -- the names from forwarded
        // institutional addresses are too noisy to auto-nickname.
        if (eff.viaName) continue;
        addSeen(eff.address, eff.name);
      } else {
        const envelope = getEnvelopeSender(msg);
        addSeen(envelope.address, envelope.name);
      }
    }

    if (seen.size === 0) return;

    const entries = Array.from(seen, ([address, rawName]) => ({ address, rawName }));
    const BATCH_SIZE = 10;
    const CONCURRENCY = 3;
    const batches: { address: string; rawName: string }[][] = [];
    for (let i = 0; i < entries.length; i += BATCH_SIZE) {
      batches.push(entries.slice(i, i + BATCH_SIZE));
    }

    let next = 0;
    const workers = Array.from({ length: Math.min(CONCURRENCY, batches.length) }, async () => {
      while (next < batches.length) {
        const batch = batches[next++];
        try {
          const map = await generateNicknamesBatch(
            s.anthropicApiKey,
            s.claudeModel,
            NICKNAME_BATCH_PROMPT,
            batch,
          );
          for (const e of batch) {
            const nickname = map.get(e.address.toLowerCase());
            if (!nickname) continue;
            // Re-check after the async gap -- the user may have deleted
            // or manually set a nickname while generation was in-flight.
            if (this.nicknameCache.has(e.address) || this.plugin.store.isNicknameDeleted(e.address)) continue;
            this.nicknameCache.set(e.address, nickname);
            this.plugin.store.setNickname(e.address, nickname);
          }
        } catch (err) {
          logger.warn("InboxView", "Nickname batch failed", err);
        }
      }
    });
    await Promise.all(workers);

    // Re-render to show nicknames
    this.regroupAndSync();
    this.renderCurrentView();
  }

  private handleSenderSelect(sender: SenderGroup): void {
    this.activeSender = sender;
    this.messageViewer.clear();
    const displayName = this.resolveName(
      sender.address,
      sender.name || sender.address,
    );
    const msgs = this.filterSenderMessages(sender.messages);
    void this.messageList.renderSenderMessages(displayName, msgs);

    // Auto-select the latest message
    if (msgs.length > 0) {
      void this.showMessageInViewer(msgs[msgs.length - 1]);
    }
  }

  /** Apply active toggle filters (unread) to a message list. */
  private applyMessageFilters(messages: Message[]): Message[] {
    let filtered = messages;
    if (this.filterUnreadOnly) {
      filtered = filtered.filter((m) => !m.isRead);
    }
    return filtered;
  }

  /** Apply active filters to a sender's message list. */
  private filterSenderMessages(messages: Message[]): Message[] {
    const filtered = this.applyMessageFilters(messages);
    const dir = this.sortNewestFirst ? -1 : 1;
    filtered.sort(
      (a, b) =>
        dir *
        (new Date(a.receivedDateTime || 0).getTime() -
          new Date(b.receivedDateTime || 0).getTime()),
    );
    return filtered;
  }

  // --- Private: tag classification ---

  private getSelectedMessage(): Message | null {
    if (!this.selectedMessageId) return null;
    return (
      this.messageState.messages.find((m) => m.id === this.selectedMessageId) ||
      null
    );
  }

  private rebuildTagWrap(): void {
    this.tagWrap.empty();

    const categories = this.getTagCategories();
    const icons = this.plugin.settings.tagIcons || {};

    // Lead icon (always visible)
    const leadBtn = this.tagWrap.createEl("button", {
      cls: "iris-topbar-btn clickable-icon",
      attr: { "aria-label": "Tags" },
    });
    setIcon(leadBtn, "tag");

    // One button per existing tag (hidden, revealed on hover)
    for (const cat of categories) {
      const wrap = this.tagWrap.createDiv({ cls: "iris-tag-icon-wrap" });
      const btn = wrap.createEl("button", {
        cls: "iris-filter-opt clickable-icon",
        attr: { "aria-label": cat },
      });
      setIcon(btn, icons[cat] || "tag");
      wrap.createSpan({ cls: "iris-tag-icon-label", text: cat });
      btn.addEventListener("click", () => {
        const msg = this.getSelectedMessage();
        if (msg) this.handleTagChange(msg, cat);
      });
      btn.addEventListener("contextmenu", (e) => {
        e.preventDefault();
        this.openTagEditModal(cat);
      });
    }

    // Plus button (also hidden, revealed on hover)
    const addBtn = this.tagWrap.createEl("button", {
      cls: "iris-filter-opt clickable-icon",
      attr: { "aria-label": "Add tag" },
    });
    setIcon(addBtn, "plus");
    addBtn.addEventListener("click", () =>
      this.showAddTagInput(this.tagWrap),
    );
  }

  private showAddTagInput(_anchor: HTMLElement): void {
    const s = this.plugin.settings;
    const categories = this.getTagCategories();
    const canGenerate = s.enableClaudeProcessing && hasClaudeAccess(s.anthropicApiKey);

    new CreateTagModal(this.plugin.app, {
      existingTags: categories,
      onGenerate: canGenerate
        ? (name) => generateTagDescription(
            s.anthropicApiKey,
            s.claudeModel,
            name,
            this.getTagCandidates(),
          )
        : undefined,
      onSubmit: async (name, criteria, icon, iconExplicit) => {
        const updated = [...categories, name];
        this.plugin.settings.tagCategories = updated.join(", ");
        if (!this.plugin.settings.tagIcons) this.plugin.settings.tagIcons = {};
        if (!this.plugin.settings.tagDescriptions) this.plugin.settings.tagDescriptions = {};
        this.plugin.settings.tagDescriptions[name] = criteria;

        const usedIcons = Object.values(this.plugin.settings.tagIcons);
        this.plugin.settings.tagIcons[name] = icon;
        void this.plugin.saveSettings();
        this.messageViewer.setTagCategories(updated);
        this.messageViewer.setTagIcons(this.getTagIconMap());
        this.rebuildTagWrap();

        if (!iconExplicit && canGenerate) {
          try {
            const picked = await pickTagIcon(
              s.anthropicApiKey, s.claudeModel, name, criteria,
              TAG_ICON_POOL, usedIcons,
            );
            if (picked) {
              this.plugin.settings.tagIcons[name] = picked;
              void this.plugin.saveSettings();
              this.messageViewer.setTagIcons(this.getTagIconMap());
              this.rebuildTagWrap();
            }
          } catch (err) {
            logger.warn("InboxView", "AI icon pick failed; kept fallback", err);
          }
        }
      },
    }).open();
  }

  private openTagEditModal(cat: string): void {
    const s = this.plugin.settings;
    const canGenerate = s.enableClaudeProcessing && hasClaudeAccess(s.anthropicApiKey);
    new CreateTagModal(this.plugin.app, {
      existingTags: this.getTagCategories(),
      initial: {
        name: cat,
        criteria: s.tagDescriptions?.[cat] || "",
        icon: s.tagIcons?.[cat] || "tag",
      },
      onGenerate: canGenerate
        ? (name) => generateTagDescription(
            s.anthropicApiKey, s.claudeModel, name,
            this.getTagCandidates().filter((t) => t.name !== cat),
          )
        : undefined,
      onSubmit: (_name, criteria, icon) => {
        const criteriaChanged = (s.tagDescriptions?.[cat] || "") !== criteria;
        if (!s.tagDescriptions) s.tagDescriptions = {};
        if (!s.tagIcons) s.tagIcons = {};
        s.tagDescriptions[cat] = criteria;
        s.tagIcons[cat] = icon;
        if (criteriaChanged) {
          if (!s.tagPromptVersions) s.tagPromptVersions = {};
          bumpTagVersion(s.tagPromptVersions, cat);
        }
        void this.plugin.saveSettings();
        this.messageViewer.setTagIcons(this.getTagIconMap());
        this.rebuildTagWrap();
      },
    }).open();
  }

  private getTagCategories(): string[] {
    return parseTagCategories(this.plugin.settings.tagCategories);
  }

  private getTagCandidates(): TagCandidate[] {
    const descriptions = this.plugin.settings.tagDescriptions || {};
    return this.getTagCategories().map((name) => ({
      name,
      description: descriptions[name] || "",
    }));
  }

  private getTagIconMap(): Map<string, string> {
    const icons = this.plugin.settings.tagIcons || {};
    return new Map(Object.entries(icons));
  }

  private getClassifiableContent(msg: Message): string {
    if (!msg.id) return [msg.subject, msg.bodyPreview].filter(Boolean).join("\n");
    const processed = this.plugin.store.getProcessed(msg.id);
    if (processed?.processedMarkdown) {
      return [msg.subject, processed.processedMarkdown].filter(Boolean).join("\n");
    }
    const body = this.plugin.store.getBody(msg.id);
    if (body?.strippedHtml) {
      return [msg.subject, body.strippedHtml].filter(Boolean).join("\n");
    }
    return [msg.subject, msg.bodyPreview].filter(Boolean).join("\n");
  }

  private getEffectiveTagPrompt(): string {
    return this.plugin.settings.tagClassifyPrompt || TAG_CLASSIFY_PROMPT;
  }

  private handleTagChange(msg: Message, tag: string | null): void {
    if (!msg.id) return;

    const existing = this.tagCache.get(msg.id) || [];
    const removedAutoTags: string[] = [];

    if (tag === null) {
      // Remove all tags — track auto-tags for prompt refinement
      for (const e of existing) {
        if (e.source === "auto") removedAutoTags.push(e.tag);
      }
      this.tagCache.delete(msg.id);
      this.plugin.store.removeTag(msg.id);
    } else {
      // Toggle: if tag already present, remove it; otherwise add it
      const has = existing.find((e) => e.tag === tag);
      if (has) {
        if (has.source === "auto") removedAutoTags.push(tag);
        const updated = existing.filter((e) => e.tag !== tag);
        if (updated.length === 0) {
          this.tagCache.delete(msg.id);
        } else {
          this.tagCache.set(msg.id, updated);
        }
        this.plugin.store.removeTag(msg.id, tag);
      } else {
        const entry: TagCacheEntry = {
          messageId: msg.id,
          tag,
          source: "manual",
          taggedAt: Date.now(),
        };
        this.tagCache.set(msg.id, [...existing, entry]);
        this.plugin.store.setTag(msg.id, tag, "manual");
      }
    }

    this.messageViewer.setTagCache(this.tagCache);
    this.renderSelectedMessage(msg);

    // Refine prompt for any removed auto-tags
    for (const removed of removedAutoTags) {
      void this.refinePromptWithFeedback(msg, removed, "incorrect");
    }
  }

  /** Save refined criteria for a single tag, bump only its version, notify, and sync. */
  private applyRefinedTagCriteria(tag: string, refined: string, noticeLabel: string): void {
    const s = this.plugin.settings;
    if (!s.tagDescriptions) s.tagDescriptions = {};
    if (!s.tagPromptVersions) s.tagPromptVersions = {};
    s.tagDescriptions[tag] = refined;
    bumpTagVersion(s.tagPromptVersions, tag);
    void this.plugin.saveSettings();
    new Notice(noticeLabel);
    this.syncPromptVersions();
  }

  /** Bump every defined tag's prompt version — used when the global meta-prompt changes. */
  private bumpAllTagVersions(): void {
    const s = this.plugin.settings;
    if (!s.tagPromptVersions) s.tagPromptVersions = {};
    for (const tag of this.getTagCategories()) {
      bumpTagVersion(s.tagPromptVersions, tag);
    }
  }

  private async refinePromptWithFeedback(
    msg: Message,
    tag: string,
    feedback: "correct" | "incorrect",
  ): Promise<void> {
    const s = this.plugin.settings;
    if (!hasClaudeAccess(s.anthropicApiKey)) return;

    const content = this.getClassifiableContent(msg);
    if (!content) return;

    try {
      const refined = await refineTagCriteria(
        s.anthropicApiKey,
        tag,
        s.tagDescriptions?.[tag] || "",
        content,
        feedback,
      );
      this.applyRefinedTagCriteria(tag, refined, `Criteria for "${tag}" changed`);
    } catch (err) {
      logger.warn("InboxView", "Criteria refinement failed", err);
    }
  }

  // autoTagAllMessages is now handled by EmailClassifier

  /** Pre-process emails with Claude in the background so results are cached before the user clicks. */
  private async prefetchAllProcessed(): Promise<void> {
    const s = this.plugin.settings;
    if (!s.enableClaudeProcessing || !hasClaudeAccess(s.anthropicApiKey)) return;

    const limit = s.prefetchLimit;
    if (limit === 0) return;

    const gen = this.prefetchGeneration;

    const effectivePrompt = s.claudeSystemPrompt || DEFAULT_CLAUDE_PROMPT;
    const promptHash = EmailStore.hashPrompt(effectivePrompt);
    const cache = this.plugin.store;

    const candidates = limit === -1
      ? this.messageState.messages
      : this.messageState.messages.slice(0, limit);

    // ── Phase 1: Fetch all bodies ────────────────────────────
    // Collect body data for each candidate before any Claude processing so
    // that originalSender resolution can update the sender list early.
    const bodyMap = new Map<string, string>(); // msgId → strippedHtml
    const skipSet = new Set<string>(); // messages that need no processing

    for (const msg of candidates) {
      if (this.prefetchGeneration !== gen) return;

      const msgId = msg.id;
      if (!msgId) continue;

      // Skip if already in L1
      if (this.processedCache.has(msgId)) {
        skipSet.add(msgId);
        continue;
      }

      // Skip if already on disk (even with stale prompt hash) — warm L1 from L2
      if (cache.hasProcessed(msgId)) {
        const entry = cache.getProcessed(msgId)!;
        this.processedCache.set(msgId, entry.processedMarkdown);
        skipSet.add(msgId);
        continue;
      }

      // --- Fetch body: L2 disk → L3 Graph API ---
      try {
        const diskBody = cache.getBody(msgId);
        if (diskBody) {
          bodyMap.set(msgId, diskBody.strippedHtml);
        } else {
          const fullMsg = await this.plugin.mailApi.getMessageBody(msgId);
          if (this.prefetchGeneration !== gen) return;
          const bodyHtml = fullMsg.body?.content || "";
          if (bodyHtml) {
            const entry = cache.setBody(msg, bodyHtml);
            bodyMap.set(msgId, entry.strippedHtml);
          }
        }
      } catch (err) {
        logger.warn("InboxView", `Prefetch body fetch failed for ${msgId}`, err);
      }
    }

    // Bodies are now cached with originalSender extracted — regroup so the
    // sender list reflects effective senders before Claude processing begins.
    if (this.prefetchGeneration === gen) {
      this.regroupAndSync();
    }

    // ── Phase 2: Claude processing ───────────────────────────
    for (const msg of candidates) {
      if (this.prefetchGeneration !== gen) return;

      const msgId = msg.id;
      if (!msgId || skipSet.has(msgId)) continue;

      const stripped = bodyMap.get(msgId);
      if (!stripped) continue;

      const parsedBody = htmlToMarkdown(stripped);
      if (!parsedBody) continue;

      if (this.prefetchGeneration !== gen) return;

      // Prepend email context (subject, sender, date) for full context
      const parsedContent = this.buildEmailContext(msg) + parsedBody;

      // Full extraction
      this.prefetchInflight.add(msgId);
      try {
        const markdown = await processEmailWithClaude(
          s.anthropicApiKey,
          s.claudeModel,
          effectivePrompt,
          parsedContent,
        );
        if (this.prefetchGeneration !== gen) return;

        this.processedCache.set(msgId, markdown);

        try {
          await cache.setProcessed(msgId, markdown, promptHash);
        } catch (err) {
          logger.warn("InboxView", `Prefetch save failed for ${msgId}`, err);
        }
      } catch (err) {
        logger.warn("InboxView", `Prefetch Claude processing failed for ${msgId}`, err);
      } finally {
        this.prefetchInflight.delete(msgId);
      }
    }
  }

  /**
   * Fetch bodies for all inbox messages so that originalSender can be
   * extracted for forwarded-sender resolution.  Used when Claude processing
   * is disabled but resolveForwardedSender is enabled.
   */
  private async prefetchBodiesForSenderResolution(): Promise<void> {
    const gen = this.prefetchGeneration;
    const cache = this.plugin.store;

    let fetched = false;
    for (const msg of this.messageState.messages) {
      if (this.prefetchGeneration !== gen) return;
      const msgId = msg.id;
      if (!msgId) continue;

      // Already cached — nothing to do
      if (cache.getBody(msgId)) continue;

      // Only bother fetching forwarded messages
      if (!/^(?:fw|fwd)\s*:/i.test(msg.subject || "")) continue;

      try {
        const fullMsg = await this.plugin.mailApi.getMessageBody(msgId);
        if (this.prefetchGeneration !== gen) return;
        const bodyHtml = fullMsg.body?.content || "";
        if (bodyHtml) {
          cache.setBody(msg, bodyHtml);
          fetched = true;
        }
      } catch (err) {
        logger.warn("InboxView", `Body fetch for sender resolution failed for ${msgId}`, err);
      }
    }

    if (fetched && this.prefetchGeneration === gen) {
      this.regroupAndSync();
    }
  }

  // ── Auto-detection of events and tasks ─────────────────────

  private async autoDetectItems(): Promise<void> {
    const s = this.plugin.settings;
    if (!s.enableAutoItemDetection || !s.enableClaudeProcessing || !hasClaudeAccess(s.anthropicApiKey)) return;

    const gen = this.prefetchGeneration;

    // Wait for prefetch so email bodies are cached
    if (this.prefetchAllPromise) {
      try { await this.prefetchAllPromise; } catch { /* proceed anyway */ }
    }
    if (this.prefetchGeneration !== gen) return;

    const cache = this.plugin.store;
    const effectivePrompt = s.itemDetectionPrompt || ITEM_DETECTION_PROMPT;

    const limit = s.prefetchLimit;
    const candidates = limit === -1
      ? this.messageState.messages
      : limit === 0
        ? []
        : this.messageState.messages.slice(0, limit);

    for (const msg of candidates) {
      if (this.prefetchGeneration !== gen) return;

      const msgId = msg.id;
      if (!msgId) continue;

      // Skip if already scanned
      if (cache.hasItemsScan(msgId)) {
        // Warm L1 from L2
        if (!this.detectedItemsCache.has(msgId)) {
          const items = cache.getDetectedItems(msgId);
          if (items.length > 0) this.detectedItemsCache.set(msgId, items);
        }
        continue;
      }

      // Only use Claude-processed markdown — skip if not yet processed.
      const processed = cache.getProcessed(msgId);
      if (!processed) continue;
      const parsed = processed.processedMarkdown;
      if (!parsed) continue;

      const account = this.plugin.authProvider.getAccount();
      const emailContext = {
        subject: msg.subject || "",
        from: msg.from?.emailAddress?.name || msg.from?.emailAddress?.address || "",
        date: msg.receivedDateTime || "",
        userName: account?.name || account?.username || "",
      };

      try {
        const detected = await detectItemsInEmail(
          s.anthropicApiKey, s.claudeModel, effectivePrompt, parsed, emailContext,
        );
        if (this.prefetchGeneration !== gen) return;

        cache.setItemsScanned(msgId);

        this.storeDetectedItems(msgId, detected);
      } catch (err) {
        logger.warn("InboxView", `Item detection failed for ${msgId}`, err);
      }
    }
  }

  /** Run item detection for a single message when the user opens it (on-demand). */
  /** Run item detection for a single message when the user opens it.
   *  Works whenever Claude processing is enabled, regardless of the
   *  auto-detection background toggle. */
  private async detectItemsOnDemand(msg: Message): Promise<void> {
    const s = this.plugin.settings;
    if (!s.enableClaudeProcessing || !hasClaudeAccess(s.anthropicApiKey)) return;

    const msgId = msg.id;
    if (!msgId) return;

    const cache = this.plugin.store;

    // Only use Claude-processed markdown — skip if not yet available
    const processed = cache.getProcessed(msgId);
    if (!processed) return;
    const content = processed.processedMarkdown;
    if (!content) return;

    const effectivePrompt = s.itemDetectionPrompt || ITEM_DETECTION_PROMPT;
    const account = this.plugin.authProvider.getAccount();
    const emailContext = {
      subject: msg.subject || "",
      from: msg.from?.emailAddress?.name || msg.from?.emailAddress?.address || "",
      date: msg.receivedDateTime || "",
      userName: account?.name || account?.username || "",
    };

    try {
      const detected = await detectItemsInEmail(
        s.anthropicApiKey, s.claudeModel, effectivePrompt, content, emailContext,
      );

      logger.debug("InboxView", `Detection returned ${detected.length} items for ${msgId}`);
      cache.setItemsScanned(msgId);

      this.storeDetectedItems(msgId, detected);
    } catch (err) {
      logger.warn("InboxView", `On-demand item detection failed for ${msgId}`, err);
    }
  }


  /** Convert DetectedItem[] to DetectedItemEntry[], persist, update caches, and refresh viewer. */
  private storeDetectedItems(msgId: string, detected: import("../utils/claudeApi").DetectedItem[]): void {
    const cache = this.plugin.store;
    if (detected.length > 0) {
      const now = Date.now();
      const entries: DetectedItemEntry[] = detected.map((d, i) => ({
        itemId: `${msgId}-${i}`,
        messageId: msgId,
        type: d.type,
        title: d.title,
        date: d.date,
        time: d.time,
        location: d.location,
        dueDate: d.dueDate,
        priority: d.priority,
        description: d.description,
        sourceText: d.sourceText,
        status: "pending" as const,
        detectedAt: now,
      }));
      cache.setDetectedItems(msgId, entries);
      this.detectedItemsCache.set(msgId, entries);

      if (msgId === this.selectedMessageId) {
        this.messageViewer.setDetectedItems(entries);
        this.messageViewer.refresh();
      }
    } else {
      cache.setDetectedItems(msgId, []);
    }
  }

  private async handleAcceptDetectedItem(messageId: string, item: DetectedItemEntry): Promise<void> {
    const s = this.plugin.settings;
    const msg = this.messageState.messages.find((m) => m.id === messageId);

    // Build an ExtractedNote from the detected item
    let note: ExtractedNote;
    if (item.type === "event") {
      note = {
        type: "event",
        title: item.title,
        date: item.date || "",
        time: item.time || "",
        location: item.location || "",
        description: item.description,
      };
    } else {
      note = {
        type: "task",
        title: item.title,
        dueDate: item.dueDate || "",
        description: item.description,
      };
    }

    try {
      const filePath = await this.saveExtractedNote(note, msg || { id: messageId } as Message);
      this.plugin.store.updateDetectedItemStatus(messageId, item.itemId, "accepted", filePath);

      // Update L1 cache
      const items = this.detectedItemsCache.get(messageId);
      if (items) {
        const entry = items.find((i) => i.itemId === item.itemId);
        if (entry) {
          entry.status = "accepted";
          entry.vaultPath = filePath;
          entry.resolvedAt = Date.now();
        }
      }

      // Refresh viewer if showing this message
      if (messageId === this.selectedMessageId) {
        this.messageViewer.setDetectedItems(items || []);
        this.messageViewer.refresh();
      }
      new Notice(`${item.type === "event" ? "Event" : "Task"} note created: ${filePath}`);
    } catch (err) {
      const errMsg = err instanceof Error ? err.message : String(err);
      new Notice(`Failed to create note: ${errMsg}`);
    }
  }

  private handleDismissDetectedItem(messageId: string, itemId: string): void {
    this.plugin.store.updateDetectedItemStatus(messageId, itemId, "dismissed");

    const items = this.detectedItemsCache.get(messageId);
    if (items) {
      const entry = items.find((i) => i.itemId === itemId);
      if (entry) {
        entry.status = "dismissed";
        entry.resolvedAt = Date.now();
      }
    }

    if (messageId === this.selectedMessageId) {
      this.messageViewer.setDetectedItems(items || []);
      this.messageViewer.refresh();
    }
  }

  private handleUpdateDetectedItem(messageId: string, itemId: string, updates: Partial<DetectedItemEntry>): void {
    const items = this.detectedItemsCache.get(messageId);
    if (!items) return;
    const entry = items.find((i) => i.itemId === itemId);
    if (!entry) return;

    Object.assign(entry, updates);
    this.plugin.store.setDetectedItems(messageId, items);

    if (messageId === this.selectedMessageId) {
      this.messageViewer.setDetectedItems(items);
      this.messageViewer.refresh();
    }
  }

  private async handleReloadDetectedItems(messageId: string): Promise<void> {
    // Clear existing scan so on-demand detection re-runs
    this.plugin.store.clearItemsScan(messageId);
    this.detectedItemsCache.delete(messageId);

    // Re-run detection
    const msg = this.messageState.messages.find((m) => m.id === messageId);
    if (!msg) return;

    const cache = this.plugin.store;
    const body = cache.getBody(messageId);
    const stripped = body ? body.strippedHtml : "";

    if (messageId === this.selectedMessageId) {
      this.messageViewer.setDetectedItems([]);
      this.messageViewer.refresh();
    }

    await this.detectItemsOnDemand(msg);
  }

  /** Push current prompt versions to MessageViewer and re-render the selected message. */
  private syncPromptVersions(): void {
    this.messageViewer.setPromptVersions(
      this.plugin.settings.tagPromptVersions || {},
    );
    if (this.selectedMessageId) {
      const msg = this.messageState.messages.find((m) => m.id === this.selectedMessageId);
      if (msg) this.renderSelectedMessage(msg);
    }
  }

  /** Re-tag a single message with the current tag prompt (replacing obsolete auto-tags). */
  private async handleRetagMessage(msg: Message): Promise<void> {
    if (!msg.id) return;
    const s = this.plugin.settings;
    if (!s.enableClaudeProcessing || !hasClaudeAccess(s.anthropicApiKey)) return;

    const candidates = this.getTagCandidates();
    if (candidates.length === 0) return;

    const content = this.getClassifiableContent(msg);
    if (!content) return;

    // Remove old auto-tags, keep manual
    const existing = this.tagCache.get(msg.id) || [];
    const manual = existing.filter((e) => e.source === "manual");
    for (const e of existing) {
      if (e.source === "auto") this.plugin.store.removeTag(msg.id, e.tag);
    }

    try {
      const tags = await classifyEmailTagsYesNo(
        s.anthropicApiKey,
        s.claudeModel,
        this.getEffectiveTagPrompt(),
        content,
        candidates,
      );

      const newEntries: TagCacheEntry[] = tags.map((tag) => ({
        messageId: msg.id!,
        tag,
        source: "auto" as const,
        promptVersion: getTagVersion(s.tagPromptVersions, tag),
        taggedAt: Date.now(),
      }));
      const merged = [...manual, ...newEntries];
      if (merged.length > 0) {
        this.tagCache.set(msg.id, merged);
      } else {
        this.tagCache.delete(msg.id);
      }
      for (const tag of tags) {
        this.plugin.store.setTag(msg.id, tag, "auto", getTagVersion(s.tagPromptVersions, tag));
      }
    } catch (err) {
      // On failure, keep manual tags only
      if (manual.length > 0) {
        this.tagCache.set(msg.id, manual);
      } else {
        this.tagCache.delete(msg.id);
      }
      logger.warn("InboxView", "Re-tag failed", err);
    }

    this.messageViewer.setTagCache(this.tagCache);
    this.renderSelectedMessage(msg);
  }

  /** Re-process the current message with the current prompt, bypassing cache. */
  private handleReprocessMessage(msg: Message): void {
    if (!msg.id) return;
    // Evict from L1 so showMessageInViewer falls through to L3 (Claude API)
    this.processedCache.delete(msg.id);
    // Evict from L2 disk cache so the hash check doesn't short-circuit
    this.plugin.store.clearProcessed(msg.id);
    // Re-run the full viewer flow
    void this.showMessageInViewer(msg);
  }

  // --- Private: helpers ---

  private groupByEffectiveSender(messages: Message[]): SenderGroup[] {
    interface GroupInfo { address: string; name: string; messages: Message[] }
    const groups = new Map<string, GroupInfo>();

    for (const msg of messages) {
      const eff = this.getEffectiveSender(msg);
      const addr = (eff.address || "unknown").toLowerCase();
      const name = eff.name || addr;
      // When different people send via the same institutional address,
      // include the sender name in the key to keep them separate.
      const groupKey = eff.viaName
        ? `${addr}::${name.toLowerCase()}`
        : addr;
      const existing = groups.get(groupKey);
      if (existing) {
        existing.messages.push(msg);
      } else {
        groups.set(groupKey, { address: addr, name, messages: [msg] });
      }
    }

    const senders: SenderGroup[] = [];
    for (const [groupKey, group] of groups) {
      group.messages.sort(
        (a, b) =>
          new Date(a.receivedDateTime || 0).getTime() -
          new Date(b.receivedDateTime || 0).getTime(),
      );

      const latestMessage = group.messages[group.messages.length - 1];

      senders.push({
        groupKey,
        address: group.address,
        name: group.name,
        messages: group.messages,
        latestMessage,
        unreadCount: group.messages.filter((m) => !m.isRead).length,
      });
    }

    const dir = this.sortNewestFirst ? -1 : 1;
    senders.sort(
      (a, b) =>
        dir * (new Date(a.latestMessage.receivedDateTime || 0).getTime() -
        new Date(b.latestMessage.receivedDateTime || 0).getTime()),
    );

    return senders;
  }

  syncBadge(): void {
    const mode = this.plugin.settings.badgeCount;
    if (mode === "off") {
      this.plugin.updateBadge(0);
      return;
    }

    // Apply the same filters the view uses so the badge matches what's shown
    const msgs = this.applyMessageFilters(this.messageState.messages);

    switch (mode) {
      case "unread":
        this.plugin.updateBadge(msgs.filter((m) => !m.isRead).length);
        break;
      case "total":
        this.plugin.updateBadge(msgs.length);
        break;
    }
  }

  /** Build an email context header (subject, sender, date) to prepend to body
   *  content before sending to Claude, so the model has full context even when
   *  the HTML body is sparse (e.g. meeting invitations, calendar events). */
  private buildEmailContext(msg: Message): string {
    const lines: string[] = [];
    if (msg.subject) lines.push(`Subject: ${msg.subject}`);
    const eff = this.getEffectiveSender(msg);
    const from = this.resolveName(eff.address, eff.name);
    if (from) lines.push(`From: ${from}`);
    if (msg.receivedDateTime) {
      const dt = new Date(msg.receivedDateTime);
      lines.push(`Date: ${dt.toISOString().split("T")[0]} ${dt.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" })}`);
    }
    return lines.length > 0 ? lines.join("\n") + "\n\n" : "";
  }

  // ── Note creation from selected email text ────────────────────

  private async handleCreateNoteFromSelection(
    selectedText: string,
    noteType: NoteType,
    msg: Message,
  ): Promise<void> {
    const s = this.plugin.settings;
    if (!hasClaudeAccess(s.anthropicApiKey)) {
      new Notice("Please configure an Anthropic API key in Iris Mail settings or enable the Iris Router plugin.");
      return;
    }

    const label = noteType === "event" ? "event" : "task";
    new Notice(`Extracting ${label} details\u2026`);

    try {
      const emailContext = {
        subject: msg.subject || "",
        from: msg.from?.emailAddress?.name || msg.from?.emailAddress?.address || "",
        date: msg.receivedDateTime || "",
      };

      const extracted = await extractNoteFromSelection(
        s.anthropicApiKey, s.claudeModel, selectedText, emailContext, noteType,
      );

      const filePath = await this.saveExtractedNote(extracted, msg);
      new Notice(`${noteType === "event" ? "Event" : "Task"} note created: ${filePath}`);
    } catch (err) {
      const errMsg = err instanceof Error ? err.message : String(err);
      new Notice(`Failed to extract ${label}: ${errMsg}`);
    }
  }

  private async saveExtractedNote(note: ExtractedNote, msg: Message): Promise<string> {
    const title = note.title
      .replace(/[\\/:*?"<>|]/g, "")
      .replace(/\s+/g, " ")
      .trim()
      .slice(0, 50) || "Untitled";

    const s = this.plugin.settings;
    const folder = note.type === "event" ? (s.eventNoteFolderPath || "Events") : (s.taskNoteFolderPath || "Tasks");
    const fromAddr = msg.from?.emailAddress?.address || "";
    const fromRaw = msg.from?.emailAddress?.name || fromAddr;
    const from = this.resolveName(fromAddr, fromRaw);
    let frontmatterLines: string[];
    let body: string;

    /** Convert YYYY-MM-DD to DD-MM-YYYY. Returns input unchanged if not valid. */
    const flip = (d: string): string => {
      const m = d.match(/^(\d{4})-(\d{2})-(\d{2})$/);
      return m ? `${m[3]}-${m[2]}-${m[1]}` : d;
    };

    /** Build the date frontmatter line(s) for a key. Single → `key: YYYY-MM-DD`, range → `key?: "Between DD-MM-YYYY and DD-MM-YYYY"`. */
    const dateLine = (key: string, raw: string | undefined): string[] => {
      if (!raw) return [];
      const parts = raw.split("/");
      if (parts.length === 2 && parts[0] && parts[1]) {
        return [`${key}?: "Between ${flip(parts[0])} and ${flip(parts[1])}"`];
      }
      return [`${key}: ${parts[0]}`];
    };

    if (note.type === "event") {
      frontmatterLines = [
        "---",
        ...dateLine("date", note.date),
        ...(note.time ? [`time: "${note.time}"`] : []),
        ...(note.location ? [`location: "${note.location.replace(/"/g, '\\"')}"`] : []),
        `from: "${from.replace(/"/g, '\\"')}"`,
        "---",
      ];
      body = note.description;
    } else {
      frontmatterLines = [
        "---",
        ...dateLine("closes", note.dueDate),
        "---",
      ];
      body = `- [ ] ${note.description}`;
    }

    const content = frontmatterLines.join("\n") + "\n\n" + body + "\n";

    // Ensure folder exists
    if (!(await this.plugin.app.vault.adapter.exists(folder))) {
      await this.plugin.app.vault.createFolder(folder);
    }

    // Determine file path, avoiding collisions
    let filePath = `${folder}/${title}.md`;
    if (await this.plugin.app.vault.adapter.exists(filePath)) {
      const hash = EmailStore.hashPrompt(msg.id || Date.now().toString()).slice(0, 5);
      filePath = `${folder}/${title} ${hash}.md`;
    }

    await this.plugin.app.vault.create(filePath, content);
    return filePath;
  }
}
