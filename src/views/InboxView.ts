import { ItemView, WorkspaceLeaf, Notice, setIcon, Menu } from "obsidian";
import type IrisMailPlugin from "../main";
import {
  VIEW_TYPE_IRIS_MAIL,
  DEFAULT_CLAUDE_PROMPT,
  IMPORTANCE_CLASSIFY_PROMPT,
  NICKNAME_PROMPT,
  TAG_CLASSIFY_PROMPT,
  TAG_ICON_CYCLE,
  ITEM_DETECTION_PROMPT,
  parseTagCategories,
} from "../constants";
import { MessageList } from "./components/MessageList";
import { MessageViewer } from "./components/MessageViewer";
import { NicknameModal } from "./components/NicknameModal";
import { SearchBar } from "./components/SearchBar";
import { Toolbar } from "./components/Toolbar";
import { processEmailWithClaude, classifyEmailTags, refineTagPrompt, refineImportancePrompt, generateNickname, mergeEmailsToFormula, refineTagPromptBulk, refineImportancePromptBulk, extractNoteFromSelection, detectItemsInEmail } from "../utils/claudeApi";
import type { NoteType, ExtractedNote } from "../utils/claudeApi";
import { htmlToMarkdown } from "../utils/htmlToMarkdown";
import { extractForwardedSender } from "../utils/extractForwardedSender";
import { getEnvelopeSender } from "../utils/envelopeSender";
import { logger } from "../utils/logger";
import { EmailStore } from "../store/EmailStore";
import { EmailClassifier } from "../services/EmailClassifier";
import type { ImportanceClass, TagCacheEntry, DetectedItemEntry } from "../store/types";
import type {
  Message,
  MessageListState,
  ConversationGroup,
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

  private inboxFolderId: string | null = null;
  private messageState: MessageListState = {
    messages: [],
    conversations: [],
    nextLink: null,
    isLoading: false,
    selectedConversationId: null,
    searchQuery: "",
  };

  // View mode: senders (default) or conversations
  private viewMode!: "conversations" | "senders";
  private senderGroups: SenderGroup[] = [];
  private activeSender: SenderGroup | null = null;
  private viewModeToggleBtn!: HTMLButtonElement;
  private sortNewestFirst!: boolean;
  private sortToggleBtn!: HTMLButtonElement;
  private filterWrap!: HTMLDivElement;
  private filterUnreadOnly!: boolean;
  private filterHideNoise!: boolean;
  private filterImportantOnly!: boolean;
  private unreadOptBtn!: HTMLButtonElement;
  private noiseOptBtn!: HTMLButtonElement;
  private importantOptBtn!: HTMLButtonElement;

  // Currently drilled-into conversation
  private activeConversation: ConversationGroup | null = null;
  private activeConversationMessages: Message[] = [];
  private selectedMessageId: string | null = null;
  private lastStrippedHtml: string = "";

  // Extracted classifier handles classification & tagging caches
  private classifier!: EmailClassifier;

  // In-memory caches
  private conversationBodyCache = new Map<string, Message[]>();
  private processedCache = new Map<string, string>();
  private nicknameCache = new Map<string, string>();
  private detectedItemsCache = new Map<string, DetectedItemEntry[]>();

  // Convenience accessors for classifier caches
  private get classificationCache() { return this.classifier.classifications; }
  private get classificationSourceCache() { return this.classifier.classificationSources; }
  private get classificationVersionCache() { return this.classifier.classificationVersions; }
  private get tagCache() { return this.classifier.tags; }
  private tagWrap!: HTMLDivElement;
  private topBar!: HTMLDivElement;

  // Prefetch state
  private prefetchGeneration = 0;
  private classifyAllPromise: Promise<void> | null = null;
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
    this.filterHideNoise = s.filterHideNoise;
    this.filterImportantOnly = s.filterImportantOnly;
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
    this.contentEl.empty();
  }

  async refresh(): Promise<void> {
    this.prefetchGeneration++;
    this.conversationBodyCache.clear();
    this.processedCache.clear();
    this.activeConversation = null;
    this.activeConversationMessages = [];

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
    browserBtn.addEventListener("click", () =>
      this.plugin.handleLoginWithAuthCode(),
    );

    const deviceBtn = btnGroup.createEl("button", {
      text: "Sign in with device code",
    });
    deviceBtn.addEventListener("click", () =>
      this.plugin.handleLoginWithDeviceCode(),
    );
  }

  private renderInboxUI(container: HTMLElement): void {
    // Top bar: senders toggle (left) + search & refresh (right)
    const topBar = container.createDiv({ cls: "iris-topbar" });

    // Senders view toggle
    this.viewModeToggleBtn = topBar.createEl("button", {
      cls:
        "iris-topbar-btn clickable-icon" +
        (this.viewMode === "senders" ? " is-active" : ""),
      attr: { "aria-label": "Senders view" },
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

    this.noiseOptBtn = this.createFilterButton(
      this.filterWrap, "volume-off", "Hide noise",
      () => this.filterHideNoise,
      () => { this.filterHideNoise = !this.filterHideNoise; },
    );

    this.importantOptBtn = this.createFilterButton(
      this.filterWrap, "circle-alert", "Important only",
      () => this.filterImportantOnly,
      () => { this.filterImportantOnly = !this.filterImportantOnly; },
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
      onConversationSelect: (conv: ConversationGroup) =>
        this.handleConversationSelect(conv),
      onMessageSelect: (msg: Message) => this.handleMessageSelect(msg),
      onSenderSelect: (sender: SenderGroup) =>
        this.handleSenderSelect(sender),
      onBack: () => this.handleBack(),
      onLoadMore: () => this.handleLoadMore(),
      onMultiSelect: (ids: Set<string>) => this.handleMultiSelect(ids),
      onEditNickname: (addr: string, rawName: string) => this.openNicknameModal(addr, rawName),
    }, nameResolver, effectiveSenderResolver);
    this.messageList.setClassifications(this.classificationCache);

    const viewerEl = rightPane.createDiv({ cls: "iris-message-viewer" });
    this.messageViewer = new MessageViewer(viewerEl, this.plugin.app, {
      onMarkAsRead: (msg: Message) => this.handleMarkAsRead(msg),
      onMarkAsUnread: (msg: Message) => this.handleMarkAsUnread(msg),
      onTagChange: (msg: Message, tag: string | null) => this.handleTagChange(msg, tag),
      onImportanceChange: (msg, importance) => this.handleImportanceChange(msg, importance),
      onRetagMessage: (msg) => this.handleRetagMessage(msg),
      onReclassifyMessage: (msg) => this.handleReclassifyMessage(msg),
      onBatchMarkAsRead: (ids) => this.handleBatchMarkAsRead(ids),
      onBatchMarkAsUnread: (ids) => this.handleBatchMarkAsUnread(ids),
      onBatchTag: (ids, tag) => this.handleBatchTag(ids, tag),
      onBulkDenyImportance: (ids, oldClass, newClass) => this.handleBulkDenyImportance(ids, oldClass, newClass),
      onBulkDenyTag: (ids, tag) => this.handleBulkDenyTag(ids, tag),
      onCreateNoteFromSelection: (text, noteType, msg) => this.handleCreateNoteFromSelection(text, noteType, msg),
      onAcceptDetectedItem: (messageId, item) => this.handleAcceptDetectedItem(messageId, item),
      onDismissDetectedItem: (messageId, itemId) => this.handleDismissDetectedItem(messageId, itemId),
      onUpdateDetectedItem: (messageId, itemId, updates) => this.handleUpdateDetectedItem(messageId, itemId, updates),
      onReloadDetectedItems: (messageId) => this.handleReloadDetectedItems(messageId),
      onReprocessMessage: (msg) => this.handleReprocessMessage(msg),
      onEditNickname: (addr: string, rawName: string) => this.openNicknameModal(addr, rawName),
    }, nameResolver);
    this.messageViewer.setEffectiveSenderResolver(effectiveSenderResolver);
    this.messageViewer.setTagCategories(this.getTagCategories());
    this.messageViewer.setTagIcons(this.getTagIconMap());
    this.messageViewer.setTagCache(this.tagCache);
    this.messageViewer.setPromptVersions(
      this.plugin.settings.tagPromptVersion,
      this.plugin.settings.importancePromptVersion,
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

  /** Recompute conversation/sender groupings and update badge after any state change. */
  private regroupAndSync(): void {
    this.messageState.conversations = this.groupByConversation(this.messageState.messages);
    this.senderGroups = this.groupByEffectiveSender(this.messageState.messages);
    this.syncBadge();
  }

  /** Push the current classification cache to the message list and refresh importance indicators. */
  private syncClassificationUI(): void {
    this.messageList.setClassifications(this.classificationCache);
    this.messageList.updateImportanceIndicators();
  }

  /** Re-render the viewer for the currently selected message with its classification data. */
  private renderSelectedMessage(msg: Message): void {
    if (!msg.id || this.selectedMessageId !== msg.id) return;
    this.messageViewer.setDetectedItems(this.detectedItemsCache.get(msg.id) || []);
    this.messageViewer.render(
      msg,
      this.lastStrippedHtml,
      this.classificationCache.get(msg.id),
      this.classificationSourceCache.get(msg.id),
      this.classificationVersionCache.get(msg.id),
    );
  }

  /** Reset drill-down/selection state and clear viewer. */
  private clearDrillDown(): void {
    this.activeConversation = null;
    this.activeConversationMessages = [];
    this.activeSender = null;
    this.selectedMessageId = null;
    this.messageList.clearMultiSelection();
    this.messageViewer.clear();
  }

  /** Classify a message via Claude and update all caches. */
  private async classifyAndCacheMessage(msgId: string, content: string): Promise<ImportanceClass> {
    const result = await this.classifier.classifyAndCache(msgId, content);
    this.syncClassificationUI();
    return result;
  }

  /** Batch mark messages as read or unread. */
  private handleBatchReadState(ids: Set<string>, markAsRead: boolean): void {
    const changed: string[] = [];

    if (!this.activeConversation && !this.activeSender) {
      for (const conv of this.messageState.conversations) {
        if (ids.has(conv.conversationId)) {
          for (const msg of conv.messages) {
            if (msg.id && msg.isRead !== markAsRead) {
              msg.isRead = markAsRead;
              if (markAsRead) this.plugin.store.markRead(msg.id);
              else this.plugin.store.markUnread(msg.id);
              changed.push(msg.id);
            }
          }
        }
      }
    } else {
      const allMessages = [
        ...this.messageState.messages,
        ...this.activeConversationMessages,
      ];
      for (const msg of allMessages) {
        if (msg.id && ids.has(msg.id) && msg.isRead !== markAsRead) {
          msg.isRead = markAsRead;
          if (markAsRead) this.plugin.store.markRead(msg.id);
          else this.plugin.store.markUnread(msg.id);
          changed.push(msg.id);
        }
      }
    }

    // Sync to Graph API in background
    const api = this.plugin.mailApi;
    for (const id of changed) {
      const call = markAsRead ? api.markAsRead(id) : api.markAsUnread(id);
      void call.catch((err) => logger.warn("InboxView", "Failed to sync read state", err));
    }

    this.regroupAndSync();
    this.messageList.clearMultiSelection();
    this.messageViewer.clear();
    this.renderCurrentView();
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

    try {
      const filter =
        !this.plugin.settings.showReadEmails ? "isRead eq false" : undefined;

      const response: GraphPagedResponse<Message> =
        await this.plugin.mailApi.listMessages(folderId, {
          top: this.plugin.settings.pageSize,
          search: this.messageState.searchQuery || undefined,
          filter,
        });

      this.messageState.messages = response.value;
      this.plugin.store.applyReadState(this.messageState.messages);
      this.messageState.nextLink = response["@odata.nextLink"] || null;
      this.messageState.selectedConversationId = null;
      this.regroupAndSync();
      this.renderCurrentView();
      this.startBackgroundProcessing();
    } catch (err: unknown) {
      if (!this.plugin.authProvider.isSignedIn()) {
        this.renderCurrentView();
      } else {
        const msg = err instanceof Error ? err.message : String(err);
        new Notice(`Failed to load messages: ${msg}`);
      }
    } finally {
      this.messageState.isLoading = false;
    }
  }

  private async handleConversationSelect(
    conv: ConversationGroup,
  ): Promise<void> {
    if (!conv.conversationId) {
      new Notice("Cannot open conversation: missing conversation ID.");
      return;
    }
    this.messageState.selectedConversationId = conv.conversationId;
    this.activeConversation = conv;

    let fullMessages: Message[];
    try {
      fullMessages = await this.fetchConversationBodies(conv);
    } catch (err: unknown) {
      const errMsg = err instanceof Error ? err.message : String(err);
      new Notice(`Failed to load conversation: ${errMsg}`);
      return;
    }

    this.activeConversationMessages = fullMessages;

    // For single-message conversations, skip the drill-down
    if (conv.messages.length === 1) {
      const msg = fullMessages[fullMessages.length - 1];
      await this.showMessageInViewer(msg);
      return;
    }

    // Multi-message: drill into message list
    this.messageViewer.clear();
    this.messageList.renderConversationMessages(conv.subject, fullMessages);

    // Auto-select the latest message
    const latest = fullMessages[fullMessages.length - 1];
    await this.showMessageInViewer(latest);
  }

  /** Fetch conversation bodies with L1 (memory) → L2 (disk) → Graph fallback. */
  private async fetchConversationBodies(
    conv: ConversationGroup,
  ): Promise<Message[]> {
    const convId = conv.conversationId;

    // L1: in-memory cache
    const l1 = this.conversationBodyCache.get(convId);
    if (l1) return l1;

    const cache = this.plugin.store;
    const messageIds = conv.messages.map((m) => m.id!).filter(Boolean);

    // L2: disk cache — check if all message bodies are cached
    if (messageIds.length > 0 && cache.hasFullConversation(messageIds)) {
      const messages = conv.messages.map((m) => {
        const entry = cache.getBody(m.id!)!;
        return {
          ...m,
          body: { content: entry.bodyHtml, contentType: "html" as const },
        };
      });
      this.conversationBodyCache.set(convId, messages);
      return messages;
    }

    // L3: Graph API fetch
    const messages = await this.plugin.mailApi.getConversationMessages(convId);

    // Backfill disk cache
    for (const msg of messages) {
      if (msg.body?.content && msg.id && !cache.hasBody(msg.id)) {
        cache.setBody(msg, msg.body.content);
      }
    }

    this.conversationBodyCache.set(convId, messages);
    return messages;
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
      void this.plugin.mailApi.markAsRead(msg.id).catch((err) =>
        logger.warn("InboxView", "Failed to sync read state", err));
    }

    this.regroupAndSync();

    // In sender view, stay on the same person's message list
    if (this.activeSender) {
      const updated = this.senderGroups.find(
        (s) => s.groupKey === this.activeSender!.groupKey,
      );
      if (updated) {
        this.activeSender = updated;
        this.selectedMessageId = null;
        this.messageViewer.clear();
        const displayName = this.resolveName(
          updated.address,
          updated.name || updated.address,
        );
        const msgs = this.filterSenderMessages(updated.messages);
        if (msgs.length > 0) {
          this.messageList.renderConversationMessages(displayName, msgs, true);
          return;
        }
        // No messages left after filtering — go back to top-level
      }
    }

    // Otherwise go back to the top-level list
    this.handleBack();
  }

  private handleMarkAsUnread(msg: Message): void {
    msg.isRead = false;

    if (msg.id) {
      const canonical = this.messageState.messages.find((m) => m.id === msg.id);
      if (canonical && canonical !== msg) canonical.isRead = false;

      this.plugin.store.markUnread(msg.id);
      void this.plugin.mailApi.markAsUnread(msg.id).catch((err) =>
        logger.warn("InboxView", "Failed to sync unread state", err));
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

  /** Bulk deny importance: merge emails into formula via Haiku, then refine prompt via Opus. */
  private async handleBulkDenyImportance(
    ids: Set<string>,
    oldClassification: ImportanceClass,
    newClassification: ImportanceClass,
  ): Promise<void> {
    const s = this.plugin.settings;
    if (!s.anthropicApiKey) return;

    const msgIds = this.resolveBatchMessageIds(ids);
    const contents: string[] = [];

    // Apply the corrected classification to each message immediately
    for (const msgId of msgIds) {
      this.classificationCache.set(msgId, newClassification);
      this.classificationSourceCache.set(msgId, "manual");
      this.classificationVersionCache.delete(msgId);
      this.plugin.store.setClassification(msgId, newClassification, "manual");

      const msg = this.messageState.messages.find((m) => m.id === msgId);
      if (msg) {
        const content = this.getClassifiableContent(msg);
        if (content) contents.push(content);
      }
    }

    this.syncClassificationUI();
    this.messageList.clearMultiSelection();
    this.messageViewer.clear();
    this.renderCurrentView();
    this.syncBadge();

    if (contents.length === 0) return;

    // Merge → refine in background
    try {
      new Notice(`Merging ${contents.length} emails into formula…`);
      const formula = await mergeEmailsToFormula(s.anthropicApiKey, contents);

      const refined = await refineImportancePromptBulk(
        s.anthropicApiKey,
        this.getEffectiveImportancePrompt(),
        formula,
        oldClassification,
        newClassification,
      );
      await this.applyRefinedPrompt("importance", refined, "Importance prompt bulk-corrected by Opus");
    } catch (err) {
      logger.warn("InboxView", "Bulk importance prompt refinement failed", err);
      new Notice("Bulk importance refinement failed — classifications still applied.");
    }
  }

  /** Bulk deny tag: remove tag from all selected, merge into formula, refine prompt. */
  private async handleBulkDenyTag(ids: Set<string>, tag: string): Promise<void> {
    const s = this.plugin.settings;
    if (!s.anthropicApiKey) return;

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

      const refined = await refineTagPromptBulk(
        s.anthropicApiKey,
        this.getEffectiveTagPrompt(),
        formula,
        tag,
        "incorrect",
      );
      await this.applyRefinedPrompt("tag", refined, "Tag prompt bulk-corrected by Opus");
    } catch (err) {
      logger.warn("InboxView", "Bulk tag prompt refinement failed", err);
      new Notice("Bulk tag refinement failed — tags still removed.");
    }
  }

  /** Expand batch IDs to message IDs (conversation IDs → their messages). */
  private resolveBatchMessageIds(ids: Set<string>): string[] {
    if (!this.activeConversation && !this.activeSender) {
      // Top-level: IDs are conversationIds
      const msgIds: string[] = [];
      for (const conv of this.messageState.conversations) {
        if (ids.has(conv.conversationId)) {
          for (const msg of conv.messages) {
            if (msg.id) msgIds.push(msg.id);
          }
        }
      }
      return msgIds;
    }
    // Drill-down: IDs are message IDs
    return [...ids];
  }

  private handleBack(): void {
    this.clearDrillDown();
    this.renderCurrentView();
  }

  private async showMessageInViewer(msg: Message): Promise<void> {
    this.selectedMessageId = msg.id || null;
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
    const msgClass = msg.id ? this.classificationCache.get(msg.id) : undefined;
    const msgClassSource = msg.id ? this.classificationSourceCache.get(msg.id) : undefined;
    const msgClassVer = msg.id ? this.classificationVersionCache.get(msg.id) : undefined;
    this.messageViewer.setDetectedItems(this.detectedItemsCache.get(msg.id!) || []);
    this.messageViewer.render(msg, stripped, msgClass, msgClassSource, msgClassVer);

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
    if (!s.anthropicApiKey) {
      logger.warn("InboxView", "Claude processing enabled but no API key set");
      return;
    }
    if (!stripped) return;

    const parsedBody = htmlToMarkdown(stripped);
    if (!parsedBody) return;

    this.messageViewer.showProcessingIndicator();

    // Classify importance first (skip full processing for noise)
    let classification = this.classificationCache.get(msgId);
    if (!classification) {
      try {
        classification = await this.classifyAndCacheMessage(msgId, parsedBody);
      } catch (err) {
        logger.warn("InboxView", "Classification failed, proceeding anyway", err);
        classification = "routine";
      }
    }

    if (this.selectedMessageId !== msgId) return;

    if (classification === "noise") {
      return;
    }

    // Prepend email context (subject, sender, date) so Claude has full context
    // even when the body is sparse (e.g. meeting invitations, calendar events).
    const parsedContent = this.buildEmailContext(msg) + parsedBody;

    processEmailWithClaude(s.anthropicApiKey, s.claudeModel, effectivePrompt, parsedContent)
      .then(async (markdown) => {
        // Store raw markdown in memory (no frontmatter — viewer renders it)
        this.processedCache.set(msgId, markdown);

        // Persist to disk with frontmatter for the vault file
        const withFrontmatter = this.prependFrontmatter(msg, markdown);
        try {
          await cache.setProcessed(
            msgId,
            markdown,
            withFrontmatter,
            promptHash,
            s.saveFolderPath,
            msg,
          );
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
    this.activeConversation = null;
    this.activeConversationMessages = [];
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
        .setTitle("Reset importance prompt")
        .setIcon("signal")
        .onClick(() => {
          this.plugin.settings.importanceClassifyPrompt = "";
          this.plugin.settings.importancePromptVersion++;
          void this.plugin.saveSettings();
          this.syncPromptVersions();
          new Notice("Importance prompt reset to default");
        }),
    );
    menu.addItem((item) =>
      item
        .setTitle("Reset tagging prompt")
        .setIcon("tags")
        .onClick(() => {
          this.plugin.settings.tagClassifyPrompt = "";
          this.plugin.settings.tagPromptVersion++;
          void this.plugin.saveSettings();
          this.syncPromptVersions();
          new Notice("Tagging prompt reset to default");
        }),
    );
    menu.showAtPosition({ x, y });
  }

  // --- Private: view mode ---

  private handleReclassify(): void {
    this.classifier.clearAutoClassifications();
    this.classifier.clearAutoTags();
    this.messageList.setClassifications(this.classificationCache);
    this.messageViewer.setTagCache(this.tagCache);
    this.renderCurrentView();
    this.startBackgroundProcessing();
  }

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
    s.filterHideNoise = this.filterHideNoise;
    s.filterImportantOnly = this.filterImportantOnly;
    void this.plugin.saveSettings();
  }

  private handleViewModeToggle(): void {
    this.viewMode =
      this.viewMode === "conversations" ? "senders" : "conversations";
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
        this.messageList.renderConversationMessages(displayName, msgs, true);
        return;
      }
    }

    // If we're drilled into a conversation, re-render with current sort
    if (this.activeConversation && this.activeConversationMessages.length > 1) {
      const dir = this.sortNewestFirst ? -1 : 1;
      const sorted = [...this.activeConversationMessages].sort(
        (a, b) =>
          dir *
          (new Date(a.receivedDateTime || 0).getTime() -
            new Date(b.receivedDateTime || 0).getTime()),
      );
      this.messageList.renderConversationMessages(
        this.activeConversation.subject,
        sorted,
      );
      return;
    }

    const hasMore = !!this.messageState.nextLink;

    // Build a message-level filter from active toggles
    const passesFilter = (m: Message) => {
      const filtered = this.applyMessageFilters([m]);
      return filtered.length > 0;
    };
    const hasFilters = this.filterUnreadOnly || this.filterImportantOnly || this.filterHideNoise;

    if (hasFilters) {
      if (this.viewMode === "senders") {
        const filtered = this.senderGroups.filter((s) =>
          s.messages.some(passesFilter),
        );
        this.messageList.renderSenders(filtered, hasMore, passesFilter);
      } else {
        const filtered = this.messageState.conversations.filter((c) =>
          c.messages.some(passesFilter),
        );
        this.messageList.renderConversations(filtered, hasMore);
      }
    } else if (this.viewMode === "senders") {
      this.messageList.renderSenders(this.senderGroups, hasMore);
    } else {
      this.messageList.renderConversations(
        this.messageState.conversations,
        hasMore,
      );
    }
  }

  /**
   * Kick off all background AI processing (classification, tagging, prefetch, detection).
   * Handles errors with user-facing notices instead of swallowing rejections.
   */
  private startBackgroundProcessing(): void {
    this.classifyAllPromise = this.classifier.classifyAllMessages(
      this.messageState.messages,
      () => this.syncClassificationUI(),
    ).then(() => {
      // Re-render if filters depend on classification
      if (this.filterUnreadOnly || this.filterHideNoise || this.filterImportantOnly) {
        this.renderCurrentView();
      }
    }).catch((err) => {
      logger.warn("InboxView", "Background classification failed", err);
    });

    void this.generateAllNicknames();

    // Tag after classification finishes
    this.classifyAllPromise.then(() => {
      void this.classifier.autoTagAllMessages(
        this.messageState.messages,
        () => this.messageViewer.setTagCache(this.tagCache),
      ).catch((err) => logger.warn("InboxView", "Auto-tagging failed", err));
    });

    this.prefetchAllPromise = this.prefetchAllProcessed();
    this.prefetchAllPromise.catch((err) =>
      logger.warn("InboxView", "Background prefetch failed", err),
    );

    // When Claude processing is disabled but forwarded-sender resolution is
    // on, bodies are never prefetched by prefetchAllProcessed().  Fetch them
    // here so originalSender gets extracted and the sender list updates.
    const s = this.plugin.settings;
    if (s.resolveForwardedSender && (!s.enableClaudeProcessing || !s.anthropicApiKey)) {
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
    if (!s.enableClaudeProcessing || !s.anthropicApiKey) return;

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

    for (const [addr, rawName] of seen) {
      try {
        const nickname = await generateNickname(
          s.anthropicApiKey,
          s.claudeModel,
          NICKNAME_PROMPT,
          rawName,
        );
        // Re-check after the async gap -- the user may have deleted
        // or manually set a nickname while generation was in-flight.
        if (this.nicknameCache.has(addr) || this.plugin.store.isNicknameDeleted(addr)) continue;
        this.nicknameCache.set(addr, nickname);
        this.plugin.store.setNickname(addr, nickname);
      } catch (err) {
        logger.warn("InboxView", "Nickname generation failed", err);
      }
    }

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
    this.messageList.renderConversationMessages(displayName, msgs, true);

    // Auto-select the latest message
    if (msgs.length > 0) {
      void this.showMessageInViewer(msgs[msgs.length - 1]);
    }
  }

  /** Apply active toggle filters (unread, important, hide-noise) to a message list. */
  private applyMessageFilters(messages: Message[]): Message[] {
    let filtered = messages;
    if (this.filterUnreadOnly) {
      filtered = filtered.filter((m) => !m.isRead);
    }
    if (this.filterImportantOnly) {
      filtered = filtered.filter(
        (m) => this.classificationCache.get(m.id || "") === "important",
      );
    }
    if (this.filterHideNoise) {
      filtered = filtered.filter(
        (m) => this.classificationCache.get(m.id || "") !== "noise",
      );
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
      this.activeConversationMessages.find((m) => m.id === this.selectedMessageId) ||
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
        const current = this.plugin.settings.tagIcons[cat] || "tag";
        const idx = TAG_ICON_CYCLE.indexOf(current);
        this.plugin.settings.tagIcons[cat] =
          TAG_ICON_CYCLE[(idx + 1) % TAG_ICON_CYCLE.length];
        void this.plugin.saveSettings();
        this.messageViewer.setTagIcons(this.getTagIconMap());
        this.rebuildTagWrap();
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

    // Update expanded width: lead (28) + (categories + plus) * 28
    const expandedCount = categories.length + 1; // tag icons + plus
    this.tagWrap.style.setProperty("--tag-expanded-width", `${28 + expandedCount * 28}px`);
  }

  private showAddTagInput(anchor: HTMLElement): void {
    // Remove any existing input from the topbar
    anchor.parentElement?.querySelector(".iris-add-tag-input")?.remove();

    const input = createEl("input", {
      cls: "iris-add-tag-input",
      attr: { type: "text", placeholder: "New tag…" },
    });
    anchor.after(input);
    input.focus();

    const commit = () => {
      const value = input.value.trim();
      input.remove();
      if (!value) return;

      // Append to existing categories
      const existing = this.plugin.settings.tagCategories;
      const categories = existing
        ? existing.split(",").map((s) => s.trim()).filter(Boolean)
        : [];
      if (categories.includes(value)) return; // already exists
      categories.push(value);
      this.plugin.settings.tagCategories = categories.join(", ");

      // Auto-assign icon from cycle
      if (!this.plugin.settings.tagIcons) this.plugin.settings.tagIcons = {};
      if (!this.plugin.settings.tagIcons[value]) {
        const usedIcons = new Set(Object.values(this.plugin.settings.tagIcons));
        const nextIcon = TAG_ICON_CYCLE.find((i) => !usedIcons.has(i)) || TAG_ICON_CYCLE[categories.length % TAG_ICON_CYCLE.length];
        this.plugin.settings.tagIcons[value] = nextIcon;
      }
      void this.plugin.saveSettings();

      // Update runtime state
      this.messageViewer.setTagCategories(categories);
      this.messageViewer.setTagIcons(this.getTagIconMap());
      this.rebuildTagWrap();
    };

    input.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        commit();
      } else if (e.key === "Escape") {
        input.remove();
      }
    });
    input.addEventListener("blur", () => commit());
  }

  private getTagCategories(): string[] {
    return parseTagCategories(this.plugin.settings.tagCategories);
  }

  private getTagIconMap(): Map<string, string> {
    const icons = this.plugin.settings.tagIcons || {};
    return new Map(Object.entries(icons));
  }

  private getClassifiableContent(msg: Message): string {
    return [msg.subject, msg.bodyPreview].filter(Boolean).join("\n");
  }

  private getEffectiveTagPrompt(): string {
    return this.plugin.settings.tagClassifyPrompt || TAG_CLASSIFY_PROMPT;
  }

  private getEffectiveImportancePrompt(): string {
    return this.plugin.settings.importanceClassifyPrompt || IMPORTANCE_CLASSIFY_PROMPT;
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

  /** Call Claude Opus to refine the tag classification prompt based on feedback. */
  /** Save a refined prompt, bump version, notify, and sync. */
  private async applyRefinedPrompt(
    type: "tag" | "importance",
    refined: string,
    noticeLabel: string,
  ): Promise<void> {
    if (type === "tag") {
      this.plugin.settings.tagClassifyPrompt = refined;
      this.plugin.settings.tagPromptVersion++;
    } else {
      this.plugin.settings.importanceClassifyPrompt = refined;
      this.plugin.settings.importancePromptVersion++;
    }
    await this.plugin.saveSettings();
    const ver = type === "tag"
      ? this.plugin.settings.tagPromptVersion
      : this.plugin.settings.importancePromptVersion;
    new Notice(`${noticeLabel} (v${ver})`);
    this.syncPromptVersions();
  }

  private async refinePromptWithFeedback(
    msg: Message,
    tag: string,
    feedback: "correct" | "incorrect",
  ): Promise<void> {
    const s = this.plugin.settings;
    if (!s.anthropicApiKey) return;

    const content = this.getClassifiableContent(msg);
    if (!content) return;

    try {
      const refined = await refineTagPrompt(
        s.anthropicApiKey,
        this.getEffectiveTagPrompt(),
        content,
        tag,
        feedback,
      );
      const label = feedback === "correct" ? "reinforced" : "corrected";
      await this.applyRefinedPrompt("tag", refined, `Tag prompt ${label} by Opus`);
    } catch (err) {
      logger.warn("InboxView", "Prompt refinement failed", err);
    }
  }

  /** Manually set importance classification. */
  private handleImportanceChange(msg: Message, importance: ImportanceClass): void {
    if (!msg.id) return;

    const wasAuto = this.classificationSourceCache.get(msg.id) === "auto";
    const oldClassification = this.classifier.setManualClassification(msg.id, importance);

    this.syncClassificationUI();
    this.renderSelectedMessage(msg);

    // If overriding an auto-classification, refine the prompt
    if (wasAuto && oldClassification && oldClassification !== importance) {
      void this.refineImportancePromptWithFeedback(msg, oldClassification, importance);
    }
  }

  /** Call Claude Opus to refine the importance classification prompt. */
  private async refineImportancePromptWithFeedback(
    msg: Message,
    oldClassification: string,
    newClassification: string,
  ): Promise<void> {
    const s = this.plugin.settings;
    if (!s.anthropicApiKey) return;

    const content = this.getClassifiableContent(msg);
    if (!content) return;

    try {
      const refined = await refineImportancePrompt(
        s.anthropicApiKey,
        this.getEffectiveImportancePrompt(),
        content,
        oldClassification,
        newClassification,
      );
      await this.applyRefinedPrompt("importance", refined, "Importance prompt corrected by Opus");
    } catch (err) {
      logger.warn("InboxView", "Importance prompt refinement failed", err);
      new Notice("Importance prompt refinement failed — classification still applied.");
    }
  }

  // autoTagAllMessages is now handled by EmailClassifier

  /** Pre-process emails with Claude in the background so results are cached before the user clicks. */
  private async prefetchAllProcessed(): Promise<void> {
    const s = this.plugin.settings;
    if (!s.enableClaudeProcessing || !s.anthropicApiKey) return;

    const limit = s.prefetchLimit;
    if (limit === 0) return;

    const gen = this.prefetchGeneration;

    // Wait for classification to finish so we can skip noise
    if (this.classifyAllPromise) {
      try { await this.classifyAllPromise; } catch { /* proceed anyway */ }
    }
    if (this.prefetchGeneration !== gen) return;

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

      // Skip noise
      const classification = this.classificationCache.get(msgId);
      if (classification === "noise") {
        this.processedCache.set(msgId, "*Classification: noise — skipped processing*");
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

      // Classify if not yet done
      let msgClass = this.classificationCache.get(msgId);
      if (!msgClass) {
        try {
          msgClass = await this.classifyAndCacheMessage(msgId, parsedBody);
        } catch {
          msgClass = "routine";
        }
      }
      if (this.prefetchGeneration !== gen) return;

      if (msgClass === "noise") {
        this.processedCache.set(msgId, "*Classification: noise — skipped processing*");
        continue;
      }

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

        const withFrontmatter = this.prependFrontmatter(msg, markdown);
        try {
          await cache.setProcessed(msgId, markdown, withFrontmatter, promptHash, s.saveFolderPath, msg);
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
    if (!s.enableAutoItemDetection || !s.enableClaudeProcessing || !s.anthropicApiKey) return;

    const gen = this.prefetchGeneration;

    // Wait for classification so we can skip noise
    if (this.classifyAllPromise) {
      try { await this.classifyAllPromise; } catch { /* proceed anyway */ }
    }
    if (this.prefetchGeneration !== gen) return;

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

      // Skip noise
      const classification = this.classificationCache.get(msgId);
      if (classification === "noise") continue;

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
    if (!s.enableClaudeProcessing || !s.anthropicApiKey) return;

    const msgId = msg.id;
    if (!msgId) return;

    const cache = this.plugin.store;
    const classification = this.classificationCache.get(msgId);
    if (classification === "noise") return;

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
      this.plugin.settings.tagPromptVersion,
      this.plugin.settings.importancePromptVersion,
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
    if (!s.enableClaudeProcessing || !s.anthropicApiKey) return;

    const categories = this.getTagCategories();
    if (categories.length === 0) return;

    const content = this.getClassifiableContent(msg);
    if (!content) return;

    // Remove old auto-tags, keep manual
    const existing = this.tagCache.get(msg.id) || [];
    const manual = existing.filter((e) => e.source === "manual");
    for (const e of existing) {
      if (e.source === "auto") this.plugin.store.removeTag(msg.id, e.tag);
    }

    try {
      const tags = await classifyEmailTags(
        s.anthropicApiKey,
        s.claudeModel,
        this.getEffectiveTagPrompt(),
        content,
        categories,
      );

      const tagVer = s.tagPromptVersion;
      const newEntries: TagCacheEntry[] = tags.map((tag) => ({
        messageId: msg.id!,
        tag,
        source: "auto" as const,
        promptVersion: tagVer,
        taggedAt: Date.now(),
      }));
      const merged = [...manual, ...newEntries];
      if (merged.length > 0) {
        this.tagCache.set(msg.id, merged);
      } else {
        this.tagCache.delete(msg.id);
      }
      for (const tag of tags) {
        this.plugin.store.setTag(msg.id, tag, "auto", tagVer);
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

  /** Re-classify a single message with the current importance prompt. */
  private async handleReclassifyMessage(msg: Message): Promise<void> {
    if (!msg.id) return;
    const s = this.plugin.settings;
    if (!s.enableClaudeProcessing || !s.anthropicApiKey) return;

    const content = this.getClassifiableContent(msg);
    if (!content) return;

    try {
      await this.classifyAndCacheMessage(msg.id, content);
    } catch (err) {
      logger.warn("InboxView", "Re-classify failed", err);
    }

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
      case "important":
        this.plugin.updateBadge(
          msgs.filter(
            (m) => !m.isRead && this.classificationCache.get(m.id || "") === "important",
          ).length,
        );
        break;
      case "total":
        this.plugin.updateBadge(msgs.length);
        break;
    }
  }

  private groupByConversation(messages: Message[]): ConversationGroup[] {
    const groups = new Map<string, Message[]>();

    for (const msg of messages) {
      const convId = msg.conversationId || msg.id || "";
      const existing = groups.get(convId);
      if (existing) {
        existing.push(msg);
      } else {
        groups.set(convId, [msg]);
      }
    }

    const conversations: ConversationGroup[] = [];
    for (const [conversationId, msgs] of groups) {
      msgs.sort(
        (a, b) =>
          new Date(a.receivedDateTime || 0).getTime() -
          new Date(b.receivedDateTime || 0).getTime(),
      );

      const latestMessage = msgs[msgs.length - 1];
      const subject = (latestMessage.subject || "").replace(
        /^(?:re|fw|fwd):\s*/i,
        "",
      );

      conversations.push({
        conversationId,
        messages: msgs,
        subject,
        latestMessage,
        unreadCount: msgs.filter((m) => !m.isRead).length,
      });
    }

    const dir = this.sortNewestFirst ? -1 : 1;
    conversations.sort(
      (a, b) =>
        dir * (new Date(a.latestMessage.receivedDateTime || 0).getTime() -
        new Date(b.latestMessage.receivedDateTime || 0).getTime()),
    );

    return conversations;
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

  private prependFrontmatter(msg: Message, markdown: string): string {
    const dt = msg.receivedDateTime ? new Date(msg.receivedDateTime) : null;
    const date = dt ? dt.toISOString().split("T")[0] : "unknown";
    const time = dt
      ? dt.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" })
      : "unknown";
    const from =
      msg.from?.emailAddress?.name ||
      msg.from?.emailAddress?.address ||
      "unknown";
    const displayTitle = msg.subject || "no subject";

    const frontmatter = [
      "---",
      `displayTitle: "${displayTitle.replace(/"/g, '\\"')}"`,
      `date: ${date}`,
      `time: "${time}"`,
      `from: "${from.replace(/"/g, '\\"')}"`,
      "---",
    ].join("\n");

    return `${frontmatter}\n\n${markdown.trimStart()}`;
  }

  // ── Note creation from selected email text ────────────────────

  private async handleCreateNoteFromSelection(
    selectedText: string,
    noteType: NoteType,
    msg: Message,
  ): Promise<void> {
    const s = this.plugin.settings;
    if (!s.anthropicApiKey) {
      new Notice("Please configure your Anthropic API key in Iris Mail settings.");
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
        ...dateLine("dueDate", note.dueDate),
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
