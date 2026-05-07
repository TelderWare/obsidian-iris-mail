import { ItemView, WorkspaceLeaf, Notice, setIcon, Menu } from "obsidian";
import type IrisMailPlugin from "../main";
import {
  VIEW_TYPE_IRIS_MAIL,
  COMPACT_MODE_CLASS,
  DEFAULT_CLAUDE_PROMPT,
  NICKNAME_PROMPT,
  NICKNAME_BATCH_PROMPT,
  TAG_CLASSIFY_PROMPT,
  parseTagCategories,
  getTagVersion,
  bumpTagVersion,
} from "../constants";
import { MessageList } from "./components/MessageList";
import { MessageViewer } from "./components/MessageViewer";
import { NicknameModal } from "./components/NicknameModal";
import { SenderRuleModal } from "./components/SenderRuleModal";
import { SearchBar } from "./components/SearchBar";
import { Toolbar } from "./components/Toolbar";
import { processEmailWithClaude, classifyEmailTagsYesNo, refineTagCriteria, generateNickname, generateNicknamesBatch, mergeEmailsToFormula, refineTagCriteriaBulk, extractNoteFromSelection, hasClaudeAccess, type TagCandidate } from "../utils/claudeApi";
import type { NoteType, ExtractedNote } from "../utils/claudeApi";
import { htmlToMarkdown } from "../utils/htmlToMarkdown";
import { getEffectiveSender, makeEffectiveSenderResolver } from "../utils/effectiveSender";
import { getEnvelopeSender } from "../utils/envelopeSender";
import { logger } from "../utils/logger";
import { EmailStore } from "../store/EmailStore";
import { EmailClassifier } from "../services/EmailClassifier";
import type {
  Box,
  Message,
  MessageListState,
} from "../types";
import { BoxEditModal } from "./components/BoxEditModal";
import { TagsModal } from "./components/TagsModal";
import type { MailListResponse } from "../mail/MailApi";
import { normalizeName } from "../utils/nameResolve";

export class InboxView extends ItemView {
  private plugin: IrisMailPlugin;
  private messageList!: MessageList;
  private messageViewer!: MessageViewer;
  private searchBar!: SearchBar;
  private toolbar!: Toolbar;

  private messageState: MessageListState = {
    messages: [],
    nextLink: null,
    isLoading: false,
    searchQuery: "",
  };

  private sortNewestFirst!: boolean;
  private sortToggleBtn!: HTMLButtonElement;

  // Boxes (named views replacing the old unread + tag filter bar)
  private selectedBoxId!: string;
  private boxStripEl!: HTMLDivElement;
  private todoIds = new Set<string>();
  private junkIds = new Set<string>();
  private pinnedIds = new Set<string>();
  private classifierUnsubscribe: (() => void) | null = null;
  private todosUnsubscribe: (() => void) | null = null;

  private selectedMessageId: string | null = null;
  private lastStrippedHtml: string = "";

  // Extracted classifier handles classification & tagging caches
  private classifier!: EmailClassifier;

  // In-memory caches
  private processedCache = new Map<string, string>();
  private nicknameCache = new Map<string, string>();
  // Render batching — coalesce multiple state changes into one paint.
  private pendingRender: { regroup?: boolean; list?: boolean; tags?: boolean } = {};
  private renderScheduled = false;

  // Prefetch state
  private prefetchGeneration = 0;
  private prefetchAllPromise: Promise<void> | null = null;
  /** Message IDs currently being summarized by Claude (prefetch or interactive).
   *  Unioned with the classifier's in-flight set to drive the Secretary box. */
  private summarizeInflight = new Set<string>();

  private markSummarizing(id: string, on: boolean): void {
    if (on) this.summarizeInflight.add(id);
    else this.summarizeInflight.delete(id);
    this.handleSecretaryUpdate();
  }

  constructor(leaf: WorkspaceLeaf, plugin: IrisMailPlugin) {
    super(leaf);
    this.plugin = plugin;
    this.classifier = new EmailClassifier(plugin.store, () => plugin.settings);
    const s = plugin.settings;
    this.sortNewestFirst = s.sortNewestFirst;
    this.selectedBoxId = s.selectedBoxId || "in";
    this.classifierUnsubscribe = this.classifier.onInFlightChange(() => this.handleSecretaryUpdate());
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

    if (!this.plugin.accounts.anySignedIn()) {
      this.renderSignInPrompt(container);
      return;
    }

    this.reloadCaches();
    this.renderInboxUI(container);

    this.todosUnsubscribe = this.plugin.onTodosChanged(() => this.handleExternalTodoChange());

    // Wire keyboard shortcuts at document level so they fire regardless of
    // which element inside the view has focus. The handler short-circuits
    // when another view is active or the user is typing in a text widget.
    this.registerDomEvent(document, "keydown", (evt: KeyboardEvent) => {
      if (this.app.workspace.getActiveViewOfType(InboxView) !== this) return;
      this.handleKeyDown(evt);
    });

    await this.loadInbox();
  }

  async onClose(): Promise<void> {
    this.prefetchGeneration++;
    if (this.classifierUnsubscribe) {
      this.classifierUnsubscribe();
      this.classifierUnsubscribe = null;
    }
    if (this.todosUnsubscribe) {
      this.todosUnsubscribe();
      this.todosUnsubscribe = null;
    }
    this.contentEl.empty();
  }

  async refresh(): Promise<void> {
    this.prefetchGeneration++;
    this.processedCache.clear();

    // Re-read persistent caches without tearing down the UI
    this.reloadCaches();

    // Merge server state into the existing list instead of overwriting —
    // preserves whatever the user has scrolled/paginated through.
    await this.mergeRefresh();
  }

  /**
   * Background-safe refresh: fetches the first page, merges by id into the
   * current list (new items appended, existing items updated, scrolled-in
   * pages preserved). Falls back to a full reload if we don't yet have any
   * messages.
   */
  private async mergeRefresh(): Promise<void> {
    if (!this.plugin.accounts.anySignedIn()) {
      this.renderCurrentView();
      return;
    }
    if (this.messageState.isLoading) return;
    if (this.messageState.messages.length === 0) {
      await this.loadInbox();
      return;
    }
    // A search or filter narrowing is active — a merge would leak unrelated
    // messages in. Defer to the normal search path.
    if (this.messageState.searchQuery) {
      await this.loadInbox();
      return;
    }

    const showRead = this.plugin.settings.showReadEmails;
    try {
      const response = await this.plugin.mailApi.listMessages("", {
        top: this.plugin.settings.pageSize,
        unreadOnly: !showRead,
        since: this.plugin.getSyncSince(),
      });
      this.plugin.store.applyReadState(response.value);

      const byId = new Map<string, Message>();
      for (const m of this.messageState.messages) {
        if (m.id) byId.set(m.id, m);
      }
      for (const m of response.value) {
        if (m.id) byId.set(m.id, m);
      }
      const merged = [...byId.values()].sort((a, b) => {
        const ad = a.receivedDateTime ?? "";
        const bd = b.receivedDateTime ?? "";
        return bd.localeCompare(ad);
      });
      this.messageState.messages = this.plugin.store.mergePersistedMessages(
        merged,
        this.plugin.getSavedBoxes(),
      );

      this.applySenderRules();
      this.regroupAndSync();
      this.renderCurrentView();
      this.startBackgroundProcessing();
      this.plugin.setInboxMessages(this.messageState.messages);
    } catch (err) {
      // Background refresh failed — leave existing state untouched and re-render.
      logger.warn("InboxView", "mergeRefresh failed", err);
    }
  }

  private reloadCaches(): void {
    this.nicknameCache = this.plugin.store.getAllNicknames();
    this.todoIds = this.plugin.store.getAllTodoIds();
    this.junkIds = this.plugin.store.getAllJunkIds();
    this.pinnedIds = this.plugin.store.getAllPinnedIds();
  }

  /** Display name of the signed-in user that owns this message, used as
   *  context for Claude prompts. Falls back to empty string if unknown. */
  private getMessageOwnerName(msg: Message): string {
    const accountId = msg._accountId;
    if (!accountId) return "";
    const entry = this.plugin.accounts.get(accountId);
    const acct = entry?.auth.getAccount();
    return acct?.name || acct?.username || "";
  }

  // --- Private: rendering ---

  private renderSignInPrompt(container: HTMLElement): void {
    const prompt = container.createDiv({ cls: "iris-sign-in-prompt" });
    const icon = prompt.createDiv({ cls: "iris-sign-in-icon" });
    setIcon(icon, "mail");
    prompt.createEl("h3", { text: "Iris Mail" });

    const accounts = this.plugin.settings.accounts;
    if (accounts.length === 0) {
      prompt.createEl("p", {
        cls: "iris-sign-in-desc",
        text: "Add an account in Iris Mail settings to get started — via Azure (OAuth) or via IMAP.",
      });
    } else {
      prompt.createEl("p", {
        cls: "iris-sign-in-desc",
        text: "Sign in to one of your configured accounts.",
      });
    }

    const btnGroup = prompt.createDiv({ cls: "iris-sign-in-buttons" });

    for (const account of accounts) {
      const btn = btnGroup.createEl("button", {
        text: `Sign in: ${account.label}`,
        cls: "mod-cta",
      });
      btn.addEventListener("click", () => void this.plugin.loginAccount(account.id));
    }

    const settingsBtn = btnGroup.createEl("button", {
      text: accounts.length === 0 ? "Open settings" : "Manage accounts",
    });
    settingsBtn.addEventListener("click", () => this.openIrisSettings());
  }

  private openIrisSettings(): void {
    // Open the plugin settings tab via Obsidian's Setting modal.
    const setting = (this.app as unknown as { setting: { open(): void; openTabById(id: string): void } }).setting;
    setting?.open();
    setting?.openTabById("iris-mail");
  }

  private renderInboxUI(container: HTMLElement): void {
    const topBar = container.createDiv({ cls: "iris-topbar" });

    // Left: sort toggle + box strip
    const leftControls = topBar.createDiv({ cls: "iris-topbar-left" });
    this.sortToggleBtn = leftControls.createEl("button", {
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

    // Box strip: named views (In, Read, To-do, Junk, Secretary, + user boxes).
    this.boxStripEl = leftControls.createDiv({ cls: "iris-box-strip" });
    this.renderBoxStrip();

    // Center: search
    const centerControls = topBar.createDiv({ cls: "iris-topbar-center" });
    this.searchBar = new SearchBar(centerControls, {
      onSearch: (query: string) => this.handleSearch(query),
    });

    const rightControls = topBar.createDiv({ cls: "iris-topbar-right" });

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

    // Tags — opens quick tag-management modal
    const tagsBtn = rightControls.createEl("button", {
      cls: "iris-topbar-btn clickable-icon",
      attr: { "aria-label": "Tags" },
    });
    setIcon(tagsBtn, "tags");
    tagsBtn.addEventListener("click", () => {
      new TagsModal(this.plugin.app, this.plugin, {
        onChange: () => this.handleTagsSettingsChanged(),
      }).open();
    });

    // Main area: message list + viewer (no sidebar)
    const mainEl = container.createDiv({ cls: "iris-main" });
    const rightPane = mainEl.createDiv({ cls: "iris-right-pane" });

    const nameResolver = (addr: string, raw: string) =>
      this.resolveName(addr, raw);
    const effectiveSenderResolver = makeEffectiveSenderResolver(this.plugin);

    const listEl = rightPane.createDiv({ cls: "iris-message-list" });
    this.messageList = new MessageList(listEl, {
      onMessageSelect: (msg: Message) => this.handleMessageSelect(msg),
      onLoadMore: () => this.handleLoadMore(),
      onMultiSelect: (ids: Set<string>) => this.handleMultiSelect(ids),
      onEditNickname: (addr: string, rawName: string) => this.openNicknameModal(addr, rawName),
      onEditSenderRule: (addr: string, rawName: string) => this.openSenderRuleModal(addr, rawName),
      onMarkSenderAsJunk: (addr: string, rawName: string) => this.markSenderAsJunk(addr, rawName),
    }, nameResolver, effectiveSenderResolver);

    const viewerEl = rightPane.createDiv({ cls: "iris-message-viewer" });
    this.messageViewer = new MessageViewer(viewerEl, this.plugin.app, {
      onMarkAsRead: (msg: Message) => this.handleMarkAsRead(msg),
      onMarkAsUnread: (msg: Message) => this.handleMarkAsUnread(msg),
      onToggleTodo: (msg: Message) => this.handleToggleTodo(msg),
      onTogglePin: (msg: Message) => this.handleTogglePin(msg),
      onTagChange: (msg: Message, tag: string | null) => this.handleTagChange(msg, tag),
      onRetagMessage: (msg) => this.handleRetagMessage(msg),
      onBatchMarkAsRead: (ids) => this.handleBatchMarkAsRead(ids),
      onBatchMarkAsUnread: (ids) => this.handleBatchMarkAsUnread(ids),
      onBatchTag: (ids, tag) => this.handleBatchTag(ids, tag),
      onBulkDenyTag: (ids, tag) => this.handleBulkDenyTag(ids, tag),
      onDeleteMessage: (msg: Message) => this.handleDeleteMessage(msg),
      onBatchDelete: (ids) => this.handleBatchDelete(ids),
      onCreateNoteFromSelection: (text, noteType, msg) => this.handleCreateNoteFromSelection(text, noteType, msg),
      onReprocessMessage: (msg) => this.handleReprocessMessage(msg),
      onEditNickname: (addr: string, rawName: string) => this.openNicknameModal(addr, rawName),
      onMarkSenderAsJunk: (addr: string, rawName: string) => this.markSenderAsJunk(addr, rawName),
      onDismiss: () => { this.messageViewer.clear(); },
    }, nameResolver);
    this.messageViewer.setEffectiveSenderResolver(effectiveSenderResolver);
    this.messageViewer.setTagCategories(this.getTagCategories());
    this.messageViewer.setTagIcons(this.getTagIconMap());
    this.messageViewer.setTagColors(this.getTagColorMap());
    this.messageViewer.setTagDescriptions(this.getTagDescriptionMap());
    this.messageViewer.setTagCache(this.plugin.store.getAllTags());
    this.messageViewer.setTodoIds(this.todoIds);
    this.messageViewer.setJunkIds(this.junkIds);
    this.messageViewer.setPinnedIds(this.pinnedIds);
    this.messageList.setTagIcons(this.getTagIconMap());
    this.messageList.setTagColors(this.getTagColorMap());
    this.messageList.setTagDescriptions(this.getTagDescriptionMap());
    this.messageList.setTagCache(this.plugin.store.getAllTags());
    this.messageList.setHiddenListTags(this.getHiddenListTagSet());
    this.messageList.setPinnedIds(this.pinnedIds);
    this.messageViewer.setPromptVersions(
      this.plugin.settings.tagPromptVersions || {},
    );
  }

  // --- Private: shared helpers ---

  /** Rebuild the box strip UI. Call when box list, selection, or counts change. */
  private renderBoxStrip(): void {
    if (!this.boxStripEl) return;
    this.boxStripEl.empty();

    const boxes = this.plugin.settings.boxes || [];
    const counts = this.computeBoxCounts();

    for (const box of boxes) {
      if (box.hidden) continue;
      const isActive = box.id === this.selectedBoxId;
      const chip = this.boxStripEl.createEl("button", {
        cls: "iris-box-chip clickable-icon" + (isActive ? " is-active" : ""),
        attr: { "aria-label": box.name, title: box.name },
      });
      const iconEl = chip.createSpan({ cls: "iris-box-chip-icon" });
      setIcon(iconEl, box.icon || "inbox");
      if (box.color) iconEl.style.color = box.color;
      const count = counts.get(box.id) ?? 0;
      if (count > 0) {
        chip.createSpan({
          cls: "iris-box-chip-count",
          text: count > 99 ? "99+" : String(count),
        });
      }
      chip.addEventListener("click", () => this.selectBox(box.id));
      chip.addEventListener("contextmenu", (evt) => {
        evt.preventDefault();
        this.openBoxContextMenu(evt, box);
      });
    }

    const addBtn = this.boxStripEl.createEl("button", {
      cls: "iris-box-chip iris-box-chip-add clickable-icon",
      attr: { "aria-label": "New box" },
    });
    setIcon(addBtn, "plus");
    addBtn.addEventListener("click", () => this.openBoxEditModal(null));
    addBtn.addEventListener("contextmenu", (evt) => {
      evt.preventDefault();
      this.openAddBoxContextMenu(evt);
    });
  }

  /** Return the count of messages matching each box's predicate. */
  private computeBoxCounts(): Map<string, number> {
    const counts = new Map<string, number>();
    const boxes = this.plugin.settings.boxes || [];
    for (const box of boxes) counts.set(box.id, 0);
    const inFlight = this.classifier.getInFlightIds();
    for (const msg of this.messageState.messages) {
      if (!msg.id) continue;
      for (const box of boxes) {
        if (this.messageMatchesBox(msg, box, inFlight)) {
          counts.set(box.id, (counts.get(box.id) ?? 0) + 1);
        }
      }
    }
    return counts;
  }

  /**
   * Predicate: does a message belong in the given box?
   *
   * Built-in boxes form an exclusive cascade: In > To-do > Junk > Read for
   * the main list; Junk also trumps Secretary so a junked message stops
   * showing up as "being processed" even if a background Claude call is
   * still finishing. An unread message always lands in In regardless of any
   * other state. Once read, a message goes to To-do if it carries an isTodo
   * flag or any tag wired into the To-do box; otherwise to Junk under the
   * same rule for the Junk box; otherwise to Read. User boxes match on
   * their tag predicate and exclude messages already claimed by To-do or
   * Junk.
   */
  private messageMatchesBox(msg: Message, box: Box, inFlight: Set<string>): boolean {
    const id = msg.id;
    if (!id) return false;

    const msgTags = this.getMessageTagSet(id);
    // Raw signals — computed independently of the In > To-do > Junk > Read
    // cascade because Secretary needs to see "is this junk?" even for unread.
    const junkSignal = this.junkIds.has(id) || this.hasBuiltinTagMatch(msgTags, "junk");
    const todoSignal = this.todoIds.has(id) || this.hasBuiltinTagMatch(msgTags, "todo");

    if (box.builtin === "secretary") {
      if (junkSignal) return false;
      return inFlight.has(id) || this.summarizeInflight.has(id);
    }

    const isUnread = !msg.isRead;
    const effectiveTodo = !isUnread && todoSignal;
    const effectiveJunk = !isUnread && !effectiveTodo && junkSignal;

    switch (box.builtin) {
      case "in":
        return isUnread;
      case "todo":
        return effectiveTodo;
      case "junk":
        return effectiveJunk;
      case "read":
        return !isUnread && !effectiveTodo && !effectiveJunk;
      default:
        if (effectiveTodo || effectiveJunk) return false;
        const boxTags = box.tags || [];
        if (boxTags.length === 0) return false;
        for (const t of boxTags) {
          if (msgTags.has(t)) return true;
        }
        return false;
    }
  }

  private getMessageTagSet(id: string): Set<string> {
    const entries = this.plugin.store.getTags(id);
    if (entries.length === 0) return new Set();
    return new Set(entries.map((e) => e.tag));
  }

  /** True if the message carries any tag wired into the given built-in box. */
  private hasBuiltinTagMatch(msgTags: Set<string>, builtin: "junk" | "todo"): boolean {
    if (msgTags.size === 0) return false;
    const box = (this.plugin.settings.boxes || []).find((b) => b.builtin === builtin);
    if (!box || !box.tags || box.tags.length === 0) return false;
    for (const t of box.tags) {
      if (msgTags.has(t)) return true;
    }
    return false;
  }

  private selectBox(boxId: string): void {
    if (this.selectedBoxId === boxId) return;
    this.selectedBoxId = boxId;
    this.renderBoxStrip();
    this.applyFilters();
    this.persistViewState();
  }

  /** Public entry point for external callers (e.g. homepage widgets). */
  showBox(boxId: string): void {
    this.selectBox(boxId);
  }

  private getSelectedBox(): Box | undefined {
    return (this.plugin.settings.boxes || []).find((b) => b.id === this.selectedBoxId);
  }

  /** Refresh the Secretary count chip (and if Secretary is selected, the list). */
  private handleSecretaryUpdate(): void {
    if (!this.boxStripEl) return;
    this.renderBoxStrip();
    if (this.selectedBoxId === "secretary") this.renderCurrentView();
  }

  private openBoxContextMenu(evt: MouseEvent, box: Box): void {
    const menu = new Menu();
    menu.addItem((item) =>
      item
        .setTitle("Edit box…")
        .setIcon("pencil")
        .onClick(() => this.openBoxEditModal(box)),
    );
    menu.addItem((item) =>
      item
        .setTitle("Hide box")
        .setIcon("eye-off")
        .onClick(() => this.setBoxHidden(box.id, true)),
    );
    if (!box.builtin) {
      menu.addItem((item) =>
        item
          .setTitle("Delete box")
          .setIcon("trash-2")
          .setWarning(true)
          .onClick(() => this.deleteBox(box.id)),
      );
    }
    menu.showAtMouseEvent(evt);
  }

  /** Right-click menu on the "+" button — lists any hidden boxes for restore. */
  private openAddBoxContextMenu(evt: MouseEvent): void {
    const boxes = this.plugin.settings.boxes || [];
    const hidden = boxes.filter((b) => b.hidden);
    if (hidden.length === 0) {
      this.openBoxEditModal(null);
      return;
    }
    const menu = new Menu();
    menu.addItem((item) =>
      item
        .setTitle("New box…")
        .setIcon("plus")
        .onClick(() => this.openBoxEditModal(null)),
    );
    menu.addSeparator();
    for (const box of hidden) {
      menu.addItem((item) =>
        item
          .setTitle(`Show ${box.name}`)
          .setIcon(box.icon || "inbox")
          .onClick(() => this.setBoxHidden(box.id, false)),
      );
    }
    menu.showAtMouseEvent(evt);
  }

  private setBoxHidden(id: string, hidden: boolean): void {
    const boxes = [...(this.plugin.settings.boxes || [])];
    const idx = boxes.findIndex((b) => b.id === id);
    if (idx < 0) return;
    boxes[idx] = { ...boxes[idx], hidden: hidden || undefined };
    this.plugin.settings.boxes = boxes;
    void this.plugin.saveSettings();
    // If the hidden box was selected, fall back to the first visible box.
    if (hidden && this.selectedBoxId === id) {
      const fallback = boxes.find((b) => !b.hidden);
      this.selectedBoxId = fallback?.id ?? "in";
      this.applyFilters();
    }
    this.renderBoxStrip();
  }

  private openBoxEditModal(existing: Box | null): void {
    new BoxEditModal(this.plugin.app, {
      initial: existing || undefined,
      existingIds: new Set((this.plugin.settings.boxes || []).map((b) => b.id)),
      availableTags: this.getTagCategories(),
      onSubmit: (draft) => {
        const boxes = [...(this.plugin.settings.boxes || [])];
        if (existing) {
          const idx = boxes.findIndex((b) => b.id === existing.id);
          if (idx >= 0) boxes[idx] = { ...boxes[idx], ...draft, id: existing.id, builtin: existing.builtin };
        } else {
          const newBox: Box = { ...draft, id: draft.id || `user-${Date.now()}` };
          boxes.push(newBox);
        }
        this.plugin.settings.boxes = boxes;
        void this.plugin.saveSettings();
        this.renderBoxStrip();
        if (this.selectedBoxId === (existing?.id ?? "")) this.renderCurrentView();
      },
    }).open();
  }

  private deleteBox(id: string): void {
    const boxes = (this.plugin.settings.boxes || []).filter((b) => b.id !== id);
    this.plugin.settings.boxes = boxes;
    if (this.selectedBoxId === id) this.selectedBoxId = "in";
    void this.plugin.saveSettings();
    this.renderBoxStrip();
    this.applyFilters();
  }

  /** Update the badge and box strip after any state change. */
  private regroupAndSync(): void {
    this.syncBadge();
    this.renderBoxStrip();
  }

  /** Push the current tag cache to the viewer and refresh list row badges in place. */
  private syncTagCacheViews(): void {
    const tags = this.plugin.store.getAllTags();
    this.messageViewer.setTagCache(tags);
    this.messageList.setTagCache(tags);
    this.messageList.refreshTagBadges();
  }

  /**
   * Coalesce a render pass. Several state-change helpers need to update the
   * badge/box strip, re-render the list, or push fresh tag data to children.
   * When called multiple times in the same tick (e.g. inside a batch loop)
   * only the union of requested work runs once on the next microtask.
   */
  private scheduleRender(opts: { regroup?: boolean; list?: boolean; tags?: boolean }): void {
    if (opts.regroup) this.pendingRender.regroup = true;
    if (opts.list) this.pendingRender.list = true;
    if (opts.tags) this.pendingRender.tags = true;
    if (this.renderScheduled) return;
    this.renderScheduled = true;
    queueMicrotask(() => this.flushRender());
  }

  private flushRender(): void {
    const p = this.pendingRender;
    this.pendingRender = {};
    this.renderScheduled = false;
    try {
      if (p.tags) this.syncTagCacheViews();
      if (p.regroup) this.regroupAndSync();
      if (p.list) this.renderCurrentView();
    } catch (err) {
      logger.warn("InboxView", "Render flush failed", err);
    }
  }

  /** Re-render the viewer for the currently selected message. */
  private renderSelectedMessage(msg: Message): void {
    if (!msg.id || this.selectedMessageId !== msg.id) return;
    this.messageViewer.render(msg, this.lastStrippedHtml);
  }

  /** Clear list selection and viewer. */
  private clearDrillDown(): void {
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

    this.messageList.clearMultiSelection();
    this.messageViewer.clear();
    this.scheduleRender({ regroup: true, list: true });
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
      // The dispatcher merges every account's Inbox; folderId is unused.
      await this.loadMessages("");
    } catch (err: unknown) {
      if (!this.plugin.accounts.anySignedIn()) {
        this.renderCurrentView();
      } else {
        const msg = err instanceof Error ? err.message : String(err);
        new Notice(`Failed to load inbox: ${msg}`);
      }
    }
  }

  private async loadMessages(folderId: string): Promise<void> {
    this.messageState.isLoading = true;
    this.messageList.showLoading();

    const showRead = this.plugin.settings.showReadEmails;
    const searchQuery = this.messageState.searchQuery || undefined;
    try {
      const response: MailListResponse<Message> =
        await this.plugin.mailApi.listMessages(folderId, {
          top: this.plugin.settings.pageSize,
          search: searchQuery,
          unreadOnly: !showRead,
          since: this.plugin.getSyncSince(),
        });

      this.messageState.messages = this.plugin.store.mergePersistedMessages(
        response.value,
        this.plugin.getSavedBoxes(),
      );
      this.plugin.store.applyReadState(this.messageState.messages);
      this.messageState.nextLink = response.nextLink;
      // Cache non-search list results so we can fall back on next failure.
      if (!searchQuery) {
        this.plugin.store.setMessageList(folderId, showRead, response.value, this.messageState.nextLink);
      }
      this.applySenderRules();
      this.regroupAndSync();
      this.renderCurrentView();
      this.startBackgroundProcessing();
      this.plugin.setInboxMessages(this.messageState.messages);
    } catch (err: unknown) {
      if (!this.plugin.accounts.anySignedIn()) {
        this.renderCurrentView();
      } else {
        // Fall back to cached list if available (non-search queries only).
        const cached = !searchQuery
          ? this.plugin.store.getMessageList(folderId, showRead)
          : undefined;
        if (cached) {
          this.messageState.messages = this.plugin.store.mergePersistedMessages(
            cached.messages as Message[],
            this.plugin.getSavedBoxes(),
          );
          this.plugin.store.applyReadState(this.messageState.messages);
          this.messageState.nextLink = cached.nextLink;
          this.regroupAndSync();
          this.renderCurrentView();
          this.plugin.setInboxMessages(this.messageState.messages);
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

    this.selectedMessageId = null;
    this.messageViewer.clear();
  }

  /** Find the next unread message in the current list, in the active sort order. */
  private findNextUnread(currentId: string | null): Message | null {
    const dir = this.sortNewestFirst ? -1 : 1;
    const list = [...this.messageState.messages].sort(
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

  /** Toggle the client-side to-do flag, auto-advance to next matching message. */
  private handleToggleTodo(msg: Message): void {
    if (!msg.id) return;
    const id = msg.id;
    const wasTodo = this.todoIds.has(id);
    if (wasTodo) {
      this.todoIds.delete(id);
      this.plugin.store.unmarkTodo(id);
    } else {
      this.todoIds.add(id);
      this.plugin.store.markTodo(id);
      // Adding to to-do implies the user has triaged the message — mark it
      // as read locally and on the server so it leaves the In box too.
      if (!msg.isRead) {
        msg.isRead = true;
        const canonical = this.messageState.messages.find((m) => m.id === id);
        if (canonical && canonical !== msg) canonical.isRead = true;
        this.plugin.store.markRead(id);
        void this.plugin.mailApi.markAsRead(id).catch((err) =>
          this.rollbackReadState(id, false, err));
      }
    }
    this.messageViewer.setTodoIds(this.todoIds);
    this.renderBoxStrip();
    this.advanceAfterStateChange(msg);
    this.plugin.notifyTodosChanged();
  }

  /** Re-sync from the store after an external `clearTodo` (e.g. from Iris Tasks). */
  private handleExternalTodoChange(): void {
    const fresh = this.plugin.store.getAllTodoIds();
    if (fresh.size === this.todoIds.size) {
      let same = true;
      for (const id of fresh) if (!this.todoIds.has(id)) { same = false; break; }
      if (same) return;
    }
    this.todoIds = fresh;
    this.messageViewer.setTodoIds(this.todoIds);
    this.renderBoxStrip();
  }

  private handleTogglePin(msg: Message): void {
    if (!msg.id) return;
    const id = msg.id;
    const wasPinned = this.pinnedIds.has(id);
    if (wasPinned) {
      this.pinnedIds.delete(id);
      this.plugin.store.unmarkPinned(id);
    } else {
      this.pinnedIds.add(id);
      // Persist the envelope so the pinned message re-appears even if it falls
      // outside the next sync window or is excluded by the read-only filter.
      this.plugin.store.markPinned(id, msg);
    }
    this.messageViewer.setPinnedIds(this.pinnedIds);
    this.messageList.setPinnedIds(this.pinnedIds);
    this.renderBoxStrip();
    // Pin/unpin never exiles the message from its current box, so re-render in
    // place rather than auto-advancing the way todo/junk do.
    this.renderCurrentView();
    this.plugin.setInboxMessages(this.messageState.messages);
  }

  /**
   * After a state change that may have filtered the current message out of the
   * selected box, re-render and advance to the next matching message.
   */
  private advanceAfterStateChange(msg: Message): void {
    this.regroupAndSync();
    const next = this.findNextBoxMatch(msg.id || null);
    this.renderCurrentView();
    if (next) {
      void this.showMessageInViewer(next);
    } else {
      this.selectedMessageId = null;
      this.messageViewer.clear();
    }
  }

  /** Find the next message matching the selected box, in the active sort order. */
  private findNextBoxMatch(currentId: string | null): Message | null {
    const box = this.getSelectedBox();
    if (!box) return null;
    const dir = this.sortNewestFirst ? -1 : 1;
    const list = [...this.messageState.messages].sort(
      (a, b) =>
        dir *
        (new Date(a.receivedDateTime || 0).getTime() -
          new Date(b.receivedDateTime || 0).getTime()),
    );
    const inFlight = this.classifier.getInFlightIds();
    for (const m of list) {
      if (m.id === currentId) continue;
      if (this.messageMatchesBox(m, box, inFlight)) return m;
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

  /** Move a single message to the provider's trash folder. Optimistically
   *  removes it from the list; restores on API failure. */
  private handleDeleteMessage(msg: Message): void {
    if (!msg.id) return;
    const id = msg.id;

    const snapshot = [...this.messageState.messages];
    this.messageState.messages = this.messageState.messages.filter((m) => m.id !== id);
    this.refreshListCache();

    const next = this.findNextUnread(id) ?? this.findNextMessage(id);

    this.regroupAndSync();
    this.renderCurrentView();
    if (next) {
      void this.showMessageInViewer(next);
    } else {
      this.selectedMessageId = null;
      this.messageViewer.clear();
    }

    void this.plugin.mailApi.deleteMessage(id).catch((err) => {
      logger.warn("InboxView", "Delete failed, restoring", err);
      this.messageState.messages = snapshot;
      this.refreshListCache();
      this.regroupAndSync();
      this.renderCurrentView();
      const errMsg = err instanceof Error ? err.message : String(err);
      new Notice(`Couldn't delete on server — restored (${errMsg}).`);
    });
  }

  /** Batch-delete selected messages. Each request fires in parallel;
   *  individual failures restore only that message. */
  private handleBatchDelete(ids: Set<string>): void {
    const victims = this.messageState.messages.filter((m) => m.id && ids.has(m.id));
    if (victims.length === 0) return;

    const victimIds = new Set(victims.map((m) => m.id!));
    this.messageState.messages = this.messageState.messages.filter((m) => !m.id || !victimIds.has(m.id));
    this.refreshListCache();

    this.regroupAndSync();
    this.messageList.clearMultiSelection();
    this.messageViewer.clear();
    this.renderCurrentView();

    const api = this.plugin.mailApi;
    for (const msg of victims) {
      const id = msg.id!;
      void api.deleteMessage(id).catch((err) => {
        logger.warn("InboxView", `Delete failed for ${id}, restoring`, err);
        if (!this.messageState.messages.some((m) => m.id === id)) {
          this.messageState.messages.push(msg);
          this.messageState.messages.sort(
            (a, b) =>
              new Date(b.receivedDateTime ?? 0).getTime() -
              new Date(a.receivedDateTime ?? 0).getTime(),
          );
          this.refreshListCache();
          this.regroupAndSync();
          this.renderCurrentView();
        }
        const errMsg = err instanceof Error ? err.message : String(err);
        new Notice(`Couldn't delete "${msg.subject || "(no subject)"}" — restored (${errMsg}).`);
      });
    }
  }

  /** Update the cached message list to match current in-memory state. */
  private refreshListCache(): void {
    const showRead = this.plugin.settings.showReadEmails;
    this.plugin.store.setMessageList(
      "", showRead,
      this.messageState.messages,
      this.messageState.nextLink,
    );
  }

  /** Find the next message after currentId in the active sort order. */
  private findNextMessage(currentId: string): Message | null {
    const dir = this.sortNewestFirst ? -1 : 1;
    const list = [...this.messageState.messages].sort(
      (a, b) =>
        dir *
        (new Date(a.receivedDateTime || 0).getTime() -
          new Date(b.receivedDateTime || 0).getTime()),
    );
    const idx = list.findIndex((m) => m.id === currentId);
    if (idx === -1) return list[0] ?? null;
    return list[idx + 1] ?? list[idx - 1] ?? null;
  }

  private handleBatchTag(ids: Set<string>, tag: string): void {
    const msgIds = this.resolveBatchMessageIds(ids);
    for (const msgId of msgIds) {
      if (this.plugin.store.getTags(msgId).some((e) => e.tag === tag)) continue;
      this.plugin.store.setTag(msgId, tag, "manual");
    }
    this.messageList.clearMultiSelection();
    this.messageViewer.clear();
    this.scheduleRender({ tags: true, regroup: true });
  }

  /** Bulk deny tag: remove tag from all selected, merge into formula, refine prompt. */
  private async handleBulkDenyTag(ids: Set<string>, tag: string): Promise<void> {
    const s = this.plugin.settings;
    if (!hasClaudeAccess(s.anthropicApiKey)) return;

    const msgIds = this.resolveBatchMessageIds(ids);
    const contents: string[] = [];

    // Remove the denied tag from each message immediately
    for (const msgId of msgIds) {
      this.plugin.store.removeTag(msgId, tag);

      const msg = this.messageState.messages.find((m) => m.id === msgId);
      if (msg) {
        const content = this.getClassifiableContent(msg);
        if (content) contents.push(content);
      }
    }

    this.messageList.clearMultiSelection();
    this.messageViewer.clear();
    this.scheduleRender({ tags: true, regroup: true });

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
    this.messageViewer.render(msg, stripped);

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
    if (this.summarizeInflight.has(msgId)) {
      this.messageViewer.showProcessingIndicator();
      const waitStart = Date.now();
      while (this.summarizeInflight.has(msgId) && Date.now() - waitStart < 10000) {
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

    this.markSummarizing(msgId, true);
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
      })
      .catch((err) => {
        const errMsg = err instanceof Error ? err.message : String(err);
        new Notice(`Claude processing failed: ${errMsg}`);
        if (this.selectedMessageId === msgId) {
          this.messageViewer.hideProcessingIndicator();
        }
      })
      .finally(() => {
        this.markSummarizing(msgId, false);
      });
  }

  private async handleSearch(query: string): Promise<void> {
    this.messageState.searchQuery = query;
    await this.loadMessages("");
  }

  private isLoadingMore = false;

  private async handleLoadMore(): Promise<void> {
    if (!this.messageState.nextLink) return;
    if (this.isLoadingMore) return;
    this.isLoadingMore = true;

    // Graph's `$filter=receivedDateTime ge X` can return an empty `value`
    // with a non-null `@odata.nextLink` while it scans older pages. Drain
    // those empty pages in one call so the sentinel doesn't flicker.
    const MAX_EMPTY_PAGES = 10;
    let emptyPages = 0;
    const seen = new Set(this.messageState.messages.map((m) => m.id));
    const collected: Message[] = [];

    try {
      const showRead = this.plugin.settings.showReadEmails;
      const since = this.plugin.getSyncSince();
      const searchQuery = this.messageState.searchQuery || undefined;
      while (this.messageState.nextLink) {
        const response: MailListResponse<Message> =
          await this.plugin.mailApi.listMessages("", {
            nextLink: this.messageState.nextLink,
            top: this.plugin.settings.pageSize,
            unreadOnly: !showRead,
            search: searchQuery,
            since,
          });
        this.messageState.nextLink = response.nextLink;

        const fresh = response.value.filter((m) => m.id && !seen.has(m.id));
        for (const m of fresh) seen.add(m.id!);
        collected.push(...fresh);

        if (fresh.length > 0) break;
        emptyPages++;
        if (emptyPages >= MAX_EMPTY_PAGES) break;
      }

      if (collected.length > 0) {
        this.plugin.store.applyReadState(collected);
        this.messageState.messages = [...this.messageState.messages, ...collected];
      }

      this.regroupAndSync();
      this.renderCurrentView();
      if (collected.length > 0) this.startBackgroundProcessing();
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      new Notice(`Failed to load more messages: ${msg}`);
    } finally {
      this.isLoadingMore = false;
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
    s.sortNewestFirst = this.sortNewestFirst;
    s.selectedBoxId = this.selectedBoxId;
    void this.plugin.saveSettings();
  }

  /** Compute the filtered, sort-ordered list as currently displayed in the view.
   *  Pinned messages always float to the top of the visible list, regardless
   *  of which box is selected or the active date sort. Within the pinned and
   *  non-pinned groups, the date sort applies normally. */
  private getVisibleMessages(): Message[] {
    const filtered = this.applyMessageFilters(this.messageState.messages);
    const dir = this.sortNewestFirst ? -1 : 1;
    const byDate = (a: Message, b: Message) =>
      dir *
      (new Date(a.receivedDateTime || 0).getTime() -
        new Date(b.receivedDateTime || 0).getTime());
    return [...filtered].sort((a, b) => {
      const ap = a.id ? this.pinnedIds.has(a.id) : false;
      const bp = b.id ? this.pinnedIds.has(b.id) : false;
      if (ap !== bp) return ap ? -1 : 1;
      return byDate(a, b);
    });
  }

  private renderCurrentView(): void {
    if (!this.plugin.accounts.anySignedIn()) {
      this.messageList.renderLoggedOut(() => this.openIrisSettings());
      return;
    }

    const hasMore = !!this.messageState.nextLink;
    const sorted = this.getVisibleMessages();
    const box = this.getSelectedBox();
    const empty = box
      ? { icon: box.icon || "inbox", title: `No messages in ${box.name}` }
      : undefined;
    void this.messageList.renderFlatMessages(sorted, hasMore, empty);
  }

  /**
   * Gmail/Mail-style keyboard navigation. Applies only when the keyboard
   * focus is outside text entry widgets — so shortcuts don't hijack typing
   * in the search bar, modals, or contenteditable regions.
   */
  private handleKeyDown(evt: KeyboardEvent): void {
    const t = evt.target as HTMLElement | null;
    if (!t) return;
    if ((evt.ctrlKey || evt.metaKey) && !evt.altKey && !evt.shiftKey && (evt.key === "a" || evt.key === "A")) {
      evt.preventDefault();
      const ids = this.messageList.selectAll();
      this.handleMultiSelect(ids);
      return;
    }

    if (t.tagName === "INPUT" || t.tagName === "TEXTAREA" || t.isContentEditable) return;
    if (evt.ctrlKey || evt.metaKey || evt.altKey) return;

    switch (evt.key) {
      case "j":
      case "ArrowDown":
        evt.preventDefault();
        this.navigateMessage(1);
        return;
      case "k":
      case "ArrowUp":
        evt.preventDefault();
        this.navigateMessage(-1);
        return;
      case "e":
        evt.preventDefault();
        this.toggleReadOnSelected();
        return;
      case "p":
        evt.preventDefault();
        this.togglePinOnSelected();
        return;
      case "#":
      case "Delete":
        evt.preventDefault();
        this.deleteSelected();
        return;
      case "/":
        evt.preventDefault();
        this.searchBar.focus();
        return;
      case "f":
        evt.preventDefault();
        this.toggleCompactMode();
        return;
    }
  }

  /** `f` — toggle compact (single-pane) mode. The DOM class is the source of truth. */
  private toggleCompactMode(): void {
    this.contentEl.toggleClass(COMPACT_MODE_CLASS, !this.contentEl.hasClass(COMPACT_MODE_CLASS));
  }

  /** Move the selection up/down through the list as it's actually rendered. */
  private navigateMessage(dir: 1 | -1): void {
    const order = this.messageList.getRenderedOrder();
    if (order.length === 0) return;
    const currentIdx = this.selectedMessageId ? order.indexOf(this.selectedMessageId) : -1;
    const nextIdx = currentIdx === -1
      ? (dir === 1 ? 0 : order.length - 1)
      : Math.max(0, Math.min(order.length - 1, currentIdx + dir));
    const targetId = order[nextIdx];
    if (!targetId || targetId === this.selectedMessageId) return;
    const target = this.messageState.messages.find((m) => m.id === targetId);
    if (target) void this.showMessageInViewer(target);
  }

  private toggleReadOnSelected(): void {
    const msg = this.getSelectedMessage();
    if (!msg) return;
    if (msg.isRead) this.handleMarkAsUnread(msg);
    else this.handleMarkAsRead(msg);
  }

  private deleteSelected(): void {
    const msg = this.getSelectedMessage();
    if (msg) this.handleDeleteMessage(msg);
  }

  private togglePinOnSelected(): void {
    const msg = this.getSelectedMessage();
    if (msg) this.handleTogglePin(msg);
  }

  private getSelectedMessage(): Message | null {
    if (!this.selectedMessageId) return null;
    return this.messageState.messages.find((m) => m.id === this.selectedMessageId) ?? null;
  }

  /**
   * Kick off all background AI processing as a hardcoded three-stage pipeline:
   *   1. Summarize — Claude body extraction for the prefetch window.
   *   2. Junk tagging — classifier pass restricted to tags wired into the
   *      Junk box, so the Junk view fills before other tags are tried.
   *   3. Other — remaining tag classification + nickname generation.
   * Each stage awaits the previous one so the Secretary view reflects the
   * user-visible ordering rather than everything racing in parallel.
   */
  private startBackgroundProcessing(): void {
    this.prefetchAllPromise = this.prefetchAllProcessed();
    this.prefetchAllPromise.catch((err) =>
      logger.warn("InboxView", "Background prefetch failed", err),
    );

    (async () => {
      try { await this.prefetchAllPromise; } catch { /* proceed */ }

      const allCandidates = this.getTagCandidates();
      const junkTagSet = this.getJunkBoxTagSet();
      const junkCandidates = allCandidates.filter((c) => junkTagSet.has(c.name));
      const otherCandidates = allCandidates.filter((c) => !junkTagSet.has(c.name));

      if (junkCandidates.length > 0) {
        await this.classifier.autoTagMessages(
          this.messageState.messages,
          junkCandidates,
          () => this.syncTagCacheViews(),
        );
      }

      if (otherCandidates.length > 0) {
        await this.classifier.autoTagMessages(
          this.messageState.messages,
          otherCandidates,
          () => this.syncTagCacheViews(),
        );
      }

      await this.generateAllNicknames();
    })().catch((err) => logger.warn("InboxView", "Background pipeline failed", err));

    // When Claude processing is disabled but forwarded-sender resolution is
    // on, bodies are never prefetched by prefetchAllProcessed().  Fetch them
    // here so originalSender gets extracted and the sender list updates.
    const s = this.plugin.settings;
    if (s.resolveForwardedSender && (!s.enableClaudeProcessing || !hasClaudeAccess(s.anthropicApiKey))) {
      void this.prefetchBodiesForSenderResolution();
    }
  }

  /** Tag names wired into the Junk box's predicate (empty if none). */
  private getJunkBoxTagSet(): Set<string> {
    const box = (this.plugin.settings.boxes || []).find((b) => b.builtin === "junk");
    return new Set(box?.tags || []);
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

  /** Open a modal to create or edit an automation rule for a sender. */
  private openSenderRuleModal(address: string, rawName: string): void {
    const key = address.toLowerCase();
    const s = this.plugin.settings;
    const current = s.senderRules?.[key];
    new SenderRuleModal(
      this.plugin.app,
      address,
      this.resolveName(address, rawName),
      this.getTagCategories(),
      current,
      (addr, rule) => {
        const k = addr.toLowerCase();
        const rules = { ...(this.plugin.settings.senderRules || {}) };
        rules[k] = rule;
        this.plugin.settings.senderRules = rules;
        this.plugin.scheduleSaveSettings();
        this.applySenderRules();
        new Notice(`Rule saved for ${addr}.`);
      },
      (addr) => {
        const k = addr.toLowerCase();
        const rules = { ...(this.plugin.settings.senderRules || {}) };
        if (!(k in rules)) return;
        delete rules[k];
        this.plugin.settings.senderRules = rules;
        this.plugin.scheduleSaveSettings();
        new Notice(`Rule removed for ${addr}.`);
      },
    ).open();
  }

  /** Add or update a sender rule with autoBin=true (preserving any existing autoTag). */
  private markSenderAsJunk(address: string, rawName: string): void {
    if (!address) return;
    const key = address.toLowerCase();
    const rules = { ...(this.plugin.settings.senderRules || {}) };
    const existing = rules[key] || {};
    if (existing.autoBin) {
      new Notice(`${this.resolveName(address, rawName)} is already marked as junk.`);
      return;
    }
    rules[key] = { ...existing, autoBin: true };
    this.plugin.settings.senderRules = rules;
    this.plugin.scheduleSaveSettings();
    this.applySenderRules();
    new Notice(`${this.resolveName(address, rawName)} marked as junk sender.`);
  }

  /** Apply all active sender rules to the currently-loaded messages.
   *  Runs after loadMessages and when a rule is saved. */
  private applySenderRules(): void {
    const rules = this.plugin.settings.senderRules;
    if (!rules || Object.keys(rules).length === 0) return;

    const toBin: Message[] = [];
    const toTag: { msgId: string; tag: string }[] = [];

    for (const msg of this.messageState.messages) {
      const addr = msg.from?.emailAddress?.address?.toLowerCase();
      if (!addr) continue;
      const rule = rules[addr];
      if (!rule) continue;

      if (rule.autoBin) {
        toBin.push(msg);
        continue; // Skip tagging — message is about to be deleted.
      }
      if (rule.autoTag && msg.id) {
        if (!this.plugin.store.getTags(msg.id).some((e) => e.tag === rule.autoTag)) {
          toTag.push({ msgId: msg.id, tag: rule.autoTag });
        }
      }
    }

    for (const { msgId, tag } of toTag) {
      this.plugin.store.setTag(msgId, tag, "manual");
    }

    if (toBin.length > 0) {
      const victimIds = new Set(toBin.map((m) => m.id!).filter(Boolean));
      this.messageState.messages = this.messageState.messages.filter(
        (m) => !m.id || !victimIds.has(m.id),
      );
      this.refreshListCache();
      this.scheduleRender({ tags: toTag.length > 0, regroup: true, list: true });

      const api = this.plugin.mailApi;
      for (const msg of toBin) {
        const id = msg.id;
        if (!id) continue;
        void api.deleteMessage(id).catch((err) => {
          logger.warn("InboxView", `Auto-bin failed for ${id}`, err);
        });
      }
    } else if (toTag.length > 0) {
      this.scheduleRender({ tags: true, list: true });
    }
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
  private getEffectiveSender(msg: Message) {
    return getEffectiveSender(this.plugin, msg);
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
    this.scheduleRender({ regroup: true, list: true });
  }

  /** Apply the selected box's predicate to a message list. */
  private applyMessageFilters(messages: Message[]): Message[] {
    const box = this.getSelectedBox();
    if (!box) return messages;
    const inFlight = this.classifier.getInFlightIds();
    return messages.filter((m) => this.messageMatchesBox(m, box, inFlight));
  }

  // --- Private: tag management ---

  private getHiddenListTagSet(): Set<string> {
    const map = this.plugin.settings.tagHiddenInList || {};
    return new Set(Object.keys(map).filter((k) => map[k]));
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

  private getTagColorMap(): Map<string, string> {
    const colors = this.plugin.settings.tagColors || {};
    return new Map(Object.entries(colors));
  }

  private getTagDescriptionMap(): Map<string, string> {
    const descriptions = this.plugin.settings.tagDescriptions || {};
    return new Map(Object.entries(descriptions).filter(([, v]) => !!v && v.trim().length > 0));
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

    const existing = this.plugin.store.getTags(msg.id);
    const removedAutoTags: string[] = [];

    if (tag === null) {
      // Remove all tags — track auto-tags for prompt refinement
      for (const e of existing) {
        if (e.source === "auto") removedAutoTags.push(e.tag);
      }
      this.plugin.store.removeTag(msg.id);
    } else {
      // Toggle: if tag already present, remove it; otherwise add it
      const has = existing.find((e) => e.tag === tag);
      if (has) {
        if (has.source === "auto") removedAutoTags.push(tag);
        this.plugin.store.removeTag(msg.id, tag);
      } else {
        this.plugin.store.setTag(msg.id, tag, "manual");
      }
    }

    this.syncTagCacheViews();
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

    const unread = this.messageState.messages.filter((m) => !m.isRead);
    const candidates = limit === -1 ? unread : unread.slice(0, limit);

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
      this.markSummarizing(msgId, true);
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
        this.markSummarizing(msgId, false);
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
      // Republish so widgets re-render with the now-resolved original senders.
      this.plugin.setInboxMessages(this.messageState.messages);
    }
  }

  /** Re-push tag icons/colors/categories and refresh the list after the tags
   *  modal mutates settings. */
  private handleTagsSettingsChanged(): void {
    this.messageViewer.setTagCategories(this.getTagCategories());
    this.messageViewer.setTagIcons(this.getTagIconMap());
    this.messageViewer.setTagColors(this.getTagColorMap());
    this.messageViewer.setTagDescriptions(this.getTagDescriptionMap());
    this.messageList.setTagIcons(this.getTagIconMap());
    this.messageList.setTagColors(this.getTagColorMap());
    this.messageList.setTagDescriptions(this.getTagDescriptionMap());
    this.messageList.setHiddenListTags(this.getHiddenListTagSet());
    this.scheduleRender({ tags: true, regroup: true, list: true });
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
    const existing = this.plugin.store.getTags(msg.id);
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
        s.tagContradictions || {},
        s.tagPrecludes || {},
      );

      for (const tag of tags) {
        this.plugin.store.setTag(msg.id, tag, "auto", getTagVersion(s.tagPromptVersions, tag));
      }
    } catch (err) {
      logger.warn("InboxView", "Re-tag failed", err);
    }

    this.syncTagCacheViews();
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

  syncBadge(): void {
    if (!this.plugin.settings.badgeCount) {
      this.plugin.updateBadge(0);
      return;
    }
    let inboxCount = 0;
    for (const m of this.messageState.messages) {
      if (!m.isRead) inboxCount++;
    }
    this.plugin.updateBadge(inboxCount);
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
