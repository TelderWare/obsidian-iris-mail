import { Plugin, WorkspaceLeaf, Notice } from "obsidian";
import { InboxView } from "./views/InboxView";
import { IrisMailSettingsTab } from "./settings/SettingsTab";
import { AccountRegistry } from "./auth/AccountRegistry";
import { MailDispatcher } from "./mail/MailDispatcher";
import { EmailStore } from "./store/EmailStore";
import {
  buildIrisHomepageWidgets,
  type IrisHomepageWidgetDescriptor,
} from "./widgets/IrisHomepageWidgets";
import {
  VIEW_TYPE_IRIS_MAIL,
  ICON_NAME,
  DEFAULT_SETTINGS,
  CACHE_STORAGE_KEY,
  freshDefaultBoxes,
} from "./constants";
import { logger, setDebugEnabled } from "./utils/logger";
import { setRelayApp } from "./utils/claudeApi";
import { newAccountId } from "./utils/compositeId";
import { encryptString, decryptString } from "./utils/safeStorage";
import type {
  Account,
  AuthMethod,
  IrisMailSettings,
  Message,
  MailProvider,
} from "./types";

/**
 * The Anthropic API key is stored with an "enc:" sentinel so older plaintext
 * values are still readable on upgrade. Once round-tripped, future writes will
 * always carry the prefix.
 */
function encryptApiKey(key: string): string {
  if (!key) return "";
  return "enc:" + encryptString(key);
}

function decryptApiKey(stored: string): string {
  if (!stored) return "";
  if (!stored.startsWith("enc:")) return stored;
  try {
    return decryptString(stored.slice(4));
  } catch {
    logger.warn("Main", "Failed to decrypt API key");
    return "";
  }
}

/**
 * Public API shape for to-do messages, exposed to other plugins (Iris Tasks).
 * Stable contract — additive changes only.
 */
export interface IrisMailTodo {
  /** Composite message id (`{accountId}:{nativeId}`). */
  id: string;
  subject: string;
  /** Display name of the sender (falls back to address). */
  from: string;
  fromAddress: string;
  /** ISO timestamp from Graph. May be empty if envelope is unavailable. */
  receivedDateTime: string;
  /** Local timestamp at which the message was flagged as a to-do. */
  todoAt: number;
  accountLabel?: string;
}

export default class IrisMailPlugin extends Plugin {
  settings: IrisMailSettings = DEFAULT_SETTINGS;
  accounts!: AccountRegistry;
  mailApi!: MailDispatcher;
  store!: EmailStore;
  private refreshIntervalId: number | null = null;
  private ribbonIconEl: HTMLElement | null = null;
  private ribbonBadgeEl: HTMLElement | null = null;
  private saveSettingsTimer: ReturnType<typeof setTimeout> | null = null;
  private lastBackgroundRefreshAt = 0;
  private visibilityHandler: (() => void) | null = null;

  /** Most recent inbox snapshot. Shared with homepage widgets and other
   *  integrations that need a read-only view of what's currently in the UI. */
  private inboxMessages: Message[] = [];
  private messagesChangedListeners = new Set<() => void>();
  private todosChangedListeners = new Set<() => void>();
  private prefetchingForwardBodies = false;

  async onload(): Promise<void> {
    await this.loadSettings();
    setRelayApp(this.app);

    this.store = new EmailStore(this);
    await this.store.load();

    // When a summary lands for a to-do-flagged message, the message becomes
    // eligible for `getTodos()` — nudge subscribers (Iris Tasks) to re-pull.
    this.register(
      this.store.onProcessedChanged((messageId) => {
        if (this.store.isMarkedTodo(messageId)) this.notifyTodosChanged();
      }),
    );

    this.mailApi = new MailDispatcher();
    this.accounts = new AccountRegistry(this.mailApi);
    await this.accounts.initializeAll(this.settings.accounts, this.settings);

    this.registerView(
      VIEW_TYPE_IRIS_MAIL,
      (leaf: WorkspaceLeaf) => new InboxView(leaf, this),
    );

    this.ribbonIconEl = this.addRibbonIcon(ICON_NAME, "Email", () => {
      this.activateView();
    });
    this.ribbonIconEl.addClass("iris-ribbon-icon");

    this.addCommand({
      id: "open-iris-mail",
      name: "Open Iris Mail",
      callback: () => this.activateView(),
    });

    this.addCommand({
      id: "iris-refresh",
      name: "Refresh Iris Mail",
      callback: () => this.refreshAllViews(),
    });

    this.addSettingTab(new IrisMailSettingsTab(this.app, this));

    this.startAutoRefresh();

    if (this.accounts.anySignedIn()) {
      void this.backgroundRefresh();
    }
  }

  onunload(): void {
    if (this.saveSettingsTimer !== null) {
      clearTimeout(this.saveSettingsTimer);
      this.saveSettingsTimer = null;
      void this.saveSettings();
    }
    this.app.workspace.detachLeavesOfType(VIEW_TYPE_IRIS_MAIL);
    this.accounts?.destroyAll();
    this.store?.flush();
    this.stopAutoRefresh();
  }

  async activateView(): Promise<void> {
    const existing =
      this.app.workspace.getLeavesOfType(VIEW_TYPE_IRIS_MAIL);
    if (existing.length > 0) {
      this.app.workspace.revealLeaf(existing[0]);
      return;
    }
    const leaf = this.app.workspace.getLeaf(false);
    await leaf.setViewState({
      type: VIEW_TYPE_IRIS_MAIL,
      active: true,
    });
    this.app.workspace.revealLeaf(leaf);
  }

  // --- Account management ---

  /** Create a new account from the partial fields supplied by Settings UI. */
  async createAccount(input: { label: string; provider: MailProvider }): Promise<Account> {
    const defaultLabel = input.provider === "imap" ? "IMAP" : "Outlook";
    const account: Account = {
      id: newAccountId(),
      label: input.label || defaultLabel,
      provider: input.provider,
      enabled: true,
      ...(input.provider === "imap" ? { imapSecure: true, imapPort: 993 } : {}),
    };
    this.settings.accounts = [...this.settings.accounts, account];
    await this.saveSettings();
    await this.accounts.add(account, this.settings);
    return account;
  }

  /** Persist an edit to an existing account (label or credentials). */
  async updateAccount(updated: Account): Promise<void> {
    this.settings.accounts = this.settings.accounts.map(
      (a) => a.id === updated.id ? updated : a,
    );
    await this.saveSettings();
    this.accounts.updateAccount(updated);
  }

  async removeAccount(accountId: string): Promise<void> {
    await this.accounts.remove(accountId);
    this.settings.accounts = this.settings.accounts.filter((a) => a.id !== accountId);
    await this.saveSettings();
    this.refreshAllViews();
  }

  async loginAccount(accountId: string, method?: AuthMethod): Promise<void> {
    const entry = this.accounts.get(accountId);
    if (!entry) throw new Error(`No such account: ${accountId}`);
    if (!this.accounts.hasCredentials(entry.account)) {
      new Notice(`${entry.account.label}: missing client credentials.`);
      return;
    }
    const useMethod = method || entry.account.authMethod || "auth-code";
    try {
      await entry.auth.initialize(this.settings);
      if (useMethod === "device-code" && entry.auth.loginWithDeviceCode) {
        await this.doDeviceCodeLogin(entry.auth);
      } else {
        await entry.auth.login(this.settings);
      }
      new Notice(`Signed in to ${entry.account.label}.`);
      this.refreshAllViews();
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      new Notice(`${entry.account.label} sign-in failed: ${msg}`);
    }
  }

  async logoutAccount(accountId: string): Promise<void> {
    const entry = this.accounts.get(accountId);
    if (!entry) return;
    await entry.auth.logout();
    new Notice(`Signed out of ${entry.account.label}.`);
    this.refreshAllViews();
  }

  private async doDeviceCodeLogin(auth: { loginWithDeviceCode?: (s: IrisMailSettings, cb: (code: string, uri: string) => void) => Promise<void> }): Promise<void> {
    if (!auth.loginWithDeviceCode) {
      throw new Error("Provider does not support device-code sign-in.");
    }
    let codeNotice: Notice | undefined;
    let copyNotice: Notice | undefined;
    try {
      await auth.loginWithDeviceCode(
        this.settings,
        (code, verificationUri) => {
          const frag = document.createDocumentFragment();
          frag.appendText("Go to ");
          const link = frag.createEl("a", { text: verificationUri, href: verificationUri });
          link.addEventListener("click", (e) => {
            e.preventDefault();
            navigator.clipboard.writeText(code);
            window.open(verificationUri);
            copyNotice = new Notice(`Code copied: ${code}`, 0);
          });
          frag.appendText(" and enter code: ");
          const codeEl = frag.createEl("b", { text: code });
          codeEl.style.cursor = "pointer";
          codeEl.title = "Click to copy";
          codeEl.addEventListener("click", () => {
            navigator.clipboard.writeText(code);
            new Notice("Code copied to clipboard.");
          });
          codeNotice = new Notice(frag, 0);
        },
      );
    } finally {
      codeNotice?.hide();
      copyNotice?.hide();
    }
  }

  // --- Settings persistence ---

  async loadSettings(): Promise<void> {
    const data = await this.loadData();
    const { __cache: _ignored, ...settingsData } = (data ?? {}) as Record<string, unknown>;
    this.settings = Object.assign({}, DEFAULT_SETTINGS, settingsData);

    if (this.settings.anthropicApiKey) {
      this.settings.anthropicApiKey = decryptApiKey(this.settings.anthropicApiKey);
    }

    // Migrate stale model IDs
    const modelFixes: Record<string, string> = {
      "claude-sonnet-4-6-20250514": "claude-sonnet-4-6",
      "claude-opus-4-6-20250514": "claude-opus-4-6",
    };
    const fixedModel = modelFixes[this.settings.claudeModel];
    if (fixedModel) {
      this.settings.claudeModel = fixedModel;
    }

    setDebugEnabled(!!this.settings.debugLogging);

    // Migrate legacy string badgeCount ("off"/"unread"/"total") → boolean.
    const legacyBadge = this.settings.badgeCount as unknown;
    if (typeof legacyBadge === "string") {
      this.settings.badgeCount = legacyBadge !== "off";
    }

    // Drop legacy triage fields if present
    const legacy = this.settings as unknown as Record<string, unknown>;
    delete legacy.triageTree;
    delete legacy.triageGraph;
    delete legacy.triageTreeVersion;
    delete legacy.enableTriage;
    delete legacy.filterUnreadOnly;

    // Seed boxes on first load, and backfill any missing builtin boxes so
    // upgrades always end up with the full default set available.
    if (!Array.isArray(this.settings.boxes) || this.settings.boxes.length === 0) {
      this.settings.boxes = freshDefaultBoxes();
    } else {
      // Drop the short-lived "pinned" built-in box — pinning now applies
      // across every box rather than having its own view.
      this.settings.boxes = this.settings.boxes.filter(
        (b) => (b.builtin as string | undefined) !== "pinned",
      );
      const existing = new Set(this.settings.boxes.map((b) => b.builtin).filter(Boolean));
      for (const def of freshDefaultBoxes()) {
        if (def.builtin && !existing.has(def.builtin)) this.settings.boxes.push(def);
      }
      // Opt the built-in To-do box into `saved` by default, but only when the
      // user hasn't already made an explicit choice (undefined vs. false).
      for (const b of this.settings.boxes) {
        if (b.builtin === "todo" && b.saved === undefined) b.saved = true;
      }
    }
    if (!this.settings.selectedBoxId || !this.settings.boxes.some((b) => b.id === this.settings.selectedBoxId)) {
      this.settings.selectedBoxId = "in";
    }

    // Migrate single-account settings → accounts[]
    const migrated = this.migrateSingleAccountSettings(settingsData as Record<string, unknown>);
    if (migrated || fixedModel) {
      await this.saveSettings();
    }
  }

  /**
   * If the settings file is from the single-account era, build one Account from
   * the old top-level fields and rename the corresponding token cache key in
   * localStorage so the user doesn't have to sign in again.
   *
   * Returns true if any migration ran.
   */
  private migrateSingleAccountSettings(rawSettings: Record<string, unknown>): boolean {
    if (this.settings.accounts.length > 0) return false;

    const provider = rawSettings.provider as string | undefined;
    const oldClientId = rawSettings.clientId as string | undefined;
    if (!provider && !oldClientId) return false;

    const account: Account = {
      id: newAccountId(),
      label: "Outlook",
      provider: "outlook",
      enabled: true,
      clientId: rawSettings.clientId as string | undefined,
      authority: rawSettings.authority as string | undefined,
      authMethod: rawSettings.authMethod as AuthMethod | undefined,
    };
    this.settings.accounts = [account];

    // Strip the legacy keys off the in-memory settings object.
    const s = this.settings as unknown as Record<string, unknown>;
    delete s.provider;
    delete s.clientId;
    delete s.authority;
    delete s.authMethod;

    // Rename the localStorage token cache so the user stays signed in.
    try {
      const old = localStorage.getItem(CACHE_STORAGE_KEY);
      if (old !== null) {
        localStorage.setItem(`${CACHE_STORAGE_KEY}:${account.id}`, old);
        localStorage.removeItem(CACHE_STORAGE_KEY);
      }
    } catch (err) {
      logger.warn("Main", "Failed to migrate token storage key", err);
    }

    logger.info("Main", `Migrated legacy single-account settings to account ${account.id}`);
    return true;
  }

  async saveSettings(): Promise<void> {
    const toSave = { ...this.settings };
    if (toSave.anthropicApiKey && !toSave.anthropicApiKey.startsWith("enc:")) {
      toSave.anthropicApiKey = encryptApiKey(toSave.anthropicApiKey);
    }
    const existing = ((await this.loadData()) ?? {}) as Record<string, unknown>;
    const merged: Record<string, unknown> = { ...existing, ...toSave };
    if (existing.__cache !== undefined) merged.__cache = existing.__cache;
    // Strip any lingering legacy single-account fields.
    delete merged.provider;
    delete merged.clientId;
    delete merged.authority;
    delete merged.authMethod;
    await this.saveData(merged);
  }

  scheduleSaveSettings(): void {
    if (this.saveSettingsTimer !== null) {
      clearTimeout(this.saveSettingsTimer);
    }
    this.saveSettingsTimer = setTimeout(() => {
      this.saveSettingsTimer = null;
      void this.saveSettings();
    }, 1000);
  }

  refreshAllViews(): void {
    for (const leaf of this.app.workspace.getLeavesOfType(VIEW_TYPE_IRIS_MAIL)) {
      (leaf.view as InboxView).refresh();
    }
  }

  /**
   * Fetch inbox messages and update the badge even when no InboxView is open.
   * If a view is open, delegates to the full view refresh instead.
   */
  /** Compute the `since` lower-bound Date from settings, or undefined when unlimited. */
  getSyncSince(): Date | undefined {
    const days = this.settings.initialSyncLookbackDays;
    if (!days || days <= 0) return undefined;
    return new Date(Date.now() - days * 24 * 60 * 60 * 1000);
  }

  async backgroundRefresh(): Promise<void> {
    this.lastBackgroundRefreshAt = Date.now();
    void this.syncLocalReadStateToServer();

    const leaves = this.app.workspace.getLeavesOfType(VIEW_TYPE_IRIS_MAIL);
    if (leaves.length > 0) {
      this.refreshAllViews();
      return;
    }

    if (!this.accounts.anySignedIn()) return;

    try {
      // The dispatcher merges every account's inbox; the folderId arg is
      // ignored because each provider resolves its own inbox internally.
      const response = await this.mailApi.listMessages("", {
        top: this.settings.pageSize,
        unreadOnly: !this.settings.showReadEmails,
        since: this.getSyncSince(),
      });

      const merged = this.store.mergePersistedMessages(
        response.value,
        this.getSavedBoxes(),
      );
      this.store.applyReadState(merged);
      this.updateBadgeFromMessages(merged);
      this.setInboxMessages(merged);

      // Widgets can render the inbox without the full view ever opening, so
      // forwarded-sender resolution needs bodies fetched here too.
      void this.prefetchForwardBodies();
    } catch (err) {
      logger.error("Main", "Background refresh failed", err);
    }
  }

  /**
   * Fetch bodies for forwarded messages so the effective-sender resolver can
   * pull `originalSender` out of the cache. InboxView has its own equivalent
   * for the view-open path; this method covers the widget-only / no-view
   * case. Runs at most one instance at a time.
   */
  async prefetchForwardBodies(): Promise<void> {
    if (!this.settings.resolveForwardedSender) return;
    if (!this.accounts.anySignedIn()) return;
    if (this.prefetchingForwardBodies) return;
    this.prefetchingForwardBodies = true;
    let fetched = false;
    try {
      for (const msg of this.inboxMessages) {
        if (!msg.id) continue;
        if (this.store.getBody(msg.id)) continue;
        if (!/^(?:fw|fwd)\s*:/i.test(msg.subject || "")) continue;
        try {
          const full = await this.mailApi.getMessageBody(msg.id);
          const body = full.body?.content || "";
          if (body) {
            this.store.setBody(msg, body);
            fetched = true;
          }
        } catch (err) {
          logger.warn("Main", `Forward body prefetch failed for ${msg.id}`, err);
        }
      }
    } finally {
      this.prefetchingForwardBodies = false;
    }
    if (fetched) this.notifyMessagesChanged();
  }

  private notifyMessagesChanged(): void {
    for (const cb of Array.from(this.messagesChangedListeners)) {
      try { cb(); } catch (err) { logger.warn("Main", "messagesChanged listener failed", err); }
    }
  }

  // --- Inbox snapshot (shared with Iris Homepage widgets) ---

  getInboxMessages(): readonly Message[] {
    return this.inboxMessages;
  }

  /** All boxes whose `saved` flag is on — used by store envelope retention. */
  getSavedBoxes(): import("./types").Box[] {
    return (this.settings.boxes || []).filter((b) => b.saved);
  }

  /** Publish a new inbox snapshot and notify subscribers. Called by the
   *  InboxView after each load, and by backgroundRefresh when no view is open. */
  setInboxMessages(messages: Message[]): void {
    this.inboxMessages = messages;
    this.notifyMessagesChanged();
  }

  /** Subscribe to inbox-snapshot updates. Returns an unsubscribe fn. */
  onMessagesChanged(cb: () => void): () => void {
    this.messagesChangedListeners.add(cb);
    return () => {
      this.messagesChangedListeners.delete(cb);
    };
  }

  // --- Public to-do API (consumed by Iris Tasks) ---

  /**
   * Snapshot of currently flagged to-do messages, resolved against the live
   * inbox snapshot when possible and the persisted-envelope cache otherwise
   * (so messages that have aged out of the sync window still surface).
   * Returned in `todoAt` descending order so newest flags appear first.
   *
   * Only messages that already have a Claude-generated summary in the
   * processed cache are emitted — Iris Tasks must never see a to-do whose
   * body it can't access in summarised form. Flagged messages without a
   * summary are held back until prefetch lands one, at which point the
   * processed-cache listener fires `notifyTodosChanged()` and Iris Tasks
   * re-pulls.
   */
  getTodos(): IrisMailTodo[] {
    const ids = this.store.getAllTodoIds();
    if (ids.size === 0) return [];

    const liveById = new Map<string, Message>();
    for (const m of this.inboxMessages) {
      if (m.id) liveById.set(m.id, m);
    }

    const accountLabelById = new Map<string, string>();
    for (const acc of this.settings.accounts ?? []) {
      accountLabelById.set(acc.id, acc.label);
    }

    const out: IrisMailTodo[] = [];
    for (const id of ids) {
      if (!this.store.getProcessed(id)?.processedMarkdown) continue;
      const msg = liveById.get(id) ?? this.store.getPersistedEnvelope(id);
      const todoAt = this.store.getTodoAt(id) ?? 0;
      const accountId = id.split(":")[0];
      out.push({
        id,
        subject: msg?.subject ?? "(no subject)",
        from: msg?.from?.emailAddress?.name ?? msg?.from?.emailAddress?.address ?? "",
        fromAddress: msg?.from?.emailAddress?.address ?? "",
        receivedDateTime: msg?.receivedDateTime ?? "",
        todoAt,
        accountLabel: msg?._accountLabel ?? accountLabelById.get(accountId),
      });
    }
    out.sort((a, b) => b.todoAt - a.todoAt);
    return out;
  }

  /**
   * Claude-summarised markdown for a to-do message, drawn from the processed
   * cache. Iris Tasks uses this when sending todos to Claude; we intentionally
   * never expose the raw email body across the plugin boundary so downstream
   * AI passes only see the summary. Returns an empty string when no summary
   * is available yet (InboxView prefetches summaries for visible messages).
   */
  getTodoBody(messageId: string): string {
    return this.store.getProcessed(messageId)?.processedMarkdown ?? "";
  }

  /** Clear a to-do flag (e.g. when Iris Tasks marks it complete). */
  clearTodo(messageId: string): void {
    if (!this.store.isMarkedTodo(messageId)) return;
    this.store.unmarkTodo(messageId);
    this.notifyTodosChanged();
  }

  /** Subscribe to to-do flag changes. Returns an unsubscribe fn. */
  onTodosChanged(cb: () => void): () => void {
    this.todosChangedListeners.add(cb);
    return () => {
      this.todosChangedListeners.delete(cb);
    };
  }

  /** Fire after any mark/unmark so subscribers can re-pull `getTodos()`. */
  notifyTodosChanged(): void {
    for (const cb of Array.from(this.todosChangedListeners)) {
      try { cb(); } catch (err) { logger.warn("Main", "todosChanged listener failed", err); }
    }
  }

  /** Widget provider consumed by the Iris Homepage plugin. Returns one
   *  descriptor per visible box (In, Read, To-do, Junk, Secretary, + user
   *  boxes). Called by the host — do not invoke from within this plugin. */
  irisHomepageWidgets(): IrisHomepageWidgetDescriptor[] {
    return buildIrisHomepageWidgets(this);
  }

  /**
   * Push locally-tracked read states to the server, then clear the local
   * markers (the server becomes the source of truth). Runs on every
   * backgroundRefresh tick rather than on every inbox load — the badge and
   * local view already reflect the read state immediately.
   */
  private async syncLocalReadStateToServer(): Promise<void> {
    if (!this.accounts.anySignedIn()) return;
    const ids = this.store.getLocallyReadIds();
    if (ids.length === 0) return;

    logger.info("Main", `Syncing ${ids.length} local read states to server`);
    await Promise.all(
      ids.map((id) =>
        this.mailApi.markAsRead(id).catch((err) => {
          // Message may have been deleted server-side — still clear it locally.
          logger.warn("Main", `Failed to sync read state for ${id}`, err);
        }),
      ),
    );
    for (const id of ids) this.store.clearLocalRead(id);
  }

  /** Compute the In-box badge count (unread messages) for a set of messages. */
  private updateBadgeFromMessages(messages: Message[]): void {
    if (!this.settings.badgeCount) {
      this.updateBadge(0);
      return;
    }
    const inboxCount = messages.filter((m) => !m.isRead).length;
    this.updateBadge(inboxCount);
  }

  updateBadge(inboxCount: number): void {
    if (!this.ribbonIconEl) return;

    if (inboxCount < 0) {
      for (const leaf of this.app.workspace.getLeavesOfType(VIEW_TYPE_IRIS_MAIL)) {
        (leaf.view as InboxView).syncBadge();
      }
      return;
    }

    const pos = this.settings.badgePosition;
    const enabled = pos !== "off";

    if (enabled && inboxCount > 0) {
      if (!this.ribbonBadgeEl) {
        this.ribbonBadgeEl = this.ribbonIconEl.createSpan({ cls: "iris-ribbon-badge" });
      }
      this.ribbonBadgeEl.className = `iris-ribbon-badge iris-ribbon-badge-${pos}`;
      this.ribbonBadgeEl.setText(inboxCount > 99 ? "99+" : String(inboxCount));
      this.ribbonBadgeEl.style.display = "";
    } else if (this.ribbonBadgeEl) {
      this.ribbonBadgeEl.style.display = "none";
    }
  }

  private startAutoRefresh(): void {
    this.stopAutoRefresh();
    const mins = this.settings.refreshIntervalMinutes;
    if (mins > 0) {
      this.refreshIntervalId = window.setInterval(
        () => void this.backgroundRefresh(),
        mins * 60 * 1000,
      );
      this.registerInterval(this.refreshIntervalId);
    }

    this.visibilityHandler = () => {
      if (document.visibilityState !== "visible") return;
      if (!this.accounts.anySignedIn()) return;
      const elapsed = Date.now() - this.lastBackgroundRefreshAt;
      if (elapsed < 30_000) return;
      void this.backgroundRefresh();
    };
    document.addEventListener("visibilitychange", this.visibilityHandler);
  }

  private stopAutoRefresh(): void {
    if (this.refreshIntervalId !== null) {
      window.clearInterval(this.refreshIntervalId);
      this.refreshIntervalId = null;
    }
    if (this.visibilityHandler) {
      document.removeEventListener("visibilitychange", this.visibilityHandler);
      this.visibilityHandler = null;
    }
  }
}
