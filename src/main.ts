import { Plugin, WorkspaceLeaf, Notice } from "obsidian";
import { InboxView } from "./views/InboxView";
import { IrisMailSettingsTab } from "./settings/SettingsTab";
import { AccountRegistry } from "./auth/AccountRegistry";
import { MailDispatcher } from "./mail/MailDispatcher";
import { EmailStore } from "./store/EmailStore";
import {
  VIEW_TYPE_IRIS_MAIL,
  ICON_NAME,
  DEFAULT_SETTINGS,
  CACHE_STORAGE_KEY,
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

  async onload(): Promise<void> {
    await this.loadSettings();
    setRelayApp(this.app);

    this.store = new EmailStore(this);
    await this.store.load();

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
    const leaf = this.app.workspace.getLeaf(true);
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

    // Drop legacy triage fields if present
    const legacy = this.settings as unknown as Record<string, unknown>;
    delete legacy.triageTree;
    delete legacy.triageGraph;
    delete legacy.triageTreeVersion;
    delete legacy.enableTriage;

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
  async backgroundRefresh(): Promise<void> {
    this.lastBackgroundRefreshAt = Date.now();
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
      });

      const messages = response.value;
      this.store.applyReadState(messages);
      this.updateBadgeFromMessages(messages);
    } catch (err) {
      logger.error("Main", "Background refresh failed", err);
    }
  }

  /** Compute badge count from a set of messages. */
  private updateBadgeFromMessages(messages: Message[]): void {
    const mode = this.settings.badgeCount;
    if (mode === "off") {
      this.updateBadge(0);
      return;
    }

    switch (mode) {
      case "unread":
        this.updateBadge(messages.filter((m) => !m.isRead).length);
        break;
      case "total":
        this.updateBadge(messages.length);
        break;
    }
  }

  updateBadge(count: number): void {
    if (!this.ribbonIconEl) return;

    if (count < 0) {
      for (const leaf of this.app.workspace.getLeavesOfType(VIEW_TYPE_IRIS_MAIL)) {
        (leaf.view as InboxView).syncBadge();
      }
      return;
    }

    const pos = this.settings.badgePosition;
    if (pos !== "off" && count > 0) {
      if (!this.ribbonBadgeEl) {
        this.ribbonBadgeEl = this.ribbonIconEl.createSpan({ cls: "iris-ribbon-badge" });
      }
      this.ribbonBadgeEl.className = `iris-ribbon-badge iris-ribbon-badge-${pos}`;
      this.ribbonBadgeEl.setText(count > 99 ? "99+" : String(count));
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
