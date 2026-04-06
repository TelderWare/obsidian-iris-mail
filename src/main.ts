import { Plugin, WorkspaceLeaf, Notice } from "obsidian";
import { InboxView } from "./views/InboxView";
import { IrisMailSettingsTab } from "./settings/SettingsTab";
import { AuthProvider } from "./auth/AuthProvider";
import { GraphMailApi } from "./graph/GraphMailApi";
import { EmailStore } from "./store/EmailStore";
import {
  VIEW_TYPE_IRIS_MAIL,
  ICON_NAME,
  DEFAULT_SETTINGS,
} from "./constants";
import { logger, setDebugEnabled } from "./utils/logger";
import { setRelayApp } from "./utils/claudeApi";
import type { AuthMethod, IrisMailSettings, AuthState, Message } from "./types";

/** Encrypt a string using Electron's safeStorage if available, else return as-is. */
function encryptApiKey(key: string): string {
  if (!key) return "";
  try {
    const { safeStorage } = require("electron");
    if (safeStorage.isEncryptionAvailable()) {
      return "enc:" + safeStorage.encryptString(key).toString("base64");
    }
  } catch { /* safeStorage unavailable */ }
  return key;
}

/** Decrypt a string using Electron's safeStorage if it was encrypted. */
function decryptApiKey(stored: string): string {
  if (!stored) return "";
  if (stored.startsWith("enc:")) {
    try {
      const { safeStorage } = require("electron");
      return safeStorage.decryptString(Buffer.from(stored.slice(4), "base64"));
    } catch {
      logger.warn("Main", "Failed to decrypt API key — safeStorage may be unavailable");
      return "";
    }
  }
  return stored;
}

export default class IrisMailPlugin extends Plugin {
  settings: IrisMailSettings = DEFAULT_SETTINGS;
  authProvider!: AuthProvider;
  mailApi!: GraphMailApi;
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

    this.store = new EmailStore(this.app);
    await this.store.load();

    this.authProvider = new AuthProvider();
    this.mailApi = new GraphMailApi(this.authProvider);

    if (this.settings.clientId) {
      try {
        await this.authProvider.initialize(this.settings);
      } catch (e) {
        logger.error("Auth", "Failed to initialize auth", e);
      }
    }

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
      id: "iris-login",
      name: "Sign in to Outlook",
      callback: () => this.handleLogin(),
    });

    this.addCommand({
      id: "iris-logout",
      name: "Sign out of Outlook",
      callback: () => this.handleLogout(),
    });

    this.addCommand({
      id: "iris-refresh",
      name: "Refresh Iris Mail",
      callback: () => this.refreshAllViews(),
    });

    this.addSettingTab(new IrisMailSettingsTab(this.app, this));

    this.startAutoRefresh();

    // Initial background badge update (doesn't require the view to be open)
    if (this.authProvider.isSignedIn()) {
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
    this.authProvider?.destroy();
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

  private async ensureInitialized(): Promise<boolean> {
    if (!this.settings.clientId) {
      new Notice("Please configure your Azure Client ID in Iris Mail settings first.");
      return false;
    }
    await this.authProvider.initialize(this.settings);
    return true;
  }

  async handleLogin(method?: AuthMethod): Promise<void> {
    const useMethod = method || this.settings.authMethod;
    try {
      if (!(await this.ensureInitialized())) return;
      if (useMethod === "device-code") {
        await this.doDeviceCodeLogin();
      } else {
        await this.authProvider.login(this.settings);
      }
      new Notice("Signed in to Outlook successfully.");
      this.refreshAllViews();
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      new Notice(`Outlook sign-in failed: ${msg}`);
    }
  }

  async handleLoginWithDeviceCode(): Promise<void> {
    return this.handleLogin("device-code");
  }

  async handleLoginWithAuthCode(): Promise<void> {
    return this.handleLogin("auth-code");
  }

  private async doDeviceCodeLogin(): Promise<void> {
    let codeNotice: Notice | undefined;
    let copyNotice: Notice | undefined;
    try {
      await this.authProvider.loginWithDeviceCode(
        this.settings,
        (code, verificationUri) => {
          const frag = document.createDocumentFragment();
          frag.appendText("Go to ");
          const link = frag.createEl("a", {
            text: verificationUri,
            href: verificationUri,
          });
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

  async handleLogout(): Promise<void> {
    await this.authProvider.logout();
    new Notice("Signed out of Outlook.");
    this.refreshAllViews();
  }

  async loadSettings(): Promise<void> {
    const data = await this.loadData();
    this.settings = Object.assign({}, DEFAULT_SETTINGS, data);

    // Decrypt API key from storage
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
      await this.saveSettings();
    }

    // Enable debug logging if configured
    setDebugEnabled(!!this.settings.debugLogging);
  }

  async saveSettings(): Promise<void> {
    // Encrypt API key before writing to disk
    const toSave = { ...this.settings };
    if (toSave.anthropicApiKey && !toSave.anthropicApiKey.startsWith("enc:")) {
      toSave.anthropicApiKey = encryptApiKey(toSave.anthropicApiKey);
    }
    await this.saveData(toSave);
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

    if (!this.authProvider.isSignedIn()) return;

    try {
      const folders = await this.mailApi.listFolders();
      const inbox = folders.find(
        (f) => f.displayName?.toLowerCase() === "inbox",
      );
      if (!inbox?.id) return;

      const filter =
        !this.settings.showReadEmails ? "isRead eq false" : undefined;
      const response = await this.mailApi.listMessages(inbox.id, {
        top: this.settings.pageSize,
        filter,
      });

      const messages = response.value;
      this.store.applyReadState(messages);
      this.updateBadgeFromMessages(messages);
    } catch (err) {
      logger.error("Main", "Background refresh failed", err);
    }
  }

  /**
   * Compute badge count from a set of messages using store classification data.
   * Applies the same default filters as InboxView (unread-only + hide-noise).
   */
  private updateBadgeFromMessages(messages: Message[]): void {
    const mode = this.settings.badgeCount;
    if (mode === "off") {
      this.updateBadge(0);
      return;
    }

    const classData = this.store.getAllClassificationData();
    const nonNoise = messages.filter(
      (m) => classData.classes.get(m.id || "") !== "noise",
    );

    switch (mode) {
      case "unread":
        this.updateBadge(nonNoise.filter((m) => !m.isRead).length);
        break;
      case "important":
        this.updateBadge(
          nonNoise.filter(
            (m) => !m.isRead && classData.classes.get(m.id || "") === "important",
          ).length,
        );
        break;
      case "total":
        this.updateBadge(nonNoise.length);
        break;
    }
  }

  updateBadge(count: number): void {
    if (!this.ribbonIconEl) return;

    // -1 signals a settings change — ask the active view to re-sync
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

    // Refresh badge when Obsidian regains focus (throttled to 30 s)
    this.visibilityHandler = () => {
      if (document.visibilityState !== "visible") return;
      if (!this.authProvider.isSignedIn()) return;
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
