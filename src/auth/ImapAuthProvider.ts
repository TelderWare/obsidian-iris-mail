import { ImapFlow } from "imapflow";
import { IMAP_PASSWORD_STORAGE_KEY } from "../constants";
import { logger } from "../utils/logger";
import { encryptString, decryptString } from "../utils/safeStorage";
import { formatImapError } from "../mail/imapError";
import type { IrisMailSettings, Account } from "../types";
import type { IAuthProvider, AuthAccount } from "./IAuthProvider";

interface StoredCreds {
  password: string;
  /** True once the credentials successfully authenticated against the server. */
  validated: boolean;
}

function storageKeyFor(accountId: string): string {
  return `${IMAP_PASSWORD_STORAGE_KEY}:${accountId}`;
}

function loadCreds(accountId: string): StoredCreds | null {
  const raw = localStorage.getItem(storageKeyFor(accountId));
  if (!raw) return null;
  try {
    const parsed = JSON.parse(decryptString(raw));
    if (typeof parsed?.password === "string") {
      // Trim on load too — passwords stored before setPassword() started
      // cleaning input would otherwise carry whitespace forever.
      return { password: parsed.password.trim(), validated: !!parsed.validated };
    }
    return null;
  } catch {
    localStorage.removeItem(storageKeyFor(accountId));
    return null;
  }
}

function saveCreds(accountId: string, creds: StoredCreds): void {
  localStorage.setItem(storageKeyFor(accountId), encryptString(JSON.stringify(creds)));
}

function clearCreds(accountId: string): void {
  localStorage.removeItem(storageKeyFor(accountId));
}

/**
 * IMAP auth is username + password (typically an app password). Unlike OAuth
 * providers there are no tokens to refresh — `getAccessToken` returns the
 * stored password for the IMAP client to use in AUTH LOGIN/PLAIN.
 *
 * Two-stage signed-in model: the password is saved (encrypted) as soon as the
 * user types it, so it survives settings-panel close and restarts. The account
 * only counts as "signed in" (isSignedIn) after `login` validates it against
 * the server at least once. Changing the password invalidates that flag.
 */
export class ImapAuthProvider implements IAuthProvider {
  private creds: StoredCreds | null = null;

  constructor(private account: Account) {}

  /** Swap in a freshly-edited account snapshot. Called by AccountRegistry when
   *  settings change so the provider sees the latest host/port/email without
   *  needing to rebuild the instance. */
  setAccount(account: Account): void {
    this.account = account;
  }

  async initialize(_settings: IrisMailSettings): Promise<void> {
    this.creds = loadCreds(this.account.id);
  }

  async login(_settings: IrisMailSettings): Promise<void> {
    if (!this.creds?.password) {
      throw new Error("Enter your app password first.");
    }
    const { imapHost, imapPort, imapSecure, imapEmail } = this.account;
    if (!imapHost || !imapPort || !imapEmail) {
      throw new Error("IMAP host, port, and email are required.");
    }
    const client = new ImapFlow({
      host: imapHost,
      port: imapPort,
      secure: imapSecure ?? true,
      auth: { user: imapEmail, pass: this.creds.password },
      logger: false,
    });
    try {
      await client.connect();
    } catch (err) {
      logger.error("ImapAuth", "IMAP connect failed", err);
      try { client.close(); } catch { /* ignore */ }
      throw new Error(formatImapError(err));
    }
    try { await client.logout(); } catch { /* ignore */ }

    this.creds = { password: this.creds.password, validated: true };
    saveCreds(this.account.id, this.creds);
  }

  async getAccessToken(): Promise<string> {
    if (!this.creds?.password) throw new Error("Not signed in");
    return this.creds.password;
  }

  async logout(): Promise<void> {
    this.creds = null;
    clearCreds(this.account.id);
  }

  isSignedIn(): boolean {
    return this.creds?.validated === true;
  }

  getAccount(): AuthAccount | null {
    if (!this.creds?.validated || !this.account.imapEmail) return null;
    return { username: this.account.imapEmail };
  }

  destroy(): void {
    /* no resources to release */
  }

  // --- IMAP-specific helpers used by SettingsTab ---

  /** Persist the password immediately (encrypted). Invalidates any prior
   *  validation — user must click Sign in again to confirm it works. An empty
   *  string clears the stored credentials entirely.
   *
   *  Trims whitespace — pasting app passwords from web pages often pulls in a
   *  trailing newline / NBSP that the server rejects as an invalid credential. */
  setPassword(password: string): void {
    const cleaned = password.trim();
    if (!cleaned) {
      this.creds = null;
      clearCreds(this.account.id);
      logger.debug("ImapAuth", `cleared stored password for ${this.account.id}`);
      return;
    }
    this.creds = { password: cleaned, validated: false };
    saveCreds(this.account.id, this.creds);
    logger.debug("ImapAuth", `saved password for ${this.account.id} (${cleaned.length} chars)`);
  }

  /** True when a password is stored (regardless of whether it's been validated
   *  against the server yet). The settings UI uses this to decide between
   *  "empty input" and "sentinel dots" rendering. */
  hasStoredPassword(): boolean {
    return !!this.creds?.password;
  }
}
