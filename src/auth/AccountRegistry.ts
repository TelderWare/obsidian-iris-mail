import { OutlookAuthProvider } from "./OutlookAuthProvider";
import { ImapAuthProvider } from "./ImapAuthProvider";
import { OutlookMailApi } from "../mail/OutlookMailApi";
import { ImapMailApi } from "../mail/ImapMailApi";
import type { Account, IrisMailSettings } from "../types";
import type { IAuthProvider } from "./IAuthProvider";
import type { MailApi } from "../mail/MailApi";
import { logger } from "../utils/logger";

export interface AccountEntry {
  account: Account;
  auth: IAuthProvider;
  mail: MailApi;
}

/**
 * Owns the per-account auth + mail backends. Mirrors the accounts array in
 * settings: every settings.accounts entry has exactly one entry here. Keeping
 * lifecycle in one place lets main.ts treat accounts as a unit and lets the
 * MailDispatcher fan messages out without knowing how each account was built.
 */
export class AccountRegistry {
  private entries: Map<string, AccountEntry> = new Map();

  constructor(private dispatcher: { setAccount(a: Account, m: MailApi): void; removeAccount(id: string): void }) {}

  get(accountId: string): AccountEntry | undefined {
    return this.entries.get(accountId);
  }

  anySignedIn(): boolean {
    for (const e of this.entries.values()) {
      if (e.auth.isSignedIn()) return true;
    }
    return false;
  }

  hasCredentials(account: Account): boolean {
    if (account.provider === "imap") {
      return !!(account.imapHost && account.imapPort && account.imapEmail);
    }
    return !!account.clientId;
  }

  /** Build providers for every account in settings and (if configured) initialize. */
  async initializeAll(accounts: Account[], settings: IrisMailSettings): Promise<void> {
    await Promise.all(
      accounts.map((account) =>
        this.add(account, settings).catch((err) => {
          logger.error("AccountRegistry", `Failed to initialize ${account.label}`, err);
        }),
      ),
    );
  }

  /** Build providers for an account and add it to the dispatcher. */
  async add(account: Account, settings: IrisMailSettings): Promise<AccountEntry> {
    if (this.entries.has(account.id)) {
      throw new Error(`Account ${account.id} already registered`);
    }
    const { auth, mail } = this.build(account);
    const entry: AccountEntry = { account, auth, mail };
    this.entries.set(account.id, entry);
    if (account.enabled) {
      this.dispatcher.setAccount(account, mail);
    }
    if (this.hasCredentials(account)) {
      try {
        await auth.initialize(settings);
      } catch (err) {
        logger.warn("AccountRegistry", `initialize failed for ${account.label}`, err);
      }
    }
    return entry;
  }

  /** Mutate an existing entry's account (e.g. after editing label/credentials). */
  updateAccount(updated: Account): void {
    const entry = this.entries.get(updated.id);
    if (!entry) return;
    entry.account = updated;
    // Providers hold their own account reference captured at construction; push
    // the new snapshot down so edits take effect without a plugin reload.
    const auth = entry.auth as { setAccount?: (a: Account) => void };
    auth.setAccount?.(updated);
    const mail = entry.mail as { setAccount?: (a: Account) => void };
    mail.setAccount?.(updated);
    if (updated.enabled) {
      this.dispatcher.setAccount(updated, entry.mail);
    } else {
      this.dispatcher.removeAccount(updated.id);
    }
  }

  /** Sign out (best-effort) and remove the account entirely. */
  async remove(accountId: string): Promise<void> {
    const entry = this.entries.get(accountId);
    if (!entry) return;
    try {
      await entry.auth.logout();
    } catch (err) {
      logger.warn("AccountRegistry", `logout during remove failed for ${entry.account.label}`, err);
    }
    try {
      entry.auth.destroy();
    } catch { /* ignore */ }
    if (entry.mail.dispose) {
      try { await entry.mail.dispose(); } catch { /* ignore */ }
    }
    this.dispatcher.removeAccount(accountId);
    this.entries.delete(accountId);
  }

  destroyAll(): void {
    for (const e of this.entries.values()) {
      try { e.auth.destroy(); } catch { /* ignore */ }
      if (e.mail.dispose) {
        e.mail.dispose().catch(() => { /* ignore */ });
      }
    }
    this.entries.clear();
  }

  private build(account: Account): { auth: IAuthProvider; mail: MailApi } {
    if (account.provider === "imap") {
      const auth = new ImapAuthProvider(account);
      const mail = new ImapMailApi(account, auth);
      return { auth, mail };
    }
    const auth = new OutlookAuthProvider(account);
    const mail = new OutlookMailApi(auth);
    return { auth, mail };
  }
}
