import type { MailFolder, Message, Account } from "../types";
import type { MailApi, MailListResponse, ListMessagesOptions } from "./MailApi";
import { compositeId, parseCompositeId } from "../utils/compositeId";
import { logger } from "../utils/logger";

const UNIFIED_INBOX_ID = "_unified_inbox";

interface AccountEntry {
  account: Account;
  api: MailApi;
  /** Cached id of the account's "Inbox"-like folder, resolved on first list. */
  inboxFolderId?: string;
}

/**
 * Fans listMessages out across every enabled account, attaches an account
 * prefix to message ids on the way back, and routes per-message methods
 * (markAsRead, getMessageBody, ...) to the correct backend by parsing that
 * prefix. The rest of the plugin sees a single MailApi.
 */
export class MailDispatcher implements MailApi {
  private entries: Map<string, AccountEntry> = new Map();

  setAccount(account: Account, api: MailApi): void {
    this.entries.set(account.id, { account, api });
  }

  removeAccount(accountId: string): void {
    this.entries.delete(accountId);
  }

  async listFolders(): Promise<MailFolder[]> {
    return [{ id: UNIFIED_INBOX_ID, displayName: "Inbox" }];
  }

  async listMessages(
    _folderId: string,
    options: ListMessagesOptions = {},
  ): Promise<MailListResponse<Message>> {
    const cursors: Record<string, string> = options.nextLink
      ? this.decodeCursors(options.nextLink)
      : {};
    const isFollowUp = !!options.nextLink;

    logger.debug("MailDispatcher", `listMessages across ${this.entries.size} account(s)`);

    const tasks = [...this.entries.values()].map(async (entry) => {
      if (isFollowUp && !cursors[entry.account.id]) return null;
      const inboxId = await this.resolveInboxId(entry);
      if (!inboxId) {
        logger.debug("MailDispatcher", `no inbox folder resolved for ${entry.account.label}`);
        return null;
      }
      try {
        const resp = await entry.api.listMessages(inboxId, {
          top: options.top,
          search: options.search,
          unreadOnly: options.unreadOnly,
          nextLink: cursors[entry.account.id],
        });
        logger.debug("MailDispatcher", `${entry.account.label}: ${resp.value.length} messages, nextLink=${resp.nextLink ?? "-"}`);
        for (const m of resp.value) this.stamp(m, entry);
        return { accountId: entry.account.id, value: resp.value, nextLink: resp.nextLink };
      } catch (err) {
        logger.warn("MailDispatcher", `listMessages failed for ${entry.account.label}`, err);
        return null;
      }
    });

    const results = (await Promise.all(tasks)).filter(
      (r): r is { accountId: string; value: Message[]; nextLink: string | null } => r !== null,
    );

    const merged = results.flatMap((r) => r.value);
    merged.sort(byReceivedDesc);

    const nextCursors: Record<string, string> = {};
    for (const r of results) {
      if (r.nextLink) nextCursors[r.accountId] = r.nextLink;
    }

    return {
      value: merged,
      nextLink: Object.keys(nextCursors).length > 0 ? JSON.stringify(nextCursors) : null,
    };
  }

  async getMessage(messageId: string): Promise<Message> {
    return await this.route(messageId, async (api, native, entry) => {
      const m = await api.getMessage(native);
      this.stamp(m, entry);
      return m;
    });
  }

  async getMessageBody(messageId: string): Promise<Message> {
    return await this.route(messageId, async (api, native, entry) => {
      const m = await api.getMessageBody(native);
      this.stamp(m, entry);
      return m;
    });
  }

  async markAsRead(messageId: string): Promise<void> {
    await this.route(messageId, (api, native) => api.markAsRead(native));
  }

  async markAsUnread(messageId: string): Promise<void> {
    await this.route(messageId, (api, native) => api.markAsUnread(native));
  }

  async deleteMessage(messageId: string): Promise<void> {
    await this.route(messageId, (api, native) => api.deleteMessage(native));
  }

  // --- private ---

  private async route<T>(
    messageId: string,
    fn: (api: MailApi, nativeId: string, entry: AccountEntry) => Promise<T>,
  ): Promise<T> {
    const parsed = parseCompositeId(messageId);
    if (!parsed) {
      throw new Error(`MailDispatcher: id has no account prefix: ${messageId}`);
    }
    const entry = this.entries.get(parsed.accountId);
    if (!entry) {
      throw new Error(`MailDispatcher: no account for id ${parsed.accountId}`);
    }
    return await fn(entry.api, parsed.nativeId, entry);
  }

  private stamp(m: Message, entry: AccountEntry): void {
    if (m.id) m.id = compositeId(entry.account.id, m.id);
    m._accountId = entry.account.id;
    if (this.entries.size > 1) m._accountLabel = entry.account.label;
  }

  private async resolveInboxId(entry: AccountEntry): Promise<string | undefined> {
    if (entry.inboxFolderId) return entry.inboxFolderId;
    try {
      const folders = await entry.api.listFolders();
      const inbox = folders.find((f) => f.displayName?.toLowerCase() === "inbox");
      if (inbox?.id) {
        entry.inboxFolderId = inbox.id;
        return entry.inboxFolderId;
      }
    } catch (err) {
      logger.warn("MailDispatcher", `listFolders failed for ${entry.account.label}`, err);
    }
    return undefined;
  }

  private decodeCursors(raw: string): Record<string, string> {
    try {
      const parsed = JSON.parse(raw);
      if (parsed && typeof parsed === "object") return parsed as Record<string, string>;
    } catch { /* ignore */ }
    return {};
  }
}

function byReceivedDesc(a: Message, b: Message): number {
  const ad = a.receivedDateTime ?? "";
  const bd = b.receivedDateTime ?? "";
  return bd.localeCompare(ad);
}
