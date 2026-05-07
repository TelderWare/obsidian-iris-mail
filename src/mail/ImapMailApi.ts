import { ImapFlow, type SearchObject } from "imapflow";
import { simpleParser } from "mailparser";
import type { ImapAuthProvider } from "../auth/ImapAuthProvider";
import type { Account, MailFolder, Message } from "../types";
import { logger } from "../utils/logger";
import type { MailApi, MailListResponse, ListMessagesOptions } from "./MailApi";
import { formatImapError } from "./imapError";
import { decodeImapId, encodeImapId, imapEnvelopeToMessage } from "./imapMapper";

/**
 * Generic IMAP implementation of MailApi. Works with Gmail, iCloud, Fastmail,
 * self-hosted Dovecot/Cyrus — anything that speaks IMAP4rev1 over TLS.
 *
 * Connection model: one persistent ImapFlow client per account, created
 * lazily. `getMailboxLock` serializes per-folder ops so concurrent calls to
 * listMessages / getMessageBody don't race on SELECT. On network errors or
 * auth failures the connection is dropped and recreated on the next call.
 */
export class ImapMailApi implements MailApi {
  private client: ImapFlow | null = null;
  private connecting: Promise<ImapFlow> | null = null;
  private trashFolderPath: string | null = null;

  constructor(
    private account: Account,
    private readonly auth: ImapAuthProvider,
  ) {}

  /** Swap in a freshly-edited account snapshot. If connection-affecting fields
   *  (host/port/email/tls) changed, drop the existing connection so the next
   *  call reconnects with the new settings. */
  setAccount(account: Account): void {
    const changed =
      this.account.imapHost !== account.imapHost ||
      this.account.imapPort !== account.imapPort ||
      this.account.imapSecure !== account.imapSecure ||
      this.account.imapEmail !== account.imapEmail;
    this.account = account;
    if (changed) this.closeSocket();
  }

  async listFolders(): Promise<MailFolder[]> {
    const client = await this.connection();
    const folders = await client.list();

    const mapped = folders
      .filter((f) => !f.flags.has("\\Noselect"))
      .map((f) => ({
        id: f.path,
        displayName: f.name,
      }));
    logger.debug("ImapMailApi", `listFolders -> ${mapped.length} folders`, mapped.map((m) => m.displayName));
    return mapped;
  }

  async listMessages(
    folderId: string,
    options: ListMessagesOptions = {},
  ): Promise<MailListResponse<Message>> {
    const top = options.top ?? 25;
    const folder = folderId || "INBOX";
    const offset = parseCursor(options.nextLink);

    logger.debug("ImapMailApi", `listMessages folder=${folder} top=${top} offset=${offset} unreadOnly=${!!options.unreadOnly}`);

    const client = await this.connection();
    const lock = await client.getMailboxLock(folder);
    try {
      const search: SearchObject = options.unreadOnly
        ? { seen: false }
        : { all: true };
      if (options.search) search.text = options.search;
      if (options.since) search.since = options.since;

      const uids = (await client.search(search, { uid: true })) || [];
      const sorted = Array.isArray(uids) ? [...uids].sort((a, b) => b - a) : [];
      logger.debug("ImapMailApi", `search returned ${sorted.length} uids`);

      const slice = sorted.slice(offset, offset + top);
      if (slice.length === 0) {
        return { value: [], nextLink: null };
      }

      const messages: Message[] = [];
      for await (const msg of client.fetch(
        slice,
        { uid: true, envelope: true, flags: true, internalDate: true },
        { uid: true },
      )) {
        messages.push(imapEnvelopeToMessage(folder, msg));
      }
      logger.debug("ImapMailApi", `fetched ${messages.length} envelopes`);

      messages.sort(
        (a, b) =>
          new Date(b.receivedDateTime ?? 0).getTime() -
          new Date(a.receivedDateTime ?? 0).getTime(),
      );

      const nextOffset = offset + slice.length;
      const nextLink = nextOffset < sorted.length ? String(nextOffset) : null;
      return { value: messages, nextLink };
    } finally {
      lock.release();
    }
  }

  async getMessage(messageId: string): Promise<Message> {
    const { folder, uid } = decodeImapId(messageId);
    const client = await this.connection();
    const lock = await client.getMailboxLock(folder);
    try {
      const msg = await client.fetchOne(String(uid), {
        uid: true,
        envelope: true,
        flags: true,
        internalDate: true,
      }, { uid: true });
      if (!msg) throw new Error(`Message not found: ${messageId}`);
      return imapEnvelopeToMessage(folder, msg);
    } finally {
      lock.release();
    }
  }

  async getMessageBody(messageId: string): Promise<Message> {
    const { folder, uid } = decodeImapId(messageId);
    const client = await this.connection();
    const lock = await client.getMailboxLock(folder);
    try {
      const msg = await client.fetchOne(String(uid), {
        uid: true,
        envelope: true,
        flags: true,
        internalDate: true,
        source: true,
      }, { uid: true });
      if (!msg) throw new Error(`Message not found: ${messageId}`);

      const base = imapEnvelopeToMessage(folder, msg);
      if (!msg.source) return base;

      const parsed = await simpleParser(msg.source);
      const html = typeof parsed.html === "string" ? parsed.html : null;
      const text = parsed.text ?? "";
      const body = html
        ? { contentType: "html" as const, content: html }
        : { contentType: "text" as const, content: text };

      return {
        ...base,
        id: encodeImapId(folder, uid),
        bodyPreview: (parsed.text ?? "").slice(0, 250).replace(/\s+/g, " ").trim(),
        hasAttachments: (parsed.attachments?.length ?? 0) > 0,
        body,
      };
    } finally {
      lock.release();
    }
  }

  async markAsRead(messageId: string): Promise<void> {
    await this.setSeenFlag(messageId, true);
  }

  async markAsUnread(messageId: string): Promise<void> {
    await this.setSeenFlag(messageId, false);
  }

  async deleteMessage(messageId: string): Promise<void> {
    const { folder, uid } = decodeImapId(messageId);
    const client = await this.connection();
    const trash = await this.resolveTrashFolder(client);

    const lock = await client.getMailboxLock(folder);
    try {
      if (trash && trash !== folder) {
        await client.messageMove(String(uid), trash, { uid: true });
      } else {
        // No trash folder (or already in it) — flag + expunge.
        await client.messageFlagsAdd(String(uid), ["\\Deleted"], { uid: true });
        await client.messageDelete(String(uid), { uid: true });
      }
    } finally {
      lock.release();
    }
  }

  /** Close the IMAP connection. Called by AccountRegistry on account removal. */
  async dispose(): Promise<void> {
    this.closeSocket();
  }

  // --- private ---

  /** Close the TCP connection without a graceful LOGOUT round-trip. LOGOUT
   *  can hang if the connection is already half-dead, and for dispose/reset
   *  we don't care about server-side cleanup — the server times idle sessions
   *  out on its own. */
  private closeSocket(): void {
    if (!this.client) return;
    try { this.client.close(); } catch { /* ignore */ }
    this.client = null;
    this.trashFolderPath = null;
  }

  /** Resolve the server's trash/deleted folder path. Prefers the RFC 6154
   *  \Trash special-use flag; falls back to common names. Cached per connection. */
  private async resolveTrashFolder(client: ImapFlow): Promise<string | null> {
    if (this.trashFolderPath) return this.trashFolderPath;

    const folders = await client.list();
    const bySpecialUse = folders.find((f) => f.specialUse === "\\Trash");
    if (bySpecialUse) {
      this.trashFolderPath = bySpecialUse.path;
      return this.trashFolderPath;
    }

    const candidates = ["Trash", "Deleted", "Deleted Messages", "Deleted Items", "[Gmail]/Trash"];
    const lower = new Map(folders.map((f) => [f.path.toLowerCase(), f.path]));
    for (const c of candidates) {
      const path = lower.get(c.toLowerCase());
      if (path) {
        this.trashFolderPath = path;
        return this.trashFolderPath;
      }
    }
    return null;
  }

  private async setSeenFlag(messageId: string, seen: boolean): Promise<void> {
    const { folder, uid } = decodeImapId(messageId);
    const client = await this.connection();
    const lock = await client.getMailboxLock(folder);
    try {
      if (seen) {
        await client.messageFlagsAdd(String(uid), ["\\Seen"], { uid: true });
      } else {
        await client.messageFlagsRemove(String(uid), ["\\Seen"], { uid: true });
      }
    } finally {
      lock.release();
    }
  }

  private async connection(): Promise<ImapFlow> {
    if (this.client?.usable) return this.client;
    if (this.connecting) return await this.connecting;

    this.connecting = this.connect().finally(() => {
      this.connecting = null;
    });
    return await this.connecting;
  }

  private async connect(): Promise<ImapFlow> {
    const { imapHost, imapPort, imapSecure, imapEmail } = this.account;
    if (!imapHost || !imapPort || !imapEmail) {
      throw new Error("IMAP host, port, and email are required.");
    }
    const pass = await this.auth.getAccessToken();

    const client = new ImapFlow({
      host: imapHost,
      port: imapPort,
      secure: imapSecure ?? true,
      auth: { user: imapEmail, pass },
      logger: false,
    });
    client.on("error", (err: Error) => {
      logger.warn("ImapMailApi", `connection error on ${imapHost}`, err);
      if (this.client === client) {
        this.client = null;
        try { client.close(); } catch { /* ignore */ }
      }
    });
    client.on("close", () => {
      if (this.client === client) this.client = null;
    });
    try {
      await client.connect();
    } catch (err) {
      try { client.close(); } catch { /* ignore */ }
      throw new Error(formatImapError(err));
    }
    this.client = client;
    logger.debug("ImapMailApi", `connected to ${imapHost}:${imapPort} as ${imapEmail}`);
    return client;
  }
}

function parseCursor(cursor: string | undefined): number {
  if (!cursor) return 0;
  const n = parseInt(cursor, 10);
  return Number.isFinite(n) && n >= 0 ? n : 0;
}
