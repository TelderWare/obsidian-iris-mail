import type { MailFolder, Message } from "../types";

export interface MailListResponse<T> {
  value: T[];
  /** Opaque cursor for the next page, or null when no more pages. */
  nextLink: string | null;
}

export interface ListMessagesOptions {
  top?: number;
  search?: string;
  unreadOnly?: boolean;
  /** Opaque pagination cursor returned from a prior MailListResponse. */
  nextLink?: string;
  /** Only return messages received on or after this instant. Omit = no lower bound. */
  since?: Date;
}

/**
 * Provider-neutral mail API. Implementations adapt a backend (Microsoft Graph,
 * IMAP, ...) to this surface so the rest of the plugin can stay agnostic.
 */
export interface MailApi {
  listFolders(): Promise<MailFolder[]>;
  listMessages(folderId: string, options?: ListMessagesOptions): Promise<MailListResponse<Message>>;
  getMessage(messageId: string): Promise<Message>;
  getMessageBody(messageId: string): Promise<Message>;
  markAsRead(messageId: string): Promise<void>;
  markAsUnread(messageId: string): Promise<void>;
  /** Move a message to the provider's trash/deleted-items folder. */
  deleteMessage(messageId: string): Promise<void>;
  /** Release any stateful resources (connections, sockets). Optional — only
   *  providers with persistent connections (currently IMAP) implement this. */
  dispose?(): Promise<void>;
}
