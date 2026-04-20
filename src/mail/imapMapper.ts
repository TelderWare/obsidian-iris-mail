import type { FetchMessageObject, MessageAddressObject } from "imapflow";
import type { Message, Recipient } from "../types";

/**
 * Encoded messageId used by ImapMailApi: `${folderPath}\u001E${uid}`.
 *
 * Folders are paths like "INBOX", "[Gmail]/All Mail", "Archive/2024" — they
 * can contain `/`, `.`, spaces etc., but never control chars, so `\u001E`
 * (Record Separator) is a safe delimiter that won't collide with real folder
 * names.
 */
const SEP = "\u001E";

export function encodeImapId(folderPath: string, uid: number): string {
  return `${folderPath}${SEP}${uid}`;
}

export function decodeImapId(id: string): { folder: string; uid: number } {
  const idx = id.indexOf(SEP);
  if (idx < 0) throw new Error(`Invalid IMAP message id: ${id}`);
  const folder = id.slice(0, idx);
  const uid = parseInt(id.slice(idx + 1), 10);
  if (!Number.isFinite(uid)) throw new Error(`Invalid IMAP UID in id: ${id}`);
  return { folder, uid };
}

function toRecipient(a: MessageAddressObject): Recipient {
  return {
    emailAddress: {
      name: a.name || undefined,
      address: a.address || "",
    },
  };
}

function toRecipients(list: MessageAddressObject[] | undefined): Recipient[] {
  if (!list) return [];
  return list.filter((a) => a.address).map(toRecipient);
}

/** Build a Message from the envelope-only fetch result returned by listMessages. */
export function imapEnvelopeToMessage(
  folderPath: string,
  fetched: FetchMessageObject,
): Message {
  const env = fetched.envelope;
  const flags = fetched.flags ?? new Set<string>();
  const from = env?.from?.[0] ? toRecipient(env.from[0]) : undefined;
  const date =
    env?.date ??
    (fetched.internalDate instanceof Date ? fetched.internalDate : undefined) ??
    (typeof fetched.internalDate === "string" ? new Date(fetched.internalDate) : undefined);

  return {
    id: encodeImapId(folderPath, fetched.uid),
    subject: env?.subject ?? "",
    bodyPreview: "",
    from,
    sender: from,
    toRecipients: toRecipients(env?.to),
    ccRecipients: toRecipients(env?.cc),
    receivedDateTime: (date ?? new Date()).toISOString(),
    isRead: flags.has("\\Seen"),
    isDraft: flags.has("\\Draft"),
    hasAttachments: false,
    importance: "normal",
  };
}
