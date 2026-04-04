import type { Message } from "../types";

export interface EnvelopeSender {
  address: string;
  name: string;
}

/**
 * Extract the envelope sender from a message.
 * Prefers `msg.from` (logical author) over `msg.sender` (technical/delegate mailbox).
 */
export function getEnvelopeSender(msg: Message): EnvelopeSender {
  const address =
    msg.from?.emailAddress?.address ||
    msg.sender?.emailAddress?.address ||
    "";
  const name =
    msg.from?.emailAddress?.name ||
    msg.sender?.emailAddress?.name ||
    address ||
    "Unknown";
  return { address, name };
}
