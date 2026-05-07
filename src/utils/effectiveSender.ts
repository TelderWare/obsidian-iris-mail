import type IrisMailPlugin from "../main";
import type { Message } from "../types";
import type { EffectiveSender, EffectiveSenderResolver } from "../views/components/MessageList";
import { getEnvelopeSender } from "./envelopeSender";
import { extractForwardedSender } from "./extractForwardedSender";

/**
 * Resolve the sender to display for a message. When `resolveForwardedSender`
 * is enabled and the cached body identifies the original sender of a forward,
 * that sender is returned with the envelope sender surfaced as `via*` fields.
 * Otherwise the envelope sender is returned as-is.
 */
export function getEffectiveSender(plugin: IrisMailPlugin, msg: Message): EffectiveSender {
  const envelope = getEnvelopeSender(msg);

  if (!plugin.settings.resolveForwardedSender) {
    return { address: envelope.address, name: envelope.name };
  }

  const cached = plugin.store.getBody(msg.id || "");
  if (cached) {
    let original = cached.originalSender;
    if (!original && /^(?:fw|fwd)\s*:/i.test(cached.subject) && cached.bodyHtml) {
      original = extractForwardedSender(cached.bodyHtml) ?? undefined;
      if (original) cached.originalSender = original;
    }
    if (original?.address) {
      return {
        address: original.address,
        name: original.name || original.address,
        viaAddress: envelope.address,
        viaName: envelope.name,
      };
    }
  }

  return { address: envelope.address, name: envelope.name };
}

/** Build a resolver suitable for MessageList / MessageViewer, or null if the
 *  feature is disabled so callers can skip the extra work. */
export function makeEffectiveSenderResolver(
  plugin: IrisMailPlugin,
): EffectiveSenderResolver | null {
  if (!plugin.settings.resolveForwardedSender) return null;
  return (msg: Message) => getEffectiveSender(plugin, msg);
}
