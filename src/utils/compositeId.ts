import * as crypto from "crypto";

/**
 * In multi-account mode, every message id flowing through the plugin is
 * `{accountId}:{nativeId}`. The MailDispatcher attaches the prefix when
 * a message comes back from a provider and strips it again before calling
 * the underlying API. Cache keys (bodies, processed, tags, etc.) inherit
 * the prefix automatically since they're keyed by message.id.
 *
 * accountIds are prefixed `acct_` + hex and never contain ':' — so a single
 * split-on-first-colon is unambiguous regardless of what the backend native id
 * looks like.
 */

const SEP = ":";

export function compositeId(accountId: string, nativeId: string): string {
  return `${accountId}${SEP}${nativeId}`;
}

export function parseCompositeId(id: string): { accountId: string; nativeId: string } | null {
  const idx = id.indexOf(SEP);
  if (idx <= 0) return null;
  return { accountId: id.slice(0, idx), nativeId: id.slice(idx + 1) };
}

export function newAccountId(): string {
  return "acct_" + crypto.randomUUID().replace(/-/g, "");
}
