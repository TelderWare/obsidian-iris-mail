/**
 * imapflow's `err.message` is almost always a generic "Command failed" /
 * "Socket timeout" / "NO response". The server's actual reply (the only
 * diagnostic info that's usually useful) is tucked onto side properties.
 * Surface whichever one is populated so errors are actionable.
 */
export function formatImapError(err: unknown): string {
  if (!(err instanceof Error)) return String(err);
  const extra = err as Error & {
    responseText?: string;
    response?: string;
    authenticationFailed?: boolean;
    code?: string;
  };
  const detail = extra.responseText || extra.response;
  if (detail) return `${err.message}: ${detail}`;
  if (extra.code) return `${err.message} (${extra.code})`;
  return err.message;
}
