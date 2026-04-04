/**
 * Deterministically strips quoted reply content from an email HTML body.
 * Returns only the "new" content from the latest reply.
 */
export function stripQuotedContent(html: string): string {
  // Common reply/forward markers across email clients
  const markers: RegExp[] = [
    // Outlook
    /<div\b[^>]*id\s*=\s*["']divRplyFwdMsg["']/i,
    /<div\b[^>]*id\s*=\s*["']appendonsend["']/i,
    /<hr\b[^>]*style\s*=\s*["'][^"']*display:\s*inline-block[^"']*width:\s*98%/i,
    /<div\b[^>]*style\s*=\s*["'][^"']*border-top:\s*solid\s*#[eE]1[eE]1[eE]1/i,
    // Gmail
    /<div\b[^>]*class\s*=\s*["']gmail_quote["']/i,
    // Generic blockquote with cite
    /<blockquote\b[^>]*type\s*=\s*["']cite["']/i,
    // "Original Message" HTML comment (some clients)
    /<!--\s*Original\s*Message\s*-->/i,
    // "On ... wrote:" pattern (common across clients)
    /<div\b[^>]*>On\s.{10,80}\swrote:\s*<\/div>/i,
    /<p\b[^>]*>On\s.{10,80}\swrote:\s*<\/p>/i,
  ];

  let earliestIndex = html.length;

  for (const marker of markers) {
    const match = marker.exec(html);
    if (match && match.index < earliestIndex) {
      earliestIndex = match.index;
    }
  }

  if (earliestIndex < html.length) {
    const stripped = html.substring(0, earliestIndex).trim();
    // If stripping leaves nothing meaningful (e.g. a pure forward),
    // fall back to the full content.
    const textOnly = stripped
      .replace(/<[^>]*>/g, "")
      .replace(/&\w+;/g, "")
      .replace(/&#\d+;/g, "")
      .replace(/\s+/g, " ")
      .trim();
    if (textOnly.length > 5) {
      return stripped;
    }
  }

  return html;
}
