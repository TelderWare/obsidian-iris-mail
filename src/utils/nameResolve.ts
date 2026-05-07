/**
 * Normalize "LastName, FirstName" to "FirstName LastName" and strip
 * Outlook delegate suffixes like "Name (via Institution)". Leaves other
 * formats untouched. Shared between InboxView and homepage widgets so
 * the sender name renders identically everywhere.
 */
export function normalizeName(raw: string): string {
  let name = raw.replace(/\s*\(via\s+[^)]+\)\s*$/i, "").trim();
  const m = name.match(/^([^,]+),\s*(.+)$/);
  if (m) name = `${m[2]} ${m[1]}`;
  return name;
}

/** Nickname-aware name resolver used by MessageList / MessageViewer. */
export function makeNameResolver(
  nicknames: Map<string, string>,
): (address: string, rawName: string) => string {
  return (address, rawName) => {
    if (!address) return normalizeName(rawName);
    return nicknames.get(address.toLowerCase()) || normalizeName(rawName);
  };
}
