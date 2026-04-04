// Minimal mock of the obsidian module for unit testing
export function requestUrl(_opts: unknown): Promise<unknown> {
  throw new Error("requestUrl is not available in tests");
}
