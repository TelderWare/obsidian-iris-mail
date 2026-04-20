import { logger } from "./logger";

/**
 * Thin wrapper around Electron's safeStorage. Falls back to base64 when the
 * OS keychain is unavailable (e.g. headless Linux without libsecret), warning
 * once so the user knows the on-disk data is not encrypted.
 *
 * Three callers use this: the Anthropic API key (in main.ts), the MSAL token
 * cache, and the IMAP app-password cache. Keeping the fallback semantics in
 * one place avoids the three of them drifting out of sync.
 */

let warned = false;

function get(): { isEncryptionAvailable(): boolean; encryptString(s: string): Buffer; decryptString(b: Buffer): string } | null {
  try {
    const { safeStorage } = require("electron");
    return safeStorage ?? null;
  } catch {
    return null;
  }
}

function warnFallback(): void {
  if (warned) return;
  warned = true;
  logger.warn(
    "safeStorage",
    "Electron safeStorage unavailable — secrets stored as base64 (NOT encrypted). " +
      "Insecure if your vault is synced to cloud or shared.",
  );
}

/** Encrypt and base64-encode. Returns plain base64 of UTF-8 if no keychain. */
export function encryptString(plaintext: string): string {
  const ss = get();
  if (ss?.isEncryptionAvailable()) {
    return ss.encryptString(plaintext).toString("base64");
  }
  warnFallback();
  return Buffer.from(plaintext, "utf-8").toString("base64");
}

/** Decrypt a value produced by encryptString. Throws on malformed input. */
export function decryptString(encoded: string): string {
  const ss = get();
  if (ss?.isEncryptionAvailable()) {
    return ss.decryptString(Buffer.from(encoded, "base64"));
  }
  warnFallback();
  return Buffer.from(encoded, "base64").toString("utf-8");
}
