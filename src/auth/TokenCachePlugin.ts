import type { ICachePlugin, TokenCacheContext } from "@azure/msal-node";
import { CACHE_STORAGE_KEY } from "../constants";
import { logger } from "../utils/logger";

export class TokenCachePlugin implements ICachePlugin {
  private insecureFallbackWarned = false;

  private canUseSafeStorage(): boolean {
    try {
      const { safeStorage } = require("electron");
      return safeStorage.isEncryptionAvailable();
    } catch {
      return false;
    }
  }

  private warnInsecureFallback(): void {
    if (!this.insecureFallbackWarned) {
      this.insecureFallbackWarned = true;
      logger.warn("TokenCache",
        "Electron safeStorage unavailable — tokens are stored as base64 (NOT encrypted). " +
        "This is insecure if your vault is synced to cloud or shared.");
    }
  }

  async beforeCacheAccess(context: TokenCacheContext): Promise<void> {
    const raw = localStorage.getItem(CACHE_STORAGE_KEY);
    if (!raw) return;

    try {
      if (this.canUseSafeStorage()) {
        const { safeStorage } = require("electron");
        const buffer = Buffer.from(raw, "base64");
        const decrypted = safeStorage.decryptString(buffer);
        context.tokenCache.deserialize(decrypted);
      } else {
        this.warnInsecureFallback();
        context.tokenCache.deserialize(Buffer.from(raw, "base64").toString("utf-8"));
      }
    } catch {
      localStorage.removeItem(CACHE_STORAGE_KEY);
    }
  }

  async afterCacheAccess(context: TokenCacheContext): Promise<void> {
    if (!context.cacheHasChanged) return;

    try {
      const serialized = context.tokenCache.serialize();
      if (this.canUseSafeStorage()) {
        const { safeStorage } = require("electron");
        const encrypted = safeStorage.encryptString(serialized);
        localStorage.setItem(CACHE_STORAGE_KEY, encrypted.toString("base64"));
      } else {
        this.warnInsecureFallback();
        localStorage.setItem(CACHE_STORAGE_KEY, Buffer.from(serialized).toString("base64"));
      }
    } catch (e) {
      logger.error("TokenCache", "Failed to persist token cache", e);
    }
  }

  deleteFromCache(): void {
    localStorage.removeItem(CACHE_STORAGE_KEY);
  }
}
