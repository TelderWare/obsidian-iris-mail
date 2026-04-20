import type { ICachePlugin, TokenCacheContext } from "@azure/msal-node";
import { CACHE_STORAGE_KEY } from "../constants";
import { logger } from "../utils/logger";
import { encryptString, decryptString } from "../utils/safeStorage";

export class TokenCachePlugin implements ICachePlugin {
  private storageKey: string;

  constructor(accountId?: string) {
    this.storageKey = accountId ? `${CACHE_STORAGE_KEY}:${accountId}` : CACHE_STORAGE_KEY;
  }

  async beforeCacheAccess(context: TokenCacheContext): Promise<void> {
    const raw = localStorage.getItem(this.storageKey);
    if (!raw) return;
    try {
      context.tokenCache.deserialize(decryptString(raw));
    } catch {
      localStorage.removeItem(this.storageKey);
    }
  }

  async afterCacheAccess(context: TokenCacheContext): Promise<void> {
    if (!context.cacheHasChanged) return;
    try {
      localStorage.setItem(this.storageKey, encryptString(context.tokenCache.serialize()));
    } catch (e) {
      logger.error("TokenCache", "Failed to persist token cache", e);
    }
  }

  deleteFromCache(): void {
    localStorage.removeItem(this.storageKey);
  }
}
