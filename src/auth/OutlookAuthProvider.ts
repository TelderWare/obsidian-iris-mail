import {
  PublicClientApplication,
  CryptoProvider,
  AccountInfo,
  AuthenticationResult,
  LogLevel,
} from "@azure/msal-node";
import type { AuthenticationProvider } from "@microsoft/microsoft-graph-client";
import { TokenCachePlugin } from "./TokenCachePlugin";
import { awaitAuthCode } from "./loopbackRedirect";
import { GRAPH_SCOPES, MSAL_AUTHORITY_DEFAULT } from "../constants";
import { logger } from "../utils/logger";
import type { IrisMailSettings, Account } from "../types";
import type { IAuthProvider, AuthAccount } from "./IAuthProvider";

/**
 * MSAL PKCE auth provider for Outlook. Implements both the plugin's neutral
 * IAuthProvider and microsoft-graph-client's AuthenticationProvider so it can
 * be passed directly to Client.initWithMiddleware().
 */
export class OutlookAuthProvider implements IAuthProvider, AuthenticationProvider {
  private pca: PublicClientApplication | null = null;
  private msalAccount: AccountInfo | null = null;
  private lastToken: string | null = null;
  private lastTokenExpiry: number = 0;
  private cachePlugin: TokenCachePlugin;
  private cryptoProvider = new CryptoProvider();

  constructor(private readonly account: Account) {
    this.cachePlugin = new TokenCachePlugin(account.id);
  }

  async initialize(_settings: IrisMailSettings): Promise<void> {
    this.pca = new PublicClientApplication({
      auth: {
        clientId: this.account.clientId ?? "",
        authority: this.account.authority || MSAL_AUTHORITY_DEFAULT,
      },
      cache: { cachePlugin: this.cachePlugin },
      system: { loggerOptions: { logLevel: LogLevel.Warning } },
    });

    const accounts = await this.pca.getTokenCache().getAllAccounts();
    if (accounts.length > 0) {
      this.msalAccount = accounts[0];
    }
  }

  async login(settings: IrisMailSettings): Promise<void> {
    if (!this.pca) throw new Error("OutlookAuthProvider not initialized");

    const { verifier, challenge } = await this.cryptoProvider.generatePkceCodes();
    const redirectUri = `http://localhost:${settings.redirectPort}/redirect`;

    const authCodeUrl = await this.pca.getAuthCodeUrl({
      scopes: GRAPH_SCOPES,
      redirectUri,
      codeChallenge: challenge,
      codeChallengeMethod: "S256",
    });

    const { shell } = require("electron");
    shell.openExternal(authCodeUrl);

    const code = await awaitAuthCode(settings.redirectPort, { host: "localhost" });

    const result = await this.pca.acquireTokenByCode({
      code,
      scopes: GRAPH_SCOPES,
      redirectUri,
      codeVerifier: verifier,
    });

    this.msalAccount = result.account;
    this.storeToken(result);
  }

  async loginWithDeviceCode(
    _settings: IrisMailSettings,
    onUserCode: (code: string, verificationUri: string) => void,
  ): Promise<void> {
    if (!this.pca) throw new Error("OutlookAuthProvider not initialized");

    const result = await this.pca.acquireTokenByDeviceCode({
      scopes: GRAPH_SCOPES,
      deviceCodeCallback: (response) => onUserCode(response.userCode, response.verificationUri),
    });

    if (!result) {
      throw new Error("Device code authentication failed — no result returned.");
    }

    this.msalAccount = result.account;
    this.storeToken(result);
  }

  async getAccessToken(): Promise<string> {
    if (!this.pca || !this.msalAccount) {
      throw new Error("Not signed in");
    }
    if (this.lastToken && Date.now() < this.lastTokenExpiry - 60_000) {
      return this.lastToken;
    }
    try {
      const result = await this.pca.acquireTokenSilent({
        scopes: GRAPH_SCOPES,
        account: this.msalAccount,
      });
      this.storeToken(result);
      return result.accessToken;
    } catch (err) {
      logger.warn("Auth", "Silent token failed, retrying with force refresh", err);
      try {
        const result = await this.pca.acquireTokenSilent({
          scopes: GRAPH_SCOPES,
          account: this.msalAccount,
          forceRefresh: true,
        });
        this.storeToken(result);
        return result.accessToken;
      } catch (err2) {
        logger.error("Auth", "Force refresh also failed", err2);
        this.msalAccount = null;
        this.lastToken = null;
        this.lastTokenExpiry = 0;
        throw new Error("Token expired. Please sign in again.");
      }
    }
  }

  async logout(): Promise<void> {
    if (this.pca && this.msalAccount) {
      await this.pca.getTokenCache().removeAccount(this.msalAccount);
    }
    this.cachePlugin.deleteFromCache();
    this.msalAccount = null;
    this.lastToken = null;
    this.lastTokenExpiry = 0;
    this.pca = null;
  }

  isSignedIn(): boolean {
    return this.msalAccount !== null;
  }

  getAccount(): AuthAccount | null {
    if (!this.msalAccount) return null;
    return {
      username: this.msalAccount.username,
      name: this.msalAccount.name,
    };
  }

  destroy(): void {
    /* no resources to release */
  }

  private storeToken(result: AuthenticationResult): void {
    this.lastToken = result.accessToken;
    this.lastTokenExpiry = result.expiresOn
      ? result.expiresOn.getTime()
      : Date.now() + 3600_000;
  }
}
