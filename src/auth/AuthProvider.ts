import {
  PublicClientApplication,
  CryptoProvider,
  AccountInfo,
  AuthenticationResult,
  LogLevel,
} from "@azure/msal-node";
import type { AuthenticationProvider } from "@microsoft/microsoft-graph-client";
import * as http from "http";
import { TokenCachePlugin } from "./TokenCachePlugin";
import { GRAPH_SCOPES } from "../constants";
import { logger } from "../utils/logger";
import type { IrisMailSettings, AuthState } from "../types";

/**
 * MSAL PKCE auth provider for Obsidian.
 * Based on obsidian-msgraph-plugin's authProvider.ts pattern.
 * Implements @microsoft/microsoft-graph-client's AuthenticationProvider
 * so it can be passed directly to Client.initWithMiddleware().
 */
export class AuthProvider implements AuthenticationProvider {
  private pca: PublicClientApplication | null = null;
  private account: AccountInfo | null = null;
  private lastToken: string | null = null;
  private lastTokenExpiry: number = 0;
  private cachePlugin = new TokenCachePlugin();
  private cryptoProvider = new CryptoProvider();
  private redirectServer: http.Server | null = null;
  private stateChangeCallback: ((state: AuthState) => void) | null = null;

  onStateChange(callback: (state: AuthState) => void): void {
    this.stateChangeCallback = callback;
  }

  async initialize(settings: IrisMailSettings): Promise<void> {
    this.pca = new PublicClientApplication({
      auth: {
        clientId: settings.clientId,
        authority: settings.authority,
      },
      cache: {
        cachePlugin: this.cachePlugin,
      },
      system: {
        loggerOptions: {
          logLevel: LogLevel.Warning,
        },
      },
    });

    // Restore session from cache
    const accounts = await this.pca.getTokenCache().getAllAccounts();
    if (accounts.length > 0) {
      this.account = accounts[0];
      this.emitState("signed-in");
    }
  }

  async login(settings: IrisMailSettings): Promise<void> {
    if (!this.pca) throw new Error("AuthProvider not initialized");
    this.emitState("signing-in");

    const { verifier, challenge } =
      await this.cryptoProvider.generatePkceCodes();

    const redirectUri = `http://localhost:${settings.redirectPort}/redirect`;

    const authCodeUrl = await this.pca.getAuthCodeUrl({
      scopes: GRAPH_SCOPES,
      redirectUri,
      codeChallenge: challenge,
      codeChallengeMethod: "S256",
    });

    // Open system browser for sign-in
    const { shell } = require("electron");
    shell.openExternal(authCodeUrl);

    // Listen for the redirect with auth code
    const code = await this.listenForAuthCode(settings.redirectPort);

    const result: AuthenticationResult =
      await this.pca.acquireTokenByCode({
        code,
        scopes: GRAPH_SCOPES,
        redirectUri,
        codeVerifier: verifier,
      });

    this.account = result.account;
    this.storeToken(result);
    this.emitState("signed-in");
  }

  /**
   * Device code flow: user visits a URL and enters a code on any device.
   * No localhost server needed — works through firewalls and restricted networks.
   */
  async loginWithDeviceCode(
    settings: IrisMailSettings,
    onUserCode: (code: string, verificationUri: string) => void,
  ): Promise<void> {
    if (!this.pca) throw new Error("AuthProvider not initialized");
    this.emitState("signing-in");

    const result = await this.pca.acquireTokenByDeviceCode({
      scopes: GRAPH_SCOPES,
      deviceCodeCallback: (response) => {
        onUserCode(response.userCode, response.verificationUri);
      },
    });

    if (!result) {
      this.emitState("error");
      throw new Error("Device code authentication failed — no result returned.");
    }

    this.account = result.account;
    this.storeToken(result);
    this.emitState("signed-in");
  }

  /**
   * Called by @microsoft/microsoft-graph-client for every API request.
   * Returns a valid access token, refreshing silently if needed.
   */
  async getAccessToken(): Promise<string> {
    if (!this.pca || !this.account) {
      throw new Error("Not signed in");
    }
    // Use in-memory token if still valid (with 60s buffer)
    if (this.lastToken && Date.now() < this.lastTokenExpiry - 60_000) {
      return this.lastToken;
    }
    try {
      const result = await this.pca.acquireTokenSilent({
        scopes: GRAPH_SCOPES,
        account: this.account,
      });
      this.storeToken(result);
      return result.accessToken;
    } catch (err) {
      logger.warn("Auth", "Silent token failed, retrying with force refresh", err);
      try {
        const result = await this.pca.acquireTokenSilent({
          scopes: GRAPH_SCOPES,
          account: this.account,
          forceRefresh: true,
        });
        this.storeToken(result);
        return result.accessToken;
      } catch (err2) {
        logger.error("Auth", "Force refresh also failed", err2);
        this.account = null;
        this.lastToken = null;
        this.lastTokenExpiry = 0;
        this.emitState("signed-out");
        throw new Error("Token expired. Please sign in again.");
      }
    }
  }

  async logout(): Promise<void> {
    if (this.pca && this.account) {
      const cache = this.pca.getTokenCache();
      await cache.removeAccount(this.account);
    }
    this.cachePlugin.deleteFromCache();
    this.account = null;
    this.lastToken = null;
    this.lastTokenExpiry = 0;
    this.pca = null;
    this.emitState("signed-out");
  }

  isSignedIn(): boolean {
    return this.account !== null;
  }

  getAccount(): AccountInfo | null {
    return this.account;
  }

  destroy(): void {
    if (this.redirectServer) {
      this.redirectServer.close();
      this.redirectServer = null;
    }
  }

  // --- private ---

  private listenForAuthCode(port: number): Promise<string> {
    return new Promise((resolve, reject) => {
      this.redirectServer = http.createServer((req, res) => {
        const url = new URL(req.url!, `http://localhost:${port}`);
        const code = url.searchParams.get("code");
        const error = url.searchParams.get("error");

        if (code) {
          res.writeHead(200, { "Content-Type": "text/html" });
          res.end(
            "<html><body><h2>Signed in successfully.</h2>" +
              "<p>You can close this window and return to Obsidian.</p></body></html>",
          );
          this.redirectServer!.close();
          this.redirectServer = null;
          resolve(code);
        } else {
          res.writeHead(400, { "Content-Type": "text/html" });
          res.end(`<html><body><h2>Error: ${error || "Unknown"}</h2></body></html>`);
          this.redirectServer!.close();
          this.redirectServer = null;
          reject(new Error(error || "No authorization code received"));
        }
      });

      this.redirectServer.on("error", (err: NodeJS.ErrnoException) => {
        if (err.code === "EADDRINUSE") {
          reject(
            new Error(
              `Port ${port} is already in use. Change the redirect port in settings.`,
            ),
          );
        } else {
          reject(err);
        }
      });

      this.redirectServer.listen(port, "127.0.0.1");

      // Timeout after 5 minutes
      setTimeout(() => {
        if (this.redirectServer) {
          this.redirectServer.close();
          this.redirectServer = null;
          reject(new Error("Login timed out after 5 minutes"));
        }
      }, 300_000);
    });
  }

  private storeToken(result: AuthenticationResult): void {
    this.lastToken = result.accessToken;
    this.lastTokenExpiry = result.expiresOn
      ? result.expiresOn.getTime()
      : Date.now() + 3600_000;
  }

  private emitState(state: AuthState): void {
    this.stateChangeCallback?.(state);
  }
}
