import type { IrisMailSettings } from "../types";

/** Lightweight, provider-neutral representation of the signed-in account. */
export interface AuthAccount {
  /** Email address. */
  username: string;
  /** Display name, if known. */
  name?: string;
}

/**
 * Provider-neutral auth surface used by the plugin. Each backend (Outlook via
 * MSAL, IMAP via app password, ...) implements this and hands an access token
 * (or equivalent credential) to its mail API client.
 */
export interface IAuthProvider {
  initialize(settings: IrisMailSettings): Promise<void>;
  login(settings: IrisMailSettings): Promise<void>;
  /** Optional — only providers that support device-code flow (currently Outlook). */
  loginWithDeviceCode?(
    settings: IrisMailSettings,
    onUserCode: (code: string, verificationUri: string) => void,
  ): Promise<void>;
  logout(): Promise<void>;
  isSignedIn(): boolean;
  getAccount(): AuthAccount | null;
  getAccessToken(): Promise<string>;
  destroy(): void;
}
