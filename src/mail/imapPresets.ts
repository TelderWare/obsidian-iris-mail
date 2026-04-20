/** Common IMAP presets so the user doesn't have to memorize host/port. */
export interface ImapPreset {
  key: string;
  label: string;
  host: string;
  port: number;
  secure: boolean;
  /** Page where the user creates an app password for this provider. */
  appPasswordUrl?: string;
  /** Short hint shown in the settings panel. */
  hint?: string;
}

export const IMAP_PRESETS: ImapPreset[] = [
  {
    key: "gmail",
    label: "Gmail",
    host: "imap.gmail.com",
    port: 993,
    secure: true,
    appPasswordUrl: "https://myaccount.google.com/apppasswords",
    hint: "Requires 2-Step Verification. Create an app password and paste it below.",
  },
  {
    key: "icloud",
    label: "iCloud",
    host: "imap.mail.me.com",
    port: 993,
    secure: true,
    appPasswordUrl: "https://account.apple.com/account/manage",
    hint: "Generate an app-specific password from your Apple ID page.",
  },
  {
    key: "fastmail",
    label: "Fastmail",
    host: "imap.fastmail.com",
    port: 993,
    secure: true,
    appPasswordUrl: "https://app.fastmail.com/settings/security/devicekeys",
    hint: "Create an app password with Mail (IMAP) access.",
  },
  {
    key: "outlook",
    label: "Outlook.com",
    host: "outlook.office365.com",
    port: 993,
    secure: true,
    appPasswordUrl: "https://account.microsoft.com/security",
    hint: "Requires 2FA. Create an app password under Security → Advanced security options.",
  },
  {
    key: "yahoo",
    label: "Yahoo",
    host: "imap.mail.yahoo.com",
    port: 993,
    secure: true,
    appPasswordUrl: "https://login.yahoo.com/account/security",
    hint: "Generate an app password under Account Security.",
  },
  {
    key: "other",
    label: "Other / custom",
    host: "",
    port: 993,
    secure: true,
  },
];

export function getImapPreset(key: string | undefined): ImapPreset | undefined {
  if (!key) return undefined;
  return IMAP_PRESETS.find((p) => p.key === key);
}
