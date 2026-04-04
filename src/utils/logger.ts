/**
 * Debug logging utility for Iris Mail.
 * All log output is gated behind a runtime flag so it can be
 * enabled/disabled without restarting Obsidian.
 */

let debugEnabled = false;

export function setDebugEnabled(enabled: boolean): void {
  debugEnabled = enabled;
}

export function isDebugEnabled(): boolean {
  return debugEnabled;
}

function timestamp(): string {
  return new Date().toISOString().slice(11, 23); // HH:MM:SS.mmm
}

export const logger = {
  debug(tag: string, message: string, ...args: unknown[]): void {
    if (debugEnabled) {
      console.debug(`[Iris ${timestamp()}] [${tag}]`, message, ...args);
    }
  },

  info(tag: string, message: string, ...args: unknown[]): void {
    console.log(`[Iris ${timestamp()}] [${tag}]`, message, ...args);
  },

  warn(tag: string, message: string, ...args: unknown[]): void {
    console.warn(`[Iris ${timestamp()}] [${tag}]`, message, ...args);
  },

  error(tag: string, message: string, ...args: unknown[]): void {
    console.error(`[Iris ${timestamp()}] [${tag}]`, message, ...args);
  },
};
