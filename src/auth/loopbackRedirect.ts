import * as http from "http";

const TIMEOUT_MS = 300_000;

/**
 * Bind an HTTP server to 127.0.0.1:{port} and resolve with the OAuth `code`
 * query param from the first request. If `expectedState` is supplied the
 * `state` param must match, otherwise the request is rejected.
 *
 * Used by the Outlook (MSAL PKCE) flow. Closes the server on first request,
 * on error, and after a 5-minute timeout.
 */
export function awaitAuthCode(
  port: number,
  options: { expectedState?: string; host?: string } = {},
): Promise<string> {
  const host = options.host ?? "127.0.0.1";
  return new Promise((resolve, reject) => {
    let server: http.Server | null = null;
    const close = () => {
      server?.close();
      server = null;
    };

    server = http.createServer((req, res) => {
      const url = new URL(req.url!, `http://${host}:${port}`);
      const code = url.searchParams.get("code");
      const state = url.searchParams.get("state");
      const error = url.searchParams.get("error");

      const stateOk = options.expectedState === undefined || state === options.expectedState;

      if (code && stateOk) {
        res.writeHead(200, { "Content-Type": "text/html" });
        res.end(
          "<html><body><h2>Signed in successfully.</h2>" +
            "<p>You can close this window and return to Obsidian.</p></body></html>",
        );
        close();
        resolve(code);
      } else {
        res.writeHead(400, { "Content-Type": "text/html" });
        res.end(`<html><body><h2>Error: ${error || (stateOk ? "no code" : "state mismatch")}</h2></body></html>`);
        close();
        reject(new Error(error || (stateOk ? "No authorization code received" : "OAuth state mismatch")));
      }
    });

    server.on("error", (err: NodeJS.ErrnoException) => {
      if (err.code === "EADDRINUSE") {
        reject(new Error(`Port ${port} is already in use. Change the redirect port in settings.`));
      } else {
        reject(err);
      }
    });

    server.listen(port, host);

    setTimeout(() => {
      if (server) {
        close();
        reject(new Error("Login timed out after 5 minutes"));
      }
    }, TIMEOUT_MS);
  });
}
