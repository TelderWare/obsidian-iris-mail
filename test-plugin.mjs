#!/usr/bin/env node
/**
 * test-plugin.mjs — Build, reload plugin, and run commands inside Obsidian.
 *
 * Two modes:
 *   CDP (default) — Connects to Obsidian's Electron debug port. Hot-reloads the
 *                   plugin and executes commands without restarting Obsidian.
 *   File-based    — Writes .test-request.json, restarts Obsidian, and polls for
 *                   .test-results.json (uses the testBridge in the plugin).
 *                   Triggered automatically when CDP isn't available.
 *
 * Usage:
 *   node test-plugin.mjs [flags] [command ...]
 *
 * Flags:
 *   --fresh       Force-kill and restart Obsidian (uses file-based mode)
 *   --no-build    Skip npm run build
 *   --list        List all registered commands for this plugin
 *   --eval "js"   Evaluate arbitrary JS in Obsidian after reload
 *
 * Commands are Obsidian command IDs (plugin prefix optional).
 * Default: open-iris-mail
 *
 * For CDP mode: Node.js 22+ (native WebSocket) and Obsidian must be running
 * with --remote-debugging-port=9222 (the script will restart it this way
 * on first use).
 */

import { execSync, spawn } from "node:child_process";
import { readFileSync, writeFileSync, existsSync, unlinkSync } from "node:fs";
import { setTimeout as sleep } from "node:timers/promises";
import { basename, dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = dirname(fileURLToPath(import.meta.url));

// ── Auto-detect from project layout ─────────────────────────────────
const manifest = JSON.parse(
  readFileSync(resolve(__dirname, "manifest.json"), "utf-8"),
);
const PLUGIN_ID = manifest.id;
const VAULT_PATH = resolve(__dirname, "..", "..", ".."); // .obsidian/plugins/<id>
const VAULT_NAME = basename(VAULT_PATH);
const DEBUG_PORT = 9222;
const OBSIDIAN_EXE = resolve(
  process.env.LOCALAPPDATA || "",
  "Obsidian",
  "Obsidian.exe",
);
const PLUGIN_LOAD_WAIT = 4000;
const DEFAULT_COMMANDS = ["open-iris-mail"];

// File-based paths (testBridge integration)
const REQUEST_FILE = resolve(__dirname, ".test-request.json");
const RESULTS_FILE = resolve(__dirname, ".test-results.json");

// ── Parse args ──────────────────────────────────────────────────────
const argv = process.argv.slice(2);
const fresh = argv.includes("--fresh");
const noBuild = argv.includes("--no-build");
const listMode = argv.includes("--list");

let evalExpr = null;
const evalIdx = argv.indexOf("--eval");
if (evalIdx !== -1) evalExpr = argv[evalIdx + 1];

const commands = argv.filter(
  (a, i) => !a.startsWith("--") && i !== evalIdx + 1,
);
if (!listMode && !evalExpr && commands.length === 0) {
  commands.push(...DEFAULT_COMMANDS);
}

// Fully qualify command IDs
function fullCommandId(cmd) {
  return cmd.includes(":") ? cmd : `${PLUGIN_ID}:${cmd}`;
}

// ── CDP helpers ─────────────────────────────────────────────────────
async function getCDPTarget() {
  const res = await fetch(`http://127.0.0.1:${DEBUG_PORT}/json`);
  const targets = await res.json();
  return targets.find((t) => t.type === "page") || targets[0];
}

function connectCDP(wsUrl) {
  return new Promise((ok, fail) => {
    const ws = new WebSocket(wsUrl);
    let nextId = 1;
    const pending = new Map();

    ws.addEventListener("open", () =>
      ok({
        eval(expression) {
          const id = nextId++;
          return new Promise((res, rej) => {
            pending.set(id, { res, rej });
            ws.send(
              JSON.stringify({
                id,
                method: "Runtime.evaluate",
                params: {
                  expression,
                  returnByValue: true,
                  awaitPromise: true,
                },
              }),
            );
          });
        },
        close() {
          ws.close();
        },
      }),
    );

    ws.addEventListener("message", (evt) => {
      const msg = JSON.parse(
        typeof evt.data === "string" ? evt.data : String(evt.data),
      );
      const p = pending.get(msg.id);
      if (!p) return;
      pending.delete(msg.id);
      if (msg.result?.exceptionDetails) {
        p.rej(
          new Error(
            msg.result.exceptionDetails.exception?.description ||
              msg.result.exceptionDetails.text ||
              "eval error",
          ),
        );
      } else {
        p.res(msg.result?.result?.value);
      }
    });

    ws.addEventListener("error", () => fail(new Error("CDP WebSocket error")));
  });
}

// ── Shared actions ──────────────────────────────────────────────────
function build() {
  console.log("[build] compiling...");
  execSync("npm run build", { stdio: "inherit", cwd: __dirname });
}

async function killObsidian() {
  try {
    execSync("taskkill /f /im Obsidian.exe", { stdio: "pipe" });
    await sleep(1500);
  } catch {
    /* not running */
  }
}

function launchObsidian(withDebugPort) {
  const args = withDebugPort
    ? [`--remote-debugging-port=${DEBUG_PORT}`]
    : [];
  console.log(
    `[start] Obsidian${withDebugPort ? ` (debug port ${DEBUG_PORT})` : ""}`,
  );

  // Try obsidian:// URI to open the right vault
  try {
    const uri = `obsidian://open?vault=${encodeURIComponent(VAULT_NAME)}`;
    if (withDebugPort) {
      // Launch exe directly with debug port, then open vault URI
      const child = spawn(OBSIDIAN_EXE, args, {
        detached: true,
        stdio: "ignore",
      });
      child.unref();
    } else {
      execSync(`start "" "${uri}"`, { shell: "cmd.exe", stdio: "ignore" });
    }
  } catch {
    const child = spawn(OBSIDIAN_EXE, args, {
      detached: true,
      stdio: "ignore",
    });
    child.unref();
  }
}

async function waitForCDP(sec = 30) {
  process.stdout.write("[cdp] waiting");
  for (let i = 0; i < sec; i++) {
    try {
      const t = await getCDPTarget();
      if (t) {
        process.stdout.write(" ok\n");
        return t;
      }
    } catch {
      /* not ready */
    }
    process.stdout.write(".");
    await sleep(1000);
  }
  process.stdout.write(" timeout\n");
  return null;
}

// ── CDP mode ────────────────────────────────────────────────────────
async function runCDP() {
  // Try to connect to existing Obsidian with debug port
  let target;
  try {
    target = await getCDPTarget();
    console.log("[cdp] connected to running Obsidian");
  } catch {
    // No debug port — restart Obsidian with one
    console.log("[cdp] no debug port — restarting Obsidian...");
    await killObsidian();
    launchObsidian(true);
    target = await waitForCDP();
    if (!target) {
      console.error("Could not connect to Obsidian via CDP.");
      process.exit(1);
    }
  }

  const conn = await connectCDP(target.webSocketDebuggerUrl);

  try {
    // Hot-reload plugin to pick up fresh build
    console.log(`[reload] ${PLUGIN_ID}`);
    await conn.eval(`app.plugins.disablePlugin("${PLUGIN_ID}")`);
    await sleep(500);
    await conn.eval(`app.plugins.enablePlugin("${PLUGIN_ID}")`);

    console.log(`[wait] plugin init (${PLUGIN_LOAD_WAIT / 1000}s)...`);
    await sleep(PLUGIN_LOAD_WAIT);

    // --list
    if (listMode) {
      const cmds = await conn.eval(`
        Object.keys(app.commands.commands)
          .filter(id => id.startsWith("${PLUGIN_ID}:"))
      `);
      console.log(`\nRegistered commands for ${PLUGIN_ID}:\n`);
      for (const id of cmds || []) console.log(`  ${id}`);
      console.log();
      return;
    }

    // --eval
    if (evalExpr) {
      console.log(`[eval] ${evalExpr}\n`);
      const result = await conn.eval(evalExpr);
      console.log(JSON.stringify(result, null, 2));
      return;
    }

    // Run commands
    console.log(`\n[test] ${commands.length} command(s):\n`);
    let passed = 0;
    let failed = 0;

    for (const cmd of commands) {
      const fullId = fullCommandId(cmd);
      try {
        const r = await conn.eval(`
          (async () => {
            const c = app.commands.commands[${JSON.stringify(fullId)}];
            if (!c) return { ok: false, error: "not registered" };
            try {
              if (c.callback) { await c.callback(); }
              else if (c.checkCallback) {
                if (!c.checkCallback(true))
                  return { ok: false, error: "checkCallback: unavailable" };
                await c.checkCallback(false);
              }
              return { ok: true };
            } catch (e) { return { ok: false, error: e.message }; }
          })()
        `);
        if (r?.ok) {
          console.log(`  PASS  ${fullId}`);
          passed++;
        } else {
          console.log(`  FAIL  ${fullId} -- ${r?.error}`);
          failed++;
        }
      } catch (e) {
        console.log(`  FAIL  ${fullId} -- ${e.message}`);
        failed++;
      }
    }

    console.log(`\n${passed} passed, ${failed} failed\n`);
    if (failed) process.exit(1);
  } finally {
    conn.close();
  }
}

// ── File-based mode (fallback / --fresh) ────────────────────────────
async function runFileBased() {
  const fullCommands = commands.map(fullCommandId);

  // Write request file for testBridge
  writeFileSync(REQUEST_FILE, JSON.stringify({ commands: fullCommands }));
  console.log(`[file] queued: ${fullCommands.join(", ")}`);

  // Clean stale results
  if (existsSync(RESULTS_FILE)) unlinkSync(RESULTS_FILE);

  // Restart Obsidian
  await killObsidian();
  launchObsidian(false);

  // Poll for results
  console.log("[file] waiting for results...");
  const timeout = 30_000;
  const start = Date.now();
  while (!existsSync(RESULTS_FILE) && Date.now() - start < timeout) {
    await sleep(500);
  }

  if (!existsSync(RESULTS_FILE)) {
    console.error(`Timed out after ${timeout / 1000}s.`);
    if (existsSync(REQUEST_FILE)) unlinkSync(REQUEST_FILE);
    process.exit(1);
  }

  const output = JSON.parse(readFileSync(RESULTS_FILE, "utf-8"));
  console.log(`\n[test] ${output.results.length} command(s):\n`);
  for (const r of output.results) {
    const status = r.status === "ok" ? "PASS" : "FAIL";
    const err = r.error ? ` -- ${r.error}` : "";
    console.log(`  ${status}  ${r.id}${err} (${r.ms}ms)`);
  }
  const allOk = output.results.every((r) => r.status === "ok");
  console.log(`\n${allOk ? "all passed" : "FAILURES"}\n`);
  unlinkSync(RESULTS_FILE);
  if (!allOk) process.exit(1);
}

// ── Main ────────────────────────────────────────────────────────────
async function main() {
  const hasCDP = typeof globalThis.WebSocket !== "undefined";

  console.log(
    `\nplugin: ${PLUGIN_ID}  vault: ${VAULT_NAME}  mode: ${fresh ? "fresh" : hasCDP ? "cdp" : "file-based"}\n`,
  );

  if (!noBuild) build();

  if (fresh || !hasCDP) {
    if (!hasCDP && !fresh) {
      console.log(
        "[warn] Node.js 22+ needed for CDP mode (current: " +
          process.version +
          "). Falling back to file-based mode.\n",
      );
    }
    await runFileBased();
  } else {
    await runCDP();
  }
}

main().catch((e) => {
  console.error(e.message);
  process.exit(1);
});
