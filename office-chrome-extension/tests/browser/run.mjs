/**
 * Browser test runner for the Word editor model (tests/browser/word-dom.html).
 *
 * NOT part of `npm test` — it needs a Chromium-based browser (Edge or Chrome).
 *
 * Run from office-chrome-extension/:
 *     node tests/browser/run.mjs
 *
 * What it does:
 *   1. serves the repository root (office-extensions/) on a random localhost
 *      port (ES modules cannot be loaded from file://);
 *   2. launches its OWN headless browser instance (own profile directory and
 *      remote-debugging port), opens word-dom.html and waits for the in-page
 *      tests to finish;
 *   3. prints PASS/FAIL per case, kills only the browser process tree it
 *      started, and exits 1 if anything failed.
 *
 * Environment overrides:
 *   BROWSER_PATH          browser executable (default: Edge, then Chrome, at the usual Windows/macOS/Linux paths)
 *   BROWSER_TEST_PORT     remote debugging port (default 9341)
 *   BROWSER_TEST_PROFILE  user-data-dir (default: <os tmpdir>/office-ext-word-browser-tests)
 *   BROWSER_TEST_FILTER   only print cases whose name contains this text
 */

import http from "node:http";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { spawn, execFileSync } from "node:child_process";
import { fileURLToPath } from "node:url";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const REPO = path.resolve(HERE, "../../..");
const PAGE = "/office-chrome-extension/tests/browser/word-dom.html";
const DEBUG_PORT = Number(process.env.BROWSER_TEST_PORT || 9341);
const PROFILE = process.env.BROWSER_TEST_PROFILE || path.join(os.tmpdir(), "office-ext-word-browser-tests");
const FILTER = process.env.BROWSER_TEST_FILTER || "";
const TIMEOUT_MS = 120000;

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

function findBrowser() {
  if (process.env.BROWSER_PATH) return process.env.BROWSER_PATH;
  const candidates = [
    "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe",
    "C:/Program Files/Microsoft/Edge/Application/msedge.exe",
    "C:/Program Files/Google/Chrome/Application/chrome.exe",
    "C:/Program Files (x86)/Google/Chrome/Application/chrome.exe",
    "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    "/usr/bin/microsoft-edge", "/usr/bin/google-chrome", "/usr/bin/chromium", "/usr/bin/chromium-browser",
  ];
  const found = candidates.find((c) => fs.existsSync(c));
  if (!found) throw new Error("No Edge/Chrome found; set BROWSER_PATH");
  return found;
}

const MIME = {
  ".html": "text/html; charset=utf-8", ".js": "text/javascript; charset=utf-8", ".mjs": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8", ".json": "application/json", ".png": "image/png", ".svg": "image/svg+xml",
  ".docx": "application/octet-stream", ".xlsx": "application/octet-stream",
};

function startServer() {
  const server = http.createServer((req, res) => {
    const url = new URL(req.url, "http://localhost");
    const file = path.resolve(REPO, "." + decodeURIComponent(url.pathname));
    if (!file.startsWith(REPO) || !fs.existsSync(file) || fs.statSync(file).isDirectory()) {
      res.writeHead(404);
      res.end("not found");
      return;
    }
    res.writeHead(200, { "Content-Type": MIME[path.extname(file).toLowerCase()] || "application/octet-stream", "Cache-Control": "no-store" });
    fs.createReadStream(file).pipe(res);
  });
  return new Promise((resolve) => server.listen(0, "127.0.0.1", () => resolve(server)));
}

async function connect(wsUrl) {
  const ws = new WebSocket(wsUrl);
  await new Promise((res, rej) => {
    ws.onopen = res;
    ws.onerror = rej;
  });
  let id = 0;
  const pending = new Map();
  const logs = [];
  ws.onmessage = (m) => {
    const msg = JSON.parse(m.data);
    if (msg.id && pending.has(msg.id)) {
      const p = pending.get(msg.id);
      pending.delete(msg.id);
      if (msg.error) p.rej(new Error(JSON.stringify(msg.error)));
      else p.res(msg.result);
    } else if (msg.method === "Runtime.exceptionThrown") {
      logs.push(msg.params.exceptionDetails.exception?.description || msg.params.exceptionDetails.text);
    } else if (msg.method === "Runtime.consoleAPICalled" && msg.params.type === "error") {
      logs.push(msg.params.args.map((a) => a.value ?? a.description).join(" "));
    }
  };
  const send = (method, params = {}) => new Promise((res, rej) => {
    const i = ++id;
    pending.set(i, { res, rej });
    ws.send(JSON.stringify({ id: i, method, params }));
  });
  return { send, logs, close: () => ws.close() };
}

function killTree(child) {
  if (!child || child.exitCode !== null) return;
  try {
    if (process.platform === "win32") execFileSync("taskkill", ["/PID", String(child.pid), "/T", "/F"], { stdio: "ignore" });
    else child.kill("SIGKILL");
  } catch {
    // already gone
  }
}

async function main() {
  const server = await startServer();
  const port = server.address().port;
  fs.mkdirSync(PROFILE, { recursive: true });
  const browser = spawn(findBrowser(), [
    "--headless=new",
    `--remote-debugging-port=${DEBUG_PORT}`,
    `--user-data-dir=${PROFILE}`,
    "--no-first-run", "--no-default-browser-check", "--disable-extensions", "--disable-sync",
    "about:blank",
  ], { stdio: "ignore" });

  let failed = 1;
  try {
    let version = null;
    for (let i = 0; i < 100 && !version; i++) {
      try {
        version = await (await fetch(`http://127.0.0.1:${DEBUG_PORT}/json/version`)).json();
      } catch {
        await sleep(150);
      }
    }
    if (!version) throw new Error(`browser did not open debug port ${DEBUG_PORT}`);

    const b = await connect(version.webSocketDebuggerUrl);
    const { targetId } = await b.send("Target.createTarget", { url: "about:blank" });
    b.close();
    const targets = await (await fetch(`http://127.0.0.1:${DEBUG_PORT}/json/list`)).json();
    const t = targets.find((x) => x.id === targetId);
    const page = await connect(t.webSocketDebuggerUrl);
    await page.send("Runtime.enable");
    await page.send("Page.enable");
    // Headless windows are never focused; execCommand-based editing needs focus.
    await page.send("Emulation.setFocusEmulationEnabled", { enabled: true });
    await page.send("Page.navigate", { url: `http://127.0.0.1:${port}${PAGE}` });

    const start = Date.now();
    let results = null;
    while (Date.now() - start < TIMEOUT_MS) {
      await sleep(250);
      const r = await page.send("Runtime.evaluate", {
        expression: "window.__done ? JSON.stringify(window.__results) : null",
        returnByValue: true,
      });
      if (r.result && r.result.value) {
        results = JSON.parse(r.result.value);
        break;
      }
    }
    if (!results) throw new Error(`tests did not finish in ${TIMEOUT_MS} ms\n${page.logs.join("\n")}`);
    page.close();

    let pass = 0;
    let fail = 0;
    for (const r of results) {
      if (r.ok) pass++;
      else fail++;
      if (FILTER && !r.name.includes(FILTER)) continue;
      console.log(`${r.ok ? "PASS" : "FAIL"}  ${r.name}${r.ok ? "" : `\n      ${r.detail.replace(/\n/g, "\n      ")}`}`);
    }
    if (page.logs.length) console.log(`\npage errors:\n  ${page.logs.join("\n  ")}`);
    console.log(`\n${pass} passed, ${fail} failed (${results.length} total)`);
    failed = fail ? 1 : 0;
  } finally {
    killTree(browser);
    server.close();
  }
  process.exit(failed);
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
