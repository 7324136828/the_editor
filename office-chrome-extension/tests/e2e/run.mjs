/**
 * End-to-end tests of the real extension UI (viewer page) with real input.
 *
 * NOT part of `npm test` — it needs a Chromium-based browser (Edge or Chrome).
 *
 * Run from office-chrome-extension/:
 *     node tests/e2e/run.mjs
 *
 * What it does:
 *   1. launches its OWN headless browser instance (own profile directory and
 *      remote-debugging port) with this folder loaded as an unpacked
 *      extension;
 *   2. turns the office backend OFF in the extension settings, so every
 *      check runs on the built-in readers/writers and needs nothing else
 *      (set E2E_BACKEND_URL=http://127.0.0.1:47831 to run with a backend);
 *   3. drives viewer.html with CDP mouse/keyboard events (typing, clicks,
 *      shortcuts) the way a person would, and checks the document model and
 *      the files written by Save (the save picker is replaced by an in-page
 *      stub that captures the bytes);
 *   4. prints PASS/FAIL per check, kills only the browser process tree it
 *      started, and exits 1 if anything failed.
 *
 * Environment overrides:
 *   BROWSER_PATH       browser executable (default: Edge, then Chrome)
 *   E2E_PORT           remote debugging port (default 9342)
 *   E2E_PROFILE        user-data-dir (default: <os tmpdir>/office-ext-e2e-<pid>, deleted afterwards)
 *   E2E_BACKEND_URL    use this office-backend instead of the built-in engines
 */

import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { spawn, execFileSync } from "node:child_process";
import { fileURLToPath } from "node:url";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const EXT = path.resolve(HERE, "../..");
const EXAMPLE = path.resolve(EXT, "../example");
const PORT = Number(process.env.E2E_PORT || 9342);
const OWN_PROFILE = !process.env.E2E_PROFILE;
const PROFILE = process.env.E2E_PROFILE || path.join(os.tmpdir(), `office-ext-e2e-${process.pid}`);
const BACKEND = process.env.E2E_BACKEND_URL || "";
const CTRL = 2, SHIFT = 8;

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

function killTree(child) {
  if (!child || child.exitCode !== null) return;
  try {
    if (process.platform === "win32") execFileSync("taskkill", ["/PID", String(child.pid), "/T", "/F"], { stdio: "ignore" });
    else child.kill("SIGKILL");
  } catch {
    // already gone
  }
}

async function connect(wsUrl) {
  const ws = new WebSocket(wsUrl);
  await new Promise((res, rej) => { ws.onopen = res; ws.onerror = rej; });
  let id = 0;
  const pending = new Map();
  const logs = [];
  const dialogs = [];
  let send;
  ws.onmessage = (m) => {
    const msg = JSON.parse(m.data);
    if (msg.id && pending.has(msg.id)) {
      const p = pending.get(msg.id);
      pending.delete(msg.id);
      if (msg.error) p.rej(new Error(JSON.stringify(msg.error)));
      else p.res(msg.result);
    } else if (msg.method === "Runtime.exceptionThrown") {
      logs.push(msg.params.exceptionDetails.exception?.description || msg.params.exceptionDetails.text);
    } else if (msg.method === "Page.javascriptDialogOpening") {
      // A native alert/confirm would block the page: record it and dismiss it.
      dialogs.push(msg.params);
      send("Page.handleJavaScriptDialog", { accept: false }).catch(() => {});
    }
  };
  send = (method, params = {}) => new Promise((res, rej) => {
    const i = ++id;
    pending.set(i, { res, rej });
    ws.send(JSON.stringify({ id: i, method, params }));
  });
  return { send, logs, dialogs, close: () => ws.close() };
}

// ---------------------------------------------------------------- page helpers

/** Captures Save output: showSaveFilePicker returns a handle that keeps the bytes. */
const SAVE_STUB = `(()=>{window.__written=[];window.showSaveFilePicker=async(o)=>({name:o.suggestedName,kind:"file",
queryPermission:async()=>"granted",requestPermission:async()=>"granted",getFile:async()=>new File([],o.suggestedName),
createWritable:async()=>{const parts=[];return{write:async(b)=>parts.push(b),close:async()=>{const u=new Uint8Array(await new Blob(parts).arrayBuffer());
let s="";for(let i=0;i<u.length;i+=0x8000)s+=String.fromCharCode.apply(null,u.subarray(i,i+0x8000));window.__written.push({name:o.suggestedName,b64:btoa(s)})},abort:async()=>{}}}});})()`;

function pageApi(c) {
  const p = {
    ...c,
    async eval(expr) {
      const r = await c.send("Runtime.evaluate", { expression: expr, awaitPromise: true, returnByValue: true, userGesture: true });
      if (r.exceptionDetails) throw new Error(r.exceptionDetails.exception?.description || r.exceptionDetails.text);
      return r.result.value;
    },
    async mouse(x, y, clicks = 1, modifiers = 0) {
      await c.send("Input.dispatchMouseEvent", { type: "mouseMoved", x, y });
      for (let n = 1; n <= clicks; n++) {
        for (const type of ["mousePressed", "mouseReleased"]) await c.send("Input.dispatchMouseEvent", { type, x, y, button: "left", clickCount: n, modifiers });
      }
      await sleep(40);
    },
    async click(selector) {
      const r = await p.eval(`(()=>{const e=document.querySelector(${JSON.stringify(selector)}); if(!e) return null; e.scrollIntoView({block:'nearest'}); const b=e.getBoundingClientRect(); return b.width ? {x:b.x+b.width/2,y:b.y+b.height/2} : null})()`);
      if (!r) throw new Error(`not clickable: ${selector}`);
      await p.mouse(r.x, r.y);
    },
    async clickText(selector, text) {
      const r = await p.eval(`(()=>{const e=[...document.querySelectorAll(${JSON.stringify(selector)})].find(n=>n.textContent.trim()===${JSON.stringify(text)} && n.getBoundingClientRect().width); if(!e) return null; const b=e.getBoundingClientRect(); return {x:b.x+b.width/2,y:b.y+b.height/2}})()`);
      if (!r) throw new Error(`no ${selector} "${text}"`);
      await p.mouse(r.x, r.y);
    },
    /** Click just inside the end of the first text occurrence of `needle` under `root`. */
    async clickAfterText(needle, root = "#word-view") {
      const r = await p.eval(`(()=>{const w=document.createTreeWalker(document.querySelector(${JSON.stringify(root)}),NodeFilter.SHOW_TEXT); for(let n=w.nextNode(); n; n=w.nextNode()){const i=n.data.indexOf(${JSON.stringify(needle)}); if(i>=0){n.parentElement.scrollIntoView({block:'center'}); const g=document.createRange(); g.setStart(n,i); g.setEnd(n,i+${needle.length}); const b=g.getBoundingClientRect(); return {x:b.right-1,y:b.y+b.height/2}}} return null})()`);
      if (!r) throw new Error(`text not found: ${needle}`);
      await p.mouse(r.x, r.y);
    },
    async key(k, modifiers = 0) {
      const map = { Enter: [13, "\r"], Tab: [9, ""], Escape: [27, ""], Delete: [46, ""], End: [35, ""] };
      const [vk, text] = map[k] || [k.toUpperCase().charCodeAt(0), k];
      const withText = !!text && !(modifiers & CTRL);
      const code = k.length === 1 ? `Key${k.toUpperCase()}` : k;
      await c.send("Input.dispatchKeyEvent", { type: withText ? "keyDown" : "rawKeyDown", key: k, code, windowsVirtualKeyCode: vk, text: withText ? text : undefined, modifiers });
      await c.send("Input.dispatchKeyEvent", { type: "keyUp", key: k, code, windowsVirtualKeyCode: vk, modifiers });
      await sleep(30);
    },
    async type(s) {
      for (const ch of s) {
        if (ch === "\n") { await p.key("Enter"); continue; }
        await c.send("Input.dispatchKeyEvent", { type: "keyDown", key: ch, text: ch, unmodifiedText: ch });
        await c.send("Input.dispatchKeyEvent", { type: "keyUp", key: ch });
      }
      await sleep(30);
    },
    /** Pick a <select> option the way the native dropdown reports it. */
    async choose(selector, value) {
      await p.eval(`(()=>{const s=document.querySelector(${JSON.stringify(selector)}); s.focus(); s.value=${JSON.stringify(value)}; s.dispatchEvent(new Event('input',{bubbles:true})); s.dispatchEvent(new Event('change',{bubbles:true}))})()`);
    },
    async waitFor(expr, ms = 15000) {
      const t0 = Date.now();
      while (Date.now() - t0 < ms) {
        try { const v = await p.eval(expr); if (v) return v; } catch { /* page still loading */ }
        await sleep(100);
      }
      throw new Error(`timed out waiting for ${expr}`);
    },
    async setFiles(files, selector = "#file-input") {
      const d = await c.send("DOM.getDocument");
      const q = await c.send("DOM.querySelector", { nodeId: d.root.nodeId, selector });
      await c.send("DOM.setFileInputFiles", { nodeId: q.nodeId, files });
    },
    async written(ms = 15000) {
      await p.waitFor("window.__written.length > 0", ms);
      await sleep(100);
      const arr = await p.eval("JSON.stringify(window.__written.splice(0))");
      return JSON.parse(arr).map((w) => ({ name: w.name, bytes: Buffer.from(w.b64, "base64") }));
    },
    state: () => p.eval(`JSON.stringify({name:__officeViewer.doc?.name, kind:__officeViewer.doc?.kind, dirty:__officeViewer.dirty, status:document.getElementById('status-msg').textContent, toast:document.getElementById('toast').hidden?null:document.getElementById('toast').textContent, modals:[...document.querySelectorAll('.modal')].filter(m=>!m.hidden).map(m=>m.id)})`).then(JSON.parse),
    wordModel: () => p.eval("JSON.stringify(__officeViewer.currentWordModel())").then(JSON.parse),
    /** Excel: centre of cell `a1` from the row/column headers. */
    async cellXY(a1) {
      const m = /^([A-Z]+)(\d+)$/.exec(a1);
      let c0 = 0;
      for (const ch of m[1]) c0 = c0 * 26 + ch.charCodeAt(0) - 64;
      return p.eval(`(()=>{const h=document.querySelector('.xv-ch[data-c="${c0 - 1}"]').getBoundingClientRect(); const r=document.querySelector('.xv-rhd[data-r="${+m[2] - 1}"]').getBoundingClientRect(); return {x:h.x+h.width/2,y:r.y+r.height/2}})()`);
    },
    async clickCell(a1) { const xy = await p.cellXY(a1); await p.mouse(xy.x, xy.y); },
    cell: (a1) => p.eval(`JSON.stringify(__officeViewer.doc.model.sheets[__officeViewer.doc.model.active||0].cells[${JSON.stringify(a1)}] ?? null)`).then(JSON.parse),
    nameBox: () => p.eval("document.getElementById('name-box').value"),
  };
  return p;
}

// ---------------------------------------------------------------- tests

const results = [];
function check(name, ok, detail = "") {
  results.push({ name, ok: !!ok });
  console.log(`${ok ? "PASS" : "FAIL"}  ${name}${ok ? "" : `\n      ${typeof detail === "string" ? detail : JSON.stringify(detail)}`}`);
}
const runsOf = (b) => (b.runs || []).map((r) => r.t).join("");
const sig = (m) => m.blocks.map((b) => `${b.type}${b.list ? `:${b.list}${b.level}` : ""}:${runsOf(b)}`);

const TESTS = [
  ["new Word document: heading, paragraph and a numbered list started on an empty line", async (open) => {
    const p = await open("?new=word");
    await p.click("#word-view");
    await p.choose("#fb-style", "h1");
    await p.type("Trip plan\nWe leave on Friday.\n");
    // The caret is on a new, empty line: the list must start right there.
    await p.click('#word-bar [data-cmd="insertOrderedList"]');
    await p.type("Book hotel\nRent car");
    const s = sig(await p.wordModel());
    check("typed structure", JSON.stringify(s) === JSON.stringify(["h1:Trip plan", "p:We leave on Friday.", "list:ol0:Book hotel", "list:ol0:Rent car"]), s);
    // Bullets on a line with text keep the caret on that line too.
    await p.clickAfterText("We leave on Friday.");
    await p.click('#word-bar [data-cmd="insertUnorderedList"]');
    await p.type(" Back Sunday.");
    const s2 = sig(await p.wordModel());
    check("list button keeps the caret in its line", s2[1] === "list:ul0:We leave on Friday. Back Sunday.", s2);
    await p.key("s", CTRL);
    await p.waitFor("!document.getElementById('saveas-modal').hidden");
    const name = await p.eval("document.getElementById('sa-name').value");
    check("Save As suggests the first line as the name of a new document", name === "Trip plan.docx", name);
    await p.click("#sa-ok");
    const [w] = await p.written();
    check("new document saved as .docx", w && w.name === "Trip plan.docx" && w.bytes[0] === 0x50, w && w.name);
  }],

  ["Word table rows: menu does not move the page, and every table edit is undoable", async (open) => {
    const p = await open("");
    await p.setFiles([path.join(EXAMPLE, "quarterly-report.docx")]);
    await p.waitFor("__officeViewer.doc?.kind === 'word'");
    await sleep(500);
    const rows = () => p.eval("document.querySelector('#word-view table').rows.length");
    const original = await p.eval("document.getElementById('word-view').innerHTML");
    await p.clickAfterText("Europe", "#word-view table");
    await p.waitFor("!document.getElementById('table-tools').hidden");
    const y0 = await p.eval("document.querySelector('#word-view table').getBoundingClientRect().y");
    await p.click("#fb-table-menu");
    await p.clickText("#menu-table button", "Insert row below");
    const y1 = await p.eval("document.querySelector('#word-view table').getBoundingClientRect().y");
    check("table menu: row inserted", (await rows()) === 5);
    check("table tools do not shift the document", Math.abs(y1 - y0) < 1, { y0, y1 });
    await p.type("LATAM");
    await p.clickAfterText("+21.0%", "#word-view table"); // last cell of the table
    await p.key("Tab"); // past the last cell: Tab adds a row, like Word
    check("Tab in the last cell adds a row", (await rows()) === 6, await rows());
    await p.key("z", CTRL);
    check("Ctrl+Z removes the Tab-added row", (await rows()) === 5);
    await p.key("z", CTRL); await p.key("z", CTRL);
    check("Ctrl+Z removes typing then the inserted row", (await rows()) === 4 && !(await p.eval("document.getElementById('word-view').textContent.includes('LATAM')")));
    await p.clickAfterText("Europe", "#word-view table");
    await p.click("#fb-table-menu");
    await p.clickText("#menu-table button", "Delete table");
    check("Delete table removes it", !(await p.eval("!!document.querySelector('#word-view table')")));
    await p.key("z", CTRL);
    check("Ctrl+Z brings a deleted table back", (await p.eval("!!document.querySelector('#word-view table')")) && (await rows()) === 4);
    check("after undoing everything the document is unchanged", (await p.eval("document.getElementById('word-view').innerHTML")) === original);
  }],

  ["Excel data entry: Tab across, Enter returns to the first column (like Excel)", async (open) => {
    const p = await open("?new=excel");
    await p.clickCell("A1");
    await p.type("Item"); await p.key("Tab"); await p.type("Cost"); await p.key("Enter");
    check("Enter after Tab goes to A2", (await p.nameBox()) === "A2", await p.nameBox());
    await p.type("Hotel"); await p.key("Tab"); await p.type("240"); await p.key("Enter");
    await p.type("Car"); await p.key("Tab"); await p.type("95.5"); await p.key("Enter");
    await p.type("Average"); await p.key("Tab"); await p.type("=average(b2:b3)"); await p.key("Enter");
    const b4 = await p.cell("B4");
    check("rows line up and AVERAGE computes", b4 && b4.f === "=AVERAGE(B2:B3)" && b4.v === 167.75 && (await p.cell("A3"))?.v === "Car", b4);
    await p.clickCell("D1"); await p.type("x"); await p.key("Enter");
    check("plain Enter still moves straight down", (await p.nameBox()) === "D2", await p.nameBox());
  }],

  ["Excel: deleting a sheet with data asks in the app's own dialog and can be undone", async (open) => {
    const p = await open("");
    await p.setFiles([path.join(EXAMPLE, "sample-workbook.xlsx")]);
    await p.waitFor("__officeViewer.doc?.kind === 'excel'");
    await sleep(400);
    const names = () => p.eval("__officeViewer.doc.model.sheets.map(s => s.name).join(',')");
    const x = await p.eval(`(()=>{const b=document.querySelectorAll('#sheet-tabs [title^="Delete sheet"]')[1].getBoundingClientRect(); return {x:b.x+b.width/2,y:b.y+b.height/2}})()`);
    await p.mouse(x.x, x.y);
    await sleep(300);
    check("no native confirm() box", p.dialogs.length === 0, p.dialogs);
    check("app dialog shown", (await p.state()).modals.includes("dialog-modal"));
    await p.clickText("#dlg-actions button", "Cancel");
    check("Cancel keeps the sheet", (await names()) === "Sales,Headcount");
    await p.mouse(x.x, x.y);
    await sleep(300);
    await p.clickText("#dlg-actions button", "Delete sheet");
    check("confirming deletes it", (await names()) === "Sales");
    await p.clickCell("A1");
    await p.key("z", CTRL);
    check("Ctrl+Z restores it", (await names()) === "Sales,Headcount");
  }],

  ["damaged files: short, plain error and the current document is kept", async (open) => {
    const p = await open("?new=word");
    await p.click("#word-view");
    await p.type("keep me");
    const bad = path.join(PROFILE, "broken.docx");
    fs.writeFileSync(bad, Buffer.from("this is not a zip file at all, just text pretending to be a document"));
    await p.setFiles([bad]);
    await sleep(300);
    const s0 = await p.state();
    check("unsaved changes prompt before opening another file", s0.modals.includes("dialog-modal"), s0);
    await p.clickText("#dlg-actions button", "Don't save");
    await p.waitFor("document.getElementById('status-msg').textContent.startsWith('Could not open')");
    const s = await p.state();
    const msg = s.status;
    check("message names the file and stays short", /^Could not open broken\.docx/.test(msg) && msg.length < 220, msg);
    check("no library internals in the message", !/central directory|jszip|https?:\/\//i.test(msg), msg);
    check("the previous document is still open", s.kind === "word" && (await p.eval("document.getElementById('word-view').textContent")).includes("keep me"), s);
  }],
];

// ---------------------------------------------------------------- main

async function main() {
  fs.mkdirSync(PROFILE, { recursive: true });
  const browser = spawn(findBrowser(), [
    "--headless=new",
    `--remote-debugging-port=${PORT}`,
    `--user-data-dir=${PROFILE}`,
    `--load-extension=${EXT}`,
    `--disable-extensions-except=${EXT}`,
    "--disable-features=DisableLoadExtensionCommandLineSwitch",
    "--no-first-run", "--no-default-browser-check", "--disable-sync",
    "--window-size=1400,900",
    "about:blank",
  ], { stdio: "ignore" });
  let failed = true;
  try {
    let version = null;
    for (let i = 0; i < 100 && !version; i++) {
      try { version = await (await fetch(`http://127.0.0.1:${PORT}/json/version`)).json(); } catch { await sleep(150); }
    }
    if (!version) throw new Error(`browser did not open debug port ${PORT}`);
    const browserConn = await connect(version.webSocketDebuggerUrl);
    // The extension id depends on its path: read it from the service worker.
    let extId = null;
    for (let i = 0; i < 80 && !extId; i++) {
      const list = await (await fetch(`http://127.0.0.1:${PORT}/json/list`)).json();
      const sw = list.find((t) => /\/background\/service-worker\.js$/.test(t.url));
      if (sw) extId = new URL(sw.url).host;
      else await sleep(150);
    }
    if (!extId) throw new Error("the extension did not load");
    const viewer = `chrome-extension://${extId}/viewer/viewer.html`;

    const open = async (query) => {
      const { targetId } = await browserConn.send("Target.createTarget", { url: "about:blank" });
      const list = await (await fetch(`http://127.0.0.1:${PORT}/json/list`)).json();
      const c = await connect(list.find((t) => t.id === targetId).webSocketDebuggerUrl);
      for (const d of ["Runtime", "Page", "DOM"]) await c.send(`${d}.enable`);
      await c.send("Emulation.setFocusEmulationEnabled", { enabled: true });
      await c.send("Emulation.setDeviceMetricsOverride", { width: 1400, height: 900, deviceScaleFactor: 1, mobile: false });
      const p = pageApi(c);
      p.targetId = targetId;
      await c.send("Page.navigate", { url: viewer + query });
      await p.waitFor("!!window.__officeViewer && document.readyState === 'complete'");
      await sleep(400);
      await p.eval(SAVE_STUB);
      return p;
    };

    // Settings: backend off (or the given URL).
    {
      const p = await open("");
      await p.eval(`chrome.storage.local.set({settings: {backend_enabled: ${!!BACKEND}, backend_url: ${JSON.stringify(BACKEND || "http://127.0.0.1:47831")}}})`);
      await browserConn.send("Target.closeTarget", { targetId: p.targetId });
    }
    for (const [name, fn] of TESTS) {
      console.log(`\n# ${name}`);
      const pages = [];
      try {
        await fn(async (q) => { const p = await open(q); pages.push(p); return p; });
      } catch (e) {
        check(`${name}: ran to the end`, false, e.stack || String(e));
      }
      for (const p of pages) {
        if (p.logs.length) check(`${name}: no page exceptions`, false, p.logs.join("\n"));
        p.close();
        // Dirty documents ask before unload; close the target from the browser side.
        await browserConn.send("Target.closeTarget", { targetId: p.targetId }).catch(() => {});
      }
    }
    const fails = results.filter((r) => !r.ok).length;
    console.log(`\n${results.length - fails} passed, ${fails} failed (${results.length} total)`);
    failed = fails > 0;
  } finally {
    killTree(browser);
    if (OWN_PROFILE) {
      await sleep(500);
      try { fs.rmSync(PROFILE, { recursive: true, force: true }); } catch { /* profile still locked */ }
    }
  }
  process.exit(failed ? 1 : 0);
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
