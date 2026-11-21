/**
 * End-to-end tests of the PowerPoint editor (viewer/slides.js) with real input.
 *
 * NOT part of `npm test` — it needs a Chromium-based browser (Edge or Chrome).
 *
 * Run from office-chrome-extension/:
 *     node tests/e2e/slides.mjs
 *
 * Launches its own headless browser (own profile and debugging port) with the
 * extension loaded, turns the office backend OFF (everything runs on the
 * built-in reader/writer), drives viewer.html with CDP mouse/keyboard events and
 * checks the model, the UI and the .pptx written by Save (the save picker is
 * replaced by an in-page stub that captures the bytes). Saved files are checked
 * by re-reading them with python-pptx when a Python with python-pptx is
 * available (env PYTHON, default "python").
 *
 * Environment:
 *   BROWSER_PATH   browser executable (default: Edge, then Chrome)
 *   E2E_PORT       remote debugging port (default 9351)
 *   E2E_PROFILE    user-data-dir (default: <os tmpdir>/office-ext-slides-<pid>, deleted afterwards)
 *   E2E_SHOTS      directory for screenshots (default: none)
 *   E2E_ONLY       run only tests whose name contains this text
 *   PYTHON         python with python-pptx for validating saved files
 */

import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { spawn, execFileSync, spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const EXT = path.resolve(HERE, "../..");
const FIX = path.resolve(EXT, "tests/fixtures/slides");
const PORT = Number(process.env.E2E_PORT || 9351);
const OWN_PROFILE = !process.env.E2E_PROFILE;
const PROFILE = process.env.E2E_PROFILE || path.join(os.tmpdir(), `office-ext-slides-${process.pid}`);
const SHOTS = process.env.E2E_SHOTS || "";
const ONLY = process.env.E2E_ONLY || "";
const PYTHON = process.env.PYTHON || "python";
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
  } catch { /* already gone */ }
}

async function connect(wsUrl) {
  const ws = new WebSocket(wsUrl);
  await new Promise((res, rej) => { ws.onopen = res; ws.onerror = rej; });
  let id = 0;
  const pending = new Map();
  const logs = [];
  const consoleErrors = [];
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
    } else if (msg.method === "Runtime.consoleAPICalled" && msg.params.type === "error") {
      consoleErrors.push(msg.params.args.map((a) => a.value ?? a.description ?? "").join(" "));
    } else if (msg.method === "Page.javascriptDialogOpening") {
      dialogs.push(msg.params);
      send("Page.handleJavaScriptDialog", { accept: false }).catch(() => {});
    }
  };
  send = (method, params = {}) => new Promise((res, rej) => {
    const i = ++id;
    pending.set(i, { res, rej });
    ws.send(JSON.stringify({ id: i, method, params }));
  });
  return { send, logs, consoleErrors, dialogs, close: () => ws.close() };
}

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
    async move(x, y, modifiers = 0) {
      await c.send("Input.dispatchMouseEvent", { type: "mouseMoved", x, y, modifiers });
    },
    async mouse(x, y, clicks = 1, modifiers = 0, button = "left") {
      await c.send("Input.dispatchMouseEvent", { type: "mouseMoved", x, y });
      for (let n = 1; n <= clicks; n++) {
        for (const type of ["mousePressed", "mouseReleased"]) await c.send("Input.dispatchMouseEvent", { type, x, y, button, clickCount: n, modifiers });
      }
      await sleep(60);
    },
    /** Real drag: press, move in steps, release. */
    async drag(x0, y0, x1, y1, modifiers = 0, steps = 12) {
      await c.send("Input.dispatchMouseEvent", { type: "mouseMoved", x: x0, y: y0 });
      await c.send("Input.dispatchMouseEvent", { type: "mousePressed", x: x0, y: y0, button: "left", clickCount: 1, modifiers });
      for (let i = 1; i <= steps; i++) {
        await c.send("Input.dispatchMouseEvent", { type: "mouseMoved", x: x0 + ((x1 - x0) * i) / steps, y: y0 + ((y1 - y0) * i) / steps, button: "left", buttons: 1, modifiers });
        await sleep(10);
      }
      await c.send("Input.dispatchMouseEvent", { type: "mouseReleased", x: x1, y: y1, button: "left", clickCount: 1, modifiers });
      await sleep(80);
    },
    async rect(selector) {
      return p.eval(`(()=>{const e=document.querySelector(${JSON.stringify(selector)}); if(!e) return null; const b=e.getBoundingClientRect(); return {x:b.x,y:b.y,w:b.width,h:b.height,cx:b.x+b.width/2,cy:b.y+b.height/2}})()`);
    },
    async click(selector, modifiers = 0) {
      const r = await p.eval(`(()=>{const e=document.querySelector(${JSON.stringify(selector)}); if(!e) return null; e.scrollIntoView({block:'nearest'}); const b=e.getBoundingClientRect(); return b.width ? {x:b.x+b.width/2,y:b.y+b.height/2} : null})()`);
      if (!r) throw new Error(`not clickable: ${selector}`);
      await p.mouse(r.x, r.y, 1, modifiers);
    },
    async clickText(selector, text) {
      const r = await p.eval(`(()=>{const e=[...document.querySelectorAll(${JSON.stringify(selector)})].find(n=>n.textContent.trim()===${JSON.stringify(text)} && n.getBoundingClientRect().width); if(!e) return null; const b=e.getBoundingClientRect(); return {x:b.x+b.width/2,y:b.y+b.height/2}})()`);
      if (!r) throw new Error(`no ${selector} "${text}"`);
      await p.mouse(r.x, r.y);
    },
    async key(k, modifiers = 0) {
      const map = { Enter: [13, "\r"], Tab: [9, ""], Escape: [27, ""], Delete: [46, ""], Backspace: [8, ""], End: [35, ""], Home: [36, ""], F5: [116, ""], ArrowLeft: [37, ""], ArrowUp: [38, ""], ArrowRight: [39, ""], ArrowDown: [40, ""], PageDown: [34, ""], PageUp: [33, ""], " ": [32, " "] };
      const [vk, text] = map[k] || [k.toUpperCase().charCodeAt(0), k];
      const withText = !!text && !(modifiers & CTRL);
      const code = k.length === 1 && /[a-z]/i.test(k) ? `Key${k.toUpperCase()}` : k === " " ? "Space" : k;
      // editing keys need their editor command when sent through CDP
      const commands = { Backspace: ["deleteBackward"], Delete: ["deleteForward"], End: [modifiers & SHIFT ? "moveToEndOfLineAndModifySelection" : "moveToEndOfLine"], Home: [modifiers & SHIFT ? "moveToBeginningOfLineAndModifySelection" : "moveToBeginningOfLine"] }[k];
      await c.send("Input.dispatchKeyEvent", { type: withText ? "keyDown" : "rawKeyDown", key: k, code, windowsVirtualKeyCode: vk, text: withText ? text : undefined, modifiers, ...(commands && !(modifiers & CTRL) ? { commands } : {}) });
      await c.send("Input.dispatchKeyEvent", { type: "keyUp", key: k, code, windowsVirtualKeyCode: vk, modifiers });
      await sleep(40);
    },
    async type(s) {
      for (const ch of s) {
        if (ch === "\n") { await p.key("Enter"); continue; }
        await c.send("Input.dispatchKeyEvent", { type: "keyDown", key: ch, text: ch, unmodifiedText: ch });
        await c.send("Input.dispatchKeyEvent", { type: "keyUp", key: ch });
      }
      await sleep(40);
    },
    async waitFor(expr, ms = 15000) {
      const t0 = Date.now();
      while (Date.now() - t0 < ms) {
        try { const v = await p.eval(expr); if (v) return v; } catch { /* still loading */ }
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
    async shot(name) {
      if (!SHOTS) return;
      fs.mkdirSync(SHOTS, { recursive: true });
      const r = await c.send("Page.captureScreenshot", { format: "png" });
      fs.writeFileSync(path.join(SHOTS, `slides-${name}.png`), Buffer.from(r.data, "base64"));
    },
    async size(w, h) {
      await c.send("Emulation.setDeviceMetricsOverride", { width: w, height: h, deviceScaleFactor: 1, mobile: false });
      await sleep(250);
    },
    deck: () => p.eval("JSON.stringify(__officeViewer.doc.model, (k, v) => typeof v === 'string' && v.startsWith('data:') ? v.slice(0, 40) : v)").then(JSON.parse),
    view: (expr) => p.eval(`(()=>{const v=window.__slidesView; return ${expr}})()`),
    state: () => p.eval(`JSON.stringify({name:__officeViewer.doc?.name, kind:__officeViewer.doc?.kind, dirty:__officeViewer.dirty, status:document.getElementById('status-msg').textContent, stats:document.getElementById('doc-stats')?.textContent, modals:[...document.querySelectorAll('.modal')].filter(m=>!m.hidden).map(m=>m.id)})`).then(JSON.parse),
    /** Screen rectangle of item i on the canvas. */
    item: (i) => p.rect(`.sv-host .sl-item[data-i="${i}"]`),
  };
  return p;
}

const results = [];
function check(name, ok, detail = "") {
  results.push({ name, ok: !!ok });
  console.log(`${ok ? "PASS" : "FAIL"}  ${name}${ok ? "" : `\n      ${typeof detail === "string" ? detail : JSON.stringify(detail)}`}`);
}

function pythonValidate(bytes, label) {
  const tmp = path.join(PROFILE, `${label}.pptx`);
  fs.writeFileSync(tmp, bytes);
  const r = spawnSync(PYTHON, [path.join(FIX, "validate_pptx.py"), tmp], { encoding: "utf8" });
  if (r.error || (r.status !== 0 && !r.stdout)) return { skipped: true, why: String(r.error || r.stderr).slice(0, 300) };
  try { return JSON.parse(r.stdout); } catch { return { ok: false, problems: [r.stdout.slice(0, 500), r.stderr.slice(0, 500)] }; }
}

const runText = (paras) => (paras || []).map((p) => (p.runs || []).map((r) => r.t).join("")).join("\n");

// ---------------------------------------------------------------- tests

const TESTS = [
  ["new presentation: layout, placeholders typed in place, slides added", async (open) => {
    const p = await open("?new=slides");
    await p.waitFor("!!document.querySelector('#slides-view .sv-host .sl-slide')");
    await p.shot("01-new");
    const d0 = await p.deck();
    check("new deck has a title slide", d0.slides.length === 1 && d0.slides[0].layout === "title" && d0.slides[0].items.length === 2, d0.slides[0]);
    check("prompt text is shown in empty placeholders", await p.eval("document.querySelector('.sv-host p[data-prompt]')?.getAttribute('data-prompt') === 'Click to add title'"));
    // click the title placeholder, then type: typing replaces the empty title
    const t = await p.item(0);
    await p.mouse(t.cx, t.cy);
    await p.type("Hello World");
    await p.key("Escape");
    let d = await p.deck();
    check("typing into the selected title placeholder sets its text", runText(d.slides[0].items[0].paras) === "Hello World", d.slides[0].items[0].paras);
    check("status bar shows slide stats", /Slide 1 of 1/.test((await p.state()).stats), (await p.state()).stats);
    // subtitle: click once to select, click again to place the caret, type
    const st = await p.item(1);
    await p.mouse(st.cx, st.cy);
    await p.mouse(st.cx, st.cy);
    await p.type("By me");
    await p.key("Escape");
    d = await p.deck();
    check("second click edits the subtitle in place", runText(d.slides[0].items[1].paras) === "By me", d.slides[0].items[1].paras);
    check("document marked dirty", (await p.state()).dirty === true);
    // New slide (default layout: title and content)
    await p.click('.sv-ribbon [data-cmd="newSlide"]');
    d = await p.deck();
    check("New slide adds a Title and Content slide after the current one", d.slides.length === 2 && d.slides[1].layout === "content", d.slides.map((s) => s.layout));
    check("thumbnail strip shows both slides", (await p.eval("document.querySelectorAll('.sv-thumb').length")) === 2);
    const title2 = await p.item(0);
    await p.mouse(title2.cx, title2.cy);
    await p.type("Agenda");
    const body = await p.item(1);
    await p.mouse(body.cx, body.cy); // commits the title, selects the body
    await p.mouse(body.x + 60, body.y + 30); // second click: caret
    await p.type("First point\nSecond point");
    await p.key("Tab");
    await p.type(" (detail)");
    await p.key("Escape");
    d = await p.deck();
    const bp = d.slides[1].items[1].paras;
    check("body bullets typed with Enter", bp.length === 2 && bp.every((x) => x.bullet === "ul") && runText([bp[0]]) === "First point", bp);
    check("Tab while typing makes a sub-bullet", bp[1].level === 1 && runText([bp[1]]) === "Second point (detail)", bp[1]);
    // layout picker for a new slide
    await p.click('.sv-ribbon [data-pop="layoutNew"]');
    await p.click('.sv-pop [data-layout="two"]');
    d = await p.deck();
    check("layout picker adds a Two Content slide", d.slides.length === 3 && d.slides[2].layout === "two" && d.slides[2].items.length === 3, d.slides[2] && d.slides[2].layout);
    await p.shot("02-three-slides");
    // undo the last slide
    await p.click(".sv-canvas");
    await p.key("z", CTRL);
    d = await p.deck();
    check("Ctrl+Z removes the added slide", d.slides.length === 2, d.slides.length);
    await p.key("y", CTRL);
    d = await p.deck();
    check("Ctrl+Y adds it back", d.slides.length === 3, d.slides.length);
  }],

  ["objects: text box, shape, picture, table; mouse move/resize; clipboard; undo", async (open) => {
    const p = await open("?new=slides");
    await p.waitFor("!!document.querySelector('.sv-host .sl-slide')");
    await p.click('.sv-ribbon [data-pop="layoutNew"]');
    await p.click('.sv-pop [data-layout="blank"]');
    let d = await p.deck();
    check("blank slide added", d.slides.length === 2 && d.slides[1].items.length === 0);
    // text box: inserted in editing mode, typing goes in
    await p.click('.sv-ribbon [data-cmd="textbox"]');
    await p.type("Note box");
    await p.key("Escape");
    d = await p.deck();
    let items = d.slides[1].items;
    check("text box inserted with typed text", items.length === 1 && items[0].type === "text" && runText(items[0].paras) === "Note box", items);
    // bold via the toolbar on the selected box
    await p.click('.sv-ribbon [data-cmd="bold"]');
    d = await p.deck();
    check("Bold button formats the selected text box", d.slides[1].items[0].paras[0].runs.every((r) => r.b), d.slides[1].items[0].paras);
    // shape
    await p.click('.sv-ribbon [data-pop="shape"]');
    await p.click('.sv-pop [data-shape="ellipse"]');
    d = await p.deck();
    items = d.slides[1].items;
    const sh0 = items[1];
    check("shape inserted with the theme accent fill", sh0 && sh0.type === "shape" && sh0.shape === "ellipse" && sh0.fill === d.theme.accent, sh0);
    // drag it by the mouse
    const k = await p.eval("window.__slidesView.scale");
    let r = await p.item(1);
    await p.drag(r.cx, r.cy, r.cx + 100, r.cy + 50);
    d = await p.deck();
    let sh = d.slides[1].items[1];
    check("dragging moves the shape", Math.abs(sh.x - (sh0.x + 100 / k)) < 2 && Math.abs(sh.y - (sh0.y + 50 / k)) < 2, { before: sh0, after: sh, k });
    // resize with the bottom-right handle
    const hnd = await p.rect(".sv-selbox .sv-h-se");
    await p.drag(hnd.cx, hnd.cy, hnd.cx + 60, hnd.cy + 40);
    d = await p.deck();
    const sh2 = d.slides[1].items[1];
    check("the corner handle resizes it", Math.abs(sh2.w - (sh.w + 60 / k)) < 2 && Math.abs(sh2.h - (sh.h + 40 / k)) < 2 && Math.abs(sh2.x - sh.x) < 1, { sh, sh2 });
    // snapping: drag so the centre lands near the slide centre
    r = await p.item(1);
    const canvas = await p.rect(".sv-canvas");
    const selBefore = await p.view("JSON.stringify(v.sel)");
    await p.drag(r.cx, r.cy, canvas.cx + 3, canvas.cy - 2);
    d = await p.deck();
    const sh3 = d.slides[1].items[1];
    check("dragging snaps to the slide centre", Math.abs(sh3.x + sh3.w / 2 - d.size.w / 2) < 0.6 && Math.abs(sh3.y + sh3.h / 2 - d.size.h / 2) < 0.6, { sh3, r, canvas, selBefore, k });
    await p.shot("03-objects");
    // arrows nudge
    await p.key("ArrowRight");
    await p.key("ArrowRight", SHIFT);
    d = await p.deck();
    check("arrow keys nudge 1pt, Shift+arrow 10pt", Math.abs(d.slides[1].items[1].x - (sh3.x + 11)) < 0.01, d.slides[1].items[1].x);
    // undo: nudges (coalesced), snap drag, resize, move
    await p.key("z", CTRL);
    d = await p.deck();
    check("Ctrl+Z undoes the nudges", Math.abs(d.slides[1].items[1].x - sh3.x) < 0.01, d.slides[1].items[1]);
    await p.key("z", CTRL); await p.key("z", CTRL);
    d = await p.deck();
    check("Ctrl+Z undoes resize and snapping", Math.abs(d.slides[1].items[1].w - sh.w) < 0.01 && Math.abs(d.slides[1].items[1].x - sh.x) < 0.01, d.slides[1].items[1]);
    // copy / paste / duplicate / delete
    await p.key("c", CTRL);
    await p.key("v", CTRL);
    await sleep(200);
    d = await p.deck();
    items = d.slides[1].items;
    check("Ctrl+C, Ctrl+V pastes an offset copy", items.length === 3 && items[2].shape === "ellipse" && Math.abs(items[2].x - items[1].x - 12) < 0.01, items.map((i) => [i.type, i.x]));
    await p.key("d", CTRL);
    d = await p.deck();
    check("Ctrl+D duplicates", d.slides[1].items.length === 4);
    await p.key("Delete");
    d = await p.deck();
    check("Delete removes the selected object", d.slides[1].items.length === 3);
    // paste onto another slide
    await p.click('.sv-thumb[data-s="0"]');
    await p.click(".sv-canvas");
    await p.key("v", CTRL);
    await sleep(200);
    d = await p.deck();
    check("paste works on another slide", d.slides[0].items.some((i) => i.type === "shape" && i.shape === "ellipse"), d.slides[0].items.map((i) => i.type));
    await p.click('.sv-thumb[data-s="1"]');
    // picture from a file
    const png = path.join(PROFILE, "pic.png");
    fs.writeFileSync(png, Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAIAAAB7QOjdAAAAD0lEQVR4nGP4z8DAwPAfAAcAAf9+CLHQAAAAAElFTkSuQmCC", "base64"));
    await p.setFiles([png], "#slides-view .sv-file");
    await p.waitFor("__officeViewer.doc.model.slides[1].items.some(i => i.type === 'image')");
    d = await p.deck();
    const img = d.slides[1].items.find((i) => i.type === "image");
    check("picture inserted from a file as a data URL", img && /^data:image\/png;base64,/.test(img.src) && Math.abs(img.w / img.h - 2) < 0.05, img);
    // fill color on the picture is not offered; on the shape it is
    r = await p.item(2); // the pasted copy (on top)
    await p.mouse(r.cx, r.cy);
    await p.click('.sv-ribbon [data-pop="fill"]');
    await p.click('.sv-pop .sv-sw[data-color="#FFC000"]');
    d = await p.deck();
    check("fill color picker recolors the shape", d.slides[1].items[2].fill === "#FFC000", d.slides[1].items[2]);
    // table picker
    await p.click('.sv-ribbon [data-pop="table"]');
    await p.click('.sv-pop .sv-cell[data-r="2"][data-c="3"]');
    await p.type("Name");
    await p.key("Tab");
    await p.type("Score");
    await p.key("Escape");
    d = await p.deck();
    const tbl = d.slides[1].items.find((i) => i.type === "table");
    check("table 2x3 inserted, typing and Tab fill its cells", tbl && tbl.rows.length === 2 && tbl.rows[0].length === 3 && runText(tbl.rows[0][0].paras) === "Name" && runText(tbl.rows[0][1].paras) === "Score", tbl && tbl.rows);
    await p.shot("04-objects-table");
    // rotation handle
    r = await p.item(2);
    await p.mouse(r.cx, r.cy);
    const rot = await p.rect(".sv-selbox .sv-rot");
    await p.drag(rot.cx, rot.cy, r.cx + 200, r.cy, SHIFT);
    d = await p.deck();
    check("rotation handle rotates (Shift snaps to 15°)", d.slides[1].items[2].rot === 90, d.slides[1].items[2].rot);
  }],

  ["text formatting while typing; size/color/alignment on selected boxes; layout switch keeps text", async (open) => {
    const p = await open("?new=slides");
    await p.waitFor("!!document.querySelector('.sv-host .sl-slide')");
    await p.click('.sv-ribbon [data-cmd="newSlide"]');
    let r = await p.item(1);
    await p.mouse(r.cx, r.cy);
    await p.mouse(r.x + 80, r.y + 30);
    await p.type("Hello world");
    for (let i = 0; i < 5; i++) await p.key("ArrowLeft", SHIFT);
    await p.click('.sv-ribbon [data-cmd="bold"]');
    await p.click('.sv-ribbon [data-pop="textColor"]');
    await p.click('.sv-pop .sv-sw[data-color="#C00000"]');
    await p.key("End");
    await p.type("!");
    await p.key("Escape");
    let d = await p.deck();
    let runs = d.slides[1].items[1].paras[0].runs;
    check("selected word made bold and red while typing", runs.length >= 2 && runs[0].t === "Hello " && !runs[0].b && runs.some((x) => x.t.startsWith("world") && x.b && x.color === "#C00000"), runs);
    check("typing after formatting continues the paragraph", runs.map((x) => x.t).join("") === "Hello world!", runs);
    // size box on the selected (not edited) box applies to all its text
    await p.eval("(()=>{const s=document.querySelector('.sv-ribbon .sv-size'); s.focus(); s.select();})()");
    await p.type("32");
    await p.key("Enter");
    d = await p.deck();
    runs = d.slides[1].items[1].paras[0].runs;
    check("font size box sets the size of the selected box's text", runs.every((x) => x.size === 32), runs);
    await p.click('.sv-ribbon [data-cmd="alignCenter"]');
    await p.click('.sv-ribbon [data-cmd="numbers"]');
    d = await p.deck();
    const para = d.slides[1].items[1].paras[0];
    check("alignment and numbering buttons apply to the selected box", para.align === "center" && para.bullet === "ol", para);
    await p.eval("(()=>{const s=document.querySelector('.sv-ribbon .sv-font'); s.value='Georgia'; s.dispatchEvent(new Event('change',{bubbles:true}))})()");
    d = await p.deck();
    check("font picker sets the font", d.slides[1].items[1].paras[0].runs.every((x) => x.font === "Georgia"));
    // title, then switch layout to Two Content: texts kept
    r = await p.item(0);
    await p.mouse(r.cx, r.cy);
    await p.type("My title");
    await p.key("Escape");
    await p.click('.sv-ribbon [data-pop="layout"]');
    await p.click('.sv-pop [data-layout="two"]');
    d = await p.deck();
    const s = d.slides[1];
    check("layout switch keeps title and body text", s.layout === "two" && runText(s.items[0].paras) === "My title" && runText(s.items[1].paras) === "Hello world!" && s.items[2].ph === "body", s.items.map((i) => [i.ph, runText(i.paras)]));
    await p.key("z", CTRL);
    d = await p.deck();
    check("Ctrl+Z restores the previous layout", d.slides[1].layout === "content");
    // Enter on a selected box edits it; Esc leaves; Delete in text does not delete the box
    await p.key("Escape"); // nothing selected
    r = await p.item(0);
    await p.mouse(r.cx, r.cy);
    await p.key("Enter");
    await p.key("End");
    await p.key("Backspace");
    await p.key("Escape");
    d = await p.deck();
    check("Enter edits the selected box; Backspace edits text, not the slide", runText(d.slides[1].items[0].paras) === "My titl" && d.slides[1].items.length === 2, d.slides[1].items.map((i) => runText(i.paras)));
    await p.shot("05b-formatting");
    // a text box grows with its text
    await p.click('.sv-ribbon [data-cmd="textbox"]');
    await p.type("one\ntwo\nthree\nfour");
    const live = await p.eval("parseFloat(document.querySelector('.sv-host .sv-editing').style.height)");
    await p.key("Escape");
    d = await p.deck();
    const box = d.slides[1].items[d.slides[1].items.length - 1];
    check("text box grows while typing and keeps the new height", live > 90 && box.h > 90 && box.paras.length === 4, { live, h: box.h });
  }],

  ["filmstrip: drag to reorder, hide, duplicate, delete", async (open) => {
    const p = await open("?new=slides");
    await p.waitFor("!!document.querySelector('.sv-host .sl-slide')");
    for (const t of ["One", "Two", "Three"]) {
      if (t !== "One") await p.click('.sv-ribbon [data-cmd="newSlide"]');
      const r = await p.item(0);
      await p.mouse(r.cx, r.cy);
      await p.type(t);
      await p.key("Escape");
    }
    const titles = async () => (await p.deck()).slides.map((s) => runText(s.items[0].paras));
    check("three slides typed", JSON.stringify(await titles()) === '["One","Two","Three"]', await titles());
    const a = await p.rect('.sv-thumb[data-s="0"]');
    const c = await p.rect('.sv-thumb[data-s="2"]');
    await p.drag(a.cx, a.cy, c.cx, c.y + c.h - 4);
    check("dragging the first thumbnail below the third reorders", JSON.stringify(await titles()) === '["Two","Three","One"]', await titles());
    const t1 = await p.rect('.sv-thumb[data-s="1"]');
    await p.mouse(t1.cx, t1.cy, 1, 0, "right");
    await p.clickText(".sv-pop .sv-menu button", "Hide slide");
    let d = await p.deck();
    check("context menu hides a slide", d.slides[1].hidden === true && (await p.eval("document.querySelector('.sv-thumb[data-s=\"1\"]').classList.contains('hidden-slide')")));
    await p.click('.sv-thumb[data-s="0"]');
    await p.key("d", CTRL);
    d = await p.deck();
    check("Ctrl+D in the filmstrip duplicates the slide", d.slides.length === 4 && runText(d.slides[1].items[0].paras) === "Two");
    await p.key("Delete");
    d = await p.deck();
    check("Delete in the filmstrip deletes the slide", d.slides.length === 3);
    await p.shot("05-filmstrip");
  }],

  ["speaker notes and the slide show (keys, hidden slides, notes, Esc)", async (open) => {
    const p = await open("?new=slides");
    await p.waitFor("!!document.querySelector('.sv-host .sl-slide')");
    let r = await p.item(0);
    await p.mouse(r.cx, r.cy);
    await p.type("Show time");
    await p.key("Escape");
    await p.click(".sv-notes");
    await p.type("Remember to smile");
    let d = await p.deck();
    check("notes typed into the notes pane", d.slides[0].notes === "Remember to smile", d.slides[0].notes);
    await p.click('.sv-ribbon [data-cmd="newSlide"]');
    r = await p.item(0); await p.mouse(r.cx, r.cy); await p.type("Hidden one"); await p.key("Escape");
    await p.click('.sv-ribbon [data-pop="present"]');
    await p.clickText(".sv-pop .sv-menu button", "Hide this slide in the slide show");
    await p.click('.sv-ribbon [data-cmd="newSlide"]');
    r = await p.item(0); await p.mouse(r.cx, r.cy); await p.type("Last"); await p.key("Escape");
    await p.click('.sv-ribbon [data-pop="transition"]');
    await p.click('.sv-pop [data-tr="fade"]');
    d = await p.deck();
    check("transition set on the slide", d.slides[2].transition === "fade" && d.slides[1].hidden === true);
    // notes keys must not start the show
    await p.click(".sv-notes");
    await p.key("F5");
    await sleep(200);
    check("F5 while typing notes still presents (it is not a typing key)", await p.eval("!!document.querySelector('.sv-show')"));
    await p.key("Escape");
    await sleep(200);
    check("Esc ends the show", !(await p.eval("!!document.querySelector('.sv-show')")));
    await p.click(".sv-canvas");
    await p.key("F5");
    await p.waitFor("!!document.querySelector('.sv-show .sv-show-slide')");
    const count = () => p.eval("document.querySelector('.sv-show-count').textContent");
    const shown = () => p.eval("document.querySelector('.sv-show-slide').textContent");
    check("F5 starts from the first slide; hidden slides are not counted", (await count()) === "1 / 2" && (await shown()).includes("Show time"), await count());
    await p.shot("06-show");
    await p.key("n");
    check("N shows the speaker notes", await p.eval("!document.querySelector('.sv-show-notes').hidden && document.querySelector('.sv-show-notes').textContent === 'Remember to smile'"));
    await p.key("ArrowRight");
    check("Right arrow skips the hidden slide", (await count()) === "2 / 2" && (await shown()).includes("Last"), await shown());
    await p.key("b");
    check("B blacks out the screen", await p.eval("!document.querySelector('.sv-show-black').hidden"));
    await p.key("b");
    await p.key(" ");
    check("past the last slide: end screen", await p.eval("!document.querySelector('.sv-show-end').hidden"));
    await p.key("Escape");
    await sleep(150);
    check("Esc closes the show and returns to the editor", !(await p.eval("!!document.querySelector('.sv-show')")) && (await p.view("v.current")) === 2);
    await p.click(".sv-canvas");
    await p.key("F5", SHIFT);
    await p.waitFor("!!document.querySelector('.sv-show .sv-show-slide')");
    check("Shift+F5 starts from the current slide", (await count()) === "2 / 2");
    const sz = await p.rect(".sv-show");
    await p.mouse(sz.cx, sz.cy);
    await p.mouse(sz.cx, sz.cy);
    check("clicking advances and finally exits", !(await p.eval("!!document.querySelector('.sv-show')")));
  }],

  ["save as .pptx: python-pptx reads it back; reopening keeps the content", async (open) => {
    const p = await open("?new=slides");
    await p.waitFor("!!document.querySelector('.sv-host .sl-slide')");
    let r = await p.item(0);
    await p.mouse(r.cx, r.cy); await p.type("Saved deck"); await p.key("Escape");
    await p.click(".sv-notes"); await p.type("Note one");
    await p.click('.sv-ribbon [data-cmd="newSlide"]');
    r = await p.item(0); await p.mouse(r.cx, r.cy); await p.type("Points"); await p.key("Escape");
    r = await p.item(1); await p.mouse(r.cx, r.cy); await p.type("Alpha\nBeta"); await p.key("Escape");
    await p.click('.sv-ribbon [data-pop="shape"]'); await p.click('.sv-pop [data-shape="rightArrow"]');
    await p.click('.sv-ribbon [data-pop="table"]'); await p.click('.sv-pop .sv-cell[data-r="2"][data-c="2"]'); await p.type("T1"); await p.key("Escape");
    const png = path.join(PROFILE, "pic2.png");
    fs.writeFileSync(png, Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAIAAAB7QOjdAAAAD0lEQVR4nGP4z8DAwPAfAAcAAf9+CLHQAAAAAElFTkSuQmCC", "base64"));
    await p.setFiles([png], "#slides-view .sv-file");
    await p.waitFor("__officeViewer.doc.model.slides[1].items.some(i => i.type === 'image')");
    await p.click('.sv-ribbon [data-pop="theme"]'); await p.click('.sv-pop [data-theme="dark"]');
    const before = await p.deck();
    check("dark theme applied", before.theme.bg === "#1E2129", before.theme);
    await p.shot("07-before-save");
    await p.key("s", CTRL);
    await p.waitFor("!document.getElementById('saveas-modal').hidden");
    const name = await p.eval("document.getElementById('sa-name').value");
    check("Save As suggests a .pptx name", /\.pptx$/.test(name), name);
    await p.click("#sa-ok");
    const [w] = await p.written();
    check("saved a .pptx (zip)", w && w.bytes[0] === 0x50 && w.bytes[1] === 0x4b && /\.pptx$/.test(w.name), w && w.name);
    const v = pythonValidate(w.bytes, "e2e-saved");
    if (v.skipped) check("python-pptx validation (skipped: no python-pptx)", true);
    else {
      check("package is consistent (content types, relationships, ids)", v.ok, v.problems);
      const s1 = v.slides && v.slides[1];
      check("python-pptx reads slides, text, table, picture and notes", v.slides.length === 2 && v.slides[0].notes === "Note one"
        && s1.shapes.some((s) => s.text === "Alpha\nBeta") && s1.shapes.some((s) => s.table && s.table[0][0] === "T1") && s1.shapes.some((s) => s.image === "image/png"), v.slides);
    }
    check("document is clean after saving", (await p.state()).dirty === false);
    // reopen the saved bytes in the editor
    await p.eval(`__officeViewer.openBytes(${JSON.stringify(w.name)}, Uint8Array.from(atob(${JSON.stringify(w.bytes.toString("base64"))}), c => c.charCodeAt(0)))`);
    await p.waitFor("__officeViewer.doc.name === " + JSON.stringify(w.name) + " && !!document.querySelector('.sv-host .sl-slide')");
    const after = await p.deck();
    const texts = (dk) => dk.slides.map((s) => s.items.map((i) => i.type === "table" ? runText(i.rows[0][0].paras) : i.type === "image" ? "img" : i.type === "shape" ? i.shape : runText(i.paras)).join("|"));
    check("reopened deck keeps texts, objects and notes", JSON.stringify(texts(after)) === JSON.stringify(texts(before)) && after.slides[0].notes === "Note one", { before: texts(before), after: texts(after) });
    check("reopened deck keeps the theme colors", after.theme.bg === before.theme.bg && after.theme.accent === before.theme.accent, after.theme);
  }],

  ["open a real .pptx (python-pptx template) through the file input", async (open) => {
    const p = await open("");
    await p.setFiles([path.join(FIX, "basic.pptx")]);
    await p.waitFor("__officeViewer.doc?.kind === 'slides' && !!document.querySelector('.sv-host .sl-slide')");
    const d = await p.deck();
    check("7 slides read, 4:3 size", d.slides.length === 7 && d.size.w === 720 && d.size.h === 540, { n: d.slides.length, size: d.size });
    check("thumbnails for every slide", (await p.eval("document.querySelectorAll('.sv-thumb').length")) === 7);
    check("hidden slide marked in the filmstrip", await p.eval("document.querySelector('.sv-thumb[data-s=\"6\"]').classList.contains('hidden-slide')"));
    const st = await p.state();
    check("status bar shows the slide count", /Slide 1 of 7/.test(st.stats), st.stats);
    check("chart warning shown to the user", await p.eval("document.body.textContent.includes('Charts are shown as placeholders')"));
    await p.shot("08-open-fixture");
    await p.click('.sv-thumb[data-s="4"]');
    await sleep(200);
    await p.shot("09-fixture-shapes");
    await p.click('.sv-thumb[data-s="5"]');
    await sleep(200);
    await p.shot("10-fixture-table");
    check("table slide renders a table", await p.eval("document.querySelectorAll('.sv-host .sl-table td').length === 9"));
  }],

  ["open a .pptx with master decorations, SmartArt, picture background; edit and save it again", async (open) => {
    const p = await open("");
    await p.setFiles([path.join(FIX, "complex.pptx")]);
    await p.waitFor("__officeViewer.doc?.kind === 'slides' && !!document.querySelector('.sv-host .sl-slide')");
    let d = await p.deck();
    check("4 slides, 16:9", d.slides.length === 4 && d.size.w === 960);
    await p.click('.sv-thumb[data-s="1"]');
    await sleep(150);
    await p.shot("15-complex-objects");
    await p.click('.sv-thumb[data-s="3"]');
    await sleep(150);
    await p.shot("16-complex-picture-bg");
    // move the text on the picture background, then save in place (Save -> Save As for a picked file)
    const r = await p.item(2);
    await p.drag(r.cx, r.cy, r.cx + 40, r.cy + 40);
    d = await p.deck();
    check("text box on the picture background moved", d.slides[3].items[2].x > 72 + 20, d.slides[3].items[2]);
    await p.key("s", CTRL);
    await p.waitFor("window.__written.length > 0 || !document.getElementById('saveas-modal').hidden");
    if (await p.eval("!document.getElementById('saveas-modal').hidden")) await p.click("#sa-ok");
    const [w] = await p.written();
    const v = pythonValidate(w.bytes, "e2e-complex");
    if (v.skipped) check("python-pptx validation (skipped)", true);
    else check("re-saved complex deck opens in python-pptx with 4 slides", v.ok && v.slides.length === 4, v.problems || v);
  }],

  ["narrow window, theme gallery and print layout", async (open) => {
    const p = await open("?new=slides&template=modern", 1000, 700);
    await p.waitFor("!!document.querySelector('.sv-host .sl-slide')");
    const d = await p.deck();
    check("template from the URL applied", d.theme.accent === "#1A56DB", d.theme);
    const r = await p.item(0);
    await p.mouse(r.cx, r.cy); await p.type("Narrow view"); await p.key("Escape");
    await p.shot("11-narrow");
    const ov = await p.eval("(()=>{const r=document.querySelector('.sv-ribbon').getBoundingClientRect(); return r.right <= innerWidth + 1})()");
    check("toolbar fits in a 1000px window", ov);
    await p.click('.sv-ribbon [data-pop="theme"]');
    await sleep(100);
    await p.shot("12-theme-gallery");
    await p.key("Escape");
    await p.eval("window.__slidesView.beforePrint()");
    check("print layout: one page per visible slide", (await p.eval("document.querySelectorAll('.sv-print-page').length")) === 1);
    await p.send("Emulation.setEmulatedMedia", { media: "print" });
    await sleep(150);
    await p.shot("13-print");
    const pdf = await p.send("Page.printToPDF", { landscape: true });
    const pdfText = Buffer.from(pdf.data, "base64").toString("latin1");
    const pages = (pdfText.match(/\/Type\s*\/Page[^s]/g) || []).length;
    check("print to PDF gives one page", pages === 1, pages);
    await p.send("Emulation.setEmulatedMedia", { media: "" });
  }],
];

// ---------------------------------------------------------------- main

async function main() {
  // a fresh profile every run: an unpacked extension in a reused profile may not
  // start its service worker, which is how the extension id is found
  try { fs.rmSync(PROFILE, { recursive: true, force: true }); } catch { /* locked */ }
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
    let extId = null;
    for (let i = 0; i < 80 && !extId; i++) {
      const list = await (await fetch(`http://127.0.0.1:${PORT}/json/list`)).json();
      const sw = list.find((t) => /\/background\/service-worker\.js$/.test(t.url));
      if (sw) extId = new URL(sw.url).host;
      else await sleep(150);
    }
    if (!extId) throw new Error("the extension did not load");
    const viewer = `chrome-extension://${extId}/viewer/viewer.html`;

    const open = async (query, w = 1400, h = 900) => {
      const { targetId } = await browserConn.send("Target.createTarget", { url: "about:blank" });
      const list = await (await fetch(`http://127.0.0.1:${PORT}/json/list`)).json();
      const c = await connect(list.find((t) => t.id === targetId).webSocketDebuggerUrl);
      for (const d of ["Runtime", "Page", "DOM"]) await c.send(`${d}.enable`);
      await c.send("Emulation.setFocusEmulationEnabled", { enabled: true });
      await c.send("Emulation.setDeviceMetricsOverride", { width: w, height: h, deviceScaleFactor: 1, mobile: false });
      const p = pageApi(c);
      p.targetId = targetId;
      await c.send("Page.navigate", { url: viewer + query });
      await p.waitFor("!!window.__officeViewer && document.readyState === 'complete'");
      await sleep(400);
      await p.eval(SAVE_STUB);
      return p;
    };

    {
      const p = await open("");
      await p.eval("chrome.storage.local.set({settings: {backend_enabled: false, backend_url: 'http://127.0.0.1:47831'}})");
      await browserConn.send("Target.closeTarget", { targetId: p.targetId });
    }
    for (const [name, fn] of TESTS) {
      if (ONLY && !name.includes(ONLY)) continue;
      console.log(`\n# ${name}`);
      const pages = [];
      try {
        await fn(async (q, w, h) => { const p = await open(q, w, h); pages.push(p); return p; });
      } catch (e) {
        check(`${name}: ran to the end`, false, e.stack || String(e));
        for (const p of pages) await p.shot(`error-${results.length}`).catch(() => {});
      }
      for (const p of pages) {
        if (p.logs.length) check(`${name}: no page exceptions`, false, p.logs.join("\n"));
        const errs = p.consoleErrors.filter((e) => !/favicon|Failed to load resource/i.test(e));
        if (errs.length) check(`${name}: no console errors`, false, errs.join("\n"));
        p.close();
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

main().catch((e) => { console.error(e); process.exit(1); });
