/**
 * End-to-end tests of the Word editor features (find & replace, navigation
 * pane, word count, zoom, symbols, change case, format painter, templates,
 * page setup, headers/footers with page numbers, document properties) with
 * real keyboard / mouse input in a headless Edge/Chrome.
 *
 * NOT part of `npm test`. Run from office-chrome-extension/:
 *     node tests/e2e/word-features.mjs
 *
 * Environment overrides:
 *   BROWSER_PATH       browser executable (default: Edge, then Chrome)
 *   E2E_PORT           remote debugging port (default 9353)
 *   E2E_PROFILE        user-data-dir (default: <os tmpdir>/office-ext-wf-<pid>, deleted afterwards)
 *   E2E_BACKEND_URL    use this office-backend (default: backend OFF, built-in engines)
 *   E2E_SHOTS          directory for screenshots (default: none)
 */

import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { spawn, execFileSync } from "node:child_process";
import { fileURLToPath } from "node:url";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const EXT = path.resolve(HERE, "../..");
const PORT = Number(process.env.E2E_PORT || 9353);
const OWN_PROFILE = !process.env.E2E_PROFILE;
const PROFILE = process.env.E2E_PROFILE || path.join(os.tmpdir(), `office-ext-wf-${process.pid}`);
const BACKEND = process.env.E2E_BACKEND_URL || "";
const SHOTS = process.env.E2E_SHOTS || "";
const CTRL = 2, SHIFT = 8;

// JSZip (vendored UMD) to inspect the saved packages.
const JSZip = (() => {
  const code = fs.readFileSync(path.join(EXT, "vendor/jszip.min.js"), "utf8");
  const mod = { exports: {} };
  new Function("module", "exports", code)(mod, mod.exports);
  return mod.exports;
})();

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
    } else if (msg.method === "Runtime.consoleAPICalled" && msg.params.type === "error") {
      logs.push(`console.error: ${msg.params.args.map((a) => a.value ?? a.description).join(" ")}`);
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
  return { send, logs, dialogs, close: () => ws.close() };
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
    async mouse(x, y, clicks = 1, modifiers = 0) {
      await c.send("Input.dispatchMouseEvent", { type: "mouseMoved", x, y });
      for (let n = 1; n <= clicks; n++) {
        for (const type of ["mousePressed", "mouseReleased"]) await c.send("Input.dispatchMouseEvent", { type, x, y, button: "left", clickCount: n, modifiers });
      }
      await sleep(60);
    },
    async xy(selector) {
      const r = await p.eval(`(()=>{const e=document.querySelector(${JSON.stringify(selector)}); if(!e) return null; e.scrollIntoView({block:'nearest'}); const b=e.getBoundingClientRect(); return b.width ? {x:b.x+b.width/2,y:b.y+b.height/2} : null})()`);
      if (!r) throw new Error(`not clickable: ${selector}`);
      return r;
    },
    async click(selector, clicks = 1) { const r = await p.xy(selector); await p.mouse(r.x, r.y, clicks); },
    async clickText(selector, text) {
      const r = await p.eval(`(()=>{const e=[...document.querySelectorAll(${JSON.stringify(selector)})].find(n=>n.textContent.trim().startsWith(${JSON.stringify(text)}) && n.getBoundingClientRect().width); if(!e) return null; const b=e.getBoundingClientRect(); return {x:b.x+b.width/2,y:b.y+b.height/2}})()`);
      if (!r) throw new Error(`no ${selector} "${text}"`);
      await p.mouse(r.x, r.y);
    },
    /** Select `needle` (first occurrence under root) with a real mouse drag. */
    async dragSelect(needle, root = "#word-view") {
      const r = await p.eval(`(()=>{const w=document.createTreeWalker(document.querySelector(${JSON.stringify(root)}),NodeFilter.SHOW_TEXT); for(let n=w.nextNode(); n; n=w.nextNode()){const i=n.data.indexOf(${JSON.stringify(needle)}); if(i>=0){n.parentElement.scrollIntoView({block:'center'}); const g=document.createRange(); g.setStart(n,i); g.setEnd(n,i+1); const a=g.getBoundingClientRect(); g.setStart(n,i+${needle.length}-1); g.setEnd(n,i+${needle.length}); const b=g.getBoundingClientRect(); return {x1:a.left+1,y:a.top+a.height/2,x2:b.right-0.5}}} return null})()`);
      if (!r) throw new Error(`text not found: ${needle}`);
      // A press inside an existing selection would start a text drag: collapse it first.
      await p.mouse(r.x2, r.y);
      await sleep(400);
      await c.send("Input.dispatchMouseEvent", { type: "mouseMoved", x: r.x1, y: r.y });
      await c.send("Input.dispatchMouseEvent", { type: "mousePressed", x: r.x1, y: r.y, button: "left", clickCount: 1 });
      for (const f of [0.3, 0.7, 1]) await c.send("Input.dispatchMouseEvent", { type: "mouseMoved", x: r.x1 + (r.x2 - r.x1) * f, y: r.y, button: "left", buttons: 1 });
      await c.send("Input.dispatchMouseEvent", { type: "mouseReleased", x: r.x2, y: r.y, button: "left", clickCount: 1 });
      await sleep(80);
    },
    async clickAfterText(needle, root = "#word-view") {
      const r = await p.eval(`(()=>{const w=document.createTreeWalker(document.querySelector(${JSON.stringify(root)}),NodeFilter.SHOW_TEXT); for(let n=w.nextNode(); n; n=w.nextNode()){const i=n.data.indexOf(${JSON.stringify(needle)}); if(i>=0){n.parentElement.scrollIntoView({block:'center'}); const g=document.createRange(); g.setStart(n,i); g.setEnd(n,i+${needle.length}); const b=g.getBoundingClientRect(); return {x:b.right-1,y:b.y+b.height/2}}} return null})()`);
      if (!r) throw new Error(`text not found: ${needle}`);
      await p.mouse(r.x, r.y);
    },
    async key(k, modifiers = 0) {
      const map = {
        Enter: [13, "\r", "Enter"], Tab: [9, "", "Tab"], Escape: [27, "", "Escape"], Delete: [46, "", "Delete"], End: [35, "", "End"],
        Home: [36, "", "Home"], ArrowDown: [40, "", "ArrowDown"], ArrowUp: [38, "", "ArrowUp"], ArrowRight: [39, "", "ArrowRight"],
        "=": [187, "=", "Equal"], "-": [189, "-", "Minus"], "0": [48, "0", "Digit0"], F3: [114, "", "F3"],
      };
      const [vk, text, code] = map[k] || [k.toUpperCase().charCodeAt(0), k, `Key${k.toUpperCase()}`];
      const withText = !!text && !(modifiers & CTRL);
      await c.send("Input.dispatchKeyEvent", { type: withText ? "keyDown" : "rawKeyDown", key: k, code, windowsVirtualKeyCode: vk, text: withText ? text : undefined, modifiers });
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
    async written(ms = 20000) {
      await p.waitFor("window.__written.length > 0", ms);
      await sleep(100);
      const arr = await p.eval("JSON.stringify(window.__written.splice(0))");
      return JSON.parse(arr).map((w) => ({ name: w.name, bytes: Buffer.from(w.b64, "base64") }));
    },
    /** Ctrl+S on a new document -> Save As -> Save; returns the unzipped package. */
    async saveDocx() {
      await p.key("s", CTRL);
      await p.waitFor("!document.getElementById('saveas-modal').hidden");
      await p.choose("#sa-format", "docx");
      await p.click("#sa-ok");
      const [w] = await p.written();
      const zip = await JSZip.loadAsync(w.bytes);
      const files = {};
      for (const name of Object.keys(zip.files)) if (!zip.files[name].dir) files[name] = await zip.file(name).async("string");
      return { name: w.name, bytes: w.bytes, files };
    },
    async shot(name) {
      if (!SHOTS) return;
      const r = await c.send("Page.captureScreenshot", { format: "png" });
      fs.mkdirSync(SHOTS, { recursive: true });
      fs.writeFileSync(path.join(SHOTS, `word-features-${name}.png`), Buffer.from(r.data, "base64"));
    },
    state: () => p.eval(`JSON.stringify({name:__officeViewer.doc?.name, kind:__officeViewer.doc?.kind, dirty:__officeViewer.dirty, status:document.getElementById('status-msg').textContent, toast:document.getElementById('toast').hidden?null:document.getElementById('toast').textContent, modals:[...document.querySelectorAll('.modal')].filter(m=>!m.hidden).map(m=>m.id), active:document.activeElement?.id||document.activeElement?.className})`).then(JSON.parse),
    wordModel: () => p.eval("JSON.stringify(__officeViewer.currentWordModel())").then(JSON.parse),
    text: () => p.eval("document.getElementById('word-view').innerText"),
  };
  return p;
}

// ---------------------------------------------------------------- tests

const results = [];
function check(name, ok, detail = "") {
  results.push({ name, ok: !!ok });
  console.log(`${ok ? "PASS" : "FAIL"}  ${name}${ok ? "" : `\n      ${typeof detail === "string" ? detail : JSON.stringify(detail)}`}`);
}
const runsText = (b) => (b.runs || []).map((r) => r.t).join("");
const allText = (m) => m.blocks.map(runsText).join("\n");

async function openMenu(p, menuBtn, itemText) {
  await p.click(menuBtn);
  await sleep(80);
  const menu = await p.eval(`document.querySelector(${JSON.stringify(menuBtn)}).getAttribute('aria-controls')`);
  await p.clickText(`#${menu} button`, itemText);
  await sleep(120);
}

const TESTS = [
  ["Find & replace with real keystrokes", async (open) => {
    const p = await open("?new=word");
    await p.click("#word-view");
    await p.type("The cat sat on the mat.\nA cat and another Cat.\nConcatenate the catalog.");
    await p.clickAfterText("The cat");
    await p.key("f", CTRL);
    await p.waitFor("!document.getElementById('find-bar').hidden");
    check("Ctrl+F opens the find bar with the cursor in it", (await p.eval("document.activeElement.id")) === "find-input");
    await p.type("cat");
    await sleep(100);
    let count = await p.eval("document.getElementById('find-count').textContent");
    check("all matches counted (case-insensitive, inside words too)", /of 5$/.test(count), count);
    check("matches are highlighted with the CSS Custom Highlight API", await p.eval("CSS.highlights.has('ow-find') && CSS.highlights.get('ow-find').size === 5 && CSS.highlights.has('ow-find-current')"));
    await p.click("#find-word");
    await sleep(80);
    count = await p.eval("document.getElementById('find-count').textContent");
    check("whole words: 3 matches", /of 3$/.test(count), count);
    await p.click("#find-case");
    await sleep(80);
    count = await p.eval("document.getElementById('find-count').textContent");
    check("match case + whole words: 2 matches", /of 2$/.test(count), count);
    await p.click("#find-case");
    await p.click("#find-input");
    const c0 = await p.eval("document.getElementById('find-count').textContent");
    await p.key("Enter");
    const c1 = await p.eval("document.getElementById('find-count').textContent");
    await p.key("Enter", SHIFT);
    const c2 = await p.eval("document.getElementById('find-count').textContent");
    check("Enter / Shift+Enter move to the next / previous match", c0 !== c1 && c2 === c0 && /^\d of 3$/.test(c1), [c0, c1, c2]);
    await p.shot("find");
    await p.key("h", CTRL);
    await p.waitFor("!document.getElementById('find-replace-row').hidden");
    check("Ctrl+H shows the replace field", (await p.eval("document.activeElement.id")) === "replace-input");
    await p.type("dog");
    await p.key("Enter");
    count = await p.eval("document.getElementById('find-count').textContent");
    const t1 = await p.text();
    check("Replace (Enter) replaces one match and moves on", /of 2$/.test(count) && (t1.match(/\bdog\b/g) || []).length === 1, [count, t1]);
    await p.click("#replace-all");
    await sleep(150);
    const m = await p.wordModel();
    const t = allText(m);
    check("Replace all: every whole-word match replaced, other words untouched",
      (t.match(/\bdog\b/gi) || []).length === 3 && !/\bcat\b/i.test(t) && t.includes("Concatenate the catalog."), t);
    check("replacing marks the document as changed", (await p.state()).dirty);
    await p.key("z", CTRL);
    const tu = await p.text();
    check("Ctrl+Z undoes the replacement", /\bcat\b/i.test(tu), tu);
    await p.key("y", CTRL);
    await p.click("#find-input");
    await p.key("Escape");
    const s = await p.state();
    check("Esc closes the bar, clears highlights and returns to the text",
      (await p.eval("document.getElementById('find-bar').hidden && !CSS.highlights.has('ow-find')")) && s.active === "word-view", s);
  }],

  ["Ctrl+F is left alone outside Word documents", async (open) => {
    const p = await open("?new=excel");
    await sleep(300);
    await p.key("f", CTRL);
    check("no Word find bar over a spreadsheet", await p.eval("document.getElementById('find-bar').hidden"));
  }],

  ["Template, navigation pane, word count and zoom", async (open) => {
    const p = await open("?new=word&template=report");
    await p.waitFor("__officeViewer.doc?.kind === 'word'");
    const m = await p.wordModel();
    check("?new=word&template=report creates the report template", m.blocks.some((b) => b.type === "h1" && runsText(b) === "1. Introduction") && m.meta?.footer?.runs?.[0]?.t.includes("{PAGE}"), m.blocks.slice(0, 3));
    check("the template footer shows page-number fields", await p.eval("document.querySelectorAll('#word-footer .ow-field').length === 2"));
    await p.click("#word-view");
    await p.key("f", CTRL | SHIFT);
    await p.waitFor("!document.getElementById('nav-pane').hidden");
    const items = await p.eval("[...document.querySelectorAll('#nav-list .nav-item')].map(b=>b.textContent)");
    check("Ctrl+Shift+F opens the navigation pane with the headings", items.length >= 8 && items.includes("3. Recommendations"), items);
    await p.clickText("#nav-list .nav-item", "3. Recommendations");
    await sleep(250);
    const where = await p.eval("(()=>{const s=getSelection(); const h=s.anchorNode && (s.anchorNode.nodeType===1?s.anchorNode:s.anchorNode.parentElement).closest('h1,h2,h3'); const r=h&&h.getBoundingClientRect(); const st=document.getElementById('stage').getBoundingClientRect(); return {h:h&&h.textContent, top:r&&Math.round(r.top-st.top), focus:document.activeElement.id}})()");
    check("clicking a heading scrolls to it and puts the cursor there", where.h === "3. Recommendations" && where.top >= 0 && where.top < 700 && where.focus === "word-view", where);
    const cur = await p.eval("document.querySelector('#nav-list .nav-item.current')?.textContent");
    check("the current section is highlighted", cur === "3. Recommendations", cur);
    await p.eval("document.getElementById('stage').scrollTop = 0");
    await sleep(200);
    const cur2 = await p.eval("document.querySelector('#nav-list .nav-item.current')?.textContent");
    check("the highlight follows scrolling", cur2 && cur2 !== "3. Recommendations", cur2);
    await p.shot("nav");
    // live outline
    await p.clickAfterText("4. Conclusion");
    await p.key("Enter");
    await p.choose("#fb-style", "h2");
    await p.type("4.1 Next steps");
    await sleep(600);
    const items2 = await p.eval("[...document.querySelectorAll('#nav-list .nav-item')].map(b=>b.textContent)");
    check("the outline updates while typing", items2.includes("4.1 Next steps"), items2);

    // word count
    await p.dragSelect("Summarize the purpose");
    await p.click("#doc-stats");
    await p.waitFor("!document.getElementById('wc-modal').hidden");
    const wc = await p.eval("JSON.stringify([...document.querySelectorAll('#wc-body tr')].map(tr=>[...tr.children].map(td=>td.textContent)))").then(JSON.parse);
    const stats = await p.eval("JSON.stringify(__officeViewer.wordStats())").then(JSON.parse);
    const row = (label) => wc.find((r) => r[0].startsWith(label)) || [];
    check("word count dialog: document words", Number(row("Words")[2].replace(/\D/g, "")) === stats.words && stats.words > 100, { wc, stats });
    check("word count dialog: selection column (3 words)", row("Words")[1] === "3" && row("Characters (with")[1] === "21", wc);
    check("word count dialog: pages estimate >= 2 (title page + break)", Number(row("Pages")[2]) >= 2, wc);
    await p.shot("wordcount");
    await p.key("Escape");
    check("Esc closes the word count", await p.eval("document.getElementById('wc-modal').hidden"));

    // zoom
    await p.click("#word-view");
    await p.key("=", CTRL);
    let z = await p.eval("__officeViewer.zoom");
    check("Ctrl+= zooms in", z === 110, z);
    await p.key("-", CTRL); await p.key("-", CTRL);
    z = await p.eval("__officeViewer.zoom");
    check("Ctrl+- zooms out", z === 90, z);
    const xy = await p.xy("#word-view");
    await p.send("Input.dispatchMouseEvent", { type: "mouseWheel", x: xy.x, y: xy.y, deltaX: 0, deltaY: -120, modifiers: CTRL });
    await sleep(150);
    z = await p.eval("__officeViewer.zoom");
    check("Ctrl+wheel zooms", z === 100, z);
    await p.click("#zoom-in"); await p.click("#zoom-in");
    const zoomCss = await p.eval("getComputedStyle(document.getElementById('word-sheet')).zoom");
    check("status-bar + button zooms the page", (await p.eval("__officeViewer.zoom")) === 120 && Math.abs(Number(zoomCss) - 1.2) < 0.01 && (await p.eval("document.getElementById('zoom-label').textContent")) === "120%", zoomCss);
    await p.shot("zoom");
    await p.key("0", CTRL);
    check("Ctrl+0 resets to 100%", (await p.eval("__officeViewer.zoom")) === 100);
    await p.click("#zoom-in");
  }],

  ["zoom is remembered", async (open) => {
    const p = await open("?new=word");
    const z = await p.eval("__officeViewer.zoom");
    check("zoom level survives a reload (localStorage)", z === 110, z);
    await p.eval("document.getElementById('zoom-label').click()");
  }],

  ["Symbols, date, change case, format painter", async (open) => {
    const p = await open("?new=word");
    await p.click("#word-view");
    await p.type("Price ");
    await openMenu(p, "#fb-insert", "Symbol");
    await p.waitFor("!document.getElementById('sym-modal').hidden");
    await p.shot("symbols");
    await p.clickText("#sym-tabs .sym-tab", "Currency");
    await p.click('#sym-grid .sym-cell[data-ch="€"]');
    await p.click("#sym-close");
    await p.type("42");
    let t = await p.text();
    check("Insert › Symbol inserts at the cursor", t.includes("Price €42"), t);
    check("recently used symbols are remembered", (await p.eval("localStorage.getItem('officeEditor.word.recentSymbols')")).includes("€"));

    await p.key("Enter");
    await openMenu(p, "#fb-insert", "Date & time");
    await p.waitFor("!document.getElementById('dt-modal').hidden");
    const iso = await p.eval("(()=>{const d=new Date(); return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`})()");
    await p.eval(`[...document.querySelectorAll('#dt-list input')].find(i=>i.value===${JSON.stringify(iso)}).click()`);
    await p.click("#dt-ok");
    t = await p.text();
    check("Insert › Date & time inserts the chosen format", t.includes(iso), t);

    await p.key("Enter");
    await p.type("hello wide world. second sentence");
    await p.dragSelect("hello wide world. second sentence");
    await p.key("b", CTRL);
    await p.dragSelect("hello wide");
    await openMenu(p, "#fb-format", "UPPERCASE");
    let m = await p.wordModel();
    let last = m.blocks[m.blocks.length - 1];
    check("Format › UPPERCASE changes the selection and keeps bold", runsText(last) === "HELLO WIDE world. second sentence" && last.runs.every((r) => r.b), last);
    await p.dragSelect("HELLO WIDE world. second sentence");
    await openMenu(p, "#fb-format", "Sentence case");
    m = await p.wordModel();
    last = m.blocks[m.blocks.length - 1];
    check("Sentence case", runsText(last) === "Hello wide world. Second sentence", runsText(last));
    await p.dragSelect("Hello wide world. Second sentence");
    await openMenu(p, "#fb-format", "Title Case");
    m = await p.wordModel();
    check("Title Case", runsText(m.blocks[m.blocks.length - 1]) === "Hello Wide World. Second Sentence");
    await p.key("z", CTRL);
    t = await p.text();
    check("Ctrl+Z undoes a case change", t.includes("Hello wide world. Second sentence") || t.includes("Hello Wide World. Second sentenc"), t);

    // format painter: copy bold from the last line to "Price"
    await p.clickAfterText("Second");
    await p.click("#fb-painter");
    check("format painter turns on", (await p.eval("document.getElementById('fb-painter').getAttribute('aria-pressed')")) === "true");
    await p.dragSelect("Price");
    await sleep(150);
    m = await p.wordModel();
    const first = m.blocks[0];
    check("format painter applies the copied formatting to the next selection", first.runs.find((r) => r.t.startsWith("Price"))?.b === true, first);
    check("single-click painter turns itself off", (await p.eval("document.getElementById('fb-painter').getAttribute('aria-pressed')")) === "false");
    await p.click("#fb-painter", 2);
    await p.dragSelect("42");
    await sleep(100);
    check("double-click keeps the painter on", (await p.eval("document.getElementById('fb-painter').getAttribute('aria-pressed')")) === "true");
    await p.key("Escape");
    check("Esc stops the painter", (await p.eval("document.getElementById('fb-painter').getAttribute('aria-pressed')")) === "false");
  }],

  ["Other editors are not affected (slides, LaTeX, Excel)", async (open) => {
    for (const kind of ["slides", "latex", "excel"]) {
      const p = await open(`?new=${kind}`);
      await p.waitFor(`__officeViewer.doc?.kind === ${JSON.stringify(kind)}`);
      await sleep(300);
      const ui = await p.eval("({zoom:document.getElementById('zoom-ctl').hidden, nav:document.getElementById('nav-pane').hidden, sheet:document.getElementById('word-sheet').hidden, find:document.getElementById('find-bar').hidden, bar:document.getElementById('word-bar').hidden})");
      check(`${kind}: no Word zoom / navigation / find / page UI`, Object.values(ui).every(Boolean), ui);
      await p.key("=", CTRL);
      check(`${kind}: Ctrl+= does not change the Word zoom`, (await p.eval("__officeViewer.zoom")) === 100);
    }
  }],

  ["New from template dialog", async (open) => {
    const p = await open("");
    await p.click("#btn-new");
    await p.click("#mi-new-template");
    await p.waitFor("!document.getElementById('tpl-modal').hidden");
    const n = await p.eval("document.querySelectorAll('#tpl-grid .tpl-item').length");
    check("template dialog lists 6 templates with previews", n === 6 && (await p.eval("document.querySelectorAll('#tpl-grid .tpl-thumb .ow-page h1, #tpl-grid .tpl-thumb .ow-page p').length")) > 10, n);
    await p.shot("templates");
    await p.clickText("#tpl-grid .tpl-name", "Résumé");
    await p.click("#tpl-ok");
    await p.waitFor("__officeViewer.doc?.kind === 'word'");
    const m = await p.wordModel();
    check("Résumé template opened", runsText(m.blocks[0]) === "[YOUR NAME]" && m.blocks.some((b) => b.type === "table"), m.blocks[0]);
    check("address bar keeps the template", (await p.eval("location.search")).includes("template=resume"));
  }],

  ["Page setup, header/footer with page numbers and properties are saved in the .docx", async (open) => {
    const p = await open("?new=word");
    await p.click("#word-view");
    await p.type("Body text of the report.");
    // page setup
    await openMenu(p, "#fb-layout", "Page setup");
    await p.waitFor("!document.getElementById('ps-modal').hidden");
    await p.choose("#ps-size", "A4");
    await p.click("#ps-landscape");
    await p.choose("#ps-preset", "narrow");
    await p.shot("pagesetup");
    await p.click("#ps-ok");
    const pw = await p.eval("document.getElementById('word-sheet').getBoundingClientRect().width / (__officeViewer.zoom/100)");
    check("the on-screen page follows the page setup (A4 landscape)", Math.abs(pw - 16838 / 15) < 3, pw);
    // header: click the band and type, insert page number with the tools
    await p.click("#word-header");
    await p.waitFor("!document.getElementById('hf-tools').hidden");
    check("clicking the header band edits it", (await p.eval("document.activeElement.id")) === "word-header");
    await p.type("Acme report — page ");
    await p.clickText("#hf-tools button", "Insert page number");
    await p.key("b", CTRL);
    await p.type(" draft");
    await p.shot("header");
    await p.key("Escape");
    check("Esc leaves the header", (await p.eval("document.activeElement.id")) === "word-view");
    // footer: Insert › Page numbers
    await openMenu(p, "#fb-insert", "Page numbers");
    check("Insert › Page numbers fills the footer", (await p.eval("document.querySelectorAll('#word-footer .ow-field').length")) === 2);
    // properties
    await p.click("#btn-more");
    await p.clickText("#menu-more button", "Document properties");
    await p.waitFor("!document.getElementById('props-modal').hidden");
    await p.click("#props-doc-title"); await p.type("Annual & <Report>");
    await p.click("#props-author"); await p.type("Ann Author");
    await p.click("#props-subject"); await p.type("Finance");
    await p.click("#props-ok");
    const m = await p.wordModel();
    check("model meta: header with {PAGE} token", JSON.stringify(m.meta.header) === JSON.stringify({ runs: [{ t: "Acme report — page {PAGE}" }, { t: " draft", b: true }] }), m.meta.header);
    check("model meta: footer, page, properties", m.meta.footer?.runs?.[0]?.t === "Page {PAGE} of {NUMPAGES}" && m.meta.page?.size === "A4" && m.meta.title === "Annual & <Report>" && m.meta.author === "Ann Author", m.meta);
    await p.shot("page");
    const saved = await p.saveDocx();
    const d = saved.files["word/document.xml"];
    check("saved docx: A4 landscape pgSz", /<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"\/>/.test(d), d.match(/<w:sectPr>.*<\/w:sectPr>/)?.[0]);
    check("saved docx: narrow margins", /<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/.test(d));
    check("saved docx: header/footer references", /<w:headerReference w:type="default"/.test(d) && /<w:footerReference w:type="default"/.test(d));
    const h = saved.files["word/header1.xml"] || "";
    const f = saved.files["word/footer1.xml"] || "";
    check("saved docx: header text + PAGE field", h.includes("Acme report — page ") && /<w:instrText[^>]*> PAGE </.test(h) && h.includes("draft"), h);
    check("saved docx: footer PAGE and NUMPAGES fields", /> PAGE </.test(f) && /> NUMPAGES </.test(f), f);
    const core = saved.files["docProps/core.xml"] || "";
    check("saved docx: core.xml title / creator / subject", core.includes("<dc:title>Annual &amp; &lt;Report&gt;</dc:title>") && core.includes("<dc:creator>Ann Author</dc:creator>") && core.includes("<dc:subject>Finance</dc:subject>"), core);
    check("saved docx: content types for header/footer", /header\+xml/.test(saved.files["[Content_Types].xml"]) && /footer\+xml/.test(saved.files["[Content_Types].xml"]));
    if (process.env.E2E_SAVE_DIR) fs.writeFileSync(path.join(process.env.E2E_SAVE_DIR, `word-features${BACKEND ? "-backend" : ""}.docx`), saved.bytes);

    // Reopen the saved file: everything comes back.
    await p.eval(`__officeViewer.openBytes("roundtrip.docx", Uint8Array.from(atob(${JSON.stringify(saved.bytes.toString("base64"))}), c => c.charCodeAt(0)), "", {skipGuard:true})`);
    await p.waitFor("__officeViewer.doc?.name === 'roundtrip.docx'");
    await sleep(300);
    const m2 = await p.wordModel();
    check("reopened: header, footer, page setup and properties are back",
      m2.meta.header?.runs?.map((r) => r.t).join("") === "Acme report — page {PAGE} draft" && m2.meta.footer?.runs?.[0]?.t === "Page {PAGE} of {NUMPAGES}"
      && m2.meta.page?.orientation === "landscape" && m2.meta.subject === "Finance", m2.meta);
    check("reopened: page-number chips in the bands", (await p.eval("document.querySelectorAll('#word-header .ow-field, #word-footer .ow-field').length")) === 3);
    // Printing: @page follows the page setup and carries the running header/footer.
    const pdf = await p.send("Page.printToPDF", { preferCSSPageSize: true, printBackground: false });
    const pdfText = Buffer.from(pdf.data, "base64").toString("latin1");
    const box = /\/MediaBox\s*\[\s*0 0 ([\d.]+) ([\d.]+)\s*\]/.exec(pdfText);
    check("print: PDF pages are A4 landscape", box && Math.abs(Number(box[1]) - 842) < 2 && Math.abs(Number(box[2]) - 595) < 2, box && box[0]);
    const css = await p.eval("document.getElementById('ow-print-page')?.textContent || ''");
    check("print: @page margin boxes show the header / footer with page counters",
      /@top-left \{ content: "Acme report — page " counter\(page\) " draft"/.test(css) && /@bottom-center \{ content: "Page " counter\(page\) " of " counter\(pages\)/.test(css), css);
    // Page view (docx-preview of the saved model) shows the header and footer too.
    await p.click("#btn-preview");
    await p.waitFor("!document.getElementById('word-preview').hidden && document.querySelector('#word-preview section.docx')", 20000);
    const pv = await p.eval("document.getElementById('word-preview').innerText");
    check("Page view shows the header and footer", pv.includes("Acme report") && /Page\s*1\s*of\s*1/.test(pv), pv.slice(0, 200));
    await p.shot("pageview");
    await p.click("#btn-preview");
    // HTML export uses the title and shows the header
    await p.key("s", CTRL | SHIFT);
    await p.waitFor("!document.getElementById('saveas-modal').hidden");
    await p.choose("#sa-format", "html");
    await p.click("#sa-ok");
    const [w] = await p.written();
    const html = w.bytes.toString("utf8");
    check("HTML export: <title> from the properties, header/footer text", html.includes("<title>Annual &amp; &lt;Report&gt;</title>") && html.includes("Acme report — page 1") && html.includes("Page 1 of 1"), html.slice(0, 400));
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
    let extId = null;
    for (let i = 0; i < 80 && !extId; i++) {
      const list = await (await fetch(`http://127.0.0.1:${PORT}/json/list`)).json();
      const sw = list.find((t) => /\/background\/service-worker\.js$/.test(t.url));
      if (sw) extId = new URL(sw.url).host;
      else await sleep(150);
    }
    if (!extId) throw new Error("the extension did not load");
    const viewer = `chrome-extension://${extId}/viewer/viewer.html`;

    const open = async (query, { width = 1400, height = 900 } = {}) => {
      const { targetId } = await browserConn.send("Target.createTarget", { url: "about:blank" });
      const list = await (await fetch(`http://127.0.0.1:${PORT}/json/list`)).json();
      const c = await connect(list.find((t) => t.id === targetId).webSocketDebuggerUrl);
      for (const d of ["Runtime", "Page", "DOM"]) await c.send(`${d}.enable`);
      await c.send("Emulation.setFocusEmulationEnabled", { enabled: true });
      await c.send("Emulation.setDeviceMetricsOverride", { width, height, deviceScaleFactor: 1, mobile: false });
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
      await p.eval(`chrome.storage.local.set({settings: {backend_enabled: ${!!BACKEND}, backend_url: ${JSON.stringify(BACKEND || "http://127.0.0.1:47831")}}})`);
      await p.eval("localStorage.removeItem('officeEditor.word.zoom'); localStorage.removeItem('officeEditor.word.nav')");
      await browserConn.send("Target.closeTarget", { targetId: p.targetId });
    }
    if (BACKEND) {
      const p = await open("");
      const ok = await p.waitFor("__officeViewer.backend.ok", 10000).catch(() => false);
      check("backend is up for this run", ok);
      await browserConn.send("Target.closeTarget", { targetId: p.targetId });
    }
    const only = process.env.E2E_ONLY ? new RegExp(process.env.E2E_ONLY, "i") : null;
    for (const [name, fn] of TESTS) {
      if (only && !only.test(name)) continue;
      console.log(`\n# ${name}`);
      const pages = [];
      try {
        await fn(async (q, o) => { const p = await open(q, o); pages.push(p); return p; });
      } catch (e) {
        check(`${name}: ran to the end`, false, e.stack || String(e));
      }
      for (const p of pages) {
        if (p.logs.length) check(`${name}: no page exceptions`, false, p.logs.join("\n"));
        p.close();
        await browserConn.send("Target.closeTarget", { targetId: p.targetId }).catch(() => {});
      }
    }
    // A narrow window: the Word toolbar still fits (no clipped controls).
    for (const width of [1000, 1400]) {
      if (only && !only.test(`toolbar at ${width}px`)) continue;
      console.log(`\n# toolbar at ${width}px`);
      const p = await open("?new=word", { width, height: 760 });
      const measure = "(()=>{const bar=document.getElementById('word-bar'); const r=bar.getBoundingClientRect(); return {h:Math.round(r.height), over:[...bar.querySelectorAll('button,select')].filter(e=>e.offsetParent && e.getBoundingClientRect().right>r.right+1).length}})()";
      const fit = await p.eval(measure);
      check(`Word toolbar at ${width}px: one row, nothing cut off`, fit.over === 0 && fit.h <= 40, fit);
      if (width === 1000) await p.shot("1000px");
      // With the cursor in a table the Table menu appears without wrapping the bar.
      await p.click("#word-view");
      await p.click("#fb-table");
      await p.click('#tp-grid button[data-r="2"][data-c="2"]');
      await p.waitFor("!document.getElementById('table-tools').hidden");
      const fit2 = await p.eval(measure);
      check(`Word toolbar at ${width}px in a table: still one row`, fit2.over === 0 && fit2.h <= 40, fit2);
      if (p.logs.length) check(`${width}px: no page exceptions`, false, p.logs.join("\n"));
      p.close();
      await browserConn.send("Target.closeTarget", { targetId: p.targetId }).catch(() => {});
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
