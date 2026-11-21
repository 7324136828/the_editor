/**
 * End-to-end tests of the excel-features stream (sort, AutoFilter, freeze
 * panes, find & replace, fill handle, formula autocomplete, charts, saving)
 * in the real viewer page with real mouse/keyboard input.
 *
 * NOT part of `npm test`. Run from office-chrome-extension/:
 *     node tests/e2e/excel-features.mjs
 * Harness copied from tests/e2e/run.mjs (own browser, own profile, save-picker
 * stub that captures the written bytes).
 *
 * Environment overrides:
 *   BROWSER_PATH       browser executable (default: Edge, then Chrome)
 *   E2E_PORT           remote debugging port (default 9354)
 *   E2E_PROFILE        user-data-dir (default: <os tmpdir>/office-ext-xf-<pid>, deleted afterwards)
 *   E2E_BACKEND_URL    use this office-backend instead of the built-in engines
 *   E2E_SHOTS          folder for screenshots (default: <os tmpdir>/excel-features-shots)
 *   E2E_PYTHON         python with openpyxl, to check the saved file
 */

import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { spawn, execFileSync } from "node:child_process";
import { fileURLToPath } from "node:url";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const EXT = path.resolve(HERE, "../..");
const EXAMPLE = path.resolve(EXT, "../example");
const PORT = Number(process.env.E2E_PORT || 9354);
const OWN_PROFILE = !process.env.E2E_PROFILE;
const PROFILE = process.env.E2E_PROFILE || path.join(os.tmpdir(), `office-ext-xf-${process.pid}`);
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
      const map = { Enter: [13, "\r"], Tab: [9, ""], Escape: [27, ""], Delete: [46, ""], End: [35, ""], Home: [36, ""], Backspace: [8, ""], ArrowDown: [40, ""], ArrowUp: [38, ""], ArrowLeft: [37, ""], ArrowRight: [39, ""], F2: [113, ""] };
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

// ---------------------------------------------------------------- excel-features helpers

const SHOTS = process.env.E2E_SHOTS || path.join(os.tmpdir(), "excel-features-shots");
const PY = process.env.E2E_PYTHON || "";

async function shot(p, name) {
  try {
    fs.mkdirSync(SHOTS, { recursive: true });
    const r = await p.send("Page.captureScreenshot", { format: "png" });
    fs.writeFileSync(path.join(SHOTS, `excel-features-${name}${BACKEND ? "-backend" : ""}.png`), Buffer.from(r.data, "base64"));
  } catch (e) {
    console.log(`(screenshot ${name} failed: ${e.message})`);
  }
}

/** A workbook with a small sales table, written with the extension's own writer. */
async function makeWorkbook() {
  const vm = await import("node:vm");
  if (!globalThis.XLSX) vm.runInThisContext(fs.readFileSync(path.join(EXT, "vendor/xlsx.full.min.js"), "utf8"));
  const L = await import(new URL(`file:///${path.join(EXT, "lib/local-xlsx.js").replace(/\\/g, "/")}`).href);
  const regions = ["East", "West", "North", "South"];
  const cells = {
    A1: { t: "s", v: "Region", s: { b: true } }, B1: { t: "s", v: "Product", s: { b: true } }, C1: { t: "s", v: "Units", s: { b: true } },
    D1: { t: "s", v: "Price", s: { b: true } }, E1: { t: "s", v: "Total", s: { b: true } },
  };
  for (let i = 0; i < 40; i++) {
    const r = i + 2;
    cells[`A${r}`] = { t: "s", v: regions[(i * 7) % 4] };
    cells[`B${r}`] = { t: "s", v: `Item ${String.fromCharCode(65 + ((i * 5) % 9))}` };
    cells[`C${r}`] = { t: "n", v: ((i * 37) % 90) + 5 };
    cells[`D${r}`] = { t: "n", v: ((i * 13) % 20) + 1.5 };
    cells[`E${r}`] = { t: "f", f: `=C${r}*D${r}`, v: null };
  }
  const model = { kind: "excel", active: 0, sheets: [{ name: "Sales", cells, colWidths: { B: 12 } }, { name: "Notes", cells: { A1: { t: "s", v: "East coast notes" } } }] };
  const file = path.join(PROFILE, "sales.xlsx");
  fs.writeFileSync(file, L.modelToBytes(model, "xlsx"));
  return file;
}

async function openSales(open) {
  const p = await open("");
  await p.setFiles([await makeWorkbook()]);
  await p.waitFor("__officeViewer.doc?.kind === 'excel'");
  await sleep(400);
  return p;
}

const sheetJ = (p, prop) => p.eval(`JSON.stringify(__officeViewer.doc.model.sheets[__officeViewer.doc.model.active||0][${JSON.stringify(prop)}] ?? null)`).then(JSON.parse);
const col = (p, letter, from, to) => p.eval(`JSON.stringify(Array.from({length:${to - from + 1}},(_, i)=>{const c=__officeViewer.doc.model.sheets[0].cells['${letter}'+(${from}+i)]; return c ? c.v : null}))`).then(JSON.parse);
const center = (p, sel) => p.eval(`(()=>{const e=document.querySelector(${JSON.stringify(sel)}); if(!e) return null; const b=e.getBoundingClientRect(); return b.width ? {x:b.x+b.width/2,y:b.y+b.height/2,w:b.width,h:b.height,l:b.x,t:b.y} : null})()`);

async function drag(p, from, to, steps = 8, modifiers = 0) {
  await p.send("Input.dispatchMouseEvent", { type: "mouseMoved", x: from.x, y: from.y });
  await p.send("Input.dispatchMouseEvent", { type: "mousePressed", x: from.x, y: from.y, button: "left", clickCount: 1, modifiers });
  for (let i = 1; i <= steps; i++) {
    const x = from.x + ((to.x - from.x) * i) / steps, y = from.y + ((to.y - from.y) * i) / steps;
    await p.send("Input.dispatchMouseEvent", { type: "mouseMoved", x, y, button: "left", buttons: 1, modifiers });
    await sleep(15);
  }
  await p.send("Input.dispatchMouseEvent", { type: "mouseReleased", x: to.x, y: to.y, button: "left", clickCount: 1, modifiers });
  await sleep(80);
}

async function keyCombo(p, key, code, vk, modifiers) {
  await p.send("Input.dispatchKeyEvent", { type: "rawKeyDown", key, code, windowsVirtualKeyCode: vk, modifiers });
  await p.send("Input.dispatchKeyEvent", { type: "keyUp", key, code, windowsVirtualKeyCode: vk, modifiers });
  await sleep(40);
}

function pyCheck(file, code) {
  if (!PY) return null;
  try {
    return execFileSync(PY, ["-I", "-c", code, file], { encoding: "utf8" }).trim();
  } catch (e) {
    return `python failed: ${e.message}`;
  }
}

// ---------------------------------------------------------------- tests

const results = [];
function check(name, ok, detail = "") {
  results.push({ name, ok: !!ok });
  console.log(`${ok ? "PASS" : "FAIL"}  ${name}${ok ? "" : `\n      ${typeof detail === "string" ? detail : JSON.stringify(detail)}`}`);
}

const TESTS = [
  ["sort with the toolbar: numbers, header kept, formulas follow their rows, undo", async (open) => {
    const p = await openSales(open);
    await p.clickCell("C5");
    await p.click('#excel-bar [data-act="sort-desc"]');
    await sleep(200);
    const units = await col(p, "C", 2, 41);
    const sorted = units.every((v, i) => i === 0 || units[i - 1] >= v);
    check("Units sorted largest first", sorted && units.length === 40, units.slice(0, 6));
    check("header row stayed in place", (await p.cell("C1"))?.v === "Units");
    const e2 = await p.cell("E2"), c2 = await p.cell("C2"), d2 = await p.cell("D2");
    check("formula moved with its row and still refers to its own row", e2 && e2.f === "=C2*D2" && Math.abs(e2.v - c2.v * d2.v) < 1e-9, e2);
    check("sorting marks the workbook dirty", (await p.state()).dirty === true);
    await shot(p, "sorted");
    await p.key("z", CTRL);
    const back = await col(p, "C", 2, 4);
    check("Ctrl+Z restores the original order", JSON.stringify(back) === JSON.stringify([5, 42, 79]), back);
    // Custom sort dialog: Region A-Z then Units Z-A
    await p.clickCell("B3");
    await p.click('#excel-bar [data-act="sort-dialog"]');
    await p.waitFor("!!document.querySelector('.xv-dlg')");
    check("Sort dialog detected the header row", await p.eval("document.querySelector('.xv-sd-hdr').checked"));
    await p.choose(".xv-sd-col", "0");
    await p.click(".xv-sd-add");
    await p.eval("(()=>{const s=document.querySelectorAll('.xv-sd-col')[1]; s.value='2'; s.dispatchEvent(new Event('change',{bubbles:true})); const o=document.querySelectorAll('.xv-sd-ord')[1]; o.value='desc'; o.dispatchEvent(new Event('change',{bubbles:true}))})()");
    await shot(p, "sort-dialog");
    await p.click(".xv-sd-ok");
    const reg = await col(p, "A", 2, 41), un = await col(p, "C", 2, 41);
    let ok = true;
    for (let i = 1; i < 40; i++) {
      if (reg[i - 1] > reg[i]) ok = false;
      if (reg[i - 1] === reg[i] && un[i - 1] < un[i]) ok = false;
    }
    check("two-level sort (Region A-Z, then Units Z-A)", ok, { reg: reg.slice(0, 12), un: un.slice(0, 12) });
  }],

  ["AutoFilter: dropdown with real clicks, value list, number filter, hidden rows, clear", async (open) => {
    const p = await openSales(open);
    await p.clickCell("B4");
    await p.click('#excel-bar [data-act="filter"]');
    check("filter range set on the current region", (await sheetJ(p, "autoFilter")) === "A1:E41", await sheetJ(p, "autoFilter"));
    check("toolbar Filter button shows pressed", await p.eval(`document.querySelector('#excel-bar [data-act="filter"]').getAttribute('aria-pressed') === 'true'`));
    check("header cells show dropdown buttons", (await p.eval("document.querySelectorAll('.xv-fbtn').length")) === 5);
    await p.click('.xv-fbtn[data-fc="0"]');
    await p.waitFor("!!document.querySelector('.xv-fpanel')");
    const items = await p.eval("[...document.querySelectorAll('.xv-fp-item span')].map(s=>s.textContent).join(',')");
    check("value list is sorted and distinct", items === "East,North,South,West", items);
    await p.click(".xv-fp-all input");      // unselect all
    await p.clickText(".xv-fp-item span", "East");
    await shot(p, "filter-dropdown");
    await p.click(".xv-fp-ok");
    await sleep(150);
    const hidden = await p.eval("__officeViewer.excelView._hidden.length");
    check("rows of other regions are hidden (30 of 40)", hidden === 30, hidden);
    const rowsShown = await p.eval("[...document.querySelectorAll('.xv-rowhdr-inner .xv-rhd')].slice(0,6).map(e=>e.textContent).join(',')");
    check("row numbers show gaps", rowsShown.startsWith("1,2,6,") || /^1,\d+,\d+/.test(rowsShown) && rowsShown !== "1,2,3,4,5,6", rowsShown);
    check("filtered header button shows the active icon", await p.eval("document.querySelector('.xv-fbtn[data-fc=\"0\"]').classList.contains('xv-fbtn-on')"));
    await shot(p, "filtered");
    // Alt+Down on a header cell opens its filter menu from the keyboard
    await p.clickCell("C1");
    await p.send("Input.dispatchKeyEvent", { type: "rawKeyDown", key: "ArrowDown", code: "ArrowDown", windowsVirtualKeyCode: 40, modifiers: 1 });
    await p.send("Input.dispatchKeyEvent", { type: "keyUp", key: "ArrowDown", code: "ArrowDown", windowsVirtualKeyCode: 40, modifiers: 1 });
    await sleep(100);
    check("Alt+Down opens the filter menu of the current header", await p.eval("document.querySelector('.xv-fpanel')?.dataset.c === '2' && document.querySelector('.xv-fpanel').contains(document.activeElement)"));
    await p.key("Escape");
    check("Esc closes it and returns to the grid", await p.eval("!document.querySelector('.xv-fpanel') && document.activeElement.classList.contains('xv-editor')"));
    // Arrow keys skip hidden rows
    await p.clickCell("A2");
    await p.key("ArrowDown");
    const nb = await p.nameBox();
    check("ArrowDown skips filtered rows", nb !== "A3" && (await p.cell(nb))?.v === "East", nb);
    // number filter: Units greater than 50
    await p.click('.xv-fbtn[data-fc="2"]');
    await p.waitFor("!!document.querySelector('.xv-fpanel')");
    await p.choose(".xv-fp-op", "gt");
    await p.click(".xv-fp-a");
    await p.type("50");
    await p.key("Enter");
    await sleep(150);
    const visible = await p.eval("(()=>{const v=__officeViewer.excelView; const out=[]; for(let r=1;r<=40;r++) if(!v._isHidden(r)) out.push(r); return out.map(r=>[__officeViewer.doc.model.sheets[0].cells['A'+(r+1)].v, __officeViewer.doc.model.sheets[0].cells['C'+(r+1)].v])})()");
    check("both filters apply (East and Units > 50)", visible.length > 0 && visible.every(([a, c]) => a === "East" && c > 50), visible);
    await p.key("z", CTRL);
    check("Ctrl+Z undoes the number filter", (await p.eval("__officeViewer.excelView._hidden.length")) === 30);
    // Clear the region filter from its menu
    await p.click('.xv-fbtn[data-fc="0"]');
    await p.waitFor("!!document.querySelector('.xv-fpanel')");
    await p.clickText(".xv-fp-act", 'Clear filter from "Region"');
    check("clear filter shows every row", (await p.eval("__officeViewer.excelView._hidden.length")) === 0);
    // Search box in the list
    await p.click('.xv-fbtn[data-fc="1"]');
    await p.waitFor("!!document.querySelector('.xv-fpanel')");
    await p.click(".xv-fp-search");
    await p.type("item a");
    await p.key("Enter");
    await sleep(100);
    const shownB = await p.eval("(()=>{const v=__officeViewer.excelView; const s=new Set(); for(let r=1;r<=40;r++) if(!v._isHidden(r)) s.add(__officeViewer.doc.model.sheets[0].cells['B'+(r+1)].v); return [...s].join(',')})()");
    check("search box keeps only matching values", shownB === "Item A", shownB);
    // Turning the filter off shows all rows and removes the range
    await p.clickCell("A1");
    await p.click('#excel-bar [data-act="filter"]');
    check("filter toggled off", (await sheetJ(p, "autoFilter")) === null && (await p.eval("__officeViewer.excelView._hidden.length")) === 0);
  }],

  ["freeze panes: top row + first column stay while scrolling; clicks on frozen cells work", async (open) => {
    const p = await openSales(open);
    await p.clickCell("B2");
    await p.click('#excel-bar [data-act="freeze"]');
    await p.waitFor("!document.querySelector('.xv-menu').hidden");
    await p.clickText(".xv-menu button span", "Freeze panes above and left of B2");
    check("freeze stored in the model", JSON.stringify(await sheetJ(p, "freeze")) === JSON.stringify({ rows: 1, cols: 1 }), await sheetJ(p, "freeze"));
    const g = await center(p, ".xv-scroll");
    await p.send("Input.dispatchMouseEvent", { type: "mouseWheel", x: g.x, y: g.y, deltaX: 0, deltaY: 480 });
    await sleep(150);
    await p.send("Input.dispatchMouseEvent", { type: "mouseWheel", x: g.x, y: g.y, deltaX: 150, deltaY: 0 });
    await sleep(300);
    const st = await p.eval("JSON.stringify({t:document.querySelector('.xv-scroll').scrollTop,l:document.querySelector('.xv-scroll').scrollLeft})").then(JSON.parse);
    check("the grid scrolled", st.t > 100 && st.l > 50, st);
    // what is drawn at the top-left of the cell area?
    const topText = await p.eval(`(()=>{const r=document.querySelector('.xv-scroll').getBoundingClientRect(); const h=document.querySelector('.xv-colhdr-inner .xv-ch[data-c="4"]').getBoundingClientRect(); const e=document.elementFromPoint(h.x+h.width/2, r.y+10); return e ? e.textContent : ''})()`);
    check("header row is still visible at the top", ["Product", "Units", "Price", "Total"].includes(topText.trim()), topText);
    const leftText = await p.eval(`(()=>{const r=document.querySelector('.xv-scroll').getBoundingClientRect(); const e=document.elementFromPoint(r.x+20, r.y+100); return e ? e.textContent : ''})()`);
    check("first column is still visible at the left", ["East", "West", "North", "South"].includes(leftText.trim()), leftText);
    await shot(p, "frozen-scrolled");
    const r = await p.eval("(()=>{const r=document.querySelector('.xv-scroll').getBoundingClientRect(); return {x:r.x+20,y:r.y+10}})()");
    await p.mouse(r.x, r.y);
    check("clicking the frozen corner cell selects A1", (await p.nameBox()) === "A1", await p.nameBox());
    await p.key("x");
    await p.key("Enter");
    check("editing a frozen cell works", (await p.cell("A1"))?.v === "x", await p.cell("A1"));
    await p.key("z", CTRL);
    await p.click('#excel-bar [data-act="freeze"]');
    await p.clickText(".xv-menu button span", "Unfreeze panes");
    check("unfreeze removes it", (await sheetJ(p, "freeze")) === null);
    await p.key("z", CTRL);
    check("Ctrl+Z brings the frozen panes back", JSON.stringify(await sheetJ(p, "freeze")) === JSON.stringify({ rows: 1, cols: 1 }));
  }],

  ["find & replace with the keyboard: Ctrl+F / Ctrl+H, workbook scope, replace all", async (open) => {
    const p = await openSales(open);
    await p.clickCell("A1");
    await p.key("f", CTRL);
    check("Ctrl+F opens find", await p.eval("!document.querySelector('.xv-find').hidden && document.activeElement.classList.contains('xv-find-input')"));
    await p.type("north");
    await p.key("Enter");
    const nb = await p.nameBox();
    check("Enter finds the next match and selects it", (await p.cell(nb))?.v === "North", nb);
    const count0 = await p.eval("document.querySelector('.xv-find-count').textContent");
    check("match counter", /of 10$/.test(count0), count0);
    await p.key("Escape");
    check("Esc closes and returns to the grid", await p.eval("document.querySelector('.xv-find').hidden && document.activeElement.classList.contains('xv-editor')"));
    await p.key("h", CTRL);
    check("Ctrl+H opens replace", await p.eval("document.querySelector('.xv-find').classList.contains('xv-find-replace')"));
    await p.eval("document.querySelector('.xv-find-input').select()");
    await p.type("East");
    await p.choose(".xv-find-within", "book");
    await p.click(".xv-rep-input");
    await p.type("Eastern");
    await shot(p, "replace");
    await p.click(".xv-rep-all");
    const all = await p.eval("__officeViewer.doc.model.sheets.map(s=>Object.values(s.cells).filter(c=>typeof c.v==='string'&&c.v.includes('Eastern')).length).join(',')");
    check("replace all in the workbook (10 cells + the note on sheet 2)", all === "10,1", all);
    const notice = await p.eval("document.querySelector('.xv-notice').textContent");
    check("replace all reports the count", /Replaced 11 occurrences in 11 cells/.test(notice), notice);
    await p.key("Escape");
    await p.key("z", CTRL);
    const after = await p.eval("__officeViewer.doc.model.sheets.map(s=>Object.values(s.cells).filter(c=>typeof c.v==='string'&&c.v.includes('Eastern')).length).join(',')");
    check("one Ctrl+Z undoes the whole replace", after === "0,0", after);
    // keys outside the Excel view are left alone
    await p.click("#formula-bar");
    await p.key("f", CTRL);
    check("Ctrl+F in the formula bar is not taken by the grid", await p.eval("document.querySelector('.xv-find').hidden"));
  }],

  ["fill handle drag with the mouse: number series, weekdays, formulas; Ctrl+D", async (open) => {
    const p = await open("?new=excel");
    await p.clickCell("A1"); await p.type("1\n2\n");
    await p.clickCell("B1"); await p.type("Mon\n");
    await p.clickCell("C1"); await p.type("Item 1\n");
    await p.clickCell("D1"); await p.type("=A1*10\n");
    // select A1:A2 then drag the handle to A6
    await p.clickCell("A1");
    await p.key("ArrowDown", SHIFT);
    const h = await center(p, ".xv-fillh");
    const a6 = await p.cellXY("A6");
    await drag(p, h, { x: a6.x, y: a6.y });
    check("1,2 extends to 3,4,5,6", JSON.stringify(await col(p, "A", 1, 6)) === "[1,2,3,4,5,6]", await col(p, "A", 1, 6));
    await shot(p, "fill");
    await p.clickCell("B1");
    const hb = await center(p, ".xv-fillh");
    const b8 = await p.cellXY("B8");
    await drag(p, hb, b8);
    check("weekday names continue", JSON.stringify(await col(p, "B", 1, 8)) === JSON.stringify(["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun", "Mon"]), await col(p, "B", 1, 8));
    await p.clickCell("C1");
    const hc = await center(p, ".xv-fillh");
    const c3 = await p.cellXY("C3");
    await drag(p, hc, c3);
    check('"Item 1" counts up', JSON.stringify(await col(p, "C", 1, 3)) === JSON.stringify(["Item 1", "Item 2", "Item 3"]), await col(p, "C", 1, 3));
    await p.clickCell("D1");
    const hd = await center(p, ".xv-fillh");
    const d4 = await p.cellXY("D4");
    await drag(p, hd, d4);
    const d4c = await p.cell("D4");
    check("formula copied with relative references shifted", d4c && d4c.f === "=A4*10" && d4c.v === 40, d4c);
    await p.key("z", CTRL);
    check("fill is undoable", (await p.cell("D4")) === null);
    // single number: plain drag copies, Ctrl+drag counts
    await p.clickCell("E1"); await p.type("7\n");
    await p.clickCell("E1");
    const he = await center(p, ".xv-fillh");
    await drag(p, he, await p.cellXY("E3"));
    check("single number copies", JSON.stringify(await col(p, "E", 1, 3)) === "[7,7,7]", await col(p, "E", 1, 3));
    await p.clickCell("F1"); await p.type("7\n");
    await p.clickCell("F1");
    const hf = await center(p, ".xv-fillh");
    await drag(p, hf, await p.cellXY("F3"), 8, CTRL);
    check("Ctrl+drag on a single number counts up", JSON.stringify(await col(p, "F", 1, 3)) === "[7,8,9]", await col(p, "F", 1, 3));
    // Ctrl+D
    await p.clickCell("D1");
    await p.key("ArrowDown", SHIFT); await p.key("ArrowDown", SHIFT);
    await p.key("d", CTRL);
    check("Ctrl+D fills down (formula shifted)", (await p.cell("D3"))?.f === "=A3*10", await p.cell("D3"));
  }],

  ["formula autocomplete: =SU lists functions, Tab inserts SUM(, argument hint shows", async (open) => {
    const p = await open("?new=excel");
    await p.clickCell("A1"); await p.type("4\n5\n6\n");
    await p.clickCell("B1");
    await p.type("=SU");
    await sleep(100);
    const list = await p.eval("[...document.querySelectorAll('.xv-ac:not([hidden]) .xv-ac-item b')].map(b=>b.textContent).join(',')");
    check("list shows matching functions", list.split(",").includes("SUM") && list.split(",").includes("SUMIFS"), list);
    const desc = await p.eval("document.querySelector('.xv-ac .xv-ac-sig')?.textContent || ''");
    check("selected function shows its signature", /^SUM\(/.test(desc), desc);
    await shot(p, "autocomplete");
    await p.key("Tab");
    const v = await p.eval("document.querySelector('.xv-editor').value");
    check("Tab inserts SUM(", v === "=SUM(", v);
    const hint = await p.eval("(()=>{const s=document.querySelector('.xv-sig'); return s && !s.hidden ? s.textContent : ''})()");
    check("argument hint while typing arguments", /SUM\(number1/.test(hint), hint);
    const bold = await p.eval("document.querySelector('.xv-sig b')?.textContent || ''");
    check("current argument is highlighted", bold === "number1", bold);
    await p.type("A1:A3,");
    const bold2 = await p.eval("document.querySelector('.xv-sig b')?.textContent || ''");
    check("hint follows to the next argument", /number2/.test(bold2), bold2);
    await shot(p, "signature");
    await p.type("10)");
    await p.key("Enter");
    const b1 = await p.cell("B1");
    check("formula computed", b1 && b1.f === "=SUM(A1:A3,10)" && b1.v === 25, b1);
    // Arrow keys move through the list and Enter accepts
    await p.clickCell("C1");
    await p.type("=vlo");
    await p.key("Enter");
    check("Enter accepts the suggestion", (await p.eval("document.querySelector('.xv-editor').value")) === "=VLOOKUP(");
    await p.key("Escape"); await p.key("Escape");
  }],

  ["charts: insert from the selection, move with the mouse, edit, delete and undo", async (open) => {
    const p = await openSales(open);
    await p.clickCell("C1");
    await p.eval("__officeViewer.excelView.select('B1:C9')");
    await p.click('#excel-bar [data-act="chart"]');
    await sleep(200);
    const charts = await sheetJ(p, "charts");
    check("chart added to the model", charts && charts.length === 1 && charts[0].range === "B1:C9" && charts[0].type === "column", charts);
    check("chart drawn as SVG over the grid", await p.eval("!!document.querySelector('.xv-chart svg rect') && document.querySelectorAll('.xv-chart svg path').length >= 8"));
    await shot(p, "chart-inserted");
    const c0 = await center(p, ".xv-chart");
    await drag(p, { x: c0.x, y: c0.y }, { x: c0.x - 150, y: c0.y + 60 });
    const moved = (await sheetJ(p, "charts"))[0];
    check("dragging moves the chart", Math.abs(moved.x - (charts[0].x - 150)) <= 2 && Math.abs(moved.y - (charts[0].y + 60)) <= 2, moved);
    const se = await center(p, ".xv-chart-sel .xv-chh-se");
    await drag(p, se, { x: se.x + 80, y: se.y + 40 });
    const resized = (await sheetJ(p, "charts"))[0];
    check("dragging a corner resizes it", Math.abs(resized.w - 560) <= 3 && Math.abs(resized.h - 328) <= 3, resized);
    // double-click: edit title and type
    const c1 = await center(p, ".xv-chart");
    await p.mouse(c1.x, c1.y, 2);
    await p.waitFor("!!document.querySelector('.xv-chartpanel')");
    await p.type("Units by product");
    await p.choose(".xv-cp-type", "line");
    await p.click('.xv-chartpanel [data-cp="apply"]');
    const edited = (await sheetJ(p, "charts"))[0];
    check("chart panel edits title and type", edited.title === "Units by product" && edited.type === "line", edited);
    await shot(p, "chart-edited");
    // Download PNG / SVG from the panel (the anchor click is captured)
    await p.eval("window.__dl=[]; HTMLAnchorElement.prototype.click=function(){ window.__dl.push({name:this.download, href:this.href}); }");
    const c2 = await center(p, ".xv-chart");
    await p.mouse(c2.x, c2.y, 2);
    await p.waitFor("!!document.querySelector('.xv-chartpanel')");
    await p.click('.xv-chartpanel [data-cp="png"]');
    await p.waitFor("window.__dl.length > 0");
    await p.click('.xv-chartpanel [data-cp="svg"]');
    await p.waitFor("window.__dl.length > 1");
    const dl = await p.eval("Promise.all(window.__dl.map(async d => ({name:d.name, type:(await (await fetch(d.href)).blob()).type, size:(await (await fetch(d.href)).blob()).size})))");
    check("Download PNG and SVG", dl[0].name === "Units by product.png" && dl[0].type === "image/png" && dl[0].size > 2000 && dl[1].name === "Units by product.svg" && dl[1].type === "image/svg+xml", dl);
    await p.key("Escape");
    await p.click(".xv-chart");
    await p.key("Delete");
    check("Delete removes the selected chart", (await sheetJ(p, "charts")) === null);
    await p.key("z", CTRL);
    check("Ctrl+Z brings it back", ((await sheetJ(p, "charts")) || []).length === 1);
  }],

  ["save with freeze, filter and a chart; the file opens in openpyxl with all three", async (open) => {
    const p = await openSales(open);
    await p.clickCell("B2");
    await p.click('#excel-bar [data-act="freeze"]');
    await p.waitFor("!document.querySelector('.xv-menu').hidden");
    await p.clickText(".xv-menu button span", "Freeze top row");
    check("freeze top row", JSON.stringify(await sheetJ(p, "freeze")) === '{"rows":1,"cols":0}', await sheetJ(p, "freeze"));
    await p.clickCell("A3");
    await p.click('#excel-bar [data-act="filter"]');
    await p.eval("__officeViewer.excelView.select('A1:C6')");
    await p.click('#excel-bar [data-act="chart"]');
    await p.clickCell("H2");
    await p.key("s", CTRL);
    await sleep(300);
    if (await p.eval("!document.getElementById('saveas-modal').hidden")) await p.click("#sa-ok");
    const [w] = await p.written(20000);
    if (BACKEND) check("saved by the backend (no fallback toast)", !(await p.state()).toast || !/built-in writer/.test((await p.state()).toast), (await p.state()).toast);
    check("saved an xlsx", w && w.bytes[0] === 0x50, w && w.name);
    const out = path.join(PROFILE, `saved${BACKEND ? "-backend" : ""}.xlsx`);
    fs.writeFileSync(out, w.bytes);
    fs.mkdirSync(SHOTS, { recursive: true });
    fs.copyFileSync(out, path.join(SHOTS, path.basename(out)));
    const py = pyCheck(out, "import sys,openpyxl; wb=openpyxl.load_workbook(sys.argv[1]); ws=wb['Sales']; print(ws.freeze_panes, ws.auto_filter.ref, len(ws._charts), type(ws._charts[0]).__name__ if ws._charts else '-')");
    if (py == null) console.log("(set E2E_PYTHON to check the file with openpyxl)");
    else check("openpyxl: freeze A2, filter A1:E41, 1 bar chart", py === "A2 A1:E41 1 BarChart", py);
    // reopen the saved file in the editor: everything comes back
    await p.setFiles([out]);
    await sleep(400);
    if (await p.eval("!!document.querySelector('#dialog-modal:not([hidden])')")) await p.clickText("#dlg-actions button", "Don't save");
    await p.waitFor("__officeViewer.doc?.name?.startsWith('saved')");
    await sleep(300);
    const fz = await sheetJ(p, "freeze"), af = await sheetJ(p, "autoFilter"), ch = await sheetJ(p, "charts");
    check("reopened: freeze, filter and chart are back", fz && fz.rows === 1 && fz.cols === 0 && af === "A1:E41" && ch && ch.length === 1 && ch[0].range === "A1:C6", { fz, af, ch });
    await shot(p, "reopened");
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
    for (const [name, fn] of TESTS.filter(([n]) => !process.env.E2E_ONLY || n.includes(process.env.E2E_ONLY))) {
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
