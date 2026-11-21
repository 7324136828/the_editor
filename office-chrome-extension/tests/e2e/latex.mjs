/**
 * End-to-end tests of the LaTeX editor (viewer/latex.js) in a real, headless
 * Edge/Chrome with the extension loaded and the office backend OFF.
 *
 * NOT part of `npm test`. Run from office-chrome-extension/:
 *     node tests/e2e/latex.mjs
 *
 * Environment overrides:
 *   BROWSER_PATH   browser executable (default: Edge, then Chrome)
 *   E2E_PORT       remote debugging port (default 9352)
 *   E2E_PROFILE    user-data-dir (default: <os tmpdir>/office-ext-latex-e2e-<pid>, deleted afterwards)
 *   E2E_SHOTS      directory for screenshots (default: none)
 *   E2E_PERF_FACTOR  multiply the 1 MB performance budgets (default 1; CI uses 3)
 */

import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { spawn, execFileSync } from "node:child_process";
import { fileURLToPath } from "node:url";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const EXT = path.resolve(HERE, "../..");
const FIX = path.resolve(HERE, "../fixtures/latex");
const PORT = Number(process.env.E2E_PORT || 9352);
const OWN_PROFILE = !process.env.E2E_PROFILE;
const PROFILE = process.env.E2E_PROFILE || path.join(os.tmpdir(), `office-ext-latex-e2e-${process.pid}`);
const SHOTS = process.env.E2E_SHOTS || "";
// Multiplies the performance budgets below (default 1). CI sets it for slow shared runners.
const PERF = Math.max(1, Number(process.env.E2E_PERF_FACTOR) || 1);
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

/** Captures Save output: showSaveFilePicker returns a handle that keeps the bytes. */
const SAVE_STUB = `(()=>{window.__written=[];window.showSaveFilePicker=async(o)=>({name:o.suggestedName,kind:"file",
queryPermission:async()=>"granted",requestPermission:async()=>"granted",getFile:async()=>new File([],o.suggestedName),
createWritable:async()=>{const parts=[];return{write:async(b)=>parts.push(b),close:async()=>{const u=new Uint8Array(await new Blob(parts).arrayBuffer());
let s="";for(let i=0;i<u.length;i+=0x8000)s+=String.fromCharCode.apply(null,u.subarray(i,i+0x8000));window.__written.push({name:o.suggestedName,b64:btoa(s)})},abort:async()=>{}}}});})()`;

const KEYS = {
  Enter: [13, "\r", "Enter"], Tab: [9, "", "Tab"], Escape: [27, "", "Escape"], Delete: [46, "", "Delete"], End: [35, "", "End"],
  Home: [36, "", "Home"], Backspace: [8, "", "Backspace"], ArrowDown: [40, "", "ArrowDown"], ArrowUp: [38, "", "ArrowUp"],
  ArrowLeft: [37, "", "ArrowLeft"], ArrowRight: [39, "", "ArrowRight"], "/": [191, "/", "Slash"],
};

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
    async hover(selector) {
      const r = await p.box(selector);
      await c.send("Input.dispatchMouseEvent", { type: "mouseMoved", x: r.x, y: r.y });
      await sleep(60);
    },
    async box(selector) {
      const r = await p.eval(`(()=>{const e=document.querySelector(${JSON.stringify(selector)}); if(!e) return null; e.scrollIntoView({block:'nearest'}); const b=e.getBoundingClientRect(); return b.width ? {x:b.x+b.width/2,y:b.y+b.height/2} : null})()`);
      if (!r) throw new Error(`not visible: ${selector}`);
      return r;
    },
    async click(selector) { const r = await p.box(selector); await p.mouse(r.x, r.y); },
    async clickText(selector, text) {
      const r = await p.eval(`(()=>{const e=[...document.querySelectorAll(${JSON.stringify(selector)})].find(n=>n.textContent.trim().startsWith(${JSON.stringify(text)}) && n.getBoundingClientRect().width); if(!e) return null; const b=e.getBoundingClientRect(); return {x:b.x+b.width/2,y:b.y+b.height/2}})()`);
      if (!r) throw new Error(`no ${selector} "${text}"`);
      await p.mouse(r.x, r.y);
    },
    async key(k, modifiers = 0) {
      const [vk, text, code] = KEYS[k] || [k.toUpperCase().charCodeAt(0), k, `Key${k.toUpperCase()}`];
      const withText = !!text && !(modifiers & CTRL);
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
      fs.writeFileSync(path.join(SHOTS, `latex-frontend-${name}.png`), Buffer.from(r.data, "base64"));
    },
    text: () => p.eval("document.querySelector('#latex-view .lx-input').value"),
    caretLine: () => p.eval("(()=>{const t=document.querySelector('#latex-view .lx-input'); return t.value.slice(0,t.selectionStart).split('\\n').length})()"),
    preview: () => p.eval("document.querySelector('#latex-view .lx-preview-inner').innerHTML"),
    state: () => p.eval(`JSON.stringify({name:__officeViewer.doc?.name, kind:__officeViewer.doc?.kind, dirty:__officeViewer.dirty, status:document.getElementById('status-msg').textContent, modals:[...document.querySelectorAll('.modal')].filter(m=>!m.hidden).map(m=>m.id)})`).then(JSON.parse),
    /** Wait until the debounced preview caught up with the source. */
    async rendered() {
      await sleep(350);
      await p.waitFor("(()=>{const v=document.querySelector('#latex-view .lx-input'); return !!v})()");
      await sleep(150);
    },
    async focusEditor() { await p.click("#latex-view .lx-input"); },
  };
  return p;
}

// ---------------------------------------------------------------- tests

const results = [];
function check(name, ok, detail = "") {
  results.push({ name, ok: !!ok });
  console.log(`${ok ? "PASS" : "FAIL"}  ${name}${ok ? "" : `\n      ${typeof detail === "string" ? detail : JSON.stringify(detail)}`}`);
}

const TESTS = [
  ["new LaTeX document: typing with real keys updates the live preview (MathML)", async (open) => {
    const p = await open("?new=latex");
    await p.waitFor("!!document.querySelector('#latex-view .lx-input')");
    const tpl = await p.text();
    check("new document starts from the article template", /\\documentclass\[11pt\]\{article\}/.test(tpl) && /\\maketitle/.test(tpl));
    await p.rendered();
    check("template preview shows the title block and MathML", /ltx-title">An Example Article/.test(await p.preview()) && /<math/.test(await p.preview()));
    await p.shot("01-new-article");
    await p.focusEditor();
    await p.key("a", CTRL);
    await p.type("\\section{Hello}\nEnergy $E=mc^2$ and \\textbf{bold} text.\n");
    const t = await p.text();
    check("auto-closed braces and $ are typed over (no doubled characters)", t === "\\section{Hello}\nEnergy $E=mc^2$ and \\textbf{bold} text.\n", JSON.stringify(t));
    await p.rendered();
    const h = await p.preview();
    check("preview shows the numbered section", /<h2 class="ltx-section"[^>]*><span class="ltx-secnum">1<\/span>Hello<\/h2>/.test(h), h.slice(0, 300));
    check("preview renders the formula as MathML", /<math[^>]*><mrow><mi>E<\/mi><mo>=<\/mo><mi>m<\/mi><msup><mi>c<\/mi><mn>2<\/mn><\/msup><\/mrow><\/math>/.test(h), h);
    check("document is marked dirty", (await p.state()).dirty === true);
    const hl = await p.eval("document.querySelector('#latex-view .lx-hl').textContent");
    check("highlight layer mirrors the text", hl === t.replace(/\n/g, ""), JSON.stringify(hl));
    check("commands are highlighted", await p.eval("!!document.querySelector('#latex-view .lx-hl .t-sec') && !!document.querySelector('#latex-view .lx-hl .t-math')"));
    await p.shot("02-typed");
  }],

  ["smart editing: \\begin + Enter adds \\end, Tab indents, Ctrl+/ comments, Ctrl+B wraps, undo works", async (open) => {
    const p = await open("?new=latex");
    await p.waitFor("!!document.querySelector('#latex-view .lx-input')");
    await p.focusEditor();
    await p.key("a", CTRL);
    await p.type("\\begin{itemize}\n");
    let t = await p.text();
    check("\\end{itemize} inserted with an \\item line", t === "\\begin{itemize}\n  \\item \n\\end{itemize}", JSON.stringify(t));
    await p.type("First\nSecond");
    t = await p.text();
    check("Enter on an item line starts the next \\item", t === "\\begin{itemize}\n  \\item First\n  \\item Second\n\\end{itemize}", JSON.stringify(t));
    await p.rendered();
    check("list shows in the preview", /<li[^>]*><span class="ltx-mark">•<\/span><div class="ltx-item-body"><p[^>]*>First<\/p>/.test(await p.preview()));
    // Ctrl+/ toggles a comment on the current line
    await p.key("/", CTRL);
    t = await p.text();
    check("Ctrl+/ comments the line", t.includes("%   \\item Second"), JSON.stringify(t));
    await p.key("/", CTRL);
    check("Ctrl+/ again uncomments it", (await p.text()).includes("\n  \\item Second\n"));
    // select "First" and bold it
    await p.eval("(()=>{const ta=document.querySelector('#latex-view .lx-input'); const i=ta.value.indexOf('First'); ta.setSelectionRange(i,i+5)})()");
    await p.key("b", CTRL);
    check("Ctrl+B wraps the selection in \\textbf", (await p.text()).includes("\\item \\textbf{First}"));
    await p.key("z", CTRL);
    check("Ctrl+Z undoes the wrap (native undo stack)", (await p.text()).includes("\\item First\n"));
    // Tab indents the selected lines, Shift+Tab outdents
    await p.eval("(()=>{const ta=document.querySelector('#latex-view .lx-input'); ta.setSelectionRange(0, ta.value.length)})()");
    await p.key("Tab");
    check("Tab indents all selected lines", (await p.text()).split("\n").every((l) => l.startsWith("  ")));
    await p.key("Tab", SHIFT);
    check("Shift+Tab outdents them again", (await p.text()).startsWith("\\begin{itemize}"));
    // autocompletion
    await p.key("End", CTRL);
    await p.type("\n\\subs");
    const ac = await p.eval("(()=>{const a=document.querySelector('#latex-view .lx-ac'); return a.hidden ? null : [...a.querySelectorAll('.lx-ac-name')].map(n=>n.textContent)})()");
    check("autocomplete popup lists matching commands", Array.isArray(ac) && ac[0] === "\\subsection", ac);
    await p.shot("03-autocomplete");
    await p.key("Enter");
    await p.type("Done");
    check("Enter accepts the completion and puts the caret inside the braces", (await p.text()).endsWith("\\subsection{Done}"), JSON.stringify((await p.text()).slice(-30)));
  }],

  ["find and replace with counts, and undo of Replace all", async (open) => {
    const p = await open("?new=latex");
    await p.waitFor("!!document.querySelector('#latex-view .lx-input')");
    await p.focusEditor();
    await p.key("h", CTRL);
    await p.waitFor("!document.querySelector('#latex-view .lx-findbar').hidden");
    check("Ctrl+H opens find with replace", await p.eval("document.querySelector('#latex-view .lx-findbar').classList.contains('lx-with-replace') && document.activeElement.classList.contains('lx-find-q')"));
    await p.type("section");
    const count = await p.eval("document.querySelector('#latex-view .lx-find-count').textContent");
    check("match count shown", /^1 of \d+$/.test(count), count);
    const n = +count.split(" of ")[1];
    check("matches are highlighted in the editor", (await p.eval("document.querySelectorAll('#latex-view .t-find').length")) >= Math.min(n, 3));
    await p.click("#latex-view .lx-opt[data-opt='word']");
    const wc = await p.eval("document.querySelector('#latex-view .lx-find-count').textContent");
    check("whole word excludes \\subsection-like hits", +wc.split(" of ")[1] <= n, wc);
    await p.shot("04-find");
    await p.click("#latex-view .lx-find-r");
    await p.type("chapter");
    await p.click("#latex-view [data-find='all']");
    const t = await p.text();
    check("Replace all replaced every match", !/\\section\{/.test(t) && /\\chapter\{Introduction\}/.test(t));
    const status = (await p.state()).status;
    check("status reports how many were replaced", /Replaced \d+ match/.test(status), status);
    await p.focusEditor();
    await p.key("z", CTRL);
    check("one Ctrl+Z restores the text", /\\section\{Introduction\}/.test(await p.text()));
    await p.key("Escape");
    check("Esc closes the find bar", await p.eval("document.querySelector('#latex-view .lx-findbar').hidden"));
    await p.key("g", CTRL);
    await p.type("20\n");
    check("Ctrl+G goes to a line", (await p.caretLine()) === 20, await p.caretLine());
  }],

  ["outline, preview click and problems navigate the source", async (open) => {
    const p = await open("?new=latex");
    await p.waitFor("!!document.querySelector('#latex-view .lx-input')");
    await p.rendered();
    if (await p.eval("document.querySelector('#latex-view .lx-side').hidden")) await p.click("#latex-view .lx-side-btn");
    const items = await p.eval("[...document.querySelectorAll('#latex-view .lx-ol-item')].map(b=>b.textContent)");
    check("outline lists the sections", JSON.stringify(items) === JSON.stringify(["1Introduction", "2Results", "3Conclusion"]), items);
    await p.clickText("#latex-view .lx-ol-item", "3Conclusion");
    await sleep(600);
    check("moving the caret scrolls the preview to the matching block", (await p.eval("document.querySelector('#latex-view .lx-preview').scrollTop")) > 100);
    await p.clickText("#latex-view .lx-ol-item", "2Results");
    const line = await p.caretLine();
    const src = (await p.text()).split("\n");
    check("clicking an outline entry jumps to its line", src[line - 1] === "\\section{Results}", { line, text: src[line - 1] });
    await p.click("#latex-view .lx-preview .ltx-tabular tr:nth-child(2)");
    const l2 = await p.caretLine();
    check("clicking a table row in the preview jumps to that row's source line", /^\s*Apples & 3 & 1\.20/.test((await p.text()).split("\n")[l2 - 1]), { l2 });
    await p.click("#latex-view .lx-preview h2.ltx-section:nth-of-type(1)");
    check("clicking a heading in the preview jumps to its \\section", /\\section\{/.test((await p.text()).split("\n")[(await p.caretLine()) - 1]));
    // broken document -> problems
    await p.setFiles([path.join(FIX, "broken.tex")]);
    await p.waitFor("__officeViewer.doc?.name === 'broken.tex'");
    const dlg = await p.state();
    if (dlg.modals.includes("dialog-modal")) await p.clickText("#dlg-actions button", "Don't save");
    await p.waitFor("__officeViewer.doc?.name === 'broken.tex' && document.querySelector('#latex-view .lx-input')?.value.includes('Broken things')");
    await p.rendered();
    const badge = await p.eval("document.querySelector('#latex-view .lx-prob-badge').textContent");
    check("problems badge counts errors and warnings", badge === "6", badge);
    check("problems button shows errors in red", await p.eval("document.querySelector('#latex-view .lx-prob-btn').classList.contains('has-err')"));
    await p.click("#latex-view .lx-prob-btn");
    const probs = await p.eval("[...document.querySelectorAll('#latex-view .lx-prob')].map(b=>b.querySelector('.lx-prob-line').textContent+' '+b.querySelector('.lx-prob-msg').textContent)");
    check("problems list shows line numbers and plain messages", probs.length === 6 && /^Line 5 Missing \}/.test(probs[0]) && probs.some((x) => /^Line 11 .*does not match/.test(x)), probs);
    check("error lines are marked in the gutter", (await p.eval("document.querySelectorAll('#latex-view .lx-l.lx-err').length")) >= 3);
    await p.shot("05-problems");
    await p.clickText("#latex-view .lx-prob", "Line 15");
    check("clicking a problem jumps to its line", (await p.caretLine()) === 15, await p.caretLine());
    check("broken document still renders the rest", /Last line\./.test(await p.preview()));
  }],

  ["templates, view modes (remembered), symbols, table picker, compile disabled offline", async (open) => {
    const p = await open("?new=latex");
    await p.waitFor("!!document.querySelector('#latex-view .lx-input')");
    await p.click("#latex-view [data-menu='tpl']");
    await p.shot("06-templates-menu");
    await p.click("#latex-view [data-tpl='math']");
    await p.waitFor("!document.getElementById('dialog-modal').hidden");
    check("template asks in the app's own dialog (no native confirm)", p.dialogs.length === 0);
    await p.clickText("#dlg-actions button", "Replace");
    await p.waitFor("document.querySelector('#latex-view .lx-input').value.includes('Homework 1')");
    await p.rendered();
    const h = await p.preview();
    check("math homework template renders problems and proofs", /Problem 1\./.test(h.replace(/<[^>]+>/g, "")) && /ltx-qed/.test(h) && /<mtable/.test(h));
    await p.shot("07-math-template");
    await p.focusEditor();
    await p.key("z", CTRL);
    check("Ctrl+Z brings the previous document back", /An Example Article/.test(await p.text()));
    // symbols palette: outside math the command is wrapped in $...$
    await p.key("End", CTRL);
    await p.click("#latex-view [data-menu='sym']");
    await p.click("#latex-view [data-sym='\\\\alpha']");
    check("symbol inserted as $\\alpha$ outside math", (await p.text()).endsWith("$\\alpha$"), (await p.text()).slice(-20));
    await p.shot("08-symbols");
    await p.click("#latex-view [data-menu='table']");
    await p.hover("#latex-view .lx-tp-grid [data-r='3'][data-c='2']");
    check("table picker shows the size", (await p.eval("document.querySelector('#latex-view .lx-tp-label').textContent")) === "3 × 2 table");
    await p.click("#latex-view .lx-tp-grid [data-r='3'][data-c='2']");
    const t = await p.text();
    check("table snippet inserted", /\\begin\{tabular\}\{\|c\|c\|\}/.test(t) && (t.match(/\\\\\n/g) || []).length >= 3);
    // compile PDF is disabled without the backend, with an explanation
    const comp = await p.eval("(()=>{const b=document.querySelector('#latex-view .lx-compile'); return {d:b.disabled, t:b.title}})()");
    check("Compile PDF disabled offline with a helpful tooltip", comp.d && /TeX engine/.test(comp.t) && /Print/.test(comp.t), comp);
    // view modes
    await p.click("#latex-view [data-mode='preview']");
    check("Preview mode hides the editor", await p.eval("document.querySelector('#latex-view .lx-editor-pane').hidden && !document.querySelector('#latex-view .lx-preview-pane').hidden"));
    await p.shot("09-preview-mode");
    await p.click("#latex-view [data-mode='source']");
    check("Source mode hides the preview", await p.eval("document.querySelector('#latex-view .lx-preview-pane').hidden && !document.querySelector('#latex-view .lx-editor-pane').hidden"));
    const p2 = await open("?new=latex");
    await p2.waitFor("!!document.querySelector('#latex-view .lx-input')");
    check("view mode is remembered", (await p2.eval("document.querySelector('#latex-view').dataset.lxView")) === "source");
    await p2.click("#latex-view [data-mode='split']");
  }],

  ["open a .tex fixture, Save (.tex bytes), Save As .docx and .html offline", async (open) => {
    const p = await open("");
    await p.setFiles([path.join(FIX, "sample.tex")]);
    await p.waitFor("__officeViewer.doc?.kind === 'latex' && !!document.querySelector('#latex-view .lx-input')");
    await p.rendered();
    const h = await p.preview();
    check("fixture renders: title, theorem, table, figure placeholder, bibliography", /Fourier Series in Practice/.test(h) && /Theorem 1\.1/.test(h.replace(/<[^>]+>/g, "")) && /ltx-tabular/.test(h) && /square-wave\.png/.test(h) && /ltx-bibitem/.test(h));
    const plain = await p.eval("document.querySelector('#latex-view .lx-preview-inner').textContent");
    check("cross references resolved", /See any textbook, for example\u00A0\[1\]/.test(plain) && /Table\u00A01 and Figure\u00A01 summarize Section\u00A02/.test(plain), plain.slice(0, 200));
    check("no problems in a valid document", (await p.eval("document.querySelector('#latex-view .lx-prob-badge').textContent")) === "0");
    await p.shot("10-sample-split");
    await p.focusEditor();
    await p.key("End", CTRL);
    await p.type("% edited\n");
    await p.key("s", CTRL);
    // a file opened through the file input has no handle: Save asks where to save (Save As dialog)
    await p.waitFor("window.__written.length > 0 || !document.getElementById('saveas-modal').hidden");
    if (await p.eval("!document.getElementById('saveas-modal').hidden")) {
      check("Save As proposes .tex for a LaTeX document", (await p.eval("document.getElementById('sa-format').value")) === "tex" && (await p.eval("document.getElementById('sa-name').value")) === "sample.tex");
      await p.click("#sa-ok");
    }
    const [w] = await p.written();
    const expected = fs.readFileSync(path.join(FIX, "sample.tex"), "utf8") + "% edited\n";
    check("Save writes the .tex source bytes", w && w.name === "sample.tex" && w.bytes.toString("utf8") === expected, w && { name: w.name, len: w.bytes.length, exp: expected.length });
    check("saving clears the dirty flag", (await p.state()).dirty === false);
    // Save As .docx (built-in writer)
    await p.key("s", CTRL | SHIFT);
    await p.waitFor("!document.getElementById('saveas-modal').hidden");
    const opts = await p.eval("[...document.querySelectorAll('#sa-format option')].map(o=>o.value)");
    check("Save As offers tex, docx, html, md, txt offline", ["tex", "docx", "html", "md", "txt"].every((f) => opts.includes(f)), opts);
    await p.choose("#sa-format", "docx");
    await p.eval("document.getElementById('sa-name').value='sample.docx'");
    await p.click("#sa-ok");
    const [d] = await p.written();
    check("Save As .docx writes a zip package", d && d.name === "sample.docx" && d.bytes[0] === 0x50 && d.bytes[1] === 0x4b, d && d.name);
    if (d) fs.writeFileSync(path.join(PROFILE, "sample-out.docx"), d.bytes);
    await sleep(300);
    await p.key("s", CTRL | SHIFT);
    await p.waitFor("!document.getElementById('saveas-modal').hidden");
    await p.choose("#sa-format", "html");
    await p.eval("document.getElementById('sa-name').value='sample.html'");
    await p.click("#sa-ok");
    const [hw] = await p.written();
    const html = hw ? hw.bytes.toString("utf8") : "";
    check("Save As .html writes a standalone page with MathML", /^<!doctype html>/.test(html) && /<math display="block"/.test(html) && html.includes(".ltx-doc{"), html.slice(0, 80));
    if (hw) fs.writeFileSync(path.join(PROFILE, "sample-out.html"), hw.bytes);
  }],

  ["print shows only the rendered preview", async (open) => {
    const p = await open("?new=latex");
    await p.waitFor("!!document.querySelector('#latex-view .lx-input')");
    await p.click("#latex-view [data-mode='source']");
    await p.send("Emulation.setEmulatedMedia", { media: "print" });
    await p.eval("window.dispatchEvent(new Event('beforeprint'))");
    await sleep(200);
    const vis = await p.eval(`(()=>{const v=(s)=>{const e=document.querySelector(s); return !!e && getComputedStyle(e).display!=='none' && e.getBoundingClientRect().height>0}; return {editor:v('#latex-view .lx-editor-pane'), toolbar:v('#latex-view .lx-toolbar'), preview:v('#latex-view .lx-preview .ltx-doc')}})()`);
    check("print media: editor and toolbar hidden, preview visible even in Source mode", !vis.editor && !vis.toolbar && vis.preview, vis);
    await p.shot("11-print-media");
    await p.send("Emulation.setEmulatedMedia", { media: "" });
    await p.click("#latex-view [data-mode='split']");
  }],

  ["beamer and CV templates from the New menu URL", async (open) => {
    const p = await open("?new=latex&template=beamer");
    await p.waitFor("!!document.querySelector('#latex-view .lx-input')");
    await p.rendered();
    const frames = await p.eval("document.querySelectorAll('#latex-view .ltx-frame').length");
    check("beamer slides render as frames", frames >= 5, frames);
    check("beamer outline lists slides", (await p.eval("[...document.querySelectorAll('#latex-view .lx-ol-item')].map(b=>b.textContent)")).includes("Why this matters"));
    await p.shot("12-beamer");
    const q = await open("?new=latex&template=cv");
    await q.waitFor("!!document.querySelector('#latex-view .lx-input')");
    await q.rendered();
    check("CV template renders without problems", (await q.eval("document.querySelector('#latex-view .lx-prob-badge').textContent")) === "0");
    await q.shot("13-cv");
  }],

  ["Compile PDF flow (with a stand-in compiler): PDF pane, back to preview, TeX errors in Problems", async (open) => {
    const p = await open("?new=latex");
    await p.waitFor("!!document.querySelector('#latex-view .lx-input')");
    // a second LatexView on its own root, with a fake backend compiler
    await p.eval(`(async()=>{
      const { LatexView } = await import(chrome.runtime.getURL('viewer/latex.js'));
      document.getElementById('latex-view').hidden = true;
      const root = document.createElement('div');
      root.id = 'lx-test-root';
      root.style.cssText = 'position:fixed;inset:46px 0 24px 0;display:flex;flex-direction:column;background:#fff;z-index:5';
      document.body.appendChild(root);
      window.__fail = false;
      window.__lv = new LatexView({ root, text: '\\\\documentclass{article}\\n\\\\begin{document}\\nHi \\\\oops\\n\\\\end{document}\\n',
        onStatus: (m) => { window.__lvStatus = m; }, onDirty: () => {}, onStats: () => {}, confirm: async () => true,
        canCompile: () => true,
        compile: async (src) => { await new Promise(r => setTimeout(r, 50)); if (window.__fail) throw new Error("! Undefined control sequence.\\nl.3 Hi \\\\oops"); return new Blob(['%PDF-1.4\\n%%EOF\\n'], { type: 'application/pdf' }); } });
      window.__lv.render();
    })()`);
    const en = await p.eval("!document.querySelector('#lx-test-root .lx-compile').disabled");
    check("Compile PDF enabled when a TeX backend is available", en);
    await p.click("#lx-test-root .lx-compile");
    await p.waitFor("!document.querySelector('#lx-test-root .lx-pdf').hidden", 5000);
    check("compiled PDF shown in the preview pane from a blob: URL", await p.eval("document.querySelector('#lx-test-root .lx-pdf').src.startsWith('blob:') && !document.querySelector('#lx-test-root .lx-pdfbar').hidden && document.querySelector('#lx-test-root .lx-preview').hidden"));
    await p.shot("15-compiled-pdf");
    await p.click("#lx-test-root .lx-pdf-back");
    check("Back to live preview returns to the HTML preview", await p.eval("document.querySelector('#lx-test-root .lx-pdf').hidden && !document.querySelector('#lx-test-root .lx-preview').hidden"));
    await p.eval("window.__fail = true");
    await p.click("#lx-test-root .lx-compile");
    await p.waitFor("document.querySelectorAll('#lx-test-root .lx-prob').length > 1 && !document.querySelector('#lx-test-root .lx-compile').classList.contains('busy')");
    const pr = await p.eval("[...document.querySelectorAll('#lx-test-root .lx-prob')].map(b=>b.textContent)");
    check("TeX errors from the compiler appear in Problems with their line", pr.some((x) => /PDF · Line 3.*Undefined control sequence/.test(x)), pr);
    check("status explains the compile failure", /could not be compiled/.test(await p.eval("window.__lvStatus")));
    await p.eval("window.__lv.destroy()");
    check("destroy() empties the root", await p.eval("document.getElementById('lx-test-root').childElementCount === 0"));
  }],

  ["a 1 MB document opens and stays responsive", async (open) => {
    const p = await open("");
    const base = fs.readFileSync(path.join(FIX, "sample.tex"), "utf8");
    const head = base.slice(0, base.indexOf("\\section{Introduction}"));
    const body = base.slice(base.indexOf("\\section{Introduction}"), base.indexOf("\\begin{thebibliography}"));
    let big = head, n = 0;
    while (big.length < 1_000_000) big += body.replace(/\\(label|ref|eqref)\{([^}]+)\}/g, (m, c, k) => `\\${c}{${k}-${n}}`), n++;
    big += base.slice(base.indexOf("\\begin{thebibliography}"));
    const file = path.join(PROFILE, "big.tex");
    fs.writeFileSync(file, big);
    const t0 = Date.now();
    await p.setFiles([file]);
    await p.waitFor("__officeViewer.doc?.name === 'big.tex' && document.querySelectorAll('#latex-view .lx-l').length > 10000 && document.querySelectorAll('#latex-view .ltx-section').length > 100", 30000);
    const openMs = Date.now() - t0;
    check(`1 MB file opens with highlighted source and preview (${openMs} ms)`, openMs < 8000 * PERF, openMs);
    const lines = await p.eval("document.querySelectorAll('#latex-view .lx-l').length");
    check("every source line has a line element", lines === big.split("\n").length, { lines, expected: big.split("\n").length });
    await p.focusEditor();
    const typing = await p.eval(`(()=>{const ta=document.querySelector('#latex-view .lx-input'); const i=ta.value.indexOf('Convergence'); ta.setSelectionRange(i,i); const t=performance.now(); for (let k=0;k<5;k++) document.execCommand('insertText', false, 'x'); return (performance.now()-t)/5})()`);
    check(`typing one character in a 1 MB document takes ${typing.toFixed(1)} ms`, typing < 120, typing);
    await sleep(1800);
    check("preview re-rendered after typing", /xxxxxConvergence/.test(await p.eval("document.querySelector('#latex-view .lx-preview-inner').textContent.slice(0, 5000)")));
    await p.shot("14-big");
    // where does the time go? (a fresh view on the same text)
    const t = await p.eval(`(async()=>{
      const { LatexView } = await import(chrome.runtime.getURL('viewer/latex.js'));
      const text = document.querySelector('#latex-view .lx-input').value;
      const root = document.createElement('div');
      root.style.cssText = 'position:fixed;inset:0;display:flex;flex-direction:column;visibility:hidden';
      document.body.appendChild(root);
      const v = new LatexView({ root, text, onStatus(){}, onDirty(){}, onStats(){}, confirm: async()=>true, canCompile: ()=>false, compile: async()=>null });
      const t0 = performance.now(); v.render(); const t1 = performance.now();
      root.getBoundingClientRect(); document.body.offsetHeight; const t2 = performance.now();
      v.lastRenderedText = null; v.renderNow(); const t3 = performance.now();
      v.fullHighlight(); root.getBoundingClientRect(); const t4 = performance.now();
      v.destroy(); root.remove();
      return { render: Math.round(t1 - t0), layout: Math.round(t2 - t1), preview: Math.round(t3 - t2), highlight: Math.round(t4 - t3) };
    })()`);
    console.log(`      timing (ms): view render ${t.render} (incl. first preview), layout ${t.layout}, preview re-render ${t.preview}`);
    check(`1 MB: the editor is ready in under ${3.5 * PERF} s (preview follows)`, t.render + t.layout < 3500 * PERF, t);
    check(`1 MB: preview re-render under ${2 * PERF} s`, t.preview < 2000 * PERF, t);
  }],
];

// ---------------------------------------------------------------- main

async function main() {
  // A reused profile keeps an idle service worker that /json/list does not show: start clean.
  try { fs.rmSync(PROFILE, { recursive: true, force: true }); } catch { /* still locked by a closing browser */ }
  fs.mkdirSync(PROFILE, { recursive: true });
  // wait until a previous run's browser released the debugging port
  for (let i = 0; i < 40; i++) {
    try { await fetch(`http://127.0.0.1:${PORT}/json/version`); await sleep(250); } catch { break; }
  }
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

    {
      const p = await open("");
      await p.eval(`chrome.storage.local.set({settings: {backend_enabled: false, backend_url: "http://127.0.0.1:47831"}})`);
      await p.eval("localStorage.clear()");
      await browserConn.send("Target.closeTarget", { targetId: p.targetId });
    }
    const only = process.argv[2];
    for (const [name, fn] of TESTS) {
      if (only && !name.includes(only)) continue;
      console.log(`\n# ${name}`);
      const pages = [];
      try {
        await fn(async (q) => { const p = await open(q); pages.push(p); return p; });
      } catch (e) {
        check(`${name}: ran to the end`, false, e.stack || String(e));
      }
      for (const p of pages) {
        if (p.logs.length) check(`${name}: no page exceptions`, false, p.logs.join("\n"));
        const ce = p.consoleErrors.filter((x) => !/favicon|ERR_CONNECTION_REFUSED|Failed to fetch|net::/i.test(x));
        if (ce.length) check(`${name}: no console errors`, false, ce.join("\n"));
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

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
