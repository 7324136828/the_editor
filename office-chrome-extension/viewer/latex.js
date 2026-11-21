/**
 * LatexView: LaTeX source editor + live preview (owned by the latex stream).
 *
 * Interface used by viewer/viewer.js:
 *   new LatexView({ root, text, onStatus, onDirty, onStats, confirm, compile, canCompile })
 *     root       #latex-view (inside #stage, a full-bleed flex column in "app-mode");
 *                the view builds ALL its UI inside root and removes it on destroy().
 *     text       initial LaTeX source
 *     onStatus   (message) status-bar text
 *     onDirty    (true) after every edit
 *     onStats    () call after edits (the viewer debounces)
 *     confirm    async (message, {title, ok}) => boolean
 *     compile    async (source) => Blob (application/pdf); throws Error whose
 *                message contains the TeX error/log excerpt
 *     canCompile () => boolean  (backend up AND a TeX engine is installed)
 *   render()          build the UI
 *   destroy()         remove listeners (incl. window/document ones), revoke object URLs, clear root
 *   focus()           focus the source editor
 *   getText()         current source
 *   setText(text)     replace the source (undoable)
 *   refreshCompile()  re-evaluate canCompile() (called when backend health changes)
 *   beforePrint()     make the printed output the rendered preview
 *   commitEdit()      no-op
 *
 * Extra (tests / automation): viewMode, setViewMode(m), renderNow(), errors, outline,
 *   jumpToLine(n), openFind(replace), insertSnippet(id).
 *
 * Editor design: a transparent <textarea> (native caret, selection, IME, spell
 * check off, NATIVE UNDO) layered over a syntax-highlighted <pre> made of one
 * <div> per line (CSS counters draw the line numbers). Only changed lines are
 * re-rendered. Every programmatic edit goes through execCommand("insertText"),
 * so Ctrl+Z / Ctrl+Y keep working.
 */

import {
  latexToHtml, newLatexDocument, TEMPLATES, PREVIEW_CSS, COMMAND_COMPLETIONS, ENVIRONMENT_COMPLETIONS,
  SYMBOL_GROUPS, userCommandNames, latexKeys,
} from "../lib/latex.js";

const LS_MODE = "officeEditor.latex.viewMode";
const LS_SIDE = "officeEditor.latex.sidebar";
const LS_SPLIT = "officeEditor.latex.split";
const VERB_ENVS = new Set(["verbatim", "verbatim*", "Verbatim", "lstlisting", "minted", "comment"]);
const MATH_ENVS = new Set(["equation", "equation*", "align", "align*", "gather", "gather*", "multline", "multline*", "eqnarray", "eqnarray*", "displaymath", "flalign", "flalign*", "alignat", "alignat*"]);
const SECTION_CMDS = new Set(["part", "chapter", "section", "subsection", "subsubsection", "paragraph", "subparagraph", "frametitle"]);
const LIST_ENVS = new Set(["itemize", "enumerate", "description"]);
const PAIRS = { "{": "}", "[": "]", "(": ")", $: "$" };
const CARET = "\u0001";   // caret marker in snippets
const SEL = "\u0002";     // selected text placeholder in snippets

const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const lsGet = (k, d) => { try { return localStorage.getItem(k) ?? d; } catch { return d; } };
const lsSet = (k, v) => { try { localStorage.setItem(k, v); } catch { /* storage unavailable */ } };

const SNIPPETS = {
  section: { label: "Section", text: `\\section{${SEL}${CARET}}`, block: true },
  subsection: { label: "Subsection", text: `\\subsection{${SEL}${CARET}}`, block: true },
  bold: { label: "Bold", text: `\\textbf{${SEL}${CARET}}`, key: "Ctrl+B" },
  italic: { label: "Italic", text: `\\textit{${SEL}${CARET}}`, key: "Ctrl+I" },
  itemize: { label: "Bulleted list", text: `\\begin{itemize}\n  \\item ${SEL}${CARET}\n  \\item \n\\end{itemize}`, block: true },
  enumerate: { label: "Numbered list", text: `\\begin{enumerate}\n  \\item ${SEL}${CARET}\n  \\item \n\\end{enumerate}`, block: true },
  figure: { label: "Figure", text: `\\begin{figure}[h]\n  \\centering\n  \\includegraphics[width=0.8\\textwidth]{${CARET}image.png}\n  \\caption{${SEL || ""}Caption}\n  \\label{fig:label}\n\\end{figure}`, block: true },
  equation: { label: "Equation (numbered)", text: `\\begin{equation}\n  ${SEL}${CARET}\n  \\label{eq:label}\n\\end{equation}`, block: true },
  inlinemath: { label: "Inline math", text: `$${SEL}${CARET}$` },
  footnote: { label: "Footnote", text: `\\footnote{${SEL}${CARET}}` },
  link: { label: "Link", text: `\\href{https://${CARET}}{${SEL}}` },
  cite: { label: "Citation", text: `\\cite{${CARET}}` },
  label: { label: "Label", text: `\\label{${CARET}}` },
  ref: { label: "Reference", text: `\\ref{${CARET}}` },
};

const ICONS = {
  doc: '<path d="M4 1.8h5.2L12.5 5v9.2H4z"/><path d="M9 1.8V5h3.5"/>',
  plus: '<path d="M8 3v10M3 8h10"/>',
  sigma: '<path d="M12 3H4.5L8.5 8l-4 5H12"/>',
  table: '<rect x="1.8" y="2.8" width="12.4" height="10.4" rx="1"/><path d="M1.8 6.3h12.4M1.8 9.8h12.4M6 2.8v10.4M10 2.8v10.4"/>',
  search: '<circle cx="7" cy="7" r="4.3"/><path d="m10.3 10.3 3.6 3.6"/>',
  side: '<rect x="1.8" y="2.5" width="12.4" height="11" rx="1.3"/><path d="M5.8 2.5v11"/>',
  pdf: '<path d="M4 1.8h5.2L12.5 5v9.2H4z"/><path d="M6 9.5h4M6 11.8h2.5M6 7.2h4"/>',
  warn: '<path d="M8 2 14.5 13.5h-13z"/><path d="M8 6.5v3.2M8 11.6v.1"/>',
  ok: '<circle cx="8" cy="8" r="6"/><path d="m5.3 8.2 1.9 1.9 3.6-3.8"/>',
  err: '<circle cx="8" cy="8" r="6"/><path d="M5.8 5.8l4.4 4.4M10.2 5.8l-4.4 4.4"/>',
  back: '<path d="M10 3.5 5.5 8l4.5 4.5"/>',
  close: '<path d="M4 4l8 8M12 4l-8 8"/>',
  up: '<path d="M4 10l4-4 4 4"/>',
  down: '<path d="M4 6l4 4 4-4"/>',
};
const icon = (n) => `<svg class="lx-ic" viewBox="0 0 16 16" width="16" height="16" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round">${ICONS[n] || ""}</svg>`;

// ---------------------------------------------------------------------------
// syntax highlighting (one line at a time, with a small carried state)
// ---------------------------------------------------------------------------

/**
 * Highlight one line. state: "" | "m" (display math) | "v:<env>" (verbatim-like).
 * Returns { segs:[[cls, text]], out: state at the end of the line }.
 */
function highlightLine(text, state) {
  const segs = [];
  const push = (cls, t) => {
    if (!t) return;
    const last = segs[segs.length - 1];
    if (last && last[0] === cls) last[1] += t;
    else segs.push([cls, t]);
  };
  let i = 0;
  const n = text.length;
  let st = state;
  if (st.startsWith("v:")) {
    const env = st.slice(2);
    const endMark = `\\end{${env}}`;
    const e = text.indexOf(endMark);
    if (e < 0) { push("t-verb", text); return { segs, out: st }; }
    push("t-verb", text.slice(0, e));
    push("t-kw", "\\end");
    push("t-br", "{");
    push("t-env", env);
    push("t-br", "}");
    i = e + endMark.length;
    st = "";
  }
  let inline = false; // $...$ or \( \) within this line
  const mathCls = () => (st === "m" || inline);
  while (i < n) {
    const c = text[i];
    if (c === "%") { push("t-com", text.slice(i)); break; }
    if (c === "\\") {
      const m = /^\\([A-Za-z@]+\*?|.)?/.exec(text.slice(i, i + 64));
      const tok = m[0];
      const name = m[1] || "";
      if (name === "begin" || name === "end") {
        const em = /^\{([^}]*)\}/.exec(text.slice(i + tok.length));
        push("t-kw", tok);
        if (em) {
          push("t-br", "{");
          push("t-env", em[1]);
          push("t-br", "}");
          i += tok.length + em[0].length;
          const env = em[1].trim();
          if (name === "begin" && VERB_ENVS.has(env)) {
            push("t-verb", text.slice(i));
            return { segs, out: `v:${env}` };
          }
          if (MATH_ENVS.has(env)) st = name === "begin" ? "m" : "";
          continue;
        }
        i += tok.length;
        continue;
      }
      if (name === "[" || name === "]") { st = name === "[" ? "m" : ""; push("t-mdel", tok); i += tok.length; continue; }
      if (name === "(" || name === ")") { inline = name === "("; push("t-mdel", tok); i += tok.length; continue; }
      if (name === "verb" || name === "verb*") {
        const d = text[i + tok.length];
        const e = d ? text.indexOf(d, i + tok.length + 1) : -1;
        const end = e < 0 ? n : e + 1;
        push("t-kw", tok);
        push("t-verb", text.slice(i + tok.length, end));
        i = end;
        continue;
      }
      const base = name.replace(/\*$/, "");
      push(mathCls() ? "t-mcmd" : SECTION_CMDS.has(base) ? "t-sec" : /^[A-Za-z@]/.test(name) ? "t-cmd" : "t-esc", tok);
      i += tok.length;
      continue;
    }
    if (c === "$") {
      if (text[i + 1] === "$") { st = st === "m" ? "" : "m"; push("t-mdel", "$$"); i += 2; continue; }
      inline = !inline;
      push("t-mdel", "$");
      i++;
      continue;
    }
    if (c === "{" || c === "}" || c === "[" || c === "]") { push("t-br", c); i++; continue; }
    if (c === "&" || c === "~" || c === "^" || c === "_" || c === "#") { push(mathCls() ? "t-math t-sp" : "t-sp", c); i++; continue; }
    let j = i + 1;
    while (j < n && !"%\\$&~^_#{}[]".includes(text[j])) j++;
    push(mathCls() ? "t-math" : "", text.slice(i, j));
    i = j;
  }
  return { segs, out: st };
}

/** Apply decorations [{s, e, cls}] (columns within the line) to segments; returns HTML. */
function segsToHtml(segs, decos) {
  if (!decos || !decos.length) {
    let h = "";
    for (const [cls, t] of segs) h += cls ? `<span class="${cls}">${esc(t)}</span>` : esc(t);
    return h;
  }
  // split at decoration boundaries
  const cuts = new Set();
  for (const d of decos) { cuts.add(d.s); cuts.add(d.e); }
  let h = "";
  let pos = 0;
  for (const [cls, t] of segs) {
    let k = 0;
    while (k < t.length) {
      let next = t.length;
      for (const c of cuts) if (c > pos + k && c < pos + next) next = c - pos;
      const piece = t.slice(k, next);
      const extra = decos.filter((d) => d.s <= pos + k && d.e >= pos + next && d.e > d.s).map((d) => d.cls);
      const all = [cls, ...extra].filter(Boolean).join(" ");
      h += all ? `<span class="${all}">${esc(piece)}</span>` : esc(piece);
      k = next;
    }
    pos += t.length;
  }
  return h;
}

// ---------------------------------------------------------------------------
// the view
// ---------------------------------------------------------------------------

export class LatexView {
  constructor(opts) {
    this.o = opts;
    this.root = opts.root;
    this.ta = null;
    this.lines = [];
    this.states = [];        // state at the start of each line
    this.lineEls = [];
    this.errors = [];
    this.compileErrors = [];
    this.outline = [];
    this.viewMode = lsGet(LS_MODE, "split");
    if (!["source", "split", "preview"].includes(this.viewMode)) this.viewMode = "split";
    this.sideOpen = lsGet(LS_SIDE, "1") === "1";
    this.sideTab = "outline";
    this.split = Math.min(0.8, Math.max(0.2, Number(lsGet(LS_SPLIT, "0.5")) || 0.5));
    this.cleanups = [];
    this.timers = new Set();
    this.objectUrls = [];
    this.renderTimer = 0;
    this.syncTimer = 0;
    this.curLine = -1;
    this.decoLines = new Set();
    this.find = { open: false, query: "", replace: "", caseSens: false, regex: false, word: false, matches: [], cur: -1, error: "" };
    this.ac = { open: false, items: [], sel: 0, start: 0, kind: "" };
    this.lastRenderedText = null;
    this.compiling = false;
    this.showingPdf = false;
    this.destroyed = false;
  }

  // ---------------------------------------------------------------- lifecycle

  render() {
    const root = this.root;
    root.innerHTML = "";
    root.classList.add("lx-root");
    root.innerHTML = this.template();
    const $ = (s) => root.querySelector(s);
    this.el = {
      style: $(".lx-style"), toolbar: $(".lx-toolbar"), main: $(".lx-main"), side: $(".lx-side"),
      outline: $(".lx-outline"), problems: $(".lx-problems"), editorPane: $(".lx-editor-pane"), code: $(".lx-code"),
      hl: $(".lx-hl"), ta: $(".lx-input"), ac: $(".lx-ac"), findbar: $(".lx-findbar"), gotobar: $(".lx-gotobar"),
      resizer: $(".lx-resizer"), previewPane: $(".lx-preview-pane"), preview: $(".lx-preview"), page: $(".lx-preview-inner"),
      pdfbar: $(".lx-pdfbar"), pdf: $(".lx-pdf"), compile: $(".lx-compile"), probBtn: $(".lx-prob-btn"),
      probBadge: $(".lx-prob-badge"), tabCount: $(".lx-tab-count"), sideBtn: $(".lx-side-btn"),
    };
    this.el.style.textContent = PREVIEW_CSS;
    this.ta = this.el.ta;
    this.ta.value = this.o.text || "";
    this.ta.setSelectionRange(0, 0);
    this.ta.scrollTop = 0;
    this.bind();
    this.applyViewMode();
    this.applySide();
    this.refreshCompile();
    this.fullHighlight();
    this.updateCaretLine();
    if (this.ta.value.length > 200000) {
      // big documents: show the editor first, build the preview right after
      this.el.page.innerHTML = '<div class="lx-loading">Rendering the preview…</div>';
      this.later(() => this.renderNow(), 30);
    } else this.renderNow();
  }

  template() {
    const tplItems = TEMPLATES.map((t) => `<button role="menuitem" data-tpl="${esc(t.id)}"><span class="lx-mi-title">${esc(t.label)}</span><span class="lx-mi-desc">${esc(t.description || "")}</span></button>`).join("");
    const insItems = ["section", "subsection", "bold", "italic", "itemize", "enumerate", "figure", "equation", "inlinemath", "footnote", "link", "cite", "label", "ref"]
      .map((id) => `<button role="menuitem" data-snippet="${id}">${esc(SNIPPETS[id].label)}${SNIPPETS[id].key ? `<kbd>${SNIPPETS[id].key}</kbd>` : ""}</button>`).join("");
    const symTabs = SYMBOL_GROUPS.map((g, k) => `<button class="lx-symtab${k === 0 ? " on" : ""}" data-symtab="${k}" type="button">${esc(g.label)}</button>`).join("");
    const symGrids = SYMBOL_GROUPS.map((g, k) => `<div class="lx-symgrid" data-symgrid="${k}"${k ? " hidden" : ""}>${g.items.map(([cmd, ch]) => `<button type="button" data-sym="${esc(cmd)}" title="${esc(cmd)}">${esc(ch)}</button>`).join("")}</div>`).join("");
    let grid = "";
    for (let r = 0; r < 8; r++) for (let c = 0; c < 8; c++) grid += `<span data-r="${r + 1}" data-c="${c + 1}"></span>`;
    return `<style class="lx-style"></style>
<div class="lx-toolbar" role="toolbar" aria-label="LaTeX tools">
  <div class="lx-menu-wrap"><button class="tb lx-menu-btn" data-menu="tpl" aria-haspopup="menu" aria-expanded="false" title="Start from a template">${icon("doc")}<span>Templates</span><span class="caret">▾</span></button>
    <div class="lx-menu lx-menu-wide" data-menu-panel="tpl" role="menu" hidden><div class="lx-menu-label">Replace the document with…</div>${tplItems}</div></div>
  <div class="lx-menu-wrap"><button class="tb lx-menu-btn" data-menu="insert" aria-haspopup="menu" aria-expanded="false" title="Insert a LaTeX snippet">${icon("plus")}<span>Insert</span><span class="caret">▾</span></button>
    <div class="lx-menu" data-menu-panel="insert" role="menu" hidden>${insItems}</div></div>
  <div class="lx-menu-wrap"><button class="tb lx-menu-btn" data-menu="table" aria-haspopup="true" aria-expanded="false" title="Insert a table">${icon("table")}<span>Table</span><span class="caret">▾</span></button>
    <div class="lx-menu lx-table-picker" data-menu-panel="table" hidden><div class="lx-tp-grid">${grid}</div><div class="lx-tp-label">Pick a size</div></div></div>
  <div class="lx-menu-wrap"><button class="tb lx-menu-btn" data-menu="sym" aria-haspopup="true" aria-expanded="false" title="Insert a math symbol">${icon("sigma")}<span>Symbols</span><span class="caret">▾</span></button>
    <div class="lx-menu lx-sym-panel" data-menu-panel="sym" hidden><div class="lx-symtabs">${symTabs}</div>${symGrids}<div class="lx-sym-hint">Symbols outside math are wrapped in $…$ automatically.</div></div></div>
  <span class="lx-sep"></span>
  <button class="tb icon lx-fmt" data-snippet="bold" title="Bold (Ctrl+B)"><b>B</b></button>
  <button class="tb icon lx-fmt" data-snippet="italic" title="Italic (Ctrl+I)"><i>I</i></button>
  <button class="tb icon lx-fmt" data-snippet="inlinemath" title="Inline math $…$"><span class="lx-mathbtn">$x$</span></button>
  <span class="lx-sep"></span>
  <button class="tb lx-find-btn" title="Find and replace (Ctrl+F, Ctrl+H)">${icon("search")}<span>Find</span></button>
  <div class="spacer"></div>
  <button class="tb lx-prob-btn" title="Show problems in the document">${icon("warn")}<span class="lx-prob-badge">0</span></button>
  <div class="lx-seg" role="group" aria-label="View">
    <button type="button" data-mode="source" title="Only the source">Source</button><button type="button" data-mode="split" title="Source and preview side by side">Split</button><button type="button" data-mode="preview" title="Only the preview">Preview</button>
  </div>
  <button class="tb toggle lx-side-btn" aria-pressed="false" title="Show or hide the outline and problems">${icon("side")}<span>Outline</span></button>
  <button class="tb lx-compile" title="">${icon("pdf")}<span>Compile PDF</span></button>
</div>
<div class="lx-main">
  <aside class="lx-side" aria-label="Outline and problems">
    <div class="lx-tabs" role="tablist"><button type="button" role="tab" data-tab="outline">Outline</button><button type="button" role="tab" data-tab="problems">Problems <span class="lx-tab-count"></span></button></div>
    <div class="lx-outline" role="tabpanel"></div>
    <div class="lx-problems" role="tabpanel" hidden></div>
  </aside>
  <section class="lx-editor-pane">
    <div class="lx-findbar" hidden>
      <div class="lx-find-row">
        <input class="lx-find-q" type="text" placeholder="Find" aria-label="Find" spellcheck="false">
        <button type="button" class="lx-opt" data-opt="caseSens" title="Match case" aria-pressed="false">Aa</button>
        <button type="button" class="lx-opt" data-opt="word" title="Whole word" aria-pressed="false"><u>W</u></button>
        <button type="button" class="lx-opt" data-opt="regex" title="Regular expression" aria-pressed="false">.*</button>
        <span class="lx-find-count" aria-live="polite"></span>
        <button type="button" class="lx-fb-btn" data-find="prev" title="Previous match (Shift+Enter)">${icon("up")}</button>
        <button type="button" class="lx-fb-btn" data-find="next" title="Next match (Enter)">${icon("down")}</button>
        <button type="button" class="lx-fb-btn" data-find="close" title="Close (Esc)">${icon("close")}</button>
      </div>
      <div class="lx-find-row lx-replace-row">
        <input class="lx-find-r" type="text" placeholder="Replace with" aria-label="Replace with" spellcheck="false">
        <button type="button" class="lx-fb-text" data-find="replace">Replace</button>
        <button type="button" class="lx-fb-text" data-find="all">Replace all</button>
      </div>
    </div>
    <div class="lx-gotobar" hidden>
      <label>Go to line <input class="lx-goto-in" type="number" min="1" aria-label="Line number"></label>
      <span class="lx-goto-hint"></span>
      <button type="button" class="lx-fb-btn" data-goto="close" title="Close (Esc)">${icon("close")}</button>
    </div>
    <div class="lx-code">
      <pre class="lx-hl" aria-hidden="true"></pre>
      <textarea class="lx-input" spellcheck="false" autocapitalize="off" autocomplete="off" wrap="soft" aria-label="LaTeX source"></textarea>
      <div class="lx-ac" role="listbox" hidden></div>
    </div>
  </section>
  <div class="lx-resizer" role="separator" aria-orientation="vertical" title="Drag to resize"></div>
  <section class="lx-preview-pane">
    <div class="lx-pdfbar" hidden><span class="lx-pdf-title">${icon("pdf")} Compiled PDF</span><button type="button" class="tb lx-pdf-back">${icon("back")}<span>Back to live preview</span></button></div>
    <div class="lx-preview" tabindex="-1" aria-label="Live preview"><div class="lx-preview-inner"></div></div>
    <iframe class="lx-pdf" title="Compiled PDF" hidden></iframe>
  </section>
</div>`;
  }

  on(target, type, fn, opts) {
    target.addEventListener(type, fn, opts);
    this.cleanups.push(() => target.removeEventListener(type, fn, opts));
  }
  later(fn, ms) {
    const id = setTimeout(() => { this.timers.delete(id); if (!this.destroyed) fn(); }, ms);
    this.timers.add(id);
    return id;
  }

  destroy() {
    this.destroyed = true;
    for (const c of this.cleanups.splice(0)) { try { c(); } catch { /* ignore */ } }
    for (const t of this.timers) clearTimeout(t);
    this.timers.clear();
    clearTimeout(this.renderTimer);
    clearTimeout(this.syncTimer);
    if (this.ro) { try { this.ro.disconnect(); } catch { /* ignore */ } }
    for (const u of this.objectUrls.splice(0)) { try { URL.revokeObjectURL(u); } catch { /* ignore */ } }
    this.root.classList.remove("lx-root");
    delete this.root.dataset.lxView;
    this.root.textContent = "";
    this.ta = null;
  }

  focus() { if (this.ta) this.ta.focus(); }
  getText() { return this.ta ? this.ta.value : this.o.text || ""; }
  commitEdit() {}

  setText(text) {
    if (!this.ta) { this.o.text = text; return; }
    const ta = this.ta;
    ta.focus();
    ta.setSelectionRange(0, ta.value.length);
    let ok = false;
    try { ok = document.execCommand("insertText", false, String(text)); } catch { ok = false; }
    if (!ok || ta.value !== String(text)) {
      ta.value = String(text);
      this.onInput();
    }
    ta.setSelectionRange(0, 0);
    ta.scrollTop = 0;
    this.syncScroll();
    this.updateCaretLine();
  }

  // ---------------------------------------------------------------- events

  bind() {
    const { ta, toolbar } = this.el;
    this.on(ta, "input", () => this.onInput());
    this.on(ta, "scroll", () => { this.syncScroll(); if (this.ac.open) this.closeAc(); });
    this.on(ta, "keydown", (e) => this.onKeydown(e));
    this.on(ta, "click", () => { this.closeAc(); this.onCaretMove(); });
    this.on(ta, "keyup", (e) => { if (!/^(Shift|Control|Alt|Meta)$/.test(e.key)) this.onCaretMove(); });
    this.on(ta, "blur", () => this.later(() => { if (document.activeElement !== this.ta && !this.el.ac.matches(":hover")) this.closeAc(); }, 150));
    this.on(document, "selectionchange", () => { if (document.activeElement === this.ta) this.onCaretMove(); });
    this.on(document, "mousedown", (e) => {
      if (!this.root.contains(e.target) || !e.target.closest(".lx-menu-wrap")) this.closeMenus();
    });
    if (typeof ResizeObserver !== "undefined") {
      this.ro = new ResizeObserver(() => this.syncWidth());
      this.ro.observe(ta);
    }
    this.on(window, "resize", () => this.syncWidth());

    // Find / replace / go to line also work while the preview or the toolbar has focus
    this.on(this.root, "keydown", (e) => {
      if (!(e.ctrlKey || e.metaKey) || e.altKey || e.shiftKey || e.target === this.ta) return;
      const k = e.key.toLowerCase();
      if (k !== "f" && k !== "h" && k !== "g") return;
      if (e.target.closest && e.target.closest(".lx-findbar, .lx-gotobar")) return;
      e.preventDefault();
      e.stopPropagation();
      if (k === "g") this.openGoto(); else this.openFind(k === "h");
    });

    // toolbar
    this.on(toolbar, "click", (e) => this.onToolbarClick(e));
    this.on(toolbar, "keydown", (e) => {
      if (e.key === "Escape") { this.closeMenus(); this.focus(); }
    });
    const grid = toolbar.querySelector(".lx-tp-grid");
    this.on(grid, "mouseover", (e) => {
      const c = e.target.closest("[data-r]");
      if (!c) return;
      const r = +c.dataset.r, k = +c.dataset.c;
      for (const s of grid.children) s.classList.toggle("on", +s.dataset.r <= r && +s.dataset.c <= k);
      toolbar.querySelector(".lx-tp-label").textContent = `${r} × ${k} table`;
    });

    // sidebar
    this.on(this.el.side, "click", (e) => {
      const tab = e.target.closest("[data-tab]");
      if (tab) { this.sideTab = tab.dataset.tab; this.applySide(); return; }
      const item = e.target.closest("[data-line]");
      if (item) this.jumpToLine(+item.dataset.line);
    });

    // preview
    this.on(this.el.preview, "click", (e) => this.onPreviewClick(e));
    this.on(this.el.pdfbar, "click", (e) => { if (e.target.closest(".lx-pdf-back")) this.showPreview(); });

    // find bar
    const fb = this.el.findbar;
    const q = fb.querySelector(".lx-find-q"), r = fb.querySelector(".lx-find-r");
    this.on(q, "input", () => { this.find.query = q.value; this.runFind(true); });
    this.on(r, "input", () => { this.find.replace = r.value; });
    this.on(fb, "click", (e) => {
      const opt = e.target.closest("[data-opt]");
      if (opt) {
        const k = opt.dataset.opt;
        this.find[k] = !this.find[k];
        opt.setAttribute("aria-pressed", String(this.find[k]));
        this.runFind(true);
        return;
      }
      const b = e.target.closest("[data-find]");
      if (!b) return;
      const a = b.dataset.find;
      if (a === "next") this.findStep(1);
      else if (a === "prev") this.findStep(-1);
      else if (a === "close") this.closeFind();
      else if (a === "replace") this.replaceOne();
      else if (a === "all") this.replaceAll();
    });
    this.on(fb, "keydown", (e) => {
      if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); this.closeFind(); return; }
      if (e.key === "Enter") {
        e.preventDefault();
        if (e.target === r) { if (e.ctrlKey || e.metaKey || e.altKey) this.replaceAll(); else this.replaceOne(); } else this.findStep(e.shiftKey ? -1 : 1);
        return;
      }
      const mod = e.ctrlKey || e.metaKey;
      if (mod && !e.altKey && (e.key === "f" || e.key === "h")) { e.preventDefault(); e.stopPropagation(); this.openFind(e.key === "h"); }
      if (mod && !e.altKey && e.key === "g") { e.preventDefault(); e.stopPropagation(); this.openGoto(); }
    });

    // go to line
    const gi = this.el.gotobar.querySelector(".lx-goto-in");
    this.on(this.el.gotobar, "keydown", (e) => {
      if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); this.closeGoto(true); }
      else if (e.key === "Enter") {
        e.preventDefault();
        const n = parseInt(gi.value, 10);
        if (n >= 1) { this.closeGoto(false); this.jumpToLine(Math.min(n, this.lines.length), true); }
      }
    });
    this.on(this.el.gotobar, "click", (e) => { if (e.target.closest("[data-goto]")) this.closeGoto(true); });

    // autocomplete popup
    this.on(this.el.ac, "mousedown", (e) => {
      const it = e.target.closest("[data-ac]");
      if (!it) return;
      e.preventDefault();
      this.ac.sel = +it.dataset.ac;
      this.acceptAc();
    });

    // split resizer
    this.on(this.el.resizer, "pointerdown", (e) => this.startResize(e));
  }

  onToolbarClick(e) {
    const mbtn = e.target.closest(".lx-menu-btn");
    if (mbtn) { this.toggleMenu(mbtn.dataset.menu); return; }
    const tpl = e.target.closest("[data-tpl]");
    if (tpl) { this.closeMenus(); this.applyTemplate(tpl.dataset.tpl); return; }
    const sn = e.target.closest("[data-snippet]");
    if (sn) { this.closeMenus(); this.insertSnippet(sn.dataset.snippet); return; }
    const cell = e.target.closest(".lx-tp-grid [data-r]");
    if (cell) { this.closeMenus(); this.insertTable(+cell.dataset.r, +cell.dataset.c); return; }
    const symtab = e.target.closest("[data-symtab]");
    if (symtab) {
      const k = symtab.dataset.symtab;
      for (const b of this.root.querySelectorAll("[data-symtab]")) b.classList.toggle("on", b === symtab);
      for (const g of this.root.querySelectorAll("[data-symgrid]")) g.hidden = g.dataset.symgrid !== k;
      return;
    }
    const sym = e.target.closest("[data-sym]");
    if (sym) { this.insertSymbol(sym.dataset.sym); return; }
    if (e.target.closest(".lx-find-btn")) { this.openFind(false); return; }
    if (e.target.closest(".lx-prob-btn")) { this.sideOpen = true; this.sideTab = "problems"; this.applySide(); return; }
    if (e.target.closest(".lx-side-btn")) { this.sideOpen = !this.sideOpen; lsSet(LS_SIDE, this.sideOpen ? "1" : "0"); this.applySide(); return; }
    const mode = e.target.closest("button[data-mode]");
    if (mode) { this.setViewMode(mode.dataset.mode); return; }
    if (e.target.closest(".lx-compile")) { this.compilePdf(); }
  }

  toggleMenu(name) {
    const panel = this.root.querySelector(`[data-menu-panel="${name}"]`);
    const wasOpen = panel && !panel.hidden;
    this.closeMenus();
    if (!panel || wasOpen) return;
    panel.hidden = false;
    const btn = this.root.querySelector(`[data-menu="${name}"]`);
    if (btn) btn.setAttribute("aria-expanded", "true");
    const first = panel.querySelector("button");
    if (first && name !== "table") first.focus();
  }
  closeMenus() {
    if (!this.root) return;
    for (const p of this.root.querySelectorAll("[data-menu-panel]")) p.hidden = true;
    for (const b of this.root.querySelectorAll(".lx-menu-btn")) b.setAttribute("aria-expanded", "false");
  }

  // ---------------------------------------------------------------- view modes, sidebar, resizing

  setViewMode(m) {
    if (!["source", "split", "preview"].includes(m)) return;
    this.viewMode = m;
    lsSet(LS_MODE, m);
    this.applyViewMode();
    if (m !== "source") this.renderNow();
    if (m === "preview") this.el.preview.focus({ preventScroll: true });
    else this.focus();
  }
  applyViewMode() {
    const m = this.viewMode;
    this.root.dataset.lxView = m;
    for (const b of this.root.querySelectorAll("[data-mode]")) b.setAttribute("aria-pressed", String(b.dataset.mode === m));
    this.el.editorPane.hidden = m === "preview";
    this.el.previewPane.hidden = m === "source";
    this.el.resizer.hidden = m !== "split";
    this.el.editorPane.style.flex = m === "split" ? `${this.split} 1 0` : "1 1 0";
    this.el.previewPane.style.flex = m === "split" ? `${1 - this.split} 1 0` : "1 1 0";
    this.later(() => this.syncWidth(), 0);
  }
  applySide() {
    this.el.side.hidden = !this.sideOpen;
    this.el.sideBtn.setAttribute("aria-pressed", String(this.sideOpen));
    for (const b of this.el.side.querySelectorAll("[data-tab]")) b.setAttribute("aria-selected", String(b.dataset.tab === this.sideTab));
    this.el.outline.hidden = this.sideTab !== "outline";
    this.el.problems.hidden = this.sideTab !== "problems";
  }
  startResize(e) {
    e.preventDefault();
    const main = this.el.main;
    const ed = this.el.editorPane.getBoundingClientRect();
    const pv = this.el.previewPane.getBoundingClientRect();
    const left = ed.left, total = pv.right - ed.left;
    this.el.resizer.setPointerCapture(e.pointerId);
    main.classList.add("lx-resizing");
    const move = (ev) => {
      this.split = Math.min(0.85, Math.max(0.15, (ev.clientX - left) / total));
      this.el.editorPane.style.flex = `${this.split} 1 0`;
      this.el.previewPane.style.flex = `${1 - this.split} 1 0`;
      this.syncWidth();
    };
    const up = () => {
      main.classList.remove("lx-resizing");
      this.el.resizer.removeEventListener("pointermove", move);
      this.el.resizer.removeEventListener("pointerup", up);
      lsSet(LS_SPLIT, String(this.split));
    };
    this.el.resizer.addEventListener("pointermove", move);
    this.el.resizer.addEventListener("pointerup", up);
  }

  // ---------------------------------------------------------------- highlighting / editor layer

  gutterWidth() {
    const digits = String(Math.max(10, this.lines.length)).length;
    return `${digits + 2.2}ch`;
  }

  lineHtml(i) {
    const text = this.lines[i];
    const r = highlightLine(text, this.states[i] || "");
    const decos = this.decosFor(i);
    const html = segsToHtml(r.segs, decos);
    return { html: html || "", out: r.out };
  }

  makeLineEl(i, html) {
    const d = document.createElement("div");
    d.className = "lx-l";
    d.innerHTML = html || "";
    if (!html) d.appendChild(document.createTextNode(""));
    return d;
  }

  fullHighlight() {
    const text = this.ta.value;
    this.lines = text.split("\n");
    this.lineStarts = null;
    this.states = new Array(this.lines.length);
    // one big HTML string is much faster than thousands of element insertions
    const parts = new Array(this.lines.length);
    let st = "";
    for (let i = 0; i < this.lines.length; i++) {
      this.states[i] = st;
      const r = highlightLine(this.lines[i], st);
      st = r.out;
      parts[i] = `<div class="lx-l">${segsToHtml(r.segs, null)}</div>`;
    }
    this.endState = st;
    const hl = this.el.hl;
    hl.innerHTML = parts.join("");
    this.lineEls = Array.from(hl.children);
    this.el.code.style.setProperty("--gutter", this.gutterWidth());
    this.applyLineMarks();
    this.syncWidth();
  }

  /** Re-render only the lines that changed (common prefix/suffix diff). */
  incrementalHighlight() {
    const text = this.ta.value;
    const nl = text.split("\n");
    const ol = this.lines;
    let a = 0;
    const min = Math.min(ol.length, nl.length);
    while (a < min && ol[a] === nl[a]) a++;
    let bo = ol.length, bn = nl.length;
    while (bo > a && bn > a && ol[bo - 1] === nl[bn - 1]) { bo--; bn--; }
    if (a === ol.length && a === nl.length) return;
    const oldStates = this.states;
    const hl = this.el.hl;
    // replace DOM lines a..bo-1 with new a..bn-1
    const ref = this.lineEls[bo] || null;
    for (let i = a; i < bo; i++) this.lineEls[i].remove();
    this.lines = nl;
    this.lineStarts = null;
    const newEls = [];
    const newStates = oldStates.slice(0, a);
    let st = a > 0 ? highlightLine(ol[a - 1], oldStates[a - 1] || "").out : "";
    const frag = document.createDocumentFragment();
    for (let i = a; i < bn; i++) {
      newStates[i] = st;
      const r = highlightLine(nl[i], st);
      st = r.out;
      const el = this.makeLineEl(i, segsToHtml(r.segs, this.decosFor(i, newStates)));
      newEls.push(el);
      frag.appendChild(el);
    }
    hl.insertBefore(frag, ref);
    this.lineEls.splice(a, bo - a, ...newEls);
    // the tail keeps its DOM unless its carried state changed
    const tailOld = oldStates.slice(bo);
    for (let k = 0; k < tailOld.length; k++) {
      const i = bn + k;
      newStates[i] = st;
      if (tailOld[k] === st) {
        for (let m = k + 1; m < tailOld.length; m++) newStates[bn + m] = tailOld[m];
        break;
      }
      const r = highlightLine(nl[i], st);
      st = r.out;
      this.lineEls[i].innerHTML = segsToHtml(r.segs, this.decosFor(i, newStates));
    }
    this.states = newStates;
    const g = this.gutterWidth();
    if (g !== this.el.code.style.getPropertyValue("--gutter")) this.el.code.style.setProperty("--gutter", g);
    if (bn !== bo) this.applyLineMarks();
  }

  rerenderLine(i) {
    if (i < 0 || i >= this.lines.length || !this.lineEls[i]) return;
    this.lineEls[i].innerHTML = this.lineHtml(i).html;
  }

  syncScroll() {
    const { ta, hl } = this.el;
    hl.scrollTop = ta.scrollTop;
    hl.scrollLeft = ta.scrollLeft;
  }
  syncWidth() {
    if (!this.ta || !this.el) return;
    const w = this.ta.clientWidth;
    if (w > 0) this.el.hl.style.width = `${w}px`;
    this.syncScroll();
  }

  getLineStarts() {
    if (!this.lineStarts) {
      const starts = new Array(this.lines.length);
      let p = 0;
      for (let i = 0; i < this.lines.length; i++) { starts[i] = p; p += this.lines[i].length + 1; }
      this.lineStarts = starts;
    }
    return this.lineStarts;
  }
  lineOf(pos) {
    const s = this.getLineStarts();
    let lo = 0, hi = s.length - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (s[mid] <= pos) lo = mid; else hi = mid - 1;
    }
    return lo;
  }

  // decorations: bracket match + find matches
  decosFor(i, states) {
    void states;
    const out = [];
    const s = this.getLineStarts()[i];
    if (s === undefined) return out;
    const e = s + this.lines[i].length;
    if (this.match) for (const p of this.match) if (p >= s && p < e) out.push({ s: p - s, e: p - s + 1, cls: "t-match" });
    const ms = this.find.open ? this.find.matches : null;
    if (ms && ms.length) {
      let lo = 0, hi = ms.length;
      while (lo < hi) { const mid = (lo + hi) >> 1; if (ms[mid][1] <= s) lo = mid + 1; else hi = mid; }
      for (let k = lo; k < ms.length && ms[k][0] <= e; k++) {
        const [a, b] = ms[k];
        if (b <= s && a < s) continue;
        out.push({ s: Math.max(0, a - s), e: Math.min(e - s, b - s), cls: k === this.find.cur ? "t-find t-find-cur" : "t-find" });
      }
    }
    return out;
  }
  redecorate(newLines) {
    const all = new Set([...this.decoLines, ...newLines]);
    this.decoLines = new Set(newLines);
    for (const i of all) this.rerenderLine(i);
  }

  applyLineMarks() {
    const want = new Map();
    for (const e of [...this.errors, ...this.compileErrors]) {
      const i = (e.line || 1) - 1;
      if (i < 0 || i >= this.lineEls.length) continue;
      if (e.severity === "error") want.set(i, "lx-err");
      else if (!want.has(i)) want.set(i, "lx-warn");
    }
    if (this.markedLines) for (const [i] of this.markedLines) if (this.lineEls[i]) this.lineEls[i].classList.remove("lx-err", "lx-warn");
    for (const [i, cls] of want) this.lineEls[i].classList.add(cls);
    this.markedLines = want;
    this.curLine = -1;
    this.updateCaretLine();
  }

  updateCaretLine() {
    if (!this.ta) return;
    const line = this.lineOf(this.ta.selectionStart);
    const el = this.lineEls[line];
    if (line === this.curLine && el && el.classList.contains("lx-cur")) return;
    const old = this.el.hl.querySelector(".lx-cur");
    if (old && old !== el) old.classList.remove("lx-cur");
    const changed = line !== this.curLine;
    this.curLine = line;
    if (el) el.classList.add("lx-cur");
    if (changed) this.highlightOutline();
  }

  updateBracketMatch() {
    const ta = this.ta;
    if (ta.selectionStart !== ta.selectionEnd) { if (this.match) { this.match = null; this.redecorate([]); } return; }
    const text = ta.value;
    const pos = ta.selectionStart;
    const open = "{[(", close = "}])";
    let at = -1;
    const isEsc = (p) => p > 0 && text[p - 1] === "\\" && !(p > 1 && text[p - 2] === "\\");
    if (pos > 0 && (open + close).includes(text[pos - 1]) && !isEsc(pos - 1)) at = pos - 1;
    else if ((open + close).includes(text[pos]) && !isEsc(pos)) at = pos;
    let match = null;
    if (at >= 0) {
      const c = text[at];
      const oi = open.indexOf(c), ci = close.indexOf(c);
      const dir = oi >= 0 ? 1 : -1;
      const me = oi >= 0 ? c : open[ci], other = oi >= 0 ? close[oi] : c;
      let depth = 0;
      for (let p = at, steps = 0; p >= 0 && p < text.length && steps < 200000; p += dir, steps++) {
        const ch = text[p];
        if ((ch !== me && ch !== other) || isEsc(p)) continue;
        if (ch === (dir > 0 ? me : other)) depth++;
        else { depth--; if (depth === 0) { match = [at, p]; break; } }
      }
      if (!match) match = [at];
    }
    const key = match ? match.join(",") : "";
    if (key === (this.match ? this.match.join(",") : "")) return;
    this.match = match;
    const lines = match ? match.map((p) => this.lineOf(p)) : [];
    this.redecorate(this.find.open ? [...lines, ...this.findLinesVisible()] : lines);
    if (match && match.length === 1) this.lineEls[lines[0]]?.querySelector(".t-match")?.classList.add("t-nomatch");
  }

  // ---------------------------------------------------------------- input handling

  onInput() {
    if (!this.ta) return;
    this.incrementalHighlight();
    this.o.onDirty(true);
    this.o.onStats();
    if (this.compileErrors.length) { this.compileErrors = []; }
    this.scheduleRender();
    if (this.find.open) this.runFind(false);
    this.updateCaretLine();
    this.updateBracketMatch();
    this.maybeAutocomplete();
  }

  onCaretMove() {
    if (!this.ta) return;
    this.updateCaretLine();
    this.updateBracketMatch();
    clearTimeout(this.syncTimer);
    this.syncTimer = setTimeout(() => { if (!this.destroyed) this.syncPreviewToCaret(); }, 180);
  }

  scheduleRender() {
    clearTimeout(this.renderTimer);
    const len = this.ta.value.length;
    const delay = Math.min(1500, 250 + Math.floor(len / 2000));
    this.renderTimer = setTimeout(() => { if (!this.destroyed) this.renderNow(); }, delay);
  }

  /** Render the live preview right away (also used before printing). */
  renderNow() {
    clearTimeout(this.renderTimer);
    if (!this.ta) return;
    const text = this.ta.value;
    if (text === this.lastRenderedText) return;
    this.lastRenderedText = text;
    let r;
    try { r = latexToHtml(text); } catch (e) { r = { html: "", outline: [], errors: [{ line: 1, message: String(e && e.message || e), severity: "error" }] }; }
    const pv = this.el.preview;
    const top = pv.scrollTop;
    this.el.page.innerHTML = r.html;
    pv.scrollTop = top;
    this.errors = r.errors;
    this.outline = r.outline;
    this.renderOutline();
    this.renderProblems();
    this.applyLineMarks();
  }

  // ---------------------------------------------------------------- keyboard

  onKeydown(e) {
    const ta = this.ta;
    if (e.isComposing || e.keyCode === 229) return;
    const mod = e.ctrlKey || e.metaKey;
    // autocomplete navigation
    if (this.ac.open) {
      if (e.key === "ArrowDown" || e.key === "ArrowUp") {
        e.preventDefault();
        const n = this.ac.items.length;
        this.ac.sel = (this.ac.sel + (e.key === "ArrowDown" ? 1 : n - 1)) % n;
        this.renderAc();
        return;
      }
      if (e.key === "Enter" || e.key === "Tab") { e.preventDefault(); this.acceptAc(); return; }
      if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); this.closeAc(); return; }
    }
    if (mod && !e.altKey) {
      const k = e.key.toLowerCase();
      if (k === "b" && !e.shiftKey) { e.preventDefault(); e.stopPropagation(); this.insertSnippet("bold"); return; }
      if (k === "i" && !e.shiftKey) { e.preventDefault(); e.stopPropagation(); this.insertSnippet("italic"); return; }
      if (k === "/" || e.code === "Slash") { e.preventDefault(); e.stopPropagation(); this.toggleComment(); return; }
      if (k === "f" && !e.shiftKey) { e.preventDefault(); e.stopPropagation(); this.openFind(false); return; }
      if (k === "h" && !e.shiftKey) { e.preventDefault(); e.stopPropagation(); this.openFind(true); return; }
      if (k === "g" && !e.shiftKey) { e.preventDefault(); e.stopPropagation(); this.openGoto(); return; }
      if (k === " " || e.code === "Space") { e.preventDefault(); this.maybeAutocomplete(true); return; }
      return;
    }
    if (e.altKey) return;
    if (e.key === "Escape") {
      if (this.find.open) { e.preventDefault(); e.stopPropagation(); this.closeFind(); }
      return;
    }
    if (e.key === "Tab") {
      e.preventDefault();
      if (e.shiftKey) this.outdent(); else this.indent();
      return;
    }
    if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); this.newline(); return; }
    if (e.key === "Backspace" && ta.selectionStart === ta.selectionEnd) {
      const p = ta.selectionStart;
      const v = ta.value;
      const a = v[p - 1], b = v[p];
      if (a && PAIRS[a] === b && (a !== "$" || v[p - 2] !== "\\")) {
        e.preventDefault();
        ta.setSelectionRange(p - 1, p + 1);
        this.insert("");
      }
      return;
    }
    if (e.key.length === 1) this.typeChar(e);
  }

  /** Auto-close pairs and over-type closing characters. */
  typeChar(e) {
    const ta = this.ta;
    const ch = e.key;
    const v = ta.value;
    const s = ta.selectionStart, en = ta.selectionEnd;
    const prev = v[s - 1] || "", next = v[en] || "";
    const escaped = prev === "\\" && v[s - 2] !== "\\";
    // over-type a closing char we inserted
    if ((ch === "}" || ch === "]" || ch === ")" || ch === "$") && s === en && next === ch && !escaped) {
      if (ch !== "$" || this.inInlineMath(s)) {
        e.preventDefault();
        ta.setSelectionRange(s + 1, s + 1);
        this.onCaretMove();
        return;
      }
    }
    if (!PAIRS[ch] || escaped) return;
    if (s !== en) {
      // wrap the selection
      e.preventDefault();
      const sel = v.slice(s, en);
      this.insert(ch + sel + PAIRS[ch]);
      ta.setSelectionRange(s + 1, s + 1 + sel.length);
      return;
    }
    const nextOk = !next || /[\s)\]},.;:!?]/.test(next);
    if (!nextOk) return;
    if (ch === "(" && /[A-Za-z0-9]/.test(prev)) return;
    if (ch === "$" && /[A-Za-z0-9]/.test(prev)) return;
    if (ch === "[" && !/[A-Za-z*}\]]/.test(prev) && prev !== "\\") return;
    e.preventDefault();
    this.insert(ch + PAIRS[ch]);
    ta.setSelectionRange(s + 1, s + 1);
    this.onCaretMove();
    this.maybeAutocomplete();
  }

  inInlineMath(pos) {
    const v = this.ta.value;
    const ls = v.lastIndexOf("\n", pos - 1) + 1;
    let n = 0;
    for (let p = ls; p < pos; p++) if (v[p] === "$" && v[p - 1] !== "\\") n++;
    return n % 2 === 1;
  }

  /** Replace the selection with text through the browser's editing stack (undoable). */
  insert(text) {
    const ta = this.ta;
    if (document.activeElement !== ta) ta.focus();
    let ok = false;
    try { ok = document.execCommand("insertText", false, text); } catch { ok = false; }
    if (!ok) {
      ta.setRangeText(text, ta.selectionStart, ta.selectionEnd, "end");
      this.onInput();
    }
  }

  newline() {
    const ta = this.ta;
    const v = ta.value;
    const s = ta.selectionStart;
    const ls = v.lastIndexOf("\n", s - 1) + 1;
    const before = v.slice(ls, s);
    const indent = /^[ \t]*/.exec(before)[0];
    const m = /\\begin\{([^}]+)\}(\[[^\]]*\])?(\{[^}]*\})*\s*$/.exec(before);
    if (m && ta.selectionStart === ta.selectionEnd) {
      const env = m[1];
      const begins = (v.match(new RegExp(`\\\\begin\\{${env.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}\\}`, "g")) || []).length;
      const ends = (v.match(new RegExp(`\\\\end\\{${env.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}\\}`, "g")) || []).length;
      if (begins > ends) {
        const inner = indent + "  " + (LIST_ENVS.has(env) ? "\\item " : "");
        this.insert(`\n${inner}\n${indent}\\end{${env}}`);
        const caret = s + 1 + inner.length;
        ta.setSelectionRange(caret, caret);
        this.onCaretMove();
        return;
      }
    }
    // continue \item lines inside lists
    const im = /^([ \t]*)\\item\b\s*(.*)$/.exec(before);
    if (im && im[2].trim() === "" && v.slice(s, v.indexOf("\n", s) < 0 ? v.length : v.indexOf("\n", s)).trim() === "") {
      // empty item + Enter: remove the item marker (like word processors)
      ta.setSelectionRange(ls, s);
      this.insert(indent);
      return;
    }
    this.insert("\n" + indent + (im ? "\\item " : ""));
    this.ensureCaretVisible();
  }

  selectedLines() {
    const ta = this.ta;
    const v = ta.value;
    const s = ta.selectionStart;
    let e = ta.selectionEnd;
    if (e > s && v[e - 1] === "\n") e--;
    const a = v.lastIndexOf("\n", s - 1) + 1;
    let b = v.indexOf("\n", e);
    if (b < 0) b = v.length;
    return { a, b, text: v.slice(a, b) };
  }

  replaceLines(fn) {
    const ta = this.ta;
    const { a, b, text } = this.selectedLines();
    const s0 = ta.selectionStart, e0 = ta.selectionEnd;
    const lines = text.split("\n");
    const out = lines.map(fn);
    const joined = out.join("\n");
    if (joined === text) return;
    ta.setSelectionRange(a, b);
    this.insert(joined);
    const delta0 = out[0].length - lines[0].length;
    const total = joined.length - text.length;
    ta.setSelectionRange(Math.max(a, s0 + delta0), Math.max(a, e0 + total));
    this.onCaretMove();
  }

  indent() {
    const ta = this.ta;
    if (ta.selectionStart === ta.selectionEnd || !ta.value.slice(ta.selectionStart, ta.selectionEnd).includes("\n")) {
      if (ta.selectionStart === ta.selectionEnd) { this.insert("  "); return; }
    }
    this.replaceLines((l) => (l.length ? "  " + l : l));
  }
  outdent() {
    this.replaceLines((l) => l.replace(/^( {1,2}|\t)/, ""));
  }

  toggleComment() {
    const { text } = this.selectedLines();
    const lines = text.split("\n");
    const nonEmpty = lines.filter((l) => l.trim());
    const all = nonEmpty.length > 0 && nonEmpty.every((l) => /^\s*%/.test(l));
    if (all) this.replaceLines((l) => l.replace(/^(\s*)% ?/, "$1"));
    else this.replaceLines((l) => (l.trim() ? "% " + l : l));
  }

  // ---------------------------------------------------------------- snippets & insertions

  insertSnippet(id) {
    const sn = SNIPPETS[id];
    if (!sn || !this.ta) return;
    this.insertTemplateText(sn.text, sn.block);
  }

  /** Insert text with CARET/SEL markers; block snippets go on their own line with the current indent. */
  insertTemplateText(tpl, block) {
    const ta = this.ta;
    if (this.viewMode === "preview") this.setViewMode("split");
    ta.focus();
    const v = ta.value;
    const s = ta.selectionStart, e = ta.selectionEnd;
    const sel = v.slice(s, e);
    const ls = v.lastIndexOf("\n", s - 1) + 1;
    const indent = /^[ \t]*/.exec(v.slice(ls, s))[0];
    let text = tpl.split(SEL).join(sel);
    if (block) {
      text = text.split("\n").join("\n" + indent);
      const before = v.slice(ls, s);
      if (before.trim()) text = "\n" + indent + text;
      const lineEnd = v.indexOf("\n", e);
      const after = v.slice(e, lineEnd < 0 ? v.length : lineEnd);
      if (after.trim()) text += "\n" + indent;
    }
    let caret = text.indexOf(CARET);
    if (caret < 0) caret = text.length;
    text = text.replace(CARET, "");
    this.insert(text);
    const pos = s + caret;
    ta.setSelectionRange(pos, pos);
    this.ensureCaretVisible();
    this.onCaretMove();
  }

  insertTable(rows, cols) {
    const spec = "|" + "c|".repeat(cols);
    const lines = ["\\begin{table}[h]", "  \\centering", `  \\begin{tabular}{${spec}}`, "    \\hline"];
    for (let r = 0; r < rows; r++) {
      const cells = Array.from({ length: cols }, (_, c) => (r === 0 && c === 0 ? CARET : "") + (r === 0 ? `Header ${c + 1}` : ""));
      lines.push(`    ${cells.join(" & ")} \\\\`, "    \\hline");
    }
    lines.push("  \\end{tabular}", "  \\caption{Caption}", "  \\label{tab:label}", "\\end{table}");
    this.insertTemplateText(lines.join("\n"), true);
  }

  insertSymbol(cmd) {
    const ta = this.ta;
    ta.focus();
    const s = ta.selectionStart;
    const v = ta.value;
    const inMath = this.inInlineMath(s) || this.inDisplayMath(s);
    const needSpace = /^\\[A-Za-z]+$/.test(cmd) && /[A-Za-z]/.test(v[ta.selectionEnd] || "");
    const text = inMath ? cmd + (needSpace ? " " : "") : `$${cmd}$`;
    this.insert(text);
    this.onCaretMove();
  }
  inDisplayMath(pos) {
    const i = this.lineOf(pos);
    const st = this.states[i] || "";
    return st === "m" || highlightLine(this.lines[i].slice(0, pos - this.getLineStarts()[i]), st).out === "m";
  }

  async applyTemplate(id) {
    const t = TEMPLATES.find((x) => x.id === id);
    if (!t) return;
    const ok = await this.o.confirm(`Replace the current document with the "${t.label}" template? You can undo this with Ctrl+Z.`, { title: "Use a template?", ok: "Replace" });
    if (!ok || this.destroyed) return;
    this.setText(newLatexDocument(id));
    this.o.onStatus(`Started from the ${t.label} template.`);
  }

  // ---------------------------------------------------------------- autocomplete

  maybeAutocomplete(force = false) {
    const ta = this.ta;
    if (!ta || ta.selectionStart !== ta.selectionEnd) { this.closeAc(); return; }
    const pos = ta.selectionStart;
    const v = ta.value;
    const ls = v.lastIndexOf("\n", pos - 1) + 1;
    const before = v.slice(Math.max(ls, pos - 200), pos);
    let m;
    let items = [];
    let start = pos, kind = "";
    if ((m = /\\begin\{([A-Za-z*]*)$/.exec(before))) {
      kind = "env";
      start = pos - m[1].length;
      const extra = latexKeys(v).envs;
      const pool = [...new Set([...ENVIRONMENT_COMPLETIONS, ...extra])];
      items = pool.filter((n) => n.startsWith(m[1]) && n !== m[1]).map((n) => ({ label: n, insert: n, desc: "environment" }));
    } else if ((m = /\\(?:ref|eqref|autoref|cref|Cref|pageref|nameref|vref)\{([^}\s]*)$/.exec(before))) {
      kind = "key";
      start = pos - m[1].length;
      items = latexKeys(v).labels.filter((k) => k.startsWith(m[1]) && k !== m[1]).map((k) => ({ label: k, insert: k, desc: "label" }));
    } else if ((m = /\\(?:cite|citep|citet|parencite|textcite|autocite)(?:\[[^\]]*\])*\{(?:[^}]*,)?\s*([^,}\s]*)$/.exec(before))) {
      kind = "key";
      start = pos - m[1].length;
      items = latexKeys(v).cites.filter((k) => k.startsWith(m[1]) && k !== m[1]).map((k) => ({ label: k, insert: k, desc: "citation" }));
    } else if ((m = /\\([A-Za-z]*)$/.exec(before)) && (m[1].length >= 1 || force)) {
      kind = "cmd";
      start = pos - m[1].length;
      const user = userCommandNames(v).map((n) => [n, n, "your command"]);
      const seen = new Set();
      for (const [name, ins, desc] of [...COMMAND_COMPLETIONS, ...user]) {
        if (seen.has(name) || !name.startsWith(m[1]) || name === m[1]) continue;
        seen.add(name);
        items.push({ label: "\\" + name, insert: ins, desc });
      }
      items.sort((a, b) => (a.label.length - b.label.length) || a.label.localeCompare(b.label));
    }
    items = items.slice(0, 50);
    if (!items.length) { this.closeAc(); return; }
    this.ac = { open: true, items, sel: 0, start, kind };
    this.renderAc();
  }

  renderAc() {
    const ac = this.el.ac;
    ac.innerHTML = this.ac.items.map((it, k) => `<div class="lx-ac-item${k === this.ac.sel ? " on" : ""}" data-ac="${k}" role="option" aria-selected="${k === this.ac.sel}"><span class="lx-ac-name">${esc(it.label)}</span><span class="lx-ac-desc">${esc(it.desc || "")}</span></div>`).join("");
    ac.hidden = false;
    const xy = this.caretXY(this.ac.start);
    const code = this.el.code.getBoundingClientRect();
    if (xy) {
      let left = xy.x - code.left - 4;
      let top = xy.y - code.top + xy.h + 2;
      const w = ac.offsetWidth || 300, h = ac.offsetHeight || 200;
      if (left + w > code.width - 8) left = Math.max(4, code.width - w - 8);
      if (top + h > code.height - 4 && xy.y - code.top - h - 2 > 0) top = xy.y - code.top - h - 2;
      ac.style.left = `${left}px`;
      ac.style.top = `${top}px`;
    }
    const on = ac.querySelector(".on");
    if (on) on.scrollIntoView({ block: "nearest" });
  }

  closeAc() {
    if (!this.ac.open) return;
    this.ac.open = false;
    if (this.el && this.el.ac) this.el.ac.hidden = true;
  }

  acceptAc() {
    const it = this.ac.items[this.ac.sel];
    if (!it) { this.closeAc(); return; }
    const ta = this.ta;
    const pos = ta.selectionStart;
    const kind = this.ac.kind;
    this.closeAc();
    ta.setSelectionRange(this.ac.start, pos);
    const v = ta.value;
    if (kind === "env") {
      const ls = v.lastIndexOf("\n", this.ac.start - 1) + 1;
      const indent = /^[ \t]*/.exec(v.slice(ls))[0];
      const hasClose = v[pos] === "}";
      const inner = indent + "  " + (LIST_ENVS.has(it.insert) ? "\\item " : "");
      const lineEnd = v.indexOf("\n", pos);
      const rest = v.slice(hasClose ? pos + 1 : pos, lineEnd < 0 ? v.length : lineEnd);
      if (hasClose) ta.setSelectionRange(this.ac.start, pos + 1);
      const text = `${it.insert}}` + (rest.trim() ? "" : `\n${inner}\n${indent}\\end{${it.insert}}`);
      this.insert(text);
      if (!rest.trim()) { const caret = this.ac.start + it.insert.length + 2 + inner.length; ta.setSelectionRange(caret, caret); }
    } else if (kind === "key") {
      this.insert(it.insert + (v[pos] === "}" ? "" : "}"));
      const c = this.ac.start + it.insert.length + 1;
      ta.setSelectionRange(c, c);
    } else {
      const ins = it.insert;
      this.insert(ins);
      const brace = ins.indexOf("{");
      const caret = this.ac.start + (brace >= 0 ? brace + 1 : ins.length);
      ta.setSelectionRange(caret, caret);
    }
    this.onCaretMove();
  }

  /** Viewport coordinates of a text offset (from the highlight layer). */
  caretXY(pos) {
    const i = this.lineOf(pos);
    const el = this.lineEls[i];
    if (!el) return null;
    let col = pos - this.getLineStarts()[i];
    const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
    let node = walker.nextNode();
    while (node && col > node.data.length) { col -= node.data.length; node = walker.nextNode(); }
    const range = document.createRange();
    if (node) { range.setStart(node, Math.min(col, node.data.length)); range.setEnd(node, Math.min(col, node.data.length)); }
    else { range.selectNodeContents(el); range.collapse(true); }
    const rects = range.getClientRects();
    const r = rects.length ? rects[rects.length - 1] : el.getBoundingClientRect();
    const lh = parseFloat(getComputedStyle(this.ta).lineHeight) || 20;
    return { x: r.left, y: r.top, h: r.height || lh };
  }

  ensureCaretVisible() {
    const ta = this.ta;
    const i = this.lineOf(ta.selectionStart);
    const el = this.lineEls[i];
    if (!el) return;
    const top = el.offsetTop, h = el.offsetHeight;
    if (top < ta.scrollTop + 4) ta.scrollTop = Math.max(0, top - 8);
    else if (top + h > ta.scrollTop + ta.clientHeight - 4) ta.scrollTop = top + h - ta.clientHeight + 8;
    this.syncScroll();
  }

  // ---------------------------------------------------------------- navigation & preview sync

  /** Put the caret at the start of a 1-based line and show it. */
  jumpToLine(line, focus = true) {
    if (!this.ta) return;
    if (this.viewMode === "preview") this.setViewMode("split");
    const i = Math.max(0, Math.min(this.lines.length - 1, (line | 0) - 1));
    const pos = this.getLineStarts()[i] || 0;
    const ta = this.ta;
    if (focus) ta.focus({ preventScroll: true });
    ta.setSelectionRange(pos, pos);
    const el = this.lineEls[i];
    if (el) {
      ta.scrollTop = Math.max(0, el.offsetTop - ta.clientHeight / 3);
      this.syncScroll();
      el.classList.remove("lx-flash");
      void el.offsetWidth;
      el.classList.add("lx-flash");
      this.later(() => el.classList.remove("lx-flash"), 1200);
    }
    this.onCaretMove();
  }

  onPreviewClick(e) {
    const a = e.target.closest("a");
    if (a) {
      const href = a.getAttribute("href") || "";
      if (href.startsWith("#")) {
        e.preventDefault();
        const id = href.slice(1);
        const t = id && this.el.page.querySelector(`[id="${CSS.escape(id)}"]`);
        if (t) t.scrollIntoView({ block: "center", behavior: "smooth" });
        return;
      }
      return; // external links open in a new tab (target=_blank)
    }
    if (this.viewMode === "preview") return;
    const sel = window.getSelection();
    if (sel && !sel.isCollapsed && this.el.preview.contains(sel.anchorNode)) return;
    const b = e.target.closest("[data-line]");
    if (b) this.jumpToLine(+b.dataset.line, true);
  }

  syncPreviewToCaret() {
    if (this.viewMode !== "split" || this.showingPdf || !this.ta) return;
    const line = this.lineOf(this.ta.selectionStart) + 1;
    if (this.lastSyncLine === line) return;
    this.lastSyncLine = line;
    const els = this.el.page.querySelectorAll("[data-line]");
    let best = null, bestLine = -1;
    for (const el of els) {
      const l = +el.dataset.line;
      if (l <= line && l >= bestLine) { best = el; bestLine = l; }
    }
    if (!best) return;
    const pv = this.el.preview.getBoundingClientRect();
    const r = best.getBoundingClientRect();
    if (r.top >= pv.top + 20 && r.bottom <= pv.bottom - 20) return;
    this.el.preview.scrollTop += r.top - pv.top - pv.height / 3;
  }

  renderOutline() {
    const items = this.outline;
    if (!items.length) {
      this.el.outline.innerHTML = '<div class="lx-empty">No sections yet. Add one with <code>\\section{…}</code> or Insert ▸ Section.</div>';
      return;
    }
    const min = Math.min(...items.map((o) => o.level));
    this.el.outline.innerHTML = items.map((o) => `<button type="button" class="lx-ol-item${o.frame ? " lx-ol-frame" : ""}" data-line="${o.line}" style="padding-left:${10 + (o.level - min) * 14}px" title="Line ${o.line}">${o.num ? `<span class="lx-ol-num">${esc(o.num)}</span>` : ""}<span class="lx-ol-title">${esc(o.title)}</span></button>`).join("");
    this.highlightOutline();
  }
  highlightOutline() {
    if (!this.el || !this.outline.length) return;
    const line = this.curLine + 1;
    let cur = -1;
    this.outline.forEach((o, k) => { if (o.line <= line) cur = k; });
    const btns = this.el.outline.querySelectorAll(".lx-ol-item");
    btns.forEach((b, k) => b.classList.toggle("on", k === cur));
  }

  renderProblems() {
    const all = [...this.compileErrors, ...this.errors];
    const nErr = all.filter((e) => e.severity === "error").length;
    const nWarn = all.length - nErr;
    const badge = this.el.probBadge;
    badge.textContent = String(all.length);
    const kind = nErr ? "err" : nWarn ? "warn" : "ok";
    if (this.el.probBtn.dataset.icon !== kind) {
      this.el.probBtn.dataset.icon = kind;
      const old = this.el.probBtn.querySelector(".lx-ic");
      if (old) old.outerHTML = icon(kind);
    }
    this.el.probBtn.classList.toggle("has-err", nErr > 0);
    this.el.probBtn.classList.toggle("has-warn", nErr === 0 && nWarn > 0);
    this.el.probBtn.title = all.length ? `${nErr} error${nErr === 1 ? "" : "s"}, ${nWarn} warning${nWarn === 1 ? "" : "s"} — click to see them` : "No problems found";
    this.el.tabCount.textContent = all.length ? String(all.length) : "";
    if (!all.length) {
      this.el.problems.innerHTML = '<div class="lx-empty lx-ok">No problems found.</div>';
      return;
    }
    this.el.problems.innerHTML = all.slice(0, 300).map((e) => `<button type="button" class="lx-prob lx-prob-${e.severity === "error" ? "error" : "warning"}" data-line="${e.line}">${icon(e.severity === "error" ? "err" : "warn")}<span class="lx-prob-text"><span class="lx-prob-line">${e.source ? `${esc(e.source)} · ` : ""}Line ${e.line}</span><span class="lx-prob-msg">${esc(e.message)}</span></span></button>`).join("");
  }

  // ---------------------------------------------------------------- find & replace, go to line

  openFind(replace) {
    if (this.viewMode === "preview") this.setViewMode("split");
    this.closeGoto(false);
    const fb = this.el.findbar;
    fb.hidden = false;
    fb.classList.toggle("lx-with-replace", !!replace);
    this.find.open = true;
    const q = fb.querySelector(".lx-find-q");
    const ta = this.ta;
    const sel = ta.value.slice(ta.selectionStart, ta.selectionEnd);
    if (sel && !sel.includes("\n") && sel.length < 200) { q.value = sel; this.find.query = sel; }
    this.runFind(true);
    (replace && q.value ? fb.querySelector(".lx-find-r") : q).focus();
    q.select();
    if (replace && q.value) fb.querySelector(".lx-find-r").focus();
  }
  closeFind() {
    if (!this.find.open) return;
    this.find.open = false;
    this.find.matches = [];
    this.el.findbar.hidden = true;
    this.redecorate(this.match ? this.match.map((p) => this.lineOf(p)) : []);
    this.focus();
  }

  findRegex() {
    const f = this.find;
    if (!f.query) return null;
    let src = f.regex ? f.query : f.query.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    if (f.word) src = `(?<![\\w])(?:${src})(?![\\w])`;
    try { f.error = ""; return new RegExp(src, f.caseSens ? "gu" : "giu"); } catch (e) { f.error = "Invalid pattern"; return null; }
  }

  runFind(selectNearest) {
    const f = this.find;
    const re = this.findRegex();
    const v = this.ta.value;
    const ms = [];
    if (re) {
      let m, guard = 0;
      while ((m = re.exec(v)) && guard++ < 20000) {
        if (m[0].length === 0) { re.lastIndex++; continue; }
        ms.push([m.index, m.index + m[0].length]);
      }
    }
    f.matches = ms;
    if (selectNearest || f.cur >= ms.length) {
      const p = this.ta.selectionStart;
      f.cur = ms.length ? Math.max(0, ms.findIndex((x) => x[0] >= p)) : -1;
      if (f.cur < 0 && ms.length) f.cur = 0;
    }
    this.updateFindCount();
    this.redecorate([...this.findLinesVisible(), ...(this.match ? this.match.map((p) => this.lineOf(p)) : [])]);
    if (selectNearest && f.cur >= 0) this.revealMatch(false);
  }
  findLinesVisible() {
    const lines = new Set();
    for (const [a, b] of this.find.matches.slice(0, 5000)) { lines.add(this.lineOf(a)); if (b - a > 0) lines.add(this.lineOf(b - 1)); }
    return [...lines];
  }
  updateFindCount() {
    const f = this.find;
    const el = this.el.findbar.querySelector(".lx-find-count");
    el.textContent = f.error ? f.error : !f.query ? "" : f.matches.length ? `${f.cur + 1} of ${f.matches.length}${f.matches.length >= 20000 ? "+" : ""}` : "No results";
    el.classList.toggle("lx-none", !!f.query && !f.matches.length);
  }
  findStep(d) {
    const f = this.find;
    if (!f.matches.length) { this.runFind(true); if (!f.matches.length) return; }
    f.cur = (f.cur + d + f.matches.length) % f.matches.length;
    this.updateFindCount();
    this.redecorate(this.findLinesVisible());
    this.revealMatch(true);
  }
  revealMatch(selectIt) {
    const m = this.find.matches[this.find.cur];
    if (!m) return;
    const ta = this.ta;
    if (selectIt) ta.setSelectionRange(m[0], m[1]);
    const el = this.lineEls[this.lineOf(m[0])];
    if (el) {
      const top = el.offsetTop;
      if (top < ta.scrollTop || top > ta.scrollTop + ta.clientHeight - 30) ta.scrollTop = Math.max(0, top - ta.clientHeight / 3);
      this.syncScroll();
    }
  }
  replaceOne() {
    const f = this.find;
    if (!f.matches.length) return;
    const m = f.matches[f.cur] || f.matches[0];
    const ta = this.ta;
    const v = ta.value;
    let rep = f.replace;
    if (f.regex) { const re = this.findRegex(); re.lastIndex = 0; rep = v.slice(m[0], m[1]).replace(new RegExp(re.source, re.flags.replace("g", "")), f.replace); }
    const active = document.activeElement;
    ta.setSelectionRange(m[0], m[1]);
    this.insert(rep);
    const next = m[0] + rep.length;
    ta.setSelectionRange(next, next);
    this.runFind(false);
    f.cur = f.matches.findIndex((x) => x[0] >= next);
    if (f.cur < 0) f.cur = f.matches.length ? 0 : -1;
    this.updateFindCount();
    this.redecorate(this.findLinesVisible());
    if (f.cur >= 0) this.revealMatch(false);
    this.o.onStatus(f.matches.length ? `Replaced 1. ${f.matches.length} more match${f.matches.length === 1 ? "" : "es"}.` : "Replaced the last match.");
    if (active && active !== ta && this.root.contains(active)) active.focus();
  }
  replaceAll() {
    const f = this.find;
    const re = this.findRegex();
    if (!re || !f.matches.length) return;
    const ta = this.ta;
    const n = f.matches.length;
    const v = ta.value;
    const out = f.regex ? v.replace(re, f.replace) : v.replace(re, () => f.replace);
    const active = document.activeElement;
    // replace only the changed span so undo restores it in one step
    let a = 0;
    while (a < v.length && a < out.length && v[a] === out[a]) a++;
    let bv = v.length, bo = out.length;
    while (bv > a && bo > a && v[bv - 1] === out[bo - 1]) { bv--; bo--; }
    ta.setSelectionRange(a, bv);
    this.insert(out.slice(a, bo));
    this.runFind(true);
    this.o.onStatus(`Replaced ${n} match${n === 1 ? "" : "es"}.`);
    if (active && active !== ta && this.root.contains(active)) active.focus();
  }

  openGoto() {
    if (this.viewMode === "preview") this.setViewMode("split");
    if (this.find.open) { this.find.open = false; this.el.findbar.hidden = true; this.redecorate([]); }
    const gb = this.el.gotobar;
    gb.hidden = false;
    const inp = gb.querySelector(".lx-goto-in");
    inp.max = String(this.lines.length);
    inp.value = String(this.curLine + 1);
    gb.querySelector(".lx-goto-hint").textContent = `1 – ${this.lines.length}`;
    inp.focus();
    inp.select();
  }
  closeGoto(refocus) {
    if (this.el.gotobar.hidden) return;
    this.el.gotobar.hidden = true;
    if (refocus) this.focus();
  }

  // ---------------------------------------------------------------- compile PDF

  refreshCompile() {
    if (!this.el) return;
    let can = false;
    try { can = !!this.o.canCompile(); } catch { can = false; }
    const b = this.el.compile;
    b.disabled = !can || this.compiling;
    b.classList.toggle("primary", can);
    b.title = can
      ? "Compile the document to PDF with the TeX engine of the office backend"
      : "Compile PDF needs the office-backend helper running with a TeX engine (tectonic, pdflatex, xelatex or lualatex). Meanwhile, Print (Ctrl+P) makes a PDF of the live preview.";
    this.el.compile.setAttribute("aria-disabled", String(!can));
  }

  async compilePdf() {
    if (this.compiling) return;
    let can = false;
    try { can = !!this.o.canCompile(); } catch { can = false; }
    if (!can) { this.o.onStatus("Compile PDF needs the office backend with a TeX engine. Use Print to save the preview as PDF."); return; }
    this.compiling = true;
    this.refreshCompile();
    this.el.compile.classList.add("busy");
    this.o.onStatus("Compiling the PDF…");
    try {
      const blob = await this.o.compile(this.getText());
      if (this.destroyed) return;
      this.compileErrors = [];
      const url = URL.createObjectURL(blob instanceof Blob ? blob : new Blob([blob], { type: "application/pdf" }));
      for (const u of this.objectUrls.splice(0)) URL.revokeObjectURL(u);
      this.objectUrls.push(url);
      this.el.pdf.src = url;
      this.showingPdf = true;
      if (this.viewMode === "source") this.setViewMode("split");
      this.el.pdf.hidden = false;
      this.el.pdfbar.hidden = false;
      this.el.preview.hidden = true;
      this.o.onStatus("PDF compiled. Use “Back to live preview” to return.");
    } catch (e) {
      if (this.destroyed) return;
      const msg = String(e && e.message ? e.message : e);
      this.compileErrors = parseTexLog(msg);
      this.sideOpen = true;
      this.sideTab = "problems";
      this.applySide();
      this.o.onStatus(`The PDF could not be compiled: ${msg.split("\n")[0].slice(0, 160)}`);
    } finally {
      this.compiling = false;
      if (!this.destroyed) {
        this.el.compile.classList.remove("busy");
        this.refreshCompile();
        this.renderProblems();
        this.applyLineMarks();
      }
    }
  }

  showPreview() {
    this.showingPdf = false;
    this.el.pdf.hidden = true;
    this.el.pdfbar.hidden = true;
    this.el.preview.hidden = false;
    this.el.pdf.removeAttribute("src");
    for (const u of this.objectUrls.splice(0)) URL.revokeObjectURL(u);
    this.renderNow();
  }

  beforePrint() {
    this.lastRenderedText = null;
    this.renderNow();
  }
}

/** Errors from a TeX log excerpt: "l.12 ...", "file.tex:12: message", "! Undefined control sequence." */
export function parseTexLog(msg) {
  const out = [];
  const lines = String(msg || "").split(/\r?\n/);
  let pending = null;
  for (const ln of lines) {
    let m = /^(?:\.\/)?[^\s:]*\.tex:(\d+):\s*(.*)$/.exec(ln);
    if (m) { out.push({ line: +m[1], message: m[2] || "TeX error", severity: "error", source: "PDF" }); pending = null; continue; }
    m = /^!\s*(.*)$/.exec(ln);
    if (m) { pending = m[1]; continue; }
    m = /^l\.(\d+)\s?(.*)$/.exec(ln);
    if (m) { out.push({ line: +m[1], message: pending ? `${pending}${m[2] ? ` (near “${m[2].trim().slice(-40)}”)` : ""}` : `TeX error near “${m[2].trim()}”`, severity: "error", source: "PDF" }); pending = null; continue; }
  }
  if (pending) out.push({ line: 1, message: pending, severity: "error", source: "PDF" });
  if (!out.length) out.push({ line: 1, message: lines.find((l) => l.trim()) || "The PDF could not be compiled", severity: "error", source: "PDF" });
  return out.slice(0, 50);
}
