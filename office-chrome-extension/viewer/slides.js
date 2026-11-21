/**
 * SlidesView: the PowerPoint editor (owned by the slides-frontend stream).
 *
 * Interface used by viewer/viewer.js:
 *   new SlidesView({ root, getModel, onStatus, onDirty, onStats, confirm })
 *     root      #slides-view (inside #stage, which is a full-bleed flex column in
 *               "app-mode"); the view builds ALL its UI (toolbar, filmstrip, canvas,
 *               notes, slideshow overlay) inside root and removes it on destroy().
 *     getModel  () => deck (lib/model-slides.js CONTRACT); edit it IN PLACE.
 *     onStatus  (message) status-bar text
 *     onDirty   (true) after every change to the model
 *     onStats   () call when the current slide / slide count changes
 *     confirm   async (message, {title, ok}) => boolean
 *   render()      build the UI for the current model
 *   destroy()     remove listeners (incl. window/document ones) and DOM
 *   focus()       put keyboard focus into the editor
 *   commitEdit()  flush an in-progress text edit into the model (before save/print)
 *   getStats()    -> { slide: 1-based current slide, slides: count }
 *   beforePrint() prepare a print layout (one slide per landscape page)
 *
 * Extra (tests / automation): current, selection (item indices), undo(), redo(),
 * startShow(index), endShow(), selectSlide(i), selectItems([i]).
 */

import {
  TEMPLATES, LAYOUTS, SHAPES, SIZES, TRANSITIONS, newSlide, renderSlideHtml, renderItemHtml, shapePath,
  cloneDeep, hexColor, tintColor, itemText, paraText, slideTitle, resizeDeck, textColorFor, defaultRunSize,
} from "../lib/model-slides.js";

// ---------------------------------------------------------------------------
// icons (16x16, stroke = currentColor)
// ---------------------------------------------------------------------------

const ICON_PATHS = {
  undo: '<path d="M4 6h6a4 4 0 0 1 0 8H7"/><path d="M6.5 3.5 4 6l2.5 2.5"/>',
  redo: '<path d="M12 6H6a4 4 0 0 0 0 8h3"/><path d="M9.5 3.5 12 6 9.5 8.5"/>',
  plus: '<path d="M8 3v10M3 8h10"/>',
  dup: '<rect x="5" y="5" width="8.5" height="8.5" rx="1.5"/><path d="M2.5 10.5V4A1.5 1.5 0 0 1 4 2.5h6.5"/>',
  trash: '<path d="M2.5 4.5h11M6.5 4.5V2.8h3v1.7M4 4.5l.7 9h6.6l.7-9"/>',
  textbox: '<rect x="1.8" y="2.3" width="12.4" height="11.4" rx="1.5"/><path d="M5 5.5h6M8 5.5v5.8"/>',
  shapes: '<rect x="1.8" y="7.2" width="7" height="7" rx="1"/><circle cx="10.6" cy="5.4" r="3.6"/>',
  image: '<rect x="1.8" y="2.8" width="12.4" height="10.4" rx="1.5"/><circle cx="5.8" cy="6.3" r="1.2"/><path d="m2.5 12.2 3.6-3.6 2.6 2.6 2-2 2.8 2.8"/>',
  table: '<rect x="1.8" y="2.8" width="12.4" height="10.4" rx="1"/><path d="M1.8 6.3h12.4M1.8 9.8h12.4M6 2.8v10.4M10 2.8v10.4"/>',
  alignL: '<path d="M2.5 3.5h11M2.5 6.5h7M2.5 9.5h11M2.5 12.5h7"/>',
  alignC: '<path d="M2.5 3.5h11M4.5 6.5h7M2.5 9.5h11M4.5 12.5h7"/>',
  alignR: '<path d="M2.5 3.5h11M6.5 6.5h7M2.5 9.5h11M6.5 12.5h7"/>',
  alignJ: '<path d="M2.5 3.5h11M2.5 6.5h11M2.5 9.5h11M2.5 12.5h11"/>',
  bullets: '<circle cx="3.2" cy="4.5" r=".6"/><circle cx="3.2" cy="8" r=".6"/><circle cx="3.2" cy="11.5" r=".6"/><path d="M6.3 4.5h7.2M6.3 8h7.2M6.3 11.5h7.2"/>',
  numbers: '<path d="M2.6 3.3l1-.6v3.4M2.5 6.1h2"/><path d="M2.5 9.4c.3-.7 1.9-.8 1.9.2 0 .8-1.9 1.3-1.9 2.3h2"/><path d="M6.8 4.5h6.7M6.8 8h6.7M6.8 11.5h6.7"/>',
  outdent: '<path d="M7.5 4h6M7.5 8h6M7.5 12h6M5 5.8 3 8l2 2.2"/>',
  indent: '<path d="M7.5 4h6M7.5 8h6M7.5 12h6M3 5.8 5 8l-2 2.2"/>',
  fill: '<path d="M3.2 8.6 7.8 4l4.6 4.6-4.1 4.1a1 1 0 0 1-1.4 0z"/><path d="M5.6 1.8 7.8 4"/><path d="M13.6 10.4s1 1.4 1 2.1a1 1 0 0 1-2 0c0-.7 1-2.1 1-2.1z"/>',
  outline: '<rect x="2.5" y="2.5" width="11" height="9" rx="1.2"/>',
  arrange: '<path d="M8 1.8 1.8 5 8 8.2 14.2 5z"/><path d="m1.8 8.2 6.2 3.2 6.2-3.2M1.8 11.2 8 14.4l6.2-3.2"/>',
  layout: '<rect x="1.8" y="2.3" width="12.4" height="11.4" rx="1.2"/><path d="M1.8 5.8h12.4M7 5.8v7.9"/>',
  background: '<rect x="1.8" y="2.3" width="12.4" height="11.4" rx="1.2"/><path d="M1.8 10.5 5.5 7l3 3 2-2 3.7 3.5"/><circle cx="11" cy="5.2" r=".9"/>',
  transition: '<rect x="1.5" y="3.5" width="8" height="9" rx="1"/><path d="M11.5 5v6M14 6.5v3"/>',
  theme: '<path d="M8 1.8a6.2 6.2 0 1 0 0 12.4c1.1 0 1.4-.9.9-1.6-.5-.8 0-1.8 1-1.8h1.3a3 3 0 0 0 3-3C14.2 4.5 11.5 1.8 8 1.8z"/><circle cx="4.9" cy="7.2" r=".7"/><circle cx="7.4" cy="4.8" r=".7"/><circle cx="10.6" cy="5.6" r=".7"/>',
  size: '<rect x="1.5" y="3.5" width="13" height="9" rx="1.2"/><path d="M4.5 6v4M11.5 6v4M4.5 8h7"/>',
  play: '<path d="M5 3.2v9.6L12.6 8z"/>',
  eye: '<path d="M1.6 8S4 3.8 8 3.8 14.4 8 14.4 8 12 12.2 8 12.2 1.6 8 1.6 8z"/><circle cx="8" cy="8" r="1.9"/>',
  eyeOff: '<path d="M1.6 8S4 3.8 8 3.8 14.4 8 14.4 8 12 12.2 8 12.2 1.6 8 1.6 8z"/><circle cx="8" cy="8" r="1.9"/><path d="M2.5 13.5 13.5 2.5"/>',
  zoom: '<circle cx="7" cy="7" r="4.5"/><path d="m10.4 10.4 3.6 3.6"/>',
  notes: '<rect x="2.5" y="2" width="11" height="12" rx="1.2"/><path d="M5 5.5h6M5 8h6M5 10.5h4"/>',
  up: '<path d="M8 13V3M4 7l4-4 4 4"/>',
  down: '<path d="M8 3v10M4 9l4 4 4-4"/>',
  front: '<rect x="5" y="5" width="8" height="8" rx="1" fill="currentColor" fill-opacity=".25"/><path d="M3 10.5V3h7.5"/>',
  back: '<rect x="2.5" y="2.5" width="8" height="8" rx="1" fill="currentColor" fill-opacity=".25"/><path d="M13 5.5V13H5.5"/>',
};

function icon(name) {
  const p = ICON_PATHS[name] || "";
  const filled = name === "play" || name === "bullets";
  return `<svg class="sv-ic" viewBox="0 0 16 16" width="16" height="16" aria-hidden="true" fill="${filled ? "currentColor" : "none"}" stroke="currentColor" stroke-width="${filled && name === "play" ? 0 : 1.35}" stroke-linecap="round" stroke-linejoin="round">${p}</svg>`;
}

const FONTS = ["Calibri", "Calibri Light", "Arial", "Arial Black", "Segoe UI", "Segoe UI Semibold", "Segoe UI Light", "Georgia", "Times New Roman", "Trebuchet MS", "Verdana", "Tahoma", "Garamond", "Cambria", "Century Gothic", "Courier New", "Consolas", "Comic Sans MS", "Impact"];
const SIZES_PT = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72, 80, 88, 96];
const STD_COLORS = ["#C00000", "#FF0000", "#FFC000", "#FFFF00", "#92D050", "#00B050", "#00B0F0", "#0070C0", "#002060", "#7030A0"];
const GRAYS = ["#FFFFFF", "#F2F2F2", "#D9D9D9", "#BFBFBF", "#A6A6A6", "#7F7F7F", "#595959", "#404040", "#262626", "#000000"];
const UNDO_CAP = 100;
const NUDGE_COALESCE_MS = 1200;
const CLIP_MIME = "application/x-office-slides+json";

const isMac = typeof navigator !== "undefined" && /Mac|iPhone|iPad/.test(navigator.platform || "");

function esc(s) {
  return String(s ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
}

function rgbToHex(s) {
  if (!s) return undefined;
  const h = hexColor(s);
  if (h) return h;
  const m = /rgba?\(\s*(\d+)[,\s]+(\d+)[,\s]+(\d+)(?:[,\s/]+([\d.]+))?/i.exec(s);
  if (!m) return undefined;
  if (m[4] !== undefined && Number(m[4]) === 0) return undefined;
  return "#" + [m[1], m[2], m[3]].map((v) => Math.max(0, Math.min(255, +v)).toString(16).padStart(2, "0")).join("").toUpperCase();
}

const r2 = (v) => Math.round(v * 100) / 100;

function isTypingEl(t) {
  return !!t && t.nodeType === 1 && (t.isContentEditable || /^(INPUT|TEXTAREA|SELECT)$/.test(t.tagName));
}

/** Read image file -> { src, w, h } (pt), converting formats PowerPoint cannot show to PNG. */
async function readImageFile(file) {
  const dataUrl = await new Promise((res, rej) => {
    const fr = new FileReader();
    fr.onload = () => res(String(fr.result));
    fr.onerror = () => rej(new Error("The picture could not be read."));
    fr.readAsDataURL(file);
  });
  const img = await new Promise((res, rej) => {
    const im = new Image();
    im.onload = () => res(im);
    im.onerror = () => rej(new Error("This picture format is not supported."));
    im.src = dataUrl;
  });
  let src = dataUrl;
  const w = img.naturalWidth || 300, h = img.naturalHeight || 200;
  if (!/^data:image\/(png|jpeg|gif);/i.test(dataUrl)) {
    const c = document.createElement("canvas");
    c.width = Math.min(4096, w); c.height = Math.min(4096, Math.round(h * (c.width / w)));
    c.getContext("2d").drawImage(img, 0, 0, c.width, c.height);
    src = c.toDataURL("image/png");
  }
  return { src, w: w * 0.75, h: h * 0.75, name: file.name || "" };
}

// ---------------------------------------------------------------------------
// contenteditable <-> paras
// ---------------------------------------------------------------------------

const BLOCK_RE = /^(P|DIV|LI|H[1-6]|BLOCKQUOTE|UL|OL|PRE)$/;

function firstFamily(ff) {
  if (!ff) return undefined;
  const f = ff.split(",")[0].trim().replace(/^["']|["']$/g, "");
  return f || undefined;
}

/** Formatting of a text node from its ancestors up to (excluding) `block`. */
function fmtOf(node, block, defaults) {
  const chain = [];
  for (let n = node.parentElement; n && n !== block; n = n.parentElement) {
    if (n.classList && n.classList.contains("sl-tx")) break;
    chain.unshift(n);
  }
  const f = {};
  if (block && block.dataset && block.dataset.psz) f.size = +block.dataset.psz;
  for (const n of chain) {
    const tag = n.tagName;
    const st = n.style;
    if (tag === "B" || tag === "STRONG") f.b = true;
    if (tag === "I" || tag === "EM") f.i = true;
    if (tag === "U" || tag === "INS") f.u = true;
    if (tag === "S" || tag === "STRIKE" || tag === "DEL") f.s = true;
    if (tag === "SUP" || n.dataset.sup) { f.sup = true; f.sub = false; }
    if (tag === "SUB" || n.dataset.sub) { f.sub = true; f.sup = false; }
    if (st.fontWeight) f.b = st.fontWeight === "bold" || st.fontWeight === "bolder" || +st.fontWeight >= 600;
    if (st.fontStyle) f.i = st.fontStyle === "italic" || st.fontStyle === "oblique";
    const deco = st.textDecorationLine || st.textDecoration;
    if (deco) { f.u = /underline/.test(deco); f.s = /line-through/.test(deco); }
    if (tag === "FONT" && n.getAttribute("color")) f.color = rgbToHex(n.getAttribute("color"));
    if (st.color) f.color = rgbToHex(st.color);
    if (st.backgroundColor) f.hl = rgbToHex(st.backgroundColor);
    if (n.dataset.sz) f.size = +n.dataset.sz;
    else if (st.fontSize && /px$/.test(st.fontSize) && !f.sup && !f.sub) f.size = parseFloat(st.fontSize);
    if (n.dataset.font) f.font = n.dataset.font;
    else if (tag === "FONT" && n.getAttribute("face")) f.font = firstFamily(n.getAttribute("face"));
    else if (st.fontFamily) f.font = firstFamily(st.fontFamily);
    if (tag === "A") f.link = n.getAttribute("href") || n.getAttribute("data-href") || undefined;
  }
  const r = {};
  for (const k of ["b", "i", "u", "s", "sup", "sub"]) if (f[k]) r[k] = true;
  if (f.color && f.color !== defaults.color) r.color = f.color;
  if (f.hl) r.hl = f.hl;
  if (f.size > 0) r.size = Math.round(f.size * 10) / 10;
  if (f.font && f.font !== defaults.font && !/^(inherit|initial)$/i.test(f.font)) r.font = f.font;
  if (f.link) r.link = f.link;
  return r;
}

function sameRunFmt(a, b) {
  for (const k of ["b", "i", "u", "s", "sup", "sub", "color", "hl", "size", "font", "link"]) if (a[k] !== b[k]) return false;
  return true;
}

function runsOfBlock(block, defaults) {
  const runs = [];
  const nodes = [];
  const walk = (n) => {
    for (const c of n.childNodes) {
      if (c.nodeType === 3) nodes.push(c);
      else if (c.nodeType === 1) {
        if (c.tagName === "BR") nodes.push(c);
        else if (!BLOCK_RE.test(c.tagName)) walk(c);
      }
    }
  };
  walk(block);
  // a trailing <br> is only the placeholder that keeps an empty line open
  let last = nodes.length - 1;
  while (last >= 0 && nodes[last].nodeType === 3 && nodes[last].data === "") last--;
  if (last >= 0 && nodes[last].nodeType === 1) nodes.splice(last, 1);
  for (const n of nodes) {
    const isBr = n.nodeType === 1;
    const t = isBr ? "\n" : n.data.replace(/ /g, " ").replace(/​/g, "");
    if (!t) continue;
    const f = fmtOf(isBr ? n : n, block, defaults);
    const run = { t, ...f };
    const prev = runs[runs.length - 1];
    if (prev && sameRunFmt(prev, run)) prev.t += t;
    else runs.push(run);
  }
  return runs;
}

function paraFromBlock(block, defaults, extra = {}) {
  const p = { runs: runsOfBlock(block, defaults) };
  const ds = block.dataset || {};
  const align = block.style.textAlign || ds.align;
  if (align && /^(center|right|justify)$/.test(align)) p.align = align;
  const lvl = Number(ds.level) || 0;
  if (lvl > 0) p.level = Math.min(8, lvl);
  if (ds.bullet === "ul" || ds.bullet === "ol") p.bullet = ds.bullet;
  Object.assign(p, extra);
  if (!p.runs.length && ds.psz) p.runs = [{ t: "", size: +ds.psz }];
  return p;
}

/** Parse an edited .sl-tx element back into paragraphs. */
export function parseTx(tx, defaults = {}) {
  const paras = [];
  let loose = null;
  const handleBlock = (el, extra) => {
    if (el.tagName === "UL" || el.tagName === "OL") {
      for (const li of el.children) handleBlock(li, { bullet: el.tagName === "OL" ? "ol" : "ul" });
      return;
    }
    const hasBlockKids = [...el.children].some((c) => BLOCK_RE.test(c.tagName));
    if (!hasBlockKids) { paras.push(paraFromBlock(el, defaults, extra)); return; }
    let buf = null;
    for (const c of el.childNodes) {
      if (c.nodeType === 1 && BLOCK_RE.test(c.tagName)) { buf = null; handleBlock(c, extra); }
      else {
        if (!buf) { buf = document.createElement("p"); buf.dataset.tmp = "1"; }
        buf.appendChild(c.cloneNode(true));
        if (!c.nextSibling || (c.nextSibling.nodeType === 1 && BLOCK_RE.test(c.nextSibling.tagName))) {
          const p = paraFromBlock(buf, defaults, extra);
          if (p.runs.length) paras.push(p);
        }
      }
    }
  };
  for (const c of [...tx.childNodes]) {
    if (c.nodeType === 1 && BLOCK_RE.test(c.tagName)) { loose = null; handleBlock(c); continue; }
    if (c.nodeType === 3 && !c.data.trim() && !loose) continue;
    if (!loose) { loose = document.createElement("p"); paras.push(null); loose.__index = paras.length - 1; }
    loose.appendChild(c.cloneNode(true));
    paras[loose.__index] = paraFromBlock(loose, defaults);
  }
  const out = paras.filter(Boolean);
  return out.length ? out : [{ runs: [] }];
}

function textNodes(root) {
  const out = [];
  const tw = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  for (let n = tw.nextNode(); n; n = tw.nextNode()) out.push(n);
  return out;
}

/** A collapsed range inside a text node at end (or start) of a paragraph. */
function caretIn(p, atEnd) {
  const r = document.createRange();
  const ts = textNodes(p);
  if (ts.length) {
    const t = atEnd ? ts[ts.length - 1] : ts[0];
    r.setStart(t, atEnd ? t.length : 0);
  } else {
    r.selectNodeContents(p);
  }
  r.collapse(true);
  return r;
}

/**
 * Caret position for a click inside editable `el`. Clicks right of a line's text
 * (or beside a paragraph) land at the end of that line, like in PowerPoint.
 */
function caretFromPoint(el, x, y) {
  let r = null;
  if (document.caretRangeFromPoint) r = document.caretRangeFromPoint(x, y);
  else if (document.caretPositionFromPoint) {
    const cp = document.caretPositionFromPoint(x, y);
    if (cp) { r = document.createRange(); r.setStart(cp.offsetNode, cp.offset); r.collapse(true); }
  }
  if (r && el.contains(r.startContainer) && r.startContainer.nodeType === 3) return r;
  const ps = [...el.querySelectorAll("p")];
  if (!ps.length) return null;
  let best = ps[ps.length - 1];
  for (const p of ps) {
    const b = p.getBoundingClientRect();
    if (y >= b.top && y <= b.bottom) { best = p; break; }
    if (y < b.top) { best = p; break; }
  }
  const tr = document.createRange();
  tr.selectNodeContents(best);
  const tb = tr.getBoundingClientRect();
  return caretIn(best, !(tb.width && x < tb.left + 2));
}

/** Re-apply indentation / bullet numbering styles after paragraph attributes changed. */
function restyleParas(tx) {
  const counters = [];
  for (const p of tx.querySelectorAll("p")) {
    const level = Number(p.dataset.level) || 0;
    const bullet = p.dataset.bullet;
    const indent = level * 36 + (bullet ? 22 : 0);
    p.style.paddingLeft = indent ? `${indent}px` : "";
    p.style.textIndent = bullet ? "-22px" : "";
    if (bullet === "ol") {
      counters.length = level + 1;
      counters[level] = (counters[level] || 0) + 1;
      p.dataset.num = `${counters[level]}.`;
    } else {
      counters.length = level;
      delete p.dataset.num;
    }
  }
}

// ---------------------------------------------------------------------------
// the view
// ---------------------------------------------------------------------------

export class SlidesView {
  constructor(opts) {
    this.o = opts || {};
    this.root = this.o.root;
    this.current = 0;
    this.sel = [];
    this.undoStack = [];
    this.redoStack = [];
    this.lastUndoKey = null;
    this.lastUndoTime = 0;
    this.zoomMode = "fit";
    this.scale = 1;
    this.clip = null;
    this.pasteN = 0;
    this.editing = null;
    this.listeners = [];
    this.timers = new Set();
    this.show = null;
    this.pop = null;
    this.thumbTimer = null;
    this.destroyed = false;
    this.notesKey = null;
  }

  get deck() { return this.o.getModel ? this.o.getModel() : null; }
  get selection() { return [...this.sel]; }

  on(target, type, fn, opts) {
    target.addEventListener(type, fn, opts);
    this.listeners.push([target, type, fn, opts]);
  }

  later(fn, ms) {
    const t = setTimeout(() => { this.timers.delete(t); if (!this.destroyed) fn(); }, ms);
    this.timers.add(t);
    return t;
  }

  status(msg) { if (this.o.onStatus) this.o.onStatus(msg); }
  dirty() { if (this.o.onDirty) this.o.onDirty(true); }
  stats() { if (this.o.onStats) this.o.onStats(); }

  // ------------------------------------------------------------------ build

  render() {
    const deck = this.deck;
    this.teardownDom();
    const root = this.root;
    if (!root) return;
    root.classList.add("sv-root");
    if (!deck) { root.textContent = "No presentation is open."; return; }
    if (this.current >= deck.slides.length) this.current = deck.slides.length - 1;
    root.innerHTML = this.shellHtml();
    this.q = (s) => root.querySelector(s);
    this.ribbon = this.q(".sv-ribbon");
    this.strip = this.q(".sv-strip");
    this.thumbs = this.q(".sv-thumbs");
    this.scroll = this.q(".sv-scroll");
    this.canvas = this.q(".sv-canvas");
    this.host = this.q(".sv-host");
    this.overlay = this.q(".sv-overlay");
    this.notesEl = this.q(".sv-notes");
    this.fileInput = this.q(".sv-file");
    this.popEl = this.q(".sv-pop");
    this.bindOnce();
    this.renderStrip();
    this.renderCanvas();
    this.updateToolbar();
    this.stats();
    window.__slidesView = this; // test / automation hook
  }

  shellHtml() {
    const b = (cmd, ic, title, extra = "") => `<button type="button" class="sv-b" data-cmd="${cmd}" title="${esc(title)}" aria-label="${esc(title.replace(/ \(.*\)$/, ""))}"${extra}>${icon(ic)}</button>`;
    const bt = (cmd, ic, label, title, extra = "") => `<button type="button" class="sv-b sv-bt" data-cmd="${cmd}" title="${esc(title)}"${extra}>${icon(ic)}<span>${label}</span></button>`;
    const pop = (name, ic, label, title) => `<button type="button" class="sv-b sv-bt sv-pb" data-pop="${name}" title="${esc(title)}" aria-haspopup="true">${ic ? icon(ic) : ""}${label ? `<span>${label}</span>` : ""}<span class="sv-caret">▾</span></button>`;
    const mod = isMac ? "⌘" : "Ctrl+";
    const fonts = FONTS.map((f) => `<option value="${esc(f)}">${esc(f)}</option>`).join("");
    return `
<div class="sv-ribbon" role="toolbar" aria-label="Presentation tools">
  <div class="sv-g">${b("undo", "undo", `Undo (${mod}Z)`)}${b("redo", "redo", `Redo (${mod}Y)`)}</div>
  <div class="sv-g">
    <div class="sv-split">${bt("newSlide", "plus", "New slide", `New slide (${mod}M)`, ' data-primary="1"')}<button type="button" class="sv-b sv-split-caret" data-pop="layoutNew" title="New slide with a layout" aria-label="Choose a layout for the new slide" aria-haspopup="true">▾</button></div>
    ${b("dupSlide", "dup", `Duplicate slide`)}${b("delSlide", "trash", "Delete slide")}
  </div>
  <div class="sv-g">
    ${bt("textbox", "textbox", "Text box", "Insert a text box")}
    ${pop("shape", "shapes", "Shape", "Insert a shape")}
    ${bt("image", "image", "Picture", "Insert a picture from a file (you can also paste or drop one)")}
    ${pop("table", "table", "Table", "Insert a table")}
  </div>
  <div class="sv-g sv-g-text">
    <select class="sv-font" data-ctl="font" title="Font" aria-label="Font"><option value="">Theme font</option>${fonts}</select>
    <input class="sv-size" data-ctl="size" title="Font size (points)" aria-label="Font size" list="sv-sizes" inputmode="decimal" autocomplete="off">
    <datalist id="sv-sizes">${SIZES_PT.map((s) => `<option value="${s}">`).join("")}</datalist>
    <button type="button" class="sv-b sv-tx" data-cmd="bold" title="Bold (${mod}B)" aria-label="Bold"><b>B</b></button>
    <button type="button" class="sv-b sv-tx" data-cmd="italic" title="Italic (${mod}I)" aria-label="Italic"><i>I</i></button>
    <button type="button" class="sv-b sv-tx" data-cmd="underline" title="Underline (${mod}U)" aria-label="Underline"><u>U</u></button>
    <button type="button" class="sv-b sv-tx" data-cmd="strike" title="Strikethrough" aria-label="Strikethrough"><s>S</s></button>
    <button type="button" class="sv-b sv-pb sv-color" data-pop="textColor" title="Text color" aria-label="Text color" aria-haspopup="true"><span class="sv-a">A</span><span class="sv-bar" data-bar="text"></span><span class="sv-caret">▾</span></button>
  </div>
  <div class="sv-g sv-g-para">
    ${b("alignLeft", "alignL", "Align left")}${b("alignCenter", "alignC", "Center")}${b("alignRight", "alignR", "Align right")}${b("alignJustify", "alignJ", "Justify")}
    ${b("bullets", "bullets", "Bullets")}${b("numbers", "numbers", "Numbering")}
    ${b("outdent", "outdent", "Decrease list level (Shift+Tab)")}${b("indent", "indent", "Increase list level (Tab)")}
  </div>
  <div class="sv-g sv-g-shape">
    <button type="button" class="sv-b sv-pb sv-color" data-pop="fill" title="Fill color" aria-label="Fill color" aria-haspopup="true">${icon("fill")}<span class="sv-bar" data-bar="fill"></span><span class="sv-caret">▾</span></button>
    <button type="button" class="sv-b sv-pb sv-color" data-pop="outline" title="Outline color and weight" aria-label="Outline" aria-haspopup="true">${icon("outline")}<span class="sv-bar" data-bar="line"></span><span class="sv-caret">▾</span></button>
    ${pop("arrange", "arrange", "Arrange", "Order, align and position objects")}
  </div>
  <div class="sv-g sv-g-slide">
    ${pop("layout", "layout", "Layout", "Change the layout of this slide")}
    ${pop("background", "background", "Background", "Slide background color")}
    ${pop("transition", "transition", "Transition", "Transition when this slide appears")}
    ${pop("theme", "theme", "Theme", "Theme: colors and fonts for the whole presentation")}
    ${pop("slideSize", "size", "Size", "Slide size")}
  </div>
  <div class="sv-g sv-g-end">
    ${pop("zoom", "zoom", "Fit", "Zoom")}
    <div class="sv-split sv-present">${bt("present", "play", "Present", "Present from the beginning (F5)")}<button type="button" class="sv-b sv-split-caret" data-pop="present" title="More slide show options" aria-label="More slide show options" aria-haspopup="true">▾</button></div>
  </div>
</div>
<div class="sv-body">
  <div class="sv-strip" aria-label="Slides">
    <div class="sv-thumbs" role="listbox" tabindex="0" aria-label="Slides (drag to reorder)"></div>
    <button type="button" class="sv-add-slide" data-cmd="newSlide" title="New slide (${mod}M)">${icon("plus")}<span>New slide</span></button>
  </div>
  <div class="sv-main">
    <div class="sv-scroll" tabindex="-1">
      <div class="sv-canvas" tabindex="-1" aria-label="Slide editor">
        <div class="sv-host"></div>
        <div class="sv-overlay"></div>
      </div>
    </div>
    <div class="sv-notes-wrap">
      <div class="sv-notes-grip" title="Drag to resize the notes" aria-hidden="true"></div>
      <textarea class="sv-notes" placeholder="Click to add speaker notes" aria-label="Speaker notes" spellcheck="true"></textarea>
    </div>
  </div>
</div>
<input type="file" class="sv-file" accept="image/*" multiple hidden>
<div class="sv-pop" hidden></div>`;
  }

  bindOnce() {
    const root = this.root;
    this.on(this.ribbon, "mousedown", (e) => {
      // keep the text selection while clicking toolbar buttons
      const b = e.target.closest("button");
      if (b && this.editing) e.preventDefault();
    });
    this.on(this.ribbon, "click", (e) => {
      const b = e.target.closest("button");
      if (!b || b.disabled) return;
      if (b.dataset.pop) { this.togglePop(b.dataset.pop, b); return; }
      if (b.dataset.cmd) { this.closePop(); this.cmd(b.dataset.cmd); }
    });
    this.on(this.q(".sv-add-slide"), "click", () => this.cmd("newSlide"));
    this.on(this.popEl, "mousedown", (e) => { if (this.editing && e.target.closest("button")) e.preventDefault(); });
    const font = this.q(".sv-font");
    const size = this.q(".sv-size");
    this.on(font, "change", () => { this.applyFont(font.value); });
    this.on(size, "change", () => { const v = parseFloat(size.value); if (v > 0 && v <= 400) this.applySize(v); });
    this.on(size, "keydown", (e) => {
      if (e.key === "Enter") { e.preventDefault(); const v = parseFloat(size.value); if (v > 0 && v <= 400) this.applySize(v); this.refocus(); }
      if (e.key === "Escape") { e.preventDefault(); this.refocus(); }
    });
    for (const ctl of [font, size]) {
      this.on(ctl, "mousedown", () => { if (this.editing) this.saveRange(); });
      this.on(ctl, "focus", () => { if (this.editing) this.saveRange(); });
    }

    // filmstrip
    this.on(this.thumbs, "pointerdown", (e) => this.onThumbDown(e));
    this.on(this.thumbs, "contextmenu", (e) => {
      const t = e.target.closest(".sv-thumb");
      if (!t) return;
      e.preventDefault();
      this.selectSlide(+t.dataset.s);
      this.openMenu(this.slideMenuItems(), e.clientX, e.clientY);
    });
    this.on(this.thumbs, "keydown", (e) => this.onThumbKey(e));

    // canvas
    this.on(this.scroll, "pointerdown", (e) => this.onCanvasDown(e));
    this.on(this.scroll, "dblclick", (e) => this.onCanvasDbl(e));
    this.on(this.scroll, "contextmenu", (e) => this.onCanvasContext(e));
    this.on(this.scroll, "wheel", (e) => {
      if (!(e.ctrlKey || e.metaKey)) return;
      e.preventDefault();
      const z = Math.max(0.1, Math.min(4, this.scale * (e.deltaY < 0 ? 1.1 : 1 / 1.1)));
      this.setZoom(z);
    }, { passive: false });
    this.on(this.scroll, "click", (e) => {
      const a = e.target.closest("a[href]");
      if (a && this.root.contains(a)) {
        e.preventDefault();
        if (e.ctrlKey || e.metaKey) window.open(a.href, "_blank", "noopener");
        else this.status(`Link: ${a.getAttribute("href")} (Ctrl+click to open)`);
      }
    });
    for (const t of ["dragenter", "dragover"]) {
      this.on(this.scroll, t, (e) => {
        const dt = e.dataTransfer;
        if (!dt || ![...(dt.types || [])].includes("Files")) return;
        const items = [...(dt.items || [])];
        if (items.length && !items.some((it) => /^image\//.test(it.type))) return;
        e.preventDefault();
        e.stopPropagation();
        dt.dropEffect = "copy";
        const hint = document.getElementById("drop-hint");
        if (hint) hint.hidden = true;
        this.canvas.classList.add("sv-dropping");
      });
    }
    this.on(this.scroll, "dragleave", () => this.canvas.classList.remove("sv-dropping"));
    this.on(this.scroll, "drop", (e) => {
      this.canvas.classList.remove("sv-dropping");
      const files = [...((e.dataTransfer && e.dataTransfer.files) || [])].filter((f) => /^image\//.test(f.type));
      if (!files.length) return;
      e.preventDefault();
      e.stopPropagation();
      const p = this.toSlide(e.clientX, e.clientY);
      this.insertImages(files, p);
    });

    // notes
    this.on(this.notesEl, "input", () => {
      const s = this.slide();
      if (!s) return;
      this.pushUndo(`notes:${this.current}`);
      s.notes = this.notesEl.value.replace(/\r\n?/g, "\n");
      if (!s.notes) delete s.notes;
      this.dirty();
    });
    this.on(this.notesEl, "blur", () => { this.lastUndoKey = null; });
    this.on(this.q(".sv-notes-grip"), "pointerdown", (e) => this.onNotesGrip(e));

    this.on(this.fileInput, "change", () => {
      const files = [...this.fileInput.files];
      this.fileInput.value = "";
      if (files.length) this.insertImages(files);
    });

    // document / window
    this.on(document, "keydown", (e) => this.onKey(e));
    this.on(document, "copy", (e) => this.onCopyEvent(e, false));
    this.on(document, "cut", (e) => this.onCopyEvent(e, true));
    this.on(document, "paste", (e) => this.onPaste(e));
    this.on(document, "selectionchange", () => { if (this.editing) { this.saveRange(); this.updateToolbarSoon(); } });
    this.on(document, "pointerdown", (e) => {
      if (this.pop && !this.popEl.contains(e.target) && !e.target.closest("[data-pop]")) this.closePop();
    }, true);
    this.on(window, "afterprint", () => this.afterPrint());
    this.on(document, "fullscreenchange", () => {
      if (this.show && this.show.fs && !document.fullscreenElement) this.endShow();
    });
    if (typeof ResizeObserver !== "undefined") {
      this.ro = new ResizeObserver(() => {
        if (this.zoomMode === "fit") this.fit();
        if (this.thumbs && this.thumbW() !== this.lastThumbW) this.renderStrip();
      });
      this.ro.observe(this.scroll);
    } else {
      this.on(window, "resize", () => { if (this.zoomMode === "fit") this.fit(); });
    }
    if (typeof IntersectionObserver !== "undefined") {
      this.io = new IntersectionObserver((ents) => {
        for (const en of ents) if (en.isIntersecting) { this.paintThumb(+en.target.dataset.s); this.io.unobserve(en.target); }
      }, { root: this.thumbs, rootMargin: "300px 0px" });
    }
    void root;
  }

  teardownDom() {
    this.closePop();
    if (this.show) this.endShow();
    for (const [t, type, fn, opts] of this.listeners) t.removeEventListener(type, fn, opts);
    this.listeners = [];
    if (this.ro) { this.ro.disconnect(); this.ro = null; }
    if (this.io) { this.io.disconnect(); this.io = null; }
    this.editing = null;
  }

  destroy() {
    this.destroyed = true;
    this.teardownDom();
    for (const t of this.timers) clearTimeout(t);
    this.timers.clear();
    if (this.printEl) { this.printEl.remove(); this.printEl = null; }
    if (this.root) { this.root.innerHTML = ""; this.root.classList.remove("sv-root"); }
    if (window.__slidesView === this) delete window.__slidesView;
  }

  focus() {
    if (!this.canvas) return;
    try { this.canvas.focus({ preventScroll: true }); } catch { /* ignore */ }
  }

  refocus() {
    if (this.editing) {
      try { this.editing.el.focus({ preventScroll: true }); } catch { /* ignore */ }
      this.restoreRange();
    } else this.focus();
  }

  getStats() {
    const deck = this.deck;
    return { slide: this.current + 1, slides: deck ? deck.slides.length : 0 };
  }

  slide() { const d = this.deck; return d ? d.slides[this.current] : null; }

  // ------------------------------------------------------------------ undo

  snapshot() {
    const d = this.deck;
    return { slides: cloneDeep(d.slides), theme: cloneDeep(d.theme), size: { ...d.size }, current: this.current, sel: [...this.sel] };
  }

  /** Record the state before a change. Calls with the same key in quick succession coalesce. */
  pushUndo(key = null, snap = null) {
    const now = Date.now();
    if (key && key === this.lastUndoKey && now - this.lastUndoTime < (key.startsWith("notes") ? 4000 : NUDGE_COALESCE_MS)) {
      this.lastUndoTime = now;
      return;
    }
    this.undoStack.push(snap || this.snapshot());
    if (this.undoStack.length > UNDO_CAP) this.undoStack.shift();
    this.redoStack = [];
    this.lastUndoKey = key;
    this.lastUndoTime = now;
    this.updateUndoButtons();
  }

  restore(s) {
    const d = this.deck;
    d.slides = s.slides;
    d.theme = s.theme;
    d.size = s.size;
    this.current = Math.max(0, Math.min(d.slides.length - 1, s.current));
    const n = d.slides[this.current].items.length;
    this.sel = s.sel.filter((i) => i < n);
    this.renderStrip();
    this.renderCanvas();
    this.dirty();
    this.stats();
  }

  undo() {
    if (this.editing) this.commitEdit();
    const s = this.undoStack.pop();
    if (!s) { this.status("Nothing to undo"); return; }
    this.redoStack.push(this.snapshot());
    this.lastUndoKey = null;
    this.restore(s);
    this.status("Undone");
  }

  redo() {
    if (this.editing) this.commitEdit();
    const s = this.redoStack.pop();
    if (!s) { this.status("Nothing to redo"); return; }
    this.undoStack.push(this.snapshot());
    this.lastUndoKey = null;
    this.restore(s);
    this.status("Redone");
  }

  updateUndoButtons() {
    if (!this.ribbon) return;
    const u = this.ribbon.querySelector('[data-cmd="undo"]');
    const r = this.ribbon.querySelector('[data-cmd="redo"]');
    if (u) u.disabled = !this.undoStack.length && !this.editing;
    if (r) r.disabled = !this.redoStack.length;
  }

  /** After a model change: re-render what is needed and report. */
  changed({ all = false, strip = false } = {}) {
    if (all || strip) this.renderStrip();
    else this.refreshThumbSoon(this.current);
    this.renderCanvas();
    this.dirty();
    if (strip || all) this.stats();
  }

  // ------------------------------------------------------------------ filmstrip

  thumbW() {
    const w = this.thumbs ? this.thumbs.clientWidth : 0;
    return Math.max(90, Math.min(220, (w || 190) - 40));
  }

  renderStrip() {
    const deck = this.deck;
    if (!this.thumbs) return;
    if (this.io) this.io.disconnect();
    const W = deck.size.w, H = deck.size.h;
    const tw = this.thumbW();
    this.lastThumbW = tw;
    const k = tw / W;
    const th = Math.round(H * k);
    let html = "";
    deck.slides.forEach((s, i) => {
      const title = slideTitle(s) || `Slide ${i + 1}`;
      html += `<div class="sv-thumb${i === this.current ? " sel" : ""}${s.hidden ? " hidden-slide" : ""}" data-s="${i}" role="option" aria-selected="${i === this.current}" aria-label="${esc(`Slide ${i + 1}: ${title}${s.hidden ? " (hidden)" : ""}`)}">`
        + `<div class="sv-thumb-num">${i + 1}${s.hidden ? `<span class="sv-hid" title="Hidden in the slide show">${icon("eyeOff")}</span>` : ""}${s.transition && s.transition !== "none" ? '<span class="sv-trn" title="Has a transition">★</span>' : ""}</div>`
        + `<div class="sv-thumb-box" style="width:${tw}px;height:${th}px"><div class="sv-thumb-scale" style="transform:scale(${k})"></div></div></div>`;
    });
    this.thumbs.innerHTML = html;
    const nodes = this.thumbs.querySelectorAll(".sv-thumb");
    nodes.forEach((n, i) => {
      if (this.io && i > 12) this.io.observe(n);
      else this.paintThumb(i);
    });
    const cur = nodes[this.current];
    if (cur && cur.scrollIntoView) cur.scrollIntoView({ block: "nearest" });
  }

  paintThumb(i) {
    const deck = this.deck;
    const n = this.thumbs && this.thumbs.querySelector(`.sv-thumb[data-s="${i}"] .sv-thumb-scale`);
    if (!n || !deck.slides[i]) return;
    n.innerHTML = renderSlideHtml(deck.slides[i], deck);
  }

  refreshThumbSoon(i) {
    if (this.thumbTimer) { clearTimeout(this.thumbTimer); this.timers.delete(this.thumbTimer); }
    this.thumbTimer = this.later(() => { this.thumbTimer = null; this.paintThumb(i); }, 120);
  }

  selectSlide(i, { focusStrip = false } = {}) {
    const deck = this.deck;
    if (!deck) return;
    i = Math.max(0, Math.min(deck.slides.length - 1, i));
    if (this.editing) this.commitEdit();
    const changedSlide = i !== this.current;
    this.current = i;
    this.sel = [];
    this.lastUndoKey = null;
    for (const t of this.thumbs.querySelectorAll(".sv-thumb")) {
      const on = +t.dataset.s === i;
      t.classList.toggle("sel", on);
      t.setAttribute("aria-selected", String(on));
      if (on && t.scrollIntoView) t.scrollIntoView({ block: "nearest" });
    }
    this.renderCanvas();
    if (focusStrip) this.thumbs.focus({ preventScroll: true });
    if (changedSlide) this.stats();
  }

  onThumbDown(e) {
    if (e.button !== 0) return;
    const t = e.target.closest(".sv-thumb");
    if (!t) return;
    const from = +t.dataset.s;
    this.selectSlide(from);
    this.thumbs.focus({ preventScroll: true });
    const startY = e.clientY;
    let dragging = false;
    let marker = null;
    let target = from;
    const move = (ev) => {
      if (!dragging && Math.abs(ev.clientY - startY) > 6) {
        dragging = true;
        marker = document.createElement("div");
        marker.className = "sv-drop-marker";
        this.thumbs.appendChild(marker);
        t.classList.add("dragging");
      }
      if (!dragging) return;
      const list = [...this.thumbs.querySelectorAll(".sv-thumb")];
      target = list.length;
      for (const n of list) {
        const r = n.getBoundingClientRect();
        if (ev.clientY < r.top + r.height / 2) { target = +n.dataset.s; break; }
      }
      const ref = list[target];
      const box = this.thumbs.getBoundingClientRect();
      const y = ref ? ref.getBoundingClientRect().top - 4 : list[list.length - 1].getBoundingClientRect().bottom + 2;
      marker.style.top = `${y - box.top + this.thumbs.scrollTop}px`;
      // auto-scroll near the edges
      if (ev.clientY < box.top + 30) this.thumbs.scrollTop -= 12;
      else if (ev.clientY > box.bottom - 30) this.thumbs.scrollTop += 12;
    };
    const up = () => {
      window.removeEventListener("pointermove", move);
      window.removeEventListener("pointerup", up);
      window.removeEventListener("pointercancel", up);
      if (marker) marker.remove();
      t.classList.remove("dragging");
      if (dragging) {
        let to = target > from ? target - 1 : target;
        if (to !== from) this.moveSlide(from, to);
      }
    };
    window.addEventListener("pointermove", move);
    window.addEventListener("pointerup", up);
    window.addEventListener("pointercancel", up);
  }

  onThumbKey(e) {
    if (e.target !== this.thumbs) return;
    const mod = e.ctrlKey || e.metaKey;
    const deck = this.deck;
    if (e.key === "ArrowDown" || e.key === "ArrowUp") {
      e.preventDefault();
      const d = e.key === "ArrowDown" ? 1 : -1;
      if (mod) this.moveSlide(this.current, this.current + d);
      else this.selectSlide(this.current + d, { focusStrip: true });
    } else if (e.key === "Home") { e.preventDefault(); this.selectSlide(0, { focusStrip: true }); }
    else if (e.key === "End") { e.preventDefault(); this.selectSlide(deck.slides.length - 1, { focusStrip: true }); }
    else if (e.key === "Delete" || e.key === "Backspace") { e.preventDefault(); this.deleteSlide(); }
    else if (e.key === "Enter") { e.preventDefault(); this.cmd("newSlide"); }
    else return;
    e.stopPropagation();
  }

  slideMenuItems() {
    const s = this.slide();
    const n = this.deck.slides.length;
    return [
      { label: "New slide", ic: "plus", run: () => this.cmd("newSlide") },
      { label: "Duplicate slide", ic: "dup", run: () => this.cmd("dupSlide") },
      { label: "Delete slide", ic: "trash", run: () => this.cmd("delSlide"), disabled: n <= 1 },
      "-",
      { label: s && s.hidden ? "Show slide" : "Hide slide", ic: s && s.hidden ? "eye" : "eyeOff", run: () => this.cmd("hideSlide") },
      { label: "Move up", ic: "up", run: () => this.moveSlide(this.current, this.current - 1), disabled: this.current === 0 },
      { label: "Move down", ic: "down", run: () => this.moveSlide(this.current, this.current + 1), disabled: this.current >= n - 1 },
      "-",
      { label: "Copy slide", run: () => { this.copySlide(); } },
      { label: "Paste", run: () => this.pasteInternal(), disabled: !this.clip },
      "-",
      { label: "Present from this slide", ic: "play", run: () => this.startShow(this.current) },
    ];
  }

  moveSlide(from, to) {
    const deck = this.deck;
    if (to < 0 || to >= deck.slides.length || from === to) return;
    this.pushUndo();
    const [s] = deck.slides.splice(from, 1);
    deck.slides.splice(to, 0, s);
    this.current = to;
    this.sel = [];
    this.changed({ strip: true });
    this.status(`Slide moved to position ${to + 1}`);
  }

  addSlide(layout) {
    const deck = this.deck;
    if (this.editing) this.commitEdit();
    this.pushUndo();
    const cur = this.slide();
    let lay = layout;
    if (!lay) lay = cur && cur.layout === "title" ? "content" : (cur && cur.layout) || "content";
    const s = newSlide(lay, deck);
    if (cur && cur.transition) s.transition = cur.transition;
    deck.slides.splice(this.current + 1, 0, s);
    this.current += 1;
    this.sel = [];
    this.changed({ strip: true });
    this.status(`Slide ${this.current + 1} added`);
  }

  duplicateSlide() {
    const deck = this.deck;
    if (this.editing) this.commitEdit();
    this.pushUndo();
    deck.slides.splice(this.current + 1, 0, cloneDeep(this.slide()));
    this.current += 1;
    this.sel = [];
    this.changed({ strip: true });
    this.status("Slide duplicated");
  }

  deleteSlide() {
    const deck = this.deck;
    if (deck.slides.length <= 1) { this.status("A presentation needs at least one slide"); return; }
    if (this.editing) this.commitEdit();
    this.pushUndo();
    deck.slides.splice(this.current, 1);
    this.current = Math.min(this.current, deck.slides.length - 1);
    this.sel = [];
    this.changed({ strip: true });
    this.status("Slide deleted (Ctrl+Z to undo)");
  }

  copySlide() {
    this.clip = { slides: [cloneDeep(this.slide())] };
    this.status("Slide copied");
  }

  // ------------------------------------------------------------------ canvas

  fitScale() {
    const deck = this.deck;
    const r = this.scroll.getBoundingClientRect();
    const pad = 28;
    const k = Math.min((r.width - pad * 2) / deck.size.w, (r.height - pad * 2) / deck.size.h);
    return Math.max(0.05, Number.isFinite(k) && k > 0 ? k : 0.5);
  }

  fit() {
    if (!this.scroll) return;
    const k = this.fitScale();
    if (Math.abs(k - this.scale) > 0.001 || !this.canvas.style.width) {
      this.scale = k;
      this.applyScale();
    }
  }

  setZoom(z) {
    if (z === "fit") { this.zoomMode = "fit"; this.scale = this.fitScale(); }
    else { this.zoomMode = "manual"; this.scale = z; }
    this.applyScale();
  }

  applyScale() {
    const deck = this.deck;
    const k = this.scale;
    this.canvas.style.width = `${deck.size.w * k}px`;
    this.canvas.style.height = `${deck.size.h * k}px`;
    this.host.style.transform = `scale(${k})`;
    this.canvas.style.setProperty("--k", k);
    const zb = this.ribbon && this.ribbon.querySelector('[data-pop="zoom"] span');
    if (zb) zb.textContent = this.zoomMode === "fit" ? "Fit" : `${Math.round(k * 100)}%`;
    this.drawOverlay();
  }

  renderCanvas() {
    const deck = this.deck;
    const s = this.slide();
    if (!s || !this.host) return;
    this.host.style.width = `${deck.size.w}px`;
    this.host.style.height = `${deck.size.h}px`;
    this.host.innerHTML = renderSlideHtml(s, deck, { edit: true });
    this.sel = this.sel.filter((i) => i < s.items.length);
    if (this.zoomMode === "fit") this.scale = this.fitScale();
    this.applyScale();
    if (this.notesEl && document.activeElement !== this.notesEl) this.notesEl.value = s.notes || "";
    this.updateToolbar();
  }

  /** Re-render a single item (after a geometry change) without rebuilding the whole slide. */
  renderItem(i) {
    const s = this.slide();
    const old = this.host.querySelector(`.sl-item[data-i="${i}"]`);
    if (!old || !s.items[i]) { this.renderCanvas(); return; }
    const tmp = document.createElement("div");
    tmp.innerHTML = renderItemHtml(s.items[i], this.deck, i, { edit: true });
    if (tmp.firstElementChild) old.replaceWith(tmp.firstElementChild);
  }

  toSlide(cx, cy) {
    const r = this.canvas.getBoundingClientRect();
    return { x: (cx - r.left) / this.scale, y: (cy - r.top) / this.scale };
  }

  itemAt(target) {
    const n = target && target.closest ? target.closest(".sv-host .sl-item") : null;
    return n ? +n.dataset.i : -1;
  }

  drawOverlay(guides = null) {
    if (!this.overlay) return;
    const s = this.slide();
    const k = this.scale;
    let html = "";
    const single = this.sel.length === 1;
    for (const i of this.sel) {
      const it = s.items[i];
      if (!it || typeof it.x !== "number") continue;
      const st = `left:${it.x * k}px;top:${it.y * k}px;width:${Math.max(0, it.w * k)}px;height:${Math.max(0, it.h * k)}px;${it.rot ? `transform:rotate(${it.rot}deg);` : ""}`;
      const editing = this.editing && this.editing.index === i;
      let inner = "";
      if (single) {
        const isLine = it.type === "shape" && it.shape === "line" && it.h < 2;
        const hs = isLine ? ["w", "e"] : ["nw", "n", "ne", "e", "se", "s", "sw", "w"];
        inner += hs.map((h) => `<div class="sv-h sv-h-${h}" data-h="${h}" style="cursor:${this.cursorFor(h, it.rot || 0)}"></div>`).join("");
        if (it.type !== "table") inner += '<div class="sv-rot" data-h="rot" title="Rotate (Shift: 15° steps)"></div>';
      }
      if (editing) inner += '<div class="sv-edge t" data-h="move"></div><div class="sv-edge b" data-h="move"></div><div class="sv-edge l" data-h="move"></div><div class="sv-edge r" data-h="move"></div>';
      html += `<div class="sv-selbox${editing ? " editing" : ""}${single ? " single" : ""}" data-i="${i}" style="${st}">${inner}</div>`;
    }
    if (this.marquee) {
      const m = this.marquee;
      html += `<div class="sv-marquee" style="left:${Math.min(m.x0, m.x1) * k}px;top:${Math.min(m.y0, m.y1) * k}px;width:${Math.abs(m.x1 - m.x0) * k}px;height:${Math.abs(m.y1 - m.y0) * k}px"></div>`;
    }
    for (const g of guides || []) {
      html += g.v != null
        ? `<div class="sv-guide v" style="left:${g.v * k}px"></div>`
        : `<div class="sv-guide h" style="top:${g.h * k}px"></div>`;
    }
    this.overlay.innerHTML = html;
  }

  cursorFor(h, rot) {
    const dirs = ["n", "ne", "e", "se", "s", "sw", "w", "nw"];
    const base = dirs.indexOf(h);
    if (base < 0) return "move";
    const idx = (base + Math.round(((rot % 360) + 360) % 360 / 45)) % 8;
    return { n: "ns-resize", s: "ns-resize", e: "ew-resize", w: "ew-resize", ne: "nesw-resize", sw: "nesw-resize", nw: "nwse-resize", se: "nwse-resize" }[dirs[idx]];
  }

  selectItems(list) {
    this.sel = [...new Set(list)].filter((i) => this.slide().items[i]);
    this.drawOverlay();
    this.updateToolbar();
  }

  onCanvasDown(e) {
    if (e.button !== 0 && e.button !== 2) return;
    if (this.show) return;
    const deck = this.deck;
    const s = this.slide();
    const handle = e.target.closest("[data-h]");
    if (handle && e.button === 0) {
      const box = handle.closest(".sv-selbox");
      const i = +box.dataset.i;
      e.preventDefault();
      if (handle.dataset.h === "rot") this.startRotate(e, i);
      else if (handle.dataset.h === "move") this.startMove(e, i, true);
      else this.startResize(e, i, handle.dataset.h);
      return;
    }
    const i = this.itemAt(e.target);
    // clicks inside the text being edited: let the browser place the caret
    if (this.editing && i === this.editing.index && this.editing.el.contains(e.target)) return;
    if (this.editing && this.editing.cell && i === this.editing.index && e.target.closest("td")) {
      const td = e.target.closest("td");
      this.commitEdit();
      this.beginEdit(i, { cell: [+td.dataset.r, +td.dataset.c], point: { x: e.clientX, y: e.clientY } });
      e.preventDefault();
      return;
    }
    if (this.editing) this.commitEdit();
    if (e.button === 2) {
      if (i >= 0 && !this.sel.includes(i)) this.selectItems([i]);
      return;
    }
    this.focus();
    if (i < 0) {
      if (!e.shiftKey) this.selectItems([]);
      this.startMarquee(e);
      return;
    }
    e.preventDefault();
    const wasSelected = this.sel.includes(i) && this.sel.length === 1;
    if (e.shiftKey || e.ctrlKey || e.metaKey) {
      this.selectItems(this.sel.includes(i) ? this.sel.filter((x) => x !== i) : [...this.sel, i]);
      return;
    }
    if (!this.sel.includes(i)) this.selectItems([i]);
    this.startMove(e, i, false, wasSelected);
    void deck; void s;
  }

  onCanvasDbl(e) {
    if (this.show) return;
    const i = this.itemAt(e.target);
    if (i < 0) return;
    const it = this.slide().items[i];
    if (!it) return;
    if (this.editing && this.editing.index === i) return;
    if (it.type === "text" || (it.type === "shape" && it.shape !== "line")) {
      this.selectItems([i]);
      this.beginEdit(i, { point: { x: e.clientX, y: e.clientY }, word: true });
    } else if (it.type === "table") {
      const td = e.target.closest("td");
      this.selectItems([i]);
      this.beginEdit(i, { cell: td ? [+td.dataset.r, +td.dataset.c] : [0, 0], point: { x: e.clientX, y: e.clientY } });
    } else if (it.type === "image") {
      this.status("Tip: drag the corner handles to resize the picture; hold Shift to keep its shape.");
    }
  }

  onCanvasContext(e) {
    if (this.show) return;
    const i = this.itemAt(e.target);
    if (this.editing && this.editing.el.contains(e.target)) {
      if (this.editing.cell) {
        e.preventDefault();
        this.openMenu(this.tableMenuItems(), e.clientX, e.clientY);
      }
      return; // native text menu (spelling, copy/paste) while typing
    }
    e.preventDefault();
    const has = this.sel.length > 0;
    const items = [];
    if (i >= 0 || has) {
      const it = this.slide().items[i >= 0 ? i : this.sel[0]];
      items.push(
        { label: "Cut", run: () => { this.copyItems(); this.deleteItems(); } },
        { label: "Copy", run: () => this.copyItems() },
        { label: "Paste", run: () => this.pasteInternal(), disabled: !this.clip },
        { label: "Duplicate", ic: "dup", run: () => this.duplicateItems() },
        { label: "Delete", ic: "trash", run: () => this.deleteItems() },
        "-",
      );
      if (it && (it.type === "text" || (it.type === "shape" && it.shape !== "line"))) items.push({ label: "Edit text", run: () => this.beginEdit(this.sel[0], { selectAll: false }) });
      if (it && it.type === "table") items.push(...this.tableMenuItems(), "-");
      items.push(
        { label: "Bring to front", ic: "front", run: () => this.arrange("front") },
        { label: "Send to back", ic: "back", run: () => this.arrange("back") },
      );
    } else {
      items.push(
        { label: "Paste", run: () => this.pasteInternal(), disabled: !this.clip },
        "-",
        { label: "New slide", ic: "plus", run: () => this.cmd("newSlide") },
        { label: "Layout…", ic: "layout", run: () => this.togglePop("layout", this.ribbon.querySelector('[data-pop="layout"]')) },
        { label: "Background…", ic: "background", run: () => this.togglePop("background", this.ribbon.querySelector('[data-pop="background"]')) },
        { label: this.slide().hidden ? "Show slide" : "Hide slide", ic: "eyeOff", run: () => this.cmd("hideSlide") },
      );
    }
    this.openMenu(items, e.clientX, e.clientY);
  }

  /** Drag selected items (snap to slide centre/edges and other objects). */
  startMove(e, i, fromEdge = false, wasSelected = false) {
    const s = this.slide();
    const deck = this.deck;
    const start = this.toSlide(e.clientX, e.clientY);
    const sel = [...this.sel];
    if (!sel.includes(i)) sel.push(i);
    const orig = sel.map((j) => ({ j, x: s.items[j].x, y: s.items[j].y }));
    const bb = this.bbox(sel);
    let moved = false;
    let snap = null;
    const others = s.items.map((it, j) => (sel.includes(j) || !it || typeof it.x !== "number" ? null : it)).filter(Boolean);
    const xs = [0, deck.size.w / 2, deck.size.w], ys = [0, deck.size.h / 2, deck.size.h];
    for (const o of others) { xs.push(o.x, o.x + o.w / 2, o.x + o.w); ys.push(o.y, o.y + o.h / 2, o.y + o.h); }
    const move = (ev) => {
      const p = this.toSlide(ev.clientX, ev.clientY);
      let dx = p.x - start.x, dy = p.y - start.y;
      if (!moved && Math.hypot(dx * this.scale, dy * this.scale) < 3) return;
      if (!moved) {
        moved = true;
        snap = this.snapshot();
        if (this.editing && !fromEdge) this.commitEdit();
        this.root.classList.add("sv-moving");
      }
      if (ev.shiftKey && !fromEdge) { if (Math.abs(dx) > Math.abs(dy)) dy = 0; else dx = 0; }
      const guides = [];
      if (!ev.altKey) {
        const th = 6 / this.scale;
        const cand = (vals, edges) => {
          let best = null;
          for (const e1 of edges) for (const v of vals) { const d = v - e1; if (Math.abs(d) <= th && (!best || Math.abs(d) < Math.abs(best.d))) best = { d, v }; }
          return best;
        };
        const nx = bb.x + dx, ny = bb.y + dy;
        const bx = cand(xs, [nx, nx + bb.w / 2, nx + bb.w]);
        const by = cand(ys, [ny, ny + bb.h / 2, ny + bb.h]);
        if (bx) { dx += bx.d; guides.push({ v: bx.v }); }
        if (by) { dy += by.d; guides.push({ h: by.v }); }
      }
      for (const o of orig) { s.items[o.j].x = r2(o.x + dx); s.items[o.j].y = r2(o.y + dy); }
      for (const o of orig) this.moveItemDom(o.j);
      this.drawOverlay(guides);
    };
    const up = () => {
      window.removeEventListener("pointermove", move);
      window.removeEventListener("pointerup", up);
      window.removeEventListener("pointercancel", up);
      this.root.classList.remove("sv-moving");
      if (moved) {
        this.pushUndo(null, snap);
        this.drawOverlay();
        this.refreshThumbSoon(this.current);
        this.dirty();
        this.updateToolbar();
      } else if (wasSelected && !fromEdge) {
        const it = s.items[i];
        if (it && (it.type === "text" || (it.type === "shape" && it.shape !== "line"))) this.beginEdit(i, { point: { x: e.clientX, y: e.clientY } });
        else if (it && it.type === "table") {
          const td = e.target.closest && e.target.closest("td");
          this.beginEdit(i, { cell: td ? [+td.dataset.r, +td.dataset.c] : [0, 0], point: { x: e.clientX, y: e.clientY } });
        }
      }
    };
    window.addEventListener("pointermove", move);
    window.addEventListener("pointerup", up);
    window.addEventListener("pointercancel", up);
  }

  moveItemDom(j) {
    const it = this.slide().items[j];
    const n = this.host.querySelector(`.sl-item[data-i="${j}"]`);
    if (!n || !it) return;
    n.style.left = `${it.x}px`;
    n.style.top = `${it.y}px`;
  }

  bbox(list) {
    const s = this.slide();
    let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
    for (const j of list) {
      const it = s.items[j];
      if (!it || typeof it.x !== "number") continue;
      x0 = Math.min(x0, it.x); y0 = Math.min(y0, it.y); x1 = Math.max(x1, it.x + it.w); y1 = Math.max(y1, it.y + it.h);
    }
    if (x0 === Infinity) return { x: 0, y: 0, w: 0, h: 0 };
    return { x: x0, y: y0, w: x1 - x0, h: y1 - y0 };
  }

  startResize(e, i, h) {
    const s = this.slide();
    const it = s.items[i];
    const o = { x: it.x, y: it.y, w: it.w, h: it.h };
    const rot = ((it.rot || 0) * Math.PI) / 180;
    const sx = h.includes("e") ? 1 : h.includes("w") ? -1 : 0;
    const sy = h.includes("s") ? 1 : h.includes("n") ? -1 : 0;
    const start = this.toSlide(e.clientX, e.clientY);
    const snap = this.snapshot();
    const aspect = o.h > 0 ? o.w / o.h : 1;
    const keepAspectDefault = it.type === "image";
    let moved = false;
    const minSize = it.type === "shape" && it.shape === "line" ? 0 : 4;
    const move = (ev) => {
      const p = this.toSlide(ev.clientX, ev.clientY);
      const dx = p.x - start.x, dy = p.y - start.y;
      // pointer delta in the item's own (unrotated) frame
      const lx = dx * Math.cos(rot) + dy * Math.sin(rot);
      const ly = -dx * Math.sin(rot) + dy * Math.cos(rot);
      let w = sx ? Math.max(minSize, o.w + sx * lx) : o.w;
      let hh = sy ? Math.max(minSize, o.h + sy * ly) : o.h;
      const keep = (ev.shiftKey ? !keepAspectDefault : keepAspectDefault) && sx && sy && o.h > 0;
      if (keep) {
        if (w / aspect > hh) hh = w / aspect; else w = hh * aspect;
      }
      // keep the opposite edge fixed (in the rotated frame)
      const dcx = sx * (w - o.w) / 2, dcy = sy * (hh - o.h) / 2;
      const cx = o.x + o.w / 2 + dcx * Math.cos(rot) - dcy * Math.sin(rot);
      const cy = o.y + o.h / 2 + dcx * Math.sin(rot) + dcy * Math.cos(rot);
      it.w = r2(w); it.h = r2(hh); it.x = r2(cx - w / 2); it.y = r2(cy - hh / 2);
      if (it.type === "table" && Array.isArray(it.colW) && o.w > 0) {
        const f = it.w / it.colW.reduce((a, b) => a + b, 0);
        it.colW = it.colW.map((c) => r2(c * f));
      }
      moved = true;
      if (!(this.editing && this.editing.index === i)) this.renderItem(i);
      else { const n = this.host.querySelector(`.sl-item[data-i="${i}"]`); Object.assign(n.style, { left: `${it.x}px`, top: `${it.y}px`, width: `${it.w}px`, height: `${it.h}px` }); }
      this.drawOverlay();
      this.status(`Size ${Math.round(it.w)} × ${Math.round(it.h)} pt`);
    };
    const up = () => {
      window.removeEventListener("pointermove", move);
      window.removeEventListener("pointerup", up);
      window.removeEventListener("pointercancel", up);
      if (moved) { this.pushUndo(null, snap); this.refreshThumbSoon(this.current); this.dirty(); }
    };
    window.addEventListener("pointermove", move);
    window.addEventListener("pointerup", up);
    window.addEventListener("pointercancel", up);
  }

  startRotate(e, i) {
    const it = this.slide().items[i];
    const snap = this.snapshot();
    const c = { x: it.x + it.w / 2, y: it.y + it.h / 2 };
    let moved = false;
    const move = (ev) => {
      const p = this.toSlide(ev.clientX, ev.clientY);
      let a = (Math.atan2(p.y - c.y, p.x - c.x) * 180) / Math.PI + 90;
      if (ev.shiftKey) a = Math.round(a / 15) * 15;
      else if (Math.abs(((a % 90) + 90) % 90) < 3 || Math.abs(((a % 90) + 90) % 90) > 87) a = Math.round(a / 90) * 90;
      a = ((a % 360) + 360) % 360;
      if (a < 0.05 || a > 359.95) delete it.rot; else it.rot = r2(a);
      moved = true;
      const n = this.host.querySelector(`.sl-item[data-i="${i}"]`);
      if (n) n.style.transform = it.rot ? `rotate(${it.rot}deg)` : "";
      this.drawOverlay();
      this.status(`Rotation ${Math.round(it.rot || 0)}°`);
    };
    const up = () => {
      window.removeEventListener("pointermove", move);
      window.removeEventListener("pointerup", up);
      if (moved) { this.pushUndo(null, snap); this.refreshThumbSoon(this.current); this.dirty(); }
    };
    window.addEventListener("pointermove", move);
    window.addEventListener("pointerup", up);
  }

  startMarquee(e) {
    const p0 = this.toSlide(e.clientX, e.clientY);
    const base = e.shiftKey ? [...this.sel] : [];
    this.marquee = { x0: p0.x, y0: p0.y, x1: p0.x, y1: p0.y };
    const move = (ev) => {
      const p = this.toSlide(ev.clientX, ev.clientY);
      this.marquee.x1 = p.x; this.marquee.y1 = p.y;
      const m = this.marquee;
      const x0 = Math.min(m.x0, m.x1), x1 = Math.max(m.x0, m.x1), y0 = Math.min(m.y0, m.y1), y1 = Math.max(m.y0, m.y1);
      const hit = [];
      this.slide().items.forEach((it, j) => {
        if (it && typeof it.x === "number" && it.x >= x0 && it.y >= y0 && it.x + it.w <= x1 && it.y + it.h <= y1) hit.push(j);
      });
      this.sel = [...new Set([...base, ...hit])];
      this.drawOverlay();
    };
    const up = () => {
      window.removeEventListener("pointermove", move);
      window.removeEventListener("pointerup", up);
      this.marquee = null;
      this.drawOverlay();
      this.updateToolbar();
    };
    window.addEventListener("pointermove", move);
    window.addEventListener("pointerup", up);
  }

  // ------------------------------------------------------------------ text editing

  textDefaults(it, cellFill) {
    const theme = this.deck.theme || {};
    return {
      color: textColorFor(it, this.deck, cellFill !== undefined ? cellFill : it.type === "shape" ? it.fill : it.fill || null),
      font: it.ph === "title" ? theme.headFont || theme.font : theme.font,
    };
  }

  /** Enter in-place editing of item i (or a table cell). */
  beginEdit(i, { point = null, selectAll = false, cell = null, word = false } = {}) {
    if (this.editing) this.commitEdit();
    const s = this.slide();
    const it = s.items[i];
    if (!it) return false;
    if (it.type === "shape" && it.shape === "line") return false;
    if (it.type === "table" && !cell) cell = [0, 0];
    if (it.type !== "text" && it.type !== "shape" && it.type !== "table") return false;
    // render the item fresh so the DOM matches the model
    this.renderItem(i);
    const node = this.host.querySelector(`.sl-item[data-i="${i}"]`);
    let el = null;
    if (cell) {
      const td = node.querySelector(`td[data-r="${cell[0]}"][data-c="${cell[1]}"]`);
      el = td && td.querySelector(".sl-tx");
    } else el = node.querySelector(".sl-tx");
    if (!el) return false;
    for (const p of el.querySelectorAll("p[data-prompt]")) { p.removeAttribute("data-prompt"); if (!p.firstChild) p.appendChild(document.createElement("br")); }
    const snap = this.snapshot();
    const before = JSON.stringify(cell ? ((it.rows[cell[0]] || [])[cell[1]] || {}).paras || [] : it.paras || []);
    this.editing = { index: i, cell, el, snap, before, node };
    node.classList.add("sv-editing");
    el.contentEditable = "true";
    el.spellcheck = true;
    try { document.execCommand("styleWithCSS", false, false); } catch { /* ignore */ }
    el.focus({ preventScroll: true });
    const sel = window.getSelection();
    let placed = false;
    if (point && !selectAll) {
      const r = caretFromPoint(el, point.x, point.y);
      if (r) {
        sel.removeAllRanges();
        sel.addRange(r);
        placed = true;
        if (word && sel.modify) {
          try { sel.modify("move", "backward", "word"); sel.modify("extend", "forward", "word"); } catch { /* ignore */ }
        }
      }
    }
    if (!placed) {
      const r = document.createRange();
      r.selectNodeContents(el);
      if (!selectAll) {
        // caret at the end of the last text (inside its run), not after the last paragraph
        const ps = el.querySelectorAll("p");
        if (ps.length) { const c = caretIn(ps[ps.length - 1], true); r.setStart(c.startContainer, c.startOffset); r.collapse(true); }
        else r.collapse(false);
      }
      sel.removeAllRanges();
      sel.addRange(r);
    }
    this.saveRange();
    this.drawOverlay();
    this.updateToolbar();
    this.updateUndoButtons();
    this.onEditInput = () => { restyleParas(el); this.growBox(); this.drawOverlay(); };
    el.addEventListener("input", this.onEditInput);
    return true;
  }

  saveRange() {
    if (!this.editing) return;
    const sel = window.getSelection();
    if (sel && sel.rangeCount && this.editing.el.contains(sel.getRangeAt(0).commonAncestorContainer)) this.editing.range = sel.getRangeAt(0).cloneRange();
  }

  restoreRange() {
    if (!this.editing || !this.editing.range) return;
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(this.editing.range);
  }

  /** Flush the in-progress text edit into the model. */
  commitEdit() {
    const ed = this.editing;
    if (!ed) return;
    this.editing = null;
    const s = this.slide();
    const it = s && s.items[ed.index];
    ed.el.contentEditable = "false";
    if (this.onEditInput) { ed.el.removeEventListener("input", this.onEditInput); this.onEditInput = null; }
    if (!it) { this.renderCanvas(); return; }
    const cellFill = ed.cell ? (((it.rows || [])[ed.cell[0]] || [])[ed.cell[1]] || {}).fill || null : undefined;
    const paras = parseTx(ed.el, this.textDefaults(it, cellFill));
    const after = JSON.stringify(paras);
    let changed = after !== ed.before;
    if (changed) {
      if (ed.cell) {
        const row = it.rows[ed.cell[0]];
        if (row && row[ed.cell[1]]) row[ed.cell[1]].paras = paras;
      } else {
        it.paras = paras;
      }
    }
    // an emptied plain text box disappears, like in PowerPoint
    if (!ed.cell && it.type === "text" && !it.ph && !it.fill && !it.line && !paras.some((p) => paraText(p).trim())) {
      s.items.splice(ed.index, 1);
      this.sel = [];
      changed = true;
    }
    if (changed) {
      this.pushUndo(null, ed.snap);
      this.dirty();
      this.refreshThumbSoon(this.current);
    }
    try { window.getSelection().removeAllRanges(); } catch { /* ignore */ }
    this.renderCanvas();
    // text boxes grow to fit their text (PowerPoint's "resize shape to fit text")
    const cur = this.slide().items[ed.index];
    if (cur === it && it.type === "text" && !it.ph && !ed.cell) {
      const need = this.neededHeight(ed.index);
      if (need > it.h + 0.5) {
        it.h = r2(need);
        if (!changed) { this.pushUndo(null, ed.snap); this.dirty(); }
        this.renderCanvas();
        this.refreshThumbSoon(this.current);
      }
    }
    this.updateUndoButtons();
  }

  /** Height (pt) the text of item i needs, from the rendered DOM. */
  neededHeight(i) {
    const node = this.host && this.host.querySelector(`.sl-item[data-i="${i}"]`);
    const tx = node && node.querySelector(".sl-tx");
    if (!tx) return 0;
    return tx.offsetHeight + 7.2; // + top/bottom insets
  }

  /** While typing in a plain text box, let the box grow with its text. */
  growBox() {
    const ed = this.editing;
    if (!ed || ed.cell) return;
    const it = this.slide().items[ed.index];
    if (!it || it.type !== "text" || it.ph) return;
    const need = this.neededHeight(ed.index);
    const n = ed.node;
    if (need > it.h + 0.5) {
      n.style.height = `${need}px`;
      it.h = r2(need); // the snapshot taken when editing began keeps the old size for undo
    }
  }

  /** Apply a function to the paragraphs touched by the current text selection (while editing). */
  editParas(fn) {
    const ed = this.editing;
    if (!ed) return;
    this.restoreRange();
    const sel = window.getSelection();
    const ps = [...ed.el.querySelectorAll("p")];
    let hit = ps;
    if (sel && sel.rangeCount) {
      const r = sel.getRangeAt(0);
      hit = ps.filter((p) => r.intersectsNode(p));
      if (!hit.length && ps.length) hit = [ps[0]];
    }
    if (!ps.length) {
      // the editor lost its paragraph structure (e.g. after select-all + delete): wrap it
      const p = document.createElement("p");
      while (ed.el.firstChild) p.appendChild(ed.el.firstChild);
      if (!p.firstChild) p.appendChild(document.createElement("br"));
      ed.el.appendChild(p);
      hit = [p];
    }
    for (const p of hit) fn(p);
    restyleParas(ed.el);
    this.drawOverlay();
    this.updateToolbar();
  }

  // ------------------------------------------------------------------ formatting

  /** Items whose text the toolbar acts on (when not editing). */
  textTargets() {
    const s = this.slide();
    return this.sel.map((i) => s.items[i]).filter((it) => it && (it.type === "text" || it.type === "table" || (it.type === "shape" && it.shape !== "line")));
  }

  allParas(it) {
    if (it.type === "table") return (it.rows || []).flatMap((r) => r.flatMap((c) => (c && c.paras) || []));
    if (it.type === "shape" && !Array.isArray(it.paras)) it.paras = [];
    return it.paras || [];
  }

  allRuns(it) {
    const out = [];
    for (const p of this.allParas(it)) {
      if (!p.runs) p.runs = [];
      if (!p.runs.length) p.runs.push({ t: "" });
      out.push(...p.runs);
    }
    return out;
  }

  modelFormat(fn, label) {
    const targets = this.textTargets();
    if (!targets.length) return false;
    this.pushUndo();
    for (const it of targets) fn(it);
    this.changed();
    if (label) this.status(label);
    return true;
  }

  toggleRunFlag(flag, execName) {
    if (this.editing) {
      this.restoreRange();
      document.execCommand(execName, false, null);
      this.saveRange();
      this.updateToolbar();
      return;
    }
    const runs = this.textTargets().flatMap((it) => this.allRuns(it));
    if (!runs.length) return;
    const all = runs.every((r) => r[flag] || !r.t);
    this.modelFormat((it) => { for (const r of this.allRuns(it)) { if (all) delete r[flag]; else r[flag] = true; } });
  }

  /** While editing: make sure something is selected (whole paragraph when the caret is collapsed). */
  ensureRangeForFormat() {
    this.restoreRange();
    const sel = window.getSelection();
    if (!sel.rangeCount) return;
    const r = sel.getRangeAt(0);
    if (!r.collapsed) return;
    let p = r.startContainer.nodeType === 1 ? r.startContainer : r.startContainer.parentElement;
    p = p && p.closest("p");
    if (!p || !this.editing.el.contains(p)) return;
    const nr = document.createRange();
    nr.selectNodeContents(p);
    sel.removeAllRanges();
    sel.addRange(nr);
  }

  applyColor(color) {
    if (this.editing) {
      this.refocus();
      this.ensureRangeForFormat();
      document.execCommand("foreColor", false, color || this.textDefaults(this.slide().items[this.editing.index]).color);
      this.saveRange();
      this.updateToolbar();
      return;
    }
    this.modelFormat((it) => { for (const r of this.allRuns(it)) { if (color) r.color = color; else delete r.color; } }, color ? "Text color changed" : "Text color reset");
  }

  applyFont(font) {
    if (this.editing) {
      this.refocus();
      this.ensureRangeForFormat();
      const theme = this.deck.theme || {};
      document.execCommand("fontName", false, font || theme.font || "Calibri");
      // tag the new <font face> so the parser keeps or drops it correctly
      for (const f of this.editing.el.querySelectorAll("font[face]")) {
        const sp = document.createElement("span");
        if (font) { sp.dataset.font = font; sp.style.fontFamily = `"${font}"`; }
        while (f.firstChild) sp.appendChild(f.firstChild);
        f.replaceWith(sp);
        if (!font) for (const n of sp.querySelectorAll("[data-font]")) { delete n.dataset.font; n.style.fontFamily = ""; }
      }
      this.refocus();
      this.saveRange();
      this.updateToolbar();
      return;
    }
    this.modelFormat((it) => { for (const r of this.allRuns(it)) { if (font) r.font = font; else delete r.font; } }, font ? `Font: ${font}` : "Theme font");
    this.refocus();
  }

  applySize(size) {
    size = Math.round(size * 10) / 10;
    if (this.editing) {
      this.refocus();
      this.ensureRangeForFormat();
      document.execCommand("fontSize", false, "7");
      for (const f of this.editing.el.querySelectorAll('font[size="7"]')) {
        const sp = document.createElement("span");
        sp.style.fontSize = `${size}px`;
        sp.dataset.sz = String(size);
        while (f.firstChild) sp.appendChild(f.firstChild);
        f.replaceWith(sp);
        for (const n of sp.querySelectorAll("[data-sz]")) { n.dataset.sz = String(size); n.style.fontSize = `${size}px`; }
      }
      // paragraphs that are now entirely this size take it as their default
      for (const p of this.editing.el.querySelectorAll("p")) {
        const spans = [...p.querySelectorAll("[data-sz]")];
        if (spans.length && spans.every((n) => n.dataset.sz === String(size))) { p.dataset.psz = String(size); p.style.fontSize = `${size}px`; }
      }
      this.refocus();
      this.saveRange();
      this.updateToolbar();
      return;
    }
    this.modelFormat((it) => { for (const r of this.allRuns(it)) r.size = size; }, `Font size ${size} pt`);
    this.refocus();
  }

  applyAlign(align) {
    if (this.editing) {
      this.editParas((p) => {
        if (align === "left") { p.style.textAlign = ""; delete p.dataset.align; } else { p.style.textAlign = align; p.dataset.align = align; }
      });
      return;
    }
    this.modelFormat((it) => { for (const p of this.allParas(it)) { if (align === "left") delete p.align; else p.align = align; } });
  }

  applyBullet(kind) {
    if (this.editing) {
      const ps = [];
      this.editParas((p) => ps.push(p));
      const all = ps.every((p) => p.dataset.bullet === kind);
      for (const p of ps) { if (all) delete p.dataset.bullet; else p.dataset.bullet = kind; }
      restyleParas(this.editing.el);
      this.updateToolbar();
      return;
    }
    const paras = this.textTargets().flatMap((it) => this.allParas(it));
    const all = paras.length && paras.every((p) => p.bullet === kind);
    this.modelFormat((it) => { for (const p of this.allParas(it)) { if (all) delete p.bullet; else p.bullet = kind; } });
  }

  applyLevel(delta) {
    if (this.editing) {
      this.editParas((p) => {
        const l = Math.max(0, Math.min(8, (Number(p.dataset.level) || 0) + delta));
        if (l) p.dataset.level = String(l); else delete p.dataset.level;
        // smaller text for deeper body levels, like PowerPoint's master
      });
      return;
    }
    this.modelFormat((it) => {
      for (const p of this.allParas(it)) {
        const l = Math.max(0, Math.min(8, (p.level || 0) + delta));
        if (l) p.level = l; else delete p.level;
      }
    });
  }

  applyValign(v) {
    const s = this.slide();
    const items = this.sel.map((i) => s.items[i]).filter((it) => it && (it.type === "text" || it.type === "shape"));
    if (!items.length) return;
    if (this.editing) this.commitEdit();
    this.pushUndo();
    for (const it of items) it.valign = v;
    this.changed();
  }

  applyFill(color) {
    const s = this.slide();
    if (this.editing) this.commitEdit();
    const items = this.sel.map((i) => s.items[i]).filter((it) => it && (it.type === "text" || it.type === "shape" || it.type === "table"));
    if (!items.length) return;
    this.pushUndo();
    for (const it of items) {
      if (it.type === "table") {
        for (const r of it.rows || []) for (const c of r) { if (color) c.fill = color; else delete c.fill; }
      } else if (color) it.fill = color; else delete it.fill;
    }
    this.changed();
    this.status(color ? "Fill changed" : "No fill");
  }

  applyLine(color, width) {
    const s = this.slide();
    if (this.editing) this.commitEdit();
    const items = this.sel.map((i) => s.items[i]).filter((it) => it && (it.type === "text" || it.type === "shape"));
    if (!items.length) return;
    this.pushUndo();
    for (const it of items) {
      if (color === null) { delete it.line; delete it.lineW; continue; }
      if (color) it.line = color;
      if (width) { it.lineW = width; if (!it.line) it.line = this.deck.theme.color || "#262626"; }
    }
    this.changed();
  }

  // ------------------------------------------------------------------ objects

  centerBox(w, h) {
    const d = this.deck.size;
    return { x: r2((d.w - w) / 2), y: r2((d.h - h) / 2), w: r2(w), h: r2(h) };
  }

  addItem(it, { edit = false } = {}) {
    if (this.editing) this.commitEdit();
    this.pushUndo();
    const s = this.slide();
    s.items.push(it);
    const i = s.items.length - 1;
    this.sel = [i];
    this.changed();
    if (edit) this.beginEdit(i, { selectAll: true });
    return i;
  }

  insertTextBox() {
    const d = this.deck.size;
    const it = { type: "text", ...this.centerBox(d.w * 0.42, 50), valign: "top", paras: [{ runs: [{ t: "", size: 20 }] }] };
    this.addItem(it, { edit: true });
    this.status("Type your text. Press Esc when done.");
  }

  insertShape(shape) {
    const theme = this.deck.theme || {};
    const accent = theme.accent || "#4472C4";
    const isLine = shape === "line";
    const w = isLine ? 220 : shape === "rightArrow" || shape === "leftArrow" || shape === "chevron" ? 200 : 160;
    const h = isLine ? 0 : shape === "rightArrow" || shape === "leftArrow" ? 90 : shape === "rect" || shape === "roundRect" || shape === "ellipse" || shape === "hexagon" ? 110 : 140;
    const it = { type: "shape", ...this.centerBox(w, h), shape };
    if (isLine) { it.line = theme.color || "#262626"; it.lineW = 2; }
    else { it.fill = accent; it.line = tintColor(accent, -0.3); it.lineW = 1; }
    this.addItem(it);
    this.status(`${(SHAPES.find((x) => x.id === shape) || {}).name || "Shape"} inserted. Drag to move; double-click to add text.`);
  }

  insertTable(rows, cols) {
    const d = this.deck.size;
    const accent = (this.deck.theme && this.deck.theme.accent) || "#4472C4";
    const w = Math.min(d.w * 0.86, cols * 140);
    const h = rows * 34;
    const tableRows = [];
    for (let r = 0; r < rows; r++) {
      const row = [];
      for (let c = 0; c < cols; c++) {
        const cell = { paras: [{ runs: r === 0 ? [{ t: "", b: true }] : [] }] };
        if (r === 0) cell.fill = accent;
        else if (r % 2 === 1) cell.fill = tintColor(accent, 0.82);
        else cell.fill = tintColor(accent, 0.92);
        row.push(cell);
      }
      tableRows.push(row);
    }
    const it = { type: "table", ...this.centerBox(w, h), rows: tableRows, colW: Array.from({ length: cols }, () => r2(w / cols)) };
    const i = this.addItem(it);
    this.beginEdit(i, { cell: [0, 0] });
    this.status(`Table ${rows} × ${cols} inserted. Tab moves to the next cell.`);
  }

  async insertImages(files, at = null) {
    const d = this.deck.size;
    let n = 0;
    for (const f of files.slice(0, 20)) {
      try {
        const img = await readImageFile(f);
        const k = Math.min(1, (d.w * 0.6) / img.w, (d.h * 0.6) / img.h);
        const w = img.w * k, h = img.h * k;
        const box = at ? { x: r2(at.x - w / 2 + n * 12), y: r2(at.y - h / 2 + n * 12), w: r2(w), h: r2(h) } : this.centerBox(w, h);
        if (!at && n) { box.x += n * 12; box.y += n * 12; }
        const it = { type: "image", ...box, src: img.src };
        if (img.name) it.alt = img.name.replace(/\.[^.]+$/, "");
        this.addItem(it);
        n++;
      } catch (e) {
        this.status(`Could not insert ${f.name || "the picture"}: ${e.message || e}`);
      }
    }
    if (n) this.status(n === 1 ? "Picture inserted" : `${n} pictures inserted`);
  }

  deleteItems() {
    const s = this.slide();
    if (!this.sel.length) return;
    this.pushUndo();
    const del = new Set(this.sel);
    s.items = s.items.filter((_, i) => !del.has(i));
    this.sel = [];
    this.changed();
    this.status(`${del.size} object${del.size > 1 ? "s" : ""} deleted`);
  }

  duplicateItems() {
    const s = this.slide();
    if (!this.sel.length) return;
    if (this.editing) this.commitEdit();
    this.pushUndo();
    const copies = this.sel.sort((a, b) => a - b).map((i) => {
      const c = cloneDeep(s.items[i]);
      if (typeof c.x === "number") { c.x = r2(c.x + 12); c.y = r2(c.y + 12); }
      return c;
    });
    const start = s.items.length;
    s.items.push(...copies);
    this.sel = copies.map((_, k) => start + k);
    this.changed();
    this.status("Duplicated");
  }

  copyItems() {
    const s = this.slide();
    if (!this.sel.length) return false;
    const items = this.sel.sort((a, b) => a - b).map((i) => cloneDeep(s.items[i]));
    this.clip = { items, fromSlide: s, text: items.map(itemText).filter(Boolean).join("\n") };
    this.pasteN = 0;
    this.status(`Copied ${items.length} object${items.length > 1 ? "s" : ""}`);
    return true;
  }

  pasteInternal(clip = this.clip) {
    if (!clip) return;
    if (this.editing) this.commitEdit();
    const deck = this.deck;
    if (clip.slides) {
      this.pushUndo();
      const copies = clip.slides.map(cloneDeep);
      deck.slides.splice(this.current + 1, 0, ...copies);
      this.current += 1;
      this.sel = [];
      this.changed({ strip: true });
      this.status(`Pasted ${copies.length} slide${copies.length > 1 ? "s" : ""}`);
      return;
    }
    const s = this.slide();
    this.pushUndo();
    const same = clip === this.clip && clip.fromSlide === s;
    if (clip === this.clip) this.pasteN = same ? this.pasteN + 1 : 0;
    const off = same ? 12 * this.pasteN : 0;
    const start = s.items.length;
    for (const it of clip.items) {
      const c = cloneDeep(it);
      if (typeof c.x === "number") { c.x = r2(c.x + off); c.y = r2(c.y + off); }
      s.items.push(c);
    }
    this.sel = clip.items.map((_, k) => start + k);
    this.changed();
    this.status(`Pasted ${clip.items.length} object${clip.items.length > 1 ? "s" : ""}`);
  }

  onCopyEvent(e, cut) {
    if (this.destroyed || this.show || this.editing || isTypingEl(e.target)) return;
    if (!this.root.contains(e.target) && e.target !== document.body) return;
    if (this.thumbs.contains(e.target) && !this.sel.length) {
      this.copySlide();
      e.clipboardData.setData("text/plain", slideTitle(this.slide()) || "Slide");
      e.clipboardData.setData(CLIP_MIME, JSON.stringify({ slides: this.clip.slides }));
      e.preventDefault();
      if (cut) this.deleteSlide();
      return;
    }
    if (!this.sel.length) return;
    if (!this.clipFresh) this.copyItems();
    this.clipFresh = false;
    e.clipboardData.setData("text/plain", this.clip.text || "");
    e.clipboardData.setData(CLIP_MIME, JSON.stringify({ items: this.clip.items }));
    e.preventDefault();
    if (cut && this.sel.length) this.deleteItems();
  }

  onPaste(e) {
    if (this.destroyed || this.show) return;
    if (this.pendingPaste) { clearTimeout(this.pendingPaste); this.timers.delete(this.pendingPaste); this.pendingPaste = null; }
    const cd = e.clipboardData;
    if (this.editing && this.editing.el.contains(e.target)) {
      // keep pasted text plain inside a text box
      const t = cd && cd.getData("text/plain");
      e.preventDefault();
      e.stopPropagation();
      if (t) document.execCommand("insertText", false, t.replace(/\r\n?/g, "\n"));
      return;
    }
    if (isTypingEl(e.target)) return;
    if (!this.root.contains(e.target) && e.target !== document.body) return;
    const custom = cd && cd.getData(CLIP_MIME);
    if (custom) {
      try {
        const data = JSON.parse(custom);
        e.preventDefault(); e.stopPropagation();
        if (data.slides) this.pasteInternal({ slides: data.slides });
        else if (Array.isArray(data.items)) {
          // the same content we copied: use the internal clip (keeps the paste offset)
          if (this.clip && this.clip.items && JSON.stringify(this.clip.items) === JSON.stringify(data.items)) this.pasteInternal();
          else this.pasteInternal({ items: data.items });
        }
        return;
      } catch { /* fall through */ }
    }
    const files = [...((cd && cd.files) || [])].filter((f) => /^image\//.test(f.type));
    if (files.length) {
      e.preventDefault(); e.stopPropagation();
      this.insertImages(files);
      return;
    }
    const text = cd && cd.getData("text/plain");
    if (text && text.trim()) {
      e.preventDefault(); e.stopPropagation();
      if (this.clip && this.clip.text === text) { this.pasteInternal(); return; }
      const lines = text.replace(/\r\n?/g, "\n").split("\n").slice(0, 200);
      const d = this.deck.size;
      const h = Math.min(d.h * 0.8, 30 * lines.length + 20);
      this.addItem({ type: "text", ...this.centerBox(d.w * 0.6, h), valign: "top", paras: lines.map((l) => ({ runs: [{ t: l, size: 20 }] })) });
      this.status("Text pasted as a new text box");
      return;
    }
    if (this.clip) { e.preventDefault(); e.stopPropagation(); this.pasteInternal(); }
  }

  arrange(op) {
    const s = this.slide();
    if (!this.sel.length) return;
    if (this.editing) this.commitEdit();
    this.pushUndo();
    const selSet = new Set(this.sel);
    const items = s.items;
    if (op === "front" || op === "back") {
      const picked = items.filter((_, i) => selSet.has(i));
      const rest = items.filter((_, i) => !selSet.has(i));
      s.items = op === "front" ? [...rest, ...picked] : [...picked, ...rest];
      this.sel = op === "front" ? picked.map((_, k) => rest.length + k) : picked.map((_, k) => k);
    } else if (op === "forward") {
      for (let i = items.length - 2; i >= 0; i--) {
        if (selSet.has(i) && !selSet.has(i + 1)) { [items[i], items[i + 1]] = [items[i + 1], items[i]]; selSet.delete(i); selSet.add(i + 1); }
      }
      this.sel = [...selSet];
    } else if (op === "backward") {
      for (let i = 1; i < items.length; i++) {
        if (selSet.has(i) && !selSet.has(i - 1)) { [items[i], items[i - 1]] = [items[i - 1], items[i]]; selSet.delete(i); selSet.add(i - 1); }
      }
      this.sel = [...selSet];
    } else {
      const d = this.deck.size;
      const ref = this.sel.length > 1 ? this.bbox(this.sel) : { x: 0, y: 0, w: d.w, h: d.h };
      for (const i of this.sel) {
        const it = items[i];
        if (op === "left") it.x = ref.x;
        else if (op === "hcenter") it.x = r2(ref.x + (ref.w - it.w) / 2);
        else if (op === "right") it.x = r2(ref.x + ref.w - it.w);
        else if (op === "top") it.y = ref.y;
        else if (op === "vmiddle") it.y = r2(ref.y + (ref.h - it.h) / 2);
        else if (op === "bottom") it.y = r2(ref.y + ref.h - it.h);
      }
    }
    this.changed();
  }

  nudge(dx, dy) {
    const s = this.slide();
    this.pushUndo(`nudge:${this.current}:${this.sel.join(",")}`);
    for (const i of this.sel) { const it = s.items[i]; if (it && typeof it.x === "number") { it.x = r2(it.x + dx); it.y = r2(it.y + dy); this.moveItemDom(i); } }
    this.drawOverlay();
    this.refreshThumbSoon(this.current);
    this.dirty();
  }

  /** Re-apply a layout to the current slide, keeping the text of its placeholders. */
  applyLayout(layout) {
    if (this.editing) this.commitEdit();
    const deck = this.deck;
    const s = this.slide();
    this.pushUndo();
    const fresh = newSlide(layout, deck);
    const oldPh = s.items.filter((it) => it && it.type === "text" && it.ph);
    const others = s.items.filter((it) => !(it && it.type === "text" && it.ph));
    const take = (kind) => {
      const k = oldPh.findIndex((it) => (kind === "title" ? it.ph === "title" : it.ph !== "title"));
      return k >= 0 ? oldPh.splice(k, 1)[0] : null;
    };
    for (const ph of fresh.items) {
      const old = take(ph.ph === "title" ? "title" : "body");
      if (old && old.paras.some((p) => paraText(p).trim())) {
        ph.paras = old.paras.map((p) => {
          const np = cloneDeep(p);
          if (ph.ph === "subtitle" || layout === "section") delete np.bullet;
          else if (ph.ph === "body" && np.bullet === undefined && (old.ph === "subtitle")) np.bullet = "ul";
          return np;
        });
        if (ph.ph === "title" && layout === "title") for (const p of ph.paras) p.align = "center";
      }
    }
    // leftover placeholder text becomes plain text boxes so nothing is lost
    const leftovers = oldPh.filter((it) => it.paras.some((p) => paraText(p).trim())).map((it) => { const c = cloneDeep(it); delete c.ph; return c; });
    s.items = [...fresh.items, ...others, ...leftovers];
    s.layout = fresh.layout;
    this.sel = [];
    this.changed();
    this.status(`Layout: ${(LAYOUTS.find((l) => l.id === layout) || {}).name || layout}`);
  }

  applyTheme(t) {
    if (this.editing) this.commitEdit();
    const deck = this.deck;
    this.pushUndo();
    const old = deck.theme || {};
    deck.theme = { ...t.theme };
    // shapes/tables that used the old accent follow the new one
    const oldAccent = hexColor(old.accent);
    if (oldAccent) {
      const map = new Map([[oldAccent, t.theme.accent], [tintColor(oldAccent, -0.3), tintColor(t.theme.accent, -0.3)], [tintColor(oldAccent, 0.82), tintColor(t.theme.accent, 0.82)], [tintColor(oldAccent, 0.92), tintColor(t.theme.accent, 0.92)]]);
      for (const s of deck.slides) for (const it of s.items) {
        if (!it) continue;
        if (it.fill && map.has(it.fill)) it.fill = map.get(it.fill);
        if (it.line && map.has(it.line)) it.line = map.get(it.line);
        if (it.type === "table") for (const r of it.rows || []) for (const c of r) if (c.fill && map.has(c.fill)) c.fill = map.get(c.fill);
      }
    }
    this.changed({ all: true });
    this.status(`Theme: ${t.name}`);
  }

  // ------------------------------------------------------------------ commands

  cmd(name, arg) {
    const deck = this.deck;
    if (!deck) return;
    switch (name) {
      case "undo": this.undo(); break;
      case "redo": this.redo(); break;
      case "newSlide": this.addSlide(arg); break;
      case "dupSlide": this.duplicateSlide(); break;
      case "delSlide": this.deleteSlide(); break;
      case "hideSlide": {
        const s = this.slide();
        this.pushUndo();
        if (s.hidden) delete s.hidden; else s.hidden = true;
        this.changed({ strip: true });
        this.status(s.hidden ? "Slide hidden: it is skipped in the slide show" : "Slide shown in the slide show again");
        break;
      }
      case "textbox": this.insertTextBox(); break;
      case "image": this.fileInput.click(); break;
      case "bold": this.toggleRunFlag("b", "bold"); break;
      case "italic": this.toggleRunFlag("i", "italic"); break;
      case "underline": this.toggleRunFlag("u", "underline"); break;
      case "strike": this.toggleRunFlag("s", "strikeThrough"); break;
      case "alignLeft": this.applyAlign("left"); break;
      case "alignCenter": this.applyAlign("center"); break;
      case "alignRight": this.applyAlign("right"); break;
      case "alignJustify": this.applyAlign("justify"); break;
      case "bullets": this.applyBullet("ul"); break;
      case "numbers": this.applyBullet("ol"); break;
      case "indent": this.applyLevel(1); break;
      case "outdent": this.applyLevel(-1); break;
      case "present": this.startShow(0); break;
      case "presentCurrent": this.startShow(this.current); break;
      default: break;
    }
    if (!this.editing && !/^(bold|italic|underline|strike|align|bullets|numbers|indent|outdent)/.test(name)) this.focus();
  }

  updateToolbarSoon() {
    if (this.tbTimer) return;
    this.tbTimer = this.later(() => { this.tbTimer = null; this.updateToolbar(); }, 60);
  }

  updateToolbar() {
    if (!this.ribbon) return;
    const deck = this.deck;
    const s = this.slide();
    const items = this.sel.map((i) => s.items[i]).filter(Boolean);
    const textOk = !!this.editing || this.textTargets().length > 0;
    const shapeOk = items.some((it) => it.type === "text" || it.type === "shape" || it.type === "table");
    const lineOk = items.some((it) => it.type === "text" || it.type === "shape");
    const set = (sel, on) => { for (const n of this.ribbon.querySelectorAll(sel)) n.disabled = !on; };
    set(".sv-g-text button, .sv-g-text select, .sv-g-text input, .sv-g-para button", textOk);
    set('[data-pop="fill"]', shapeOk);
    set('[data-pop="outline"]', lineOk);
    set('[data-pop="arrange"]', items.length > 0);
    set('[data-cmd="delSlide"]', deck.slides.length > 1);
    this.updateUndoButtons();
    // states
    let st = {};
    if (this.editing) {
      const q = (c) => { try { return document.queryCommandState(c); } catch { return false; } };
      st = { b: q("bold"), i: q("italic"), u: q("underline"), s: q("strikeThrough") };
      const sel = window.getSelection();
      const node = sel && sel.rangeCount ? sel.getRangeAt(0).startContainer : null;
      const elx = node ? (node.nodeType === 1 ? node : node.parentElement) : null;
      if (elx && this.editing.el.contains(elx)) {
        const cs = getComputedStyle(elx);
        st.size = Math.round(parseFloat(cs.fontSize) * 10) / 10;
        const fam = elx.closest("[data-font]");
        st.font = fam ? fam.dataset.font : "";
        st.color = rgbToHex(cs.color);
        const p = elx.closest("p");
        if (p) { st.align = p.dataset.align || "left"; st.bullet = p.dataset.bullet || ""; }
      }
    } else if (textOk) {
      const it = this.textTargets()[0];
      const runs = this.allRuns(it).filter((r) => r.t) ;
      const r0 = runs[0] || this.allRuns(it)[0] || {};
      const p0 = this.allParas(it)[0] || {};
      st = { b: runs.length > 0 && runs.every((r) => r.b), i: runs.length > 0 && runs.every((r) => r.i), u: runs.length > 0 && runs.every((r) => r.u), s: runs.length > 0 && runs.every((r) => r.s), size: r0.size || defaultRunSize(it, p0.level || 0), font: r0.font || "", color: r0.color || textColorFor(it, deck), align: p0.align || "left", bullet: p0.bullet || "" };
    }
    const press = (cmd, on) => { const n = this.ribbon.querySelector(`[data-cmd="${cmd}"]`); if (n) n.setAttribute("aria-pressed", String(!!on)); };
    press("bold", st.b); press("italic", st.i); press("underline", st.u); press("strike", st.s);
    press("alignLeft", textOk && st.align === "left"); press("alignCenter", st.align === "center"); press("alignRight", st.align === "right"); press("alignJustify", st.align === "justify");
    press("bullets", st.bullet === "ul"); press("numbers", st.bullet === "ol");
    const size = this.q(".sv-size");
    if (size && document.activeElement !== size) size.value = textOk && st.size ? String(st.size) : "";
    const font = this.q(".sv-font");
    if (font && document.activeElement !== font) {
      const f = st.font || "";
      if (f && ![...font.options].some((o) => o.value === f)) font.insertAdjacentHTML("beforeend", `<option value="${esc(f)}">${esc(f)}</option>`);
      font.value = f;
      font.options[0].textContent = `${(items[0] && items[0].ph === "title" ? deck.theme.headFont : deck.theme.font) || deck.theme.font || "Calibri"} (theme)`;
    }
    const bar = (k, c) => { const n = this.ribbon.querySelector(`[data-bar="${k}"]`); if (n) n.style.background = c || "transparent"; };
    bar("text", st.color || (deck.theme && deck.theme.color));
    const it0 = items[0];
    bar("fill", it0 && it0.fill ? it0.fill : "");
    bar("line", it0 && it0.line ? it0.line : "");
    const hideBtn = this.ribbon.querySelector('[data-cmd="hideSlide"]');
    if (hideBtn) hideBtn.setAttribute("aria-pressed", String(!!(s && s.hidden)));
  }

  // ------------------------------------------------------------------ keyboard

  onKey(e) {
    if (this.destroyed || !this.root || this.root.hidden) return;
    if (this.show) { this.onShowKey(e); return; }
    if (document.querySelector(".modal:not([hidden])")) return;
    const mod = e.ctrlKey || e.metaKey;
    const k = e.key;
    if (this.pop && k === "Escape") { e.preventDefault(); this.closePop(); this.refocus(); return; }
    // slide show keys work from anywhere in the editor
    if (k === "F5" && !mod) {
      if (!this.root.contains(e.target) && e.target !== document.body) return;
      e.preventDefault();
      this.startShow(e.shiftKey ? this.current : 0);
      return;
    }
    const inRoot = this.root.contains(e.target);
    if (!inRoot && e.target !== document.body && e.target !== document.documentElement) return;
    // ----- text editing
    if (this.editing && this.editing.el.contains(e.target)) {
      if (k === "Escape") { e.preventDefault(); this.commitEdit(); this.focus(); return; }
      if (k === "Tab" && !mod) {
        e.preventDefault();
        if (this.editing.cell) { this.moveCell(e.shiftKey ? -1 : 1); return; }
        this.applyLevel(e.shiftKey ? -1 : 1);
        return;
      }
      if (mod && !e.altKey) {
        const kl = k.toLowerCase();
        if (kl === "b" || kl === "i" || kl === "u") { e.preventDefault(); this.cmd({ b: "bold", i: "italic", u: "underline" }[kl]); return; }
        if (kl === "e" || kl === "l" || kl === "r" || kl === "j") { e.preventDefault(); this.cmd({ e: "alignCenter", l: "alignLeft", r: "alignRight", j: "alignJustify" }[kl]); return; }
        if (kl === "]" || kl === "[") { e.preventDefault(); const cur = this.q(".sv-size"); const v = parseFloat(cur.value) || 18; const idx = SIZES_PT.findIndex((x) => x >= v); const nv = kl === "]" ? SIZES_PT[Math.min(SIZES_PT.length - 1, (SIZES_PT[idx] === v ? idx + 1 : idx))] : SIZES_PT[Math.max(0, idx - 1)]; this.applySize(nv || v); return; }
      }
      return; // let the browser handle typing, selection, native undo
    }
    if (isTypingEl(e.target)) return;
    if (this.thumbs && this.thumbs.contains(e.target)) {
      if (mod && (k === "d" || k === "D")) { e.preventDefault(); this.duplicateSlide(); return; }
      if (mod && !e.shiftKey && (k === "z" || k === "Z")) { e.preventDefault(); this.undo(); return; }
      if (mod && (k === "y" || k === "Y" || (e.shiftKey && (k === "z" || k === "Z")))) { e.preventDefault(); this.redo(); return; }
      if (mod && (k === "c" || k === "C")) { this.copySlide(); return; }
      if (mod && (k === "v" || k === "V")) { this.schedulePaste(); return; }
      if (mod && (k === "m" || k === "M")) { e.preventDefault(); this.addSlide(); return; }
      return;
    }
    const s = this.slide();
    if (mod && !e.altKey) {
      const kl = k.toLowerCase();
      if (kl === "z" && !e.shiftKey) { e.preventDefault(); this.undo(); return; }
      if (kl === "y" || (kl === "z" && e.shiftKey)) { e.preventDefault(); this.redo(); return; }
      if (kl === "d") { e.preventDefault(); if (this.sel.length) this.duplicateItems(); else this.duplicateSlide(); return; }
      if (kl === "m") { e.preventDefault(); this.addSlide(); return; }
      if (kl === "a") { e.preventDefault(); this.selectItems(s.items.map((_, i) => i)); return; }
      if (kl === "c" || kl === "x") {
        if (this.sel.length) {
          this.copyItems();
          this.clipFresh = true;
          if (kl === "x") this.later(() => { if (this.sel.length) this.deleteItems(); }, 0);
        } else if (kl === "c") this.copySlide();
        return; // the copy/cut event also fills the system clipboard
      }
      if (kl === "v") { this.schedulePaste(); return; }
      if (kl === "b" || kl === "i" || kl === "u") { if (this.textTargets().length) { e.preventDefault(); this.cmd({ b: "bold", i: "italic", u: "underline" }[kl]); } return; }
      if (kl === "e" || kl === "l" || kl === "r" || kl === "j") { if (this.textTargets().length) { e.preventDefault(); this.cmd({ e: "alignCenter", l: "alignLeft", r: "alignRight", j: "alignJustify" }[kl]); } return; }
      if (k === "ArrowUp" || k === "ArrowDown") return;
      return;
    }
    if (k === "Escape") { if (this.sel.length) { e.preventDefault(); this.selectItems([]); } return; }
    if (k === "Delete" || k === "Backspace") {
      e.preventDefault();
      if (this.sel.length) this.deleteItems();
      return;
    }
    if (k.startsWith("Arrow")) {
      e.preventDefault();
      if (this.sel.length) {
        const st = e.shiftKey ? 10 : 1;
        const d = { ArrowLeft: [-st, 0], ArrowRight: [st, 0], ArrowUp: [0, -st], ArrowDown: [0, st] }[k];
        this.nudge(d[0], d[1]);
      } else if (k === "ArrowDown" || k === "ArrowRight") this.selectSlide(this.current + 1);
      else this.selectSlide(this.current - 1);
      return;
    }
    if (k === "PageDown") { e.preventDefault(); this.selectSlide(this.current + 1); return; }
    if (k === "PageUp") { e.preventDefault(); this.selectSlide(this.current - 1); return; }
    if (k === "Home" && !this.sel.length) { e.preventDefault(); this.selectSlide(0); return; }
    if (k === "End" && !this.sel.length) { e.preventDefault(); this.selectSlide(this.deck.slides.length - 1); return; }
    if (k === "Tab") {
      e.preventDefault();
      const n = s.items.length;
      if (!n) return;
      const cur = this.sel.length ? this.sel[this.sel.length - 1] : -1;
      this.selectItems([((cur + (e.shiftKey ? -1 : 1)) % n + n) % n]);
      return;
    }
    if (this.sel.length === 1) {
      const it = s.items[this.sel[0]];
      const editable = it && (it.type === "text" || (it.type === "shape" && it.shape !== "line") || it.type === "table");
      if (!editable) return;
      if (k === "Enter" || k === "F2") {
        e.preventDefault();
        this.beginEdit(this.sel[0], { selectAll: k === "Enter" && false });
        return;
      }
      if (k.length === 1 && !e.altKey) {
        // typing replaces the text of the selected box
        e.preventDefault();
        if (this.beginEdit(this.sel[0], { selectAll: true })) document.execCommand("insertText", false, k);
      }
    }
  }

  schedulePaste() {
    // the paste event normally follows; without it (no clipboard access) use our own clipboard
    if (this.pendingPaste) clearTimeout(this.pendingPaste);
    this.pendingPaste = this.later(() => { this.pendingPaste = null; this.pasteInternal(); }, 80);
  }

  moveCell(dir) {
    const ed = this.editing;
    const it = this.slide().items[ed.index];
    const rows = it.rows.length, cols = it.rows.reduce((m, r) => Math.max(m, r.length), 0);
    let n = ed.cell[0] * cols + ed.cell[1] + dir;
    const i = ed.index;
    if (n >= rows * cols) {
      // Tab in the last cell adds a row, like PowerPoint
      this.commitEdit();
      this.pushUndo();
      it.rows.push(Array.from({ length: cols }, () => ({ paras: [{ runs: [] }] })));
      it.h = r2(it.h + it.h / rows);
      this.changed();
      this.selectItems([i]);
    } else this.commitEdit();
    n = Math.max(0, n);
    this.selectItems([i]);
    this.beginEdit(i, { cell: [Math.floor(n / cols), n % cols], selectAll: true });
  }

  tableMenuItems() {
    const ed = this.editing;
    const i = ed ? ed.index : this.sel[0];
    const it = this.slide().items[i];
    if (!it || it.type !== "table") return [];
    const cell = ed && ed.cell ? ed.cell : [it.rows.length - 1, (it.rows[0] || []).length - 1];
    const mutate = (fn, msg) => () => {
      if (this.editing) this.commitEdit();
      this.pushUndo();
      fn();
      this.sel = [i];
      this.changed();
      this.status(msg);
    };
    const cols = it.rows.reduce((m, r) => Math.max(m, r.length), 0);
    return [
      { label: "Insert row above", run: mutate(() => { it.rows.splice(cell[0], 0, Array.from({ length: cols }, () => ({ paras: [{ runs: [] }] }))); it.h = r2(it.h * (it.rows.length / (it.rows.length - 1))); }, "Row inserted") },
      { label: "Insert row below", run: mutate(() => { it.rows.splice(cell[0] + 1, 0, Array.from({ length: cols }, () => ({ paras: [{ runs: [] }] }))); it.h = r2(it.h * (it.rows.length / (it.rows.length - 1))); }, "Row inserted") },
      { label: "Insert column left", run: mutate(() => { for (const r of it.rows) r.splice(cell[1], 0, { paras: [{ runs: [] }] }); if (it.colW) it.colW.splice(cell[1], 0, it.colW[cell[1]] || 100); this.fitCols(it); }, "Column inserted") },
      { label: "Insert column right", run: mutate(() => { for (const r of it.rows) r.splice(cell[1] + 1, 0, { paras: [{ runs: [] }] }); if (it.colW) it.colW.splice(cell[1] + 1, 0, it.colW[cell[1]] || 100); this.fitCols(it); }, "Column inserted") },
      { label: "Delete row", disabled: it.rows.length <= 1, run: mutate(() => { const n = it.rows.length; it.rows.splice(cell[0], 1); it.h = r2(it.h * ((n - 1) / n)); }, "Row deleted") },
      { label: "Delete column", disabled: cols <= 1, run: mutate(() => { for (const r of it.rows) r.splice(cell[1], 1); if (it.colW) it.colW.splice(cell[1], 1); this.fitCols(it); }, "Column deleted") },
    ];
  }

  fitCols(it) {
    const cols = it.rows.reduce((m, r) => Math.max(m, r.length), 0);
    if (!Array.isArray(it.colW) || it.colW.length !== cols) it.colW = Array.from({ length: cols }, () => it.w / cols);
    const sum = it.colW.reduce((a, b) => a + b, 0) || 1;
    it.colW = it.colW.map((c) => r2((c / sum) * it.w));
  }

  onNotesGrip(e) {
    e.preventDefault();
    const wrap = this.q(".sv-notes-wrap");
    const startY = e.clientY;
    const h0 = wrap.getBoundingClientRect().height;
    const move = (ev) => {
      const h = Math.max(36, Math.min(window.innerHeight * 0.6, h0 - (ev.clientY - startY)));
      wrap.style.height = `${h}px`;
    };
    const up = () => {
      window.removeEventListener("pointermove", move);
      window.removeEventListener("pointerup", up);
      if (this.zoomMode === "fit") this.fit();
    };
    window.addEventListener("pointermove", move);
    window.addEventListener("pointerup", up);
  }

  // ------------------------------------------------------------------ popovers / menus

  closePop() {
    if (!this.pop) return;
    this.pop = null;
    if (this.popEl) { this.popEl.hidden = true; this.popEl.innerHTML = ""; }
    for (const b of this.root.querySelectorAll("[data-pop][aria-expanded]")) b.removeAttribute("aria-expanded");
  }

  placePop(x, y, anchorRect = null) {
    const p = this.popEl;
    p.hidden = false;
    p.style.left = "0px"; p.style.top = "0px";
    const rr = this.root.getBoundingClientRect();
    const pr = p.getBoundingClientRect();
    let left = x - rr.left, top = y - rr.top;
    if (left + pr.width > rr.width - 6) left = Math.max(6, rr.width - pr.width - 6);
    if (top + pr.height > rr.height - 6) top = anchorRect ? Math.max(6, anchorRect.top - rr.top - pr.height - 4) : Math.max(6, rr.height - pr.height - 6);
    p.style.left = `${left}px`;
    p.style.top = `${top}px`;
  }

  openMenu(items, x, y) {
    this.closePop();
    this.pop = "menu";
    this.popEl.innerHTML = `<div class="sv-menu" role="menu">${items.map((m, k) => (m === "-" ? '<div class="sv-sep"></div>'
      : `<button type="button" role="menuitem" data-k="${k}"${m.disabled ? " disabled" : ""}>${m.ic ? icon(m.ic) : '<span class="sv-ic"></span>'}<span>${esc(m.label)}</span></button>`)).join("")}</div>`;
    this.placePop(x, y);
    this.popEl.onclick = (e) => {
      const b = e.target.closest("button[data-k]");
      if (!b || b.disabled) return;
      const m = items[+b.dataset.k];
      this.closePop();
      m.run();
    };
    const first = this.popEl.querySelector("button:not([disabled])");
    if (first) first.focus({ preventScroll: true });
    this.popEl.onkeydown = (e) => this.menuKeys(e);
  }

  menuKeys(e) {
    const btns = [...this.popEl.querySelectorAll("button:not([disabled])")];
    const i = btns.indexOf(document.activeElement);
    if (e.key === "ArrowDown") { e.preventDefault(); (btns[i + 1] || btns[0]).focus(); }
    else if (e.key === "ArrowUp") { e.preventDefault(); (btns[i - 1] || btns[btns.length - 1]).focus(); }
    else if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); this.closePop(); this.refocus(); }
  }

  colorGrid(cur, { none = null, auto = null } = {}) {
    const theme = this.deck.theme || {};
    const accent = theme.accent || "#4472C4";
    const themeRow = [theme.bg || "#FFFFFF", theme.color || "#262626", accent, tintColor(accent, 0.4), tintColor(accent, 0.6), tintColor(accent, 0.8), tintColor(accent, -0.25), tintColor(accent, -0.5), tintColor(theme.color || "#262626", 0.5), tintColor(theme.color || "#262626", 0.8)];
    const sw = (c) => `<button type="button" class="sv-sw${hexColor(cur) === c ? " on" : ""}" data-color="${c}" title="${c}" style="background:${c}"></button>`;
    return `${none ? `<button type="button" class="sv-wide" data-color="">${esc(none)}</button>` : ""}${auto ? `<button type="button" class="sv-wide" data-color="auto">${esc(auto)}</button>` : ""}`
      + `<div class="sv-cap">Theme colors</div><div class="sv-swatches">${themeRow.map((c) => sw(hexColor(c))).join("")}</div>`
      + `<div class="sv-cap">Standard colors</div><div class="sv-swatches">${STD_COLORS.map(sw).join("")}</div><div class="sv-swatches">${GRAYS.map(sw).join("")}</div>`
      + `<label class="sv-more">More colors… <input type="color" value="${hexColor(cur) || "#000000"}"></label>`;
  }

  bindColor(fn) {
    this.popEl.onclick = (e) => {
      const b = e.target.closest("[data-color]");
      if (!b) return;
      const c = b.dataset.color;
      this.closePop();
      fn(c === "" ? null : c === "auto" ? undefined : c);
    };
    const inp = this.popEl.querySelector('input[type="color"]');
    if (inp) inp.onchange = () => { const c = hexColor(inp.value); this.closePop(); fn(c); };
  }

  togglePop(name, anchor) {
    if (this.pop === name) { this.closePop(); return; }
    if (this.editing) this.saveRange();
    this.closePop();
    const deck = this.deck;
    const s = this.slide();
    this.pop = name;
    anchor.setAttribute("aria-expanded", "true");
    const p = this.popEl;
    p.onclick = null;
    p.onkeydown = (e) => { if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); this.closePop(); this.refocus(); } };
    let html = "";
    const items = this.sel.map((i) => s.items[i]).filter(Boolean);
    const thumbFor = (layout) => {
      const sl = newSlide(layout, deck);
      for (const it of sl.items) if (it.ph) it.paras = [{ ...(it.paras[0] || {}), runs: [{ t: it.ph === "title" ? "Title" : it.ph === "subtitle" ? "Subtitle" : "• Text", size: it.ph === "title" ? 60 : 40 }] }];
      return `<div class="sv-mini" style="width:${deck.size.w * 0.13}px;height:${deck.size.h * 0.13}px"><div style="transform:scale(0.13);transform-origin:0 0">${renderSlideHtml(sl, deck)}</div></div>`;
    };
    switch (name) {
      case "layoutNew":
      case "layout":
        html = `<div class="sv-cap">${name === "layout" ? "Layout of this slide" : "New slide with layout"}</div><div class="sv-tiles">${LAYOUTS.map((l) => `<button type="button" class="sv-tile${name === "layout" && s.layout === l.id ? " on" : ""}" data-layout="${l.id}">${thumbFor(l.id)}<span>${esc(l.name)}</span></button>`).join("")}</div>`;
        p.innerHTML = `<div class="sv-panel">${html}</div>`;
        p.onclick = (e) => {
          const b = e.target.closest("[data-layout]");
          if (!b) return;
          this.closePop();
          if (name === "layout") this.applyLayout(b.dataset.layout); else this.addSlide(b.dataset.layout);
        };
        break;
      case "shape":
        html = `<div class="sv-cap">Shapes</div><div class="sv-shapes">${SHAPES.map((sh) => `<button type="button" class="sv-shp" data-shape="${sh.id}" title="${esc(sh.name)}" aria-label="${esc(sh.name)}"><svg width="30" height="24" viewBox="-2 -2 34 28" overflow="visible"><path d="${shapePath(sh.id, 30, sh.id === "line" ? 24 : 24)}" fill="${sh.id === "line" ? "none" : "currentColor"}" fill-opacity=".18" stroke="currentColor" stroke-width="1.4"/></svg></button>`).join("")}</div>`;
        p.innerHTML = `<div class="sv-panel">${html}</div>`;
        p.onclick = (e) => { const b = e.target.closest("[data-shape]"); if (!b) return; this.closePop(); this.insertShape(b.dataset.shape); };
        break;
      case "table": {
        const R = 8, C = 10;
        let g = "";
        for (let r = 0; r < R; r++) for (let c = 0; c < C; c++) g += `<button type="button" class="sv-cell" data-r="${r + 1}" data-c="${c + 1}" aria-label="${r + 1} by ${c + 1} table"></button>`;
        p.innerHTML = `<div class="sv-panel"><div class="sv-cap sv-tsize">Insert table</div><div class="sv-tgrid" style="grid-template-columns:repeat(${C},16px)">${g}</div></div>`;
        const label = p.querySelector(".sv-tsize");
        p.onmouseover = (e) => {
          const b = e.target.closest(".sv-cell");
          if (!b) return;
          const rr = +b.dataset.r, cc = +b.dataset.c;
          for (const n of p.querySelectorAll(".sv-cell")) n.classList.toggle("on", +n.dataset.r <= rr && +n.dataset.c <= cc);
          label.textContent = `Insert table: ${rr} × ${cc}`;
        };
        p.onclick = (e) => { const b = e.target.closest(".sv-cell"); if (!b) return; this.closePop(); this.insertTable(+b.dataset.r, +b.dataset.c); };
        break;
      }
      case "textColor": {
        let cur = null;
        const it = items[0];
        if (it) { const r = this.allRuns(it).find((x) => x.t) || {}; cur = r.color; }
        p.innerHTML = `<div class="sv-panel">${this.colorGrid(cur, { auto: "Automatic" })}</div>`;
        this.bindColor((c) => { this.refocus(); this.applyColor(c || null); });
        break;
      }
      case "fill": {
        const it = items[0];
        p.innerHTML = `<div class="sv-panel">${this.colorGrid(it && it.fill, { none: "No fill" })}</div>`;
        this.bindColor((c) => this.applyFill(c || null));
        break;
      }
      case "outline": {
        const it = items[0] || {};
        const widths = [0.75, 1, 1.5, 2.25, 3, 4.5, 6];
        p.innerHTML = `<div class="sv-panel">${this.colorGrid(it.line, { none: "No outline" })}<div class="sv-cap">Weight</div><div class="sv-weights">${widths.map((w) => `<button type="button" class="sv-wt${it.lineW === w ? " on" : ""}" data-w="${w}"><span style="border-top:${Math.max(1, w * 1.33)}px solid currentColor"></span>${w} pt</button>`).join("")}</div></div>`;
        this.bindColor((c) => this.applyLine(c === undefined ? null : c, null));
        const onclick = p.onclick;
        p.onclick = (e) => {
          const w = e.target.closest("[data-w]");
          if (w) { this.closePop(); this.applyLine(undefined, +w.dataset.w); return; }
          onclick(e);
        };
        break;
      }
      case "arrange": {
        const multi = this.sel.length > 1;
        const ops = [
          ["front", "Bring to front", "front"], ["forward", "Bring forward", "up"], ["backward", "Send backward", "down"], ["back", "Send to back", "back"], "-",
          ["left", `Align left${multi ? "" : " (to slide)"}`, "alignL"], ["hcenter", "Align center", "alignC"], ["right", "Align right", "alignR"],
          ["top", "Align top", "alignL"], ["vmiddle", "Align middle", "alignC"], ["bottom", "Align bottom", "alignR"], "-",
          ["vtop", "Text at top", null], ["vmid", "Text in the middle", null], ["vbot", "Text at bottom", null], "-",
          ["dup", "Duplicate (Ctrl+D)", "dup"], ["del", "Delete", "trash"],
        ];
        const menu = ops.map((o) => (o === "-" ? "-" : {
          label: o[1], ic: o[2], run: () => {
            if (o[0] === "dup") this.duplicateItems();
            else if (o[0] === "del") this.deleteItems();
            else if (o[0][0] === "v" && o[0] !== "vmiddle") this.applyValign({ vtop: "top", vmid: "middle", vbot: "bottom" }[o[0]]);
            else this.arrange(o[0]);
          },
        }));
        this.pop = null;
        const r = anchor.getBoundingClientRect();
        this.openMenu(menu, r.left, r.bottom + 4);
        this.pop = "arrange";
        return;
      }
      case "background": {
        p.innerHTML = `<div class="sv-panel">${this.colorGrid(s.bg || deck.theme.bg, { none: "Theme background" })}<button type="button" class="sv-wide" data-all="1">Apply this background to all slides</button></div>`;
        this.bindColor((c) => {
          this.pushUndo();
          if (c) s.bg = c; else delete s.bg;
          this.changed();
          this.status(c ? "Background changed" : "Theme background");
        });
        const onclick = p.onclick;
        p.onclick = (e) => {
          if (e.target.closest("[data-all]")) {
            this.closePop();
            this.pushUndo();
            for (const sl of deck.slides) { if (s.bg) sl.bg = s.bg; else delete sl.bg; }
            this.changed({ all: true });
            this.status("Background applied to all slides");
            return;
          }
          onclick(e);
        };
        break;
      }
      case "transition": {
        const names = { none: "None", fade: "Fade", push: "Push", wipe: "Wipe", zoom: "Zoom" };
        p.innerHTML = `<div class="sv-panel"><div class="sv-cap">Transition to this slide</div><div class="sv-list">${TRANSITIONS.map((t) => `<button type="button" class="sv-li${(s.transition || "none") === t ? " on" : ""}" data-tr="${t}">${names[t]}</button>`).join("")}</div><button type="button" class="sv-wide" data-all="1">Apply to all slides</button></div>`;
        p.onclick = (e) => {
          const b = e.target.closest("[data-tr]");
          if (b) {
            this.closePop();
            this.pushUndo();
            if (b.dataset.tr === "none") delete s.transition; else s.transition = b.dataset.tr;
            this.changed({ strip: true });
            this.previewTransition(s.transition);
            this.status(`Transition: ${names[b.dataset.tr]}`);
          } else if (e.target.closest("[data-all]")) {
            this.closePop();
            this.pushUndo();
            for (const sl of deck.slides) { if (s.transition) sl.transition = s.transition; else delete sl.transition; }
            this.changed({ all: true });
            this.status("Transition applied to all slides");
          }
        };
        break;
      }
      case "theme": {
        p.innerHTML = `<div class="sv-panel"><div class="sv-cap">Themes</div><div class="sv-tiles">${TEMPLATES.map((t) => {
          const th = t.theme;
          const on = deck.theme.font === th.font && deck.theme.accent === th.accent && deck.theme.bg === th.bg;
          return `<button type="button" class="sv-tile sv-theme${on ? " on" : ""}" data-theme="${t.id}"><div class="sv-tprev" style="background:${th.bg};color:${th.color};font-family:'${esc(th.headFont)}',sans-serif"><span class="sv-tt">Aa</span><span class="sv-tbar" style="background:${th.accent}"></span></div><span>${esc(t.name)}</span></button>`;
        }).join("")}</div></div>`;
        p.onclick = (e) => { const b = e.target.closest("[data-theme]"); if (!b) return; this.closePop(); this.applyTheme(TEMPLATES.find((t) => t.id === b.dataset.theme)); };
        break;
      }
      case "slideSize": {
        const cur = `${deck.size.w}x${deck.size.h}`;
        p.innerHTML = `<div class="sv-panel"><div class="sv-cap">Slide size</div><div class="sv-list">${Object.values(SIZES).map((z) => `<button type="button" class="sv-li${cur === `${z.w}x${z.h}` ? " on" : ""}" data-w="${z.w}" data-h="${z.h}">${esc(z.label)}</button>`).join("")}</div></div>`;
        p.onclick = (e) => {
          const b = e.target.closest("[data-w]");
          if (!b) return;
          this.closePop();
          const w = +b.dataset.w, h = +b.dataset.h;
          if (w === deck.size.w && h === deck.size.h) return;
          if (this.editing) this.commitEdit();
          this.pushUndo();
          resizeDeck(deck, w, h);
          this.zoomMode = "fit";
          this.changed({ all: true });
          this.status(`Slide size: ${b.textContent}`);
        };
        break;
      }
      case "zoom": {
        const zs = [0.5, 0.75, 1, 1.25, 1.5, 2];
        p.innerHTML = `<div class="sv-panel"><div class="sv-list"><button type="button" class="sv-li${this.zoomMode === "fit" ? " on" : ""}" data-z="fit">Fit to window</button>${zs.map((z) => `<button type="button" class="sv-li${this.zoomMode !== "fit" && Math.abs(this.scale - z) < 0.01 ? " on" : ""}" data-z="${z}">${z * 100}%</button>`).join("")}</div><div class="sv-hint">Ctrl + mouse wheel also zooms</div></div>`;
        p.onclick = (e) => { const b = e.target.closest("[data-z]"); if (!b) return; this.closePop(); this.setZoom(b.dataset.z === "fit" ? "fit" : +b.dataset.z); };
        break;
      }
      case "present": {
        this.pop = null;
        const r = anchor.getBoundingClientRect();
        this.openMenu([
          { label: "From beginning (F5)", ic: "play", run: () => this.startShow(0) },
          { label: "From current slide (Shift+F5)", ic: "play", run: () => this.startShow(this.current) },
          "-",
          { label: s.hidden ? "Show this slide in the slide show" : "Hide this slide in the slide show", ic: s.hidden ? "eye" : "eyeOff", run: () => this.cmd("hideSlide") },
        ], r.left, r.bottom + 4);
        this.pop = "present";
        return;
      }
      default:
        this.pop = null;
        return;
    }
    const r = anchor.getBoundingClientRect();
    this.placePop(r.left, r.bottom + 4, r);
    const f = p.querySelector("button");
    if (f && name !== "textColor") f.focus({ preventScroll: true });
  }

  previewTransition(t) {
    if (!t || t === "none" || !this.host) return;
    this.host.classList.remove("sv-tr-fade", "sv-tr-push", "sv-tr-wipe", "sv-tr-zoom");
    void this.host.offsetWidth;
    this.host.classList.add(`sv-tr-${t}`);
    this.later(() => this.host && this.host.classList.remove(`sv-tr-${t}`), 700);
  }

  // ------------------------------------------------------------------ slide show

  startShow(from = 0) {
    if (this.show) return;
    if (this.editing) this.commitEdit();
    this.closePop();
    const deck = this.deck;
    const visible = deck.slides.map((s, i) => (s.hidden ? -1 : i)).filter((i) => i >= 0);
    let list = visible;
    if (deck.slides[from] && deck.slides[from].hidden) list = [...new Set([from, ...visible.filter((i) => i > from)])].sort((a, b) => a - b);
    if (!list.length) { this.status("All slides are hidden: there is nothing to present"); return; }
    let pos = list.findIndex((i) => i >= from);
    if (pos < 0) pos = 0;
    const el = document.createElement("div");
    el.className = "sv-show";
    el.tabIndex = -1;
    el.setAttribute("role", "dialog");
    el.setAttribute("aria-label", "Slide show");
    el.innerHTML = '<div class="sv-show-stage"></div><div class="sv-show-black" hidden></div><div class="sv-show-end" hidden>End of slide show. Click or press any key to exit.</div>'
      + '<div class="sv-show-notes" hidden></div>'
      + '<div class="sv-show-bar"><button type="button" data-a="prev" title="Previous (Left arrow)" aria-label="Previous slide">‹</button><span class="sv-show-count"></span><button type="button" data-a="next" title="Next (Right arrow, Space)" aria-label="Next slide">›</button><button type="button" data-a="notes" title="Speaker notes (N)">Notes</button><button type="button" data-a="exit" title="End show (Esc)">End show</button></div>';
    this.root.appendChild(el);
    this.show = { el, list, pos, end: false, fs: false, black: false, notes: false };
    this.showListeners = [];
    const on = (t, type, fn, o) => { t.addEventListener(type, fn, o); this.showListeners.push([t, type, fn, o]); };
    on(el, "click", (e) => {
      const b = e.target.closest("button[data-a]");
      if (b) { e.stopPropagation(); const a = b.dataset.a; if (a === "prev") this.showGo(-1); else if (a === "next") this.showGo(1); else if (a === "notes") this.toggleShowNotes(); else this.endShow(); return; }
      if (e.target.closest("a[href]")) return;
      this.showGo(1);
    });
    on(el, "contextmenu", (e) => { e.preventDefault(); this.showGo(-1); });
    on(el, "wheel", (e) => { if (Math.abs(e.deltaY) > 20) this.showGo(e.deltaY > 0 ? 1 : -1); }, { passive: true });
    on(el, "mousemove", () => {
      el.classList.add("sv-show-ui");
      if (this.show.uiTimer) clearTimeout(this.show.uiTimer);
      this.show.uiTimer = setTimeout(() => { if (this.show) this.show.el.classList.remove("sv-show-ui"); }, 2200);
    });
    on(window, "resize", () => this.layoutShow());
    on(window, "keydown", (e) => this.onShowKey(e), true);
    this.paintShow(false);
    el.focus({ preventScroll: true });
    if (el.requestFullscreen) {
      el.requestFullscreen().then(() => { if (this.show) { this.show.fs = true; this.layoutShow(); } }).catch(() => { /* overlay only */ });
    }
    this.status("Slide show: arrows or click to move, Esc to end, N for notes, B for a black screen");
  }

  onShowKey(e) {
    if (!this.show) return;
    if (e.__svHandled) return;
    e.__svHandled = true;
    const k = e.key;
    const sh = this.show;
    e.stopPropagation();
    if (k === "Escape") { e.preventDefault(); this.endShow(); return; }
    if (sh.end) { e.preventDefault(); if (k === "ArrowLeft" || k === "PageUp" || k === "Backspace") { sh.end = false; this.paintShow(false); } else this.endShow(); return; }
    if (k === "ArrowRight" || k === "ArrowDown" || k === " " || k === "PageDown" || k === "Enter") { e.preventDefault(); this.showGo(1); }
    else if (k === "ArrowLeft" || k === "ArrowUp" || k === "PageUp" || k === "Backspace") { e.preventDefault(); this.showGo(-1); }
    else if (k === "Home") { e.preventDefault(); sh.pos = 0; this.paintShow(false); }
    else if (k === "End") { e.preventDefault(); sh.pos = sh.list.length - 1; this.paintShow(false); }
    else if (k === "n" || k === "N") { e.preventDefault(); this.toggleShowNotes(); }
    else if (k === "b" || k === "B" || k === ".") { e.preventDefault(); sh.black = !sh.black; sh.el.querySelector(".sv-show-black").hidden = !sh.black; }
    else if (k === "w" || k === "W" || k === ",") { e.preventDefault(); sh.black = !sh.black; const b = sh.el.querySelector(".sv-show-black"); b.hidden = !sh.black; b.classList.toggle("white", sh.black); }
    else if (/^[0-9]$/.test(k)) { sh.typed = (sh.typed || "") + k; }
    else if (k === "Enter" && sh.typed) { e.preventDefault(); }
    else if (k === "F5") e.preventDefault();
    else if (e.ctrlKey || e.metaKey) e.preventDefault();
  }

  toggleShowNotes() {
    const sh = this.show;
    sh.notes = !sh.notes;
    this.paintShow(false);
  }

  showGo(d) {
    const sh = this.show;
    if (!sh) return;
    if (sh.black) { sh.black = false; sh.el.querySelector(".sv-show-black").hidden = true; return; }
    if (sh.typed) {
      const n = parseInt(sh.typed, 10) - 1;
      sh.typed = "";
      const at = sh.list.indexOf(n);
      if (at >= 0) { sh.pos = at; this.paintShow(true); return; }
    }
    if (sh.end) { if (d > 0) this.endShow(); else { sh.end = false; this.paintShow(false); } return; }
    const np = sh.pos + d;
    if (np >= sh.list.length) { sh.end = true; this.paintShow(false); return; }
    if (np < 0) return;
    sh.pos = np;
    this.paintShow(d > 0);
  }

  paintShow(animate) {
    const sh = this.show;
    const deck = this.deck;
    const stage = sh.el.querySelector(".sv-show-stage");
    sh.el.querySelector(".sv-show-end").hidden = !sh.end;
    stage.hidden = sh.end;
    if (!sh.end) {
      const idx = sh.list[sh.pos];
      const s = deck.slides[idx];
      stage.innerHTML = `<div class="sv-show-slide">${renderSlideHtml(s, deck)}</div>`;
      const t = s.transition;
      if (animate && t && t !== "none") stage.firstElementChild.classList.add(`sv-tr-${t}`);
      sh.el.querySelector(".sv-show-count").textContent = `${sh.pos + 1} / ${sh.list.length}`;
      const notes = sh.el.querySelector(".sv-show-notes");
      notes.hidden = !sh.notes;
      notes.textContent = s.notes && s.notes.trim() ? s.notes : "No speaker notes for this slide.";
      sh.lastIndex = idx;
    }
    this.layoutShow();
  }

  layoutShow() {
    const sh = this.show;
    if (!sh) return;
    const deck = this.deck;
    const slideEl = sh.el.querySelector(".sv-show-slide");
    if (!slideEl) return;
    const r = sh.el.getBoundingClientRect();
    const notesH = sh.notes ? Math.min(r.height * 0.3, 180) : 0;
    const k = Math.min(r.width / deck.size.w, (r.height - notesH) / deck.size.h);
    slideEl.style.width = `${deck.size.w}px`;
    slideEl.style.height = `${deck.size.h}px`;
    slideEl.style.left = `${(r.width - deck.size.w * k) / 2}px`;
    slideEl.style.top = `${(r.height - notesH - deck.size.h * k) / 2}px`;
    slideEl.style.transform = `scale(${k})`;
  }

  endShow() {
    const sh = this.show;
    if (!sh) return;
    this.show = null;
    for (const [t, type, fn, o] of this.showListeners || []) t.removeEventListener(type, fn, o);
    this.showListeners = [];
    if (sh.uiTimer) clearTimeout(sh.uiTimer);
    if (document.fullscreenElement === sh.el && document.exitFullscreen) document.exitFullscreen().catch(() => {});
    sh.el.remove();
    if (!this.destroyed && sh.lastIndex != null && sh.lastIndex !== this.current && this.thumbs) this.selectSlide(sh.lastIndex);
    if (!this.destroyed) { this.status("Slide show ended"); this.focus(); }
  }

  // ------------------------------------------------------------------ print

  beforePrint() {
    if (this.editing) this.commitEdit();
    const deck = this.deck;
    if (!deck || !this.root) return;
    if (this.printEl) this.printEl.remove();
    const W = deck.size.w, H = deck.size.h;
    // fit inside a landscape page (A4 or Letter, 0.4in margins) at 96 px per inch
    const k = Math.min(1, 980 / W, 680 / H);
    const el = document.createElement("div");
    el.className = "sv-print";
    el.innerHTML = deck.slides.filter((s) => !s.hidden).map((s) => `<div class="sv-print-page"><div class="sv-print-box" style="width:${W * k}px;height:${H * k}px"><div style="transform:scale(${k});transform-origin:0 0">${renderSlideHtml(s, deck)}</div></div></div>`).join("");
    this.root.appendChild(el);
    this.printEl = el;
  }

  afterPrint() {
    if (this.printEl) { this.printEl.remove(); this.printEl = null; }
  }
}
