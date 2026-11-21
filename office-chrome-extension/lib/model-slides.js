/**
 * Slides (PowerPoint) document model helpers (owned by the slides-frontend stream).
 *
 * CONTRACT (shared with src/office-backend/src/model.rs "slides" section):
 * deck  = { kind:"slides", meta:{ title?, author?, warnings?:[string] },
 *           size:{ w, h }   points; 16:9 = 960x540 (default), 4:3 = 720x540
 *           theme:{ font?, headFont?, color?:"#RRGGBB" (body text), accent?:"#RRGGBB",
 *                   bg?:"#RRGGBB" (default slide background) },
 *           slides:[Slide] }
 * slide = { layout?:"title"|"content"|"section"|"two"|"blank", bg?:"#RRGGBB",
 *           notes?:string ("\n" between paragraphs), hidden?:bool,
 *           transition?:"none"|"fade"|"push"|"wipe"|"zoom", items:[Item] }  (z-order = array order)
 * Item  — geometry in points from the slide's top-left; rot = degrees clockwise:
 *   text  { type:"text",  x,y,w,h, rot?, ph?:"title"|"subtitle"|"body", paras:[Para],
 *           fill?, line?, lineW?, valign?:"top"|"middle"|"bottom" }
 *   shape { type:"shape", x,y,w,h, rot?, shape:"rect"|"roundRect"|"ellipse"|"triangle"|"diamond"
 *           |"rightArrow"|"leftArrow"|"line"|"star5"|"hexagon"|"chevron", fill?, line?, lineW?,
 *           paras?:[Para], valign? }
 *   image { type:"image", x,y,w,h, rot?, src:"data:image/...;base64,...", alt? }
 *   table { type:"table", x,y,w,h, rows:[[{ paras:[Para], fill? }]], colW?:[pt] }
 *   Unknown item types must be kept as they are.
 * Para  = { runs:[Run], align?:"left"|"center"|"right"|"justify", level?:0..8, bullet?:"ul"|"ol" }
 * Run   = the word Run { t, b?, i?, u?, s?, sup?, sub?, color?, size?(pt), font?, link?, hl? }
 *         (a "\n" inside run text is a line break inside the paragraph)
 * Colors "#RRGGBB"; absent fill/line = none. lineW in points.
 *
 * Exports used by viewer/viewer.js:
 *   newDeck(template?)                -> deck   (template: undefined = blank title slide)
 *   normalizeDeck(model)              -> deck   (fills defaults, never throws on partial input)
 *   slidesToHtmlDocument(deck, title) -> string (standalone HTML slideshow)
 *   slidesToMarkdown(deck)            -> string (outline)
 *   slidesToText(deck)                -> string (outline)
 * Also used by viewer/slides.js and lib/local-pptx.js:
 *   TEMPLATES, LAYOUTS, SHAPES, newSlide(layout, deck), renderSlideHtml(slide, deck, opts),
 *   renderItemHtml(item, deck, index, opts), shapePath(shape, w, h), defaultRunSize(item, level),
 *   textColorFor(item, deck), slideBg(slide, deck), SLIDE_CSS, cloneDeep, paraText, itemText,
 *   slideTitle, hexColor, luminance.
 */

// ---------------------------------------------------------------------------
// constants
// ---------------------------------------------------------------------------

export const SIZES = {
  wide: { w: 960, h: 540, label: "Widescreen (16:9)" },
  standard: { w: 720, h: 540, label: "Standard (4:3)" },
};

/** Built-in templates: a theme is colors + fonts applied to placeholders and backgrounds. */
export const TEMPLATES = [
  { id: "blank", name: "Blank", theme: { font: "Calibri", headFont: "Calibri Light", color: "#262626", accent: "#4472C4", bg: "#FFFFFF" } },
  { id: "modern", name: "Modern blue", theme: { font: "Segoe UI", headFont: "Segoe UI Semibold", color: "#1F2A44", accent: "#1A56DB", bg: "#F5F8FF" } },
  { id: "dark", name: "Dark", theme: { font: "Segoe UI", headFont: "Segoe UI Light", color: "#F2F4F7", accent: "#4CC2FF", bg: "#1E2129" } },
  { id: "warm", name: "Warm", theme: { font: "Georgia", headFont: "Georgia", color: "#3D2C1E", accent: "#C8553D", bg: "#FBF5EC" } },
  { id: "fresh", name: "Fresh green", theme: { font: "Trebuchet MS", headFont: "Trebuchet MS", color: "#1E3A2F", accent: "#2E9D6A", bg: "#FFFFFF" } },
  { id: "bold", name: "Bold red", theme: { font: "Arial", headFont: "Arial Black", color: "#202124", accent: "#C43E1C", bg: "#FFFFFF" } },
];

export const DEFAULT_THEME = TEMPLATES[0].theme;

export const LAYOUTS = [
  { id: "title", name: "Title slide" },
  { id: "content", name: "Title and content" },
  { id: "section", name: "Section header" },
  { id: "two", name: "Two content" },
  { id: "blank", name: "Blank" },
];

export const SHAPES = [
  { id: "rect", name: "Rectangle" },
  { id: "roundRect", name: "Rounded rectangle" },
  { id: "ellipse", name: "Oval" },
  { id: "triangle", name: "Triangle" },
  { id: "diamond", name: "Diamond" },
  { id: "rightArrow", name: "Right arrow" },
  { id: "leftArrow", name: "Left arrow" },
  { id: "line", name: "Line" },
  { id: "star5", name: "Star" },
  { id: "hexagon", name: "Hexagon" },
  { id: "chevron", name: "Chevron" },
];

export const TRANSITIONS = ["none", "fade", "push", "wipe", "zoom"];

const LAYOUT_IDS = new Set(LAYOUTS.map((l) => l.id));
const SHAPE_IDS = new Set(SHAPES.map((s) => s.id));
const PH_IDS = new Set(["title", "subtitle", "body"]);
const ALIGNS = new Set(["left", "center", "right", "justify"]);
const VALIGNS = new Set(["top", "middle", "bottom"]);

export const PROMPTS = { title: "Click to add title", subtitle: "Click to add subtitle", body: "Click to add text" };

// ---------------------------------------------------------------------------
// small helpers
// ---------------------------------------------------------------------------

/** "#abc" / "ABCDEF" / "#aabbcc" -> "#AABBCC"; anything else -> undefined. */
export function hexColor(c) {
  if (typeof c !== "string") return undefined;
  let s = c.trim();
  if (s[0] === "#") s = s.slice(1);
  if (/^[0-9a-fA-F]{3}$/.test(s)) s = s.split("").map((ch) => ch + ch).join("");
  if (!/^[0-9a-fA-F]{6}$/.test(s)) return undefined;
  return `#${s.toUpperCase()}`;
}

/** Relative luminance 0..1 of "#RRGGBB" (0.5 for invalid input). */
export function luminance(hex) {
  const h = hexColor(hex);
  if (!h) return 0.5;
  const ch = [1, 3, 5].map((i) => {
    const v = parseInt(h.slice(i, i + 2), 16) / 255;
    return v <= 0.03928 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4;
  });
  return 0.2126 * ch[0] + 0.7152 * ch[1] + 0.0722 * ch[2];
}

/** Mix a color with white (amt > 0) or black (amt < 0). */
export function tintColor(hex, amt) {
  const h = hexColor(hex) || "#808080";
  const out = [1, 3, 5].map((i) => {
    const v = parseInt(h.slice(i, i + 2), 16);
    const r = amt >= 0 ? v + (255 - v) * amt : v * (1 + amt);
    return Math.max(0, Math.min(255, Math.round(r))).toString(16).padStart(2, "0");
  });
  return `#${out.join("").toUpperCase()}`;
}

/** Deep copy of plain JSON-like data (strings are shared, so image data URLs are cheap). */
export function cloneDeep(v) {
  if (Array.isArray(v)) return v.map(cloneDeep);
  if (v && typeof v === "object") {
    const o = {};
    for (const k of Object.keys(v)) o[k] = cloneDeep(v[k]);
    return o;
  }
  return v;
}

const num = (v, d = 0) => (typeof v === "number" && Number.isFinite(v) ? v : (typeof v === "string" && v.trim() !== "" && Number.isFinite(+v) ? +v : d));

export function paraText(p) {
  return ((p && p.runs) || []).map((r) => (r && r.t) || "").join("");
}

export function itemText(it) {
  if (!it) return "";
  if (it.type === "table") return (it.rows || []).map((r) => (r || []).map((c) => ((c && c.paras) || []).map(paraText).join(" ")).join("\t")).join("\n");
  return (it.paras || []).map(paraText).join("\n");
}

/** Title text of a slide: the title placeholder, else "" */
export function slideTitle(slide) {
  const t = ((slide && slide.items) || []).find((it) => it && it.type === "text" && it.ph === "title");
  return t ? itemText(t).replace(/\s+/g, " ").trim() : "";
}

function esc(s) {
  return String(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
}

// ---------------------------------------------------------------------------
// normalization
// ---------------------------------------------------------------------------

function normRun(r) {
  if (!r || typeof r !== "object") return { t: String(r ?? "") };
  const o = { ...r };
  o.t = typeof r.t === "string" ? r.t : String(r.t ?? "");
  for (const k of ["b", "i", "u", "s", "sup", "sub"]) if (o[k] !== undefined && typeof o[k] !== "boolean") o[k] = !!o[k];
  for (const k of ["color", "hl"]) if (o[k] !== undefined) { const h = hexColor(o[k]); if (h) o[k] = h; else delete o[k]; }
  if (o.size !== undefined) { const s = num(o.size, 0); if (s > 0 && s < 4000) o.size = s; else delete o.size; }
  if (o.font !== undefined && typeof o.font !== "string") delete o.font;
  if (o.link !== undefined && typeof o.link !== "string") delete o.link;
  return o;
}

export function normPara(p) {
  if (!p || typeof p !== "object") return { runs: [] };
  const o = { ...p };
  o.runs = Array.isArray(p.runs) ? p.runs.map(normRun) : [];
  if (o.align !== undefined && !ALIGNS.has(o.align)) delete o.align;
  if (o.level !== undefined) { const l = Math.round(num(o.level, 0)); if (l > 0) o.level = Math.min(8, l); else delete o.level; }
  if (o.bullet !== undefined && o.bullet !== "ul" && o.bullet !== "ol") delete o.bullet;
  return o;
}

function normItem(it) {
  if (!it || typeof it !== "object") return null;
  if (!["text", "shape", "image", "table"].includes(it.type)) return it; // unknown: kept as-is
  const o = it;
  for (const k of ["x", "y", "w", "h"]) o[k] = num(o[k], 0);
  if (o.w < 0) { o.x += o.w; o.w = -o.w; }
  if (o.h < 0) { o.y += o.h; o.h = -o.h; }
  if (o.rot !== undefined) { const r = num(o.rot, 0) % 360; if (r) o.rot = r < 0 ? r + 360 : r; else delete o.rot; }
  for (const k of ["fill", "line"]) if (o[k] !== undefined) { const h = hexColor(o[k]); if (h) o[k] = h; else delete o[k]; }
  if (o.lineW !== undefined) { const w = num(o.lineW, 0); if (w > 0) o.lineW = Math.min(w, 200); else delete o.lineW; }
  if (o.valign !== undefined && !VALIGNS.has(o.valign)) delete o.valign;
  if (o.type === "text") {
    o.paras = Array.isArray(o.paras) ? o.paras.map(normPara) : [];
    if (o.ph !== undefined && !PH_IDS.has(o.ph)) delete o.ph;
  } else if (o.type === "shape") {
    if (!SHAPE_IDS.has(o.shape)) o.shape = "rect";
    if (o.paras !== undefined) o.paras = Array.isArray(o.paras) ? o.paras.map(normPara) : [];
  } else if (o.type === "image") {
    if (typeof o.src !== "string") o.src = "";
    if (o.alt !== undefined && typeof o.alt !== "string") delete o.alt;
  } else if (o.type === "table") {
    o.rows = Array.isArray(o.rows) ? o.rows.filter(Array.isArray).map((r) => r.map((c) => {
      const cell = c && typeof c === "object" ? { ...c } : { paras: [] };
      cell.paras = Array.isArray(cell.paras) ? cell.paras.map(normPara) : [];
      if (cell.fill !== undefined) { const h = hexColor(cell.fill); if (h) cell.fill = h; else delete cell.fill; }
      return cell;
    })) : [];
    if (o.colW !== undefined) {
      if (Array.isArray(o.colW)) o.colW = o.colW.map((v) => Math.max(0, num(v, 0)));
      else delete o.colW;
    }
  }
  return o;
}

/** Fill defaults in place; never throws on partial input. */
export function normalizeDeck(model) {
  const deck = model && typeof model === "object" && !Array.isArray(model) ? model : {};
  deck.kind = "slides";
  if (!deck.meta || typeof deck.meta !== "object" || Array.isArray(deck.meta)) deck.meta = {};
  if (!deck.size || typeof deck.size !== "object") deck.size = { w: 960, h: 540 };
  deck.size.w = num(deck.size.w, 960);
  deck.size.h = num(deck.size.h, 540);
  if (!(deck.size.w > 0) || deck.size.w > 20000) deck.size.w = 960;
  if (!(deck.size.h > 0) || deck.size.h > 20000) deck.size.h = 540;
  if (!deck.theme || typeof deck.theme !== "object" || Array.isArray(deck.theme)) deck.theme = {};
  for (const k of ["color", "accent", "bg"]) if (deck.theme[k] !== undefined) { const h = hexColor(deck.theme[k]); if (h) deck.theme[k] = h; else delete deck.theme[k]; }
  for (const k of ["font", "headFont"]) if (deck.theme[k] !== undefined && (typeof deck.theme[k] !== "string" || !deck.theme[k].trim())) delete deck.theme[k];
  if (!Array.isArray(deck.slides)) deck.slides = [];
  deck.slides = deck.slides.map((s) => (s && typeof s === "object" && !Array.isArray(s) ? s : { items: [] }));
  if (!deck.slides.length) deck.slides.push({ layout: "blank", items: [] });
  for (const s of deck.slides) {
    if (!Array.isArray(s.items)) s.items = [];
    s.items = s.items.map(normItem).filter((x) => x !== null);
    if (s.layout !== undefined && !LAYOUT_IDS.has(s.layout)) delete s.layout;
    if (s.bg !== undefined) { const h = hexColor(s.bg); if (h) s.bg = h; else delete s.bg; }
    if (s.notes !== undefined && typeof s.notes !== "string") s.notes = s.notes == null ? "" : String(s.notes);
    if (s.hidden !== undefined) s.hidden = !!s.hidden;
    if (s.transition !== undefined && !TRANSITIONS.includes(s.transition)) delete s.transition;
  }
  return deck;
}

// ---------------------------------------------------------------------------
// factories
// ---------------------------------------------------------------------------

function templateById(id) {
  const key = String(id || "").toLowerCase().replace(/[^a-z]/g, "");
  return TEMPLATES.find((t) => t.id === key || t.name.toLowerCase().replace(/[^a-z]/g, "") === key) || null;
}

function phPara(size, extra = {}) {
  return { ...extra, runs: [{ t: "", size }] };
}

/** Placeholder geometry for a layout (fractions of the slide size, PowerPoint-like). */
export function layoutBoxes(layout, size) {
  const W = (size && size.w) || 960;
  const H = (size && size.h) || 540;
  const mx = W * 0.0688;
  const cw = W - 2 * mx;
  const titleBox = { x: mx, y: H * 0.0533, w: cw, h: H * 0.1935 };
  switch (layout) {
    case "title":
      return { title: { x: W * 0.125, y: H * 0.1637, w: W * 0.75, h: H * 0.348 }, subtitle: { x: W * 0.125, y: H * 0.5265, w: W * 0.75, h: H * 0.2413 } };
    case "content":
      return { title: titleBox, body: { x: mx, y: H * 0.2667, w: cw, h: H * 0.635 } };
    case "section":
      return { title: { x: W * 0.0682, y: H * 0.2493, w: cw, h: H * 0.416 }, body: { x: W * 0.0682, y: H * 0.6693, w: cw, h: H * 0.2187 } };
    case "two": {
      const gap = W * 0.0125;
      const half = (cw - gap) / 2;
      return { title: titleBox, left: { x: mx, y: H * 0.2667, w: half, h: H * 0.635 }, right: { x: mx + half + gap, y: H * 0.2667, w: half, h: H * 0.635 } };
    }
    default:
      return {};
  }
}

const r2 = (v) => Math.round(v * 100) / 100;
const box = (b) => ({ x: r2(b.x), y: r2(b.y), w: r2(b.w), h: r2(b.h) });

/** A new slide with placeholder text items for `layout`. */
export function newSlide(layout = "content", deck = null) {
  const size = (deck && deck.size) || { w: 960, h: 540 };
  const lay = LAYOUT_IDS.has(layout) ? layout : "content";
  const b = layoutBoxes(lay, size);
  const items = [];
  if (lay === "title") {
    items.push({ type: "text", ph: "title", ...box(b.title), valign: "bottom", paras: [phPara(54, { align: "center" })] });
    items.push({ type: "text", ph: "subtitle", ...box(b.subtitle), valign: "top", paras: [phPara(24, { align: "center" })] });
  } else if (lay === "content") {
    items.push({ type: "text", ph: "title", ...box(b.title), valign: "middle", paras: [phPara(40)] });
    items.push({ type: "text", ph: "body", ...box(b.body), valign: "top", paras: [phPara(24, { bullet: "ul" })] });
  } else if (lay === "section") {
    items.push({ type: "text", ph: "title", ...box(b.title), valign: "bottom", paras: [phPara(54)] });
    items.push({ type: "text", ph: "body", ...box(b.body), valign: "top", paras: [phPara(24)] });
  } else if (lay === "two") {
    items.push({ type: "text", ph: "title", ...box(b.title), valign: "middle", paras: [phPara(40)] });
    items.push({ type: "text", ph: "body", ...box(b.left), valign: "top", paras: [phPara(24, { bullet: "ul" })] });
    items.push({ type: "text", ph: "body", ...box(b.right), valign: "top", paras: [phPara(24, { bullet: "ul" })] });
  }
  return { layout: lay, items };
}

/** A new deck: one title slide, themed by `template` (id or name; default "Blank"). */
export function newDeck(template) {
  const t = templateById(template) || TEMPLATES[0];
  const size = /4[:x]3|standard/i.test(String(template || "")) ? { w: 720, h: 540 } : { w: 960, h: 540 };
  const deck = { kind: "slides", meta: {}, size, theme: { ...t.theme }, slides: [] };
  deck.slides.push(newSlide("title", deck));
  return normalizeDeck(deck);
}

/** Move/scale every item when the slide size changes (keeps relative placement). */
export function resizeDeck(deck, w, h) {
  const sx = w / deck.size.w;
  const sy = h / deck.size.h;
  const s = Math.min(sx, sy);
  for (const slide of deck.slides) {
    for (const it of slide.items) {
      if (!it || typeof it.x !== "number") continue;
      // keep aspect of pictures/shapes: scale uniformly around the scaled centre
      const cx = (it.x + it.w / 2) * sx;
      const cy = (it.y + it.h / 2) * sy;
      const uniform = it.type === "image" || it.type === "shape";
      const nw = uniform ? it.w * s : it.w * sx;
      const nh = uniform ? it.h * s : it.h * sy;
      it.x = r2(cx - nw / 2); it.y = r2(cy - nh / 2); it.w = r2(nw); it.h = r2(nh);
      if (it.type === "table" && Array.isArray(it.colW)) it.colW = it.colW.map((c) => r2(c * sx));
    }
  }
  deck.size = { w, h };
}

// ---------------------------------------------------------------------------
// rendering (shared by the editor, thumbnails, the slideshow and HTML export)
// ---------------------------------------------------------------------------

/** Default font size (pt) of a paragraph at `level` in `item` when its runs have none. */
export function defaultRunSize(item, level = 0) {
  if (!item) return 18;
  if (item.type === "table") return 18;
  if (item.ph === "title") return 44;
  if (item.ph === "subtitle") return 24;
  if (item.ph === "body") return [28, 24, 20, 20, 20, 20, 20, 20, 20][Math.min(8, level || 0)];
  return 18;
}

export function slideBg(slide, deck) {
  return hexColor(slide && slide.bg) || hexColor(deck && deck.theme && deck.theme.bg) || "#FFFFFF";
}

/** Text color used when a run has none (contrast-aware for filled shapes / table cells). */
export function textColorFor(item, deck, fill) {
  const theme = (deck && deck.theme) || {};
  const base = hexColor(theme.color) || "#262626";
  const f = hexColor(fill !== undefined ? fill : item && item.type === "shape" ? item.fill : undefined);
  if (!f) return base;
  const lf = luminance(f);
  const lb = luminance(base);
  // keep the theme text color when it contrasts enough with the fill
  const contrast = (Math.max(lf, lb) + 0.05) / (Math.min(lf, lb) + 0.05);
  if (contrast >= 3) return base;
  return lf < 0.4 ? "#FFFFFF" : "#1F1F1F";
}

function fontStack(name) {
  const n = String(name || "Calibri").replace(/["\\<>;{}]/g, "");
  return `"${n}", "Segoe UI", Calibri, Carlito, Arial, sans-serif`;
}

function cssSafeUrl(u) {
  return String(u).replace(/["\\\n\r]/g, "");
}

const SAFE_LINK = /^(https?:|mailto:|tel:|#)/i;

function runHtml(r) {
  const st = [];
  if (r.b) st.push("font-weight:700");
  if (r.i) st.push("font-style:italic");
  const deco = [];
  if (r.u) deco.push("underline");
  if (r.s) deco.push("line-through");
  if (deco.length) st.push(`text-decoration:${deco.join(" ")}`);
  if (r.color) st.push(`color:${r.color}`);
  if (r.hl) st.push(`background-color:${r.hl}`);
  if (r.size) st.push(`font-size:${r.size}px`);
  if (r.font) st.push(`font-family:${esc(fontStack(r.font))}`);
  if (r.sup) st.push("vertical-align:super;font-size:" + (r.size ? r.size * 0.65 + "px" : "65%"));
  else if (r.sub) st.push("vertical-align:sub;font-size:" + (r.size ? r.size * 0.65 + "px" : "65%"));
  const attrs = [];
  if (st.length) attrs.push(`style="${st.join(";")}"`);
  if (r.sup) attrs.push('data-sup="1"');
  if (r.sub) attrs.push('data-sub="1"');
  if (r.size) attrs.push(`data-sz="${r.size}"`);
  if (r.font) attrs.push(`data-font="${esc(r.font)}"`);
  const text = esc(r.t || "").replace(/\n/g, "<br>");
  let html = `<span${attrs.length ? " " + attrs.join(" ") : ""}>${text}</span>`;
  if (r.link && SAFE_LINK.test(r.link)) html = `<a href="${esc(r.link)}" target="_blank" rel="noopener noreferrer">${html}</a>`;
  else if (r.link) html = `<a data-href="${esc(r.link)}">${html}</a>`;
  return html;
}

/** Paragraphs -> HTML (<p> elements); `ctx` = { item, deck, color, edit, prompt }. */
export function parasHtml(paras, ctx) {
  ctx = ctx || {};
  const out = [];
  const counters = [];
  const list = Array.isArray(paras) && paras.length ? paras : [{ runs: [] }];
  list.forEach((p, idx) => {
    const level = p.level || 0;
    const firstSized = (p.runs || []).find((r) => r && r.size);
    const size = firstSized ? firstSized.size : defaultRunSize(ctx.item, level);
    let num = 0;
    if (p.bullet === "ol") {
      counters.length = level + 1;
      num = (counters[level] || 0) + 1;
      counters[level] = num;
    } else {
      counters.length = level;
    }
    const st = [`font-size:${size}px`];
    if (p.align) st.push(`text-align:${p.align}`);
    const indent = level * 36 + (p.bullet ? 22 : 0);
    if (indent) st.push(`padding-left:${indent}px`);
    if (p.bullet) st.push("text-indent:-22px");
    if (idx > 0 && ctx.item && (ctx.item.ph === "body" || ctx.item.ph === "subtitle")) st.push(`margin-top:${Math.round(size * 0.35)}px`);
    const firstRun = (p.runs || []).find((r) => r && r.t);
    if (p.bullet && firstRun && firstRun.color) st.push(`--bc:${firstRun.color}`);
    const attrs = [`style="${st.join(";")}"`];
    if (level) attrs.push(`data-level="${level}"`);
    if (p.bullet) attrs.push(`data-bullet="${p.bullet}"`);
    if (p.bullet === "ol") attrs.push(`data-num="${num}."`);
    if (p.align) attrs.push(`data-align="${p.align}"`);
    if (firstSized) attrs.push(`data-psz="${size}"`);
    const runs = (p.runs || []).filter((r) => r && (r.t || "") !== "");
    if (idx === 0 && ctx.prompt) attrs.push(`data-prompt="${esc(ctx.prompt)}"`);
    const body = runs.length ? runs.map(runHtml).join("") : idx === 0 && ctx.prompt ? "" : "<br>";
    out.push(`<p ${attrs.join(" ")}>${body}</p>`);
  });
  return out.join("");
}

function hasText(paras) {
  return (paras || []).some((p) => paraText(p).trim() !== "");
}

/** SVG path data for a preset shape in a w x h box. */
export function shapePath(shape, w, h) {
  const f = (v) => Math.round(v * 100) / 100;
  const P = (pts) => "M" + pts.map(([x, y]) => `${f(x)} ${f(y)}`).join(" L") + " Z";
  const ss = Math.min(w, h);
  switch (shape) {
    case "roundRect": {
      const r = ss * 0.1667;
      return `M${f(r)} 0 H${f(w - r)} A${f(r)} ${f(r)} 0 0 1 ${f(w)} ${f(r)} V${f(h - r)} A${f(r)} ${f(r)} 0 0 1 ${f(w - r)} ${f(h)} H${f(r)} A${f(r)} ${f(r)} 0 0 1 0 ${f(h - r)} V${f(r)} A${f(r)} ${f(r)} 0 0 1 ${f(r)} 0 Z`;
    }
    case "ellipse": {
      const rx = w / 2, ry = h / 2;
      return `M0 ${f(ry)} A${f(rx)} ${f(ry)} 0 1 1 ${f(w)} ${f(ry)} A${f(rx)} ${f(ry)} 0 1 1 0 ${f(ry)} Z`;
    }
    case "triangle": return P([[w / 2, 0], [w, h], [0, h]]);
    case "diamond": return P([[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]]);
    case "rightArrow": {
      const hd = Math.min(w, h * 0.5);
      return P([[0, h * 0.25], [w - hd, h * 0.25], [w - hd, 0], [w, h / 2], [w - hd, h], [w - hd, h * 0.75], [0, h * 0.75]]);
    }
    case "leftArrow": {
      const hd = Math.min(w, h * 0.5);
      return P([[w, h * 0.25], [hd, h * 0.25], [hd, 0], [0, h / 2], [hd, h], [hd, h * 0.75], [w, h * 0.75]]);
    }
    case "line": return `M0 0 L${f(w)} ${f(h)}`;
    case "star5": {
      const pts = [];
      const cx = w / 2, cy = h / 2;
      for (let k = 0; k < 10; k++) {
        const a = -Math.PI / 2 + (k * Math.PI) / 5;
        const r = k % 2 ? 0.382 : 1;
        pts.push([cx + Math.cos(a) * r * (w / 2), cy + Math.sin(a) * r * (h / 2) * 1.05 + h * 0.025]);
      }
      return P(pts);
    }
    case "hexagon": {
      const o = ss * 0.25;
      return P([[o, 0], [w - o, 0], [w, h / 2], [w - o, h], [o, h], [0, h / 2]]);
    }
    case "chevron": {
      const o = Math.min(w, ss * 0.5);
      return P([[0, 0], [w - o, 0], [w, h / 2], [w - o, h], [0, h], [o, h / 2]]);
    }
    default: return `M0 0 H${f(w)} V${f(h)} H0 Z`;
  }
}

function geomStyle(it) {
  const st = [`left:${num(it.x)}px`, `top:${num(it.y)}px`, `width:${Math.max(0, num(it.w))}px`, `height:${Math.max(0, num(it.h))}px`];
  if (it.rot) st.push(`transform:rotate(${it.rot}deg)`);
  return st;
}

function textBoxHtml(it, deck, opts, color) {
  const theme = deck.theme || {};
  const font = it.ph === "title" ? theme.headFont || theme.font : theme.font;
  const valign = it.valign || (it.type === "shape" ? "middle" : it.ph === "title" ? "middle" : "top");
  const jc = valign === "middle" ? "center" : valign === "bottom" ? "flex-end" : "flex-start";
  const prompt = opts.edit && it.ph && !hasText(it.paras) ? PROMPTS[it.ph] || PROMPTS.body : "";
  const lh = it.ph === "title" ? 1.0 : 1.18;
  return `<div class="sl-tb" style="justify-content:${jc};font-family:${esc(fontStack(font))};color:${color};line-height:${lh}">`
    + `<div class="sl-tx">${parasHtml(it.paras, { item: it, deck, prompt })}</div></div>`;
}

/** HTML of one item (absolutely positioned in slide points). */
export function renderItemHtml(it, deck, index, opts = {}) {
  if (!it || typeof it !== "object") return "";
  const di = ` data-i="${index}"`;
  const st = geomStyle(it);
  if (it.type === "text") {
    if (it.fill) st.push(`background:${it.fill}`);
    if (it.line) st.push(`box-shadow:inset 0 0 0 ${it.lineW || 0.75}px ${it.line}`);
    const ph = it.ph ? ` data-ph="${it.ph}"` : "";
    // empty placeholders are invisible outside the editor (no stray bullets)
    const inner = !opts.edit && !hasText(it.paras) ? "" : textBoxHtml(it, deck, opts, textColorFor(it, deck, it.fill || null));
    return `<div class="sl-item sl-text"${di}${ph} style="${st.join(";")}">${inner}</div>`;
  }
  if (it.type === "shape") {
    const w = Math.max(0, num(it.w)), h = Math.max(0, num(it.h));
    const isLine = it.shape === "line";
    const stroke = it.line || (isLine ? (deck.theme && deck.theme.color) || "#262626" : null);
    const sw = stroke ? it.lineW || (isLine ? 1.5 : 0.75) : 0;
    const fill = !isLine && it.fill ? it.fill : "none";
    const svg = `<svg class="sl-svg" width="${Math.max(w, 1)}" height="${Math.max(h, 1)}" overflow="visible" aria-hidden="true">`
      + `<path d="${shapePath(it.shape, w, h)}" fill="${fill}" stroke="${stroke || "none"}" stroke-width="${sw}" stroke-linejoin="round" vector-effect="non-scaling-stroke"/></svg>`;
    const text = !isLine && (hasText(it.paras) || opts.edit) ? textBoxHtml({ ...it, paras: it.paras || [] }, deck, opts, textColorFor(it, deck)) : "";
    return `<div class="sl-item sl-shape" data-shape="${esc(it.shape)}"${di} style="${st.join(";")}">${svg}${text}</div>`;
  }
  if (it.type === "image") {
    const src = /^data:image\/[a-z0-9.+-]+;base64,/i.test(it.src || "") || /^(blob:|https?:)/i.test(it.src || "") ? it.src : "";
    const img = src ? `<img src="${esc(src)}" alt="${esc(it.alt || "")}" draggable="false">` : `<div class="sl-missing">${esc(it.alt || "Picture")}</div>`;
    return `<div class="sl-item sl-image"${di} style="${st.join(";")}">${img}</div>`;
  }
  if (it.type === "table") {
    const rows = it.rows || [];
    const ncol = rows.reduce((m, r) => Math.max(m, r.length), 0);
    const colW = Array.isArray(it.colW) && it.colW.length === ncol ? it.colW : null;
    const total = colW ? colW.reduce((a, b) => a + b, 0) || 1 : 1;
    const cols = colW ? `<colgroup>${colW.map((c) => `<col style="width:${(c / total) * 100}%">`).join("")}</colgroup>` : "";
    const theme = deck.theme || {};
    const body = rows.map((r, ri) => `<tr>${r.map((c, ci) => {
      const fill = c && c.fill;
      const cst = [];
      if (fill) cst.push(`background:${fill}`);
      cst.push(`color:${textColorFor(it, deck, fill || null)}`);
      return `<td data-r="${ri}" data-c="${ci}" style="${cst.join(";")}"><div class="sl-tx">${parasHtml((c && c.paras) || [], { item: it, deck })}</div></td>`;
    }).join("")}</tr>`).join("");
    return `<div class="sl-item sl-table"${di} style="${st.join(";")}"><table style="font-family:${esc(fontStack(theme.font))}">${cols}<tbody>${body}</tbody></table></div>`;
  }
  // unknown item type: keep it, show a neutral box where it sits
  if (typeof it.x === "number" && typeof it.w === "number") {
    return `<div class="sl-item sl-unknown"${di} style="${st.join(";")}"><span>${esc(it.type || "object")}</span></div>`;
  }
  return "";
}

/**
 * One slide as HTML: a w x h px box (1 px = 1 pt) to be scaled with a CSS transform.
 * opts.edit adds placeholder prompts ("Click to add title").
 */
export function renderSlideHtml(slide, deck, opts = {}) {
  const W = deck.size.w, H = deck.size.h;
  const bg = slideBg(slide, deck);
  const items = ((slide && slide.items) || []).map((it, i) => renderItemHtml(it, deck, i, opts)).join("");
  const cls = opts.className ? ` ${opts.className}` : "";
  return `<div class="sl-slide${cls}" style="width:${W}px;height:${H}px;background:${bg}">${items}</div>`;
}

/** CSS needed to show rendered slides (used inside the editor via slides.css and in exports). */
export const SLIDE_CSS = `
.sl-slide{position:relative;overflow:hidden;box-sizing:border-box;-webkit-font-smoothing:antialiased;text-rendering:optimizeLegibility}
.sl-slide *{box-sizing:border-box}
.sl-item{position:absolute;transform-origin:50% 50%}
.sl-tb{position:absolute;inset:0;padding:3.6px 7.2px;display:flex;flex-direction:column;overflow:visible;word-wrap:break-word;overflow-wrap:break-word;white-space:pre-wrap}
.sl-tx{outline:none;min-height:1em;flex:none}
.sl-tx p{margin:0;min-height:1em}
.sl-tx p[data-bullet]::before{content:"\\2022";display:inline-block;min-width:22px;padding-right:.15em;white-space:nowrap;text-indent:0;text-align:left;color:var(--bc,inherit)}
.sl-tx p[data-bullet="ol"]::before{content:attr(data-num)}
.sl-tx a{color:inherit;text-decoration:underline}
.sl-tx p[data-prompt]::after{content:attr(data-prompt);opacity:.55}
.sl-svg{position:absolute;left:0;top:0;overflow:visible}
.sl-image img{width:100%;height:100%;display:block;object-fit:fill;user-select:none}
.sl-missing{width:100%;height:100%;display:flex;align-items:center;justify-content:center;background:#e9ecef;color:#6a737d;font:14px sans-serif}
.sl-table table{width:100%;height:100%;border-collapse:collapse;table-layout:fixed;font-size:18px;line-height:1.15}
.sl-table td{border:1px solid #BFBFBF;padding:3.6px 7.2px;vertical-align:top;overflow:hidden;word-wrap:break-word;white-space:pre-wrap}
.sl-unknown{border:1px dashed #9aa;background:rgba(200,200,200,.25);display:flex;align-items:center;justify-content:center;font:12px sans-serif;color:#667}
`;

// ---------------------------------------------------------------------------
// exports: text / markdown / html slideshow
// ---------------------------------------------------------------------------

function nonTitleParas(slide) {
  const out = [];
  for (const it of slide.items || []) {
    if (!it || it.ph === "title") continue;
    if ((it.type === "text" || it.type === "shape") && Array.isArray(it.paras)) out.push(...it.paras.filter((p) => paraText(p).trim() !== ""));
  }
  return out;
}

export function slidesToText(deck) {
  const out = [];
  (deck.slides || []).forEach((s, i) => {
    const title = slideTitle(s);
    out.push(`Slide ${i + 1}${title ? ": " + title : ""}${s.hidden ? " (hidden)" : ""}`);
    for (const p of nonTitleParas(s)) out.push(`${"  ".repeat(p.level || 0)}${p.bullet ? "- " : ""}${paraText(p).replace(/\n/g, " ")}`);
    for (const it of s.items || []) {
      if (it && it.type === "table") for (const r of it.rows || []) out.push(r.map((c) => ((c && c.paras) || []).map(paraText).join(" ")).join("\t"));
    }
    if (s.notes && s.notes.trim()) out.push(`Notes: ${s.notes.trim()}`);
    out.push("");
  });
  return out.join("\n");
}

function mdEsc(s) {
  return String(s).replace(/([\\`*_[\]])/g, "\\$1").replace(/^(#|>|\d+\.|[-+] )/, "\\$1");
}

export function slidesToMarkdown(deck) {
  const out = [];
  const title = deck.meta && deck.meta.title;
  if (title) out.push(`# ${mdEsc(title)}`, "");
  (deck.slides || []).forEach((s, i) => {
    const t = slideTitle(s);
    out.push(`## ${i + 1}. ${t ? mdEsc(t) : "Untitled slide"}`, "");
    const paras = nonTitleParas(s);
    const counters = [];
    for (const p of paras) {
      const lvl = p.level || 0;
      let marker = "-";
      if (p.bullet === "ol") { counters.length = lvl + 1; counters[lvl] = (counters[lvl] || 0) + 1; marker = `${counters[lvl]}.`; } else counters.length = lvl;
      out.push(`${"  ".repeat(lvl)}${marker} ${mdEsc(paraText(p).replace(/\n/g, " "))}`);
    }
    if (paras.length) out.push("");
    for (const it of s.items || []) {
      if (it && it.type === "table" && (it.rows || []).length) {
        const rows = it.rows.map((r) => r.map((c) => mdEsc(((c && c.paras) || []).map(paraText).join(" ")).replace(/\|/g, "\\|")));
        const n = rows.reduce((m, r) => Math.max(m, r.length), 0);
        const line = (r) => `| ${Array.from({ length: n }, (_, k) => r[k] || "").join(" | ")} |`;
        out.push(line(rows[0]), `|${" --- |".repeat(n)}`, ...rows.slice(1).map(line), "");
      } else if (it && it.type === "image") {
        out.push(`*[Picture${it.alt ? ": " + mdEsc(it.alt) : ""}]*`, "");
      }
    }
    if (s.notes && s.notes.trim()) {
      out.push(...s.notes.trim().split("\n").map((l) => `> ${l}`), "");
    }
  });
  return out.join("\n").replace(/\n{3,}/g, "\n\n").trim() + "\n";
}

/** Standalone HTML slideshow: scaled slides, arrows/space/click navigation, notes hidden. */
export function slidesToHtmlDocument(deck, title = "Presentation") {
  const d = normalizeDeck(cloneDeep(deck));
  const W = d.size.w, H = d.size.h;
  const slides = d.slides.map((s, i) => `<section class="sv-x-slide" data-i="${i}" data-tr="${esc(s.transition || "none")}"${s.hidden ? " data-hidden=\"1\"" : ""}>${renderSlideHtml(s, d)}</section>`).join("\n");
  const css = `${SLIDE_CSS}
html,body{margin:0;height:100%;background:#111;overflow:hidden;font-family:"Segoe UI",Arial,sans-serif}
.sv-x-slide{position:absolute;left:0;top:0;width:${W}px;height:${H}px;transform-origin:0 0;display:none}
.sv-x-slide.on{display:block}
.sv-x-slide.tr-fade{animation:svfade .5s ease}
.sv-x-slide.tr-push{animation:svpush .5s ease}
.sv-x-slide.tr-wipe{animation:svwipe .5s ease}
.sv-x-slide.tr-zoom{animation:svzoom .5s ease}
@keyframes svfade{from{opacity:0}to{opacity:1}}
@keyframes svpush{from{translate:0 100%}to{translate:0 0}}
@keyframes svwipe{from{clip-path:inset(0 100% 0 0)}to{clip-path:inset(0 0 0 0)}}
@keyframes svzoom{from{opacity:0;scale:.6}to{opacity:1;scale:1}}
#count{position:fixed;right:14px;bottom:10px;color:#aaa;font-size:13px;opacity:.7}
@media print{html,body{background:#fff;overflow:visible;height:auto}#count{display:none}
.sv-x-slide{display:block!important;position:relative;transform:none!important;break-after:page;page-break-after:always;margin:0 auto}
.sv-x-slide[data-hidden]{display:none!important}@page{size:landscape;margin:.3in}}`;
  const js = `(function(){var W=${W},H=${H};var all=[].slice.call(document.querySelectorAll('.sv-x-slide'));
var list=all.filter(function(s){return !s.hasAttribute('data-hidden')});if(!list.length)list=all;var cur=0;var cnt=document.getElementById('count');
function fit(){var k=Math.min(innerWidth/W,innerHeight/H);var x=(innerWidth-W*k)/2,y=(innerHeight-H*k)/2;
all.forEach(function(s){s.style.transform='translate('+x+'px,'+y+'px) scale('+k+')'})}
function show(n,anim){n=Math.max(0,Math.min(list.length-1,n));list.forEach(function(s){s.classList.remove('on','tr-fade','tr-push','tr-wipe','tr-zoom')});
var s=list[n];s.classList.add('on');var t=s.getAttribute('data-tr');if(anim&&t&&t!=='none'){void s.offsetWidth;s.classList.add('tr-'+t)}cur=n;cnt.textContent=(n+1)+' / '+list.length}
addEventListener('resize',fit);addEventListener('keydown',function(e){var k=e.key;
if(k==='ArrowRight'||k==='ArrowDown'||k===' '||k==='PageDown'||k==='Enter'||k==='n'){e.preventDefault();if(cur<list.length-1)show(cur+1,true)}
else if(k==='ArrowLeft'||k==='ArrowUp'||k==='PageUp'||k==='Backspace'||k==='p'){e.preventDefault();show(cur-1,false)}
else if(k==='Home'){show(0,false)}else if(k==='End'){show(list.length-1,false)}
else if(k==='f'||k==='F5'){e.preventDefault();if(document.documentElement.requestFullscreen)document.documentElement.requestFullscreen().catch(function(){})}});
addEventListener('click',function(e){if(e.target.closest&&e.target.closest('a'))return;if(cur<list.length-1)show(cur+1,true)});
addEventListener('contextmenu',function(e){e.preventDefault();show(cur-1,false)});fit();show(0,false)})();`;
  return `<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>${esc(title)}</title>
<meta name="generator" content="Office Editor">
<style>${css}</style></head>
<body>
${slides}
<div id="count"></div>
<script>${js}</script>
</body></html>
`;
}
