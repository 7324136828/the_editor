/**
 * Word document model: JSON <-> HTML rendering and DOM -> model parsing.
 *
 * Model shape (frozen contract shared with the Rust backend, see
 * src/office-backend/src/model.rs "WORD CONTRACT"):
 * {
 *   kind: "word",
 *   meta?: { title?, author?, subject?: string, warnings?: string[],
 *            page?: {size, orientation, width, height, margins} (twips),
 *            header?/footer?: { runs: Run[], align? } — "{PAGE}" / "{NUMPAGES}"
 *            inside a run's text are page-number fields },
 *   blocks: Block[]
 * }
 *
 * Block =
 *   { type:"p"|"h1".."h6"|"quote"|"code", align?:"left"|"center"|"right"|"justify", runs: Run[] }
 * | { type:"list", list:"ul"|"ol", level:0..8, runs: Run[] }      // one block per list item
 * | { type:"table", rows: Cell[][] }            // Cell = { blocks: Block[], colspan?: number }
 * | { type:"img", src:"data:...", alt?:string, w?:number, h?:number }   // w/h in CSS px
 * | { type:"pagebreak" }
 *
 * Run = { t, b?, i?, u?, s?, sup?, sub?, color?:"#RRGGBB", size?:pt, font?, link?, hl?:"#RRGGBB" }
 *   - t may contain "\n" (soft line break, rendered as <br>) and "\t".
 *
 * Design notes
 * ------------
 * renderWordHtml() produces HTML that is *parse-stable*: parseWordDom() of the
 * rendered DOM gives back the same model (modulo normalizeWordModel()).
 * Rendering conventions the parser relies on:
 *   - every run is a <span> carrying its marks as inline CSS (b/i/u/s/color/
 *     background-color/size/font); sup/sub use <sup>/<sub> tags;
 *   - links are <a class="ow-link"> with NO inline presentation styles (the
 *     link look comes from CSS), so nothing display-only leaks into the model;
 *   - real quotes are <blockquote data-ow="quote">; a <blockquote> carrying an
 *     inline margin style is Chrome's execCommand('indent') wrapper and is
 *     parsed as a transparent container (its paragraphs stay paragraphs);
 *   - alignment is ONLY inline `text-align` (no stale data-align attribute);
 *   - runs of spaces / leading / trailing spaces are emitted as &nbsp; so the
 *     browser displays them, and the parser maps &nbsp; back to spaces;
 *   - a paragraph whose text ends with "\n" gets one extra <br> (the browser
 *     does not display a single trailing <br>); the parser strips one trailing
 *     line break per block accordingly.
 *
 * parseWordDom() reads whatever DOM Chrome's contenteditable/execCommand turns
 * that into: top-level text nodes, <div>/<br> line structure, <p><ol>…</ol></p>
 * (list toggled on a paragraph), <ul><li>a</li><ul><li>b</li></ul></ul>
 * (indent), bare text in <li>, <font>, styled spans, tables with bare text in
 * cells, inline <img> (split out into img blocks), etc.
 */

export const BLOCK_TAGS = {
  p: "p",
  h1: "h1", h2: "h2", h3: "h3", h4: "h4", h5: "h5", h6: "h6",
  quote: "blockquote",
  code: "pre",
};

const MAX_LIST_LEVEL = 8;
const PARA_TYPES = new Set(["p", "h1", "h2", "h3", "h4", "h5", "h6", "quote", "code"]);

export function escapeHtml(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

// ---------------------------------------------------------------------------
// Model -> HTML
// ---------------------------------------------------------------------------

/** Font family names go into a single-quoted CSS string; strip anything that could break out. */
function cssFontName(name) {
  return String(name).replace(/['"\\;<>{}]/g, "").trim();
}

function cssColor(c) {
  const hex = colorToHex(c);
  return hex || null;
}

function runStyle(run) {
  const st = [];
  if (run.b) st.push("font-weight:bold");
  if (run.i) st.push("font-style:italic");
  if (run.u && run.s) st.push("text-decoration:underline line-through");
  else if (run.u) st.push("text-decoration:underline");
  else if (run.s) st.push("text-decoration:line-through");
  const color = run.color && cssColor(run.color);
  if (color) st.push(`color:${color}`);
  const hl = run.hl && cssColor(run.hl);
  if (hl) st.push(`background-color:${hl}`);
  const size = Number(run.size);
  if (size > 0 && size < 1639) st.push(`font-size:${size}pt`);
  const font = run.font && cssFontName(run.font);
  if (font) st.push(`font-family:'${font}'`);
  return st.length ? ` style="${escapeHtml(st.join(";"))}"` : "";
}

/** Schemes we are willing to put into a live href. Everything else is kept in data-href. */
function isSafeHref(url) {
  const u = String(url).trim();
  const m = /^([a-z][a-z0-9+.-]*):/i.exec(u);
  if (!m) return true; // relative URL / #anchor
  return /^(https?|mailto|tel|ftp|ftps|news|file)$/i.test(m[1]);
}

function linkOpenTag(url) {
  if (isSafeHref(url)) {
    return `<a class="ow-link" href="${escapeHtml(url)}" target="_blank" rel="noopener noreferrer">`;
  }
  return `<a class="ow-link" href="#" data-href="${escapeHtml(url)}">`;
}

/**
 * Escape one run's text for HTML. `all` is the whole block text and `start`
 * this run's offset in it, so spacing decisions can look across run borders.
 */
function textHtml(t, all, start, pre) {
  let out = "";
  for (let k = 0; k < t.length; k++) {
    const ch = t[k];
    const gi = start + k;
    switch (ch) {
      case "\n": out += "<br>"; break;
      case "\r": break;
      case "\t":
        out += pre ? "\t" : '<span class="ow-tab" style="white-space:pre">\t</span>';
        break;
      case " ": {
        if (pre) { out += " "; break; }
        const prev = gi === 0 ? "\n" : all[gi - 1];
        const next = gi + 1 >= all.length ? "\n" : all[gi + 1];
        // A space the browser would collapse or hide must be non-breaking.
        out += prev === " " || prev === "\n" || next === "\n" ? "&nbsp;" : " ";
        break;
      }
      case "&": out += "&amp;"; break;
      case "<": out += "&lt;"; break;
      case ">": out += "&gt;"; break;
      case '"': out += "&quot;"; break;
      default: out += ch;
    }
  }
  return out;
}

function renderRunsHtml(runs, pre = false) {
  const list = (runs || []).filter((r) => r && r.t != null && String(r.t) !== "");
  if (!list.length) return "<br>";
  const all = list.map((r) => String(r.t)).join("");
  let html = "";
  let pos = 0;
  let i = 0;
  while (i < list.length) {
    const link = list[i].link || null;
    let inner = "";
    let j = i;
    // Consecutive runs with the same link share one <a>.
    while (j < list.length && (list[j].link || null) === link) {
      const r = list[j];
      const t = String(r.t);
      let body = textHtml(t, all, pos, pre);
      if (r.sub) body = `<sub>${body}</sub>`;
      if (r.sup) body = `<sup>${body}</sup>`;
      inner += `<span${runStyle(r)}>${body}</span>`;
      pos += t.length;
      j++;
    }
    html += link ? `${linkOpenTag(link)}${inner}</a>` : inner;
    i = j;
  }
  // The browser does not display a lone trailing <br>; add the placeholder.
  if (all.endsWith("\n")) html += "<br>";
  return html;
}

function alignStyle(align) {
  return align === "center" || align === "right" || align === "justify"
    ? ` style="text-align:${align}"`
    : "";
}

function renderParagraph(b) {
  const align = alignStyle(b.align);
  switch (b.type) {
    case "quote":
      return `<blockquote data-ow="quote"${align}>${renderRunsHtml(b.runs)}</blockquote>`;
    case "code":
      return `<pre${align}>${renderRunsHtml(b.runs, true)}</pre>`;
    default: {
      const tag = BLOCK_TAGS[b.type] || "p";
      return `<${tag}${align}>${renderRunsHtml(b.runs)}</${tag}>`;
    }
  }
}

function renderImage(b) {
  const w = Number(b.w) > 0 ? ` width="${Math.round(Number(b.w))}"` : "";
  const h = Number(b.h) > 0 ? ` height="${Math.round(Number(b.h))}"` : "";
  return `<img class="ow-img" src="${escapeHtml(b.src || "")}" alt="${escapeHtml(b.alt || "")}"${w}${h}>`;
}

function renderTable(b) {
  let html = '<table class="ow-table"><tbody>';
  for (const row of b.rows || []) {
    html += "<tr>";
    for (const cell of row || []) {
      const cs = cell && cell.colspan > 1 ? ` colspan="${cell.colspan | 0}"` : "";
      const inner = renderBlocks((cell && cell.blocks) || []);
      html += `<td${cs}>${inner || "<p><br></p>"}</td>`;
    }
    html += "</tr>";
  }
  html += "</tbody></table>";
  return html;
}

function renderBlocks(blocks) {
  let html = "";
  const stack = []; // open lists: {type, level}

  const closeTop = () => {
    const top = stack.pop();
    html += `</li></${top.type}>`;
  };

  for (const b of blocks || []) {
    if (!b) continue;
    if (b.type === "list") {
      const lvl = Math.min(MAX_LIST_LEVEL, Math.max(0, b.level | 0));
      const type = b.list === "ol" ? "ol" : "ul";
      // Close deeper levels, or the same level when the list type changes.
      while (stack.length) {
        const top = stack[stack.length - 1];
        if (top.level > lvl || (top.level === lvl && top.type !== type)) closeTop();
        else break;
      }
      const top = stack[stack.length - 1];
      if (!top || top.level < lvl) {
        // Open the missing levels; skipped intermediate levels get an item
        // without own content (the parser drops such content-less items).
        const start = top ? top.level + 1 : 0;
        for (let l = start; l <= lvl; l++) {
          html += `<${type}><li>`;
          stack.push({ type, level: l });
        }
      } else {
        html += "</li><li>";
      }
      html += renderRunsHtml(b.runs);
      continue;
    }

    while (stack.length) closeTop();

    if (PARA_TYPES.has(b.type)) {
      html += renderParagraph(b);
    } else if (b.type === "table") {
      html += renderTable(b);
    } else if (b.type === "img") {
      if (b.src) html += renderImage(b);
    } else if (b.type === "pagebreak") {
      html += '<hr class="ow-pagebreak" data-ow="pagebreak" contenteditable="false">';
    } else if (Array.isArray(b.runs)) {
      // Unknown block type with text: show it as a paragraph rather than drop it.
      html += `<p>${renderRunsHtml(b.runs)}</p>`;
    }
  }
  while (stack.length) closeTop();
  return html;
}

/** Render the whole word model to an HTML string for the article element. */
export function renderWordHtml(model) {
  // An empty document still needs one paragraph to put the caret in;
  // otherwise typed text would land directly in the editor root.
  return renderBlocks((model && model.blocks) || []) || "<p><br></p>";
}

// ---------------------------------------------------------------------------
// Colors and sizes
// ---------------------------------------------------------------------------

const NAMED_COLORS = {
  black: "#000000", white: "#FFFFFF", red: "#FF0000", lime: "#00FF00", green: "#008000",
  blue: "#0000FF", yellow: "#FFFF00", cyan: "#00FFFF", aqua: "#00FFFF", magenta: "#FF00FF",
  fuchsia: "#FF00FF", gray: "#808080", grey: "#808080", silver: "#C0C0C0", maroon: "#800000",
  olive: "#808000", purple: "#800080", teal: "#008080", navy: "#000080", orange: "#FFA500",
  pink: "#FFC0CB", brown: "#A52A2A", gold: "#FFD700", darkblue: "#00008B", darkred: "#8B0000",
  darkgreen: "#006400", darkgray: "#A9A9A9", darkgrey: "#A9A9A9", lightgray: "#D3D3D3",
  lightgrey: "#D3D3D3", darkcyan: "#008B8B", darkmagenta: "#8B008B", indigo: "#4B0082",
  violet: "#EE82EE", orangered: "#FF4500", crimson: "#DC143C",
};

/** Any CSS-ish color -> "#RRGGBB" (uppercase), or null for transparent/unknown. */
export function colorToHex(c) {
  if (c == null) return null;
  const s = String(c).trim().toLowerCase();
  if (!s) return null;
  let m = /^#?([0-9a-f]{6})([0-9a-f]{2})?$/.exec(s);
  if (m && (s.startsWith("#") || /^[0-9a-f]{6}$/.test(s))) {
    if (m[2] === "00") return null;
    return `#${m[1].toUpperCase()}`;
  }
  m = /^#([0-9a-f])([0-9a-f])([0-9a-f])([0-9a-f])?$/.exec(s);
  if (m) {
    if (m[4] === "0") return null;
    return `#${m[1]}${m[1]}${m[2]}${m[2]}${m[3]}${m[3]}`.toUpperCase();
  }
  m = /^rgba?\(\s*([\d.]+)%?[\s,]+([\d.]+)%?[\s,]+([\d.]+)%?(?:[\s,/]+([\d.]+)(%?))?\s*\)$/.exec(s);
  if (m) {
    if (m[4] !== undefined) {
      const a = parseFloat(m[4]) / (m[5] ? 100 : 1);
      if (a === 0) return null;
    }
    const h = (n) => Math.max(0, Math.min(255, Math.round(parseFloat(n)))).toString(16).padStart(2, "0");
    return `#${h(m[1])}${h(m[2])}${h(m[3])}`.toUpperCase();
  }
  return NAMED_COLORS[s] || null;
}

const SIZE_KEYWORDS_PX = {
  "xx-small": 9, "x-small": 10, small: 13, medium: 16, large: 18,
  "x-large": 24, "xx-large": 32, "xxx-large": 48,
};
// <font size="1..7"> in px (same scale Chrome uses).
const FONT_TAG_SIZE_PX = { 1: 10, 2: 13, 3: 16, 4: 18, 5: 24, 6: 32, 7: 48 };

/** CSS length/keyword -> points (unrounded), or null for relative units. */
function cssSizeToPt(v) {
  const s = String(v || "").trim().toLowerCase();
  if (!s) return null;
  if (SIZE_KEYWORDS_PX[s]) return SIZE_KEYWORDS_PX[s] * 0.75;
  const m = /^([\d.]+)(pt|px|pc|in|cm|mm)$/.exec(s);
  if (!m) return null;
  const n = parseFloat(m[1]);
  const factor = { pt: 1, px: 0.75, pc: 12, in: 72, cm: 72 / 2.54, mm: 72 / 25.4 }[m[2]];
  const pt = n * factor;
  return pt > 0 && pt < 1639 ? pt : null;
}

/** Word stores sizes in half points; snap to that grid. */
function roundPt(pt) {
  return Math.round(pt * 2) / 2;
}

const GENERIC_FONTS = new Set([
  "serif", "sans-serif", "monospace", "cursive", "fantasy", "system-ui", "ui-serif",
  "ui-sans-serif", "ui-monospace", "inherit", "initial", "unset", "revert", "-webkit-body",
]);

function firstFamily(ff) {
  const first = String(ff || "").split(",")[0].trim().replace(/^['"]|['"]$/g, "").trim();
  if (!first || GENERIC_FONTS.has(first.toLowerCase())) return null;
  return first;
}

// ---------------------------------------------------------------------------
// DOM -> model
// ---------------------------------------------------------------------------

const SKIP_TAGS = new Set([
  "SCRIPT", "STYLE", "TEMPLATE", "META", "LINK", "TITLE", "NOSCRIPT", "HEAD", "OBJECT",
  "EMBED", "IFRAME", "SVG", "MATH", "CANVAS", "VIDEO", "AUDIO", "SELECT", "BUTTON", "INPUT",
  "TEXTAREA", "COLGROUP", "COL",
]);
const BLOCK_SET = new Set([
  "P", "H1", "H2", "H3", "H4", "H5", "H6", "BLOCKQUOTE", "PRE", "UL", "OL", "LI", "TABLE",
  "DIV", "SECTION", "ARTICLE", "MAIN", "HEADER", "FOOTER", "NAV", "ASIDE", "FIGURE",
  "FIGCAPTION", "ADDRESS", "DL", "DT", "DD", "CENTER", "FORM", "FIELDSET", "DETAILS",
  "SUMMARY", "BODY", "HTML", "TR", "TD", "TH", "TBODY", "THEAD", "TFOOT", "CAPTION", "LEGEND",
]);
const MARK_KEYS = ["b", "i", "u", "s", "sup", "sub", "color", "size", "font", "link", "hl"];

function tagOf(el) {
  return String(el.tagName || "").toUpperCase();
}

function styleAttr(el) {
  return (el.getAttribute && el.getAttribute("style")) || "";
}

function isHiddenEl(el) {
  if (el.hasAttribute && el.hasAttribute("hidden")) return true;
  const st = el.style;
  return !!(st && (st.display === "none" || st.visibility === "hidden"));
}

function isAllWs(s) {
  return !/[^ \t\n\r\f\u200B\uFEFF]/.test(s);
}

/** Marks implied by the element's tag (and <font> attributes). */
function marksFromTag(el, tag, marks) {
  const m = { ...marks };
  switch (tag) {
    case "B": case "STRONG": m.b = true; break;
    case "I": case "EM": case "CITE": case "VAR": case "DFN": m.i = true; break;
    case "U": case "INS": m.u = true; break;
    case "S": case "STRIKE": case "DEL": m.s = true; break;
    case "SUP": m.sup = true; delete m.sub; break;
    case "SUB": m.sub = true; delete m.sup; break;
    case "CODE": case "KBD": case "SAMP": case "TT": if (!m.font) m.font = "Consolas"; break;
    case "MARK": if (!m.hl) m.hl = "#FFFF00"; break;
    case "TH": m.b = true; break;
    case "FONT": {
      const c = colorToHex(el.getAttribute("color"));
      if (c) m.color = c;
      const f = firstFamily(el.getAttribute("face"));
      if (f) m.font = f;
      const sz = (el.getAttribute("size") || "").trim();
      if (FONT_TAG_SIZE_PX[sz]) m.size = roundPt(FONT_TAG_SIZE_PX[sz] * 0.75);
      break;
    }
    default: break;
  }
  return m;
}

/**
 * Marks from an element's inline style. `inline` = the element is an inline
 * element (background colour only counts there). `isLink` = an <a>: its
 * underline and the legacy viewer link colour are presentation, not content.
 */
function marksFromStyle(el, marks, defaults, inline, isLink) {
  if (!styleAttr(el)) return marks;
  const st = el.style;
  if (!st) return marks;
  const m = { ...marks };

  const deco = String(st.textDecorationLine || st.textDecoration || "").toLowerCase();
  if (deco) {
    if (/underline/.test(deco) && !isLink) m.u = true;
    if (/line-through/.test(deco)) m.s = true;
  }

  const fw = String(st.fontWeight || "").toLowerCase();
  if (fw) {
    const n = parseInt(fw, 10);
    if (fw === "bold" || fw === "bolder" || n >= 600) m.b = true;
    else if (fw === "normal" || fw === "lighter" || (n > 0 && n < 600)) delete m.b;
  }

  const fst = String(st.fontStyle || "").toLowerCase();
  if (fst === "italic" || fst === "oblique") m.i = true;
  else if (fst === "normal") delete m.i;

  if (st.color) {
    const hex = colorToHex(st.color);
    const legacyLink = isLink && hex === "#1A5FB4";
    // Chrome copies the editor's own text colour into spans when merging
    // blocks; that colour is not a user choice. Explicit black etc. is kept.
    if (hex && !legacyLink && hex !== defaults.color) m.color = hex;
  }

  if (inline && st.backgroundColor) {
    const hex = colorToHex(st.backgroundColor);
    if (hex) m.hl = hex;
  }

  if (st.fontSize) {
    const pt = cssSizeToPt(st.fontSize);
    if (pt && !(defaults.sizePt && Math.abs(pt - defaults.sizePt) < 0.01)) m.size = roundPt(pt);
  }

  if (st.fontFamily) {
    const f = firstFamily(st.fontFamily);
    if (f) m.font = f;
  }

  const va = String(st.verticalAlign || "").toLowerCase();
  if (va === "super") { m.sup = true; delete m.sub; }
  else if (va === "sub") { m.sub = true; delete m.sup; }
  else if (va === "baseline") { delete m.sup; delete m.sub; }
  return m;
}

/** null = not specified; "" = explicitly left/start; else center|right|justify. */
function alignValue(v) {
  const s = String(v || "").trim().toLowerCase();
  if (!s) return null;
  if (s === "center" || s === "-webkit-center") return "center";
  if (s === "right" || s === "end" || s === "-webkit-right") return "right";
  if (s === "justify") return "justify";
  if (s === "left" || s === "start" || s === "-webkit-left") return "";
  return null;
}

/** Inline style wins over the legacy align attribute, which wins over old data-align. */
function alignOf(el) {
  let a = alignValue(el.style && el.style.textAlign);
  if (a !== null) return a;
  a = alignValue(el.getAttribute && el.getAttribute("align"));
  if (a !== null) return a;
  a = alignValue(el.getAttribute && el.getAttribute("data-align"));
  return a;
}

function blockCtx(el, ctx) {
  const tag = tagOf(el);
  let marks = marksFromStyle(el, ctx.marks, ctx.defaults, false, false);
  if (tag === "TH") marks = { ...marks, b: true };
  let align = ctx.align;
  const a = alignOf(el);
  if (a !== null) align = a || undefined;
  let pre = ctx.pre;
  const ws = el.style && String(el.style.whiteSpace || "").toLowerCase();
  if (tag === "PRE") pre = true;
  else if (ws) pre = ws.startsWith("pre") || ws === "break-spaces";
  return { ...ctx, marks, align, pre };
}

function inlineCtx(el, tag, ctx) {
  const isLink = tag === "A";
  let marks = marksFromTag(el, tag, ctx.marks);
  marks = marksFromStyle(el, marks, ctx.defaults, true, isLink);
  if (isLink) {
    const href = (el.getAttribute("data-href") || el.getAttribute("href") || "").trim();
    if (href && href !== "#") marks.link = href;
  }
  let pre = ctx.pre;
  const ws = el.style && String(el.style.whiteSpace || "").toLowerCase();
  if (ws) pre = ws.startsWith("pre") || ws === "break-spaces";
  return { ...ctx, marks, pre };
}

function sameMarks(a, b) {
  for (const k of MARK_KEYS) if ((a[k] || undefined) !== (b[k] || undefined)) return false;
  return true;
}

function cleanMarks(marks) {
  const out = {};
  for (const k of MARK_KEYS) if (marks[k]) out[k] = marks[k];
  return out;
}

function pushRun(runs, text, marks) {
  if (!text) return;
  const last = runs[runs.length - 1];
  const clean = cleanMarks(marks);
  if (last && sameMarks(last, clean)) last.t += text;
  else runs.push({ t: text, ...clean });
}

function lastChar(runs) {
  for (let i = runs.length - 1; i >= 0; i--) {
    const t = runs[i].t;
    if (t) return t[t.length - 1];
  }
  return undefined;
}

/** Append a DOM text node's content, applying HTML white-space collapsing unless in pre. */
function pushText(runs, text, ctx) {
  let t = String(text || "").replace(/[\u200B\uFEFF]/g, "");
  if (!ctx.pre) {
    t = t.replace(/[ \t\n\r\f]+/g, " ");
    if (t[0] === " ") {
      const prev = lastChar(runs);
      if (prev === undefined || prev === " " || prev === "\n") t = t.slice(1);
    }
  } else {
    t = t.replace(/\r\n?/g, "\n");
  }
  if (t) pushRun(runs, t, ctx.marks);
}

/**
 * Final clean-up of a block's runs: drop collapsible edge spaces (unless
 * pre), strip one trailing line break (the contenteditable placeholder <br>),
 * turn &nbsp; back into spaces and merge equal neighbours.
 */
function finishRuns(runs, pre) {
  const rs = runs.map((r) => ({ ...r }));
  const trimEnd = () => {
    while (rs.length) {
      const last = rs[rs.length - 1];
      last.t = last.t.replace(/ +$/, "");
      if (last.t) break;
      rs.pop();
    }
  };
  if (!pre) trimEnd();
  if (rs.length) {
    const last = rs[rs.length - 1];
    if (last.t.endsWith("\n")) {
      last.t = last.t.slice(0, -1);
      if (!last.t) rs.pop();
    }
  }
  if (!pre) {
    trimEnd();
    while (rs.length) {
      rs[0].t = rs[0].t.replace(/^ +/, "");
      if (rs[0].t) break;
      rs.shift();
    }
  }
  const out = [];
  for (const r of rs) {
    const t = r.t.replace(/\u00A0/g, " ");
    if (t) pushRun(out, t, r);
  }
  return out;
}

function imgBlockOf(el) {
  const src = (el.getAttribute("src") || "").trim();
  if (!src) return null;
  const b = { type: "img", src };
  const alt = el.getAttribute("alt");
  if (alt) b.alt = alt;
  const dim = (attr, css) => {
    const v = parseInt(el.getAttribute(attr) || "", 10);
    if (v > 0) return v;
    const c = el.style && el.style[css];
    const m = /^([\d.]+)px$/.exec(String(c || "").trim());
    return m ? Math.round(parseFloat(m[1])) : 0;
  };
  const w = dim("width", "width");
  const h = dim("height", "height");
  if (w > 0) b.w = w;
  if (h > 0) b.h = h;
  return b;
}

function isPagebreakHr(el) {
  const cls = " " + String(el.getAttribute("class") || "") + " ";
  return cls.includes(" ow-pagebreak ") || el.getAttribute("data-ow") === "pagebreak" || breaksAfter(el) || breaksBefore(el);
}

function breaksBefore(el) {
  const st = el.style;
  if (!st || !styleAttr(el)) return false;
  return /^(always|page|left|right)$/.test(st.pageBreakBefore || "") || /^(page|left|right)$/.test(st.breakBefore || "");
}

function breaksAfter(el) {
  const st = el.style;
  if (!st || !styleAttr(el)) return false;
  return /^(always|page|left|right)$/.test(st.pageBreakAfter || "") || /^(page|left|right)$/.test(st.breakAfter || "");
}

/** A real quote (ours or pasted) vs. Chrome's indent wrapper (inline margin style). */
function isQuoteBlock(el) {
  if (el.getAttribute("data-ow") === "quote") return true;
  if (/(^|\s)ow-quote(\s|$)/.test(el.getAttribute("class") || "")) return true;
  return !/margin/i.test(styleAttr(el));
}

// --- accumulators -----------------------------------------------------------
// The walker visits nodes in document order and reports text, <br>, <img>,
// <hr> and block elements to an accumulator; the accumulator decides how they
// map to model blocks for its kind of container.

function emitPara(out, ctx, runs) {
  const type = PARA_TYPES.has(ctx.paraType) ? ctx.paraType : "p";
  const block = { type, runs: finishRuns(runs, ctx.pre || type === "code") };
  if (ctx.align) block.align = ctx.align;
  out.push(block);
}

/**
 * Flow container (document root, <div>, table cell, indent wrapper…):
 * inline content forms anonymous paragraphs; a <br> ends the current line,
 * and a <br> on an empty line is an empty paragraph.
 */
class FlowAcc {
  /** `base` = the container's context: anonymous paragraphs take type/align/pre from it. */
  constructor(out, base = null) {
    this.out = out;
    this.base = base;
    this.runs = [];
    this.pctx = null;
  }
  text(s, ctx) {
    if (!this.pctx && !ctx.pre && isAllWs(s)) return;
    if (!this.pctx) this.pctx = this.base || ctx;
    pushText(this.runs, s, ctx);
  }
  br(ctx) {
    if (this.pctx) this.flush();
    else emitPara(this.out, ctx, []);
  }
  img(el) {
    const b = imgBlockOf(el);
    if (!b) return;
    this.flush();
    this.out.push(b);
  }
  hr(el) {
    this.flush();
    if (isPagebreakHr(el)) this.out.push({ type: "pagebreak" });
  }
  block(el, ctx) {
    this.flush();
    parseBlockEl(el, ctx, this.out);
  }
  flush() {
    if (this.pctx) emitPara(this.out, this.pctx, this.runs);
    this.runs = [];
    this.pctx = null;
  }
  finish() {
    this.flush();
  }
}

/**
 * Paragraph element (<p>, <hN>, <pre>, quote): inline content becomes one
 * block of ctx.paraType, <br> is a soft line break, nested blocks (a list
 * toggled inside the paragraph, an image…) split it.
 */
class ParaAcc {
  constructor(out, ctx) {
    this.out = out;
    this.ctx = ctx;
    this.runs = [];
    this.raw = false; // any content (incl. placeholder <br>) since last flush
    this.emitted = false;
    this.hadBlock = false;
  }
  text(s, ctx) {
    if (!this.raw && !ctx.pre && isAllWs(s)) return;
    this.raw = true;
    pushText(this.runs, s, ctx);
  }
  br(ctx) {
    this.raw = true;
    pushRun(this.runs, "\n", ctx.marks);
  }
  img(el) {
    const b = imgBlockOf(el);
    if (!b) return;
    this.flush();
    this.hadBlock = true;
    this.out.push(b);
  }
  hr(el) {
    this.flush();
    this.hadBlock = true;
    if (isPagebreakHr(el)) this.out.push({ type: "pagebreak" });
  }
  block(el, ctx) {
    this.flush();
    this.hadBlock = true;
    parseBlockEl(el, ctx, this.out);
  }
  flush() {
    if (this.raw) {
      emitPara(this.out, this.ctx, this.runs);
      this.emitted = true;
    }
    this.runs = [];
    this.raw = false;
  }
  finish() {
    if (this.raw) this.flush();
    else if (!this.emitted && !this.hadBlock) emitPara(this.out, this.ctx, []);
  }
}

/**
 * List item (<li>, or stray content directly inside <ul>/<ol>). Paragraph-like
 * children (<p>, <div>…) contribute to the item text (joined by line breaks);
 * nested lists become deeper levels; images are emitted after the item.
 */
class ItemAcc {
  constructor(out, list, level, stray = false) {
    this.out = out;
    this.list = list;
    this.level = level;
    this.stray = stray;
    this.runs = [];
    this.raw = false;
    this.needBreak = false;
    this.emitted = false;
    this.hadNested = false;
    this.deferred = [];
  }
  hasText() {
    return this.runs.some((r) => r.t && r.t.replace(/[\s\u00A0]/g, ""));
  }
  text(s, ctx) {
    if (!ctx.pre && isAllWs(s) && (!this.raw || this.needBreak)) return;
    if (this.needBreak && this.hasText()) pushRun(this.runs, "\n", ctx.marks);
    this.needBreak = false;
    this.raw = true;
    pushText(this.runs, s, ctx);
  }
  br(ctx) {
    this.raw = true;
    this.needBreak = false;
    pushRun(this.runs, "\n", ctx.marks);
  }
  img(el) {
    const b = imgBlockOf(el);
    if (b) this.deferred.push(b);
  }
  hr(el) {
    this.flushItem();
    if (isPagebreakHr(el)) this.out.push({ type: "pagebreak" });
  }
  block(el, ctx) {
    const tag = tagOf(el);
    if (tag === "UL" || tag === "OL") {
      this.flushItem();
      this.hadNested = true;
      parseList(el, Math.min(MAX_LIST_LEVEL, this.level + 1), ctx, this.out);
    } else if (tag === "TABLE") {
      this.flushItem();
      this.hadNested = true;
      parseTable(el, ctx, this.out);
    } else if (tag === "LI") {
      this.flushItem();
      this.hadNested = true;
      parseItem(el, this.list, this.level, ctx, this.out);
    } else {
      // Paragraph-like wrapper inside the item: its text is item text.
      if (this.hasText()) this.needBreak = true;
      walkChildren(el, blockCtx(el, ctx), this);
      if (this.hasText()) this.needBreak = true;
    }
  }
  emitDeferred() {
    for (const b of this.deferred) this.out.push(b);
    this.deferred = [];
  }
  flushItem() {
    if (this.raw && (!this.stray || this.hasText())) {
      this.out.push({ type: "list", list: this.list, level: this.level, runs: finishRuns(this.runs, false) });
      this.emitted = true;
    }
    this.runs = [];
    this.raw = false;
    this.needBreak = false;
    this.emitDeferred();
  }
  finish() {
    if (this.raw) {
      this.flushItem();
    } else {
      // An empty <li> is an empty bullet the user created; an <li> that only
      // wraps a nested list or an image has no bullet of its own.
      if (!this.stray && !this.emitted && !this.hadNested && !this.deferred.length) {
        this.out.push({ type: "list", list: this.list, level: this.level, runs: [] });
      }
      this.emitDeferred();
    }
  }
}

/** Inline-only accumulator used by the legacy parseInlines() export. */
class InlineAcc {
  constructor(runs) { this.runs = runs; }
  text(s, ctx) { pushText(this.runs, s, ctx); }
  br(ctx) { pushRun(this.runs, "\n", ctx.marks); }
  img(el, ctx) { pushRun(this.runs, el.getAttribute("alt") || "", ctx.marks); }
  hr() {}
  block(el, ctx) { walkChildren(el, blockCtx(el, ctx), this); }
}

// --- walker -----------------------------------------------------------------

function walkNode(node, ctx, acc) {
  if (node.nodeType === 3) {
    acc.text(node.nodeValue || "", ctx);
    return;
  }
  if (node.nodeType !== 1) return;
  const tag = tagOf(node);
  if (SKIP_TAGS.has(tag) || isHiddenEl(node)) return;
  switch (tag) {
    case "BR": acc.br(ctx); return;
    case "IMG": acc.img(node, ctx); return;
    case "HR": acc.hr(node, ctx); return;
    case "WBR": return;
    default: break;
  }
  if (BLOCK_SET.has(tag)) {
    acc.block(node, ctx);
    return;
  }
  walkChildren(node, inlineCtx(node, tag, ctx), acc);
}

function walkChildren(parent, ctx, acc) {
  for (const child of parent.childNodes || []) walkNode(child, ctx, acc);
}

function parseItem(li, list, level, ctx, out) {
  const acc = new ItemAcc(out, list, level);
  walkChildren(li, blockCtx(li, ctx), acc);
  acc.finish();
}

/** <ul>/<ol>: <li> items, nested lists directly inside (Chrome indent) = level+1. */
function parseList(el, level, ctx, out) {
  const list = tagOf(el) === "OL" ? "ol" : "ul";
  const lctx = blockCtx(el, ctx);
  let stray = null;
  const strayAcc = () => stray || (stray = new ItemAcc(out, list, level, true));
  const endStray = () => {
    if (stray) stray.finish();
    stray = null;
  };
  for (const child of el.childNodes || []) {
    if (child.nodeType === 3) {
      if (!stray && isAllWs(child.nodeValue || "")) continue;
      strayAcc().text(child.nodeValue || "", lctx);
      continue;
    }
    if (child.nodeType !== 1) continue;
    const tag = tagOf(child);
    if (SKIP_TAGS.has(tag) || isHiddenEl(child)) continue;
    if (tag === "LI") {
      endStray();
      parseItem(child, list, level, lctx, out);
    } else if (tag === "UL" || tag === "OL") {
      endStray();
      parseList(child, Math.min(MAX_LIST_LEVEL, level + 1), lctx, out);
    } else if (tag === "TABLE") {
      endStray();
      parseTable(child, lctx, out);
    } else if (tag === "IMG" || tag === "HR") {
      endStray();
      walkNode(child, lctx, new FlowAcc(out));
    } else if (BLOCK_SET.has(tag)) {
      // <p>/<div> directly inside a list: treat as an item.
      endStray();
      parseItem(child, list, level, lctx, out);
    } else {
      walkNode(child, lctx, strayAcc());
    }
  }
  endStray();
}

function tableRows(table) {
  const trs = [];
  for (const child of table.children || []) {
    const tag = tagOf(child);
    if (tag === "TR") trs.push(child);
    else if (tag === "TBODY" || tag === "THEAD" || tag === "TFOOT") {
      for (const tr of child.children || []) if (tagOf(tr) === "TR") trs.push(tr);
    }
  }
  return trs;
}

function parseTable(table, ctx, out) {
  // A caption has no model equivalent: keep its text as a paragraph above.
  for (const child of table.children || []) {
    if (tagOf(child) === "CAPTION") {
      const cap = { ...ctx, marks: {}, align: undefined, paraType: "p", pre: false };
      const acc = new FlowAcc(out, cap);
      walkChildren(child, cap, acc);
      acc.finish();
    }
  }
  const rows = [];
  for (const tr of tableRows(table)) {
    const cells = [];
    for (const td of tr.children || []) {
      const tag = tagOf(td);
      if (tag !== "TD" && tag !== "TH") continue;
      const cctx = blockCtx(td, { ...ctx, marks: {}, align: undefined, paraType: "p", pre: false });
      const blocks = [];
      const acc = new FlowAcc(blocks, cctx);
      walkChildren(td, cctx, acc);
      acc.finish();
      const cell = { blocks: blocks.length ? blocks : [{ type: "p", runs: [] }] };
      const cs = parseInt(td.getAttribute("colspan") || "1", 10);
      if (cs > 1) cell.colspan = Math.min(cs, 1000);
      cells.push(cell);
    }
    if (cells.length) rows.push(cells);
  }
  if (rows.length) out.push({ type: "table", rows });
}

function parseBlockEl(el, ctx, out) {
  const tag = tagOf(el);
  const bctx = blockCtx(el, ctx);
  if (breaksBefore(el)) out.push({ type: "pagebreak" });
  switch (tag) {
    case "H1": case "H2": case "H3": case "H4": case "H5": case "H6":
      paraEl(el, { ...bctx, paraType: tag.toLowerCase() }, out);
      break;
    case "P":
      paraEl(el, bctx, out);
      break;
    case "PRE":
      paraEl(el, { ...bctx, paraType: "code", pre: true }, out);
      break;
    case "BLOCKQUOTE":
      if (isQuoteBlock(el)) paraEl(el, { ...bctx, paraType: "quote" }, out);
      else flowEl(el, bctx, out); // execCommand('indent') wrapper
      break;
    case "UL": case "OL":
      parseList(el, 0, ctx, out);
      break;
    case "LI":
      parseItem(el, "ul", 0, ctx, out);
      break;
    case "TABLE":
      parseTable(el, ctx, out);
      break;
    case "CENTER":
      flowEl(el, { ...bctx, align: "center" }, out);
      break;
    case "TR": case "TD": case "TH": case "TBODY": case "THEAD": case "TFOOT":
      // Table parts outside a table (broken paste): keep the text.
      flowEl(el, bctx, out);
      break;
    default:
      flowEl(el, bctx, out);
  }
  if (breaksAfter(el)) out.push({ type: "pagebreak" });
}

function paraEl(el, ctx, out) {
  const acc = new ParaAcc(out, ctx);
  walkChildren(el, ctx, acc);
  acc.finish();
}

function flowEl(el, ctx, out) {
  const acc = new FlowAcc(out, ctx);
  walkChildren(el, ctx, acc);
  acc.finish();
}

/** Editor-wide defaults (computed style of the root) used to ignore Chrome's style junk. */
function rootDefaults(root, opts) {
  if (opts && opts.defaults) return opts.defaults;
  try {
    const view = root.ownerDocument && root.ownerDocument.defaultView;
    if (view && root.isConnected) {
      const cs = view.getComputedStyle(root);
      return { color: colorToHex(cs.color), sizePt: cssSizeToPt(cs.fontSize) };
    }
  } catch {
    // not in a browser / detached: no defaults
  }
  return {};
}

function rootCtx(root, opts) {
  return { marks: {}, align: undefined, paraType: "p", pre: false, defaults: rootDefaults(root, opts) };
}

/** Collect inline content (runs) from a DOM node, inheriting marks. (Legacy helper.) */
export function parseInlines(node, marks = {}, runs = [], link = null) {
  const ctx = { marks: link ? { ...marks, link } : { ...marks }, align: undefined, paraType: "p", pre: false, defaults: {} };
  walkChildren(node, ctx, new InlineAcc(runs));
  return runs;
}

/** Parse the children of a container element into model blocks. */
export function parseBlocks(root, opts = {}) {
  const out = [];
  const ctx = rootCtx(root, opts);
  const acc = new FlowAcc(out, ctx);
  walkChildren(root, ctx, acc);
  acc.finish();
  return out;
}

/**
 * Parse a contenteditable/article element into a full word model.
 * opts.meta: meta object to carry over (e.g. the title of the opened model);
 * warnings are not carried (they described the original import).
 */
export function parseWordDom(root, opts = {}) {
  const model = { kind: "word", blocks: parseBlocks(root, opts) };
  if (opts && opts.meta && typeof opts.meta === "object") {
    const meta = { ...opts.meta };
    delete meta.warnings;
    if (Object.keys(meta).length) model.meta = meta;
  }
  return model;
}

// ---------------------------------------------------------------------------
// Normalization (for comparisons / tests)
// ---------------------------------------------------------------------------

function normRuns(runs) {
  const out = [];
  for (const r of runs || []) {
    if (!r || r.t == null) continue;
    const t = String(r.t).replace(/\u00A0/g, " ").replace(/\r/g, "");
    if (!t) continue;
    const m = {};
    for (const k of ["b", "i", "u", "s", "sup", "sub"]) if (r[k]) m[k] = true;
    if (r.color) { const c = colorToHex(r.color); if (c) m.color = c; }
    if (r.hl) { const c = colorToHex(r.hl); if (c) m.hl = c; }
    if (Number(r.size) > 0) m.size = Number(r.size);
    if (r.font) m.font = String(r.font);
    if (r.link) m.link = String(r.link);
    pushRun(out, t, m);
  }
  return out;
}

function normBlocks(blocks) {
  const out = [];
  for (const b of blocks || []) {
    if (!b) continue;
    if (PARA_TYPES.has(b.type)) {
      const nb = { type: b.type, runs: normRuns(b.runs) };
      if (b.align === "center" || b.align === "right" || b.align === "justify") nb.align = b.align;
      out.push(nb);
    } else if (b.type === "list") {
      out.push({
        type: "list",
        list: b.list === "ol" ? "ol" : "ul",
        level: Math.min(MAX_LIST_LEVEL, Math.max(0, b.level | 0)),
        runs: normRuns(b.runs),
      });
    } else if (b.type === "table") {
      out.push({
        type: "table",
        rows: (b.rows || []).map((row) => (row || []).map((c) => {
          const blocks = normBlocks((c && c.blocks) || []);
          const cell = { blocks: blocks.length ? blocks : [{ type: "p", runs: [] }] };
          if (c && c.colspan > 1) cell.colspan = c.colspan | 0;
          return cell;
        })).filter((row) => row.length),
      });
    } else if (b.type === "img") {
      if (!b.src) continue;
      const nb = { type: "img", src: b.src };
      if (b.alt) nb.alt = b.alt;
      if (Number(b.w) > 0) nb.w = Math.round(Number(b.w));
      if (Number(b.h) > 0) nb.h = Math.round(Number(b.h));
      out.push(nb);
    } else if (b.type === "pagebreak") {
      out.push({ type: "pagebreak" });
    } else if (Array.isArray(b.runs)) {
      out.push({ type: "p", runs: normRuns(b.runs) });
    }
  }
  return out;
}

/**
 * Canonical form of a word model: equivalent models normalize to deep-equal
 * values (color case, left alignment, nbsp, run splitting, empty cells…).
 */
export function normalizeWordModel(model) {
  const blocks = normBlocks((model && model.blocks) || []);
  // An empty document and a document with one empty paragraph are the same.
  return { kind: "word", blocks: blocks.length ? blocks : [{ type: "p", runs: [] }] };
}

// ---------------------------------------------------------------------------
// Plain text / Markdown / standalone HTML
// ---------------------------------------------------------------------------

const runsText = (runs) => (runs || []).map((r) => (r && r.t != null ? String(r.t) : "")).join("");

/**
 * Track ordered-list numbering the way renderWordHtml nests lists: a counter
 * per level, reset when that level's list is closed (a shallower item, a
 * type change, or a non-list block).
 */
function listCounter() {
  const stack = []; // {type, level, n}
  return {
    next(b) {
      const lvl = Math.min(MAX_LIST_LEVEL, Math.max(0, b.level | 0));
      const type = b.list === "ol" ? "ol" : "ul";
      while (stack.length) {
        const top = stack[stack.length - 1];
        if (top.level > lvl || (top.level === lvl && top.type !== type)) stack.pop();
        else break;
      }
      const top = stack[stack.length - 1];
      if (!top || top.level < lvl) {
        stack.push({ type, level: lvl, n: 1 });
        return 1;
      }
      top.n += 1;
      return top.n;
    },
    reset() { stack.length = 0; },
  };
}

/** Plain-text serialization of a word model (for .txt export / word count). */
export function wordModelToText(model) {
  const lines = [];
  const counter = listCounter();
  const walk = (blocks, out, ctr) => {
    for (const b of blocks || []) {
      if (!b) continue;
      if (b.type !== "list") ctr.reset();
      switch (b.type) {
        case "list": {
          const n = ctr.next(b);
          const indent = "  ".repeat(Math.max(0, b.level | 0));
          const marker = b.list === "ol" ? `${n}.` : "•";
          out.push(`${indent}${marker} ${runsText(b.runs)}`);
          break;
        }
        case "table":
          for (const row of b.rows || []) {
            out.push((row || []).map((c) => {
              const cellLines = [];
              walk((c && c.blocks) || [], cellLines, listCounter());
              return cellLines.join(" ").replace(/[\t\n]+/g, " ");
            }).join("\t"));
          }
          break;
        case "img":
          out.push(`[image${b.alt ? `: ${b.alt}` : ""}]`);
          break;
        case "pagebreak":
          out.push("\f");
          break;
        default:
          if (Array.isArray(b.runs)) out.push(runsText(b.runs));
      }
    }
  };
  walk((model && model.blocks) || [], lines, counter);
  return lines.join("\n");
}

function mdEscape(text) {
  return String(text)
    .replace(/[\\`*_[\]<>~|]/g, "\\$&")
    .replace(/\u00A0/g, " ");
}

/** Escape characters that would start a block construct at line start. */
function mdEscapeLineStart(line) {
  return line
    .replace(/^(\s*)([#+=-])/, "$1\\$2")
    .replace(/^(\s*\d+)([.)])/, "$1\\$2");
}

function mdInline(runs) {
  let out = "";
  const list = (runs || []).filter((r) => r && r.t != null && String(r.t) !== "");
  let i = 0;
  while (i < list.length) {
    const link = list[i].link || null;
    let inner = "";
    let j = i;
    while (j < list.length && (list[j].link || null) === link) {
      inner += mdRun(list[j]);
      j++;
    }
    if (link) {
      // encodeURIComponent leaves "(" and ")" alone, so map them explicitly.
      const url = String(link).replace(/[()\s<>]/g, (c) => ({ "(": "%28", ")": "%29" }[c] || encodeURIComponent(c)));
      out += `[${inner}](${url})`;
    } else {
      out += inner;
    }
    i = j;
  }
  return out;
}

function mdRun(r) {
  const t = String(r.t);
  // Markdown emphasis may not start/end with whitespace: keep edges outside.
  const m = /^(\s*)([\s\S]*?)(\s*)$/.exec(t);
  const lead = m[1], core = m[2], trail = m[3];
  const enc = (s) => s.split("\n").map(mdEscape).join("\\\n").replace(/\t/g, "    ");
  if (!core) return enc(t);
  let s = enc(core);
  const mono = r.font && /consolas|courier|mono|menlo|lucida console/i.test(r.font);
  if (mono && !/\n/.test(core)) {
    const ticks = (core.match(/`+/g) || []).reduce((a, b) => Math.max(a, b.length), 0);
    const fence = "`".repeat(ticks + 1);
    s = `${fence}${ticks ? " " : ""}${core}${ticks ? " " : ""}${fence}`;
  }
  if (r.sup) s = `<sup>${s}</sup>`;
  if (r.sub) s = `<sub>${s}</sub>`;
  if (r.u) s = `<u>${s}</u>`;
  if (r.s) s = `~~${s}~~`;
  if (r.i) s = `*${s}*`;
  if (r.b) s = `**${s}**`;
  return enc(lead) + s + enc(trail);
}

function mdCellText(cell) {
  const parts = [];
  for (const b of (cell && cell.blocks) || []) {
    if (b.type === "img") parts.push(`![${mdEscape(b.alt || "")}](${b.src})`);
    else if (b.type === "table") parts.push(mdEscape(wordModelToText({ blocks: [b] }).replace(/\s+/g, " ")));
    else if (Array.isArray(b.runs)) parts.push(mdInline(b.runs));
  }
  return parts.join("<br>").replace(/\\\n/g, "<br>").replace(/\n/g, " ") || " ";
}

/** Markdown (GFM flavoured) serialization of a word model. */
export function wordModelToMarkdown(model) {
  const chunks = [];
  const counter = listCounter();
  let listBuf = null; // consecutive list items are one chunk (no blank lines)
  let mdLevel = -1; // level of the previous item in the current list
  const endList = () => {
    if (listBuf) chunks.push(listBuf.join("\n"));
    listBuf = null;
    counter.reset();
  };
  for (const b of (model && model.blocks) || []) {
    if (!b) continue;
    if (b.type === "list") {
      if (!listBuf) { listBuf = []; mdLevel = -1; }
      // Markdown cannot skip levels (an over-indented item reads back as code
      // or continuation text): at most one level deeper than the previous item.
      const lvl = Math.min(Math.max(0, Math.min(MAX_LIST_LEVEL, b.level | 0)), mdLevel + 1);
      mdLevel = lvl;
      const n = counter.next(lvl === b.level ? b : { ...b, level: lvl });
      const indent = "    ".repeat(lvl);
      const marker = b.list === "ol" ? `${n}.` : "-";
      const text = mdInline(b.runs).replace(/\\\n/g, `\\\n${indent}  `);
      listBuf.push(`${indent}${marker} ${text}`.replace(/\s+$/, ""));
      continue;
    }
    endList();
    switch (b.type) {
      case "h1": case "h2": case "h3": case "h4": case "h5": case "h6": {
        const text = mdInline(b.runs).replace(/\\\n/g, " ");
        if (text.trim()) chunks.push(`${"#".repeat(Number(b.type[1]))} ${text}`);
        break;
      }
      case "quote": {
        const text = mdInline(b.runs);
        chunks.push(text.split("\n").map((l) => `> ${l}`).join("\n"));
        break;
      }
      case "code": {
        const text = runsText(b.runs);
        const ticks = (text.match(/`{3,}/g) || []).reduce((a, s) => Math.max(a, s.length), 2);
        const fence = "`".repeat(ticks + 1);
        chunks.push(`${fence}\n${text}\n${fence}`);
        break;
      }
      case "table": {
        const rows = (b.rows || []).filter((r) => r && r.length);
        if (!rows.length) break;
        const width = Math.max(...rows.map((r) => r.reduce((n, c) => n + Math.max(1, (c && c.colspan) | 0 || 1), 0)));
        const line = (row) => {
          const cells = [];
          for (const c of row) {
            cells.push(mdCellText(c));
            for (let k = 1; k < ((c && c.colspan) | 0); k++) cells.push(" ");
          }
          while (cells.length < width) cells.push(" ");
          return `| ${cells.join(" | ")} |`;
        };
        const out = [line(rows[0]), `|${" --- |".repeat(width)}`];
        for (const r of rows.slice(1)) out.push(line(r));
        chunks.push(out.join("\n"));
        break;
      }
      case "img":
        if (b.src) chunks.push(`![${mdEscape(b.alt || "")}](${b.src})`);
        break;
      case "pagebreak":
        chunks.push("---");
        break;
      default: {
        if (!Array.isArray(b.runs)) break;
        const text = mdInline(b.runs);
        if (!text.trim()) break; // markdown cannot express empty paragraphs
        // Alignment has no markdown equivalent (an HTML wrapper would stop
        // the inline markdown inside it from rendering), so it is dropped.
        const lines = text.split("\n").map(mdEscapeLineStart);
        chunks.push(lines.join("\n"));
      }
    }
  }
  endList();
  return chunks.join("\n\n") + (chunks.length ? "\n" : "");
}

const HTML_DOC_CSS = `
body { margin: 0; background: #f3f3f3; }
.ow-page { box-sizing: border-box; max-width: 816px; margin: 24px auto; padding: 72px 80px; background: #fff;
  box-shadow: 0 1px 4px rgba(0,0,0,.18); font: 11pt/1.5 Calibri, "Segoe UI", Arial, sans-serif; color: #000; }
.ow-page h1 { font-size: 20pt; margin: .6em 0 .3em; } .ow-page h2 { font-size: 16pt; margin: .6em 0 .3em; }
.ow-page h3 { font-size: 14pt; margin: .5em 0 .25em; } .ow-page h4, .ow-page h5, .ow-page h6 { font-size: 12pt; margin: .5em 0 .25em; }
.ow-page p { margin: 0 0 .5em; }
.ow-page blockquote { margin: .6em 0; padding: 2px 14px; border-left: 3px solid #ccc; color: #444; font-style: italic; }
.ow-page pre { font: 10pt/1.4 Consolas, "Courier New", monospace; background: #f6f8fa; padding: 8px 10px; white-space: pre-wrap; }
.ow-page ul, .ow-page ol { margin: .3em 0; padding-left: 1.6em; }
.ow-page table.ow-table { border-collapse: collapse; margin: .6em 0; width: 100%; }
.ow-page table.ow-table td { border: 1px solid #999; padding: 4px 8px; vertical-align: top; }
.ow-page img.ow-img { max-width: 100%; height: auto; }
.ow-page a.ow-link { color: #0563C1; }
.ow-page hr.ow-pagebreak { border: none; border-top: 1px dashed #bbb; margin: 1.5em 0; break-after: page; page-break-after: always; }
.ow-page .ow-doc-header, .ow-page .ow-doc-footer { color: #595959; font-size: 9.5pt; }
.ow-page .ow-doc-header { border-bottom: 1px solid #e3e3e3; padding-bottom: 6px; margin-bottom: 18px; }
.ow-page .ow-doc-footer { border-top: 1px solid #e3e3e3; padding-top: 6px; margin-top: 24px; }
@media print {
  body { background: none; }
  .ow-page { box-shadow: none; margin: 0; max-width: none; padding: 0; }
  .ow-page hr.ow-pagebreak { border: none; margin: 0; }
}
`;

/** A complete, self-contained HTML document for a word model (Save as .html). */
export function wordModelToHtmlDocument(model, title) {
  const t = title || (model && model.meta && model.meta.title) || "Document";
  const meta = (model && model.meta) || {};
  // A web page is one page: page-number fields read "1".
  const band = (hf, cls) => {
    if (!hf || !Array.isArray(hf.runs) || !hf.runs.some((r) => r && r.t && String(r.t).trim())) return "";
    const runs = hf.runs.map((r) => ({ ...r, t: String(r.t == null ? "" : r.t).replace(/\{PAGE\}|\{NUMPAGES\}/g, "1") }));
    const tag = cls === "ow-doc-header" ? "header" : "footer";
    return `<${tag} class="${cls}"${alignStyle(hf.align)}>${renderRunsHtml(runs)}</${tag}>\n`;
  };
  const body = band(meta.header, "ow-doc-header") + renderWordHtml(model).replace(/ contenteditable="false"/g, "") +
    (meta.footer ? `\n${band(meta.footer, "ow-doc-footer")}` : "");
  return `<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${escapeHtml(t)}</title>
<style>${HTML_DOC_CSS}</style>
</head>
<body>
<article class="ow-page">
${body}
</article>
</body>
</html>
`;
}

// ---------------------------------------------------------------------------
// Editing helpers (pure; used by the viewer, tested under Node)
// ---------------------------------------------------------------------------

// A word: letters/digits with inner apostrophes and hyphens, and "." / ","
// between digits (3.5, 1,000), as Word counts them.
const WORD_RE = /[\p{L}\p{N}](?:[\p{L}\p{N}'’_-]|[.,](?=\p{N}))*/gu;

/** Word / character counts of a piece of text (Word-like word rules). */
export function textStats(text) {
  const s = String(text || "");
  const m = s.match(WORD_RE);
  return {
    words: m ? m.length : 0,
    chars: [...s.replace(/[\r\n]/g, "")].length,
    charsNoSpaces: [...s.replace(/\s/g, "")].length,
  };
}

/**
 * Change the case of `text`: "upper", "lower", "title" (Every Word
 * Capitalized), "sentence" (first letter of each sentence), "toggle".
 * The result has the same length as the input whenever the case mapping
 * does (callers rely on that to keep formatting runs aligned).
 */
export function changeCase(text, mode) {
  const s = String(text || "");
  const up1 = (ch) => { const u = ch.toUpperCase(); return u.length === ch.length ? u : ch; };
  const lo1 = (ch) => { const l = ch.toLowerCase(); return l.length === ch.length ? l : ch; };
  const map = (fn) => [...s].map(fn).join("");
  switch (mode) {
    case "upper": return map(up1);
    case "lower": return map(lo1);
    case "toggle": return map((ch) => (ch === lo1(ch) ? up1(ch) : lo1(ch)));
    case "title": {
      let prevWord = false;
      return [...s].map((ch) => {
        const isWord = /[\p{L}\p{N}'’]/u.test(ch);
        const out = isWord ? (prevWord ? lo1(ch) : up1(ch)) : ch;
        prevWord = isWord;
        return out;
      }).join("");
    }
    case "sentence": {
      let start = true;
      return [...s].map((ch) => {
        if (/\p{L}/u.test(ch)) {
          const out = start ? up1(ch) : lo1(ch);
          start = false;
          return out;
        }
        if (/[.!?\n]/.test(ch)) start = true;
        else if (/[\p{N}]/u.test(ch)) start = false;
        return ch;
      }).join("");
    }
    default: return s;
  }
}

/**
 * Find `query` in `text`. Returns [[start, end], …] (UTF-16 offsets).
 * opts.matchCase, opts.wholeWord; opts.limit caps the number of matches.
 */
export function findMatches(text, query, opts = {}) {
  const s = String(text || "");
  const q = String(query || "");
  if (!q || !s) return [];
  const esc = q.replace(/[.*+?^${}()|[\]\\]/g, "\\$&").replace(/\s+/g, "\\s+");
  const src = opts.wholeWord ? `(?<![\\p{L}\\p{N}_])${esc}(?![\\p{L}\\p{N}_])` : esc;
  let re;
  try { re = new RegExp(src, `gu${opts.matchCase ? "" : "i"}`); } catch { return []; }
  const out = [];
  const limit = opts.limit > 0 ? opts.limit : 10000;
  let m;
  while ((m = re.exec(s)) && out.length < limit) {
    if (!m[0].length) { re.lastIndex++; continue; }
    out.push([m.index, m.index + m[0].length]);
  }
  return out;
}
