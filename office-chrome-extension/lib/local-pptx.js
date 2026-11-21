/**
 * In-page PowerPoint (.pptx) reader / writer over the slides model (see
 * lib/model-slides.js CONTRACT). Uses the global JSZip (vendor/jszip.min.js).
 * Owned by the slides-frontend stream.
 *
 * Exports (used by viewer/viewer.js):
 *   async pptxToModel(bytes: Uint8Array) -> { model: deck, warnings: string[] }
 *        throws Error with a plain-language message for non-pptx / damaged input
 *   async modelToPptx(deck) -> Uint8Array   (valid .pptx that PowerPoint opens without repair)
 *
 * Reader: slide order/size, layout/master placeholder geometry and text-style
 * inheritance (lstStyle / txStyles), theme colors (schemeClr with lumMod/lumOff/
 * tint/shade) and fonts (+mj-lt/+mn-lt), paragraphs/runs/bullets/links, shapes,
 * pictures, tables, groups (flattened), connectors, backgrounds (solid, image),
 * layout/master decorations, notes, transitions, hidden slides, core properties.
 * Charts, SmartArt, media, OLE objects and freeforms become visible placeholders
 * with one warning per kind.
 *
 * Writer: a minimal but complete package (content types, rels, core/app
 * properties, presentation, presProps, viewProps, tableStyles, theme from
 * deck.theme, one master with six layouts, slides with placeholders, shapes,
 * pictures, tables, transitions, notes master + notes slides).
 *
 * XML is parsed with a small built-in parser so this also runs under Node
 * (tests); names are matched by local name.
 */

import { normalizeDeck, cloneDeep, layoutBoxes, hexColor, tintColor, textColorFor, DEFAULT_THEME, paraText } from "./model-slides.js";

// ---------------------------------------------------------------------------
// small XML parser (same approach as lib/local-docx.js; kept local so the two
// files do not depend on each other)
// ---------------------------------------------------------------------------

const ENTITY = { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'" };

function decodeEntities(s) {
  if (s.indexOf("&") < 0) return s;
  return s.replace(/&(#x[0-9a-fA-F]+|#\d+|amp|lt|gt|quot|apos);/g, (m, e) => {
    if (e[0] === "#") {
      const cp = e[1] === "x" || e[1] === "X" ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
      try { return String.fromCodePoint(cp); } catch { return m; }
    }
    return ENTITY[e];
  });
}

function localOf(name) {
  const i = name.indexOf(":");
  return i < 0 ? name : name.slice(i + 1);
}

const ATTR_RE = /\s*([^\s=/>]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/y;
const NAME_RE = /[^\s/>]+/y;

export function parseXml(text) {
  const root = { name: "#document", local: "#document", attrs: {}, children: [] };
  const stack = [root];
  const n = text.length;
  let i = text.charCodeAt(0) === 0xfeff ? 1 : 0;
  const addText = (s) => { if (s) stack[stack.length - 1].children.push({ text: s }); };
  while (i < n) {
    const lt = text.indexOf("<", i);
    if (lt < 0) { addText(decodeEntities(text.slice(i))); break; }
    if (lt > i) addText(decodeEntities(text.slice(i, lt)));
    if (text.startsWith("<!--", lt)) { const e = text.indexOf("-->", lt + 4); i = e < 0 ? n : e + 3; continue; }
    if (text.startsWith("<![CDATA[", lt)) { const e = text.indexOf("]]>", lt + 9); addText(text.slice(lt + 9, e < 0 ? n : e)); i = e < 0 ? n : e + 3; continue; }
    if (text.startsWith("<?", lt)) { const e = text.indexOf("?>", lt + 2); i = e < 0 ? n : e + 2; continue; }
    if (text.startsWith("<!", lt)) {
      let depth = 0, j = lt + 2;
      for (; j < n; j++) { const c = text[j]; if (c === "[") depth++; else if (c === "]") depth--; else if (c === ">" && depth <= 0) break; }
      i = j + 1;
      continue;
    }
    if (text[lt + 1] === "/") {
      const e = text.indexOf(">", lt);
      const name = text.slice(lt + 2, e < 0 ? n : e).trim();
      for (let k = stack.length - 1; k > 0; k--) if (stack[k].name === name) { stack.length = k; break; }
      i = e < 0 ? n : e + 1;
      continue;
    }
    NAME_RE.lastIndex = lt + 1;
    const nm = NAME_RE.exec(text);
    if (!nm) { addText("<"); i = lt + 1; continue; }
    const el = { name: nm[0], local: localOf(nm[0]), attrs: {}, children: [] };
    let j = NAME_RE.lastIndex;
    for (;;) {
      ATTR_RE.lastIndex = j;
      const am = ATTR_RE.exec(text);
      if (!am) break;
      el.attrs[am[1]] = decodeEntities(am[2] !== undefined ? am[2] : am[3]);
      j = ATTR_RE.lastIndex;
    }
    const close = text.indexOf(">", j);
    if (close < 0) break;
    const selfClosing = text[close - 1] === "/";
    if (stack.length > 400) { i = close + 1; continue; } // nesting cap (hostile input)
    stack[stack.length - 1].children.push(el);
    if (!selfClosing) stack.push(el);
    i = close + 1;
  }
  return root;
}

function kids(node, local) {
  const out = [];
  if (!node || !node.children) return out;
  for (const c of node.children) if (c.local && (!local || c.local === local)) out.push(c);
  return out;
}

function kid(node, local) {
  if (!node || !node.children) return null;
  for (const c of node.children) if (c.local === local) return c;
  return null;
}

/** kid chain: path(node, "spPr", "xfrm", "off") */
function path(node, ...names) {
  let n = node;
  for (const nm of names) { n = kid(n, nm); if (!n) return null; }
  return n;
}

function findDesc(node, local, depth = 0) {
  if (!node || !node.children || depth > 60) return null;
  for (const c of node.children) {
    if (!c.local) continue;
    if (c.local === local) return c;
    const d = findDesc(c, local, depth + 1);
    if (d) return d;
  }
  return null;
}

function attr(node, name) {
  if (!node) return null;
  const a = node.attrs;
  if (name in a) return a[name];
  const want = localOf(name);
  for (const k in a) if (localOf(k) === want) return a[k];
  return null;
}

function textOf(node) {
  if (!node) return "";
  if (node.text !== undefined) return node.text;
  let s = "";
  for (const c of node.children) s += textOf(c);
  return s;
}

function rootEl(doc) {
  return doc ? kids(doc)[0] || null : null;
}

function decodeXmlBytes(u8) {
  if (u8.length >= 2 && u8[0] === 0xff && u8[1] === 0xfe) return new TextDecoder("utf-16le").decode(u8.subarray(2));
  if (u8.length >= 2 && u8[0] === 0xfe && u8[1] === 0xff) return new TextDecoder("utf-16be").decode(u8.subarray(2));
  return new TextDecoder("utf-8").decode(u8);
}

// ---------------------------------------------------------------------------
// shared helpers
// ---------------------------------------------------------------------------

function getJSZip() {
  const J = typeof globalThis !== "undefined" ? globalThis.JSZip : undefined;
  if (!J) throw new Error("The PowerPoint support library (JSZip) is not loaded.");
  return J;
}

export function bytesToBase64(u8) {
  let s = "";
  for (let i = 0; i < u8.length; i += 0x8000) s += String.fromCharCode.apply(null, u8.subarray(i, i + 0x8000));
  return btoa(s);
}

export function base64ToBytes(b64) {
  const s = atob(String(b64).replace(/\s+/g, ""));
  const u8 = new Uint8Array(s.length);
  for (let i = 0; i < s.length; i++) u8[i] = s.charCodeAt(i);
  return u8;
}

const EXT_MIME = {
  png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", jpe: "image/jpeg", jfif: "image/jpeg", gif: "image/gif",
  bmp: "image/bmp", dib: "image/bmp", tif: "image/tiff", tiff: "image/tiff", svg: "image/svg+xml",
  webp: "image/webp", emf: "image/x-emf", wmf: "image/x-wmf", ico: "image/x-icon", avif: "image/avif",
};
const MIME_EXT = { "image/png": "png", "image/jpeg": "jpeg", "image/jpg": "jpeg", "image/gif": "gif", "image/bmp": "bmp", "image/tiff": "tiff", "image/svg+xml": "svg", "image/webp": "webp", "image/x-icon": "ico", "image/avif": "avif" };

/** Sniff the real image type from bytes (files are sometimes misnamed). */
function sniffMime(u8, fallback) {
  if (u8.length > 8 && u8[0] === 0x89 && u8[1] === 0x50 && u8[2] === 0x4e && u8[3] === 0x47) return "image/png";
  if (u8.length > 3 && u8[0] === 0xff && u8[1] === 0xd8 && u8[2] === 0xff) return "image/jpeg";
  if (u8.length > 6 && u8[0] === 0x47 && u8[1] === 0x49 && u8[2] === 0x46) return "image/gif";
  if (u8.length > 2 && u8[0] === 0x42 && u8[1] === 0x4d) return "image/bmp";
  if (u8.length > 12 && u8[0] === 0x52 && u8[1] === 0x49 && u8[2] === 0x46 && u8[3] === 0x46 && u8[8] === 0x57 && u8[9] === 0x45) return "image/webp";
  return fallback;
}

const EMU = 12700;
const pt = (v) => Math.round((Number(v) || 0) / EMU * 100) / 100;
const emu = (v) => Math.round((Number(v) || 0) * EMU);

function dirOf(p) { const i = p.lastIndexOf("/"); return i < 0 ? "" : p.slice(0, i); }

function resolvePath(base, target) {
  if (/^[a-z][a-z0-9+.-]*:/i.test(target)) return target;
  let t = target.replace(/\\/g, "/");
  try { t = decodeURIComponent(t); } catch { /* keep */ }
  const parts = t.startsWith("/") ? [] : (dirOf(base) ? dirOf(base).split("/") : []);
  for (const seg of t.replace(/^\//, "").split("/")) {
    if (seg === "..") parts.pop();
    else if (seg && seg !== ".") parts.push(seg);
  }
  return parts.join("/");
}

function relsPathOf(partPath) {
  const d = dirOf(partPath);
  const f = partPath.slice(d.length ? d.length + 1 : 0);
  return `${d ? d + "/" : ""}_rels/${f}.rels`;
}

function checkZipLimits(zip) {
  const names = Object.keys(zip.files);
  if (names.length > 20000) throw new Error("This presentation contains too many parts to open safely.");
  let total = 0;
  for (const name of names) {
    const f = zip.files[name];
    const d = f && f._data;
    if (d && typeof d.uncompressedSize === "number") total += d.uncompressedSize;
  }
  if (total > 1.5e9) throw new Error("This presentation is too large to open in the browser (over 1.5 GB uncompressed).");
}

// ---------------------------------------------------------------------------
// colors
// ---------------------------------------------------------------------------

const PRESET_COLORS = {
  black: "000000", white: "FFFFFF", red: "FF0000", green: "008000", blue: "0000FF", yellow: "FFFF00",
  gray: "808080", grey: "808080", darkGray: "A9A9A9", lightGray: "D3D3D3", orange: "FFA500", purple: "800080",
  navy: "000080", teal: "008080", silver: "C0C0C0", maroon: "800000", olive: "808000", lime: "00FF00",
  aqua: "00FFFF", cyan: "00FFFF", fuchsia: "FF00FF", magenta: "FF00FF", darkBlue: "00008B", darkRed: "8B0000",
  darkGreen: "006400", lightBlue: "ADD8E6", pink: "FFC0CB", brown: "A52A2A", gold: "FFD700",
};

function rgbToHsl(r, g, b) {
  r /= 255; g /= 255; b /= 255;
  const max = Math.max(r, g, b), min = Math.min(r, g, b);
  let h = 0, s = 0;
  const l = (max + min) / 2;
  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
    else if (max === g) h = (b - r) / d + 2;
    else h = (r - g) / d + 4;
    h /= 6;
  }
  return [h, s, l];
}

function hslToRgb(h, s, l) {
  if (s === 0) return [l * 255, l * 255, l * 255];
  const hue = (p, q, t) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  return [hue(p, q, h + 1 / 3) * 255, hue(p, q, h) * 255, hue(p, q, h - 1 / 3) * 255];
}

const toHex = (rgb) => "#" + rgb.map((v) => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, "0")).join("").toUpperCase();

function applyMods(hex, el) {
  let [r, g, b] = [1, 3, 5].map((i) => parseInt(hex.slice(i, i + 2), 16));
  for (const m of kids(el)) {
    const v = Number(attr(m, "val")) / 100000;
    if (!Number.isFinite(v)) continue;
    if (m.local === "lumMod" || m.local === "lumOff") {
      const [h, s, l] = rgbToHsl(r, g, b);
      const nl = m.local === "lumMod" ? l * v : l + v;
      [r, g, b] = hslToRgb(h, s, Math.max(0, Math.min(1, nl)));
    } else if (m.local === "satMod") {
      const [h, s, l] = rgbToHsl(r, g, b);
      [r, g, b] = hslToRgb(h, Math.max(0, Math.min(1, s * v)), l);
    } else if (m.local === "tint") {
      [r, g, b] = [r, g, b].map((c) => c + (255 - c) * (1 - v));
    } else if (m.local === "shade") {
      [r, g, b] = [r, g, b].map((c) => c * v);
    } else if (m.local === "inv") {
      [r, g, b] = [255 - r, 255 - g, 255 - b];
    } else if (m.local === "gray") {
      const y = 0.3 * r + 0.59 * g + 0.11 * b;
      [r, g, b] = [y, y, y];
    }
  }
  return toHex([r, g, b]);
}

/** Color of a color-container element (solidFill, fgClr, a:lnRef ...) -> "#RRGGBB" | null */
function colorOf(container, ctx, phClr) {
  if (!container) return null;
  for (const el of kids(container)) {
    let hex = null;
    switch (el.local) {
      case "srgbClr": hex = hexColor(attr(el, "val")); break;
      case "sysClr": hex = hexColor(attr(el, "lastClr")) || (attr(el, "val") === "window" ? "#FFFFFF" : "#000000"); break;
      case "prstClr": hex = hexColor(PRESET_COLORS[attr(el, "val")] || ""); break;
      case "scrgbClr": {
        const c = ["r", "g", "b"].map((k) => {
          const v = Number(attr(el, k)) / 100000;
          const s = v <= 0.0031308 ? v * 12.92 : 1.055 * v ** (1 / 2.4) - 0.055;
          return s * 255;
        });
        hex = toHex(c);
        break;
      }
      case "hslClr": hex = toHex(hslToRgb(Number(attr(el, "hue")) / 21600000, Number(attr(el, "sat")) / 100000, Number(attr(el, "lum")) / 100000)); break;
      case "schemeClr": {
        let name = attr(el, "val");
        if (name === "phClr") { hex = phClr || null; break; }
        const map = (ctx && ctx.clrMap) || {};
        if (map[name]) name = map[name];
        else if (name === "bg1") name = "lt1"; else if (name === "tx1") name = "dk1"; else if (name === "bg2") name = "lt2"; else if (name === "tx2") name = "dk2";
        hex = ctx && ctx.theme && ctx.theme.colors[name] ? ctx.theme.colors[name] : null;
        break;
      }
      default: continue;
    }
    if (!hex) return null;
    return applyMods(hex, el);
  }
  return null;
}

// ---------------------------------------------------------------------------
// reader
// ---------------------------------------------------------------------------

const PRST_MAP = {
  rect: "rect", roundRect: "roundRect", round1Rect: "roundRect", round2SameRect: "roundRect", round2DiagRect: "roundRect",
  snipRect: "rect", snip1Rect: "rect", snip2SameRect: "rect", snipRoundRect: "roundRect", plaque: "roundRect",
  ellipse: "ellipse", flowChartConnector: "ellipse", donut: "ellipse", pie: "ellipse", chord: "ellipse",
  triangle: "triangle", rtTriangle: "triangle", flowChartExtract: "triangle", flowChartMerge: "triangle",
  diamond: "diamond", flowChartDecision: "diamond",
  rightArrow: "rightArrow", notchedRightArrow: "rightArrow", stripedRightArrow: "rightArrow", rightArrowCallout: "rightArrow",
  leftArrow: "leftArrow", leftArrowCallout: "leftArrow",
  line: "line", straightConnector1: "line", bentConnector2: "line", bentConnector3: "line", bentConnector4: "line",
  curvedConnector3: "line", curvedConnector2: "line", curvedConnector4: "line",
  star5: "star5", star4: "star5", star6: "star5", star7: "star5", star8: "star5",
  hexagon: "hexagon", octagon: "hexagon", flowChartPreparation: "hexagon",
  chevron: "chevron", homePlate: "chevron",
  flowChartProcess: "rect", flowChartAlternateProcess: "roundRect", flowChartTerminator: "roundRect",
  frame: "rect", bevel: "rect", plus: "rect", can: "rect", cube: "rect", parallelogram: "rect", trapezoid: "rect",
  pentagon: "hexagon", heptagon: "hexagon", decagon: "hexagon", dodecagon: "hexagon",
};

const ALIGN_MAP = { l: "left", ctr: "center", r: "right", just: "justify", dist: "justify", justLow: "justify" };

function phInfo(sp) {
  const nv = kids(sp).find((k) => /^nv/.test(k.local));
  const ph = nv ? path(nv, "nvPr", "ph") : null;
  if (!ph) return null;
  return { type: attr(ph, "type") || "body", idx: attr(ph, "idx") };
}

const TITLE_TYPES = new Set(["title", "ctrTitle"]);
const BODY_TYPES = new Set(["body", "subTitle", "obj", "tbl", "chart", "dgm", "media", "clipArt", "pic"]);

function phFamily(type) {
  if (TITLE_TYPES.has(type)) return "title";
  if (BODY_TYPES.has(type)) return "body";
  return type;
}

class Reader {
  constructor(zip) {
    this.zip = zip;
    this.parts = new Map();
    this.relCache = new Map();
    this.warnings = new Set();
    this.media = new Map();
  }

  warn(msg) { this.warnings.add(msg); }

  async xml(p) {
    if (this.parts.has(p)) return this.parts.get(p);
    const f = this.zip.file(p);
    let root = null;
    if (f) {
      const u8 = await f.async("uint8array");
      root = rootEl(parseXml(decodeXmlBytes(u8)));
    }
    this.parts.set(p, root);
    return root;
  }

  async rels(p) {
    if (this.relCache.has(p)) return this.relCache.get(p);
    const map = new Map();
    const root = await this.xml(relsPathOf(p));
    for (const r of kids(root, "Relationship")) {
      const ext = attr(r, "TargetMode") === "External";
      const target = attr(r, "Target") || "";
      map.set(attr(r, "Id"), { type: attr(r, "Type") || "", target: ext ? target : resolvePath(p, target), external: ext });
    }
    this.relCache.set(p, map);
    return map;
  }

  async relOfType(p, suffix) {
    const rels = await this.rels(p);
    for (const r of rels.values()) if (r.type.endsWith(suffix)) return r;
    return null;
  }

  async imageDataUrl(target) {
    if (this.media.has(target)) return this.media.get(target);
    const f = this.zip.file(target);
    let res = null;
    if (f) {
      const u8 = await f.async("uint8array");
      const ext = (target.split(".").pop() || "").toLowerCase();
      const mime = sniffMime(u8, EXT_MIME[ext] || "application/octet-stream");
      res = { mime, url: `data:${mime};base64,${bytesToBase64(u8)}`, size: u8.length };
    }
    this.media.set(target, res);
    return res;
  }

  // ---- theme / master / layout contexts

  async theme(p) {
    const key = `theme:${p}`;
    if (this.parts.has(key)) return this.parts.get(key);
    const root = p ? await this.xml(p) : null;
    const colors = {
      dk1: "#000000", lt1: "#FFFFFF", dk2: "#44546A", lt2: "#E7E6E6", accent1: "#4472C4", accent2: "#ED7D31", accent3: "#A5A5A5",
      accent4: "#FFC000", accent5: "#5B9BD5", accent6: "#70AD47", hlink: "#0563C1", folHlink: "#954F72",
    };
    let major = "Calibri Light", minor = "Calibri";
    const te = path(root, "themeElements");
    const cs = kid(te, "clrScheme");
    for (const c of kids(cs)) {
      const v = colorOf({ children: [kids(c)[0]].filter(Boolean) }, null);
      if (v) colors[c.local] = v;
    }
    const fs = kid(te, "fontScheme");
    const mj = attr(path(fs, "majorFont", "latin"), "typeface");
    const mn = attr(path(fs, "minorFont", "latin"), "typeface");
    if (mj) major = mj;
    if (mn) minor = mn;
    const bgFills = kids(path(te, "fmtScheme", "bgFillStyleLst"));
    const t = { colors, major, minor, bgFills };
    this.parts.set(key, t);
    return t;
  }

  async master(p) {
    const key = `master:${p}`;
    if (this.parts.has(key)) return this.parts.get(key);
    const root = await this.xml(p);
    const themeRel = await this.relOfType(p, "/theme");
    const theme = await this.theme(themeRel ? themeRel.target : null);
    const clrMap = {};
    const cm = kid(root, "clrMap");
    if (cm) for (const k in cm.attrs) clrMap[k] = cm.attrs[k];
    const tx = kid(root, "txStyles");
    const m = {
      path: p, root, theme, clrMap, rels: await this.rels(p),
      spTree: path(root, "cSld", "spTree"),
      bg: path(root, "cSld", "bg"),
      titleStyle: kid(tx, "titleStyle"), bodyStyle: kid(tx, "bodyStyle"), otherStyle: kid(tx, "otherStyle"),
    };
    m.phs = this.collectPhs(m.spTree);
    this.parts.set(key, m);
    return m;
  }

  async layout(p) {
    const key = `layout:${p}`;
    if (this.parts.has(key)) return this.parts.get(key);
    const root = await this.xml(p);
    const mRel = await this.relOfType(p, "/slideMaster");
    const master = mRel ? await this.master(mRel.target) : null;
    const ovr = path(root, "clrMapOvr", "overrideClrMapping");
    let clrMap = master ? master.clrMap : {};
    if (ovr) { clrMap = {}; for (const k in ovr.attrs) clrMap[k] = ovr.attrs[k]; }
    const l = {
      path: p, root, master, clrMap, rels: await this.rels(p),
      type: attr(root, "type") || "",
      spTree: path(root, "cSld", "spTree"),
      bg: path(root, "cSld", "bg"),
      showMasterSp: attr(root, "showMasterSp") !== "0",
    };
    l.phs = this.collectPhs(l.spTree);
    this.parts.set(key, l);
    return l;
  }

  collectPhs(spTree) {
    const out = [];
    for (const sp of kids(spTree)) {
      const ph = phInfo(sp);
      if (ph) out.push({ ...ph, sp });
    }
    return out;
  }

  /** Inheritance chain for a placeholder: [layoutPh, masterPh] (sp nodes, may be null). */
  phChain(ph, layout) {
    let lp = null, mp = null;
    if (layout && ph) {
      if (ph.idx != null) lp = layout.phs.find((x) => x.idx === ph.idx && (ph.idx !== "0" || phFamily(x.type) === phFamily(ph.type))) || null;
      if (!lp) lp = layout.phs.find((x) => x.type === ph.type) || layout.phs.find((x) => phFamily(x.type) === phFamily(ph.type) && (phFamily(ph.type) !== "body" || ph.idx == null || x.idx == null)) || null;
      const m = layout.master;
      if (m) {
        const t = lp ? lp.type : ph.type;
        mp = m.phs.find((x) => x.type === t) || m.phs.find((x) => phFamily(x.type) === phFamily(t)) || null;
      }
    }
    return [lp && lp.sp, mp && mp.sp];
  }

  // ---- geometry

  xfrmOf(node) {
    // sp/pic/cxnSp: spPr/xfrm; graphicFrame: p:xfrm; grpSp: grpSpPr/xfrm
    const x = path(node, "spPr", "xfrm") || kid(node, "xfrm") || path(node, "grpSpPr", "xfrm");
    if (!x) return null;
    const off = kid(x, "off"), ext = kid(x, "ext");
    if (!off || !ext) return null;
    return {
      x: Number(attr(off, "x")) || 0, y: Number(attr(off, "y")) || 0,
      cx: Number(attr(ext, "cx")) || 0, cy: Number(attr(ext, "cy")) || 0,
      rot: (Number(attr(x, "rot")) || 0) / 60000,
      flipH: attr(x, "flipH") === "1", flipV: attr(x, "flipV") === "1",
      chOff: kid(x, "chOff"), chExt: kid(x, "chExt"),
    };
  }

  // ---- text

  /** lvlNpPr from a list-style container for level L (0-based). */
  lvlPr(lst, L) {
    if (!lst) return null;
    return kid(lst, `lvl${L + 1}pPr`) || null;
  }

  /**
   * Resolve paragraph-level defaults for level L from a chain of list styles
   * (most specific first): { align, bullet, size, b, i, u, color, font }.
   */
  resolveLevel(lsts, L, ctx, ownCount = 0) {
    const out = {};
    for (let li = 0; li < lsts.length; li++) {
      const lst = lsts[li];
      const p = this.lvlPr(lst, L);
      if (!p) continue;
      if (out.align === undefined && attr(p, "algn")) out.align = ALIGN_MAP[attr(p, "algn")] || "left";
      if (out.bullet === undefined) {
        if (kid(p, "buNone")) out.bullet = null;
        else if (kid(p, "buAutoNum")) out.bullet = "ol";
        else if (kid(p, "buChar") || kid(p, "buBlip")) out.bullet = "ul";
      }
      const d = kid(p, "defRPr");
      if (d) {
        if (out.size === undefined && attr(d, "sz")) out.size = Number(attr(d, "sz")) / 100;
        if (out.b === undefined && attr(d, "b") != null) out.b = attr(d, "b") === "1" || attr(d, "b") === "true";
        if (out.i === undefined && attr(d, "i") != null) out.i = attr(d, "i") === "1" || attr(d, "i") === "true";
        if (out.color === undefined && kid(d, "solidFill")) { out.color = colorOf(kid(d, "solidFill"), ctx); out.own = li < ownCount; }
        if (out.font === undefined && attr(kid(d, "latin"), "typeface")) out.font = attr(kid(d, "latin"), "typeface");
      }
    }
    return out;
  }

  fontName(face, ctx) {
    if (!face) return undefined;
    if (face === "+mj-lt" || face === "+mj-ea" || face === "+mj-cs") return ctx.theme.major;
    if (face === "+mn-lt" || face === "+mn-ea" || face === "+mn-cs") return ctx.theme.minor;
    return face;
  }

  /**
   * txBody -> paras. opts: { lsts: [list styles most specific first], defColor, defFont,
   * fontScale, item }
   */
  readParas(txBody, ctx, opts) {
    const paras = [];
    if (!txBody) return paras;
    const ownLst = kid(txBody, "lstStyle");
    const lsts = [ownLst, ...(opts.lsts || [])].filter(Boolean);
    const ownN = ownLst ? 1 : 0;
    const scale = opts.fontScale || 1;
    for (const p of kids(txBody, "p")) {
      const pPr = kid(p, "pPr");
      const L = Math.max(0, Math.min(8, Number(attr(pPr, "lvl")) || 0));
      const lv = this.resolveLevel(pPr ? [{ children: [{ ...pPr, local: `lvl${L + 1}pPr` }] }, ...lsts] : lsts, L, ctx, ownN + (pPr ? 1 : 0));
      const para = { runs: [] };
      if (lv.align && lv.align !== "left") para.align = lv.align;
      if (L) para.level = L;
      if (lv.bullet) para.bullet = lv.bullet;
      const baseSize = lv.size || opts.defSize || 18;
      const runFrom = (rPr, t, extra) => {
        const r = { t };
        const sz = attr(rPr, "sz") ? Number(attr(rPr, "sz")) / 100 : baseSize;
        r.size = Math.round(sz * scale * 10) / 10;
        const b = attr(rPr, "b") != null ? attr(rPr, "b") === "1" || attr(rPr, "b") === "true" : lv.b;
        const i = attr(rPr, "i") != null ? attr(rPr, "i") === "1" || attr(rPr, "i") === "true" : lv.i;
        if (b) r.b = true;
        if (i) r.i = true;
        const u = attr(rPr, "u");
        if (u && u !== "none") r.u = true;
        const st = attr(rPr, "strike");
        if (st && st !== "noStrike") r.s = true;
        const bl = Number(attr(rPr, "baseline")) || 0;
        if (bl > 0) r.sup = true; else if (bl < 0) r.sub = true;
        let color = rPr && kid(rPr, "solidFill") ? colorOf(kid(rPr, "solidFill"), ctx) : null;
        if (!color && rPr && kid(rPr, "gradFill")) color = colorOf(findDesc(kid(rPr, "gradFill"), "gs"), ctx);
        if (!color) color = (opts.styleColor && !lv.own ? opts.styleColor : null) || lv.color || opts.defColor || null;
        if (color) r.color = color; // colors equal to the rendering default are dropped later
        const hl = rPr && kid(rPr, "highlight") ? colorOf(kid(rPr, "highlight"), ctx) : null;
        if (hl) r.hl = hl;
        const face = this.fontName(attr(kid(rPr, "latin"), "typeface"), ctx) || this.fontName(lv.font, ctx) || opts.defFont;
        if (face && face !== opts.baseFont) r.font = face;
        const hk = kid(rPr, "hlinkClick");
        if (hk) {
          const rid = attr(hk, "r:id") || attr(hk, "id");
          const rel = rid && ctx.rels ? ctx.rels.get(rid) : null;
          if (rel && rel.external) r.link = rel.target;
        }
        Object.assign(r, extra || {});
        return r;
      };
      for (const c of kids(p)) {
        if (c.local === "r") para.runs.push(runFrom(kid(c, "rPr"), textOf(kid(c, "t"))));
        else if (c.local === "fld") para.runs.push(runFrom(kid(c, "rPr"), textOf(kid(c, "t"))));
        else if (c.local === "br") para.runs.push(runFrom(kid(c, "rPr"), "\n"));
        else if (c.local === "AlternateContent" || c.local === "m") {
          const t = textOf(c).trim();
          if (t) { para.runs.push(runFrom(null, t)); this.warn("Equations are shown as plain text."); }
        }
      }
      // merge adjacent runs with identical formatting
      const merged = [];
      for (const r of para.runs) {
        const prev = merged[merged.length - 1];
        if (prev && sameFmt(prev, r)) prev.t += r.t;
        else merged.push(r);
      }
      para.runs = merged.filter((r) => r.t !== "");
      if (!para.runs.length) {
        const end = kid(p, "endParaRPr");
        const sz = end && attr(end, "sz") ? Number(attr(end, "sz")) / 100 : baseSize;
        para.runs = [{ t: "", size: Math.round(sz * scale * 10) / 10 }];
      }
      paras.push(para);
    }
    return paras;
  }

  bodyProps(chain) {
    const out = {};
    for (const bp of chain) {
      if (!bp) continue;
      if (out.anchor === undefined && attr(bp, "anchor")) out.anchor = attr(bp, "anchor");
      if (out.fontScale === undefined) {
        const na = kid(bp, "normAutofit");
        if (na && attr(na, "fontScale")) out.fontScale = Number(attr(na, "fontScale")) / 100000;
      }
      if (out.vert === undefined && attr(bp, "vert") && attr(bp, "vert") !== "horz") out.vert = attr(bp, "vert");
    }
    return out;
  }

  // ---- fills

  /** spPr fill -> { fill?: hex | null, explicit: bool } */
  fillOf(spPr, style, ctx) {
    if (spPr) {
      if (kid(spPr, "noFill")) return { fill: null, explicit: true };
      if (kid(spPr, "solidFill")) return { fill: colorOf(kid(spPr, "solidFill"), ctx), explicit: true };
      if (kid(spPr, "gradFill")) {
        this.warn("Gradient fills are shown as solid colors.");
        const gs = findDesc(kid(spPr, "gradFill"), "gs");
        return { fill: colorOf(gs, ctx), explicit: true };
      }
      if (kid(spPr, "pattFill")) return { fill: colorOf(kid(kid(spPr, "pattFill"), "fgClr"), ctx), explicit: true };
      if (kid(spPr, "blipFill")) { this.warn("Picture fills inside shapes are not shown."); return { fill: "#D9D9D9", explicit: true }; }
      if (kid(spPr, "grpFill")) return { fill: ctx.grpFill || null, explicit: true };
    }
    const fr = style ? kid(style, "fillRef") : null;
    if (fr && Number(attr(fr, "idx")) > 0) return { fill: colorOf(fr, ctx), explicit: false };
    return { fill: null, explicit: false };
  }

  lineOf(spPr, style, ctx) {
    const ln = spPr ? kid(spPr, "ln") : null;
    let color = null, w = null, none = false;
    if (ln) {
      if (kid(ln, "noFill")) none = true;
      else if (kid(ln, "solidFill")) color = colorOf(kid(ln, "solidFill"), ctx);
      else if (kid(ln, "gradFill")) color = colorOf(findDesc(kid(ln, "gradFill"), "gs"), ctx);
      if (attr(ln, "w")) w = pt(attr(ln, "w"));
    }
    if (!none && !color) {
      const lr = style ? kid(style, "lnRef") : null;
      if (lr && Number(attr(lr, "idx")) > 0) {
        color = colorOf(lr, ctx);
        if (w == null) w = 0.75;
      }
    }
    if (none || !color) return {};
    return { line: color, lineW: w || 0.75 };
  }

  // ---- shapes

  /**
   * Convert the children of an spTree (or group) into items.
   * env = { ctx (theme/clrMap/rels/path), layout, tf: {sx,sy,dx,dy}, fromSlide, depth }
   */
  async readTree(tree, env, items) {
    if (!tree || env.depth > 30) return;
    for (const node of kids(tree)) {
      if (items.length > 5000) { this.warn("Some slides have too many objects; only the first 5000 are shown."); return; }
      let n = node;
      if (n.local === "AlternateContent") {
        const fb = kid(n, "Fallback");
        const ch = kid(n, "Choice");
        const pick = fb && kids(fb).length ? fb : ch;
        if (pick) await this.readTree(pick, env, items);
        continue;
      }
      switch (n.local) {
        case "sp": await this.readSp(n, env, items); break;
        case "cxnSp": this.readCxn(n, env, items); break;
        case "pic": await this.readPic(n, env, items); break;
        case "graphicFrame": await this.readFrame(n, env, items); break;
        case "grpSp": await this.readGroup(n, env, items); break;
        case "contentPart": this.warn("Ink drawings are not shown."); break;
        default: break;
      }
    }
  }

  place(xf, env) {
    const t = env.tf;
    const out = {
      x: Math.round(pt(xf.x * t.sx + t.dx) * 100) / 100,
      y: Math.round(pt(xf.y * t.sy + t.dy) * 100) / 100,
      w: Math.round(pt(xf.cx * t.sx) * 100) / 100,
      h: Math.round(pt(xf.cy * t.sy) * 100) / 100,
    };
    const rot = ((xf.rot + (env.rot || 0)) % 360 + 360) % 360;
    if (rot > 0.01) out.rot = Math.round(rot * 100) / 100;
    return out;
  }

  async readSp(sp, env, items) {
    const ctx = env.ctx;
    const ph = phInfo(sp);
    if (ph && !env.fromSlide) return; // layout/master placeholders are templates, not content
    const [lp, mp] = ph ? this.phChain(ph, env.layout) : [null, null];
    let xf = this.xfrmOf(sp) || (lp && this.xfrmOf(lp)) || (mp && this.xfrmOf(mp));
    if (!xf) {
      if (!ph) return;
      xf = { x: 0, y: 0, cx: 0, cy: 0, rot: 0 };
    }
    const spPr = kid(sp, "spPr");
    const style = kid(sp, "style");
    const geomEl = kid(spPr, "prstGeom");
    const prst = geomEl ? attr(geomEl, "prst") : kid(spPr, "custGeom") ? "custom" : "rect";
    const txBody = kid(sp, "txBody");
    const fillInfo = this.fillOf(spPr, style, ctx);
    const lineInfo = this.lineOf(spPr, style, ctx);
    // inherited fill/line for placeholders
    let fill = fillInfo.fill;
    if (!fillInfo.explicit && ph) {
      for (const s of [lp, mp]) {
        if (!s) continue;
        const f = this.fillOf(kid(s, "spPr"), kid(s, "style"), ctx);
        if (f.explicit || f.fill) { fill = f.fill; break; }
      }
    }
    const geom = this.place(xf, env);
    const family = ph ? phFamily(ph.type) : null;
    const lstChain = [];
    if (ph) {
      for (const s of [lp, mp]) if (s) lstChain.push(path(s, "txBody", "lstStyle"));
      const m = env.layout && env.layout.master;
      if (m) lstChain.push(family === "title" ? m.titleStyle : family === "body" ? m.bodyStyle : m.otherStyle);
    } else {
      lstChain.push(env.defaultTextStyle, env.layout && env.layout.master && env.layout.master.otherStyle);
    }
    const bp = this.bodyProps([kid(txBody, "bodyPr"), lp && path(lp, "txBody", "bodyPr"), mp && path(mp, "txBody", "bodyPr")]);
    if (bp.vert) this.warn("Vertical text is shown horizontally.");
    const fontRef = style ? kid(style, "fontRef") : null;
    const styleColor = fontRef ? colorOf(fontRef, ctx) : null;
    const isTitle = family === "title";
    const baseFont = isTitle ? ctx.theme.major : ctx.theme.minor;
    const themeText = env.themeText;
    const defSize = isTitle ? 44 : family === "body" ? 28 : 18;
    const paras = this.readParas(txBody, ctx, {
      lsts: lstChain.filter(Boolean), styleColor: styleColor || undefined, baseFont,
      fontScale: bp.fontScale, defSize,
    });
    const hasTxt = paras.some((p) => paraText(p).trim() !== "");
    const valign = { t: "top", ctr: "middle", b: "bottom" }[bp.anchor || (isTitle ? "ctr" : "t")];
    const isTextBox = attr(path(sp, "nvSpPr", "cNvSpPr"), "txBox") === "1";
    if (ph) {
      const t = ph.type;
      const phName = t === "ctrTitle" || t === "title" ? "title" : t === "subTitle" ? "subtitle" : (t === "body" || t === "obj") ? "body" : null;
      if (!phName && !hasTxt) return; // empty date/footer/picture placeholders are invisible
      const it = { type: "text", ...geom };
      if (phName) it.ph = phName;
      it.paras = paras;
      if (fill) it.fill = fill;
      if (lineInfo.line) { it.line = lineInfo.line; it.lineW = lineInfo.lineW; }
      it.valign = valign || "top";
      items.push(it);
      return;
    }
    let shape = prst === "custom" ? null : PRST_MAP[prst];
    if (prst === "custom") {
      if (!fill && !lineInfo.line && hasTxt) shape = "rect";
      else this.warn("Freeform shapes are shown as rectangles.");
    } else if (!shape) {
      this.warn("Some shapes are shown as simpler shapes (rectangles).");
    }
    const plainText = (isTextBox || (!fill && !lineInfo.line)) && (shape === "rect" || !shape) && hasTxt;
    if (plainText || (isTextBox && (shape === "rect" || !shape))) {
      const it = { type: "text", ...geom, paras };
      if (fill) it.fill = fill;
      if (lineInfo.line) { it.line = lineInfo.line; it.lineW = lineInfo.lineW; }
      it.valign = valign || "top";
      items.push(it);
      return;
    }
    const it = { type: "shape", ...geom, shape: shape || "rect" };
    if (fill && it.shape !== "line") it.fill = fill;
    if (lineInfo.line) { it.line = lineInfo.line; it.lineW = lineInfo.lineW; }
    if (hasTxt) { it.paras = paras; it.valign = { t: "top", ctr: "middle", b: "bottom" }[bp.anchor || "ctr"] || "middle"; }
    if (it.shape === "line") this.fixLine(it, xf);
    items.push(it);
  }

  fixLine(it, xf) {
    // our line runs from the top-left to the bottom-right corner; a flipped line
    // runs the other way: express it with a rotation around the centre.
    if ((xf.flipH ? 1 : 0) ^ (xf.flipV ? 1 : 0)) {
      const w = it.w, h = it.h;
      const len = Math.sqrt(w * w + h * h);
      const ang = -Math.atan2(h, w) * 180 / Math.PI; // direction of the flipped diagonal
      const cx = it.x + w / 2, cy = it.y + h / 2;
      it.w = Math.round(len * 100) / 100; it.h = 0;
      it.x = Math.round((cx - len / 2) * 100) / 100; it.y = Math.round(cy * 100) / 100;
      const rot = ((ang + (it.rot || 0)) % 360 + 360) % 360;
      if (rot) it.rot = Math.round(rot * 100) / 100; else delete it.rot;
    }
  }

  readCxn(n, env, items) {
    const xf = this.xfrmOf(n);
    if (!xf) return;
    const spPr = kid(n, "spPr");
    const ln = this.lineOf(spPr, kid(n, "style"), env.ctx);
    const it = { type: "shape", ...this.place(xf, env), shape: "line", line: ln.line || env.themeText || "#000000", lineW: ln.lineW || 0.75 };
    this.fixLine(it, xf);
    items.push(it);
  }

  async readPic(n, env, items) {
    const xf = this.xfrmOf(n);
    if (!xf) return;
    const geom = this.place(xf, env);
    const nvPr = path(n, "nvPicPr", "nvPr");
    if (nvPr && (kid(nvPr, "videoFile") || kid(nvPr, "audioFile") || findDesc(nvPr, "media"))) {
      this.warn("Video and audio are not played here; their preview picture is shown.");
    }
    const blip = path(n, "blipFill", "blip");
    const rid = blip ? attr(blip, "r:embed") || attr(blip, "embed") : null;
    const rel = rid ? env.ctx.rels.get(rid) : null;
    const descr = attr(path(n, "nvPicPr", "cNvPr"), "descr") || "";
    let img = rel && !rel.external ? await this.imageDataUrl(rel.target) : null;
    if (rel && rel.external) this.warn("Linked (external) pictures are not shown.");
    if (img && (img.mime === "image/x-emf" || img.mime === "image/x-wmf" || img.mime === "image/tiff")) {
      this.warn("Some pictures (EMF/WMF/TIFF) cannot be shown in the browser and appear as placeholders.");
      img = null;
    }
    if (path(n, "blipFill", "srcRect") && kids(path(n, "blipFill", "srcRect")).length === 0 && Object.keys(path(n, "blipFill", "srcRect").attrs).length) {
      this.warn("Picture cropping is not shown; cropped pictures appear uncropped.");
    }
    if (!img) {
      items.push(placeholderItem(geom, "[Picture]"));
      return;
    }
    const it = { type: "image", ...geom, src: img.url };
    if (descr) it.alt = descr;
    items.push(it);
  }

  async readFrame(n, env, items) {
    const xf = this.xfrmOf(n);
    if (!xf) return;
    const geom = this.place(xf, env);
    const gd = path(n, "graphic", "graphicData");
    const uri = attr(gd, "uri") || "";
    if (/\/table$/.test(uri) && kid(gd, "tbl")) {
      items.push(this.readTable(kid(gd, "tbl"), geom, env));
      return;
    }
    let label = "[Object]";
    if (/chart/i.test(uri)) { label = "[Chart]"; this.warn("Charts are shown as placeholders."); }
    else if (/diagram/i.test(uri)) { label = "[SmartArt]"; this.warn("SmartArt graphics are shown as placeholders."); }
    else if (/ole/i.test(uri)) { label = "[Embedded object]"; this.warn("Embedded objects are shown as placeholders."); }
    else if (/media|video|audio/i.test(uri)) { label = "[Media]"; this.warn("Video and audio are shown as placeholders."); }
    else this.warn("Some objects are shown as placeholders.");
    const g = { ...geom };
    delete g.rot;
    items.push(placeholderItem(g, label));
  }

  readTable(tbl, geom, env) {
    const ctx = env.ctx;
    const colW = kids(kid(tbl, "tblGrid"), "gridCol").map((c) => pt(attr(c, "w")) * env.tf.sx);
    const tblPr = kid(tbl, "tblPr");
    const firstRow = attr(tblPr, "firstRow") === "1";
    const bandRow = attr(tblPr, "bandRow") === "1";
    const styleId = textOf(kid(tblPr, "tableStyleId")).trim();
    const noStyle = styleId === "{2D5ABB26-0587-4C30-8999-92F81FD0307C}" || styleId === "{5940675A-B579-460E-94D1-54222C63F5DA}";
    const accent = ctx.theme.colors.accent1 || "#4472C4";
    const rows = [];
    let merged = false;
    kids(tbl, "tr").forEach((tr, ri) => {
      const row = [];
      for (const tc of kids(tr, "tc")) {
        if (attr(tc, "hMerge") === "1" || attr(tc, "vMerge") === "1" || Number(attr(tc, "gridSpan")) > 1 || Number(attr(tc, "rowSpan")) > 1) merged = true;
        const tcPr = kid(tc, "tcPr");
        let fill = null;
        let explicit = false;
        if (tcPr && kid(tcPr, "solidFill")) { fill = colorOf(kid(tcPr, "solidFill"), ctx); explicit = true; }
        else if (tcPr && kid(tcPr, "noFill")) explicit = true;
        let header = false;
        if (!explicit && !noStyle && styleId) {
          if (firstRow && ri === 0) { fill = accent; header = true; }
          else if (bandRow) fill = tintColor(accent, (firstRow ? ri - 1 : ri) % 2 === 0 ? 0.6 : 0.8);
          else fill = tintColor(accent, 0.8);
        }
        const paras = this.readParas(kid(tc, "txBody"), ctx, {
          lsts: [env.defaultTextStyle].filter(Boolean), baseFont: ctx.theme.minor, defSize: 18,
          styleColor: header ? "#FFFFFF" : undefined,
        });
        if (header) for (const p of paras) for (const r of p.runs) if (r.b === undefined) r.b = true;
        const cell = { paras };
        if (fill) cell.fill = fill;
        row.push(cell);
      }
      rows.push(row);
    });
    if (merged) this.warn("Merged table cells are shown as separate cells.");
    const g = { ...geom };
    delete g.rot;
    const it = { type: "table", ...g, rows };
    if (colW.length) it.colW = colW.map((v) => Math.round(v * 100) / 100);
    return it;
  }

  async readGroup(n, env, items) {
    const xf = this.xfrmOf(n);
    if (!xf) { await this.readTree(n, { ...env, depth: env.depth + 1 }, items); return; }
    const chOff = xf.chOff, chExt = xf.chExt;
    const chx = Number(attr(chOff, "x")) || 0, chy = Number(attr(chOff, "y")) || 0;
    const ccx = Number(attr(chExt, "cx")) || xf.cx || 1, ccy = Number(attr(chExt, "cy")) || xf.cy || 1;
    const gsx = xf.cx && ccx ? xf.cx / ccx : 1;
    const gsy = xf.cy && ccy ? xf.cy / ccy : 1;
    const t = env.tf;
    const tf = { sx: t.sx * gsx, sy: t.sy * gsy, dx: t.dx + t.sx * (xf.x - chx * gsx), dy: t.dy + t.sy * (xf.y - chy * gsy) };
    if (xf.flipH || xf.flipV) this.warn("Flipped groups are shown unflipped.");
    const grpFill = this.fillOf(kid(n, "grpSpPr"), null, env.ctx).fill;
    const start = items.length;
    await this.readTree(n, { ...env, tf, depth: env.depth + 1, ctx: { ...env.ctx, grpFill } }, items);
    if (xf.rot) {
      // rotate every member around the group's centre (in slide points)
      const gx = pt(xf.x * t.sx + t.dx) + pt(xf.cx * t.sx) / 2;
      const gy = pt(xf.y * t.sy + t.dy) + pt(xf.cy * t.sy) / 2;
      const a = (xf.rot * Math.PI) / 180;
      for (let k = start; k < items.length; k++) {
        const it = items[k];
        if (typeof it.x !== "number") continue;
        const cx = it.x + it.w / 2 - gx, cy = it.y + it.h / 2 - gy;
        const nx = gx + cx * Math.cos(a) - cy * Math.sin(a);
        const ny = gy + cx * Math.sin(a) + cy * Math.cos(a);
        it.x = Math.round((nx - it.w / 2) * 100) / 100;
        it.y = Math.round((ny - it.h / 2) * 100) / 100;
        if (it.type !== "table") {
          const r = (((it.rot || 0) + xf.rot) % 360 + 360) % 360;
          if (r > 0.01) it.rot = Math.round(r * 100) / 100; else delete it.rot;
        }
      }
    }
  }

  // ---- backgrounds

  async background(bg, ctx) {
    if (!bg) return null;
    const bgPr = kid(bg, "bgPr");
    if (bgPr) {
      if (kid(bgPr, "solidFill")) return { color: colorOf(kid(bgPr, "solidFill"), ctx) };
      if (kid(bgPr, "gradFill")) {
        this.warn("Gradient backgrounds are shown as solid colors.");
        return { color: colorOf(findDesc(kid(bgPr, "gradFill"), "gs"), ctx) };
      }
      if (kid(bgPr, "pattFill")) return { color: colorOf(kid(kid(bgPr, "pattFill"), "bgClr"), ctx) };
      if (kid(bgPr, "blipFill")) {
        const blip = path(bgPr, "blipFill", "blip");
        const rid = blip ? attr(blip, "r:embed") || attr(blip, "embed") : null;
        const rel = rid ? ctx.rels.get(rid) : null;
        const img = rel && !rel.external ? await this.imageDataUrl(rel.target) : null;
        if (img && /png|jpeg|gif|bmp|webp|svg/.test(img.mime)) return { image: img.url };
        this.warn("Some background pictures are not shown.");
        return null;
      }
      if (kid(bgPr, "noFill")) return { color: "#FFFFFF" };
    }
    const ref = kid(bg, "bgRef");
    if (ref) {
      const idx = Number(attr(ref, "idx")) || 0;
      const color = colorOf(ref, ctx);
      const style = idx >= 1001 ? ctx.theme.bgFills[idx - 1001] : null;
      if (style && style.local === "gradFill") {
        this.warn("Gradient backgrounds are shown as solid colors.");
        const gs = findDesc(style, "gs");
        return { color: gs ? colorOf(gs, ctx, color) || color : color };
      }
      return { color };
    }
    return null;
  }
}

/** Run colors equal to what the renderer shows anyway are dropped (they then follow the theme). */
function dropDefaultColors(deck) {
  const strip = (paras, def) => {
    for (const p of paras || []) for (const r of p.runs || []) if (r.color && r.color === def) delete r.color;
  };
  for (const s of deck.slides) {
    for (const it of s.items) {
      if (!it) continue;
      if (it.type === "table") for (const row of it.rows || []) for (const c of row) strip(c.paras, textColorFor(it, deck, c.fill || null));
      else if (it.type === "shape") strip(it.paras, textColorFor(it, deck));
      else if (it.type === "text") strip(it.paras, textColorFor(it, deck, it.fill || null));
    }
  }
}

function sameFmt(a, b) {
  const keys = ["b", "i", "u", "s", "sup", "sub", "color", "size", "font", "link", "hl"];
  return keys.every((k) => a[k] === b[k]);
}

function placeholderItem(geom, label) {
  return {
    type: "shape", ...geom, shape: "rect", fill: "#EEF0F3", line: "#AAB2BD", lineW: 1, valign: "middle",
    paras: [{ align: "center", runs: [{ t: label, color: "#5F6B7A", size: 14 }] }],
  };
}

const LAYOUT_TYPE_MAP = {
  title: "title", obj: "content", tx: "content", secHead: "section", twoObj: "two", twoTxTwoObj: "two",
  twoColTx: "two", objTx: "two", txAndObj: "two", objAndTx: "two", picTx: "two", blank: "blank", titleOnly: "blank",
  vertTx: "content", objOnly: "content", cust: undefined,
};

function plainError(msg) { return new Error(msg); }

/** .pptx bytes -> { model, warnings } */
export async function pptxToModel(bytes) {
  const u8 = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes || []);
  if (!u8.length) throw plainError("This file is empty, so it cannot be opened as a presentation.");
  if (!(u8[0] === 0x50 && u8[1] === 0x4b)) {
    if (u8[0] === 0xd0 && u8[1] === 0xcf) throw plainError("This is an old PowerPoint 97-2003 (.ppt) file or another legacy Office file. Save it as .pptx in PowerPoint first, or open it with the backend.");
    throw plainError("This file is not a PowerPoint presentation (.pptx): it is not a valid Office package.");
  }
  const JSZip = getJSZip();
  let zip;
  try { zip = await JSZip.loadAsync(u8); } catch { throw plainError("This presentation is damaged: the file could not be unpacked (it is not a valid .pptx)."); }
  checkZipLimits(zip);
  const R = new Reader(zip);
  // find the main part
  let main = "ppt/presentation.xml";
  const top = await R.rels("");
  for (const r of top.values()) if (r.type.endsWith("/officeDocument")) { main = r.target; break; }
  const pres = await R.xml(main);
  if (!pres || pres.local !== "presentation") {
    if (zip.file("word/document.xml")) throw plainError("This is a Word document, not a PowerPoint presentation.");
    if (zip.file("xl/workbook.xml")) throw plainError("This is an Excel workbook, not a PowerPoint presentation.");
    throw plainError("This file is not a PowerPoint presentation (.pptx): it has no presentation part.");
  }
  const presRels = await R.rels(main);
  const sz = kid(pres, "sldSz");
  const size = { w: sz ? pt(attr(sz, "cx")) : 960, h: sz ? pt(attr(sz, "cy")) : 540 };
  if (!(size.w > 0 && size.h > 0)) { size.w = 960; size.h = 540; }
  const defaultTextStyle = kid(pres, "defaultTextStyle");

  // theme from the first master
  const masterId = kids(kid(pres, "sldMasterIdLst"), "sldMasterId")[0];
  const masterRel = masterId ? presRels.get(attr(masterId, "r:id") || attr(masterId, "id")) : null;
  const firstMaster = masterRel ? await R.master(masterRel.target) : null;
  const theme0 = firstMaster ? firstMaster.theme : await R.theme(null);
  const mctx0 = firstMaster ? { theme: theme0, clrMap: firstMaster.clrMap, rels: firstMaster.rels } : { theme: theme0, clrMap: {}, rels: new Map() };
  const themeText = colorOf({ children: [{ local: "schemeClr", attrs: { val: "tx1" }, children: [] }] }, mctx0) || "#000000";
  const mbg = firstMaster ? await R.background(firstMaster.bg, mctx0) : null;
  const deckTheme = {
    font: theme0.minor, headFont: theme0.major, color: themeText,
    accent: theme0.colors.accent1, bg: (mbg && mbg.color) || colorOf({ children: [{ local: "schemeClr", attrs: { val: "bg1" }, children: [] }] }, mctx0) || "#FFFFFF",
  };

  const slides = [];
  const ids = kids(kid(pres, "sldIdLst"), "sldId");
  if (ids.length > 3000) throw plainError("This presentation has too many slides to open in the browser.");
  for (const sid of ids) {
    const rel = presRels.get(attr(sid, "r:id") || attr(sid, "id"));
    if (!rel || rel.external) continue;
    const sp = rel.target;
    const sroot = await R.xml(sp);
    if (!sroot) { R.warn("Some slides are missing from the file and were skipped."); continue; }
    const srels = await R.rels(sp);
    const lRel = [...srels.values()].find((r) => r.type.endsWith("/slideLayout"));
    const layout = lRel ? await R.layout(lRel.target) : null;
    const master = layout ? layout.master : firstMaster;
    const ovr = path(sroot, "clrMapOvr", "overrideClrMapping");
    let clrMap = layout ? layout.clrMap : master ? master.clrMap : {};
    if (ovr) { clrMap = {}; for (const k in ovr.attrs) clrMap[k] = ovr.attrs[k]; }
    const theme = master ? master.theme : theme0;
    const sctx = { theme, clrMap, rels: srels, path: sp };
    const slide = { items: [] };
    const lt = layout ? LAYOUT_TYPE_MAP[layout.type] : undefined;
    if (lt) slide.layout = lt;
    if (attr(sroot, "show") === "0") slide.hidden = true;

    // background: slide -> layout -> master
    let bg = await R.background(path(sroot, "cSld", "bg"), sctx);
    if (!bg && layout) bg = await R.background(layout.bg, { theme, clrMap: layout.clrMap, rels: layout.rels });
    if (!bg && master) bg = await R.background(master.bg, { theme, clrMap: master.clrMap, rels: master.rels });
    if (bg && bg.color && bg.color !== deckTheme.bg) slide.bg = bg.color;
    if (bg && bg.image) slide.items.push({ type: "image", x: 0, y: 0, w: size.w, h: size.h, src: bg.image, alt: "Background picture" });

    const baseEnv = { tf: { sx: 1, sy: 1, dx: 0, dy: 0 }, depth: 0, layout, defaultTextStyle, themeText };
    const showMaster = attr(sroot, "showMasterSp") !== "0";
    if (showMaster && layout && master && layout.showMasterSp) {
      await R.readTree(master.spTree, { ...baseEnv, ctx: { theme, clrMap, rels: master.rels, path: master.path }, fromSlide: false }, slide.items);
    }
    if (showMaster && layout) {
      await R.readTree(layout.spTree, { ...baseEnv, ctx: { theme, clrMap, rels: layout.rels, path: layout.path }, fromSlide: false }, slide.items);
    }
    await R.readTree(path(sroot, "cSld", "spTree"), { ...baseEnv, ctx: sctx, fromSlide: true }, slide.items);

    // transition
    const TRANS = { fade: "fade", dissolve: "fade", cut: "none", push: "push", cover: "push", pull: "push", wipe: "wipe", split: "wipe", strips: "wipe", wedge: "wipe", wheel: "wipe", circle: "zoom", diamond: "wipe", plus: "zoom", blinds: "wipe", checker: "wipe", comb: "wipe", randomBar: "wipe", random: "fade", zoom: "zoom", newsflash: "zoom", reveal: "fade", morph: "fade", flash: "fade", prism: "zoom", ripple: "zoom", vortex: "zoom", conveyor: "push", pan: "push", gallery: "push", flythrough: "zoom", warp: "zoom", glitter: "fade", honeycomb: "zoom", shred: "wipe", switch: "push", flip: "push", ferris: "push", doors: "wipe", window: "wipe" };
    const trKind = (t) => { const k = t ? kids(t)[0] : null; return k ? k.local : ""; };
    const cands = [kid(sroot, "transition")];
    for (const ac of kids(sroot, "AlternateContent")) cands.push(kid(kid(ac, "Choice"), "transition"), kid(kid(ac, "Fallback"), "transition"));
    const trs = cands.filter(Boolean);
    if (trs.length) {
      const known = trs.find((t) => TRANS[trKind(t)]);
      if (known) slide.transition = TRANS[trKind(known)];
      else if (trs.some((t) => trKind(t))) { slide.transition = "fade"; R.warn("Some slide transitions are shown as simple fades."); }
    }
    if (findDesc(sroot, "timing") && findDesc(kid(sroot, "timing"), "par")) R.warn("Animations are not played here.");

    // notes
    const nRel = [...srels.values()].find((r) => r.type.endsWith("/notesSlide"));
    if (nRel) {
      const nroot = await R.xml(nRel.target);
      const tree = path(nroot, "cSld", "spTree");
      const lines = [];
      for (const s of kids(tree, "sp")) {
        const ph = phInfo(s);
        if (!ph || ph.type !== "body") continue;
        for (const p of kids(kid(s, "txBody"), "p")) {
          let t = "";
          for (const c of kids(p)) {
            if (c.local === "r" || c.local === "fld") t += textOf(kid(c, "t"));
            else if (c.local === "br") t += "\n";
          }
          lines.push(t);
        }
      }
      while (lines.length && !lines[lines.length - 1].trim()) lines.pop();
      const notes = lines.join("\n");
      if (notes.trim()) slide.notes = notes;
    }
    slides.push(slide);
  }

  const meta = {};
  const core = await R.xml("docProps/core.xml");
  if (core) {
    const t = textOf(kid(core, "title")).trim();
    const a = textOf(kid(core, "creator")).trim();
    if (t) meta.title = t;
    if (a) meta.author = a;
  }
  const warnings = [...R.warnings];
  if (warnings.length) meta.warnings = warnings;
  const model = normalizeDeck({ kind: "slides", meta, size, theme: deckTheme, slides });
  dropDefaultColors(model);
  return { model, warnings };
}

// ---------------------------------------------------------------------------
// writer
// ---------------------------------------------------------------------------

const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main";
const NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships";
const RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const CT = "application/vnd.openxmlformats-officedocument.presentationml";
const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NSDECL = `xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}"`;

// eslint-disable-next-line no-control-regex
const BAD_XML = /[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/g;

export function xmlEsc(s) {
  return String(s ?? "").replace(BAD_XML, "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
}

const hex6 = (c) => (hexColor(c) || "#000000").slice(1);
const srgb = (c) => `<a:srgbClr val="${hex6(c)}"/>`;
const solid = (c) => `<a:solidFill>${srgb(c)}</a:solidFill>`;

function rels(list) {
  return `${XML_HEAD}<Relationships xmlns="${NS_PKG}">${list.map((r) => `<Relationship Id="${r.id}" Type="${r.type}" Target="${xmlEsc(r.target)}"${r.external ? ' TargetMode="External"' : ""}/>`).join("")}</Relationships>`;
}

const SP_TREE_HEAD = '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';

function xfrm(it, tag = "a:xfrm") {
  const rot = it.rot ? ` rot="${Math.round(((it.rot % 360) + 360) % 360 * 60000)}"` : "";
  return `<${tag}${rot}><a:off x="${emu(it.x)}" y="${emu(it.y)}"/><a:ext cx="${Math.max(0, emu(it.w))}" cy="${Math.max(0, emu(it.h))}"/></${tag}>`;
}

function lvlStyles(sizes, opts = {}) {
  let s = "";
  for (let i = 0; i < 9; i++) {
    const sz = sizes[Math.min(i, sizes.length - 1)];
    const bullet = opts.bullets;
    const marL = bullet ? emu(i * 36 + 22) : emu(i * 36);
    const ind = bullet ? -emu(22) : 0;
    s += `<a:lvl${i + 1}pPr marL="${opts.flat ? 0 : marL}" indent="${opts.flat ? 0 : ind}" algn="${opts.algn || "l"}" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">`
      + (opts.lnSpc ? `<a:lnSpc><a:spcPct val="${opts.lnSpc}"/></a:lnSpc>` : "")
      + `<a:spcBef><a:spcPts val="${opts.spcBef || 0}"/></a:spcBef>`
      + (bullet ? '<a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/><a:buChar char="&#8226;"/>' : "<a:buNone/>")
      + `<a:defRPr sz="${Math.round(sz * 100)}" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="${opts.major ? "+mj-lt" : "+mn-lt"}"/><a:ea typeface="${opts.major ? "+mj-ea" : "+mn-ea"}"/><a:cs typeface="${opts.major ? "+mj-cs" : "+mn-cs"}"/></a:defRPr></a:lvl${i + 1}pPr>`;
  }
  return s;
}

const ANCHOR = { top: "t", middle: "ctr", bottom: "b" };

class Writer {
  constructor(deck) {
    this.deck = deck;
    this.theme = { ...DEFAULT_THEME, ...(deck.theme || {}) };
    this.media = new Map(); // src -> { path, ext }
    this.mediaN = 0;
    this.exts = new Set();
  }

  /** rPr XML for a run. */
  rPr(r, tag, ctx) {
    const a = ['lang="en-US"'];
    if (r.size) a.push(`sz="${Math.max(100, Math.min(400000, Math.round(r.size * 100)))}"`);
    if (r.b) a.push('b="1"');
    if (r.i) a.push('i="1"');
    if (r.u) a.push('u="sng"');
    if (r.s) a.push('strike="sngStrike"');
    if (r.sup) a.push('baseline="30000"');
    else if (r.sub) a.push('baseline="-25000"');
    a.push('dirty="0"');
    let kidsXml = "";
    const color = r.color || ctx.defColor;
    if (color) kidsXml += solid(color);
    if (r.hl) kidsXml += `<a:highlight>${srgb(r.hl)}</a:highlight>`;
    if (r.font) kidsXml += `<a:latin typeface="${xmlEsc(r.font)}"/><a:cs typeface="${xmlEsc(r.font)}"/>`;
    if (r.link && ctx.rels && /^(https?:|mailto:|ftp:|tel:)/i.test(r.link)) {
      let target = r.link;
      try { target = encodeURI(decodeURI(r.link)); } catch { target = encodeURI(r.link); }
      const id = ctx.rels.add(`${RT}/hyperlink`, target, true);
      kidsXml += `<a:hlinkClick r:id="${id}"/>`;
    }
    return kidsXml ? `<${tag} ${a.join(" ")}>${kidsXml}</${tag}>` : `<${tag} ${a.join(" ")}/>`;
  }

  paras(paras, ctx) {
    const list = Array.isArray(paras) && paras.length ? paras : [{ runs: [] }];
    return list.map((p) => {
      const level = Math.max(0, Math.min(8, p.level || 0));
      const pa = [];
      if (ctx.lists !== false) {
        pa.push(`marL="${p.bullet ? emu(level * 36 + 22) : emu(level * 36)}"`);
        pa.push(`indent="${p.bullet ? -emu(22) : 0}"`);
      }
      if (level) pa.push(`lvl="${level}"`);
      pa.push(`algn="${{ left: "l", center: "ctr", right: "r", justify: "just" }[p.align] || "l"}"`);
      let bu = "";
      if (p.bullet === "ul") bu = '<a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/><a:buChar char="&#8226;"/>';
      else if (p.bullet === "ol") bu = '<a:buFont typeface="+mj-lt"/><a:buAutoNum type="arabicPeriod"/>';
      else if (ctx.lists !== false) bu = "<a:buNone/>";
      const pPr = `<a:pPr ${pa.join(" ")}>${bu}</a:pPr>`;
      let runs = "";
      let lastSize = null;
      for (const r of p.runs || []) {
        if (!r) continue;
        if (r.size) lastSize = r.size;
        const t = String(r.t ?? "");
        if (!t) continue;
        const segs = t.split("\n");
        segs.forEach((seg, k) => {
          if (k > 0) runs += `<a:br>${this.rPr(r, "a:rPr", ctx)}</a:br>`;
          if (seg) runs += `<a:r>${this.rPr(r, "a:rPr", ctx)}<a:t>${xmlEsc(seg)}</a:t></a:r>`;
        });
      }
      const endSize = lastSize ? ` sz="${Math.round(lastSize * 100)}"` : "";
      return `<a:p>${pPr}${runs}<a:endParaRPr lang="en-US"${endSize} dirty="0"/></a:p>`;
    }).join("");
  }

  fillXml(fill) { return fill ? solid(fill) : "<a:noFill/>"; }

  lnXml(line, lineW) {
    if (!line) return "<a:ln><a:noFill/></a:ln>";
    return `<a:ln w="${Math.max(0, emu(lineW || 0.75))}">${solid(line)}</a:ln>`;
  }

  bodyPr(it, isPh) {
    const anchor = ANCHOR[it.valign] || (it.type === "shape" ? "ctr" : it.ph === "title" ? "ctr" : "t");
    if (isPh) return `<a:bodyPr anchor="${anchor}"><a:normAutofit/></a:bodyPr>`;
    return `<a:bodyPr wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="${anchor}"><a:noAutofit/></a:bodyPr>`;
  }

  async mediaFor(src) {
    if (this.media.has(src)) return this.media.get(src);
    const m = /^data:([^;,]+)(;[^,]*)?,(.*)$/s.exec(src || "");
    if (!m) { this.media.set(src, null); return null; }
    let bytes;
    try {
      bytes = /;base64/i.test(m[2] || "") ? base64ToBytes(m[3]) : new TextEncoder().encode(decodeURIComponent(m[3]));
    } catch { this.media.set(src, null); return null; }
    let mime = sniffMime(bytes, m[1].toLowerCase());
    let ext = MIME_EXT[mime] || null;
    if (!ext) { this.media.set(src, null); return null; }
    this.mediaN++;
    const p = `ppt/media/image${this.mediaN}.${ext}`;
    this.zip.file(p, bytes);
    this.exts.add(ext);
    const rec = { path: p, ext, mime };
    if (ext === "svg") {
      // PowerPoint shows SVG through an extension; a PNG fallback is required.
      this.mediaN++;
      const fp = `ppt/media/image${this.mediaN}.png`;
      this.zip.file(fp, base64ToBytes(TRANSPARENT_PNG));
      this.exts.add("png");
      rec.fallback = fp;
    }
    this.media.set(src, rec);
    return rec;
  }

  async slideXml(slide, index, phMap) {
    const relList = [];
    const ctxRels = {
      add: (type, target, external = false) => {
        const found = relList.find((r) => r.type === type && r.target === target && !!r.external === external);
        if (found) return found.id;
        const id = `rId${relList.length + 1}`;
        relList.push({ id, type, target, external });
        return id;
      },
    };
    ctxRels.add(`${RT}/slideLayout`, `../slideLayouts/slideLayout${phMap.layoutN}.xml`);
    const theme = this.theme;
    let id = 1;
    const shapes = [];
    let bodyN = 0;
    for (const it of slide.items || []) {
      if (!it || typeof it !== "object") continue;
      id++;
      if (it.type === "text") {
        let phXml = "";
        if (it.ph === "title" && phMap.title) phXml = phMap.title;
        else if (it.ph === "subtitle" && phMap.subtitle) phXml = phMap.subtitle;
        else if (it.ph === "body" && phMap.bodies[bodyN]) phXml = phMap.bodies[bodyN++];
        else if (it.ph === "subtitle" && phMap.bodies[bodyN]) phXml = phMap.bodies[bodyN++];
        const isPh = !!phXml;
        const name = isPh ? (it.ph === "title" ? "Title" : it.ph === "subtitle" ? "Subtitle" : "Content Placeholder") : "TextBox";
        const defColor = it.fill ? textColorFor(it, this.deck, it.fill) : null;
        const ctx = { rels: ctxRels, defColor: defColor && defColor !== hexColor(theme.color) ? defColor : null };
        shapes.push(`<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name} ${id - 1}"/><p:cNvSpPr${isPh ? "><a:spLocks noGrp=\"1\"/></p:cNvSpPr>" : ' txBox="1"/>'}<p:nvPr>${phXml}</p:nvPr></p:nvSpPr>`
          + `<p:spPr>${xfrm(it)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${isPh && !it.fill ? "" : this.fillXml(it.fill)}${isPh && !it.line ? "" : this.lnXml(it.line, it.lineW)}</p:spPr>`
          + `<p:txBody>${this.bodyPr(it, isPh)}<a:lstStyle/>${this.paras(it.paras, ctx)}</p:txBody></p:sp>`);
      } else if (it.type === "shape") {
        const shape = it.shape || "rect";
        const prst = shape;
        const isLine = shape === "line";
        const line = it.line || (isLine ? theme.color || "#000000" : null);
        const lineW = it.lineW || (isLine ? 1.5 : 0.75);
        const tc = textColorFor(it, this.deck);
        const ctx = { rels: ctxRels, defColor: tc !== hexColor(theme.color) ? tc : null };
        const hasParas = !isLine && Array.isArray(it.paras) && it.paras.length;
        shapes.push(`<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${isLine ? "Straight Connector" : "Shape"} ${id - 1}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>`
          + `<p:spPr>${xfrm(it)}<a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom>${isLine ? "<a:noFill/>" : this.fillXml(it.fill)}${this.lnXml(line, lineW)}</p:spPr>`
          + (hasParas ? `<p:txBody>${this.bodyPr({ ...it, valign: it.valign || "middle" }, false)}<a:lstStyle/>${this.paras(it.paras, ctx)}</p:txBody>` : "")
          + "</p:sp>");
      } else if (it.type === "image") {
        const med = await this.mediaFor(it.src);
        if (!med) {
          shapes.push(`<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Picture ${id - 1}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>${xfrm(it)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${solid("#EEF0F3")}${this.lnXml("#AAB2BD", 1)}</p:spPr>`
            + `<p:txBody>${this.bodyPr({ valign: "middle", type: "shape" }, false)}<a:lstStyle/>${this.paras([{ align: "center", runs: [{ t: it.alt || "[Picture]", color: "#5F6B7A", size: 14 }] }], { rels: ctxRels })}</p:txBody></p:sp>`);
          continue;
        }
        let blip;
        if (med.fallback) {
          const rid = ctxRels.add(`${RT}/image`, `../media/${med.fallback.split("/").pop()}`);
          const sid = ctxRels.add(`${RT}/image`, `../media/${med.path.split("/").pop()}`);
          blip = `<a:blip r:embed="${rid}"><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="${sid}"/></a:ext></a:extLst></a:blip>`;
        } else {
          const rid = ctxRels.add(`${RT}/image`, `../media/${med.path.split("/").pop()}`);
          blip = `<a:blip r:embed="${rid}"/>`;
        }
        shapes.push(`<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="Picture ${id - 1}" descr="${xmlEsc(it.alt || "")}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>`
          + `<p:blipFill>${blip}<a:stretch><a:fillRect/></a:stretch></p:blipFill>`
          + `<p:spPr>${xfrm(it)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`);
      } else if (it.type === "table") {
        const rows = (it.rows || []).filter((r) => Array.isArray(r));
        if (!rows.length) { id--; continue; }
        const ncol = Math.max(1, rows.reduce((m, r) => Math.max(m, r.length), 0));
        let colW = Array.isArray(it.colW) && it.colW.length === ncol && it.colW.every((v) => v > 0) ? it.colW : null;
        if (!colW) colW = Array.from({ length: ncol }, () => (it.w || 400) / ncol);
        const colE = colW.map((v) => Math.max(12700, emu(v)));
        const totalW = colE.reduce((a, b) => a + b, 0);
        const rowH = Math.max(12700, Math.round(emu(it.h || rows.length * 30) / rows.length));
        const border = (tag) => `<a:${tag} w="12700" cap="flat" cmpd="sng" algn="ctr">${solid("#BFBFBF")}<a:prstDash val="solid"/></a:${tag}>`;
        const trs = rows.map((r) => `<a:tr h="${rowH}">${Array.from({ length: ncol }, (_, ci) => {
          const c = r[ci] || { paras: [] };
          const tc = textColorFor(it, this.deck, c.fill || null);
          const ctx = { rels: ctxRels, defColor: tc !== hexColor(theme.color) ? tc : null, lists: true };
          return `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>${this.paras(c.paras, ctx)}</a:txBody><a:tcPr>${border("lnL")}${border("lnR")}${border("lnT")}${border("lnB")}${c.fill ? solid(c.fill) : "<a:noFill/>"}</a:tcPr></a:tc>`;
        }).join("")}</a:tr>`).join("");
        shapes.push(`<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="Table ${id - 1}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>`
          + `<p:xfrm><a:off x="${emu(it.x)}" y="${emu(it.y)}"/><a:ext cx="${totalW}" cy="${rowH * rows.length}"/></p:xfrm>`
          + `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>{2D5ABB26-0587-4C30-8999-92F81FD0307C}</a:tableStyleId></a:tblPr>`
          + `<a:tblGrid>${colE.map((w) => `<a:gridCol w="${w}"/>`).join("")}</a:tblGrid>${trs}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`);
      } else {
        id--; // unknown item types cannot be written
      }
    }
    const bg = hexColor(slide.bg);
    const bgXml = bg ? `<p:bg><p:bgPr>${solid(bg)}<a:effectLst/></p:bgPr></p:bg>` : "";
    const tr = { fade: "<p:fade/>", push: '<p:push dir="u"/>', wipe: '<p:wipe dir="r"/>', zoom: "<p:zoom/>" }[slide.transition];
    const trXml = tr ? `<p:transition spd="med">${tr}</p:transition>` : "";
    const xml = `${XML_HEAD}<p:sld ${NSDECL}${slide.hidden ? ' show="0"' : ""}><p:cSld>${bgXml}<p:spTree>${SP_TREE_HEAD}${shapes.join("")}</p:spTree></p:cSld>`
      + `<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>${trXml}</p:sld>`;
    return { xml, relList, addRel: ctxRels.add };
  }
}

const TRANSPARENT_PNG = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGNgYGD4DwABBAEAwS2OUAAAAABJRU5ErkJggg==";

/** Layout definitions written into every file (index = layout number - 1). */
function layoutDefs() {
  return [
    { id: "title", name: "Title Slide", type: "title" },
    { id: "content", name: "Title and Content", type: "obj" },
    { id: "section", name: "Section Header", type: "secHead" },
    { id: "two", name: "Two Content", type: "twoObj" },
    { id: "titleOnly", name: "Title Only", type: "titleOnly" },
    { id: "blank", name: "Blank", type: "blank" },
  ];
}

/** Which layout a slide uses and how its ph items map to the layout placeholders. */
function phMapFor(slide) {
  const items = (slide.items || []).filter((it) => it && it.type === "text");
  const hasTitle = items.some((it) => it.ph === "title");
  const bodies = items.filter((it) => it.ph === "body").length;
  const hasSub = items.some((it) => it.ph === "subtitle");
  let lay = slide.layout;
  if (!lay) lay = hasSub ? "title" : bodies >= 2 ? "two" : bodies === 1 ? "content" : hasTitle ? "titleOnly" : "blank";
  if (lay === "blank" && hasTitle) lay = "titleOnly";
  const TITLE = '<p:ph type="title"/>';
  switch (lay) {
    case "title": return { layoutN: 1, title: '<p:ph type="ctrTitle"/>', subtitle: '<p:ph type="subTitle" idx="1"/>', bodies: [] };
    case "content": return { layoutN: 2, title: TITLE, bodies: ['<p:ph idx="1"/>'] };
    case "section": return { layoutN: 3, title: TITLE, bodies: ['<p:ph type="body" idx="1"/>'] };
    case "two": return { layoutN: 4, title: TITLE, bodies: ['<p:ph sz="half" idx="1"/>', '<p:ph sz="half" idx="2"/>'] };
    case "titleOnly": return { layoutN: 5, title: TITLE, bodies: [] };
    default: return { layoutN: 6, title: null, bodies: [] };
  }
}

function phSp(id, name, ph, box, txBodyInner, spPrExtra = "") {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>${ph}</p:nvPr></p:nvSpPr>`
    + `<p:spPr>${box ? xfrm(box) : ""}${spPrExtra}</p:spPr><p:txBody>${txBodyInner}</p:txBody></p:sp>`;
}

const PROMPT_P = (t, lvl = 0) => `<a:p>${lvl ? `<a:pPr lvl="${lvl}"/>` : ""}<a:r><a:rPr lang="en-US"/><a:t>${t}</a:t></a:r></a:p>`;
const BODY_PROMPTS = PROMPT_P("Click to edit Master text styles") + PROMPT_P("Second level", 1) + PROMPT_P("Third level", 2) + PROMPT_P("Fourth level", 3) + PROMPT_P("Fifth level", 4);

function layoutXml(def, size) {
  let tree = "";
  const b = layoutBoxes(def.id === "titleOnly" ? "content" : def.id, size);
  const lst = (inner) => `<a:lstStyle>${inner}</a:lstStyle>`;
  if (def.id === "title") {
    tree += phSp(2, "Title 1", '<p:ph type="ctrTitle"/>', b.title, `<a:bodyPr anchor="b"><a:normAutofit/></a:bodyPr>${lst('<a:lvl1pPr algn="ctr"><a:defRPr sz="4400"/></a:lvl1pPr>')}${PROMPT_P("Click to edit Master title style")}`);
    tree += phSp(3, "Subtitle 2", '<p:ph type="subTitle" idx="1"/>', b.subtitle, `<a:bodyPr><a:normAutofit/></a:bodyPr>${lst(`<a:lvl1pPr marL="0" indent="0" algn="ctr"><a:buNone/><a:defRPr sz="2400"/></a:lvl1pPr>`)}${PROMPT_P("Click to edit Master subtitle style")}`);
  } else if (def.id === "content" || def.id === "two") {
    tree += phSp(2, "Title 1", '<p:ph type="title"/>', b.title, `<a:bodyPr/><a:lstStyle/>${PROMPT_P("Click to edit Master title style")}`);
    if (def.id === "content") {
      tree += phSp(3, "Content Placeholder 2", '<p:ph idx="1"/>', b.body, `<a:bodyPr><a:normAutofit/></a:bodyPr><a:lstStyle/>${BODY_PROMPTS}`);
    } else {
      tree += phSp(3, "Content Placeholder 2", '<p:ph sz="half" idx="1"/>', b.left, `<a:bodyPr><a:normAutofit/></a:bodyPr><a:lstStyle/>${BODY_PROMPTS}`);
      tree += phSp(4, "Content Placeholder 3", '<p:ph sz="half" idx="2"/>', b.right, `<a:bodyPr><a:normAutofit/></a:bodyPr><a:lstStyle/>${BODY_PROMPTS}`);
    }
  } else if (def.id === "section") {
    tree += phSp(2, "Title 1", '<p:ph type="title"/>', b.title, `<a:bodyPr anchor="b"><a:normAutofit/></a:bodyPr><a:lstStyle/>${PROMPT_P("Click to edit Master title style")}`);
    tree += phSp(3, "Text Placeholder 2", '<p:ph type="body" idx="1"/>', b.body, `<a:bodyPr><a:normAutofit/></a:bodyPr>${lst('<a:lvl1pPr marL="0" indent="0"><a:buNone/></a:lvl1pPr>')}${PROMPT_P("Click to edit Master text styles")}`);
  } else if (def.id === "titleOnly") {
    tree += phSp(2, "Title 1", '<p:ph type="title"/>', b.title, `<a:bodyPr/><a:lstStyle/>${PROMPT_P("Click to edit Master title style")}`);
  }
  return `${XML_HEAD}<p:sldLayout ${NSDECL} type="${def.type}" preserve="1"><p:cSld name="${def.name}"><p:spTree>${SP_TREE_HEAD}${tree}</p:spTree></p:cSld>`
    + "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>";
}

function masterXml(size, nLayouts) {
  const b = layoutBoxes("content", size);
  const tree = phSp(2, "Title Placeholder 1", '<p:ph type="title"/>', b.title, `<a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"><a:normAutofit/></a:bodyPr><a:lstStyle/>${PROMPT_P("Click to edit Master title style")}`)
    + phSp(3, "Text Placeholder 2", '<p:ph type="body" idx="1"/>', b.body, `<a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"><a:normAutofit/></a:bodyPr><a:lstStyle/>${BODY_PROMPTS}`);
  const ids = Array.from({ length: nLayouts }, (_, i) => `<p:sldLayoutId id="${2147483649 + i}" r:id="rId${i + 1}"/>`).join("");
  return `${XML_HEAD}<p:sldMaster ${NSDECL}><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree>${SP_TREE_HEAD}${tree}</p:spTree></p:cSld>`
    + '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
    + `<p:sldLayoutIdLst>${ids}</p:sldLayoutIdLst>`
    + `<p:txStyles><p:titleStyle>${lvlStyles([44], { major: true, lnSpc: 90000 }).replace(/<a:lvl[2-9]pPr[\s\S]*$/, "")}</p:titleStyle>`
    + `<p:bodyStyle>${lvlStyles([28, 24, 20, 20, 20], { bullets: true, spcBef: 1000 })}</p:bodyStyle>`
    + `<p:otherStyle>${lvlStyles([18], {})}</p:otherStyle></p:txStyles></p:sldMaster>`;
}

const DEFAULT_ACCENTS = ["#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47"];

function themeXml(theme, name = "Office Theme") {
  const accent1 = hexColor(theme.accent) || DEFAULT_ACCENTS[0];
  const accents = [accent1, ...DEFAULT_ACCENTS.slice(1)];
  const text = hexColor(theme.color) || "#000000";
  const bg = hexColor(theme.bg) || "#FFFFFF";
  const dark = luminanceOf(bg) < 0.4;
  const major = xmlEsc(theme.headFont || theme.font || "Calibri Light");
  const minor = xmlEsc(theme.font || "Calibri");
  const clr = (n, c) => `<a:${n}>${srgb(c)}</a:${n}>`;
  return `${XML_HEAD}<a:theme xmlns:a="${NS_A}" name="${xmlEsc(name)}"><a:themeElements><a:clrScheme name="Office Editor">`
    + clr("dk1", text) + clr("lt1", bg) + clr("dk2", dark ? "#E7E6E6" : "#44546A") + clr("lt2", dark ? "#3A3F4B" : "#E7E6E6")
    + accents.map((c, i) => clr(`accent${i + 1}`, c)).join("") + clr("hlink", dark ? "#6CB4FF" : "#0563C1") + clr("folHlink", dark ? "#C59BD6" : "#954F72")
    + `</a:clrScheme><a:fontScheme name="Office Editor"><a:majorFont><a:latin typeface="${major}"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>`
    + `<a:minorFont><a:latin typeface="${minor}"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>`
    + '<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    + '<a:solidFill><a:schemeClr val="phClr"><a:tint val="50000"/></a:schemeClr></a:solidFill>'
    + '<a:solidFill><a:schemeClr val="phClr"><a:shade val="80000"/></a:schemeClr></a:solidFill></a:fillStyleLst>'
    + '<a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>'
    + '<a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>'
    + '<a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst>'
    + "<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>"
    + '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/></a:schemeClr></a:solidFill>'
    + '<a:solidFill><a:schemeClr val="phClr"><a:shade val="90000"/></a:schemeClr></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements>'
    + "<a:objectDefaults/><a:extraClrSchemeLst/></a:theme>";
}

function luminanceOf(hex) {
  const h = hexColor(hex) || "#FFFFFF";
  const [r, g, b] = [1, 3, 5].map((i) => parseInt(h.slice(i, i + 2), 16) / 255);
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

function notesMasterXml() {
  const tree = phSp(2, "Slide Image Placeholder 1", '<p:ph type="sldImg" idx="2"/>', { x: 54, y: 54, w: 432, h: 243 }, "", '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="12700"><a:solidFill><a:prstClr val="black"/></a:solidFill></a:ln>').replace("<p:txBody></p:txBody>", "")
    + phSp(3, "Notes Placeholder 2", '<p:ph type="body" sz="quarter" idx="3"/>', { x: 54, y: 324, w: 432, h: 360 }, `<a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"/><a:lstStyle/>${PROMPT_P("Click to edit Master text styles")}`, '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>');
  let notesStyle = "";
  for (let i = 0; i < 9; i++) notesStyle += `<a:lvl${i + 1}pPr marL="${i * 457200}" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl${i + 1}pPr>`;
  return `${XML_HEAD}<p:notesMaster ${NSDECL}><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree>${SP_TREE_HEAD}${tree}</p:spTree></p:cSld>`
    + '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
    + `<p:notesStyle>${notesStyle}</p:notesStyle></p:notesMaster>`;
}

function notesSlideXml(notes) {
  const paras = String(notes).split("\n").map((l) => `<a:p>${l ? `<a:r><a:rPr lang="en-US" dirty="0"/><a:t>${xmlEsc(l)}</a:t></a:r>` : ""}<a:endParaRPr lang="en-US" dirty="0"/></a:p>`).join("");
  return `${XML_HEAD}<p:notes ${NSDECL}><p:cSld><p:spTree>${SP_TREE_HEAD}`
    + '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>'
    + `<p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>${paras}</p:txBody></p:sp>`
    + "</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>";
}

function isoNow() {
  return new Date().toISOString().replace(/\.\d{3}Z$/, "Z");
}

/** deck -> .pptx bytes */
export async function modelToPptx(input) {
  const JSZip = getJSZip();
  const deck = normalizeDeck(cloneDeep(input || {}));
  const zip = new JSZip();
  const W = new Writer(deck);
  W.zip = zip;
  const size = {
    w: Math.max(72, Math.min(4032, deck.size.w)),
    h: Math.max(72, Math.min(4032, deck.size.h)),
  };
  deck.size = size;
  zip.file("[Content_Types].xml", ""); // first entry; filled in at the end
  const overrides = [];
  const ov = (p, type) => overrides.push(`<Override PartName="/${p}" ContentType="${type}"/>`);

  // slides
  const presRels = [{ id: "rId1", type: `${RT}/slideMaster`, target: "slideMasters/slideMaster1.xml" }];
  const sldIds = [];
  const slides = deck.slides;
  const hasNotes = slides.some((s) => s.notes && s.notes.trim());
  let notesN = 0;
  for (let i = 0; i < slides.length; i++) {
    const s = slides[i];
    const map = phMapFor(s);
    const res = await W.slideXml(s, i, map);
    if (s.notes && s.notes.trim()) {
      notesN++;
      res.addRel(`${RT}/notesSlide`, `../notesSlides/notesSlide${notesN}.xml`);
      zip.file(`ppt/notesSlides/notesSlide${notesN}.xml`, notesSlideXml(s.notes));
      zip.file(`ppt/notesSlides/_rels/notesSlide${notesN}.xml.rels`, rels([
        { id: "rId1", type: `${RT}/notesMaster`, target: "../notesMasters/notesMaster1.xml" },
        { id: "rId2", type: `${RT}/slide`, target: `../slides/slide${i + 1}.xml` },
      ]));
      ov(`ppt/notesSlides/notesSlide${notesN}.xml`, `${CT}.notesSlide+xml`);
    }
    zip.file(`ppt/slides/slide${i + 1}.xml`, res.xml);
    zip.file(`ppt/slides/_rels/slide${i + 1}.xml.rels`, rels(res.relList));
    ov(`ppt/slides/slide${i + 1}.xml`, `${CT}.slide+xml`);
    const rid = `rId${presRels.length + 1}`;
    presRels.push({ id: rid, type: `${RT}/slide`, target: `slides/slide${i + 1}.xml` });
    sldIds.push(`<p:sldId id="${256 + i}" r:id="${rid}"/>`);
  }

  // master, layouts, theme
  const defs = layoutDefs();
  defs.forEach((d, i) => {
    zip.file(`ppt/slideLayouts/slideLayout${i + 1}.xml`, layoutXml(d, size));
    zip.file(`ppt/slideLayouts/_rels/slideLayout${i + 1}.xml.rels`, rels([{ id: "rId1", type: `${RT}/slideMaster`, target: "../slideMasters/slideMaster1.xml" }]));
    ov(`ppt/slideLayouts/slideLayout${i + 1}.xml`, `${CT}.slideLayout+xml`);
  });
  zip.file("ppt/slideMasters/slideMaster1.xml", masterXml(size, defs.length));
  zip.file("ppt/slideMasters/_rels/slideMaster1.xml.rels", rels([
    ...defs.map((_, i) => ({ id: `rId${i + 1}`, type: `${RT}/slideLayout`, target: `../slideLayouts/slideLayout${i + 1}.xml` })),
    { id: `rId${defs.length + 1}`, type: `${RT}/theme`, target: "../theme/theme1.xml" },
  ]));
  ov("ppt/slideMasters/slideMaster1.xml", `${CT}.slideMaster+xml`);
  zip.file("ppt/theme/theme1.xml", themeXml(W.theme));
  ov("ppt/theme/theme1.xml", "application/vnd.openxmlformats-officedocument.theme+xml");

  let notesMasterIdLst = "";
  if (hasNotes) {
    zip.file("ppt/notesMasters/notesMaster1.xml", notesMasterXml());
    zip.file("ppt/notesMasters/_rels/notesMaster1.xml.rels", rels([{ id: "rId1", type: `${RT}/theme`, target: "../theme/theme2.xml" }]));
    zip.file("ppt/theme/theme2.xml", themeXml({ ...DEFAULT_THEME }, "Notes Theme"));
    ov("ppt/notesMasters/notesMaster1.xml", `${CT}.notesMaster+xml`);
    ov("ppt/theme/theme2.xml", "application/vnd.openxmlformats-officedocument.theme+xml");
    const rid = `rId${presRels.length + 1}`;
    presRels.push({ id: rid, type: `${RT}/notesMaster`, target: "notesMasters/notesMaster1.xml" });
    notesMasterIdLst = `<p:notesMasterIdLst><p:notesMasterId r:id="${rid}"/></p:notesMasterIdLst>`;
  }
  for (const [p, t, target] of [
    ["ppt/presProps.xml", "presProps", "presProps.xml"], ["ppt/viewProps.xml", "viewProps", "viewProps.xml"],
    ["ppt/theme/theme1.xml", "theme", "theme/theme1.xml"], ["ppt/tableStyles.xml", "tableStyles", "tableStyles.xml"],
  ]) {
    void p;
    presRels.push({ id: `rId${presRels.length + 1}`, type: `${RT}/${t}`, target });
  }
  zip.file("ppt/presProps.xml", `${XML_HEAD}<p:presentationPr ${NSDECL}/>`);
  ov("ppt/presProps.xml", `${CT}.presProps+xml`);
  zip.file("ppt/viewProps.xml", `${XML_HEAD}<p:viewPr ${NSDECL}><p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr><p:gridSpacing cx="76200" cy="76200"/></p:viewPr>`);
  ov("ppt/viewProps.xml", `${CT}.viewProps+xml`);
  zip.file("ppt/tableStyles.xml", `${XML_HEAD}<a:tblStyleLst xmlns:a="${NS_A}" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>`);
  ov("ppt/tableStyles.xml", `${CT}.tableStyles+xml`);

  const cx = emu(size.w), cy = emu(size.h);
  const type = cx === 12192000 && cy === 6858000 ? "" : cx === 9144000 && cy === 6858000 ? ' type="screen4x3"' : ' type="custom"';
  let dts = "";
  for (let i = 0; i < 9; i++) dts += `<a:lvl${i + 1}pPr marL="${i * 457200}" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl${i + 1}pPr>`;
  zip.file("ppt/presentation.xml", `${XML_HEAD}<p:presentation ${NSDECL} saveSubsetFonts="1">`
    + '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'
    + notesMasterIdLst
    + `<p:sldIdLst>${sldIds.join("")}</p:sldIdLst>`
    + `<p:sldSz cx="${cx}" cy="${cy}"${type}/><p:notesSz cx="6858000" cy="9144000"/>`
    + `<p:defaultTextStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr>${dts}</p:defaultTextStyle></p:presentation>`);
  ov("ppt/presentation.xml", `${CT}.presentation.main+xml`);
  zip.file("ppt/_rels/presentation.xml.rels", rels(presRels));

  // package-level
  const meta = deck.meta || {};
  const now = isoNow();
  zip.file("docProps/core.xml", `${XML_HEAD}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">`
    + (meta.title ? `<dc:title>${xmlEsc(meta.title)}</dc:title>` : "")
    + (meta.author ? `<dc:creator>${xmlEsc(meta.author)}</dc:creator>` : "")
    + `<cp:lastModifiedBy>Office Editor</cp:lastModifiedBy><cp:revision>1</cp:revision>`
    + `<dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`);
  ov("docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml");
  const hidden = slides.filter((s) => s.hidden).length;
  const fmt = type === "" ? "Widescreen" : type.includes("4x3") ? "On-screen Show (4:3)" : "Custom";
  zip.file("docProps/app.xml", `${XML_HEAD}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">`
    + `<TotalTime>0</TotalTime><Words>0</Words><Application>Microsoft Office PowerPoint</Application><PresentationFormat>${fmt}</PresentationFormat>`
    + `<Paragraphs>0</Paragraphs><Slides>${slides.length}</Slides><Notes>${notesN}</Notes><HiddenSlides>${hidden}</HiddenSlides><MMClips>0</MMClips>`
    + "<ScaleCrop>false</ScaleCrop><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0000</AppVersion></Properties>");
  ov("docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml");
  zip.file("_rels/.rels", rels([
    { id: "rId1", type: `${RT}/officeDocument`, target: "ppt/presentation.xml" },
    { id: "rId2", type: "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", target: "docProps/core.xml" },
    { id: "rId3", type: `${RT}/extended-properties`, target: "docProps/app.xml" },
  ]));

  const defaults = { rels: "application/vnd.openxmlformats-package.relationships+xml", xml: "application/xml" };
  const IMG_CT = { png: "image/png", jpeg: "image/jpeg", jpg: "image/jpeg", gif: "image/gif", bmp: "image/bmp", tiff: "image/tiff", svg: "image/svg+xml", webp: "image/webp", ico: "image/x-icon", avif: "image/avif" };
  for (const e of ["png", "jpeg", "gif", "svg", ...W.exts]) defaults[e] = IMG_CT[e] || "application/octet-stream";
  zip.file("[Content_Types].xml", `${XML_HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`
    + Object.entries(defaults).map(([e, t]) => `<Default Extension="${e}" ContentType="${t}"/>`).join("")
    + overrides.join("") + "</Types>");
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE", compressionOptions: { level: 6 } });
}
