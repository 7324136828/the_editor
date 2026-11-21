/**
 * Local spreadsheet engine: SheetJS (window.XLSX) <-> office excel model.
 * Used when the Rust backend is disabled/unreachable. Reads xlsx/xlsm/xls/ods
 * (values, formulas, number formats, basic cell styles, merges, column
 * widths, hidden sheets, active sheet) and writes xlsx/xlsm/xls/ods/csv/tsv.
 *
 * Dates never shift with the timezone: SheetJS is used with cellDates:false
 * and serial numbers are converted with pure UTC arithmetic.
 */

import {
  usedRange, colToIndex, indexToCol, parseRange, addrToRC, rangeToRef, rcToAddr,
  isoToSerial, serialToIso, isDateFormat, isDateOnlyOrDateTimeFormat,
  csvBytesToModel, modelSheetToCsvBytes, sanitizeSheetName, BUILTIN_FORMATS,
  pxToChars, chartLayout, chartAnchorFromXY, chartXYFromAnchor, excelColPx, excelXOf, excelYOf,
  CHART_COLORS, CHART_TYPES, cellNumericValue, formatCell, MAX_COLS, MAX_ROWS,
} from "./model-excel.js";
import { stripXlfnPrefixes, addXlfnPrefixes, quoteSheetName } from "./formula.js";

const X = () => globalThis.XLSX;

export function xlsxAvailable() {
  const x = X();
  return !!(x && x.read && x.write);
}

// SheetJS numeric error codes <-> Excel error text.
const ERR_BY_CODE = { 0x00: "#NULL!", 0x07: "#DIV/0!", 0x0f: "#VALUE!", 0x17: "#REF!", 0x1d: "#NAME?", 0x24: "#NUM!", 0x2a: "#N/A", 0x2b: "#GETTING_DATA" };
const CODE_BY_ERR = Object.fromEntries(Object.entries(ERR_BY_CODE).map(([k, v]) => [v, Number(k)]));

const DATE1904_OFFSET = 1462;

function argbToHex(c) {
  if (!c) return null;
  const rgb = typeof c === "string" ? c : c.rgb;
  if (!rgb || !/^[0-9a-f]{6,8}$/i.test(rgb)) return null;
  return "#" + rgb.slice(-6).toUpperCase();
}

/** Map a style index (cellXfs) to our CellStyle using SheetJS's parsed Styles. */
function makeStyleResolver(Styles) {
  const cache = new Map();
  if (!Styles || !Styles.CellXf) return () => null;
  const defaultFont = Styles.Fonts && Styles.Fonts[0];
  const defColor = defaultFont && argbToHex(defaultFont.color);
  return (idx) => {
    if (cache.has(idx)) return cache.get(idx);
    const xf = Styles.CellXf[idx];
    let st = null;
    if (xf) {
      st = {};
      const font = Styles.Fonts && Styles.Fonts[xf.fontId];
      if (font) {
        if (font.bold) st.b = true;
        if (font.italic) st.i = true;
        if (font.underline) st.u = true;
        const col = argbToHex(font.color);
        if (col && col !== defColor && col !== "#000000") st.color = col;
      }
      const fill = Styles.Fills && Styles.Fills[xf.fillId];
      if (fill && fill.patternType === "solid") {
        const bg = argbToHex(fill.fgColor) || argbToHex(fill.bgColor);
        if (bg) st.bg = bg;
      }
      const al = xf.alignment;
      if (al) {
        if (["left", "center", "right"].includes(al.horizontal)) st.align = al.horizontal;
        if (al.wrapText === true || al.wrapText === "1" || al.wrapText === 1) st.wrap = true;
      }
      if (!Object.keys(st).length) st = null;
    }
    cache.set(idx, st);
    return st;
  };
}

function bytesToString(content) {
  if (content == null) return "";
  if (typeof content === "string") return content;
  return new TextDecoder("utf-8").decode(content instanceof Uint8Array ? content : new Uint8Array(content));
}

/** addr -> style index from raw sheet XML (SheetJS does not expose it per cell). */
function readStyleIndexes(xml) {
  const out = new Map();
  const re = /<c\b([^>]*?)\/?>/g;
  let m;
  while ((m = re.exec(xml))) {
    const attrs = m[1];
    const s = /\ss="(\d+)"/.exec(attrs);
    if (!s || s[1] === "0") continue;
    const r = /\sr="([A-Z]+\d+)"/.exec(attrs);
    if (r) out.set(r[1], Number(s[1]));
  }
  return out;
}

/** Number format code of cellXfs[idx] (undefined for General). */
function numFmtOfXf(wb, idx) {
  const xf = wb.Styles && wb.Styles.CellXf && wb.Styles.CellXf[idx];
  if (!xf || xf.numFmtId == null) return undefined;
  const id = Number(xf.numFmtId);
  const code = (wb.SSF && wb.SSF[id]) || BUILTIN_FORMATS[id];
  return code && !/^general$/i.test(code) ? String(code) : undefined;
}

/**
 * Cells that carry only formatting (a filled header slot, a column
 * pre-formatted as currency) have no value, so SheetJS skips them. Add them
 * as value-less cells so their style / number format survives a save.
 */
function addFormattedBlanks(cells, styleIdx, resolve, numFmtOf) {
  for (const [addr, idx] of styleIdx) {
    if (cells[addr]) continue;
    const st = resolve(idx);
    const z = numFmtOf(idx);
    if (!st && !z) continue;
    const cell = { t: "s", v: null };
    if (z) cell.z = z;
    if (st) cell.s = { ...st };
    cells[addr] = cell;
  }
}

/** Workbook paths of each sheet XML, in SheetNames order. */
function sheetPaths(wb) {
  const dir = wb.Directory && wb.Directory.sheets;
  if (!Array.isArray(dir)) return [];
  return dir.map((p) => String(p).replace(/^\//, ""));
}

function fileText(wb, path) {
  const f = wb.files && (wb.files[path] || wb.files["/" + path]);
  return f ? bytesToString(f.content) : null;
}

/**
 * SheetJS worksheet -> our cells map.
 * opts: { styleOf?: (addr)=>CellStyle|null, date1904?: boolean }
 */
export function sheetToCells(ws, opts = {}) {
  const XL = X();
  const cells = {};
  if (!ws || !ws["!ref"]) return cells;
  const offset = opts.date1904 ? DATE1904_OFFSET : 0;
  for (const addr of Object.keys(ws)) {
    if (addr[0] === "!") continue;
    const cell = ws[addr];
    if (!cell) continue;
    const f = cell.f != null && cell.f !== "" ? "=" + stripXlfnPrefixes(String(cell.f)) : undefined;
    let z = cell.z != null ? String(typeof cell.z === "number" ? (BUILTIN_FORMATS[cell.z] || "General") : cell.z) : undefined;
    if (z && /^general$/i.test(z)) z = undefined;
    let out = null;
    switch (cell.t) {
      case "n": {
        let v = cell.v;
        if (typeof v !== "number") v = Number(v);
        if (z && isDateFormat(z) && isDateOnlyOrDateTimeFormat(z) && v >= 1 && v < 2958466) {
          out = { t: "d", v: serialToIso(v + offset) };
        } else out = { t: "n", v };
        break;
      }
      case "s": out = { t: "s", v: cell.v == null ? "" : String(cell.v) }; break;
      case "b": out = { t: "b", v: !!cell.v }; break;
      case "e": out = { t: "e", v: (cell.w && /^#/.test(cell.w) ? cell.w : ERR_BY_CODE[cell.v]) || "#VALUE!" }; break;
      case "d": {
        // Only happens if someone reads with cellDates:true; use local parts (SheetJS builds local dates).
        const d = cell.v instanceof Date ? cell.v : new Date(cell.v);
        if (!isNaN(d)) {
          const pad = (n) => String(n).padStart(2, "0");
          const iso = `${String(d.getFullYear()).padStart(4, "0")}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
          const time = d.getHours() || d.getMinutes() || d.getSeconds() ? `T${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}` : "";
          out = { t: "d", v: iso + time };
        }
        break;
      }
      default: break; // "z" stubs
    }
    if (!out) {
      if (!f) continue;
      out = { t: "f", v: null };
    }
    if (f) out.f = f;
    if (z) out.z = z;
    const st = opts.styleOf ? opts.styleOf(addr) : null;
    if (st) out.s = { ...st };
    cells[XL ? XL.utils.encode_cell(XL.utils.decode_cell(addr)) : addr] = out;
  }
  return cells;
}

// ---- frozen panes, AutoFilter and charts from the raw package ----

const xmlUnesc = (s) => String(s).replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'")
  .replace(/&#(\d+);/g, (m, d) => String.fromCodePoint(+d)).replace(/&#x([0-9a-f]+);/gi, (m, h) => String.fromCodePoint(parseInt(h, 16))).replace(/&amp;/g, "&");
const attrOf = (tag, name) => { const m = new RegExp(`\\s(?:\\w+:)?${name}="([^"]*)"`).exec(tag); return m ? xmlUnesc(m[1]) : null; };

/** Frozen pane of a sheet XML: {rows, cols} or null. */
export function readFrozenPane(xml) {
  const m = /<(?:\w+:)?pane\b[^>]*>/.exec(xml || "");
  if (!m) return null;
  const state = attrOf(m[0], "state");
  if (state !== "frozen" && state !== "frozenSplit") return null;
  const rows = Math.max(0, Math.round(Number(attrOf(m[0], "ySplit")) || 0));
  const cols = Math.max(0, Math.round(Number(attrOf(m[0], "xSplit")) || 0));
  return rows || cols ? { rows: Math.min(rows, MAX_ROWS - 1), cols: Math.min(cols, MAX_COLS - 1) } : null;
}

function resolvePart(baseDir, target) {
  if (target.startsWith("/")) return target.slice(1);
  const parts = baseDir.split("/").filter(Boolean);
  for (const seg of target.split("/")) {
    if (seg === "..") parts.pop();
    else if (seg && seg !== ".") parts.push(seg);
  }
  return parts.join("/");
}

/** Relationships of a part: Map id -> resolved target path. */
function partRels(readText, path) {
  const dir = path.includes("/") ? path.slice(0, path.lastIndexOf("/")) : "";
  const name = path.slice(path.lastIndexOf("/") + 1);
  const xml = readText(`${dir ? dir + "/" : ""}_rels/${name}.rels`);
  const out = new Map();
  if (!xml) return out;
  for (const m of xml.matchAll(/<(?:\w+:)?Relationship\b[^>]*>/g)) {
    const id = attrOf(m[0], "Id"), t = attrOf(m[0], "Target");
    if (id && t && attrOf(m[0], "TargetMode") !== "External") out.set(id, resolvePart(dir, t));
  }
  return out;
}

/** "'My Sheet'!$A$1:$B$4" -> {sheet, r1, c1, r2, c2} or null. */
function parseChartRef(f) {
  const s = xmlUnesc(f).trim().replace(/^\(|\)$/g, "");
  const m = /^(?:'((?:[^']|'')+)'|([^!]+))!(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?)$/.exec(s);
  if (!m) return null;
  const rg = parseRange(m[3].replace(/\$/g, ""));
  return rg ? { sheet: m[1] != null ? m[1].replace(/''/g, "'") : m[2], ...rg } : null;
}

const TYPE_BY_TAG = {
  barChart: "column", bar3DChart: "column", lineChart: "line", line3DChart: "line", pieChart: "pie", pie3DChart: "pie",
  doughnutChart: "pie", ofPieChart: "pie", scatterChart: "scatter", areaChart: "area", area3DChart: "area",
};

/** Parse a chart part into {type, title?, range} (refs must be on `sheetName`), or null. */
export function parseChartXml(xml, sheetName) {
  if (!xml) return null;
  const plotAt = xml.search(/<(?:\w+:)?plotArea\b/);
  if (plotAt < 0) return null;
  const tm = /<(?:\w+:)?(barChart|bar3DChart|lineChart|line3DChart|pieChart|pie3DChart|doughnutChart|ofPieChart|scatterChart|areaChart|area3DChart)\b/.exec(xml.slice(plotAt));
  if (!tm) return null;
  let type = TYPE_BY_TAG[tm[1]];
  if (type === "column") {
    const dir = /<(?:\w+:)?barDir\b[^>]*val="(\w+)"/.exec(xml.slice(plotAt));
    if (dir && dir[1] === "bar") type = "bar";
  }
  const head = xml.slice(0, plotAt);
  let title = "";
  const tt = /<(?:\w+:)?title>([\s\S]*?)<\/(?:\w+:)?title>/.exec(head);
  if (tt) {
    title = [...tt[1].matchAll(/<a:t>([^<]*)<\/a:t>|<(?:\w+:)?t>([^<]*)<\/(?:\w+:)?t>/g)].map((m) => xmlUnesc(m[1] ?? m[2])).join("");
    if (!title) { const v = /<(?:\w+:)?v>([^<]*)<\/(?:\w+:)?v>/.exec(tt[1]); if (v) title = xmlUnesc(v[1]); }
  }
  const refs = [...xml.slice(plotAt).matchAll(/<(?:\w+:)?f>([^<]*)<\/(?:\w+:)?f>/g)].map((m) => parseChartRef(m[1]));
  if (!refs.length || refs.some((r) => !r || r.sheet.toLowerCase() !== String(sheetName).toLowerCase())) return null;
  const range = rangeToRef({
    r1: Math.min(...refs.map((r) => r.r1)), c1: Math.min(...refs.map((r) => r.c1)),
    r2: Math.max(...refs.map((r) => r.r2)), c2: Math.max(...refs.map((r) => r.c2)),
  });
  const out = { type, range };
  if (title.trim()) out.title = title.trim();
  return out;
}

const numTag = (xml, tag) => { const m = new RegExp(`<(?:\\w+:)?${tag}>\\s*(-?\\d+)\\s*<`).exec(xml); return m ? Number(m[1]) : 0; };

/** Charts of one worksheet: {charts, failed, other} (other: non-chart drawings present). */
function readSheetCharts(readText, sheetPath, sheetXml, sheet) {
  const res = { charts: [], failed: 0, other: false };
  const dm = /<(?:\w+:)?drawing\b[^>]*>/.exec(sheetXml || "");
  if (!dm) return res;
  const rid = attrOf(dm[0], "id");
  const dpath = rid && partRels(readText, sheetPath).get(rid);
  const dxml = dpath && readText(dpath);
  if (!dxml) return res;
  const drels = partRels(readText, dpath);
  const anchorRe = /<(?:\w+:)?(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b[\s\S]*?<\/(?:\w+:)?\1>/g;
  for (const am of dxml.matchAll(anchorRe)) {
    const a = am[0];
    const cm = /<(?:\w+:)?chart\b[^>]*\b(?:\w+:)?id="([^"]+)"[^>]*\/?>/.exec(a);
    if (!cm) { if (/<(?:\w+:)?(pic|sp|cxnSp|grpSp|graphicFrame)\b/.test(a)) res.other = true; continue; }
    if (res.charts.length >= 200) { res.failed++; continue; }
    const cpath = drels.get(cm[1]);
    const parsed = cpath ? parseChartXml(readText(cpath), sheet.name) : null;
    if (!parsed) { res.failed++; continue; }
    const from = /<(?:\w+:)?from>([\s\S]*?)<\/(?:\w+:)?from>/.exec(a);
    const to = /<(?:\w+:)?to>([\s\S]*?)<\/(?:\w+:)?to>/.exec(a);
    const beforeFrame = a.slice(0, Math.max(0, a.search(/<(?:\w+:)?graphicFrame\b/)));
    const ext = /<(?:\w+:)?ext\b[^>]*>/.exec(beforeFrame);
    let pos;
    if (from) {
      const f = from[1];
      const col = numTag(f, "col"), row = numTag(f, "row"), co = numTag(f, "colOff") / EMU, ro = numTag(f, "rowOff") / EMU;
      const xy = chartXYFromAnchor(sheet, col, co, row, ro);
      let w = 480, h = 288;
      if (to && am[1] === "twoCellAnchor") {
        const t = to[1];
        w = excelXOf(sheet, numTag(t, "col"), numTag(t, "colOff") / EMU) - excelXOf(sheet, col, co);
        h = excelYOf(numTag(t, "row"), numTag(t, "rowOff") / EMU) - excelYOf(row, ro);
      } else if (ext) {
        w = Number(attrOf(ext[0], "cx")) / EMU; h = Number(attrOf(ext[0], "cy")) / EMU;
      }
      pos = { ...xy, w, h };
    } else {
      const p = /<(?:\w+:)?pos\b[^>]*>/.exec(a);
      const px = p ? Number(attrOf(p[0], "x")) / EMU : 0, py = p ? Number(attrOf(p[0], "y")) / EMU : 0;
      const ac = excelColAt(sheet, px);
      const xy = chartXYFromAnchor(sheet, ac.col, ac.off, Math.floor(py / 20), py - Math.floor(py / 20) * 20);
      pos = { ...xy, w: ext ? Number(attrOf(ext[0], "cx")) / EMU : 480, h: ext ? Number(attrOf(ext[0], "cy")) / EMU : 288 };
    }
    const clampPx = (v, lo, d) => (isFinite(v) && v >= lo ? Math.round(v) : d);
    res.charts.push({ ...parsed, x: clampPx(pos.x, 0, 0), y: clampPx(pos.y, 0, 0), w: clampPx(pos.w, 40, 480), h: clampPx(pos.h, 30, 288) });
  }
  return res;
}

function workbookWarnings(wb, info = {}) {
  const warn = [];
  const keys = wb.keys || Object.keys(wb.files || {});
  const has = (re) => keys.some((k) => re.test(k));
  if (info.failedCharts) warn.push(`${info.failedCharts} chart${info.failedCharts === 1 ? "" : "s"} could not be read and will be lost if you save.`);
  else if (has(/xl\/chartsheets\//)) warn.push("Chart sheets are not shown and will not be kept when saving from this editor.");
  if (has(/xl\/media\//) || info.otherDrawings) warn.push("Images and drawings are not shown and will not be kept when saving.");
  if (has(/xl\/pivot(Tables|CacheDefinitions)?\//i)) warn.push("Pivot tables are shown as plain values; pivot definitions will not be kept.");
  if (has(/vbaProject\.bin$/)) warn.push("Macros (VBA) will not be kept when saving.");
  if (has(/xl\/comments\d*\.xml$/) || has(/threadedComments/)) warn.push("Cell comments are not shown and will not be kept.");
  if (wb.Workbook && Array.isArray(wb.Workbook.Names) && wb.Workbook.Names.some((n) => !/^_xlnm\./.test(n.Name || ""))) {
    warn.push("Defined names are not supported; formulas that use them keep their last saved values.");
  }
  let cf = false, dv = false;
  for (const p of sheetPaths(wb)) {
    const xml = fileText(wb, p) || "";
    if (/<conditionalFormatting\b/.test(xml)) cf = true;
    if (/<dataValidations\b/.test(xml)) dv = true;
  }
  if (cf) warn.push("Conditional formatting is not shown and will not be kept.");
  if (dv) warn.push("Data validation rules will not be kept.");
  return warn;
}

/** Column widths (characters) of a SheetJS worksheet, or null. */
function colWidthsOf(ws) {
  if (!Array.isArray(ws && ws["!cols"])) return null;
  const cw = {};
  ws["!cols"].forEach((col, ci) => {
    if (!col) return;
    let w = null;
    if (col.width != null) w = storedWidthToChars(Number(col.width));
    else if (col.wch != null) w = Number(col.wch);
    else if (col.wpx != null) w = pxToChars(col.wpx);
    if (col.hidden) w = 0; // hidden column (same as the backend)
    if (w != null && isFinite(w)) cw[indexToCol(ci)] = Math.round(w * 100) / 100;
  });
  return Object.keys(cw).length ? cw : null;
}

/** bytes -> our excel model. CSV/TSV (by name) use the native codec. */
export function bytesToModel(bytes, name) {
  const u8 = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  const lower = String(name || "").toLowerCase();
  if (/\.(csv|tsv|txt)$/.test(lower)) return csvBytesToModel(u8, name);
  const XL = X();
  const wb = XL.read(u8, {
    type: "array", cellDates: false, cellFormula: true, cellNF: true, cellStyles: true, bookFiles: true,
  });
  const date1904 = !!(wb.Workbook && wb.Workbook.WBProps && wb.Workbook.WBProps.date1904);
  const resolve = makeStyleResolver(wb.Styles);
  const paths = sheetPaths(wb);
  let failedCharts = 0, otherDrawings = false;
  const sheets = wb.SheetNames.map((sheetName, i) => {
    const ws = wb.Sheets[sheetName];
    let styleOf = null;
    const xml = paths[i] && wb.Styles && wb.Styles.CellXf && wb.Styles.CellXf.length > 1 ? fileText(wb, paths[i]) : null;
    if (xml) {
      const idx = readStyleIndexes(xml);
      styleOf = (addr) => (idx.has(addr) ? resolve(idx.get(addr)) : null);
    }
    const meta = wb.Workbook && wb.Workbook.Sheets && wb.Workbook.Sheets[i];
    const sheet = {
      name: sheetName,
      cells: sheetToCells(ws, { styleOf, date1904 }),
    };
    if (xml) addFormattedBlanks(sheet.cells, readStyleIndexes(xml), resolve, (i) => numFmtOfXf(wb, i));
    if (meta && meta.Hidden) sheet.hidden = true;
    const rawXml = paths[i] ? (xml ?? fileText(wb, paths[i])) : null;
    if (rawXml) {
      const fz = readFrozenPane(rawXml);
      if (fz) sheet.freeze = fz;
      const af = /<(?:\w+:)?autoFilter\b[^>]*>/.exec(rawXml);
      const afRef = (ws && ws["!autofilter"] && ws["!autofilter"].ref) || (af && attrOf(af[0], "ref"));
      const afRg = afRef && parseRange(String(afRef).replace(/\$/g, ""));
      if (afRg) sheet.autoFilter = rangeToRef(afRg);
      try {
        const ch = readSheetCharts((p) => fileText(wb, p), paths[i], rawXml, { ...sheet, colWidths: colWidthsOf(ws) });
        if (ch.charts.length) sheet.charts = ch.charts;
        failedCharts += ch.failed;
        if (ch.other) otherDrawings = true;
      } catch (e) {
        failedCharts++;
      }
    }
    if (Array.isArray(ws && ws["!merges"]) && ws["!merges"].length) {
      sheet.merges = ws["!merges"].map((m) => XL.utils.encode_range(m));
    }
    const cw = colWidthsOf(ws);
    if (cw) sheet.colWidths = cw;
    return sheet;
  });
  let active = Number(wb.Workbook && wb.Workbook.WBView && wb.Workbook.WBView[0] && wb.Workbook.WBView[0].activeTab) || 0;
  if (!sheets[active] || sheets[active].hidden) {
    const firstVisible = sheets.findIndex((s) => !s.hidden);
    active = firstVisible >= 0 ? firstVisible : 0;
  }
  const model = { kind: "excel", active, sheets };
  const warnings = workbookWarnings(wb, { failedCharts, otherDrawings });
  if (warnings.length) model.warnings = warnings;
  return model;
}

// ---------------------------------------------------------------------------
// Writing
// ---------------------------------------------------------------------------

function defaultDateFormat(iso) {
  return /T(?!00:00:00)/.test(String(iso)) ? "yyyy-mm-dd hh:mm:ss" : "yyyy-mm-dd";
}

// Column widths: the model (like Excel's UI and the backend) counts
// characters; <col width> stores characters + 5px padding at the default
// 7px digit width (Calibri 11). 14.7109375 stored = 14 characters.
const PAD_CHARS = 5 / 7;
export function storedWidthToChars(w) {
  if (!(w > 0) || !isFinite(w)) return 0;
  const chars = w >= 12 / 7 ? w - PAD_CHARS : (w * 7) / 12;
  return Math.round(chars * 100) / 100;
}
export function charsToStoredWidth(chars) {
  if (!(chars > 0)) return 0;
  return chars >= 1 ? chars + PAD_CHARS : (chars * 12) / 7;
}

/** Our cells map -> SheetJS worksheet. */
export function cellsToSheet(cells, colWidths, merges) {
  const XL = X();
  const ws = {};
  const range = usedRange(cells);
  ws["!ref"] = range ? XL.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: range.r2, c: range.c2 } }) : "A1:A1";
  for (const [addr, cell] of Object.entries(cells || {})) {
    if (!cell) continue;
    const js = {};
    let v = cell.v;
    let t = cell.t;
    if (t === "d") {
      const serial = isoToSerial(v);
      if (serial != null) {
        js.t = "n"; js.v = serial;
        js.z = cell.z && isDateFormat(cell.z) ? cell.z : defaultDateFormat(v);
      } else { js.t = "s"; js.v = String(v); }
    } else if (t === "n") {
      js.t = "n"; js.v = typeof v === "number" ? v : Number(v) || 0;
    } else if (t === "b") {
      js.t = "b"; js.v = !!v;
    } else if (t === "e") {
      js.t = "e"; js.v = CODE_BY_ERR[v] ?? 0x0f; js.w = String(v);
    } else if (t === "f") {
      js.t = "n"; js.v = 0;
    } else {
      if (v == null && !cell.f) continue; // format-only cell
      js.t = "s"; js.v = v == null ? "" : String(v);
    }
    if (cell.z && !js.z) js.z = cell.z;
    if (cell.f) js.f = addXlfnPrefixes(String(cell.f)).replace(/^=/, "");
    ws[addr] = js;
  }
  if (colWidths && Object.keys(colWidths).length) {
    ws["!cols"] = [];
    for (const [col, w] of Object.entries(colWidths)) {
      const idx = colToIndex(col);
      if (idx < 0) continue;
      const chars = Number(w) >= 0 ? Number(w) : 8.43;
      ws["!cols"][idx] = { width: charsToStoredWidth(chars) };
    }
  }
  if (Array.isArray(merges) && merges.length) {
    ws["!merges"] = merges.map((m) => parseRange(m)).filter(Boolean)
      .map((r) => ({ s: { r: r.r1, c: r.c1 }, e: { r: r.r2, c: r.c2 } }));
  }
  return ws;
}

// ---- styles.xml generation (SheetJS community does not write cell styles) ----

const xmlEsc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const BUILTIN_ID = Object.fromEntries(Object.entries(BUILTIN_FORMATS).map(([id, code]) => [code, Number(id)]));

function hexToArgb(hex) {
  return "FF" + String(hex).replace(/^#/, "").toUpperCase();
}

/** Build styles.xml for all (z, s) combinations used by the model. Returns {xml, indexOf(cell)}. */
function buildStyles(model) {
  const numFmts = new Map(); // code -> id
  let nextFmt = 164;
  const fonts = ['<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>'];
  const fontIdx = new Map([["", 0]]);
  const fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
  const fillIdx = new Map([["", 0]]);
  const xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'];
  const xfIdx = new Map([["0|0|0||", 0]]);

  const fmtId = (z) => {
    if (!z || /^general$/i.test(z)) return 0;
    if (BUILTIN_ID[z] != null) return BUILTIN_ID[z];
    if (!numFmts.has(z)) numFmts.set(z, nextFmt++);
    return numFmts.get(z);
  };
  const fontId = (s) => {
    if (!s || (!s.b && !s.i && !s.u && !s.color)) return 0;
    const key = `${s.b ? 1 : 0}${s.i ? 1 : 0}${s.u ? 1 : 0}${s.color || ""}`;
    if (!fontIdx.has(key)) {
      fontIdx.set(key, fonts.length);
      fonts.push(`<font>${s.b ? "<b/>" : ""}${s.i ? "<i/>" : ""}${s.u ? "<u/>" : ""}<sz val="11"/>${
        s.color && /^#[0-9a-f]{6}$/i.test(s.color) ? `<color rgb="${hexToArgb(s.color)}"/>` : '<color theme="1"/>'
      }<name val="Calibri"/><family val="2"/><scheme val="minor"/></font>`);
    }
    return fontIdx.get(key);
  };
  const fillId = (s) => {
    if (!s || !s.bg || !/^#[0-9a-f]{6}$/i.test(s.bg)) return 0;
    if (!fillIdx.has(s.bg)) {
      fillIdx.set(s.bg, fills.length);
      fills.push(`<fill><patternFill patternType="solid"><fgColor rgb="${hexToArgb(s.bg)}"/><bgColor indexed="64"/></patternFill></fill>`);
    }
    return fillIdx.get(s.bg);
  };
  const indexOf = (cell) => {
    if (!cell) return 0;
    let z = cell.z;
    if (cell.t === "d" && !(z && isDateFormat(z))) z = defaultDateFormat(cell.v);
    const s = cell.s;
    if (!z && !s) return 0;
    const nf = fmtId(z), fo = fontId(s), fi = fillId(s);
    const al = s && (s.align || s.wrap) ? `${s.align || ""}${s.wrap ? "w" : ""}` : "";
    const key = `${nf}|${fo}|${fi}|${al}|`;
    if (!xfIdx.has(key)) {
      xfIdx.set(key, xfs.length);
      const align = al ? `<alignment${s.align ? ` horizontal="${s.align}"` : ""}${s.wrap ? ' wrapText="1"' : ""}/>` : "";
      xfs.push(`<xf numFmtId="${nf}" fontId="${fo}" fillId="${fi}" borderId="0" xfId="0"${nf ? ' applyNumberFormat="1"' : ""}${fo ? ' applyFont="1"' : ""}${fi ? ' applyFill="1"' : ""}${align ? ' applyAlignment="1">' + align + "</xf>" : "/>"}`);
    }
    return xfIdx.get(key);
  };
  // Pre-register all cells so the xml is complete.
  for (const sheet of model.sheets || []) for (const cell of Object.values(sheet.cells || {})) indexOf(cell);
  const nf = [...numFmts].map(([code, id]) => `<numFmt numFmtId="${id}" formatCode="${xmlEsc(code)}"/>`).join("");
  const xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    (numFmts.size ? `<numFmts count="${numFmts.size}">${nf}</numFmts>` : "") +
    `<fonts count="${fonts.length}">${fonts.join("")}</fonts>` +
    `<fills count="${fills.length}">${fills.join("")}</fills>` +
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    `<cellXfs count="${xfs.length}">${xfs.join("")}</cellXfs>` +
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
    '<dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>' +
    "</styleSheet>";
  return { xml, indexOf };
}

/**
 * SheetJS does not write value-less cells; add `<c r=".." s=".."/>` for cells
 * that only carry formatting (style and/or number format), in row/column order.
 */
function injectFormatOnlyCells(xml, cells, indexOf) {
  const byRow = new Map(); // row number -> [{c, xml}]
  for (const [addr, cell] of Object.entries(cells)) {
    if (!cell || cell.v != null || cell.f) continue;
    const idx = indexOf(cell);
    if (!idx) continue;
    const rc = addrToRC(addr);
    if (!rc) continue;
    if (!byRow.has(rc.r + 1)) byRow.set(rc.r + 1, []);
    byRow.get(rc.r + 1).push({ c: rc.c, xml: `<c r="${addr}" s="${idx}"/>` });
  }
  if (!byRow.size) return xml;
  const m = /<sheetData\s*\/>|<sheetData>([\s\S]*?)<\/sheetData>/.exec(xml);
  if (!m) return xml;
  const rows = new Map(); // row number -> {attrs, cells:[{c, xml}]}
  const rowRe = /<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g;
  let rm;
  while ((rm = rowRe.exec(m[1] || ""))) {
    const n = Number((/\sr="(\d+)"/.exec(rm[1]) || [])[1]);
    const list = [];
    const cellRe = /<c\b[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g;
    let cm;
    while ((cm = cellRe.exec(rm[2] || ""))) {
      const a = /\sr="([A-Z]+\d+)"/.exec(cm[0]);
      const rc = a && addrToRC(a[1]);
      list.push({ c: rc ? rc.c : list.length, xml: cm[0] });
    }
    rows.set(n, { attrs: rm[1], cells: list });
  }
  for (const [n, extra] of byRow) {
    if (!rows.has(n)) rows.set(n, { attrs: ` r="${n}"`, cells: [] });
    const row = rows.get(n);
    for (const e of extra) if (!row.cells.some((x) => x.c === e.c)) row.cells.push(e);
    row.cells.sort((a, b) => a.c - b.c);
  }
  const body = [...rows.entries()].sort((a, b) => a[0] - b[0])
    .map(([, r]) => `<row${r.attrs.replace(/\sspans="[^"]*"/, "")}>${r.cells.map((x) => x.xml).join("")}</row>`).join("");
  return xml.slice(0, m.index) + `<sheetData>${body}</sheetData>` + xml.slice(m.index + m[0].length);
}

// ---- frozen panes and charts (SheetJS CE writes neither) ----

/** <sheetView> with a frozen <pane> for sheet.freeze, or null. */
function frozenSheetView(freeze) {
  const rows = Math.max(0, Math.min(MAX_ROWS - 1, Math.floor(Number(freeze && freeze.rows) || 0)));
  const cols = Math.max(0, Math.min(MAX_COLS - 1, Math.floor(Number(freeze && freeze.cols) || 0)));
  if (!rows && !cols) return null;
  const pane = rows && cols ? "bottomRight" : rows ? "bottomLeft" : "topRight";
  const tl = rcToAddr(rows, cols);
  return `<pane${cols ? ` xSplit="${cols}"` : ""}${rows ? ` ySplit="${rows}"` : ""} topLeftCell="${tl}" activePane="${pane}" state="frozen"/>` +
    `<selection pane="${pane}" activeCell="${tl}" sqref="${tl}"/>`;
}

const EMU = 9525; // EMU per pixel
const C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

function absRef(sheetName, r1, c1, r2, c2) {
  const a = `$${indexToCol(c1)}$${r1 + 1}`;
  const b = `$${indexToCol(c2)}$${r2 + 1}`;
  return `${quoteSheetName(sheetName)}!${a}${r1 === r2 && c1 === c2 ? "" : ":" + b}`;
}

function strCache(texts) {
  return `<c:strCache><c:ptCount val="${texts.length}"/>${texts.map((t, i) => `<c:pt idx="${i}"><c:v>${xmlEsc(t)}</c:v></c:pt>`).join("")}</c:strCache>`;
}
function numCache(values) {
  return `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${values.map((v, i) => (v == null ? "" : `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`)).join("")}</c:numCache>`;
}

/** DrawingML chart part for a model chart, or null when its range is invalid. */
export function chartXml(sheet, sheetName, chart) {
  const L = chartLayout(chart && chart.range);
  if (!L || L.r2 <= L.r1) return null;
  const type = CHART_TYPES.some(([t]) => t === chart.type) ? chart.type : "column";
  const cellAt = (r, c) => (sheet.cells || {})[rcToAddr(r, c)];
  const text = (cell) => (cell && cell.v != null ? formatCell(cell).text : "");
  const nums = (c) => {
    const out = [];
    for (let r = L.r1 + 1; r <= L.r2; r++) {
      const v = cellNumericValue(cellAt(r, c));
      out.push(typeof v === "number" && isFinite(v) ? v : null);
    }
    return out;
  };
  const texts = (c) => { const out = []; for (let r = L.r1 + 1; r <= L.r2; r++) out.push(text(cellAt(r, c))); return out; };
  let catXml = "";
  if (L.cat) {
    const ref = absRef(sheetName, L.r1 + 1, L.cat.c, L.r2, L.cat.c);
    const cv = nums(L.cat.c);
    const numeric = cv.every((v) => v != null);
    const tag = type === "scatter" ? "xVal" : "cat";
    catXml = numeric || type === "scatter"
      ? `<c:${tag}><c:numRef><c:f>${xmlEsc(ref)}</c:f>${numCache(cv)}</c:numRef></c:${tag}>`
      : `<c:${tag}><c:strRef><c:f>${xmlEsc(ref)}</c:f>${strCache(texts(L.cat.c))}</c:strRef></c:${tag}>`;
  }
  const valTag = type === "scatter" ? "yVal" : "val";
  const sers = L.series.slice(0, 255).map((s, i) => {
    const col = CHART_COLORS[i % CHART_COLORS.length].slice(1).toUpperCase();
    const tx = `<c:tx><c:strRef><c:f>${xmlEsc(absRef(sheetName, L.r1, s.c, L.r1, s.c))}</c:f>${strCache([text(cellAt(L.r1, s.c)) || `Series ${i + 1}`])}</c:strRef></c:tx>`;
    const val = `<c:${valTag}><c:numRef><c:f>${xmlEsc(absRef(sheetName, L.r1 + 1, s.c, L.r2, s.c))}</c:f>${numCache(nums(s.c))}</c:numRef></c:${valTag}>`;
    const fill = `<a:solidFill><a:srgbClr val="${col}"/></a:solidFill>`;
    let sp = "", marker = "", tail = "";
    if (type === "line") {
      sp = `<c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="${col}"/></a:solidFill><a:round/></a:ln></c:spPr>`;
      marker = `<c:marker><c:symbol val="circle"/><c:size val="5"/><c:spPr>${fill}</c:spPr></c:marker>`;
      tail = '<c:smooth val="0"/>';
    } else if (type === "scatter") {
      sp = '<c:spPr><a:ln w="19050"><a:noFill/></a:ln></c:spPr>';
      marker = `<c:marker><c:symbol val="circle"/><c:size val="7"/><c:spPr>${fill}</c:spPr></c:marker>`;
      tail = '<c:smooth val="0"/>';
    } else if (type !== "pie") {
      sp = `<c:spPr>${fill}</c:spPr>`;
    }
    return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${tx}${sp}${marker}${catXml}${val}${tail}</c:ser>`;
  }).join("");
  const axIds = '<c:axId val="500000001"/><c:axId val="500000002"/>';
  const catAx = (pos) => `<c:catAx><c:axId val="500000001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${pos}"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="500000002"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>`;
  const valAx = (id, cross, pos, grid, cb = "between") => `<c:valAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${pos}"/>${grid ? "<c:majorGridlines/>" : ""}<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="${cross}"/><c:crosses val="autoZero"/><c:crossBetween val="${cb}"/></c:valAx>`;
  let plot;
  switch (type) {
    case "pie":
      plot = `<c:pieChart><c:varyColors val="1"/>${sers}<c:firstSliceAng val="0"/></c:pieChart>`;
      break;
    case "line":
      plot = `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${sers}<c:marker val="1"/>${axIds}</c:lineChart>${catAx("b")}${valAx("500000002", "500000001", "l", true)}`;
      break;
    case "area":
      plot = `<c:areaChart><c:grouping val="standard"/><c:varyColors val="0"/>${sers}${axIds}</c:areaChart>${catAx("b")}${valAx("500000002", "500000001", "l", true)}`;
      break;
    case "scatter":
      plot = `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${sers}${axIds}</c:scatterChart>` +
        `${valAx("500000001", "500000002", "b", false, "midCat")}${valAx("500000002", "500000001", "l", true, "midCat")}`;
      break;
    default: {
      const bar = type === "bar";
      plot = `<c:barChart><c:barDir val="${bar ? "bar" : "col"}"/><c:grouping val="clustered"/><c:varyColors val="0"/>${sers}<c:gapWidth val="150"/>${axIds}</c:barChart>` +
        `${catAx(bar ? "l" : "b")}${valAx("500000002", "500000001", bar ? "b" : "l", true)}`;
    }
  }
  const title = String(chart.title || "").trim();
  const titleXml = title
    ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400" b="1"/></a:pPr><a:r><a:rPr lang="en-US" sz="1400" b="1"/><a:t>${xmlEsc(title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>`
    : '<c:autoTitleDeleted val="1"/>';
  const legend = type === "pie" || L.series.length > 1 ? `<c:legend><c:legendPos val="${type === "pie" ? "r" : "b"}"/><c:overlay val="0"/></c:legend>` : "";
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<c:chartSpace xmlns:c="${C_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}"><c:roundedCorners val="0"/>` +
    `<c:chart>${titleXml}<c:plotArea><c:layout/>${plot}</c:plotArea>${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>` +
    "</c:chartSpace>";
}

/** Excel px position -> {cell index, offset px} along columns. */
function excelColAt(sheet, px) {
  let c = 0, x = 0;
  while (c < MAX_COLS - 1 && x + excelColPx(sheet, c) <= px) { x += excelColPx(sheet, c); c++; }
  return { col: c, off: Math.round(px - x) };
}

/** <xdr:twoCellAnchor> for a model chart (rId = its chart relationship). */
function chartAnchorXml(sheet, chart, rId, id) {
  const a = chartAnchorFromXY(sheet, chart.x, chart.y);
  const w = Math.max(60, Math.round(Number(chart.w) || 480));
  const h = Math.max(40, Math.round(Number(chart.h) || 288));
  const x0 = excelXOf(sheet, a.col, a.dx), y0 = excelYOf(a.row, a.dy);
  const end = excelColAt(sheet, x0 + w);
  const endRow = Math.floor((y0 + h) / 20), endRowOff = Math.round(y0 + h - endRow * 20);
  const name = xmlEsc(String(chart.title || "").trim() || `Chart ${id - 1}`);
  return '<xdr:twoCellAnchor editAs="oneCell">' +
    `<xdr:from><xdr:col>${a.col}</xdr:col><xdr:colOff>${a.dx * EMU}</xdr:colOff><xdr:row>${a.row}</xdr:row><xdr:rowOff>${a.dy * EMU}</xdr:rowOff></xdr:from>` +
    `<xdr:to><xdr:col>${end.col}</xdr:col><xdr:colOff>${end.off * EMU}</xdr:colOff><xdr:row>${endRow}</xdr:row><xdr:rowOff>${endRowOff * EMU}</xdr:rowOff></xdr:to>` +
    `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="${name}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
    '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>' +
    `<a:graphic><a:graphicData uri="${C_NS}"><c:chart xmlns:c="${C_NS}" xmlns:r="${R_NS}" r:id="${rId}"/></a:graphicData></a:graphic>` +
    "</xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>";
}

const relsXml = (rels) => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.map(([id, type, target]) => `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`).join("")}</Relationships>`;

/**
 * Patch a SheetJS-written .xlsx: real cell styles, active tab and
 * fullCalcOnLoad (so Excel recalculates formulas on open), frozen panes and
 * charts (DrawingML parts referencing the sheet's cells).
 */
function postProcessXlsx(bytes, model, sheetOrder, names = []) {
  const XL = X();
  if (!XL.CFB) return bytes;
  const zip = XL.CFB.read(bytes, { type: "array" });
  const entryOf = (p) => {
    const i = zip.FullPaths.findIndex((x) => x.endsWith("/" + p));
    return i >= 0 ? zip.FileIndex[i] : null;
  };
  const read = (p) => { const e = entryOf(p); return e ? bytesToString(e.content) : null; };
  const write = (p, text) => { const e = entryOf(p); if (e) { e.content = new TextEncoder().encode(text); e.size = e.content.length; } };
  const add = (p, text) => XL.CFB.utils.cfb_add(zip, p, new TextEncoder().encode(text));
  const contentTypes = [];
  let drawingNo = 0, chartNo = 0;

  const styles = buildStyles(model);
  write("xl/styles.xml", styles.xml);

  sheetOrder.forEach((sheet, i) => {
    const p = `xl/worksheets/sheet${i + 1}.xml`;
    let xml = read(p);
    if (xml == null) return;
    const cells = sheet.cells || {};
    xml = xml.replace(/<c\b([^>]*?)(\/?)>/g, (m, attrs, selfClose) => {
      const r = /\sr="([A-Z]+\d+)"/.exec(attrs);
      if (!r) return m;
      const idx = styles.indexOf(cells[r[1]]);
      let a = attrs.replace(/\ss="\d+"/, "");
      if (idx) a += ` s="${idx}"`;
      return `<c${a}${selfClose}>`;
    });
    xml = injectFormatOnlyCells(xml, cells, styles.indexOf);
    const pane = frozenSheetView(sheet.freeze);
    if (pane) xml = xml.replace(/<sheetView\b([^>]*?)(?:\/>|>[\s\S]*?<\/sheetView>)/, (m, a) => `<sheetView${a}>${pane}</sheetView>`);
    const charts = (Array.isArray(sheet.charts) ? sheet.charts : []).slice(0, 200)
      .map((ch) => ({ ch, xml: chartXml(sheet, names[i] || sheet.name, ch) })).filter((x) => x.xml);
    if (charts.length) {
      const dn = ++drawingNo;
      const anchors = [], drels = [];
      charts.forEach(({ ch, xml: cx }, k) => {
        const cn = ++chartNo;
        add(`xl/charts/chart${cn}.xml`, cx);
        contentTypes.push([`/xl/charts/chart${cn}.xml`, "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"]);
        drels.push([`rId${k + 1}`, `${R_NS}/chart`, `../charts/chart${cn}.xml`]);
        anchors.push(chartAnchorXml(sheet, ch, `rId${k + 1}`, k + 2));
      });
      add(`xl/drawings/drawing${dn}.xml`, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
        `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="${A_NS}">${anchors.join("")}</xdr:wsDr>`);
      add(`xl/drawings/_rels/drawing${dn}.xml.rels`, relsXml(drels));
      contentTypes.push([`/xl/drawings/drawing${dn}.xml`, "application/vnd.openxmlformats-officedocument.drawing+xml"]);
      const relPath = `xl/worksheets/_rels/sheet${i + 1}.xml.rels`;
      const old = read(relPath);
      const rid = "rIdDrawing1";
      const rel = `<Relationship Id="${rid}" Type="${R_NS}/drawing" Target="../drawings/drawing${dn}.xml"/>`;
      if (old != null) write(relPath, old.replace("</Relationships>", rel + "</Relationships>"));
      else add(relPath, relsXml([[rid, `${R_NS}/drawing`, `../drawings/drawing${dn}.xml`]]));
      if (!/xmlns:r=/.test(xml.slice(0, 400))) xml = xml.replace(/<worksheet\b/, `<worksheet xmlns:r="${R_NS}"`);
      const tag = `<drawing r:id="${rid}"/>`;
      const after = /<(legacyDrawing|legacyDrawingHF|picture|oleObjects|controls|webPublishItems|tableParts|extLst)\b/.exec(xml);
      xml = after ? xml.slice(0, after.index) + tag + xml.slice(after.index) : xml.replace("</worksheet>", tag + "</worksheet>");
    }
    write(p, xml);
  });
  if (contentTypes.length) {
    const ct = read("[Content_Types].xml");
    if (ct != null) write("[Content_Types].xml", ct.replace("</Types>", contentTypes.map(([part, type]) => `<Override PartName="${part}" ContentType="${type}"/>`).join("") + "</Types>"));
  }

  let wbx = read("xl/workbook.xml");
  if (wbx != null) {
    const active = Math.max(0, Math.min(model.active || 0, sheetOrder.length - 1));
    if (/<bookViews>/.test(wbx)) {
      wbx = wbx.replace(/<workbookView\b([^>]*?)\/>/, (m, a) => `<workbookView${a.replace(/\s(activeTab|firstSheet)="\d+"/g, "")} activeTab="${active}"/>`);
    } else {
      wbx = wbx.replace("<sheets>", `<bookViews><workbookView activeTab="${active}"/></bookViews><sheets>`);
    }
    wbx = wbx.replace(/<calcPr\b[^>]*\/>/, "");
    const calc = '<calcPr calcId="191029" fullCalcOnLoad="1"/>';
    if (/<\/definedNames>/.test(wbx)) wbx = wbx.replace("</definedNames>", "</definedNames>" + calc);
    else wbx = wbx.replace("</sheets>", "</sheets>" + calc);
    write("xl/workbook.xml", wbx);
  }
  const out = XL.CFB.write(zip, { fileType: "zip", type: "array" });
  return out instanceof Uint8Array ? out : new Uint8Array(out);
}

/** our excel model -> bytes for the given format (xlsx default). */
export function modelToBytes(model, format = "xlsx") {
  const fmt = String(format || "xlsx").toLowerCase();
  if (fmt === "csv" || fmt === "tsv") return modelSheetToCsvBytes(model, model.active || 0, fmt);
  const XL = X();
  const wb = XL.utils.book_new();
  const sheets = model.sheets && model.sheets.length ? model.sheets : [{ name: "Sheet1", cells: {} }];
  const names = [];
  sheets.forEach((sheet) => {
    const name = sanitizeSheetName(sheet.name || "Sheet", names);
    names.push(name);
    const ws = cellsToSheet(sheet.cells, sheet.colWidths, sheet.merges);
    const af = sheet.autoFilter && parseRange(sheet.autoFilter);
    if (af) {
      ws["!autofilter"] = { ref: rangeToRef(af) };
      // the sheet's extent must cover the filter range
      const cur = parseRange(ws["!ref"]) || af;
      ws["!ref"] = rangeToRef({ r1: 0, c1: 0, r2: Math.max(cur.r2, af.r2), c2: Math.max(cur.c2, af.c2) });
    }
    XL.utils.book_append_sheet(wb, ws, name);
  });
  wb.Workbook = wb.Workbook || {};
  wb.Workbook.Sheets = sheets.map((s, i) => ({ name: names[i], Hidden: s.hidden ? 1 : 0 }));
  const bookType = fmt === "xls" ? "xls" : fmt === "ods" ? "ods" : fmt === "xlsm" ? "xlsm" : "xlsx";
  let out = XL.write(wb, { type: "array", bookType, bookSST: bookType !== "ods", compression: true });
  out = out instanceof Uint8Array ? out : new Uint8Array(out);
  if (bookType === "xlsx" || bookType === "xlsm") {
    try {
      out = postProcessXlsx(out, model, sheets, names);
    } catch (e) {
      if (typeof console !== "undefined") console.warn("xlsx style post-processing failed; saving without styles", e);
    }
  }
  return out;
}

