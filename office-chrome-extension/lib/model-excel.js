/**
 * Excel workbook model: addressing, number-format rendering, data-entry
 * semantics, date serials, sheet -> HTML table rendering and the CSV codec.
 * Shared contract with the Rust backend (src/office-backend/src/model.rs).
 *
 * Model shape:
 * {
 *   kind: "excel",
 *   active?: number,
 *   warnings?: [string],
 *   csvMeta?: { encoding, bom, delim, eol, trailingNewline, quoteAll }  // JS only, set by csv import
 *   sheets: [{
 *     name: string,
 *     hidden?: boolean,
 *     cells: { "A1": { v, t:"n"|"s"|"b"|"d"|"e"|"f", f?, z?, s? } },
 *     colWidths?: { "A": number },   // character units
 *     merges?: ["A1:B2"],
 *     freeze?: { rows, cols },       // frozen panes: rows above / columns left of the split
 *     autoFilter?: "A1:D20",         // AutoFilter range (first row = headers); the filter
 *                                    // criteria are view state of the grid, not saved
 *     charts?: [Chart]
 *   }]
 * }
 *
 * Chart = { type: "column"|"bar"|"line"|"pie"|"scatter"|"area", title?, range: "A1:C10",
 *           x, y, w, h }
 *   range: data INCLUDING a header row; first column = categories (x values for
 *   scatter), each further column = one series named by its header cell (a
 *   single-column range is one series without categories).
 *   x, y, w, h: px relative to the grid's top-left cell A1 at 100% zoom (grid
 *   columns: colWidths -> charsToPx, else 88px; rows 24px). In .xlsx files the
 *   anchor cell + offset is scaled to Excel's metrics (default column 64px, row
 *   20px); see chartAnchorFromXY / chartXYFromAnchor (same rules in xlsx.rs).
 *
 * Cell semantics:
 *   t "n" number, "s" string, "b" boolean, "d" NAIVE local ISO date/datetime
 *   ("YYYY-MM-DD" | "YYYY-MM-DDTHH:MM:SS", never timezone-shifted),
 *   "e" error text ("#DIV/0!"), "f" formula whose value is not known yet.
 *   `f` is the formula including the leading "=". `v` is the cached/computed
 *   value. `z` is an Excel number-format code; `s` = {b,i,u,color,bg,align,wrap}.
 */

export const MAX_COLS = 16384; // XFD
export const MAX_ROWS = 1048576;

// ---------------------------------------------------------------------------
// Addressing
// ---------------------------------------------------------------------------

export function colToIndex(col) {
  let n = 0;
  for (const ch of String(col).toUpperCase()) {
    n = n * 26 + (ch.charCodeAt(0) - 64);
  }
  return n - 1;
}

export function indexToCol(i) {
  let s = "";
  i += 1;
  while (i > 0) {
    const r = (i - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    i = Math.floor((i - 1) / 26);
  }
  return s;
}

export function addrToRC(addr) {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/.exec(String(addr).trim());
  if (!m) return null;
  const c = colToIndex(m[1]);
  const r = parseInt(m[2], 10) - 1;
  if (c < 0 || c >= MAX_COLS || r < 0 || r >= MAX_ROWS) return null;
  return { r, c };
}

export function rcToAddr(r, c) {
  return `${indexToCol(c)}${r + 1}`;
}

export function parseRange(ref) {
  const m = /^([A-Za-z$]+[\d$]+)(?::([A-Za-z$]+[\d$]+))?$/.exec(String(ref).trim());
  if (!m) return null;
  const a = addrToRC(m[1]);
  const b = m[2] ? addrToRC(m[2]) : a;
  if (!a || !b) return null;
  return {
    r1: Math.min(a.r, b.r), r2: Math.max(a.r, b.r),
    c1: Math.min(a.c, b.c), c2: Math.max(a.c, b.c),
  };
}

export function rangeToRef(rg) {
  const a = rcToAddr(rg.r1, rg.c1);
  const b = rcToAddr(rg.r2, rg.c2);
  return a === b ? a : `${a}:${b}`;
}

/** Used range of a cells map; null when empty. */
export function usedRange(cells) {
  let r1 = Infinity, r2 = -1, c1 = Infinity, c2 = -1;
  for (const addr of Object.keys(cells || {})) {
    const rc = addrToRC(addr);
    if (!rc) continue;
    if (rc.r < r1) r1 = rc.r;
    if (rc.r > r2) r2 = rc.r;
    if (rc.c < c1) c1 = rc.c;
    if (rc.c > c2) c2 = rc.c;
  }
  if (r2 < 0) return null;
  return { r1, r2, c1, c2 };
}

export function setCell(cells, addr, value) {
  const rc = addrToRC(addr);
  if (!rc) return cells;
  const key = rcToAddr(rc.r, rc.c);
  if (value == null) {
    delete cells[key];
    return cells;
  }
  cells[key] = value;
  return cells;
}

export function getCell(cells, addr) {
  return cells ? cells[String(addr).toUpperCase().replace(/\$/g, "")] : undefined;
}

// ---------------------------------------------------------------------------
// Sheet names
// ---------------------------------------------------------------------------

const BAD_SHEET_CHARS = /[\[\]:*?/\\]/;

/** Returns an error message for an invalid sheet name, or null if valid. */
export function validateSheetName(name, sheets = [], selfIndex = -1) {
  const n = String(name ?? "");
  if (!n.trim()) return "Sheet name cannot be empty.";
  if (n.length > 31) return "Sheet name cannot exceed 31 characters.";
  if (BAD_SHEET_CHARS.test(n)) return "Sheet name cannot contain any of: [ ] : * ? / \\";
  if (n.startsWith("'") || n.endsWith("'")) return "Sheet name cannot begin or end with an apostrophe.";
  if (n.toLowerCase() === "history") return "\"History\" is a reserved sheet name.";
  const lower = n.toLowerCase();
  if (sheets.some((s, i) => i !== selfIndex && String(s.name).toLowerCase() === lower)) {
    return `A sheet named "${n}" already exists (names are case-insensitive).`;
  }
  return null;
}

/**
 * Turn any string into a valid, unique (case-insensitive) sheet name:
 * replaces []:*?/\ with "_", trims apostrophes, truncates to 31 chars and
 * appends " (2)" etc. for duplicates. Falls back to "Sheet".
 */
export function sanitizeSheetName(name, existing = []) {
  let n = String(name ?? "").replace(/[\[\]:*?/\\]/g, "_").replace(/[\u0000-\u001f]/g, "").trim();
  n = n.replace(/^'+|'+$/g, "").trim();
  if (!n || n.toLowerCase() === "history") n = n ? `${n}_` : "Sheet";
  n = n.slice(0, 31);
  const taken = new Set(existing.map((e) => String(typeof e === "string" ? e : e.name).toLowerCase()));
  if (!taken.has(n.toLowerCase())) return n;
  for (let i = 2; i < 10000; i++) {
    const suffix = ` (${i})`;
    const cand = n.slice(0, 31 - suffix.length) + suffix;
    if (!taken.has(cand.toLowerCase())) return cand;
  }
  return n;
}

// ---------------------------------------------------------------------------
// Date serials (Excel 1900 system, incl. the fake 1900-02-29). All math is
// done in UTC so nothing ever shifts with the local timezone.
// ---------------------------------------------------------------------------

const DAY_MS = 86400000;
const EPOCH_MS = Date.UTC(1899, 11, 30);

function utcDays(y, m, d) {
  const dt = new Date(0);
  dt.setUTCFullYear(y, m - 1, d);
  dt.setUTCHours(0, 0, 0, 0);
  return Math.round((dt.getTime() - EPOCH_MS) / DAY_MS);
}

/** Excel serial for a calendar date (month/day overflow normalizes like DATE()). */
export function ymdToSerial(y, m, d) {
  const s = utcDays(y, m, d);
  return s <= 60 ? s - 1 : s;
}

/** Calendar parts of a serial: {y,m,d,H,M,S,ms,dow} (dow 0=Sunday). */
export function serialToParts(serial, fracDigits = 0) {
  let days = Math.floor(serial);
  const unit = Math.pow(10, fracDigits);
  let ticks = Math.round((serial - days) * 86400 * unit);
  if (ticks >= 86400 * unit) { days += 1; ticks -= 86400 * unit; }
  const secs = Math.floor(ticks / unit);
  const sub = ticks - secs * unit;
  const H = Math.floor(secs / 3600), M = Math.floor((secs % 3600) / 60), S = secs % 60;
  if (days === 60) return { y: 1900, m: 2, d: 29, H, M, S, sub, dow: 3 };
  if (days === 0) return { y: 1900, m: 1, d: 0, H, M, S, sub, dow: 6 };
  const dd = days < 60 ? days + 1 : days;
  const dt = new Date(EPOCH_MS + dd * DAY_MS);
  return {
    y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate(),
    H, M, S, sub, dow: days < 60 ? (dt.getUTCDay() + 6) % 7 : dt.getUTCDay(),
  };
}

const pad2 = (n) => String(n).padStart(2, "0");
const pad4 = (n) => String(n).padStart(4, "0");

/** "YYYY-MM-DD[THH:MM:SS]" -> serial (null if not a valid date string). */
export function isoToSerial(iso) {
  const m = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{1,2}):(\d{2})(?::(\d{2})(?:\.(\d+))?)?)?/.exec(String(iso || ""));
  if (!m) return null;
  const y = +m[1], mo = +m[2], d = +m[3];
  if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
  let s = ymdToSerial(y, mo, d);
  if (m[4] != null) {
    s += (+m[4] * 3600 + +m[5] * 60 + +(m[6] || 0) + (m[7] ? Number("0." + m[7]) : 0)) / 86400;
  }
  return s;
}

/** Serial -> naive ISO ("YYYY-MM-DD" when there is no time part). */
export function serialToIso(serial, forceTime = false) {
  const p = serialToParts(serial);
  const date = `${pad4(p.y)}-${pad2(p.m)}-${pad2(p.d)}`;
  if (!forceTime && p.H === 0 && p.M === 0 && p.S === 0) return date;
  return `${date}T${pad2(p.H)}:${pad2(p.M)}:${pad2(p.S)}`;
}

/** Number of days in month m (1-based) of year y. */
function daysInMonth(y, m) {
  return new Date(Date.UTC(y, m, 0)).getUTCDate();
}

// ---------------------------------------------------------------------------
// Number formats
// ---------------------------------------------------------------------------

const MONTHS = ["January", "February", "March", "April", "May", "June", "July",
  "August", "September", "October", "November", "December"];
const DAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
const COLORS = {
  black: "#000000", blue: "#0000ff", cyan: "#00ffff", green: "#008000",
  magenta: "#ff00ff", red: "#ff0000", white: "#ffffff", yellow: "#ffff00",
};

/** Built-in number formats by id (used by the xlsx readers). */
export const BUILTIN_FORMATS = {
  0: "General", 1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00", 9: "0%", 10: "0.00%",
  11: "0.00E+00", 12: "# ?/?", 13: "# ??/??", 14: "m/d/yyyy", 15: "d-mmm-yy", 16: "d-mmm",
  17: "mmm-yy", 18: "h:mm AM/PM", 19: "h:mm:ss AM/PM", 20: "h:mm", 21: "h:mm:ss",
  22: "m/d/yyyy h:mm", 37: "#,##0 ;(#,##0)", 38: "#,##0 ;[Red](#,##0)",
  39: "#,##0.00;(#,##0.00)", 40: "#,##0.00;[Red](#,##0.00)", 45: "mm:ss", 46: "[h]:mm:ss",
  47: "mmss.0", 48: "##0.0E+0", 49: "@",
};

const fmtCache = new Map();

function tokenizeSection(src) {
  const toks = [];
  let i = 0;
  const lit = (t) => {
    const last = toks[toks.length - 1];
    if (last && last.k === "lit") last.v += t; else toks.push({ k: "lit", v: t });
  };
  const sec = { toks, color: null, cond: null };
  while (i < src.length) {
    const ch = src[i];
    const rest = src.slice(i);
    if (ch === '"') {
      const j = src.indexOf('"', i + 1);
      lit(j < 0 ? src.slice(i + 1) : src.slice(i + 1, j));
      i = j < 0 ? src.length : j + 1;
    } else if (ch === "\\") { lit(src[i + 1] ?? ""); i += 2; }
    else if (ch === "_") { lit(" "); i += 2; }
    else if (ch === "*") { i += 2; } // fill character: ignored
    else if (ch === "[") {
      const j = src.indexOf("]", i);
      const body = j < 0 ? src.slice(i + 1) : src.slice(i + 1, j);
      i = j < 0 ? src.length : j + 1;
      let m;
      if ((m = /^(h+|m+|s+)$/i.exec(body))) toks.push({ k: "el", v: m[1].toLowerCase() });
      else if ((m = /^\$([^-]*)(-[0-9A-Fa-f]+)?$/.exec(body))) lit(m[1]);
      else if (COLORS[body.toLowerCase()]) sec.color = COLORS[body.toLowerCase()];
      else if ((m = /^(<=|>=|<>|<|>|=)\s*(-?[\d.]+)$/.exec(body))) sec.cond = { op: m[1], v: Number(m[2]) };
      // other bracket codes (Color10, locale ids, DBNum…) are ignored
    }
    else if (/^general/i.test(rest)) { toks.push({ k: "gen" }); i += 7; }
    else if (/^am\/pm/i.test(rest)) { toks.push({ k: "ampm", v: rest.slice(0, 5) }); i += 5; }
    else if (/^a\/p/i.test(rest)) { toks.push({ k: "ampm", v: rest.slice(0, 3) }); i += 3; }
    else if (ch === "0" || ch === "#" || ch === "?") { toks.push({ k: "ph", v: ch }); i++; }
    else if (ch === ".") { toks.push({ k: "dot" }); i++; }
    else if (ch === ",") { toks.push({ k: "comma" }); i++; }
    else if (ch === "%") { toks.push({ k: "pct" }); i++; }
    else if (ch === "@") { toks.push({ k: "at" }); i++; }
    else if (/^[eE][+-][0#?]/.test(rest)) { toks.push({ k: "exp", v: rest[1] }); i += 2; }
    else if (ch === "/") { toks.push({ k: "slash" }); i++; }
    else if (/[yYmMdDhHsS]/.test(ch)) {
      let j = i;
      while (j < src.length && src[j].toLowerCase() === ch.toLowerCase()) j++;
      toks.push({ k: "dt", v: src.slice(i, j).toLowerCase() });
      i = j;
    }
    else if (/[eE]/.test(ch) && !toks.some((t) => t.k === "ph")) { toks.push({ k: "dt", v: "yyyy" }); i++; }
    else if (/[bBgG]/.test(ch) && /^(b+|g+)/i.test(rest)) { i += /^(b+|g+)/i.exec(rest)[0].length; }
    else { lit(ch); i++; }
  }
  const hasPh = toks.some((t) => t.k === "ph");
  const hasDate = toks.some((t) => t.k === "el" || t.k === "ampm" ||
    (t.k === "dt" && (!hasPh || t.v[0] !== "m")));
  sec.kind = hasDate ? "date" : hasPh ? "num" : toks.some((t) => t.k === "at") ? "text"
    : toks.some((t) => t.k === "gen") ? "gen" : "lit";
  if (sec.kind === "date") resolveMinutes(toks);
  return sec;
}

/** "m"/"mm" means minutes right after hours or right before seconds. */
function resolveMinutes(toks) {
  const dts = toks.filter((t) => t.k === "dt" || t.k === "el");
  for (let i = 0; i < dts.length; i++) {
    const t = dts[i];
    if (t.k !== "dt" || t.v[0] !== "m" || t.v.length > 2) continue;
    const prev = dts[i - 1], next = dts[i + 1];
    if ((prev && (prev.v[0] === "h")) || (next && next.v[0] === "s")) t.min = true;
  }
}

function splitSections(z) {
  const out = [];
  let cur = "", q = false;
  for (let i = 0; i < z.length; i++) {
    const ch = z[i];
    if (ch === '"') q = !q;
    if (ch === "\\" && !q) { cur += ch + (z[i + 1] ?? ""); i++; continue; }
    if (ch === ";" && !q) { out.push(cur); cur = ""; continue; }
    cur += ch;
  }
  out.push(cur);
  return out;
}

function parseFormat(z) {
  let f = fmtCache.get(z);
  if (!f) {
    f = splitSections(z).map(tokenizeSection);
    if (fmtCache.size > 500) fmtCache.clear();
    fmtCache.set(z, f);
  }
  return f;
}

/** True if the format code shows dates/times. */
export function isDateFormat(z) {
  if (!z || /^general$/i.test(z)) return false;
  return parseFormat(String(z))[0].kind === "date";
}

/** True if the format shows a calendar date (not just a time of day). */
export function isDateOnlyOrDateTimeFormat(z) {
  if (!isDateFormat(z)) return false;
  return parseFormat(String(z))[0].toks.some((t) => t.k === "dt" && (t.v[0] === "y" || t.v[0] === "d" || (t.v[0] === "m" && !t.min)));
}

/** Round half away from zero to `d` decimals without binary artefacts. */
export function roundDecimal(x, d) {
  if (!isFinite(x)) return x;
  if (d > 15) return x;
  const sign = x < 0 ? -1 : 1;
  const a = Math.abs(x);
  const s = Math.round(Number(`${a}e${d}`));
  let r = Number(`${s}e-${d}`);
  if (!isFinite(r) || isNaN(r)) r = Math.round(a * Math.pow(10, d)) / Math.pow(10, d);
  return sign * r;
}

function fixedString(a, d) {
  // a >= 0, already rounded.
  if (a >= 1e21) return a.toLocaleString("en-US", { useGrouping: false, maximumFractionDigits: d, minimumFractionDigits: d });
  return a.toFixed(d);
}

/**
 * Excel "General" display: integers up to 15 digits shown in full, other
 * values with up to `digits` significant digits, scientific for very
 * large/small magnitudes.
 */
export function formatGeneral(v, digits = 11) {
  if (typeof v !== "number") return String(v ?? "");
  if (!isFinite(v)) return v > 0 ? "#NUM!" : v < 0 ? "-#NUM!" : "#NUM!";
  if (v === 0) return "0";
  const a = Math.abs(v);
  if (a >= 1e15 || a < 1e-9) return expString(v, 5);
  if (Number.isInteger(v)) return String(v);
  let s = String(parseFloat(v.toPrecision(digits)));
  if (/e/i.test(s)) {
    // e.g. 1.5e-7: write it out in plain decimal.
    const p = parseFloat(v.toPrecision(digits));
    s = p.toFixed(Math.min(20, Math.max(0, -Math.floor(Math.log10(Math.abs(p))) + digits)));
    s = s.replace(/(\.\d*?)0+$/, "$1").replace(/\.$/, "");
  }
  return s;
}

function expString(v, mantDecimals) {
  const d = Math.max(0, mantDecimals);
  let [m, e] = v.toExponential(d).split("e");
  if (m.includes(".")) m = m.replace(/0+$/, "").replace(/\.$/, "");
  const en = Number(e);
  return `${m}E${en < 0 ? "-" : "+"}${pad2(Math.abs(en))}`;
}

/** General formatting with full (15 digit) precision, used for CSV/clipboard/formula bar. */
export function formatNumberFull(v) {
  if (typeof v !== "number") return String(v ?? "");
  if (!isFinite(v)) return "#NUM!";
  if (Number.isInteger(v) && Math.abs(v) < 1e21) return String(v);
  const s = String(parseFloat(v.toPrecision(15)));
  return s;
}

function pickSection(secs, v) {
  const conds = secs.some((s) => s.cond);
  if (conds) {
    const test = (c, x) => c.op === "<" ? x < c.v : c.op === ">" ? x > c.v : c.op === "<=" ? x <= c.v
      : c.op === ">=" ? x >= c.v : c.op === "=" ? x === c.v : x !== c.v;
    for (let i = 0; i < Math.min(secs.length, 3); i++) {
      const s = secs[i];
      if (s.kind === "text") continue;
      if (!s.cond) return { sec: s, abs: i > 0 && v < 0 };
      if (test(s.cond, v)) return { sec: s, abs: s.cond.op === "<" || s.cond.op === "<=" ? v < 0 : false };
    }
    return { sec: secs[0], abs: false };
  }
  const numeric = secs.filter((s, i) => i < 3);
  if (v > 0 || (v === 0 && numeric.length < 3) || (v < 0 && numeric.length < 2)) {
    return { sec: numeric[0], abs: false };
  }
  if (v < 0) return { sec: numeric[1], abs: true };
  return { sec: numeric[2], abs: false };
}

/**
 * Format a value with an Excel number-format code.
 * @returns {{text:string, color:string|null}}
 */
export function formatValueEx(v, z) {
  const zz = z == null ? "" : String(z);
  if (typeof v === "boolean") return { text: v ? "TRUE" : "FALSE", color: null };
  if (v == null) return { text: "", color: null };
  if (!zz || /^general$/i.test(zz)) {
    return { text: typeof v === "number" ? formatGeneral(v) : String(v), color: null };
  }
  const secs = parseFormat(zz);
  if (typeof v !== "number") {
    const s = String(v);
    const ts = secs.length >= 4 ? secs[3] : secs.find((x) => x.kind === "text");
    if (!ts) return { text: s, color: null };
    let out = "";
    for (const t of ts.toks) out += t.k === "at" ? s : t.k === "lit" ? t.v : "";
    return { text: out, color: ts.color };
  }
  if (!isFinite(v)) return { text: "#NUM!", color: null };
  const { sec, abs } = pickSection(secs, v);
  if (!sec) return { text: formatGeneral(v), color: null };
  const x = abs ? Math.abs(v) : v;
  let text;
  switch (sec.kind) {
    case "date": text = renderDate(sec.toks, x); break;
    case "num": text = renderNumber(sec.toks, x); break;
    case "gen": {
      text = "";
      for (const t of sec.toks) text += t.k === "gen" ? formatGeneral(abs ? Math.abs(x) : x) : t.k === "lit" ? t.v : "";
      break;
    }
    case "text": case "lit": default: {
      text = "";
      for (const t of sec.toks) text += t.k === "lit" ? t.v : t.k === "at" ? formatGeneral(x) : "";
      if (sec.kind === "lit" && secs.length === 1 && x < 0) text = "-" + text;
      break;
    }
  }
  return { text, color: sec.color };
}

export function formatValue(v, z) {
  return formatValueEx(v, z).text;
}

function renderDate(toks, serial) {
  if (serial < 0) return "#".repeat(8);
  let fracDigits = 0;
  for (let i = 0; i < toks.length; i++) {
    if (toks[i].k === "dot" && toks[i - 1] && (toks[i - 1].v || "")[0] === "s") {
      let n = 0;
      while (toks[i + 1 + n] && toks[i + 1 + n].k === "ph") n++;
      fracDigits = Math.max(fracDigits, Math.min(3, n));
    }
  }
  const p = serialToParts(serial, fracDigits);
  const ampm = toks.some((t) => t.k === "ampm");
  let out = "";
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    switch (t.k) {
      case "lit": out += t.v; break;
      case "dt": {
        const c = t.v[0], n = t.v.length;
        if (c === "y") out += n <= 2 ? pad2(p.y % 100) : pad4(p.y);
        else if (c === "m" && t.min) out += n === 1 ? String(p.M) : pad2(p.M);
        else if (c === "m") {
          out += n === 1 ? String(p.m) : n === 2 ? pad2(p.m) : n === 3 ? MONTHS[p.m - 1].slice(0, 3)
            : n === 5 ? MONTHS[p.m - 1][0] : MONTHS[p.m - 1];
        } else if (c === "d") {
          out += n === 1 ? String(p.d) : n === 2 ? pad2(p.d) : n === 3 ? DAYS[p.dow].slice(0, 3) : DAYS[p.dow];
        } else if (c === "h") {
          let h = p.H;
          if (ampm) { h = h % 12; if (h === 0) h = 12; }
          out += n === 1 ? String(h) : pad2(h);
        } else if (c === "s") out += n === 1 ? String(p.S) : pad2(p.S);
        break;
      }
      case "el": {
        const total = serial * 86400;
        const c = t.v[0];
        const val = c === "h" ? Math.floor(total / 3600) : c === "m" ? Math.floor(total / 60) : Math.round(total);
        out += String(val).padStart(t.v.length, "0");
        break;
      }
      case "ampm": {
        const pm = p.H >= 12;
        const lower = t.v[0] === t.v[0].toLowerCase();
        let s = t.v.length === 3 ? (pm ? "P" : "A") : (pm ? "PM" : "AM");
        out += lower ? s.toLowerCase() : s;
        break;
      }
      case "dot": {
        if (fracDigits && toks[i + 1] && toks[i + 1].k === "ph") {
          let n = 0;
          while (toks[i + 1 + n] && toks[i + 1 + n].k === "ph") n++;
          out += "." + String(p.sub).padStart(fracDigits, "0").slice(0, n);
          i += n;
        } else out += ".";
        break;
      }
      case "ph": out += t.v === "0" ? "0" : ""; break;
      case "comma": out += ","; break;
      case "slash": out += "/"; break;
      case "pct": out += "%"; break;
      default: break;
    }
  }
  return out;
}

function renderNumber(toks, x) {
  const neg = x < 0;
  let a = Math.abs(x);
  const pctCount = toks.filter((t) => t.k === "pct").length;
  a *= Math.pow(100, pctCount);
  const firstPh = toks.findIndex((t) => t.k === "ph");
  let lastPh = -1;
  toks.forEach((t, i) => { if (t.k === "ph") lastPh = i; });
  const dotIdx = toks.findIndex((t, i) => t.k === "dot" && i > firstPh - 1 && i <= lastPh + 1);
  const expIdx = toks.findIndex((t) => t.k === "exp");
  const slashIdx = toks.findIndex((t, i) => t.k === "slash" && i > firstPh && i < lastPh);

  // Comma scaling: commas right after the last integer placeholder.
  const intEnd = dotIdx >= 0 ? dotIdx : expIdx >= 0 ? expIdx : lastPh + 1;
  let scale = 0;
  for (let i = intEnd - 1; i > firstPh && toks[i].k === "comma"; i--) scale++;
  if (dotIdx < 0 && expIdx < 0) {
    for (let i = lastPh + 1; i < toks.length && toks[i].k === "comma"; i++) scale++;
  }
  a /= Math.pow(1000, scale);
  const grouping = toks.some((t, i) => t.k === "comma" && i > firstPh && i < intEnd - scale && toks.slice(i + 1, intEnd).some((u) => u.k === "ph"));

  let body;
  if (slashIdx > 0) body = renderFraction(toks, a, firstPh, lastPh, slashIdx);
  else if (expIdx > 0) body = renderSci(toks, a, firstPh, expIdx);
  else {
    const intPh = [], fracPh = [];
    for (let i = firstPh; i <= lastPh; i++) {
      const t = toks[i];
      if (t.k !== "ph") continue;
      if (dotIdx >= 0 && i > dotIdx) fracPh.push(t.v); else intPh.push(t.v);
    }
    const litInside = toks.slice(firstPh, dotIdx >= 0 ? dotIdx : lastPh + 1).some((t) => t.k === "lit");
    const r = roundDecimal(a, fracPh.length);
    const fixed = fixedString(r, fracPh.length);
    let [ip, fp = ""] = fixed.split(".");
    if (litInside && dotIdx < 0) {
      // Digit-fill mode, e.g. "000-00-0000" or "(000) 000-0000".
      return wrapNumber(toks, firstPh, lastPh, fillDigits(toks, firstPh, lastPh, ip), neg);
    }
    if (ip === "0" && !intPh.includes("0")) ip = "";
    // pad integer part according to placeholders
    if (ip.length < intPh.length) {
      const missing = intPh.slice(0, intPh.length - ip.length);
      ip = missing.map((ph) => (ph === "0" ? "0" : ph === "?" ? " " : "")).join("") + ip;
    }
    if (grouping) ip = ip.replace(/\d+/, (d) => d.replace(/\B(?=(\d{3})+(?!\d))/g, ","));
    // trim optional fraction digits
    if (fracPh.length) {
      const chars = fp.split("");
      for (let k = fracPh.length - 1; k >= 0; k--) {
        if (chars[k] !== "0") break;
        if (fracPh[k] === "#") chars[k] = "";
        else if (fracPh[k] === "?") chars[k] = " ";
        else break;
      }
      fp = chars.join("");
    }
    body = ip + (dotIdx >= 0 ? "." + fp : "");
  }
  // Only show a minus sign when a non-zero digit is actually displayed.
  return wrapNumber(toks, firstPh, lastPh, body, neg && /[1-9]/.test(body));
}

/** Emit literals around the numeric core; the core (firstPh..lastPh) becomes `body`. */
function wrapNumber(toks, firstPh, lastPh, body, neg) {
  let out = "";
  let start = firstPh;
  // Extend core start backwards over a leading dot (".00" formats).
  if (start > 0 && toks[start - 1].k === "dot") start--;
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    if (i === start) { out += body; i = Math.max(i, lastPh); continue; }
    if (t.k === "lit") out += t.v;
    else if (t.k === "pct") out += "%";
    else if (t.k === "slash") out += "/";
    else if (t.k === "dt") out += t.v; // stray letters
  }
  return neg ? "-" + out : out;
}

function fillDigits(toks, firstPh, lastPh, digits) {
  const slots = [];
  for (let i = firstPh; i <= lastPh; i++) if (toks[i].k === "ph") slots.push(i);
  const fill = new Map();
  let d = digits === "0" ? "" : digits;
  for (let k = slots.length - 1; k >= 0; k--) {
    const ph = toks[slots[k]].v;
    if (k === 0) fill.set(slots[k], d || (ph === "0" ? "0" : ph === "?" ? " " : ""));
    else if (d.length) { fill.set(slots[k], d[d.length - 1]); d = d.slice(0, -1); }
    else fill.set(slots[k], ph === "0" ? "0" : ph === "?" ? " " : "");
  }
  let out = "";
  for (let i = firstPh; i <= lastPh; i++) {
    const t = toks[i];
    out += t.k === "ph" ? fill.get(i) : t.k === "lit" ? t.v : t.k === "comma" ? "," : "";
  }
  return out;
}

function renderSci(toks, a, firstPh, expIdx) {
  const intPh = [], fracPh = [], expPh = [];
  let seenDot = false;
  for (let i = firstPh; i < toks.length; i++) {
    const t = toks[i];
    if (i > expIdx) { if (t.k === "ph") expPh.push(t.v); else break; continue; }
    if (t.k === "dot") seenDot = true;
    else if (t.k === "ph") (seenDot ? fracPh : intPh).push(t.v);
  }
  const step = intPh.length > 1 && intPh.includes("#") ? intPh.length : 1;
  let e = a === 0 ? 0 : Math.floor(Math.log10(a));
  e = Math.floor(e / step) * step;
  let m = a === 0 ? 0 : a / Math.pow(10, e);
  m = roundDecimal(m, fracPh.length);
  if (m >= Math.pow(10, step)) { e += step; m = roundDecimal(a / Math.pow(10, e), fracPh.length); }
  let ms = m.toFixed(fracPh.length);
  if (fracPh.length && fracPh.every((p) => p === "#")) ms = ms.replace(/\.?0+$/, "");
  const sign = toks[expIdx].v;
  const es = String(Math.abs(e)).padStart(Math.max(1, expPh.length), "0");
  return `${ms}E${e < 0 ? "-" : sign === "+" ? "+" : ""}${es}`;
}

function renderFraction(toks, a, firstPh, lastPh, slashIdx) {
  // Split placeholders: [whole] [numerator] / [denominator]
  const before = toks.slice(firstPh, slashIdx);
  const after = toks.slice(slashIdx + 1, lastPh + 1);
  let litBetween = before.findIndex((t) => t.k === "lit");
  const hasWhole = litBetween > 0;
  const denomFixed = after.filter((t) => t.k === "lit").map((t) => t.v).join("").trim();
  const denomDigits = after.filter((t) => t.k === "ph").length;
  let whole = hasWhole ? Math.floor(a) : 0;
  let frac = a - whole;
  let num, den;
  if (/^\d+$/.test(denomFixed)) {
    den = Number(denomFixed);
    num = Math.round(frac * den);
  } else {
    const maxDen = Math.pow(10, Math.max(1, denomDigits)) - 1;
    let best = { n: 0, d: 1, err: frac };
    for (let d = 1; d <= maxDen; d++) {
      const n = Math.round(frac * d);
      const err = Math.abs(frac - n / d);
      if (err < best.err - 1e-12) best = { n, d, err };
      if (err < 1e-12) break;
    }
    num = best.n; den = best.d;
  }
  if (num === den) { whole += 1; num = 0; }
  if (!hasWhole) { num += whole * den; whole = 0; }
  const wholeStr = hasWhole ? (whole || num === 0 ? String(whole) : "") : "";
  if (num === 0) return hasWhole ? (wholeStr || "0") : "0";
  return `${wholeStr}${wholeStr ? " " : ""}${num}/${den}`;
}

// ---------------------------------------------------------------------------
// Cell display / raw text / data entry
// ---------------------------------------------------------------------------

/** Value of a cell in "number space": date cells become serials. */
export function cellNumericValue(cell) {
  if (!cell) return null;
  if (cell.t === "d") return isoToSerial(cell.v);
  return cell.v;
}

/** Display text + color for a cell (what the user sees in the grid). */
export function formatCell(cell) {
  if (!cell) return { text: "", color: null };
  if (cell.t === "f" || (cell.f && cell.v == null)) return { text: cell.f || "", color: null };
  const v = cell.v;
  if (v == null) return { text: "", color: null };
  if (cell.t === "b") return { text: v ? "TRUE" : "FALSE", color: null };
  if (cell.t === "e") return { text: String(v), color: null };
  if (cell.t === "d") {
    const serial = isoToSerial(v);
    if (serial == null) return { text: String(v), color: null };
    if (cell.z && isDateFormat(cell.z)) return formatValueEx(serial, cell.z);
    return { text: formatDateDisplay(v), color: null };
  }
  if (typeof v === "number") return formatValueEx(v, cell.z);
  if (cell.z && cell.z !== "@" && /@/.test(cell.z)) return formatValueEx(String(v), cell.z);
  return { text: String(v), color: null };
}

/** Display string for a cell in the grid. */
export function cellDisplay(cell) {
  return formatCell(cell).text;
}

/** Default horizontal alignment of a cell's value (Excel "General" alignment). */
export function cellAlign(cell) {
  if (!cell) return "left";
  if (cell.s && cell.s.align) return cell.s.align;
  const t = cell.t === "f" ? null : cell.t;
  if (t === "n" || t === "d") return "right";
  if (t === "b" || t === "e") return "center";
  return "left";
}

/** Would `parseCellInput(s)` give back exactly the string s? */
function stringRoundTrips(s) {
  if (s === "") return true;
  const p = parseCellInput(s);
  return !!p && p.t === "s" && p.v === s;
}

/**
 * Raw content for the formula bar / in-cell editor. Feeding it back through
 * parseCellInput reproduces the same value and type.
 */
export function cellRaw(cell) {
  if (!cell) return "";
  if (cell.f) return cell.f;
  const v = cell.v;
  if (v == null) return "";
  switch (cell.t) {
    case "b": return v ? "TRUE" : "FALSE";
    case "e": return String(v);
    case "d": {
      const serial = isoToSerial(v);
      if (serial == null) return String(v);
      return formatDateDisplay(serialToIso(serial));
    }
    case "n": {
      if (typeof v !== "number") return String(v);
      const z = cell.z || "";
      if (z && isDateFormat(z)) {
        if (isDateOnlyOrDateTimeFormat(z)) return formatDateDisplay(serialToIso(v));
        const p = serialToParts(v);
        if (v >= 0 && v < 1) return `${p.H}:${pad2(p.M)}:${pad2(p.S)}`;
        return formatNumberFull(v);
      }
      if (/%/.test(z.replace(/"[^"]*"/g, ""))) return formatNumberFull(roundDecimal(v * 100, 13)) + "%";
      return formatNumberFull(v);
    }
    default: {
      const s = String(v);
      return stringRoundTrips(s) ? s : `'${s}`;
    }
  }
}

export function formatDateDisplay(iso) {
  const s = String(iso);
  const m = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2}))?)?/.exec(s);
  if (!m) return s;
  const date = `${m[1]}-${m[2]}-${m[3]}`;
  if (m[4] == null || (m[4] === "00" && m[5] === "00" && (m[6] == null || m[6] === "00"))) return date;
  return `${date} ${m[4]}:${m[5]}${m[6] ? ":" + m[6] : ""}`;
}

export const ERROR_VALUES = ["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A",
  "#GETTING_DATA", "#SPILL!", "#CALC!", "#CIRC!"];

const MONTH_ABBR = MONTHS.map((m) => m.slice(0, 3).toLowerCase());
const CURRENCY = "$€£¥₹";

function validDate(y, m, d) {
  return m >= 1 && m <= 12 && d >= 1 && d <= daysInMonth(y, m);
}

function parseTimeParts(s) {
  const m = /^(\d{1,2}):(\d{2})(?::(\d{2})(?:\.(\d+))?)?\s*([AaPp][Mm])?$/.exec(s);
  if (!m) return null;
  let h = +m[1];
  const mi = +m[2], se = +(m[3] || 0);
  if (mi > 59 || se > 59) return null;
  if (m[5]) {
    if (h < 1 || h > 12) return null;
    const pm = /p/i.test(m[5]);
    h = (h % 12) + (pm ? 12 : 0);
  } else if (h > 23) return null;
  return { h, mi, se, frac: m[4] ? Number("0." + m[4]) : 0, hasSec: m[3] != null, ampm: !!m[5], lead: m[1].length === 2 };
}

function isoOf(y, m, d, t) {
  const date = `${pad4(y)}-${pad2(m)}-${pad2(d)}`;
  if (!t) return date;
  return `${date}T${pad2(t.h)}:${pad2(t.mi)}:${pad2(t.se)}`;
}

function timeFormat(t) {
  const h = t.lead && !t.ampm ? "hh" : "h";
  return `${h}:mm${t.hasSec ? ":ss" : ""}${t.ampm ? " AM/PM" : ""}`;
}

/** Recognise typed dates/datetimes/times. Returns a cell or null. */
function parseDateInput(s) {
  let m;
  // ISO: 2024-01-31, 2024-01-31 10:30[:15], 2024-01-31T10:30:15[Z|+hh:mm]
  if ((m = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:([T ])(\d{1,2}:\d{2}(?::\d{2}(?:\.\d+)?)?)(?:Z|[+-]\d{2}:?\d{2})?)?$/.exec(s))) {
    const y = +m[1], mo = +m[2], d = +m[3];
    if (!validDate(y, mo, d)) return null;
    const dz = `yyyy-${m[2].length === 2 ? "mm" : "m"}-${m[3].length === 2 ? "dd" : "d"}`;
    if (!m[5]) return { t: "d", v: isoOf(y, mo, d), z: dz };
    const t = parseTimeParts(m[5]);
    if (!t) return null;
    const sep = m[4] === "T" ? '"T"' : " ";
    return { t: "d", v: isoOf(y, mo, d, t), z: `${dz}${sep}${t.lead ? "hh" : "h"}:mm${t.hasSec ? ":ss" : ""}` };
  }
  // US: 1/31/2024, 01/31/24, 1/31/2024 10:30 PM
  if ((m = /^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})(?:\s+(\d{1,2}:\d{2}(?::\d{2})?(?:\s*[AaPp][Mm])?))?$/.exec(s))) {
    const mo = +m[1], d = +m[2];
    let y = +m[3];
    if (m[3].length === 2) y += y < 30 ? 2000 : 1900;
    if (!validDate(y, mo, d)) return null;
    const dz = `${m[1].length === 2 ? "mm" : "m"}/${m[2].length === 2 ? "dd" : "d"}/${m[3].length === 2 ? "yy" : "yyyy"}`;
    if (!m[4]) return { t: "d", v: isoOf(y, mo, d), z: dz };
    const t = parseTimeParts(m[4]);
    if (!t) return null;
    return { t: "d", v: isoOf(y, mo, d, t), z: `${dz} ${timeFormat(t)}` };
  }
  // 31-Jan-2024, 31 Jan 24
  if ((m = /^(\d{1,2})[- ]([A-Za-z]{3,9})[- ](\d{2}|\d{4})$/.exec(s))) {
    const mo = MONTH_ABBR.indexOf(m[2].slice(0, 3).toLowerCase()) + 1;
    if (!mo || (m[2].length > 3 && MONTHS[mo - 1].toLowerCase() !== m[2].toLowerCase())) return null;
    let y = +m[3];
    if (m[3].length === 2) y += y < 30 ? 2000 : 1900;
    const d = +m[1];
    if (!validDate(y, mo, d)) return null;
    const sep = s.includes("-") ? "-" : " ";
    return { t: "d", v: isoOf(y, mo, d), z: `d${sep}${m[2].length > 3 ? "mmmm" : "mmm"}${sep}${m[3].length === 2 ? "yy" : "yyyy"}` };
  }
  // Time of day: 10:30, 10:30:15, 9:05 PM  -> fraction of a day
  const t = parseTimeParts(s);
  if (t) {
    return { t: "n", v: (t.h * 3600 + t.mi * 60 + t.se + t.frac) / 86400, z: timeFormat(t) };
  }
  return null;
}

/** Strict number syntax with optional thousands separators. */
const NUM_RE = /^[+-]?(?:\d{1,3}(?:,\d{3})+|\d+)?(?:\.\d*)?(?:[eE][+-]?\d+)?$/;

function parseNumberText(s) {
  if (!s || !/\d/.test(s) || !NUM_RE.test(s)) return null;
  const n = Number(s.replace(/,/g, ""));
  return isFinite(n) ? n : null;
}

/**
 * Parse what a user typed into a cell, with Excel-like entry semantics.
 *  - ""            -> null (clear)
 *  - "'text"       -> forced text (apostrophe removed)
 *  - "=..."        -> formula {f, t:"f", v:null}
 *  - TRUE / FALSE  -> boolean
 *  - #N/A etc.     -> error value
 *  - 1,234.5  (12) -> number;  3% -> 0.03 z "0%";  $1,234.50 -> 1234.5 z "$"#,##0.00
 *  - 2024-01-31, 1/31/2024, 31-Jan-2024 [time] -> {t:"d", v:naive ISO, z}
 *  - 10:30         -> {t:"n", v:0.4375, z:"h:mm"}
 *  - anything else -> string (kept verbatim)
 */
export function parseCellInput(text) {
  if (text == null) return null;
  const s = String(text);
  if (s === "") return null;
  if (s[0] === "'") return { t: "s", v: s.slice(1) };
  if (s[0] === "=" && s.length > 1) return { f: s, t: "f", v: null };
  const trimmed = s.trim();
  if (!trimmed) return { t: "s", v: s };
  if (/^TRUE$/i.test(trimmed)) return { t: "b", v: true };
  if (/^FALSE$/i.test(trimmed)) return { t: "b", v: false };
  const err = ERROR_VALUES.find((e) => e === trimmed.toUpperCase());
  if (err) return { t: "e", v: err };

  // Plain / thousands-separated number.
  let n = parseNumberText(trimmed);
  if (n != null) return { t: "n", v: n };

  // Accounting negatives "(123)".
  let m;
  if ((m = /^\((.+)\)$/.exec(trimmed))) {
    const inner = parseCellInput(m[1]);
    if (inner && inner.t === "n" && inner.v >= 0) return { ...inner, v: -inner.v };
    return { t: "s", v: s };
  }

  // Percent.
  if ((m = /^([+-]?[\d.,eE+-]+)\s*%$/.exec(trimmed))) {
    n = parseNumberText(m[1]);
    if (n != null) {
      const dec = (m[1].split(".")[1] || "").replace(/[eE].*/, "").length;
      const v = roundDecimal(n / 100, 15);
      return { t: "n", v, z: dec ? `0.${"0".repeat(Math.min(dec, 10))}%` : "0%" };
    }
  }

  // Currency prefix/suffix: $1,234.50  -$5  $-5  1.234 €
  if ((m = new RegExp(`^([+-]?)\\s*([${CURRENCY}])\\s*([+-]?[\\d.,]+)$`).exec(trimmed)) ||
      (m = new RegExp(`^([+-]?)([\\d.,]+)\\s*([${CURRENCY}])$`).exec(trimmed))) {
    const prefix = CURRENCY.includes(m[2]);
    const sym = prefix ? m[2] : m[3];
    const numText = prefix ? m[3] : m[2];
    n = parseNumberText(numText.replace(/^[+-]/, ""));
    if (n != null) {
      const negative = (m[1] === "-") !== (numText[0] === "-");
      const dec = (numText.split(".")[1] || "").length;
      const body = `#,##0${dec ? "." + "0".repeat(Math.min(dec, 10)) : ""}`;
      return { t: "n", v: negative ? -n : n, z: prefix ? `"${sym}"${body}` : `${body} "${sym}"` };
    }
  }

  const d = parseDateInput(trimmed);
  if (d) return d;
  return { t: "s", v: s };
}

/** Does the value family of `cell` fit the number format z? Used to keep a cell's format on re-entry. */
export function formatFitsCell(z, cell) {
  if (!z || !cell) return false;
  if (cell.t === "s" || cell.t === "b" || cell.t === "e" || cell.t === "f") return true;
  if (cell.t === "d") return isDateOnlyOrDateTimeFormat(z);
  if (cell.t === "n") return !isDateFormat(z) || (!isDateOnlyOrDateTimeFormat(z) && cell.v >= 0 && cell.v < 1);
  return false;
}

/**
 * Turn typed text into the new cell for `prev`, preserving style and (when
 * it still fits) number format. Returns `prev` itself when nothing changed.
 */
export function cellFromInput(text, prev) {
  if (prev && text === cellRaw(prev)) return prev;
  let next;
  if (prev && prev.z === "@" && text !== "") next = { t: "s", v: String(text).replace(/^'/, "") };
  else next = parseCellInput(text);
  if (!next) {
    // Clearing keeps formatting only (Excel keeps formats on Delete).
    if (prev && (prev.s || prev.z)) {
      const keep = {};
      if (prev.s) keep.s = { ...prev.s };
      if (prev.z) keep.z = prev.z;
      return { t: "s", v: null, ...keep };
    }
    return null;
  }
  if (prev) {
    if (prev.s && !next.s) next.s = { ...prev.s };
    if (prev.z && !next.z && formatFitsCell(prev.z, next)) next.z = prev.z;
    else if (prev.z && next.z && next.t === "d" && isDateOnlyOrDateTimeFormat(prev.z) &&
      /h/.test(next.z) === /h/i.test(prev.z)) next.z = prev.z;
  }
  return next;
}

/** True when a cell carries no value (it may still carry formatting). */
export function cellIsEmpty(cell) {
  return !cell || (cell.v == null && !cell.f);
}

// ---------------------------------------------------------------------------
// Sheet -> static HTML table (used for printing and simple previews)
// ---------------------------------------------------------------------------

export function colLetterWidth(index) {
  return indexToCol(index).length;
}

export function escapeHtmlText(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

/** Inline CSS for a cell's style object. */
export function cellStyleCss(cell, align) {
  const st = cell && cell.s;
  let css = "";
  if (st) {
    if (st.b) css += "font-weight:700;";
    if (st.i) css += "font-style:italic;";
    if (st.u) css += "text-decoration:underline;";
    if (st.color && /^#[0-9a-f]{6}$/i.test(st.color)) css += `color:${st.color};`;
    if (st.bg && /^#[0-9a-f]{6}$/i.test(st.bg)) css += `background:${st.bg};`;
    if (st.wrap) css += "white-space:pre-wrap;";
  }
  if (align && align !== "left") css += `text-align:${align};`;
  return css;
}

/** Excel column width (characters) -> CSS pixels. */
export function charsToPx(chars) {
  return Math.max(4, Math.round(Number(chars) * 7 + 5));
}
export function pxToChars(px) {
  return Math.max(0.5, Math.round(((px - 5) / 7) * 100) / 100);
}

/**
 * Render one sheet as an HTML table string (row/column headers included,
 * cells carry data-addr). Bounded by opts.maxRows/maxCols/cellBudget.
 */
export function renderSheetHtml(sheet, opts = {}) {
  const cells = sheet.cells || {};
  let range = opts.range || usedRange(cells);
  // merged areas count as used (a merge may extend past the last value)
  for (const ref of sheet.merges || []) {
    const m = parseRange(ref);
    if (!m) continue;
    range = range ? { ...range, r2: Math.max(range.r2, m.r2), c2: Math.max(range.c2, m.c2) } : { r1: 0, c1: 0, r2: m.r2, c2: m.c2 };
  }
  const maxRows = opts.maxRows ?? 1500;
  const maxCols = opts.maxCols ?? 150;
  const cellBudget = opts.cellBudget ?? 120000;
  let r2 = Math.min(Math.max(range ? range.r2 : 0, (opts.minRows ?? 10) - 1), maxRows - 1);
  let c2 = Math.min(Math.max(range ? range.c2 : 0, (opts.minCols ?? 5) - 1), maxCols - 1);
  if ((r2 + 1) * (c2 + 1) > cellBudget) {
    r2 = Math.min(r2, Math.max((opts.minRows ?? 10) - 1, Math.floor(cellBudget / (c2 + 1)) - 1));
  }
  const r1 = 0, c1 = 0;

  // merges: top-left -> span, covered -> skip
  const spans = new Map();
  const covered = new Set();
  for (const ref of sheet.merges || []) {
    const m = parseRange(ref);
    if (!m) continue;
    spans.set(rcToAddr(m.r1, m.c1), { rs: m.r2 - m.r1 + 1, cs: m.c2 - m.c1 + 1 });
    for (let r = m.r1; r <= m.r2; r++) for (let c = m.c1; c <= m.c2; c++) {
      if (r !== m.r1 || c !== m.c1) covered.add(rcToAddr(r, c));
    }
  }

  let html = '<table class="ow-grid"><thead><tr><th class="ow-corner"></th>';
  for (let c = c1; c <= c2; c++) {
    html += `<th class="ow-colhead" data-col="${indexToCol(c)}">${indexToCol(c)}</th>`;
  }
  html += "</tr></thead><tbody>";
  const widths = sheet.colWidths || {};
  for (let r = r1; r <= r2; r++) {
    html += `<tr><th class="ow-rowhead">${r + 1}</th>`;
    for (let c = c1; c <= c2; c++) {
      const addr = rcToAddr(r, c);
      if (covered.has(addr)) continue;
      const cell = cells[addr];
      const { text, color } = formatCell(cell);
      const w = widths[indexToCol(c)];
      let css = w ? `min-width:${charsToPx(w)}px;max-width:${charsToPx(w)}px;` : "";
      css += cellStyleCss(cell, cellAlign(cell));
      if (color) css += `color:${color};`;
      const span = spans.get(addr);
      const spanAttr = span ? `${span.rs > 1 ? ` rowspan="${span.rs}"` : ""}${span.cs > 1 ? ` colspan="${span.cs}"` : ""}` : "";
      const cls = cell && (cell.t === "n" || cell.t === "d") ? "ow-cell ow-num" : "ow-cell";
      html += `<td class="${cls}" data-addr="${addr}"${spanAttr}${css ? ` style="${css}"` : ""}>${escapeHtmlText(text)}</td>`;
    }
    html += "</tr>";
  }
  html += "</tbody></table>";
  return { html, range: { r1, r2, c1, c2 } };
}

// ---------------------------------------------------------------------------
// CSV codec (RFC 4180)
// ---------------------------------------------------------------------------

/**
 * Parse delimited text. A quote only opens a quoted field when it is the
 * first character of the field; a quote anywhere else is literal (12" long).
 */
export function csvParse(text, delim = ",") {
  const rows = [];
  let row = [];
  let field = "";
  let inQ = false;
  let fieldStart = true;
  const n = text.length;
  for (let i = 0; i < n; i++) {
    const ch = text[i];
    if (inQ) {
      if (ch === '"') {
        if (text[i + 1] === '"') { field += '"'; i++; }
        else inQ = false;
      } else field += ch;
      continue;
    }
    if (ch === '"' && fieldStart) { inQ = true; fieldStart = false; continue; }
    if (ch === delim) { row.push(field); field = ""; fieldStart = true; continue; }
    if (ch === "\n" || ch === "\r") {
      if (ch === "\r" && text[i + 1] === "\n") i++;
      row.push(field); field = "";
      rows.push(row); row = [];
      fieldStart = true;
      continue;
    }
    field += ch;
    fieldStart = false;
  }
  if (field !== "" || row.length || inQ) { row.push(field); rows.push(row); }
  return rows;
}

/**
 * Serialize rows; fields containing the delimiter, quotes or newlines are quoted.
 * opts.bareQuotes: the source file left quotes inside unquoted fields (12" ruler);
 * mirror that — only fields that start or end with a quote or contain the
 * delimiter / a line break are quoted (a quote with text on both sides, like an
 * inch mark, stays bare) — so an unedited file saves byte-identical. Our
 * reader, Excel and Python's csv module all read such fields literally.
 */
export function csvSerialize(rows, delim = ",", opts = {}) {
  const eol = opts.eol || "\r\n";
  const delimClass = delim === "\t" ? "\\t" : delim.replace(/[\\\]^-]/g, "\\$&");
  const special = new RegExp(opts.bareQuotes ? `^"|"$|[\\r\\n${delimClass}]` : `["\\r\\n${delimClass}]`);
  const body = rows.map((row) =>
    row.map((f) => {
      const s = String(f == null ? "" : f);
      return (opts.quoteAll && s !== "") || special.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
    }).join(delim)
  ).join(eol);
  return opts.trailingNewline && rows.length ? body + eol : body;
}

/** Guess the delimiter of CSV text: "," ";" tab or "|". */
export function sniffDelimiter(text, fallback = ",") {
  const cands = [",", ";", "\t", "|"];
  const lines = [];
  let cur = "", inQ = false;
  for (let i = 0; i < text.length && lines.length < 30; i++) {
    const ch = text[i];
    if (ch === '"') inQ = !inQ;
    if (!inQ && (ch === "\n" || ch === "\r")) {
      if (cur) lines.push(cur);
      cur = "";
      continue;
    }
    if (!inQ) cur += ch; else if (ch !== '"') cur += "x";
  }
  if (cur && lines.length < 30) lines.push(cur);
  if (!lines.length) return fallback;
  let best = fallback, bestScore = 0;
  for (const d of cands) {
    const counts = lines.map((l) => l.split(d).length - 1);
    const freq = new Map();
    for (const c of counts) if (c > 0) freq.set(c, (freq.get(c) || 0) + 1);
    if (!freq.size) continue;
    let mode = 0, modeN = 0;
    for (const [c, k] of freq) if (k > modeN || (k === modeN && c > mode)) { mode = c; modeN = k; }
    const consistency = modeN / lines.length;
    let score = consistency * 10 + Math.min(mode, 50) / 50;
    // "1,5" style decimal commas are a strong hint for ";" (European Excel exports).
    if (d === ";" && /\d,\d/.test(text.slice(0, 4000))) score += 0.5;
    if (score > bestScore + 1e-9) { bestScore = score; best = d; }
  }
  return best;
}

// Windows-1252 code points for bytes 0x80..0x9F.
const CP1252_HIGH = [0x20ac, 0x81, 0x201a, 0x192, 0x201e, 0x2026, 0x2020, 0x2021, 0x2c6, 0x2030, 0x160,
  0x2039, 0x152, 0x8d, 0x17d, 0x8f, 0x90, 0x2018, 0x2019, 0x201c, 0x201d, 0x2022, 0x2013, 0x2014,
  0x2dc, 0x2122, 0x161, 0x203a, 0x153, 0x9d, 0x17e, 0x178];

function decodeCp1252(u8) {
  let s = "";
  for (let i = 0; i < u8.length; i += 0x4000) {
    const chunk = u8.subarray(i, i + 0x4000);
    const codes = new Array(chunk.length);
    for (let k = 0; k < chunk.length; k++) {
      const b = chunk[k];
      codes[k] = b >= 0x80 && b <= 0x9f ? CP1252_HIGH[b - 0x80] : b;
    }
    s += String.fromCharCode.apply(null, codes);
  }
  return s;
}

/**
 * Decode text bytes: BOM detection (UTF-8, UTF-16LE/BE), strict UTF-8, then
 * Windows-1252 fallback. Returns {text, encoding, bom}.
 */
export function decodeTextBytes(bytes) {
  const u8 = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  if (u8[0] === 0xef && u8[1] === 0xbb && u8[2] === 0xbf) {
    return { text: new TextDecoder("utf-8").decode(u8.subarray(3)), encoding: "utf-8", bom: true };
  }
  if (u8[0] === 0xff && u8[1] === 0xfe) {
    return { text: new TextDecoder("utf-16le").decode(u8.subarray(2)), encoding: "utf-16le", bom: true };
  }
  if (u8[0] === 0xfe && u8[1] === 0xff) {
    return { text: new TextDecoder("utf-16be").decode(u8.subarray(2)), encoding: "utf-16be", bom: true };
  }
  // BOM-less UTF-16: lots of NUL bytes in odd/even positions.
  if (u8.length >= 4) {
    let oddZero = 0, evenZero = 0;
    const n = Math.min(u8.length, 4000);
    for (let i = 0; i < n; i++) if (u8[i] === 0) (i % 2 ? oddZero++ : evenZero++);
    if (oddZero > n / 4 && evenZero === 0) return { text: new TextDecoder("utf-16le").decode(u8), encoding: "utf-16le", bom: false };
    if (evenZero > n / 4 && oddZero === 0) return { text: new TextDecoder("utf-16be").decode(u8), encoding: "utf-16be", bom: false };
  }
  try {
    return { text: new TextDecoder("utf-8", { fatal: true }).decode(u8), encoding: "utf-8", bom: false };
  } catch {
    let text;
    try { text = new TextDecoder("windows-1252").decode(u8); } catch { text = decodeCp1252(u8); }
    return { text, encoding: "windows-1252", bom: false };
  }
}

/** Encode text in the given encoding. Unencodable Windows-1252 text falls back to UTF-8 with BOM. */
export function encodeText(text, encoding = "utf-8", bom = false) {
  const enc = String(encoding || "utf-8").toLowerCase();
  if (enc === "utf-16le" || enc === "utf-16be") {
    const le = enc === "utf-16le";
    const off = bom ? 2 : 0;
    const out = new Uint8Array(off + text.length * 2);
    if (bom) { out[0] = le ? 0xff : 0xfe; out[1] = le ? 0xfe : 0xff; }
    for (let i = 0; i < text.length; i++) {
      const c = text.charCodeAt(i);
      out[off + i * 2] = le ? c & 0xff : c >> 8;
      out[off + i * 2 + 1] = le ? c >> 8 : c & 0xff;
    }
    return out;
  }
  if (enc === "windows-1252") {
    const out = new Uint8Array(text.length);
    for (let i = 0; i < text.length; i++) {
      const c = text.charCodeAt(i);
      if (c < 0x80 || (c >= 0xa0 && c <= 0xff)) { out[i] = c; continue; }
      const k = CP1252_HIGH.indexOf(c);
      if (k < 0) return encodeText(text, "utf-8", true); // can't represent: switch to UTF-8 + BOM
      out[i] = 0x80 + k;
    }
    return out;
  }
  const body = new TextEncoder().encode(text);
  if (!bom) return body;
  const out = new Uint8Array(body.length + 3);
  out.set([0xef, 0xbb, 0xbf]);
  out.set(body, 3);
  return out;
}

/** Text written to CSV / plain-text clipboard for a cell: the displayed value, never the formula. */
export function cellText(cell) {
  if (!cell) return "";
  if (cell.v == null) return "";
  if (cell.t === "f") return "";
  switch (cell.t) {
    case "b": return cell.v ? "TRUE" : "FALSE";
    case "e": return String(cell.v);
    case "d": {
      const serial = isoToSerial(cell.v);
      if (serial == null) return String(cell.v);
      if (cell.z && isDateFormat(cell.z)) return formatValue(serial, cell.z);
      return formatDateDisplay(cell.v);
    }
    case "n":
      if (typeof cell.v !== "number") return String(cell.v);
      if (cell.z && !/^general$/i.test(cell.z)) return formatValue(cell.v, cell.z);
      return formatNumberFull(cell.v);
    default:
      return String(cell.v);
  }
}

/**
 * Convert one CSV field into a cell, losslessly: a typed value is used only
 * when writing it back reproduces the exact same text; otherwise the field
 * stays a string (keeps 00501, +1555…, 12345678901234567890, "1e5" as-is).
 */
export function csvFieldToCell(f) {
  if (f === "") return null;
  if (f[0] === "=" || f[0] === "'" || f[0] === "@") return { t: "s", v: f };
  const p = parseCellInput(f);
  if (!p || p.t === "s" || p.t === "f") return { t: "s", v: f };
  if (cellText(p) === f) return p;
  if (p.t === "n" && !p.z) {
    // 1.50 -> 1.5 with "0.00";  1,234 -> 1234 with "#,##0"
    const dec = (f.split(".")[1] || "").length;
    const z = `${f.includes(",") ? "#,##0" : "0"}${dec ? "." + "0".repeat(dec) : ""}`;
    const cand = { ...p, z };
    if (cellText(cand) === f) return cand;
  }
  return { t: "s", v: f };
}

/** Convert CSV text into an excel model (one sheet). */
export function csvToModel(text, name = "Sheet1", delim = ",") {
  const rows = csvParse(text, delim);
  const cells = {};
  rows.forEach((row, r) => {
    row.forEach((f, c) => {
      const cell = csvFieldToCell(f);
      if (cell) cells[rcToAddr(r, c)] = cell;
    });
  });
  const sheetName = sanitizeSheetName(name);
  return { kind: "excel", active: 0, sheets: [{ name: sheetName, cells }] };
}

/**
 * CSV/TSV bytes -> excel model with encoding detection and delimiter
 * sniffing. The detected properties are kept in model.csvMeta so saving can
 * reproduce the file byte-for-byte when nothing was edited.
 */
export function csvBytesToModel(bytes, name = "Sheet1", delim) {
  const { text, encoding, bom } = decodeTextBytes(bytes);
  const isTsv = /\.tsv$/i.test(String(name));
  const d = delim || sniffDelimiter(text, isTsv ? "\t" : ",");
  const base = String(name || "Sheet1").replace(/^.*[\\/]/, "").replace(/\.[^.]+$/, "") || "Sheet1";
  const model = csvToModel(text, base, d);
  const firstEol = /\r\n|\n|\r/.exec(text);
  const rows = csvParse(text, d);
  const nonEmpty = [];
  let quoted = 0;
  // Detect "quote every field" writers so we can mirror them.
  const qre = new RegExp(`(^|${d === "\t" ? "\\t" : d.replace(/[|\\]/g, "\\$&")}|\\r?\\n)"`, "g");
  quoted = (text.match(qre) || []).length;
  for (const row of rows) for (const f of row) if (f !== "") nonEmpty.push(f);
  model.csvMeta = {
    encoding, bom, delim: d,
    eol: firstEol ? firstEol[0] : "\r\n",
    trailingNewline: /[\r\n]$/.test(text),
    quoteAll: nonEmpty.length > 0 && quoted >= nonEmpty.length,
    bareQuotes: hasBareQuotes(text, d),
  };
  if (encoding === "windows-1252") {
    model.warnings = [...(model.warnings || []), "File is not UTF-8; it was read as Windows-1252 and will be saved the same way."];
  }
  return model;
}

/** True when some unquoted field contains a literal quote (e.g. 12" ruler). Mirrors csvParse's rules. */
function hasBareQuotes(text, d) {
  let inQ = false, fieldStart = true;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (inQ) {
      if (ch === '"') {
        if (text[i + 1] === '"') i++;
        else inQ = false;
      }
      continue;
    }
    if (ch === '"') {
      if (fieldStart) { inQ = true; fieldStart = false; continue; }
      return true;
    }
    fieldStart = ch === d || ch === "\n" || ch === "\r";
  }
  return false;
}

/** Rows of display text for a sheet, starting at A1 (absolute positions kept). */
export function sheetToRows(sheet, range) {
  const cells = (sheet && sheet.cells) || {};
  const used = range || usedRange(cells);
  if (!used) return [];
  const rows = [];
  for (let r = 0; r <= used.r2; r++) {
    const row = [];
    for (let c = 0; c <= used.c2; c++) row.push(cellText(cells[rcToAddr(r, c)]));
    rows.push(row);
  }
  return rows;
}

/**
 * Serialize one sheet to CSV text: computed values (not formulas), from A1 to
 * the last used cell, honoring model.csvMeta (eol, quoting, trailing newline).
 */
export function modelSheetToCsv(model, sheetIndex = model?.active ?? 0, delim) {
  const sheet = model?.sheets?.[sheetIndex];
  if (!sheet) return "";
  const meta = model.csvMeta || {};
  const d = delim || meta.delim || ",";
  const rows = sheetToRows(sheet);
  if (!rows.length) return "";
  return csvSerialize(rows, d, {
    eol: meta.eol || "\r\n",
    trailingNewline: meta.trailingNewline ?? true,
    quoteAll: !!meta.quoteAll && d === meta.delim,
    bareQuotes: !!meta.bareQuotes && d === meta.delim,
  });
}

/**
 * Bytes for saving a sheet as CSV/TSV. format "csv" keeps the delimiter the
 * file was read with (e.g. ";"), "tsv" always uses tabs. Encoding/BOM follow
 * model.csvMeta (UTF-8 for new files).
 */
export function modelSheetToCsvBytes(model, sheetIndex = model?.active ?? 0, format = "csv") {
  const meta = model?.csvMeta || {};
  const delim = format === "tsv" ? "\t" : (meta.delim && meta.delim !== "\t" ? meta.delim : ",");
  const text = modelSheetToCsv(model, sheetIndex, delim);
  return encodeText(text, meta.encoding || "utf-8", meta.encoding ? !!meta.bom : false);
}

// ---------------------------------------------------------------------------
// Sheet operations: current region, sort, AutoFilter, fill series, charts.
// Pure functions over a sheet ({cells, merges, ...}); the grid UI
// (viewer/excel.js) turns their results into undoable changes.
// ---------------------------------------------------------------------------

const cellOf = (sheet, r, c) => (sheet && sheet.cells ? sheet.cells[rcToAddr(r, c)] : undefined);
const cloneCell = (cell) => (cell == null ? null : JSON.parse(JSON.stringify(cell)));

/**
 * Excel's "current region" around (r, c): the rectangle of data bounded by
 * empty rows and columns (what Ctrl+A / Sort / Filter use). A lone empty cell
 * with no data around it gives that single cell.
 */
export function currentRegion(sheet, r, c) {
  const has = (rr, cc) => rr >= 0 && cc >= 0 && rr < MAX_ROWS && cc < MAX_COLS && !cellIsEmpty(cellOf(sheet, rr, cc));
  const rowHas = (rr, c1, c2) => { for (let cc = Math.max(0, c1); cc <= Math.min(MAX_COLS - 1, c2); cc++) if (has(rr, cc)) return true; return false; };
  const colHas = (cc, r1, r2) => { for (let rr = Math.max(0, r1); rr <= Math.min(MAX_ROWS - 1, r2); rr++) if (has(rr, cc)) return true; return false; };
  const rg = { r1: r, r2: r, c1: c, c2: c };
  for (let pass = 0; pass < 1000; pass++) {
    let grew = false;
    while (rg.r2 + 1 < MAX_ROWS && rowHas(rg.r2 + 1, rg.c1 - 1, rg.c2 + 1)) { rg.r2++; grew = true; }
    while (rg.r1 > 0 && rowHas(rg.r1 - 1, rg.c1 - 1, rg.c2 + 1)) { rg.r1--; grew = true; }
    while (rg.c2 + 1 < MAX_COLS && colHas(rg.c2 + 1, rg.r1 - 1, rg.r2 + 1)) { rg.c2++; grew = true; }
    while (rg.c1 > 0 && colHas(rg.c1 - 1, rg.r1 - 1, rg.r2 + 1)) { rg.c1--; grew = true; }
    if (!grew) break;
  }
  return rg;
}

/** Merged ranges of the sheet that intersect rg. */
export function mergesIn(sheet, rg) {
  return (sheet.merges || []).map(parseRange).filter((m) => m && !(m.r2 < rg.r1 || m.r1 > rg.r2 || m.c2 < rg.c1 || m.c1 > rg.c2));
}

/**
 * Guess whether the first row of rg is a header row (Excel's heuristic, in
 * short): every non-empty cell of the first row is text, and either the row
 * below holds numbers/dates/booleans under a text header or the first row is
 * bold / filled while the second is not.
 */
export function detectHeaderRow(sheet, rg) {
  if (rg.r2 <= rg.r1) return false;
  let nonEmpty = 0, text = 0, typedBelow = false, styledTop = false, styledBelow = false;
  for (let c = rg.c1; c <= rg.c2; c++) {
    const top = cellOf(sheet, rg.r1, c), below = cellOf(sheet, rg.r1 + 1, c);
    if (top && top.s && (top.s.b || top.s.bg)) styledTop = true;
    if (below && below.s && (below.s.b || below.s.bg)) styledBelow = true;
    if (cellIsEmpty(top)) continue;
    nonEmpty++;
    if (top.t === "s" && !top.f) text++;
    if (below && !cellIsEmpty(below) && (below.t === "n" || below.t === "d" || below.t === "b")) typedBelow = true;
  }
  if (!nonEmpty || text < nonEmpty) return false;
  return typedBelow || (styledTop && !styledBelow);
}

let collator = null;
function textCompare(a, b) {
  if (!collator) {
    try {
      collator = new Intl.Collator(undefined, { sensitivity: "base", numeric: false });
    } catch {
      collator = { compare: (x, y) => (x.toLowerCase() < y.toLowerCase() ? -1 : x.toLowerCase() > y.toLowerCase() ? 1 : 0) };
    }
  }
  return collator.compare(a, b);
}

/**
 * Sort class of a cell like Excel's ascending order: numbers and dates (0) <
 * text (1) < logical (2) < errors (3) < blanks (4). Formula cells sort by
 * their computed value.
 */
export function sortKeyOf(cell) {
  if (!cell || cell.v == null || cell.v === "") return { k: 4, v: 0 };
  switch (cell.t) {
    case "n": return typeof cell.v === "number" ? { k: 0, v: cell.v } : { k: 1, v: String(cell.v) };
    case "d": { const s = isoToSerial(cell.v); return s == null ? { k: 1, v: String(cell.v) } : { k: 0, v: s }; }
    case "b": return { k: 2, v: cell.v ? 1 : 0 };
    case "e": return { k: 3, v: 0 };
    case "f": return { k: 4, v: 0 };
    default: return { k: 1, v: String(cell.v) };
  }
}

/** Compare two sort keys; blanks are always last, `desc` reverses the rest. */
export function compareSortKeys(a, b, desc = false) {
  if (a.k === 4 || b.k === 4) return (a.k === 4 ? 1 : 0) - (b.k === 4 ? 1 : 0);
  let d;
  if (a.k !== b.k) d = a.k - b.k;
  else if (a.k === 1) d = textCompare(a.v, b.v);
  else d = a.v < b.v ? -1 : a.v > b.v ? 1 : 0;
  return desc ? -d : d;
}

/**
 * Sort the rows of rg by keys [{col, desc}] (col = absolute column index).
 * opts.header: the first row stays in place. opts.translate(f, dr, dc):
 * formula rewriting. Formulas move with their row and their RELATIVE
 * references shift by the distance moved (Excel's behaviour, like copy and
 * paste), so "=B2*C2" moving from row 2 to row 5 becomes "=B5*C5";
 * absolute ($) references stay. Returns {changes:[{addr, after}], moved} or
 * {error} (merged cells, nothing to sort).
 */
export function sortRange(sheet, rg, keys, opts = {}) {
  if (mergesIn(sheet, rg).length) return { error: "Cells that are merged cannot be sorted. Unmerge them first (or select a range without merged cells)." };
  const first = rg.r1 + (opts.header ? 1 : 0);
  if (first >= rg.r2) return { error: "There is nothing to sort: select at least two rows of data." };
  const ks = (keys || []).filter((k) => k && k.col >= rg.c1 && k.col <= rg.c2);
  if (!ks.length) return { error: "Choose a column to sort by." };
  const rows = [];
  for (let r = first; r <= rg.r2; r++) rows.push({ r, keys: ks.map((k) => sortKeyOf(cellOf(sheet, r, k.col))) });
  rows.sort((a, b) => {
    for (let i = 0; i < ks.length; i++) {
      const d = compareSortKeys(a.keys[i], b.keys[i], !!ks[i].desc);
      if (d) return d;
    }
    return a.r - b.r;
  });
  const changes = [];
  let moved = 0;
  rows.forEach((row, i) => {
    const target = first + i;
    if (row.r !== target) moved++;
    for (let c = rg.c1; c <= rg.c2; c++) {
      const src = cellOf(sheet, row.r, c);
      const next = cloneCell(src);
      if (next && next.f && row.r !== target && opts.translate) next.f = opts.translate(next.f, target - row.r, 0);
      changes.push({ addr: rcToAddr(target, c), after: next });
    }
  });
  return { changes, moved };
}

// ---- AutoFilter -----------------------------------------------------------

/** Text a cell shows in a filter list ("" for blanks). */
export function filterText(cell) {
  if (cellIsEmpty(cell)) return "";
  return formatCell(cell).text;
}

/** Distinct values of column c in data rows r1..r2, sorted like Excel's list. */
export function filterValues(sheet, c, r1, r2, limit = 10000) {
  const seen = new Map();
  for (let r = r1; r <= r2 && seen.size < limit; r++) {
    const cell = cellOf(sheet, r, c);
    const t = filterText(cell);
    if (!seen.has(t)) seen.set(t, sortKeyOf(cell));
  }
  return [...seen.entries()].sort((a, b) => compareSortKeys(a[1], b[1]) || textCompare(a[0], b[0])).map(([t]) => t);
}

export const FILTER_OPS = [
  ["eq", "Equals"], ["ne", "Does not equal"], ["gt", "Greater than"], ["ge", "Greater than or equal to"],
  ["lt", "Less than"], ["le", "Less than or equal to"], ["between", "Between"],
  ["top", "Top N"], ["bottom", "Bottom N"], ["above", "Above average"], ["below", "Below average"],
  ["contains", "Contains"], ["notcontains", "Does not contain"], ["begins", "Begins with"], ["ends", "Ends with"],
];

function numOrNull(x) {
  if (x == null || x === "") return null;
  if (typeof x === "number") return x;
  const p = parseCellInput(String(x));
  if (p && p.t === "n") return p.v;
  if (p && p.t === "d") return isoToSerial(p.v);
  return null;
}

/**
 * Per-column criteria: { values?: [text] (allowed filterText()s; "" = blanks),
 * cond?: {op, a, b} } (both must pass when both are present).
 * Returns a predicate (cell) => boolean; `columnCells` are the column's data
 * cells (needed by top/bottom N and average conditions).
 */
export function makeFilterTest(crit, columnCells = []) {
  if (!crit) return () => true;
  const allowed = Array.isArray(crit.values) ? new Set(crit.values) : null;
  const cond = crit.cond && crit.cond.op ? crit.cond : null;
  let condTest = null;
  if (cond) {
    const op = cond.op;
    const nums = () => columnCells.map(cellNumericValue).filter((v) => typeof v === "number" && isFinite(v));
    const a = numOrNull(cond.a), b = numOrNull(cond.b);
    const sa = String(cond.a ?? "").toLowerCase();
    if (op === "top" || op === "bottom") {
      const n = Math.max(1, Math.min(500, Math.floor(a || 10)));
      const sorted = nums().sort((x, y) => (op === "top" ? y - x : x - y));
      const edge = sorted.length ? sorted[Math.min(n, sorted.length) - 1] : null;
      condTest = (cell) => { const v = cellNumericValue(cell); return typeof v === "number" && edge != null && (op === "top" ? v >= edge : v <= edge); };
    } else if (op === "above" || op === "below") {
      const ns = nums();
      const avg = ns.length ? ns.reduce((s, x) => s + x, 0) / ns.length : 0;
      condTest = (cell) => { const v = cellNumericValue(cell); return typeof v === "number" && (op === "above" ? v > avg : v < avg); };
    } else if (["gt", "ge", "lt", "le", "between"].includes(op)) {
      condTest = (cell) => {
        const v = cellNumericValue(cell);
        if (typeof v !== "number" || a == null) return false;
        switch (op) {
          case "gt": return v > a;
          case "ge": return v >= a;
          case "lt": return v < a;
          case "le": return v <= a;
          default: return b != null && v >= Math.min(a, b) && v <= Math.max(a, b);
        }
      };
    } else {
      condTest = (cell) => {
        const t = filterText(cell).toLowerCase();
        switch (op) {
          case "eq": { const v = cellNumericValue(cell); return a != null && typeof v === "number" ? v === a : t === sa; }
          case "ne": { const v = cellNumericValue(cell); return a != null && typeof v === "number" ? v !== a : t !== sa; }
          case "contains": return t.includes(sa);
          case "notcontains": return !t.includes(sa);
          case "begins": return t.startsWith(sa);
          case "ends": return t.endsWith(sa);
          default: return true;
        }
      };
    }
  }
  return (cell) => (!allowed || allowed.has(filterText(cell))) && (!condTest || condTest(cell));
}

/**
 * Rows hidden by an AutoFilter: ref = "A1:D20" (first row = headers),
 * criteria = { [absoluteColumnIndex]: crit }. Returns ascending row indexes.
 */
export function filterHiddenRows(sheet, ref, criteria) {
  const rg = parseRange(ref || "");
  if (!rg || !criteria) return [];
  const tests = [];
  for (const [k, crit] of Object.entries(criteria)) {
    const c = Number(k);
    if (!crit || c < rg.c1 || c > rg.c2) continue;
    const colCells = [];
    for (let r = rg.r1 + 1; r <= rg.r2; r++) colCells.push(cellOf(sheet, r, c));
    tests.push([c, makeFilterTest(crit, colCells)]);
  }
  if (!tests.length) return [];
  const hidden = [];
  for (let r = rg.r1 + 1; r <= rg.r2; r++) {
    for (const [c, t] of tests) if (!t(cellOf(sheet, r, c))) { hidden.push(r); break; }
  }
  return hidden;
}

// ---- Fill series (fill handle, Ctrl+D / Ctrl+R) ---------------------------

const SHORT_MONTHS = MONTHS.map((m) => m.slice(0, 3));
const SHORT_DAYS = DAYS.map((d) => d.slice(0, 3));
const NAME_LISTS = [MONTHS, SHORT_MONTHS, DAYS, SHORT_DAYS];

function listItem(text) {
  const t = String(text).trim();
  const lower = t.toLowerCase();
  for (const list of NAME_LISTS) {
    const idx = list.findIndex((x) => x.toLowerCase() === lower);
    if (idx >= 0) {
      const kase = t === t.toUpperCase() && t !== t.toLowerCase() ? "upper" : t === lower ? "lower" : "title";
      return { list, idx, kase };
    }
  }
  return null;
}
const applyCase = (s, kase) => (kase === "upper" ? s.toUpperCase() : kase === "lower" ? s.toLowerCase() : s);

function textNumber(text) {
  const m = /^(.*?)(\d{1,15})(\D*)$/.exec(String(text));
  if (!m || (!m[1] && !m[3])) return null; // a bare number stored as text is not a series
  return { pre: m[1], num: Number(m[2]), width: m[2].length, post: m[3], zero: m[2].length > 1 && m[2][0] === "0" };
}

function kindOf(cell) {
  if (cellIsEmpty(cell)) return "empty";
  if (cell.f) return "formula";
  if (cell.t === "n" && typeof cell.v === "number") return "num";
  if (cell.t === "d" && isoToSerial(cell.v) != null) return "date";
  if (cell.t === "s") return listItem(cell.v) ? "list" : textNumber(cell.v) ? "textnum" : "text";
  return "other";
}

const clean15 = (x) => Number(Number(x).toPrecision(15));

/** Least-squares line through (i, ys[i]); returns p => value. */
function linearFit(ys) {
  const n = ys.length;
  if (n === 1) return () => ys[0];
  const mx = (n - 1) / 2;
  const my = ys.reduce((s, y) => s + y, 0) / n;
  let sxy = 0, sxx = 0;
  ys.forEach((y, i) => { sxy += (i - mx) * (y - my); sxx += (i - mx) * (i - mx); });
  const slope = sxy / sxx;
  return (p) => clean15(my + slope * (p - mx));
}

function addMonthsIso(iso, months) {
  const m = /^(\d{4})-(\d{2})-(\d{2})(.*)$/.exec(iso);
  if (!m) return iso;
  let y = +m[1], mo = +m[2] - 1 + months;
  y += Math.floor(mo / 12);
  mo = ((mo % 12) + 12) % 12;
  const d = Math.min(+m[3], daysInMonth(y, mo + 1));
  return `${pad4(y)}-${pad2(mo + 1)}-${pad2(d)}${m[4]}`;
}

const modn = (p, n) => ((p % n) + n) % n;

/**
 * AutoFill like Excel's fill handle. src: the source cells along the fill
 * axis (in order, nulls allowed); count: number of cells to produce.
 * opts.reverse: filling up/left; result[0] is the cell next to the source
 * (index -1), result[1] index -2, ... ; otherwise result[j] is index n + j.
 * opts.ctrl: Ctrl held while dragging (a single number then counts up, a
 * series is copied instead). opts.translate(f, delta) rewrites formulas
 * copied `delta` cells along the axis.
 *
 * Rules: a single number/text copies (Ctrl: number +1), a single date steps
 * one day, "Item 1" -> "Item 2", weekday / month names (full and short)
 * continue, two or more numbers extend their linear trend (1,2 -> 3,4;
 * 2,4,6 -> 8,10), dates step by their day / month interval; anything else
 * repeats the source pattern (formulas shifted like copy & paste).
 */
export function fillSeries(src, count, opts = {}) {
  const n = src.length;
  const out = [];
  if (!n || count <= 0) return out;
  const pos = (j) => (opts.reverse ? -1 - j : n + j);
  const kinds = src.map(kindOf);
  const same = kinds.every((k) => k === kinds[0]) ? kinds[0] : null;
  const withValue = (k, v) => { const c = cloneCell(src[k]); c.v = v; return c; };
  const ctrl = !!opts.ctrl;

  if (same === "num" && (n > 1 ? !ctrl : ctrl)) {
    const ys = src.map((c) => c.v);
    const fit = n > 1 ? linearFit(ys) : (p) => clean15(ys[0] + p);
    for (let j = 0; j < count; j++) { const p = pos(j); out.push(withValue(modn(p, n), fit(p))); }
    return out;
  }
  if (same === "date" && !ctrl) {
    const isos = src.map((c) => String(c.v));
    const serials = isos.map(isoToSerial);
    let step = null;
    if (n > 1) {
      const parts = isos.map((s) => /^(\d{4})-(\d{2})-(\d{2})/.exec(s)).map((m) => ({ y: +m[1], m: +m[2], d: +m[3] }));
      const months = parts.map((p) => p.y * 12 + p.m - 1);
      const dm = months[1] - months[0];
      const monthly = dm !== 0 && !/T/.test(isos[0]) && parts.every((p) => p.d === parts[0].d || (p.d >= 28 && parts[0].d >= 28))
        && months.every((m, i) => m - months[0] === dm * i);
      if (monthly) step = { months: dm };
      else {
        const dd = serials[1] - serials[0];
        if (serials.every((s, i) => Math.abs(s - serials[0] - dd * i) < 1e-9)) step = { days: dd };
      }
    } else step = { days: 1 };
    if (step) {
      for (let j = 0; j < count; j++) {
        const p = pos(j);
        const v = step.months != null ? addMonthsIso(isos[0], step.months * p)
          : serialToIso(serials[0] + step.days * p, /T/.test(isos[0]));
        out.push(withValue(modn(p, n), v));
      }
      return out;
    }
  }
  if (same === "list" && !ctrl) {
    const items = src.map((c) => listItem(c.v));
    const list = items[0].list, L = list.length;
    if (items.every((it) => it.list === list)) {
      const step = n > 1 ? modn(items[1].idx - items[0].idx, L) : 1;
      if (items.every((it, i) => modn(it.idx - items[0].idx - step * i, L) === 0)) {
        for (let j = 0; j < count; j++) {
          const p = pos(j);
          const k = modn(p, n);
          out.push(withValue(k, applyCase(list[modn(items[0].idx + step * p, L)], items[k].kase)));
        }
        return out;
      }
    }
  }
  if (same === "textnum" && !ctrl) {
    const parts = src.map((c) => textNumber(c.v));
    if (parts.every((p) => p.pre === parts[0].pre && p.post === parts[0].post)) {
      const fit = n > 1 ? linearFit(parts.map((p) => p.num)) : (p) => parts[0].num + p;
      for (let j = 0; j < count; j++) {
        const p = pos(j);
        const num = Math.abs(Math.round(fit(p)));
        const digits = parts[0].zero ? String(num).padStart(parts[0].width, "0") : String(num);
        out.push(withValue(modn(p, n), parts[0].pre + digits + parts[0].post));
      }
      return out;
    }
  }

  // Repeat the pattern (formulas shifted; "Item 1" style text still counts up per cycle).
  for (let j = 0; j < count; j++) {
    const p = pos(j);
    const k = modn(p, n);
    const cycle = Math.floor(p / n);
    const cell = cloneCell(src[k]);
    if (cell && cell.f && opts.translate) {
      cell.f = opts.translate(cell.f, p - k);
      cell.v = null;
      cell.t = "f";
    } else if (cell && !ctrl && kinds[k] === "textnum") {
      const t = textNumber(cell.v);
      const num = Math.abs(t.num + cycle);
      cell.v = t.pre + (t.zero ? String(num).padStart(t.width, "0") : String(num)) + t.post;
    }
    out.push(cell);
  }
  return out;
}

// ---- Charts -----------------------------------------------------------------

export const CHART_TYPES = [
  ["column", "Column"], ["bar", "Bar"], ["line", "Line"], ["pie", "Pie"], ["scatter", "Scatter (XY)"], ["area", "Area"],
];
/** Categorical series colours, assigned in this fixed order. */
export const CHART_COLORS = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300", "#6250d6", "#e34948"];
const MAX_SERIES = 8;
export const GRID_ROW_PX = 24;
export const GRID_DEFAULT_COL_PX = 88;
const EXCEL_ROW_PX = 20;
const EXCEL_DEFAULT_COL_PX = 64;

/** Width in px of column c in the grid / in Excel (chart geometry, see the CONTRACT). */
export function gridColPx(sheet, c) {
  const w = sheet && sheet.colWidths ? sheet.colWidths[indexToCol(c)] : undefined;
  return w != null ? charsToPx(w) : GRID_DEFAULT_COL_PX;
}
export function excelColPx(sheet, c) {
  const w = sheet && sheet.colWidths ? sheet.colWidths[indexToCol(c)] : undefined;
  return w != null ? charsToPx(w) : EXCEL_DEFAULT_COL_PX;
}

/** Grid px (x, y) -> Excel anchor {col, row, dx, dy} (offsets in Excel px). */
export function chartAnchorFromXY(sheet, x, y) {
  let col = 0, left = 0;
  x = Math.max(0, Number(x) || 0);
  while (col < MAX_COLS - 1 && left + gridColPx(sheet, col) <= x) { left += gridColPx(sheet, col); col++; }
  const dx = Math.round((x - left) * excelColPx(sheet, col) / gridColPx(sheet, col));
  y = Math.max(0, Number(y) || 0);
  const row = Math.min(MAX_ROWS - 1, Math.floor(y / GRID_ROW_PX));
  const dy = Math.round((y - row * GRID_ROW_PX) * EXCEL_ROW_PX / GRID_ROW_PX);
  return { col, row, dx, dy };
}

/** Excel anchor (col, offset px, row, offset px) -> grid px {x, y}. */
export function chartXYFromAnchor(sheet, col, colOffPx, row, rowOffPx) {
  let x = 0;
  for (let c = 0; c < col && c < MAX_COLS; c++) x += gridColPx(sheet, c);
  x += Math.round((colOffPx || 0) * gridColPx(sheet, col) / excelColPx(sheet, col));
  const y = row * GRID_ROW_PX + Math.round((rowOffPx || 0) * GRID_ROW_PX / EXCEL_ROW_PX);
  return { x, y };
}

/** Excel x / y position (px) of a column / row plus an offset (anchor -> size). */
export function excelXOf(sheet, col, offPx = 0) {
  let x = 0;
  for (let c = 0; c < col && c < MAX_COLS; c++) x += excelColPx(sheet, c);
  return x + (offPx || 0);
}
export const excelYOf = (row, offPx = 0) => row * EXCEL_ROW_PX + (offPx || 0);

/**
 * Series layout of a chart range (header row + first column categories):
 * {r1, r2, c1, c2, cat: {c}|null, series: [{c}]}.
 */
export function chartLayout(range) {
  const rg = parseRange(range || "");
  if (!rg) return null;
  const series = [];
  if (rg.c2 > rg.c1) for (let c = rg.c1 + 1; c <= rg.c2; c++) series.push({ c });
  else series.push({ c: rg.c1 });
  return { ...rg, cat: rg.c2 > rg.c1 ? { c: rg.c1 } : null, series };
}

/** Data of a chart: {categories:[text], xs:[number|null], series:[{name, values:[number|null]}]}. */
export function chartData(sheet, chart) {
  const L = chartLayout(chart && chart.range);
  if (!L) return { categories: [], xs: [], series: [] };
  const categories = [], xs = [];
  const rowsMax = Math.min(L.r2, L.r1 + 2000);
  for (let r = L.r1 + 1; r <= rowsMax; r++) {
    const cell = L.cat ? cellOf(sheet, r, L.cat.c) : null;
    categories.push(L.cat ? filterText(cell) : String(r - L.r1));
    const v = L.cat ? cellNumericValue(cell) : r - L.r1;
    xs.push(typeof v === "number" && isFinite(v) ? v : null);
  }
  const series = L.series.slice(0, MAX_SERIES).map((s, i) => {
    const values = [];
    for (let r = L.r1 + 1; r <= rowsMax; r++) {
      const v = cellNumericValue(cellOf(sheet, r, s.c));
      values.push(typeof v === "number" && isFinite(v) ? v : null);
    }
    return { name: filterText(cellOf(sheet, L.r1, s.c)) || `Series ${i + 1}`, values };
  });
  return { categories, xs, series };
}

function niceTicks(min, max, count = 5) {
  if (!(isFinite(min) && isFinite(max))) { min = 0; max = 1; }
  if (min === max) { if (min === 0) max = 1; else if (min > 0) min = 0; else max = 0; }
  const raw = (max - min) / count;
  const mag = Math.pow(10, Math.floor(Math.log10(raw)));
  const norm = raw / mag;
  const step = (norm <= 1 ? 1 : norm <= 2 ? 2 : norm <= 2.5 ? 2.5 : norm <= 5 ? 5 : 10) * mag;
  const lo = Math.floor(min / step) * step, hi = Math.ceil(max / step) * step;
  const ticks = [];
  for (let v = lo; v <= hi + step / 2 && ticks.length < 50; v += step) ticks.push(clean15(v));
  return { lo, hi, ticks };
}

function tickLabel(v) {
  const a = Math.abs(v);
  if (a >= 1e9) return formatGeneral(v / 1e9, 4) + "B";
  if (a >= 1e6) return formatGeneral(v / 1e6, 4) + "M";
  if (a >= 1e4) return formatGeneral(v / 1e3, 4) + "K";
  return formatGeneral(v, 6);
}

const svgEsc = (s) => escapeHtmlText(String(s)).replace(/"/g, "&quot;");
const trunc = (s, n) => (String(s).length > n ? String(s).slice(0, n - 1) + "…" : String(s));
const f1 = (x) => (Math.round(x * 10) / 10).toString();

/**
 * Render a chart as a standalone SVG string (width x height px). Marks carry
 * <title> tooltips; two or more series get a legend; colours follow the
 * series order (never cycled: at most 8 series / pie slices, the rest of a
 * pie folds into "Other").
 */
export function chartSvg(sheet, chart, width, height) {
  const W = Math.max(120, Math.round(width || chart.w || 480)), H = Math.max(90, Math.round(height || chart.h || 288));
  const data = chartData(sheet, chart);
  const type = CHART_TYPES.some(([t]) => t === chart.type) ? chart.type : "column";
  const ink = "#0b0b0b", ink2 = "#52514e", gridc = "#e4e3df";
  const open = `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" font-family="Segoe UI, Calibri, Arial, sans-serif"><rect width="${W}" height="${H}" fill="#ffffff"/>`;
  let body = "";
  let top = 10;
  const title = String(chart.title || "").trim();
  if (title) { body += `<text x="${W / 2}" y="24" text-anchor="middle" font-size="15" font-weight="600" fill="${ink}">${svgEsc(trunc(title, 80))}</text>`; top = 36; }
  const series = data.series;
  if (!series.length || !series.some((s) => s.values.some((v) => v != null))) {
    body += `<text x="${W / 2}" y="${H / 2}" text-anchor="middle" font-size="13" fill="${ink2}">No numbers to chart in ${svgEsc(chart.range || "this range")}</text>`;
    return `${open}${body}</svg>`;
  }

  if (type === "pie") {
    const s = series[0];
    let items = s.values.map((v, i) => [data.categories[i] || `${i + 1}`, v]).filter(([, v]) => v != null && v > 0);
    let other = false;
    if (items.length > MAX_SERIES) {
      const rest = items.slice(MAX_SERIES - 1).reduce((a, [, v]) => a + v, 0);
      items = items.slice(0, MAX_SERIES - 1).concat([["Other", rest]]);
      other = true;
    }
    const total = items.reduce((a, [, v]) => a + v, 0);
    const legendW = Math.min(170, W * 0.38);
    const cx = (W - legendW) / 2 + 4, cy = top + (H - top - 8) / 2;
    const R = Math.max(10, Math.min((W - legendW) / 2 - 12, (H - top - 16) / 2));
    let a0 = -Math.PI / 2;
    items.forEach(([name, v], i) => {
      const col = other && i === items.length - 1 ? "#8a8984" : CHART_COLORS[i];
      const frac = v / total;
      const a1 = a0 + frac * Math.PI * 2;
      const tip = `<title>${svgEsc(`${name}: ${formatGeneral(v)} (${(frac * 100).toFixed(1)}%)`)}</title>`;
      if (frac >= 0.9999) body += `<circle cx="${f1(cx)}" cy="${f1(cy)}" r="${f1(R)}" fill="${col}" stroke="#fff" stroke-width="2">${tip}</circle>`;
      else {
        const large = a1 - a0 > Math.PI ? 1 : 0;
        const p = (a) => `${f1(cx + R * Math.cos(a))},${f1(cy + R * Math.sin(a))}`;
        body += `<path d="M${f1(cx)},${f1(cy)} L${p(a0)} A${f1(R)},${f1(R)} 0 ${large} 1 ${p(a1)} Z" fill="${col}" stroke="#fff" stroke-width="2">${tip}</path>`;
      }
      if (frac > 0.06) {
        const am = (a0 + a1) / 2;
        body += `<text x="${f1(cx + R * 0.62 * Math.cos(am))}" y="${f1(cy + R * 0.62 * Math.sin(am) + 4)}" text-anchor="middle" font-size="11" font-weight="600" fill="#fff" pointer-events="none">${Math.round(frac * 100)}%</text>`;
      }
      const ly = top + 6 + i * 18;
      if (ly < H - 6) body += `<rect x="${f1(W - legendW + 6)}" y="${ly}" width="10" height="10" rx="2" fill="${col}"/><text x="${f1(W - legendW + 22)}" y="${ly + 9}" font-size="11" fill="${ink2}">${svgEsc(trunc(name, 22))}</text>`;
      a0 = a1;
    });
    return `${open}${body}</svg>`;
  }

  // legend (2+ series): one centred row, wrapping when long
  let bottomPad = 8;
  if (series.length >= 2) {
    const items = series.map((s, i) => [trunc(s.name, 18), CHART_COLORS[i]]);
    const itemW = items.map(([n]) => Math.min(140, 24 + n.length * 6.6));
    const total = itemW.reduce((s, w) => s + w, 0);
    const rows = total > W - 16 ? 2 : 1;
    let x = rows === 1 ? Math.max(8, (W - total) / 2) : 8, y = H - (rows === 1 ? 10 : 26);
    items.forEach(([n, col], i) => {
      if (x + itemW[i] > W - 4) { x = 8; y += 16; }
      body += `<rect x="${f1(x)}" y="${y - 9}" width="10" height="10" rx="2" fill="${col}"/><text x="${f1(x + 14)}" y="${y}" font-size="11" fill="${ink2}">${svgEsc(n)}</text>`;
      x += itemW[i];
    });
    bottomPad = rows === 1 ? 26 : 42;
  }

  const horizontal = type === "bar";
  const scatter = type === "scatter";
  const allVals = series.flatMap((s) => s.values).filter((v) => v != null);
  let vmin = Math.min(...allVals), vmax = Math.max(...allVals);
  if (type !== "line" && type !== "scatter") { vmin = Math.min(0, vmin); vmax = Math.max(0, vmax); }
  const vt = niceTicks(vmin, vmax, Math.max(2, Math.min(6, Math.floor((horizontal ? W : H) / 60))));
  const labelW = Math.max(...vt.ticks.map((t) => tickLabel(t).length)) * 6.5 + 10;
  const ncat = data.categories.length;
  const catLabelW = horizontal ? Math.min(W * 0.3, Math.max(...data.categories.map((c) => trunc(c, 16).length)) * 6.3 + 10) : 0;
  const L = 8 + (horizontal ? catLabelW : labelW), R = W - 14, T = top + 4, B = H - bottomPad - 18;
  if (R - L < 20 || B - T < 20) return `${open}${body}</svg>`;
  const vpos = (v) => (horizontal ? L + ((v - vt.lo) / (vt.hi - vt.lo)) * (R - L) : B - ((v - vt.lo) / (vt.hi - vt.lo)) * (B - T));
  for (const t of vt.ticks) {
    const p = f1(vpos(t));
    if (horizontal) body += `<line x1="${p}" y1="${T}" x2="${p}" y2="${B}" stroke="${gridc}"/><text x="${p}" y="${B + 14}" text-anchor="middle" font-size="10.5" fill="${ink2}">${tickLabel(t)}</text>`;
    else body += `<line x1="${f1(L)}" y1="${p}" x2="${R}" y2="${p}" stroke="${gridc}"/><text x="${f1(L - 6)}" y="${f1(vpos(t) + 3.5)}" text-anchor="end" font-size="10.5" fill="${ink2}">${tickLabel(t)}</text>`;
  }
  let xpos, band = 0;
  if (scatter) {
    const xs = data.xs.map((x, i) => (x == null ? i + 1 : x));
    const xt = niceTicks(Math.min(...xs), Math.max(...xs), Math.max(2, Math.min(6, Math.floor(W / 80))));
    xpos = (i) => L + ((xs[i] - xt.lo) / (xt.hi - xt.lo)) * (R - L);
    for (const t of xt.ticks) {
      const p = f1(L + ((t - xt.lo) / (xt.hi - xt.lo)) * (R - L));
      body += `<line x1="${p}" y1="${T}" x2="${p}" y2="${B}" stroke="${gridc}"/><text x="${p}" y="${B + 14}" text-anchor="middle" font-size="10.5" fill="${ink2}">${tickLabel(t)}</text>`;
    }
  } else {
    const span = horizontal ? B - T : R - L;
    band = span / Math.max(1, ncat);
    const every = Math.max(1, Math.ceil(ncat / Math.max(1, Math.floor(span / (horizontal ? 16 : 56)))));
    xpos = (i) => (horizontal ? T : L) + band * (i + 0.5);
    data.categories.forEach((c, i) => {
      if (i % every) return;
      const p = xpos(i);
      if (horizontal) body += `<text x="${f1(L - 6)}" y="${f1(p + 3.5)}" text-anchor="end" font-size="10.5" fill="${ink2}">${svgEsc(trunc(c, 16))}</text>`;
      else body += `<text x="${f1(p)}" y="${B + 14}" text-anchor="middle" font-size="10.5" fill="${ink2}">${svgEsc(trunc(c, Math.max(3, Math.floor((band * every) / 6.5))))}</text>`;
    });
  }
  const base = vpos(Math.max(vt.lo, Math.min(vt.hi, 0)));
  if (horizontal) body += `<line x1="${f1(base)}" y1="${T}" x2="${f1(base)}" y2="${B}" stroke="#8a8984"/>`;
  else body += `<line x1="${f1(L)}" y1="${f1(base)}" x2="${R}" y2="${f1(base)}" stroke="#8a8984"/>`;

  const tip = (s, i, v) => `<title>${svgEsc(`${s.name}, ${scatter ? `x ${formatGeneral(data.xs[i] ?? i + 1)}` : data.categories[i]}: ${formatGeneral(v)}`)}</title>`;
  if (type === "column" || type === "bar") {
    const ns = series.length;
    const inner = band * 0.72;
    const bw = Math.max(1, (inner - (ns - 1) * 2) / ns);
    series.forEach((s, si) => {
      const col = CHART_COLORS[si];
      s.values.forEach((v, i) => {
        if (v == null) return;
        const a = xpos(i) - inner / 2 + si * (bw + 2), b = a + bw;
        const p = vpos(v);
        const rr = Math.min(4, Math.abs(p - base) / 2, bw / 2);
        const sg = (horizontal ? p >= base : p <= base) ? 1 : -1; // direction of the value end
        let d;
        if (horizontal) {
          d = `M${f1(base)},${f1(a)} H${f1(p - sg * rr)} Q${f1(p)},${f1(a)} ${f1(p)},${f1(a + rr)} V${f1(b - rr)} Q${f1(p)},${f1(b)} ${f1(p - sg * rr)},${f1(b)} H${f1(base)} Z`;
        } else {
          d = `M${f1(a)},${f1(base)} V${f1(p + sg * rr)} Q${f1(a)},${f1(p)} ${f1(a + rr)},${f1(p)} H${f1(b - rr)} Q${f1(b)},${f1(p)} ${f1(b)},${f1(p + sg * rr)} V${f1(base)} Z`;
        }
        body += `<path d="${d}" fill="${col}">${tip(s, i, v)}</path>`;
      });
    });
  } else {
    series.forEach((s, si) => {
      const col = CHART_COLORS[si];
      const pts = s.values.map((v, i) => (v == null ? null : [xpos(i), vpos(v)]));
      if (!scatter) {
        const segs = [];
        let seg = [];
        pts.forEach((p) => { if (p) seg.push(p); else if (seg.length) { segs.push(seg); seg = []; } });
        if (seg.length) segs.push(seg);
        for (const sg of segs) {
          const line = sg.map((p) => `${f1(p[0])},${f1(p[1])}`).join(" ");
          if (type === "area") body += `<polygon points="${f1(sg[0][0])},${f1(base)} ${line} ${f1(sg[sg.length - 1][0])},${f1(base)}" fill="${col}" fill-opacity="0.28"/>`;
          body += `<polyline points="${line}" fill="none" stroke="${col}" stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>`;
        }
      }
      const markers = scatter || pts.length <= 40;
      pts.forEach((p, i) => {
        if (!p) return;
        body += `<circle cx="${f1(p[0])}" cy="${f1(p[1])}" r="${scatter ? 4.5 : markers ? 3.5 : 6}" fill="${col}"${markers ? ' stroke="#fff" stroke-width="1.5"' : ' fill-opacity="0"'}>${tip(s, i, s.values[i])}</circle>`;
      });
    });
  }
  return `${open}${body}</svg>`;
}

/** Shift a range ref "A1:C10" by dr rows / dc columns (null when it leaves the sheet). */
export function offsetRef(ref, dr, dc) {
  const rg = parseRange(ref || "");
  if (!rg) return null;
  const n = { r1: rg.r1 + dr, r2: rg.r2 + dr, c1: rg.c1 + dc, c2: rg.c2 + dc };
  if (n.r1 < 0 || n.c1 < 0 || n.r2 >= MAX_ROWS || n.c2 >= MAX_COLS) return null;
  return rangeToRef(n);
}
