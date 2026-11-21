/**
 * Formula engine for the Excel model: tokenizer, parser, evaluator,
 * whole-book recalculation and reference rewriting (copy/paste translation,
 * row/column insertion and deletion, sheet rename/delete).
 *
 * Values inside the evaluator:
 *   number | string | boolean | null (blank) | XErr | Ref | Matrix (array of rows)
 *
 * recalcBook(model) recomputes every formula cell (cell.f) in every sheet and
 * stores the result in cell.v / cell.t (cell.f is kept). Formulas that use
 * unsupported functions or names keep their cached value (or become #NAME?).
 * Circular references evaluate to "#CIRC!".
 */

import {
  rcToAddr, indexToCol, colToIndex, MAX_ROWS, MAX_COLS,
  isoToSerial, serialToIso, serialToParts, ymdToSerial,
  formatValue, formatGeneral, parseCellInput, roundDecimal, ERROR_VALUES,
  isDateFormat, isDateOnlyOrDateTimeFormat,
} from "./model-excel.js";

// ---------------------------------------------------------------------------
// Errors and value helpers
// ---------------------------------------------------------------------------

export class XErr {
  constructor(code) { this.code = code; }
  toString() { return this.code; }
}
const ERR = Object.fromEntries(ERROR_VALUES.map((e) => [e, new XErr(e)]));
const DIV0 = ERR["#DIV/0!"], VALUE = ERR["#VALUE!"], REF = ERR["#REF!"], NAME = ERR["#NAME?"];
const NUM = ERR["#NUM!"], NA = ERR["#N/A"], CIRC = ERR["#CIRC!"];

export class Ref {
  constructor(sheet, r1, c1, r2, c2) {
    this.sheet = sheet; this.r1 = r1; this.c1 = c1; this.r2 = r2; this.c2 = c2;
  }
  get rows() { return this.r2 - this.r1 + 1; }
  get cols() { return this.c2 - this.c1 + 1; }
  get single() { return this.r1 === this.r2 && this.c1 === this.c2; }
}

/** Thrown when a formula uses something we can't evaluate (keep cached value). */
class Unsupported extends Error {}
/** Thrown when evaluation nests too deep; the engine evaluates `cell` first and retries. */
class TooDeep extends Error {
  constructor(si, key, cell) { super("deep"); this.si = si; this.key = key; this.cell = cell; }
}
class ParseError extends Error {}

const isErr = (v) => v instanceof XErr;

// ---------------------------------------------------------------------------
// Tokenizer
// ---------------------------------------------------------------------------

const REF_END = "(?![A-Za-z0-9_.(!\\[])";
const RE_AREA = new RegExp(`^(\\$?)([A-Za-z]{1,3})(\\$?)(\\d{1,7})(?::(\\$?)([A-Za-z]{1,3})(\\$?)(\\d{1,7}))?${REF_END}`);
const RE_COLS = new RegExp(`^(\\$?)([A-Za-z]{1,3}):(\\$?)([A-Za-z]{1,3})${REF_END}`);
const RE_ROWS = new RegExp(`^(\\$?)(\\d{1,7}):(\\$?)(\\d{1,7})${REF_END}`);
const RE_QSHEET = /^'((?:[^']|'')+)'!/;
const RE_SHEET = /^([A-Za-z_¡-￿][A-Za-z0-9_.¡-￿]*)!/;
const RE_NUM = /^(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?/;
const RE_IDENT = /^[A-Za-z_\\¡-￿][A-Za-z0-9_.¡-￿]*/;

function normFuncName(name) {
  return name.toUpperCase().replace(/^(_XLFN\.)?(_XLWS\.)?/, "").replace(/^_XLFN\./, "");
}

function makeRefToken(text, sheet, sheetText, kind, a, b) {
  return { k: "ref", text, sheet, sheetText, kind, a, b };
}

/** Tokenize formula text (without the leading "="). Whitespace tokens are kept. */
export function tokenize(src) {
  const toks = [];
  let i = 0;
  const lastSig = () => {
    for (let j = toks.length - 1; j >= 0; j--) if (toks[j].k !== "ws") return toks[j];
    return null;
  };
  const expectOperand = () => {
    const t = lastSig();
    return !t || (t.k === "op" && t.text !== "%") || t.k === "(" || t.k === "," || t.k === ";" || t.k === "{";
  };
  while (i < src.length) {
    const ch = src[i];
    const rest = src.slice(i);
    let m;
    if (/\s/.test(ch)) {
      m = /^\s+/.exec(rest);
      toks.push({ k: "ws", text: m[0] });
      i += m[0].length;
      continue;
    }
    if (ch === '"') {
      let j = i + 1, v = "";
      for (;;) {
        if (j >= src.length) throw new ParseError("Unterminated string");
        if (src[j] === '"') {
          if (src[j + 1] === '"') { v += '"'; j += 2; continue; }
          break;
        }
        v += src[j++];
      }
      toks.push({ k: "str", text: src.slice(i, j + 1), v });
      i = j + 1;
      continue;
    }
    if (ch === "#") {
      const up = rest.toUpperCase();
      const e = ERROR_VALUES.find((x) => up.startsWith(x));
      if (!e) throw new ParseError("Bad error literal");
      toks.push({ k: "err", text: rest.slice(0, e.length), v: e });
      i += e.length;
      continue;
    }
    if ("(),;{}".includes(ch)) { toks.push({ k: ch, text: ch }); i++; continue; }
    if ((m = /^(<>|<=|>=|[-+*/^&=<>%:])/.exec(rest))) {
      toks.push({ k: "op", text: m[1] });
      i += m[1].length;
      continue;
    }
    // Sheet-qualified references.
    let sheet = null, sheetText = "";
    if ((m = RE_QSHEET.exec(rest)) || (m = RE_SHEET.exec(rest))) {
      sheetText = m[0];
      sheet = m[0][0] === "'" ? m[1].replace(/''/g, "'") : m[1];
    }
    const after = rest.slice(sheetText.length);
    const operandPos = expectOperand() || sheetText;
    if (sheetText && /^#REF!/i.test(after)) {
      toks.push({ k: "err", text: sheetText + after.slice(0, 5), v: "#REF!" });
      i += sheetText.length + 5;
      continue;
    }
    if ((m = RE_AREA.exec(after))) {
      const a = { ac: !!m[1], c: colToIndex(m[2]), ar: !!m[3], r: +m[4] - 1 };
      const b = m[6] ? { ac: !!m[5], c: colToIndex(m[6]), ar: !!m[7], r: +m[8] - 1 } : null;
      if (a.c < MAX_COLS && a.r >= 0 && a.r < MAX_ROWS && (!b || (b.c < MAX_COLS && b.r >= 0 && b.r < MAX_ROWS))) {
        toks.push(makeRefToken(sheetText + m[0], sheet, sheetText, b ? "area" : "cell", a, b));
        i += sheetText.length + m[0].length;
        continue;
      }
    }
    if ((m = RE_COLS.exec(after)) && operandPos) {
      const a = { ac: !!m[1], c: colToIndex(m[2]), ar: false, r: null };
      const b = { ac: !!m[3], c: colToIndex(m[4]), ar: false, r: null };
      if (a.c < MAX_COLS && b.c < MAX_COLS) {
        toks.push(makeRefToken(sheetText + m[0], sheet, sheetText, "cols", a, b));
        i += sheetText.length + m[0].length;
        continue;
      }
    }
    if ((m = RE_ROWS.exec(after)) && operandPos) {
      const a = { ac: false, c: null, ar: !!m[1], r: +m[2] - 1 };
      const b = { ac: false, c: null, ar: !!m[3], r: +m[4] - 1 };
      if (a.r >= 0 && b.r >= 0 && a.r < MAX_ROWS && b.r < MAX_ROWS) {
        toks.push(makeRefToken(sheetText + m[0], sheet, sheetText, "rows", a, b));
        i += sheetText.length + m[0].length;
        continue;
      }
    }
    if (sheetText) {
      // Sheet prefix followed by a name/structured ref: unsupported but tokenizable.
      if ((m = RE_IDENT.exec(after))) {
        toks.push({ k: "name", text: sheetText + m[0], v: m[0] });
        i += sheetText.length + m[0].length;
        continue;
      }
      throw new ParseError("Bad reference");
    }
    if ((m = RE_NUM.exec(rest))) {
      toks.push({ k: "num", text: m[0], v: Number(m[0]) });
      i += m[0].length;
      continue;
    }
    if ((m = RE_IDENT.exec(rest))) {
      const word = m[0];
      if (src[i + word.length] === "(") {
        toks.push({ k: "func", text: word, v: normFuncName(word) });
      } else if (/^(TRUE|FALSE)$/i.test(word)) {
        toks.push({ k: "bool", text: word, v: /^TRUE$/i.test(word) });
      } else if (src[i + word.length] === "[") {
        // Structured table reference: Table1[Column] / Table1[[#This Row],[Col]]
        let depth = 0, j = i + word.length;
        for (; j < src.length; j++) {
          if (src[j] === "[") depth++;
          else if (src[j] === "]" && --depth === 0) break;
        }
        toks.push({ k: "name", text: src.slice(i, j + 1), v: src.slice(i, j + 1) });
        i = j + 1;
        continue;
      } else {
        toks.push({ k: "name", text: word, v: word });
      }
      i += word.length;
      continue;
    }
    if (ch === "[") {
      // [Book.xlsx]Sheet!A1 external references: unsupported
      throw new ParseError("External references are not supported");
    }
    throw new ParseError(`Unexpected character ${ch}`);
  }
  return toks;
}

// ---------------------------------------------------------------------------
// Parser
// ---------------------------------------------------------------------------

function refNodeFromToken(t) {
  let r1, c1, r2, c2;
  if (t.kind === "cols") { r1 = 0; r2 = MAX_ROWS - 1; c1 = t.a.c; c2 = t.b.c; }
  else if (t.kind === "rows") { c1 = 0; c2 = MAX_COLS - 1; r1 = t.a.r; r2 = t.b.r; }
  else {
    r1 = t.a.r; c1 = t.a.c;
    r2 = t.b ? t.b.r : r1; c2 = t.b ? t.b.c : c1;
  }
  return {
    t: "ref", sheet: t.sheet,
    r1: Math.min(r1, r2), r2: Math.max(r1, r2), c1: Math.min(c1, c2), c2: Math.max(c1, c2),
    cell: t.kind === "cell",
  };
}

function parseTokens(all) {
  const toks = all.filter((t) => t.k !== "ws");
  let p = 0;
  const peek = () => toks[p];
  const isOp = (op) => toks[p] && toks[p].k === "op" && toks[p].text === op;
  const expect = (k) => {
    if (!toks[p] || toks[p].k !== k) throw new ParseError(`Expected ${k}`);
    return toks[p++];
  };

  function parseExpr() { return parseCompare(); }
  function parseCompare() {
    let a = parseConcat();
    while (toks[p] && toks[p].k === "op" && ["=", "<>", "<", ">", "<=", ">="].includes(toks[p].text)) {
      const op = toks[p++].text;
      a = { t: "bin", op, a, b: parseConcat() };
    }
    return a;
  }
  function parseConcat() {
    let a = parseAdd();
    while (isOp("&")) { p++; a = { t: "bin", op: "&", a, b: parseAdd() }; }
    return a;
  }
  function parseAdd() {
    let a = parseMul();
    while (isOp("+") || isOp("-")) { const op = toks[p++].text; a = { t: "bin", op, a, b: parseMul() }; }
    return a;
  }
  function parseMul() {
    let a = parsePow();
    while (isOp("*") || isOp("/")) { const op = toks[p++].text; a = { t: "bin", op, a, b: parsePow() }; }
    return a;
  }
  function parsePow() {
    let a = parsePct();
    while (isOp("^")) { p++; a = { t: "bin", op: "^", a, b: parsePct() }; }
    return a;
  }
  function parsePct() {
    let a = parseUnary();
    while (isOp("%")) { p++; a = { t: "pct", a }; }
    return a;
  }
  function parseUnary() {
    if (isOp("-")) { p++; return { t: "neg", a: parseUnary() }; }
    if (isOp("+")) { p++; return parseUnary(); }
    return parsePrimary();
  }
  function parsePrimary() {
    const t = peek();
    if (!t) throw new ParseError("Unexpected end of formula");
    switch (t.k) {
      case "num": p++; return { t: "num", v: t.v };
      case "str": p++; return { t: "str", v: t.v };
      case "bool": p++; return { t: "bool", v: t.v };
      case "err": p++; return { t: "err", v: t.v };
      case "ref": {
        p++;
        const node = refNodeFromToken(t);
        if (isOp(":")) throw new ParseError("Range operator on non-reference");
        return node;
      }
      case "name": p++; return { t: "name", v: t.v };
      case "(": {
        p++;
        const e = parseExpr();
        expect(")");
        return e;
      }
      case "{": {
        p++;
        const rows = [[]];
        for (;;) {
          const v = parseUnary();
          rows[rows.length - 1].push(v);
          if (peek() && peek().k === ",") { p++; continue; }
          if (peek() && peek().k === ";") { p++; rows.push([]); continue; }
          expect("}");
          break;
        }
        return { t: "arr", rows };
      }
      case "func": {
        p++;
        expect("(");
        const args = [];
        if (peek() && peek().k === ")") { p++; return { t: "fn", name: t.v, args }; }
        for (;;) {
          if (peek() && (peek().k === "," || peek().k === ")")) args.push({ t: "missing" });
          else args.push(parseExpr());
          if (peek() && peek().k === ",") { p++; continue; }
          expect(")");
          break;
        }
        return { t: "fn", name: t.v, args };
      }
      default:
        throw new ParseError(`Unexpected ${t.text}`);
    }
  }
  const ast = parseExpr();
  if (p < toks.length) throw new ParseError(`Unexpected ${toks[p].text}`);
  return ast;
}

const astCache = new Map();

/** Parse formula text ("=..." or without "="). Returns {ast} or {error}. Cached. */
export function parseFormula(f) {
  const src = String(f).replace(/^=/, "");
  let hit = astCache.get(src);
  if (hit) return hit;
  try {
    hit = { ast: parseTokens(tokenize(src)) };
  } catch (e) {
    hit = { error: e.message || String(e) };
  }
  if (astCache.size > 50000) astCache.clear();
  astCache.set(src, hit);
  return hit;
}

// ---------------------------------------------------------------------------
// Coercions & comparison
// ---------------------------------------------------------------------------

function toNum(v) {
  if (typeof v === "number") return v;
  if (v == null) return 0;
  if (typeof v === "boolean") return v ? 1 : 0;
  if (isErr(v)) return v;
  if (typeof v === "string") {
    const t = v.trim();
    if (!t) return VALUE;
    const p = parseCellInput(t);
    if (p && p.t === "n") return p.v;
    if (p && p.t === "d") return isoToSerial(p.v);
    // ISO-like "2006/1/2" dates (accepted by Excel in every locale)
    const m = /^(\d{4})[/-](\d{1,2})[/-](\d{1,2})$/.exec(t);
    if (m && +m[2] >= 1 && +m[2] <= 12 && +m[3] >= 1 && +m[3] <= 31) return ymdToSerial(+m[1], +m[2], +m[3]);
    return VALUE;
  }
  return VALUE;
}

function toStr(v) {
  if (typeof v === "string") return v;
  if (v == null) return "";
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
  if (typeof v === "number") return formatGeneral(v, 15);
  if (isErr(v)) return v;
  return String(v);
}

function toBool(v) {
  if (typeof v === "boolean") return v;
  if (typeof v === "number") return v !== 0;
  if (v == null) return false;
  if (isErr(v)) return v;
  if (typeof v === "string") {
    if (/^true$/i.test(v.trim())) return true;
    if (/^false$/i.test(v.trim())) return false;
  }
  return VALUE;
}

const round15 = (x) => (isFinite(x) && x !== 0 ? Number(x.toPrecision(15)) : x);
const typeRank = (v) => (typeof v === "number" ? 0 : typeof v === "string" ? 1 : typeof v === "boolean" ? 2 : 3);

/** Excel comparison: numbers < text < booleans; text is case-insensitive. */
function compare(a, b) {
  if (a == null) a = typeof b === "string" ? "" : typeof b === "boolean" ? false : 0;
  if (b == null) b = typeof a === "string" ? "" : typeof a === "boolean" ? false : 0;
  const ra = typeRank(a), rb = typeRank(b);
  if (ra !== rb) return ra < rb ? -1 : 1;
  if (ra === 0) { const x = round15(a), y = round15(b); return x < y ? -1 : x > y ? 1 : 0; }
  if (ra === 1) {
    const x = a.toLowerCase(), y = b.toLowerCase();
    return x < y ? -1 : x > y ? 1 : 0;
  }
  return a === b ? 0 : a ? 1 : -1;
}

function wildcardRegex(pattern) {
  let re = "";
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i];
    if (ch === "~" && i + 1 < pattern.length) { re += pattern[++i].replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); continue; }
    if (ch === "*") re += "[\\s\\S]*";
    else if (ch === "?") re += "[\\s\\S]";
    else re += ch.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  }
  return new RegExp(`^${re}$`, "i");
}
const hasWildcard = (s) => /[*?~]/.test(s);

/** Build a predicate for SUMIF/COUNTIF style criteria. */
function makeCriteria(crit) {
  if (isErr(crit)) return () => false;
  if (typeof crit === "number") {
    return (v) => (typeof v === "number" && round15(v) === round15(crit)) ||
      (typeof v === "string" && v.trim() !== "" && toNum(v) === crit);
  }
  if (typeof crit === "boolean") return (v) => v === crit;
  const s = crit == null ? "" : String(crit);
  const m = /^(<=|>=|<>|<|>|=)?([\s\S]*)$/.exec(s);
  const op = m[1] || "=";
  const operand = m[2];
  const p = operand === "" ? null : parseCellInput(operand);
  const numeric = p && (p.t === "n" || p.t === "d") ? (p.t === "d" ? isoToSerial(p.v) : p.v) : null;
  const boolOperand = p && p.t === "b" ? p.v : null;
  if (numeric != null) {
    return (v) => {
      let x = v;
      if (typeof x !== "number") {
        if (op === "<>") return true;
        if (op !== "=" || typeof x !== "string") return false;
        x = toNum(x);
        if (typeof x !== "number") return false;
      }
      const c = compare(x, numeric);
      return op === "=" ? c === 0 : op === "<>" ? c !== 0 : op === "<" ? c < 0 : op === ">" ? c > 0 : op === "<=" ? c <= 0 : c >= 0;
    };
  }
  if (boolOperand != null) {
    return (v) => (op === "<>" ? v !== boolOperand : op === "=" ? v === boolOperand : false);
  }
  if (op === "=" || op === "<>") {
    if (operand === "") {
      const blank = (v) => v == null || v === "";
      return op === "=" ? blank : (v) => !blank(v);
    }
    const re = hasWildcard(operand) ? wildcardRegex(operand) : null;
    const lower = operand.toLowerCase();
    const eq = (v) => typeof v === "string" && (re ? re.test(v) : v.toLowerCase() === lower);
    return op === "=" ? eq : (v) => !eq(v);
  }
  return (v) => {
    if (typeof v !== "string") return false;
    const c = compare(v, operand);
    return op === "<" ? c < 0 : op === ">" ? c > 0 : op === "<=" ? c <= 0 : c >= 0;
  };
}

// ---------------------------------------------------------------------------
// Engine
// ---------------------------------------------------------------------------

const MAX_DEPTH = 120;
const KEEP = Symbol("keep-cached");
const indexCache = new WeakMap(); // cells object -> {formulas:Set, sorted, maxR, maxC, count}

function rcOfKey(key) {
  let i = 0, c = 0;
  while (i < key.length) {
    const code = key.charCodeAt(i);
    if (code < 65) break;
    c = c * 26 + ((code & ~32) - 64);
    i++;
  }
  return { r: parseInt(key.slice(i), 10) - 1, c: c - 1 };
}

function buildIndex(cells) {
  const formulas = new Set();
  let maxR = -1, maxC = -1, count = 0;
  for (const key in cells) {
    const cell = cells[key];
    if (!cell) continue;
    count++;
    const { r, c } = rcOfKey(key);
    if (r > maxR) maxR = r;
    if (c > maxC) maxC = c;
    if (cell.f) formulas.add(key);
  }
  const idx = { formulas, sorted: null, maxR, maxC, count };
  indexCache.set(cells, idx);
  return idx;
}

function sheetIndex(sheet) {
  if (!sheet.cells) sheet.cells = {};
  return indexCache.get(sheet.cells) || buildIndex(sheet.cells);
}

/**
 * Tell the engine which cells of a sheet changed (keeps the per-sheet formula
 * index incremental so recalc does not have to scan huge sheets).
 */
export function noteCellsChanged(sheet, addrs) {
  const idx = indexCache.get(sheet.cells);
  if (!idx) return;
  for (const key of addrs) {
    const cell = sheet.cells[key];
    const { r, c } = rcOfKey(key);
    if (cell) {
      if (r > idx.maxR) idx.maxR = r;
      if (c > idx.maxC) idx.maxC = c;
    }
    const had = idx.formulas.has(key);
    if (cell && cell.f) { if (!had) { idx.formulas.add(key); idx.sorted = null; } }
    else if (had) { idx.formulas.delete(key); idx.sorted = null; }
  }
  idx.count = Math.max(idx.count, 0);
}

/**
 * Last used row/column of a sheet ({r2, c2}, -1 when empty), from the cached
 * per-sheet index (cheap after the first call; kept current by noteCellsChanged).
 * Bounds only grow on edits; call invalidateSheet() to recompute exactly.
 */
export function sheetBounds(sheet) {
  const idx = sheetIndex(sheet);
  return { r2: idx.maxR, c2: idx.maxC };
}

/** Forget cached indexes for a sheet (after replacing its cells wholesale). */
export function invalidateSheet(sheet) {
  if (sheet && sheet.cells) indexCache.delete(sheet.cells);
}

function sortedFormulaKeys(idx) {
  if (!idx.sorted) {
    idx.sorted = [...idx.formulas].map((k) => ({ k, ...rcOfKey(k) }))
      .sort((x, y) => x.r - y.r || x.c - y.c).map((x) => x.k);
  }
  return idx.sorted;
}

function cellToValue(cell) {
  if (!cell) return null;
  const v = cell.v;
  switch (cell.t) {
    case "n": return typeof v === "number" ? v : v == null ? null : (typeof toNum(v) === "number" ? toNum(v) : 0);
    case "s": return v == null ? null : String(v);
    case "b": return !!v;
    case "e": return ERR[v] || new XErr(String(v));
    case "d": { const s = isoToSerial(v); return s == null ? (v == null ? null : String(v)) : s; }
    case "f": return null;
    default: return v == null ? null : v;
  }
}

class Engine {
  constructor(book, opts = {}) {
    this.book = book;
    this.sheets = book.sheets || [];
    this.byName = new Map(this.sheets.map((s, i) => [String(s.name).toLowerCase(), i]));
    this.now = opts.now || new Date();
    this.done = new Set();
    this.inProgress = new Set();
    this.depth = 0;
    this.cycles = [];
    this.matrixCache = new Map();
    this.lookupCache = new Map();
    this.colLetters = [];
  }

  col(c) {
    return this.colLetters[c] || (this.colLetters[c] = indexToCol(c));
  }

  sheetIdxOf(name, host) {
    if (name == null) return host;
    const i = this.byName.get(String(name).toLowerCase());
    return i == null ? -1 : i;
  }

  // ---- cell access ----
  valueAt(si, r, c) {
    const sheet = this.sheets[si];
    if (!sheet) return REF;
    const key = this.col(c) + (r + 1);
    return this.valueOf(si, key, sheet.cells[key]);
  }

  valueOf(si, key, cell) {
    if (!cell) return null;
    if (cell.f && !this.done.has(cell)) {
      if (this.inProgress.has(cell)) {
        this.cycles.push(`${this.sheets[si].name}!${key}`);
        return CIRC;
      }
      this.evalCell(si, key, cell);
    }
    return cellToValue(cell);
  }

  evalCell(si, key, cell) {
    if (this.depth >= MAX_DEPTH) throw new TooDeep(si, key, cell);
    this.inProgress.add(cell);
    this.depth++;
    let res;
    try {
      const parsed = parseFormula(cell.f);
      if (parsed.error) res = KEEP;
      else {
        const { r, c } = rcOfKey(key);
        // Resolve references while the cell is still marked in-progress so a
        // self-reference (=A1 in A1) is detected as a cycle.
        res = this.finalValue(this.evalTop(parsed.ast, { si, r, c }), { r, c });
      }
    } catch (e) {
      if (e instanceof TooDeep) throw e;
      if (e instanceof Unsupported) res = KEEP;
      else {
        if (typeof console !== "undefined") console.warn("formula error", cell.f, e);
        res = VALUE;
      }
    } finally {
      this.inProgress.delete(cell);
      this.depth--;
    }
    this.store(cell, res);
    this.done.add(cell);
  }

  /** Evaluate a cell, handling deep dependency chains without blowing the JS stack. */
  compute(si, key, cell) {
    if (this.done.has(cell)) return;
    const stack = [[si, key, cell]];
    const onStack = new Set([cell]);
    while (stack.length) {
      const [s, k, c] = stack[stack.length - 1];
      if (this.done.has(c)) { stack.pop(); onStack.delete(c); continue; }
      try {
        this.depth = 0;
        this.evalCell(s, k, c);
        stack.pop();
        onStack.delete(c);
      } catch (e) {
        if (!(e instanceof TooDeep)) throw e;
        this.inProgress.clear();
        if (onStack.has(e.cell)) {
          // A dependency chain that loops back onto pending work: circular.
          this.cycles.push(`${this.sheets[e.si].name}!${e.key}`);
          this.store(e.cell, CIRC);
          this.done.add(e.cell);
        } else {
          stack.push([e.si, e.key, e.cell]);
          onStack.add(e.cell);
        }
      }
    }
  }

  /** Reduce a formula result to a scalar (refs via implicit intersection, arrays top-left). */
  finalValue(v, pos) {
    if (v instanceof Ref) v = this.implicit(v, pos);
    if (Array.isArray(v)) v = v.length && v[0].length ? v[0][0] : null;
    if (v instanceof Ref) v = this.valueAt(v.sheet, v.r1, v.c1);
    return v;
  }

  store(cell, res) {
    if (res === KEEP) {
      if (cell.v == null || cell.t === "f") { cell.t = "e"; cell.v = "#NAME?"; }
      return;
    }
    const v = res;
    const wasDate = cell.t === "d";
    if (v == null) { cell.t = "n"; cell.v = 0; }
    else if (typeof v === "number") {
      if (!isFinite(v)) { cell.t = "e"; cell.v = "#NUM!"; }
      else if (wasDate && v >= 1 && v < 2958466) { cell.t = "d"; cell.v = serialToIso(v); }
      else { cell.t = "n"; cell.v = round15(v); }
    } else if (typeof v === "string") { cell.t = "s"; cell.v = v; }
    else if (typeof v === "boolean") { cell.t = "b"; cell.v = v; }
    else if (isErr(v)) { cell.t = "e"; cell.v = v.code; }
    else { cell.t = "s"; cell.v = String(v); }
  }

  // ---- bounds / iteration ----
  bounds(si) {
    const sheet = this.sheets[si];
    if (!sheet) return { maxR: -1, maxC: -1, count: 0 };
    return sheetIndex(sheet);
  }

  clip(ref) {
    const b = this.bounds(ref.sheet);
    return { r1: ref.r1, c1: ref.c1, r2: Math.min(ref.r2, b.maxR), c2: Math.min(ref.c2, b.maxC) };
  }

  /** Call fn(value, r, c) for every non-blank cell in the ref. */
  eachInRef(ref, fn) {
    const sheet = this.sheets[ref.sheet];
    if (!sheet) return;
    const cl = this.clip(ref);
    if (cl.r2 < cl.r1 || cl.c2 < cl.c1) return;
    const area = (cl.r2 - cl.r1 + 1) * (cl.c2 - cl.c1 + 1);
    const cells = sheet.cells;
    const idx = this.bounds(ref.sheet);
    if (area > idx.count * 4 + 64) {
      // Sparse sheet relative to range: walk existing cells instead.
      const hits = [];
      for (const key in cells) {
        const { r, c } = rcOfKey(key);
        if (r >= cl.r1 && r <= cl.r2 && c >= cl.c1 && c <= cl.c2) hits.push([r, c, key]);
      }
      hits.sort((x, y) => x[0] - y[0] || x[1] - y[1]);
      for (const [r, c, key] of hits) {
        const v = this.valueOf(ref.sheet, key, cells[key]);
        if (v != null) fn(v, r, c);
      }
      return;
    }
    for (let r = cl.r1; r <= cl.r2; r++) {
      const rs = String(r + 1);
      for (let c = cl.c1; c <= cl.c2; c++) {
        const key = this.col(c) + rs;
        const cell = cells[key];
        if (!cell) continue;
        const v = this.valueOf(ref.sheet, key, cell);
        if (v != null) fn(v, r, c);
      }
    }
  }

  /** 2D array of values (clipped to the used range of the sheet). */
  matrix(ref) {
    const ck = `${ref.sheet}:${ref.r1}:${ref.c1}:${ref.r2}:${ref.c2}`;
    let m = this.matrixCache.get(ck);
    if (m) return m;
    const cl = this.clip(ref);
    const rows = Math.max(0, cl.r2 - cl.r1 + 1), cols = Math.max(0, cl.c2 - cl.c1 + 1);
    if (rows * cols > 5e6) throw new Unsupported("range too large");
    m = [];
    for (let i = 0; i < rows; i++) m.push(new Array(cols).fill(null));
    this.eachInRef(ref, (v, r, c) => { m[r - ref.r1][c - ref.c1] = v; });
    m.fullRows = ref.rows;
    m.fullCols = ref.cols;
    this.matrixCache.set(ck, m);
    return m;
  }

  /** Implicit intersection of a multi-cell ref with the formula's row/column. */
  implicit(ref, pos) {
    if (ref.single) return this.valueAt(ref.sheet, ref.r1, ref.c1);
    if (pos) {
      if (ref.c1 === ref.c2 && pos.r >= ref.r1 && pos.r <= ref.r2) return this.valueAt(ref.sheet, pos.r, ref.c1);
      if (ref.r1 === ref.r2 && pos.c >= ref.c1 && pos.c <= ref.c2) return this.valueAt(ref.sheet, ref.r1, pos.c);
    }
    return this.valueAt(ref.sheet, ref.r1, ref.c1);
  }

  // ---- evaluation ----
  evalTop(ast, ctx) {
    return this.ev(ast, ctx);
  }

  ev(node, ctx) {
    switch (node.t) {
      case "num": return node.v;
      case "str": return node.v;
      case "bool": return node.v;
      case "err": return ERR[node.v] || new XErr(node.v);
      case "missing": return null;
      case "ref": {
        const si = this.sheetIdxOf(node.sheet, ctx.si);
        if (si < 0) return REF;
        return new Ref(si, node.r1, node.c1, node.r2, node.c2);
      }
      case "name": throw new Unsupported(`name ${node.v}`);
      case "arr": return node.rows.map((row) => row.map((n) => this.scalar(this.ev(n, ctx), ctx)));
      case "neg": return this.unary(this.ev(node.a, ctx), ctx, (x) => -x);
      case "pct": return this.unary(this.ev(node.a, ctx), ctx, (x) => x / 100);
      case "bin": return this.binary(node.op, this.ev(node.a, ctx), this.ev(node.b, ctx), ctx);
      case "fn": {
        const def = FUNCS[node.name];
        if (!def) throw new Unsupported(`function ${node.name}`);
        if (def.lazy) return def.fn(node.args, ctx, this);
        const args = node.args.map((a) => this.ev(a, ctx));
        return def.fn(args, ctx, this, node.args);
      }
      default: throw new Unsupported(`node ${node.t}`);
    }
  }

  /** Reduce any value to a scalar (implicit intersection for refs, top-left for arrays). */
  scalar(v, ctx) {
    if (v instanceof Ref) return this.implicit(v, ctx);
    if (Array.isArray(v)) return v.length && v[0].length ? v[0][0] : null;
    return v;
  }

  operand(v) {
    if (v instanceof Ref) return v.single ? this.valueAt(v.sheet, v.r1, v.c1) : this.fullMatrix(v);
    return v;
  }

  /** Like matrix() but padded to the ref's full size (bounded). */
  fullMatrix(ref) {
    const m = this.matrix(ref);
    const rows = Math.min(ref.rows, Math.max(m.length, 1));
    const cols = Math.min(ref.cols, Math.max(m[0] ? m[0].length : 0, 1));
    const out = [];
    for (let i = 0; i < rows; i++) {
      const row = new Array(cols).fill(null);
      if (m[i]) for (let j = 0; j < cols; j++) row[j] = m[i][j] ?? null;
      out.push(row);
    }
    return out;
  }

  unary(v, ctx, f) {
    v = this.operand(v);
    const one = (x) => {
      if (isErr(x)) return x;
      const n = toNum(x);
      return isErr(n) ? n : f(n);
    };
    return Array.isArray(v) ? v.map((row) => row.map(one)) : one(v);
  }

  binary(op, a, b) {
    a = this.operand(a);
    b = this.operand(b);
    if (Array.isArray(a) || Array.isArray(b)) {
      const A = Array.isArray(a) ? a : [[a]];
      const B = Array.isArray(b) ? b : [[b]];
      const rows = Math.max(A.length, B.length), cols = Math.max(A[0]?.length || 0, B[0]?.length || 0);
      const pick = (M, i, j) => {
        const ri = M.length === 1 ? 0 : i, cj = M[0].length === 1 ? 0 : j;
        return M[ri] && cj < M[ri].length ? M[ri][cj] : NA;
      };
      const out = [];
      for (let i = 0; i < rows; i++) {
        const row = [];
        for (let j = 0; j < cols; j++) row.push(binScalar(op, pick(A, i, j), pick(B, i, j)));
        out.push(row);
      }
      return out;
    }
    return binScalar(op, a, b);
  }
}

function binScalar(op, a, b) {
  if (isErr(a)) return a;
  if (isErr(b)) return b;
  switch (op) {
    case "+": case "-": case "*": case "/": case "^": {
      const x = toNum(a), y = toNum(b);
      if (isErr(x)) return x;
      if (isErr(y)) return y;
      let r;
      if (op === "+") r = x + y;
      else if (op === "-") r = x - y;
      else if (op === "*") r = x * y;
      else if (op === "/") { if (y === 0) return DIV0; r = x / y; }
      else {
        if (x === 0 && y < 0) return DIV0;
        r = Math.pow(x, y);
      }
      return isFinite(r) ? r : NUM;
    }
    case "&": return toStr(a) + toStr(b);
    default: {
      const c = compare(a, b);
      switch (op) {
        case "=": return c === 0;
        case "<>": return c !== 0;
        case "<": return c < 0;
        case ">": return c > 0;
        case "<=": return c <= 0;
        case ">=": return c >= 0;
      }
    }
  }
  return VALUE;
}

// ---------------------------------------------------------------------------
// Function helpers
// ---------------------------------------------------------------------------

/**
 * Visit numbers for aggregate functions. Values from ranges/arrays: numbers
 * only (text/booleans ignored). Direct arguments: coerced. Returns an error or null.
 */
function eachNumber(args, eng, fn, { directText = true } = {}) {
  for (const a of args) {
    if (a instanceof Ref) {
      let err = null;
      eng.eachInRef(a, (v) => {
        if (err) return;
        if (isErr(v)) err = v;
        else if (typeof v === "number") fn(v);
      });
      if (err) return err;
    } else if (Array.isArray(a)) {
      for (const row of a) for (const v of row) {
        if (isErr(v)) return v;
        if (typeof v === "number") fn(v);
      }
    } else if (a == null) {
      continue;
    } else if (isErr(a)) {
      return a;
    } else {
      if (typeof a === "string" && !directText) continue;
      const n = toNum(a);
      if (isErr(n)) return n;
      fn(n);
    }
  }
  return null;
}

function numbers(args, eng, opts) {
  const out = [];
  const err = eachNumber(args, eng, (n) => out.push(n), opts);
  return err || out;
}

/** Every value (including text/booleans; blanks skipped) of refs/arrays/scalars. */
function eachValue(args, eng, fn, includeBlankScalars = false) {
  for (const a of args) {
    if (a instanceof Ref) eng.eachInRef(a, fn);
    else if (Array.isArray(a)) { for (const row of a) for (const v of row) if (v != null) fn(v); }
    else if (a != null || includeBlankScalars) fn(a);
  }
}

function num(v, eng, ctx) {
  return toNum(eng.scalar(v, ctx));
}
function str(v, eng, ctx) {
  return toStr(eng.scalar(v, ctx));
}
function bool(v, eng, ctx) {
  return toBool(eng.scalar(v, ctx));
}

/** Matrix (2D values) for a range/array/scalar argument. */
function asMatrix(v, eng) {
  if (v instanceof Ref) return eng.matrix(v);
  if (Array.isArray(v)) return v;
  return [[v]];
}

/** First error found among scalars, if any. */
function firstErr(...vals) {
  for (const v of vals) if (isErr(v)) return v;
  return null;
}

function intArg(v, eng, ctx) {
  const n = num(v, eng, ctx);
  return isErr(n) ? n : Math.trunc(n);
}

const FUNCS = {};
function def(name, fn, lazy = false) {
  for (const n of name.split(" ")) FUNCS[n] = { fn, lazy };
}

// ---------------------------------------------------------------------------
// Math & aggregates
// ---------------------------------------------------------------------------

def("SUM", (args, ctx, eng) => {
  let s = 0;
  const err = eachNumber(args, eng, (n) => { s += n; });
  return err || s;
});
def("PRODUCT", (args, ctx, eng) => {
  let p = 1, any = false;
  const err = eachNumber(args, eng, (n) => { p *= n; any = true; });
  return err || (any ? p : 0);
});
def("AVERAGE", (args, ctx, eng) => {
  const ns = numbers(args, eng);
  if (isErr(ns)) return ns;
  return ns.length ? ns.reduce((a, b) => a + b, 0) / ns.length : DIV0;
});
def("AVERAGEA", (args, ctx, eng) => {
  let s = 0, n = 0, err = null;
  eachValue(args, eng, (v) => {
    if (isErr(v)) { err = err || v; return; }
    s += typeof v === "number" ? v : v === true ? 1 : 0;
    n++;
  });
  return err || (n ? s / n : DIV0);
});
def("MIN", (args, ctx, eng) => {
  const ns = numbers(args, eng);
  if (isErr(ns)) return ns;
  return ns.length ? ns.reduce((a, b) => (b < a ? b : a), Infinity) : 0;
});
def("MAX", (args, ctx, eng) => {
  const ns = numbers(args, eng);
  if (isErr(ns)) return ns;
  return ns.length ? ns.reduce((a, b) => (b > a ? b : a), -Infinity) : 0;
});
def("COUNT", (args, ctx, eng) => {
  let n = 0;
  for (const a of args) {
    if (a instanceof Ref || Array.isArray(a)) eachValue([a], eng, (v) => { if (typeof v === "number") n++; });
    else if (typeof a === "number" || typeof a === "boolean" || (typeof a === "string" && typeof toNum(a) === "number")) n++;
  }
  return n;
});
def("COUNTA", (args, ctx, eng, nodes) => {
  let n = 0;
  args.forEach((a, i) => {
    if (a instanceof Ref || Array.isArray(a)) eachValue([a], eng, () => { n++; });
    else if (nodes[i].t !== "missing") n++;
  });
  return n;
});
def("COUNTBLANK", (args, ctx, eng) => {
  const ref = args[0];
  if (!(ref instanceof Ref)) return ref == null || ref === "" ? 1 : 0;
  let filled = 0;
  eng.eachInRef(ref, (v) => { if (v !== "") filled++; });
  return ref.rows * ref.cols - filled;
});
def("MEDIAN", (args, ctx, eng) => {
  const ns = numbers(args, eng);
  if (isErr(ns)) return ns;
  if (!ns.length) return NUM;
  ns.sort((a, b) => a - b);
  const mid = ns.length >> 1;
  return ns.length % 2 ? ns[mid] : (ns[mid - 1] + ns[mid]) / 2;
});
function variance(ns, sample) {
  const n = ns.length;
  if (n < (sample ? 2 : 1)) return DIV0;
  const mean = ns.reduce((a, b) => a + b, 0) / n;
  const ss = ns.reduce((a, b) => a + (b - mean) * (b - mean), 0);
  return ss / (sample ? n - 1 : n);
}
def("STDEV STDEV.S", (args, ctx, eng) => {
  const ns = numbers(args, eng); if (isErr(ns)) return ns;
  const v = variance(ns, true); return isErr(v) ? v : Math.sqrt(v);
});
def("STDEVP STDEV.P", (args, ctx, eng) => {
  const ns = numbers(args, eng); if (isErr(ns)) return ns;
  const v = variance(ns, false); return isErr(v) ? v : Math.sqrt(v);
});
def("VAR VAR.S", (args, ctx, eng) => {
  const ns = numbers(args, eng); return isErr(ns) ? ns : variance(ns, true);
});
def("VARP VAR.P", (args, ctx, eng) => {
  const ns = numbers(args, eng); return isErr(ns) ? ns : variance(ns, false);
});
def("LARGE", (args, ctx, eng) => kth(args, ctx, eng, true));
def("SMALL", (args, ctx, eng) => kth(args, ctx, eng, false));
function kth(args, ctx, eng, large) {
  const ns = numbers([args[0]], eng);
  if (isErr(ns)) return ns;
  const k = intArg(args[1], eng, ctx);
  if (isErr(k)) return k;
  if (k < 1 || k > ns.length) return NUM;
  ns.sort((a, b) => (large ? b - a : a - b));
  return ns[k - 1];
}
def("RANK RANK.EQ", (args, ctx, eng) => {
  const x = num(args[0], eng, ctx);
  if (isErr(x)) return x;
  const ns = numbers([args[1]], eng);
  if (isErr(ns)) return ns;
  const asc = args[2] != null && toBool(eng.scalar(args[2], ctx)) === true;
  if (!ns.some((n) => round15(n) === round15(x))) return NA;
  return 1 + ns.filter((n) => (asc ? n < x : n > x)).length;
});
def("ROUND", (args, ctx, eng) => {
  const x = num(args[0], eng, ctx), d = intArg(args[1] ?? 0, eng, ctx);
  return firstErr(x, d) || roundDecimal(x, d);
});
function roundDir(x, d, up) {
  const f = Math.pow(10, d);
  let s = Math.abs(x) * f;
  if (Math.abs(s - Math.round(s)) < 1e-9 * Math.max(1, Math.abs(s))) s = Math.round(s);
  s = up ? Math.ceil(s) : Math.floor(s);
  return Math.sign(x) * (s / f);
}
def("ROUNDUP", (args, ctx, eng) => {
  const x = num(args[0], eng, ctx), d = intArg(args[1] ?? 0, eng, ctx);
  return firstErr(x, d) || roundDir(x, d, true);
});
def("ROUNDDOWN TRUNC", (args, ctx, eng) => {
  const x = num(args[0], eng, ctx), d = intArg(args[1] ?? 0, eng, ctx);
  return firstErr(x, d) || roundDir(x, d, false);
});
def("INT", (args, ctx, eng) => { const x = num(args[0], eng, ctx); return isErr(x) ? x : Math.floor(x); });
def("MOD", (args, ctx, eng) => {
  const n = num(args[0], eng, ctx), d = num(args[1], eng, ctx);
  const e = firstErr(n, d); if (e) return e;
  if (d === 0) return DIV0;
  const r = n - d * Math.floor(n / d);
  return round15(r);
});
def("ABS", (args, ctx, eng) => { const x = num(args[0], eng, ctx); return isErr(x) ? x : Math.abs(x); });
def("SIGN", (args, ctx, eng) => { const x = num(args[0], eng, ctx); return isErr(x) ? x : Math.sign(x); });
def("SQRT", (args, ctx, eng) => {
  const x = num(args[0], eng, ctx); if (isErr(x)) return x;
  return x < 0 ? NUM : Math.sqrt(x);
});
def("POWER", (args, ctx, eng) => binScalar("^", num(args[0], eng, ctx), num(args[1], eng, ctx)));
def("EXP", (args, ctx, eng) => { const x = num(args[0], eng, ctx); return isErr(x) ? x : Math.exp(x); });
def("LN", (args, ctx, eng) => { const x = num(args[0], eng, ctx); return isErr(x) ? x : x <= 0 ? NUM : Math.log(x); });
def("LOG10", (args, ctx, eng) => { const x = num(args[0], eng, ctx); return isErr(x) ? x : x <= 0 ? NUM : Math.log10(x); });
def("LOG", (args, ctx, eng) => {
  const x = num(args[0], eng, ctx), b = args[1] == null ? 10 : num(args[1], eng, ctx);
  const e = firstErr(x, b); if (e) return e;
  if (x <= 0 || b <= 0 || b === 1) return b === 1 ? DIV0 : NUM;
  return Math.log(x) / Math.log(b);
});
def("PI", () => Math.PI);
def("RAND", () => Math.random());
def("RANDBETWEEN", (args, ctx, eng) => {
  const lo = num(args[0], eng, ctx), hi = num(args[1], eng, ctx);
  const e = firstErr(lo, hi); if (e) return e;
  const a = Math.ceil(lo), b = Math.floor(hi);
  return a > b ? NUM : a + Math.floor(Math.random() * (b - a + 1));
});
def("CEILING CEILING.MATH", (args, ctx, eng) => {
  const x = num(args[0], eng, ctx), s = args[1] == null ? 1 : num(args[1], eng, ctx);
  const e = firstErr(x, s); if (e) return e;
  if (s === 0) return 0;
  return round15(Math.ceil(x / Math.abs(s)) * Math.abs(s));
});
def("FLOOR FLOOR.MATH", (args, ctx, eng) => {
  const x = num(args[0], eng, ctx), s = args[1] == null ? 1 : num(args[1], eng, ctx);
  const e = firstErr(x, s); if (e) return e;
  if (s === 0) return x === 0 ? 0 : DIV0;
  return round15(Math.floor(x / Math.abs(s)) * Math.abs(s));
});
def("MROUND", (args, ctx, eng) => {
  const x = num(args[0], eng, ctx), m = num(args[1], eng, ctx);
  const e = firstErr(x, m); if (e) return e;
  if (m === 0) return 0;
  if (Math.sign(x) * Math.sign(m) < 0) return NUM;
  return round15(roundDecimal(x / m, 0) * m);
});
def("QUOTIENT", (args, ctx, eng) => {
  const x = num(args[0], eng, ctx), y = num(args[1], eng, ctx);
  const e = firstErr(x, y); if (e) return e;
  return y === 0 ? DIV0 : Math.trunc(x / y);
});
def("SUMPRODUCT", (args, ctx, eng) => {
  const ms = args.map((a) => (a instanceof Ref ? eng.fullMatrix(a) : Array.isArray(a) ? a : [[a]]));
  const rows = ms[0].length, cols = ms[0][0] ? ms[0][0].length : 0;
  if (ms.some((m) => m.length !== rows || (m[0] ? m[0].length : 0) !== cols)) {
    // allow clipped matrices from whole-column refs: compare full sizes
    const sizes = args.map((a) => (a instanceof Ref ? `${a.rows}x${a.cols}` : null)).filter(Boolean);
    if (new Set(sizes).size > 1) return VALUE;
  }
  let s = 0;
  const R = Math.min(...ms.map((m) => m.length));
  for (let i = 0; i < R; i++) {
    const C = Math.min(...ms.map((m) => m[i].length));
    for (let j = 0; j < C; j++) {
      let p = 1;
      for (const m of ms) {
        const v = m[i][j];
        if (isErr(v)) return v;
        p *= typeof v === "number" ? v : v === true && ms.length === 1 ? 1 : 0;
      }
      s += p;
    }
  }
  return s;
});
def("SUMSQ", (args, ctx, eng) => {
  let s = 0;
  const err = eachNumber(args, eng, (n) => { s += n * n; });
  return err || s;
});

// ---- conditional aggregates ----
function condPairs(args, eng, ctx, start) {
  const pairs = [];
  for (let i = start; i + 1 < args.length; i += 2) {
    const rng = args[i];
    if (!(rng instanceof Ref) && !Array.isArray(rng)) return VALUE;
    pairs.push({ m: asMatrix(rng, eng), rows: rng instanceof Ref ? rng.rows : rng.length,
      cols: rng instanceof Ref ? rng.cols : rng[0].length, test: makeCriteria(eng.scalar(args[i + 1], ctx)) });
  }
  return pairs;
}
/** Visit (i,j) positions where all criteria pass; blanks beyond the used range are considered too. */
function eachMatch(pairs, fn) {
  const { rows, cols } = pairs[0];
  if (pairs.some((p) => p.rows !== rows || p.cols !== cols)) return VALUE;
  const blankOk = pairs.every((p) => p.test(null));
  const maxI = blankOk ? rows : Math.max(...pairs.map((p) => p.m.length));
  for (let i = 0; i < Math.min(rows, maxI); i++) {
    for (let j = 0; j < cols; j++) {
      let ok = true;
      for (const p of pairs) {
        const row = p.m[i];
        const v = row ? (row[j] ?? null) : null;
        if (!p.test(v)) { ok = false; break; }
      }
      if (ok) fn(i, j);
    }
  }
  return null;
}
function valueAtOffset(target, i, j, eng) {
  if (target instanceof Ref) {
    if (target.r1 + i > MAX_ROWS - 1 || target.c1 + j > MAX_COLS - 1) return null;
    return eng.valueAt(target.sheet, target.r1 + i, target.c1 + j);
  }
  if (Array.isArray(target)) return target[i] ? target[i][j] ?? null : null;
  return null;
}
function ifsAggregate(target, pairs, eng, kind) {
  if (isErr(pairs)) return pairs;
  let s = 0, n = 0, mn = Infinity, mx = -Infinity, err = null;
  const e = eachMatch(pairs, (i, j) => {
    if (kind === "count") { n++; return; }
    const v = valueAtOffset(target, i, j, eng);
    if (isErr(v)) { err = err || v; return; }
    if (typeof v !== "number") return;
    s += v; n++;
    if (v < mn) mn = v;
    if (v > mx) mx = v;
  });
  if (e) return e;
  if (err) return err;
  switch (kind) {
    case "count": return n;
    case "sum": return s;
    case "avg": return n ? s / n : DIV0;
    case "min": return n ? mn : 0;
    case "max": return n ? mx : 0;
  }
  return VALUE;
}
def("SUMIF", (args, ctx, eng) => ifsAggregate(args[2] ?? args[0], condPairs(args.slice(0, 2), eng, ctx, 0), eng, "sum"));
def("AVERAGEIF", (args, ctx, eng) => ifsAggregate(args[2] ?? args[0], condPairs(args.slice(0, 2), eng, ctx, 0), eng, "avg"));
def("COUNTIF", (args, ctx, eng) => ifsAggregate(null, condPairs(args, eng, ctx, 0), eng, "count"));
def("COUNTIFS", (args, ctx, eng) => ifsAggregate(null, condPairs(args, eng, ctx, 0), eng, "count"));
def("SUMIFS", (args, ctx, eng) => ifsAggregate(args[0], condPairs(args, eng, ctx, 1), eng, "sum"));
def("AVERAGEIFS", (args, ctx, eng) => ifsAggregate(args[0], condPairs(args, eng, ctx, 1), eng, "avg"));
def("MINIFS", (args, ctx, eng) => ifsAggregate(args[0], condPairs(args, eng, ctx, 1), eng, "min"));
def("MAXIFS", (args, ctx, eng) => ifsAggregate(args[0], condPairs(args, eng, ctx, 1), eng, "max"));

// ---------------------------------------------------------------------------
// Logic & info
// ---------------------------------------------------------------------------

def("IF", (nodes, ctx, eng) => {
  const c = bool(eng.ev(nodes[0], ctx), eng, ctx);
  if (isErr(c)) return c;
  if (c) return nodes.length > 1 ? (nodes[1].t === "missing" ? 0 : eng.ev(nodes[1], ctx)) : true;
  return nodes.length > 2 ? (nodes[2].t === "missing" ? 0 : eng.ev(nodes[2], ctx)) : false;
}, true);
def("IFS", (nodes, ctx, eng) => {
  for (let i = 0; i + 1 < nodes.length; i += 2) {
    const c = bool(eng.ev(nodes[i], ctx), eng, ctx);
    if (isErr(c)) return c;
    if (c) return eng.ev(nodes[i + 1], ctx);
  }
  return NA;
}, true);
def("IFERROR", (nodes, ctx, eng) => {
  const v = eng.ev(nodes[0], ctx);
  const s = v instanceof Ref || Array.isArray(v) ? eng.scalar(v, ctx) : v;
  if (isErr(s)) return nodes[1] ? eng.ev(nodes[1], ctx) : 0;
  return v;
}, true);
def("IFNA", (nodes, ctx, eng) => {
  const v = eng.ev(nodes[0], ctx);
  const s = v instanceof Ref || Array.isArray(v) ? eng.scalar(v, ctx) : v;
  if (s === NA || (isErr(s) && s.code === "#N/A")) return nodes[1] ? eng.ev(nodes[1], ctx) : 0;
  return v;
}, true);
function logicVals(args, eng) {
  const out = [];
  for (const a of args) {
    if (a instanceof Ref || Array.isArray(a)) {
      let err = null;
      eachValue([a], eng, (v) => {
        if (isErr(v)) err = err || v;
        else if (typeof v === "boolean") out.push(v);
        else if (typeof v === "number") out.push(v !== 0);
      });
      if (err) return err;
    } else if (a != null) {
      const b = toBool(a);
      if (isErr(b)) return b;
      out.push(b);
    }
  }
  return out.length ? out : VALUE;
}
def("AND", (args, ctx, eng) => { const v = logicVals(args, eng); return isErr(v) ? v : v.every(Boolean); });
def("OR", (args, ctx, eng) => { const v = logicVals(args, eng); return isErr(v) ? v : v.some(Boolean); });
def("XOR", (args, ctx, eng) => { const v = logicVals(args, eng); return isErr(v) ? v : v.filter(Boolean).length % 2 === 1; });
def("NOT", (args, ctx, eng) => { const b = bool(args[0], eng, ctx); return isErr(b) ? b : !b; });
def("TRUE", () => true);
def("FALSE", () => false);
def("SWITCH", (nodes, ctx, eng) => {
  const x = eng.scalar(eng.ev(nodes[0], ctx), ctx);
  if (isErr(x)) return x;
  let i = 1;
  for (; i + 1 < nodes.length; i += 2) {
    const v = eng.scalar(eng.ev(nodes[i], ctx), ctx);
    if (compare(x, v) === 0 && typeRank(x ?? 0) === typeRank(v ?? 0)) return eng.ev(nodes[i + 1], ctx);
  }
  return i < nodes.length ? eng.ev(nodes[i], ctx) : NA;
}, true);
def("CHOOSE", (nodes, ctx, eng) => {
  const i = intArg(eng.ev(nodes[0], ctx), eng, ctx);
  if (isErr(i)) return i;
  if (i < 1 || i >= nodes.length) return VALUE;
  return eng.ev(nodes[i], ctx);
}, true);
def("ISBLANK", (args, ctx, eng) => eng.scalar(args[0], ctx) == null);
def("ISNUMBER", (args, ctx, eng) => typeof eng.scalar(args[0], ctx) === "number");
def("ISTEXT", (args, ctx, eng) => typeof eng.scalar(args[0], ctx) === "string");
def("ISNONTEXT", (args, ctx, eng) => typeof eng.scalar(args[0], ctx) !== "string");
def("ISLOGICAL", (args, ctx, eng) => typeof eng.scalar(args[0], ctx) === "boolean");
def("ISERROR", (args, ctx, eng) => isErr(eng.scalar(args[0], ctx)));
def("ISERR", (args, ctx, eng) => { const v = eng.scalar(args[0], ctx); return isErr(v) && v.code !== "#N/A"; });
def("ISNA", (args, ctx, eng) => { const v = eng.scalar(args[0], ctx); return isErr(v) && v.code === "#N/A"; });
def("ISEVEN", (args, ctx, eng) => { const x = num(args[0], eng, ctx); return isErr(x) ? x : Math.trunc(x) % 2 === 0; });
def("ISODD", (args, ctx, eng) => { const x = num(args[0], eng, ctx); return isErr(x) ? x : Math.abs(Math.trunc(x)) % 2 === 1; });
def("NA", () => NA);
def("N", (args, ctx, eng) => { const v = eng.scalar(args[0], ctx); return typeof v === "number" ? v : v === true ? 1 : isErr(v) ? v : 0; });
def("T", (args, ctx, eng) => { const v = eng.scalar(args[0], ctx); return typeof v === "string" ? v : isErr(v) ? v : ""; });
def("ROW", (args, ctx) => (args[0] instanceof Ref ? args[0].r1 + 1 : args.length ? VALUE : ctx.r + 1));
def("COLUMN", (args, ctx) => (args[0] instanceof Ref ? args[0].c1 + 1 : args.length ? VALUE : ctx.c + 1));
def("ROWS", (args) => (args[0] instanceof Ref ? args[0].rows : Array.isArray(args[0]) ? args[0].length : 1));
def("COLUMNS", (args) => (args[0] instanceof Ref ? args[0].cols : Array.isArray(args[0]) ? args[0][0].length : 1));

// ---------------------------------------------------------------------------
// Lookup
// ---------------------------------------------------------------------------

/** Find index in a vector. mode: 0 exact (wildcards for text), 1 largest <= (sorted asc), -1 smallest >= (sorted desc). */
function findIndex(vec, x, mode, eng, cacheKey) {
  if (mode === 0) {
    if (typeof x === "string" && hasWildcard(x)) {
      const re = wildcardRegex(x);
      return vec.findIndex((v) => typeof v === "string" && re.test(v));
    }
    if (cacheKey) {
      let map = eng.lookupCache.get(cacheKey);
      if (!map) {
        map = new Map();
        vec.forEach((v, i) => {
          if (v == null) return;
          const k = typeof v === "string" ? "s" + v.toLowerCase() : typeof v === "number" ? "n" + round15(v) : typeof v === "boolean" ? "b" + v : null;
          if (k != null && !map.has(k)) map.set(k, i);
        });
        eng.lookupCache.set(cacheKey, map);
      }
      const k = typeof x === "string" ? "s" + x.toLowerCase() : typeof x === "number" ? "n" + round15(x) : typeof x === "boolean" ? "b" + x : null;
      const i = k == null ? undefined : map.get(k);
      return i == null ? -1 : i;
    }
    return vec.findIndex((v) => v != null && typeRank(v) === typeRank(x) && compare(v, x) === 0);
  }
  let best = -1;
  for (let i = 0; i < vec.length; i++) {
    const v = vec[i];
    if (v == null || typeRank(v) !== typeRank(x)) continue;
    const c = compare(v, x);
    if (mode === 1) { if (c <= 0) best = i; else break; }
    else { if (c >= 0) best = i; else break; }
  }
  return best;
}

function refKey(r) {
  return r instanceof Ref ? `${r.sheet}:${r.r1}:${r.c1}:${r.r2}:${r.c2}` : null;
}

function vhlookup(args, ctx, eng, vertical) {
  const x = eng.scalar(args[0], ctx);
  if (isErr(x)) return x;
  const table = args[1];
  const k = intArg(args[2], eng, ctx);
  if (isErr(k)) return k;
  const approx = args.length < 4 || args[3] == null ? true : toBool(eng.scalar(args[3], ctx));
  if (isErr(approx)) return approx;
  const m = asMatrix(table, eng);
  const width = table instanceof Ref ? (vertical ? table.cols : table.rows) : (vertical ? m[0]?.length : m.length) || 0;
  if (k < 1) return VALUE;
  if (k > width) return REF;
  const vec = vertical ? m.map((row) => row[0]) : (m[0] || []).slice();
  const ck = table instanceof Ref ? `${vertical ? "v" : "h"}${refKey(table)}` : null;
  const i = findIndex(vec, x, approx ? 1 : 0, eng, approx ? null : ck);
  if (i < 0) return NA;
  const v = vertical ? (m[i] ? m[i][k - 1] : null) : (m[k - 1] ? m[k - 1][i] : null);
  return v == null ? 0 : v;
}
def("VLOOKUP", (a, c, e) => vhlookup(a, c, e, true));
def("HLOOKUP", (a, c, e) => vhlookup(a, c, e, false));

function vectorOf(v, eng) {
  const m = asMatrix(v, eng);
  if (v instanceof Ref) {
    if (v.cols === 1) return { vec: m.map((r) => r[0]), vertical: true };
    if (v.rows === 1) return { vec: (m[0] || []).slice(), vertical: false };
    return null;
  }
  if (m.length === 1) return { vec: m[0].slice(), vertical: false };
  if (m[0] && m[0].length === 1) return { vec: m.map((r) => r[0]), vertical: true };
  return null;
}

def("MATCH", (args, ctx, eng) => {
  const x = eng.scalar(args[0], ctx);
  if (isErr(x)) return x;
  const vv = vectorOf(args[1], eng);
  if (!vv) return NA;
  const t = args.length > 2 && args[2] != null ? intArg(args[2], eng, ctx) : 1;
  if (isErr(t)) return t;
  const mode = t === 0 ? 0 : t > 0 ? 1 : -1;
  const i = findIndex(vv.vec, x, mode, eng, mode === 0 ? `m${refKey(args[1])}` : null);
  return i < 0 ? NA : i + 1;
});
def("INDEX", (args, ctx, eng) => {
  const src = args[0];
  let r = args.length > 1 && args[1] != null ? intArg(args[1], eng, ctx) : 0;
  let c = args.length > 2 && args[2] != null ? intArg(args[2], eng, ctx) : 0;
  const e = firstErr(r, c); if (e) return e;
  if (r < 0 || c < 0) return VALUE;
  if (src instanceof Ref) {
    if (args.length === 2 && src.rows === 1 && src.cols > 1) { c = r; r = 1; }
    if (r > src.rows || c > src.cols) return REF;
    if (r === 0 && c === 0) return src;
    if (r === 0) return new Ref(src.sheet, src.r1, src.c1 + c - 1, src.r2, src.c1 + c - 1);
    if (c === 0) {
      if (src.cols === 1) return new Ref(src.sheet, src.r1 + r - 1, src.c1, src.r1 + r - 1, src.c1);
      return new Ref(src.sheet, src.r1 + r - 1, src.c1, src.r1 + r - 1, src.c2);
    }
    return new Ref(src.sheet, src.r1 + r - 1, src.c1 + c - 1, src.r1 + r - 1, src.c1 + c - 1);
  }
  const m = asMatrix(src, eng);
  if (args.length === 2 && m.length === 1) { c = r; r = 1; }
  if (r === 0) r = 1;
  if (c === 0) c = 1;
  if (!m[r - 1] || c > m[r - 1].length) return REF;
  return m[r - 1][c - 1];
});
def("XLOOKUP", (args, ctx, eng) => {
  const x = eng.scalar(args[0], ctx);
  if (isErr(x)) return x;
  const vv = vectorOf(args[1], eng);
  if (!vv) return VALUE;
  const ret = args[2];
  const matchMode = args.length > 4 && args[4] != null ? intArg(args[4], eng, ctx) : 0;
  const searchMode = args.length > 5 && args[5] != null ? intArg(args[5], eng, ctx) : 1;
  const e = firstErr(matchMode, searchMode); if (e) return e;
  const vec = vv.vec;
  const order = [...vec.keys()];
  if (searchMode < 0) order.reverse();
  let idx = -1;
  const re = matchMode === 2 && typeof x === "string" ? wildcardRegex(x) : null;
  let best = -1;
  for (const i of order) {
    const v = vec[i];
    if (v == null) continue;
    if (re) { if (typeof v === "string" && re.test(v)) { idx = i; break; } continue; }
    if (typeRank(v) !== typeRank(x)) continue;
    const c = compare(v, x);
    if (c === 0) { idx = i; break; }
    if (matchMode === -1 && c < 0 && (best < 0 || compare(v, vec[best]) > 0)) best = i;
    if (matchMode === 1 && c > 0 && (best < 0 || compare(v, vec[best]) < 0)) best = i;
  }
  if (idx < 0) idx = best;
  if (idx < 0) return args.length > 3 && args[3] != null ? args[3] : NA;
  if (ret instanceof Ref) {
    if (vv.vertical) {
      if (ret.r1 + idx > ret.r2) return VALUE;
      return new Ref(ret.sheet, ret.r1 + idx, ret.c1, ret.r1 + idx, ret.c2);
    }
    if (ret.c1 + idx > ret.c2) return VALUE;
    return new Ref(ret.sheet, ret.r1, ret.c1 + idx, ret.r2, ret.c1 + idx);
  }
  const m = asMatrix(ret, eng);
  return vv.vertical ? [m[idx] || [NA]] : m.map((row) => [row[idx]]);
});

// ---------------------------------------------------------------------------
// Text
// ---------------------------------------------------------------------------

def("CONCAT", (args, ctx, eng) => {
  let s = "", err = null;
  for (const a of args) {
    if (a instanceof Ref || Array.isArray(a)) eachValue([a], eng, (v) => { if (isErr(v)) err = err || v; else s += toStr(v); });
    else { const t = toStr(a); if (isErr(t)) return t; s += t; }
  }
  return err || s;
});
def("CONCATENATE", (args, ctx, eng) => {
  let s = "";
  for (const a of args) {
    const t = str(a, eng, ctx);
    if (isErr(t)) return t;
    s += t;
  }
  return s;
});
def("TEXTJOIN", (args, ctx, eng) => {
  const d = str(args[0], eng, ctx);
  const ignore = bool(args[1], eng, ctx);
  const e = firstErr(d, ignore); if (e) return e;
  const parts = [];
  let err = null;
  for (const a of args.slice(2)) {
    if (a instanceof Ref) {
      // include blanks inside the range when not ignoring empties
      const m = eng.fullMatrix(a);
      for (const row of m) for (const v of row) {
        if (isErr(v)) { err = err || v; continue; }
        const t = toStr(v);
        if (!(ignore && t === "")) parts.push(t);
      }
    } else if (Array.isArray(a)) {
      for (const row of a) for (const v of row) {
        if (isErr(v)) { err = err || v; continue; }
        const t = toStr(v);
        if (!(ignore && t === "")) parts.push(t);
      }
    } else {
      const t = toStr(a);
      if (isErr(t)) return t;
      if (!(ignore && t === "")) parts.push(t);
    }
  }
  return err || parts.join(d);
});
def("LEFT", (args, ctx, eng) => {
  const s = str(args[0], eng, ctx), n = args.length > 1 ? intArg(args[1], eng, ctx) : 1;
  const e = firstErr(s, n); if (e) return e;
  return n < 0 ? VALUE : [...s].slice(0, n).join("");
});
def("RIGHT", (args, ctx, eng) => {
  const s = str(args[0], eng, ctx), n = args.length > 1 ? intArg(args[1], eng, ctx) : 1;
  const e = firstErr(s, n); if (e) return e;
  if (n < 0) return VALUE;
  const chars = [...s];
  return n === 0 ? "" : chars.slice(Math.max(0, chars.length - n)).join("");
});
def("MID", (args, ctx, eng) => {
  const s = str(args[0], eng, ctx), st = intArg(args[1], eng, ctx), n = intArg(args[2], eng, ctx);
  const e = firstErr(s, st, n); if (e) return e;
  if (st < 1 || n < 0) return VALUE;
  return [...s].slice(st - 1, st - 1 + n).join("");
});
def("LEN", (args, ctx, eng) => { const s = str(args[0], eng, ctx); return isErr(s) ? s : [...s].length; });
def("UPPER", (args, ctx, eng) => { const s = str(args[0], eng, ctx); return isErr(s) ? s : s.toUpperCase(); });
def("LOWER", (args, ctx, eng) => { const s = str(args[0], eng, ctx); return isErr(s) ? s : s.toLowerCase(); });
def("PROPER", (args, ctx, eng) => {
  const s = str(args[0], eng, ctx);
  return isErr(s) ? s : s.toLowerCase().replace(/(^|[^A-Za-zÀ-￿])([a-zà-￿])/g, (m, a, b) => a + b.toUpperCase());
});
def("TRIM", (args, ctx, eng) => { const s = str(args[0], eng, ctx); return isErr(s) ? s : s.replace(/ +/g, " ").trim(); });
def("CLEAN", (args, ctx, eng) => { const s = str(args[0], eng, ctx); return isErr(s) ? s : s.replace(/[\x00-\x1f]/g, ""); });
def("SUBSTITUTE", (args, ctx, eng) => {
  const s = str(args[0], eng, ctx), o = str(args[1], eng, ctx), n = str(args[2], eng, ctx);
  const inst = args.length > 3 && args[3] != null ? intArg(args[3], eng, ctx) : null;
  const e = firstErr(s, o, n, inst); if (e) return e;
  if (o === "") return s;
  if (inst == null) return s.split(o).join(n);
  if (inst < 1) return VALUE;
  let idx = -1;
  for (let k = 0; k < inst; k++) {
    idx = s.indexOf(o, idx + 1);
    if (idx < 0) return s;
  }
  return s.slice(0, idx) + n + s.slice(idx + o.length);
});
def("REPLACE", (args, ctx, eng) => {
  const s = str(args[0], eng, ctx), st = intArg(args[1], eng, ctx), n = intArg(args[2], eng, ctx), t = str(args[3], eng, ctx);
  const e = firstErr(s, st, n, t); if (e) return e;
  if (st < 1 || n < 0) return VALUE;
  return s.slice(0, st - 1) + t + s.slice(st - 1 + n);
});
def("FIND", (args, ctx, eng) => {
  const f = str(args[0], eng, ctx), s = str(args[1], eng, ctx), st = args.length > 2 ? intArg(args[2], eng, ctx) : 1;
  const e = firstErr(f, s, st); if (e) return e;
  if (st < 1 || st > s.length + 1) return VALUE;
  const i = s.indexOf(f, st - 1);
  return i < 0 ? VALUE : i + 1;
});
def("SEARCH", (args, ctx, eng) => {
  const f = str(args[0], eng, ctx), s = str(args[1], eng, ctx), st = args.length > 2 ? intArg(args[2], eng, ctx) : 1;
  const e = firstErr(f, s, st); if (e) return e;
  if (st < 1 || st > s.length + 1) return VALUE;
  if (hasWildcard(f)) {
    const re = new RegExp(wildcardRegex(f).source.slice(1, -1), "i");
    const m = re.exec(s.slice(st - 1));
    return m ? m.index + st : VALUE;
  }
  const i = s.toLowerCase().indexOf(f.toLowerCase(), st - 1);
  return i < 0 ? VALUE : i + 1;
});
def("TEXT", (args, ctx, eng) => {
  const v = eng.scalar(args[0], ctx);
  const f = str(args[1], eng, ctx);
  const e = firstErr(v, f); if (e) return e;
  if (typeof v === "string") {
    const n = toNum(v);
    return typeof n === "number" ? formatValue(n, f) : formatValue(v, f);
  }
  return formatValue(typeof v === "boolean" ? v : v == null ? 0 : v, f);
});
def("VALUE", (args, ctx, eng) => {
  const v = eng.scalar(args[0], ctx);
  if (typeof v === "number") return v;
  if (v == null) return 0;
  if (typeof v === "boolean") return VALUE;
  return toNum(v);
});
def("REPT", (args, ctx, eng) => {
  const s = str(args[0], eng, ctx), n = intArg(args[1], eng, ctx);
  const e = firstErr(s, n); if (e) return e;
  return n < 0 || s.length * n > 32767 ? VALUE : s.repeat(n);
});
def("EXACT", (args, ctx, eng) => {
  const a = str(args[0], eng, ctx), b = str(args[1], eng, ctx);
  return firstErr(a, b) || a === b;
});
def("CHAR", (args, ctx, eng) => {
  const n = intArg(args[0], eng, ctx); if (isErr(n)) return n;
  return n < 1 || n > 255 ? VALUE : String.fromCharCode(n);
});
def("CODE", (args, ctx, eng) => {
  const s = str(args[0], eng, ctx); if (isErr(s)) return s;
  return s ? s.charCodeAt(0) : VALUE;
});

// ---------------------------------------------------------------------------
// Dates
// ---------------------------------------------------------------------------

function todaySerial(eng) {
  const d = eng.now;
  return ymdToSerial(d.getFullYear(), d.getMonth() + 1, d.getDate());
}
def("TODAY", (args, ctx, eng) => todaySerial(eng));
def("NOW", (args, ctx, eng) => {
  const d = eng.now;
  return todaySerial(eng) + (d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds()) / 86400;
});
def("DATE", (args, ctx, eng) => {
  let y = intArg(args[0], eng, ctx);
  const m = intArg(args[1], eng, ctx), d = intArg(args[2], eng, ctx);
  const e = firstErr(y, m, d); if (e) return e;
  if (y < 0 || y > 9999) return NUM;
  if (y < 1900) y += 1900;
  const s = ymdToSerial(y, m, d);
  return s < 0 ? NUM : s;
});
def("TIME", (args, ctx, eng) => {
  const h = intArg(args[0], eng, ctx), m = intArg(args[1], eng, ctx), s = intArg(args[2], eng, ctx);
  const e = firstErr(h, m, s); if (e) return e;
  const t = (h * 3600 + m * 60 + s) / 86400;
  return t < 0 ? NUM : t - Math.floor(t);
});
function datePart(args, ctx, eng, f) {
  const x = num(args[0], eng, ctx);
  if (isErr(x)) return x;
  if (x < 0) return NUM;
  return f(serialToParts(x));
}
def("YEAR", (a, c, e) => datePart(a, c, e, (p) => p.y));
def("MONTH", (a, c, e) => datePart(a, c, e, (p) => p.m));
def("DAY", (a, c, e) => datePart(a, c, e, (p) => p.d));
def("HOUR", (a, c, e) => datePart(a, c, e, (p) => p.H));
def("MINUTE", (a, c, e) => datePart(a, c, e, (p) => p.M));
def("SECOND", (a, c, e) => datePart(a, c, e, (p) => p.S));
def("WEEKDAY", (args, ctx, eng) => {
  const t = args.length > 1 && args[1] != null ? intArg(args[1], eng, ctx) : 1;
  if (isErr(t)) return t;
  // return_type -> weekday (0=Sunday) that counts as day 1 (type 3 counts from 0)
  const first = t === 1 || t === 17 ? 0 : t === 2 || t === 3 ? 1 : t >= 11 && t <= 16 ? t - 10 : null;
  if (first == null) return NUM;
  return datePart(args, ctx, eng, (p) => ((p.dow - first + 7) % 7) + (t === 3 ? 0 : 1));
});
function addMonths(serial, months, endOfMonth) {
  const p = serialToParts(Math.floor(serial));
  const total = p.y * 12 + (p.m - 1) + months;
  const y = Math.floor(total / 12), m = (total % 12) + 1;
  const dim = new Date(Date.UTC(y, m, 0)).getUTCDate();
  return ymdToSerial(y, m, endOfMonth ? dim : Math.min(p.d, dim));
}
def("EDATE", (args, ctx, eng) => {
  const s = num(args[0], eng, ctx), m = intArg(args[1], eng, ctx);
  const e = firstErr(s, m); if (e) return e;
  return s < 0 ? NUM : addMonths(s, m, false);
});
def("EOMONTH", (args, ctx, eng) => {
  const s = num(args[0], eng, ctx), m = intArg(args[1], eng, ctx);
  const e = firstErr(s, m); if (e) return e;
  return s < 0 ? NUM : addMonths(s, m, true);
});
def("DATEDIF", (args, ctx, eng) => {
  const a = num(args[0], eng, ctx), b = num(args[1], eng, ctx), u = str(args[2], eng, ctx);
  const e = firstErr(a, b, u); if (e) return e;
  if (a > b) return NUM;
  const p = serialToParts(Math.floor(a)), q = serialToParts(Math.floor(b));
  let months = (q.y - p.y) * 12 + (q.m - p.m);
  if (q.d < p.d) months--;
  switch (u.toUpperCase()) {
    case "Y": return Math.floor(months / 12);
    case "M": return months;
    case "D": return Math.floor(b) - Math.floor(a);
    case "YM": return ((months % 12) + 12) % 12;
    case "MD": {
      let d = q.d - p.d;
      if (d < 0) d += new Date(Date.UTC(q.y, q.m - 1, 0)).getUTCDate();
      return d;
    }
    case "YD": {
      let start = ymdToSerial(q.y, p.m, p.d);
      if (start > Math.floor(b)) start = ymdToSerial(q.y - 1, p.m, p.d);
      return Math.floor(b) - start;
    }
    default: return NUM;
  }
});
def("DAYS", (args, ctx, eng) => {
  const a = num(args[0], eng, ctx), b = num(args[1], eng, ctx);
  return firstErr(a, b) || Math.floor(a) - Math.floor(b);
});
def("DATEVALUE", (args, ctx, eng) => {
  const s = str(args[0], eng, ctx);
  if (isErr(s)) return s;
  const p = parseCellInput(s.trim());
  if (p && p.t === "d") return Math.floor(isoToSerial(p.v));
  return VALUE;
});

/** Excel weekday of a serial (0=Sunday), consistent with WEEKDAY (serial 1 is a Sunday). */
const dowOf = (s) => (((Math.floor(s) - 1) % 7) + 7) % 7;

/** Weekend mask (index 0=Sunday, true = weekend) from a WORKDAY.INTL/NETWORKDAYS.INTL weekend argument. */
function weekendMask(v, eng, ctx) {
  const x = v == null ? 1 : eng.scalar(v, ctx);
  if (x == null) return weekendMask(null);
  if (isErr(x)) return x;
  const mask = new Array(7).fill(false);
  if (typeof x === "string") {
    if (!/^[01]{7}$/.test(x) || x === "1111111") return VALUE;
    for (let i = 0; i < 7; i++) mask[(i + 1) % 7] = x[i] === "1"; // string starts on Monday
    return mask;
  }
  const n = toNum(x);
  if (isErr(n)) return n;
  const t = Math.trunc(n);
  if (t >= 1 && t <= 7) { mask[(t + 5) % 7] = true; mask[(t + 6) % 7] = true; }
  else if (t >= 11 && t <= 17) mask[t - 11] = true;
  else return NUM;
  return mask;
}

/** Set of holiday serials (whole days) from a range/array/scalar argument. */
function holidaySet(v, eng) {
  const set = new Set();
  if (v == null) return set;
  let err = null;
  const add = (x) => {
    if (err || x == null || x === "") return;
    if (isErr(x)) { err = x; return; }
    const n = toNum(x);
    if (isErr(n)) { err = n; return; }
    set.add(Math.floor(n));
  };
  if (v instanceof Ref) eng.eachInRef(v, add);
  else if (Array.isArray(v)) for (const row of v) for (const x of row) add(x);
  else add(v);
  return err || set;
}

function networkDays(args, ctx, eng, intl) {
  const a = num(args[0], eng, ctx), b = num(args[1], eng, ctx);
  const e = firstErr(a, b); if (e) return e;
  const mask = intl ? weekendMask(args[2], eng, ctx) : weekendMask(null);
  if (isErr(mask)) return mask;
  const hol = holidaySet(intl ? args[3] : args[2], eng);
  if (isErr(hol)) return hol;
  if (a < 0 || b < 0) return NUM;
  const s = Math.floor(a), t = Math.floor(b);
  const lo = Math.min(s, t), hi = Math.max(s, t);
  let n = 0;
  for (let d = lo; d <= hi; d++) if (!mask[dowOf(d)] && !hol.has(d)) n++;
  return s <= t ? n : -n;
}
def("NETWORKDAYS", (a, c, e) => networkDays(a, c, e, false));
def("NETWORKDAYS.INTL", (a, c, e) => networkDays(a, c, e, true));

function workday(args, ctx, eng, intl) {
  const a = num(args[0], eng, ctx), k = num(args[1], eng, ctx);
  const e = firstErr(a, k); if (e) return e;
  const mask = intl ? weekendMask(args[2], eng, ctx) : weekendMask(null);
  if (isErr(mask)) return mask;
  const hol = holidaySet(intl ? args[3] : args[2], eng);
  if (isErr(hol)) return hol;
  if (a < 0) return NUM;
  let d = Math.floor(a), left = Math.trunc(k);
  const step = left < 0 ? -1 : 1;
  left = Math.abs(left);
  while (left > 0) {
    d += step;
    if (d < 0 || d > 2958465) return NUM;
    if (!mask[dowOf(d)] && !hol.has(d)) left--;
  }
  return d;
}
def("WORKDAY", (a, c, e) => workday(a, c, e, false));
def("WORKDAY.INTL", (a, c, e) => workday(a, c, e, true));

function isoWeek(s) {
  const d = Math.floor(s);
  const wd = ((dowOf(d) + 6) % 7) + 1; // Monday=1 .. Sunday=7
  const th = d - wd + 4; // Thursday of the same ISO week
  const y = serialToParts(th).y;
  return Math.floor((th - ymdToSerial(y, 1, 1)) / 7) + 1;
}
def("WEEKNUM", (args, ctx, eng) => {
  const s = num(args[0], eng, ctx);
  const t = args.length > 1 && args[1] != null ? intArg(args[1], eng, ctx) : 1;
  const e = firstErr(s, t); if (e) return e;
  if (s < 0) return NUM;
  if (t === 21) return isoWeek(s);
  const first = t === 1 || t === 17 ? 0 : t === 2 || t === 11 ? 1 : t >= 12 && t <= 16 ? t - 10 : null;
  if (first == null) return NUM;
  const d = Math.floor(s);
  const jan1 = ymdToSerial(serialToParts(d).y, 1, 1);
  const offset = (dowOf(jan1) - first + 7) % 7;
  return Math.floor((d - jan1 + offset) / 7) + 1;
});
def("ISOWEEKNUM", (args, ctx, eng) => {
  const s = num(args[0], eng, ctx);
  if (isErr(s)) return s;
  return s < 0 ? NUM : isoWeek(s);
});

// ---------------------------------------------------------------------------
// Statistics: percentiles, quartiles, mode
// ---------------------------------------------------------------------------

function percentile(arr, k, eng, exclusive) {
  const ns = numbers([arr], eng);
  if (isErr(ns)) return ns;
  if (isErr(k)) return k;
  const n = ns.length;
  if (!n) return NUM;
  ns.sort((a, b) => a - b);
  let h;
  if (exclusive) {
    if (k <= 0 || k >= 1) return NUM;
    h = (n + 1) * k - 1;
    if (h < 0 || h > n - 1) return NUM;
  } else {
    if (k < 0 || k > 1) return NUM;
    h = (n - 1) * k;
  }
  const lo = Math.floor(h);
  const frac = h - lo;
  return lo + 1 < n ? ns[lo] + frac * (ns[lo + 1] - ns[lo]) : ns[lo];
}
def("PERCENTILE PERCENTILE.INC", (args, ctx, eng) => percentile(args[0], num(args[1], eng, ctx), eng, false));
def("PERCENTILE.EXC", (args, ctx, eng) => percentile(args[0], num(args[1], eng, ctx), eng, true));
def("QUARTILE QUARTILE.INC", (args, ctx, eng) => {
  const q = intArg(args[1], eng, ctx);
  if (isErr(q)) return q;
  return q < 0 || q > 4 ? NUM : percentile(args[0], q / 4, eng, false);
});
def("QUARTILE.EXC", (args, ctx, eng) => {
  const q = intArg(args[1], eng, ctx);
  if (isErr(q)) return q;
  return q < 1 || q > 3 ? NUM : percentile(args[0], q / 4, eng, true);
});
def("MODE MODE.SNGL", (args, ctx, eng) => {
  const ns = numbers(args, eng);
  if (isErr(ns)) return ns;
  const counts = new Map();
  let best = null, bestCount = 1;
  for (const x of ns) {
    const k = round15(x);
    const c = (counts.get(k) || 0) + 1;
    counts.set(k, c);
    if (c > bestCount) { bestCount = c; best = x; }
  }
  if (best == null) return NA;
  // ties: the value that occurs first in the data wins (like Excel)
  for (const x of ns) if (counts.get(round15(x)) === bestCount) return x;
  return best;
});

// ---------------------------------------------------------------------------
// Financial
// ---------------------------------------------------------------------------

/** Evaluate numeric arguments (missing optional ones default to 0). Returns an array or an error. */
function numArgs(args, eng, ctx, count) {
  const out = [];
  for (let i = 0; i < count; i++) {
    const x = args[i] == null ? 0 : num(args[i], eng, ctx);
    if (isErr(x)) return x;
    out.push(x);
  }
  return out;
}
const finite = (x) => (isFinite(x) ? x : NUM);

/** pv*(1+r)^n + pmt*(1+r*type)*((1+r)^n-1)/r + fv: zero when the cash flows balance. */
function tvmBalance(r, n, pmt, pv, fv, type) {
  if (Math.abs(r) < 1e-12) return pv + pmt * n + fv;
  const g = Math.pow(1 + r, n);
  return pv * g + pmt * (1 + r * type) * (g - 1) / r + fv;
}

def("PMT", (args, ctx, eng) => {
  const a = numArgs(args, eng, ctx, 5); if (isErr(a)) return a;
  const [r, n, pv, fv, t] = a, type = t ? 1 : 0;
  if (n === 0) return NUM;
  if (r === 0) return finite(-(pv + fv) / n);
  const g = Math.pow(1 + r, n);
  return finite(-(r * (fv + pv * g)) / ((1 + r * type) * (g - 1)));
});
def("FV", (args, ctx, eng) => {
  const a = numArgs(args, eng, ctx, 5); if (isErr(a)) return a;
  const [r, n, pmt, pv, t] = a, type = t ? 1 : 0;
  if (r === 0) return finite(-(pv + pmt * n));
  const g = Math.pow(1 + r, n);
  return finite(-(pv * g + pmt * (1 + r * type) * (g - 1) / r));
});
def("PV", (args, ctx, eng) => {
  const a = numArgs(args, eng, ctx, 5); if (isErr(a)) return a;
  const [r, n, pmt, fv, t] = a, type = t ? 1 : 0;
  if (r === 0) return finite(-(fv + pmt * n));
  const g = Math.pow(1 + r, n);
  return finite(-(fv + pmt * (1 + r * type) * (g - 1) / r) / g);
});
def("NPER", (args, ctx, eng) => {
  const a = numArgs(args, eng, ctx, 5); if (isErr(a)) return a;
  const [r, pmt, pv, fv, t] = a, type = t ? 1 : 0;
  if (r === 0) return pmt === 0 ? NUM : finite(-(pv + fv) / pmt);
  const num_ = pmt * (1 + r * type) - fv * r, den = pv * r + pmt * (1 + r * type);
  if (den === 0 || num_ / den <= 0 || r <= -1) return NUM;
  return finite(Math.log(num_ / den) / Math.log(1 + r));
});

/** Find a root of f near `guess`: Newton's method, then bisection over a bracket. */
function solveRate(f, guess) {
  let r = guess;
  for (let i = 0; i < 100; i++) {
    const y = f(r);
    if (!isFinite(y)) break;
    if (Math.abs(y) < 1e-10) return r;
    const h = Math.max(1e-7, Math.abs(r) * 1e-7);
    const dy = (f(r + h) - f(r - h)) / (2 * h);
    if (!isFinite(dy) || dy === 0) break;
    const next = r - y / dy;
    if (!isFinite(next) || next <= -1) break;
    if (Math.abs(next - r) < 1e-12) return next;
    r = next;
  }
  if (isFinite(r) && r > -1 && Math.abs(f(r)) < 1e-7) return r;
  // Bisection: scan for a sign change, nearest to the guess first.
  const pts = [-0.999999, -0.99, -0.9, -0.75, -0.5, -0.3, -0.2, -0.1, -0.05, 0, 0.05, 0.1, 0.2, 0.3, 0.5,
    0.75, 1, 1.5, 2, 3, 5, 10, 20, 50, 100, 1000];
  const brackets = [];
  for (let i = 0; i + 1 < pts.length; i++) {
    const fa = f(pts[i]), fb = f(pts[i + 1]);
    if (isFinite(fa) && isFinite(fb) && Math.sign(fa) !== Math.sign(fb)) brackets.push([pts[i], pts[i + 1], fa]);
  }
  if (!brackets.length) return null;
  brackets.sort((x, y) => Math.min(Math.abs(x[0] - guess), Math.abs(x[1] - guess)) - Math.min(Math.abs(y[0] - guess), Math.abs(y[1] - guess)));
  let [lo, hi, flo] = brackets[0];
  for (let i = 0; i < 200; i++) {
    const mid = (lo + hi) / 2, fm = f(mid);
    if (fm === 0 || (hi - lo) / 2 < 1e-15) return mid;
    if (Math.sign(fm) === Math.sign(flo)) { lo = mid; flo = fm; } else hi = mid;
  }
  return (lo + hi) / 2;
}

def("RATE", (args, ctx, eng) => {
  const a = numArgs(args, eng, ctx, 5); if (isErr(a)) return a;
  const [n, pmt, pv, fv, t] = a, type = t ? 1 : 0;
  const guess = args[5] == null ? 0.1 : num(args[5], eng, ctx);
  if (isErr(guess)) return guess;
  if (n <= 0) return NUM;
  const r = solveRate((x) => tvmBalance(x, n, pmt, pv, fv, type), guess);
  return r == null ? NUM : r;
});
def("NPV", (args, ctx, eng) => {
  const r = num(args[0], eng, ctx);
  if (isErr(r)) return r;
  if (r === -1) return DIV0;
  let s = 0, i = 0;
  const err = eachNumber(args.slice(1), eng, (v) => { i++; s += v / Math.pow(1 + r, i); });
  return err || finite(s);
});
def("IRR", (args, ctx, eng) => {
  const vs = numbers([args[0]], eng);
  if (isErr(vs)) return vs;
  const guess = args[1] == null ? 0.1 : num(args[1], eng, ctx);
  if (isErr(guess)) return guess;
  if (!vs.some((v) => v > 0) || !vs.some((v) => v < 0)) return NUM;
  const npv = (r) => vs.reduce((s, v, i) => s + v / Math.pow(1 + r, i), 0);
  const r = solveRate(npv, guess);
  return r == null ? NUM : r;
});

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/** Names of supported functions (for UI hints / tests). */
export const SUPPORTED_FUNCTIONS = Object.keys(FUNCS).sort();

/**
 * Signature and one-line description of every supported function (for the
 * formula bar's autocomplete / signature tooltip): { NAME: { sig, desc } }.
 */
export const FUNCTION_INFO = Object.fromEntries(Object.entries({
  ABS: ["ABS(number)", "Returns the absolute value of a number."],
  AND: ["AND(logical1, [logical2], ...)", "Returns TRUE if all of its arguments are TRUE."],
  AVERAGE: ["AVERAGE(number1, [number2], ...)", "Returns the average (arithmetic mean) of the numbers."],
  AVERAGEA: ["AVERAGEA(value1, [value2], ...)", "Returns the average of the values, counting text as 0 and TRUE as 1."],
  AVERAGEIF: ["AVERAGEIF(range, criteria, [average_range])", "Averages the cells that meet a condition."],
  AVERAGEIFS: ["AVERAGEIFS(average_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)", "Averages the cells that meet all the given conditions."],
  CEILING: ["CEILING(number, [significance])", "Rounds a number up to the nearest multiple of significance."],
  "CEILING.MATH": ["CEILING.MATH(number, [significance], [mode])", "Rounds a number up to the nearest integer or multiple of significance."],
  CHAR: ["CHAR(number)", "Returns the character with the given code number."],
  CHOOSE: ["CHOOSE(index_num, value1, [value2], ...)", "Picks a value from the list by its position."],
  CLEAN: ["CLEAN(text)", "Removes non-printable characters from text."],
  CODE: ["CODE(text)", "Returns the code number of the first character of text."],
  COLUMN: ["COLUMN([reference])", "Returns the column number of a reference."],
  COLUMNS: ["COLUMNS(array)", "Returns the number of columns in a range or array."],
  CONCAT: ["CONCAT(text1, [text2], ...)", "Joins several texts or ranges into one text."],
  CONCATENATE: ["CONCATENATE(text1, [text2], ...)", "Joins several texts into one text."],
  COUNT: ["COUNT(value1, [value2], ...)", "Counts the cells that contain numbers."],
  COUNTA: ["COUNTA(value1, [value2], ...)", "Counts the cells that are not empty."],
  COUNTBLANK: ["COUNTBLANK(range)", "Counts the empty cells in a range."],
  COUNTIF: ["COUNTIF(range, criteria)", "Counts the cells that meet a condition."],
  COUNTIFS: ["COUNTIFS(criteria_range1, criteria1, [criteria_range2, criteria2], ...)", "Counts the cells that meet all the given conditions."],
  DATE: ["DATE(year, month, day)", "Returns the date for the given year, month and day."],
  DATEDIF: ["DATEDIF(start_date, end_date, unit)", "Returns the difference between two dates in years (\"Y\"), months (\"M\"), days (\"D\"), or \"YM\", \"YD\", \"MD\"."],
  DATEVALUE: ["DATEVALUE(date_text)", "Converts a date written as text into a date."],
  DAY: ["DAY(serial_number)", "Returns the day of the month (1-31) of a date."],
  DAYS: ["DAYS(end_date, start_date)", "Returns the number of days between two dates."],
  EDATE: ["EDATE(start_date, months)", "Returns the date a number of months before or after a date."],
  EOMONTH: ["EOMONTH(start_date, months)", "Returns the last day of the month a number of months before or after a date."],
  EXACT: ["EXACT(text1, text2)", "Checks whether two texts are exactly the same (case-sensitive)."],
  EXP: ["EXP(number)", "Returns e raised to the power of a number."],
  FALSE: ["FALSE()", "Returns the logical value FALSE."],
  FIND: ["FIND(find_text, within_text, [start_num])", "Returns the position of one text inside another (case-sensitive)."],
  FLOOR: ["FLOOR(number, [significance])", "Rounds a number down to the nearest multiple of significance."],
  "FLOOR.MATH": ["FLOOR.MATH(number, [significance], [mode])", "Rounds a number down to the nearest integer or multiple of significance."],
  FV: ["FV(rate, nper, pmt, [pv], [type])", "Returns the future value of an investment with constant payments and interest rate."],
  HLOOKUP: ["HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])", "Looks for a value in the top row of a table and returns a value from the given row."],
  HOUR: ["HOUR(serial_number)", "Returns the hour (0-23) of a time."],
  IF: ["IF(logical_test, [value_if_true], [value_if_false])", "Returns one value if a condition is TRUE and another if it is FALSE."],
  IFERROR: ["IFERROR(value, value_if_error)", "Returns value_if_error if the value is an error, otherwise the value."],
  IFNA: ["IFNA(value, value_if_na)", "Returns value_if_na if the value is #N/A, otherwise the value."],
  IFS: ["IFS(logical_test1, value_if_true1, [logical_test2, value_if_true2], ...)", "Returns the value for the first condition that is TRUE."],
  INDEX: ["INDEX(array, row_num, [column_num])", "Returns the value at the given row and column of a range."],
  INT: ["INT(number)", "Rounds a number down to the nearest integer."],
  IRR: ["IRR(values, [guess])", "Returns the internal rate of return of a series of cash flows."],
  ISBLANK: ["ISBLANK(value)", "Returns TRUE if the cell is empty."],
  ISERR: ["ISERR(value)", "Returns TRUE if the value is any error except #N/A."],
  ISERROR: ["ISERROR(value)", "Returns TRUE if the value is any error."],
  ISEVEN: ["ISEVEN(number)", "Returns TRUE if the number is even."],
  ISLOGICAL: ["ISLOGICAL(value)", "Returns TRUE if the value is TRUE or FALSE."],
  ISNA: ["ISNA(value)", "Returns TRUE if the value is the #N/A error."],
  ISNONTEXT: ["ISNONTEXT(value)", "Returns TRUE if the value is not text."],
  ISNUMBER: ["ISNUMBER(value)", "Returns TRUE if the value is a number."],
  ISODD: ["ISODD(number)", "Returns TRUE if the number is odd."],
  ISOWEEKNUM: ["ISOWEEKNUM(date)", "Returns the ISO week number of the year for a date."],
  ISTEXT: ["ISTEXT(value)", "Returns TRUE if the value is text."],
  LARGE: ["LARGE(array, k)", "Returns the k-th largest number in a set of data."],
  LEFT: ["LEFT(text, [num_chars])", "Returns the first characters of a text."],
  LEN: ["LEN(text)", "Returns the number of characters in a text."],
  LN: ["LN(number)", "Returns the natural logarithm of a number."],
  LOG: ["LOG(number, [base])", "Returns the logarithm of a number to the given base (10 by default)."],
  LOG10: ["LOG10(number)", "Returns the base-10 logarithm of a number."],
  LOWER: ["LOWER(text)", "Converts text to lowercase."],
  MATCH: ["MATCH(lookup_value, lookup_array, [match_type])", "Returns the position of a value in a range."],
  MAX: ["MAX(number1, [number2], ...)", "Returns the largest number."],
  MAXIFS: ["MAXIFS(max_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)", "Returns the largest number among the cells that meet all the given conditions."],
  MEDIAN: ["MEDIAN(number1, [number2], ...)", "Returns the middle number of the given numbers."],
  MID: ["MID(text, start_num, num_chars)", "Returns characters from the middle of a text."],
  MIN: ["MIN(number1, [number2], ...)", "Returns the smallest number."],
  MINIFS: ["MINIFS(min_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)", "Returns the smallest number among the cells that meet all the given conditions."],
  MINUTE: ["MINUTE(serial_number)", "Returns the minute (0-59) of a time."],
  MOD: ["MOD(number, divisor)", "Returns the remainder after division."],
  MODE: ["MODE(number1, [number2], ...)", "Returns the most frequently occurring number."],
  "MODE.SNGL": ["MODE.SNGL(number1, [number2], ...)", "Returns the most frequently occurring number."],
  MONTH: ["MONTH(serial_number)", "Returns the month (1-12) of a date."],
  MROUND: ["MROUND(number, multiple)", "Rounds a number to the nearest multiple."],
  N: ["N(value)", "Converts a value to a number (text becomes 0)."],
  NA: ["NA()", "Returns the #N/A error (value not available)."],
  NETWORKDAYS: ["NETWORKDAYS(start_date, end_date, [holidays])", "Counts the working days (Monday to Friday) between two dates, excluding holidays."],
  "NETWORKDAYS.INTL": ["NETWORKDAYS.INTL(start_date, end_date, [weekend], [holidays])", "Counts the working days between two dates with custom weekend days."],
  NOT: ["NOT(logical)", "Reverses a logical value: TRUE becomes FALSE and FALSE becomes TRUE."],
  NOW: ["NOW()", "Returns the current date and time."],
  NPER: ["NPER(rate, pmt, pv, [fv], [type])", "Returns the number of payment periods for a loan or investment."],
  NPV: ["NPV(rate, value1, [value2], ...)", "Returns the net present value of future cash flows at a discount rate."],
  OR: ["OR(logical1, [logical2], ...)", "Returns TRUE if any argument is TRUE."],
  PERCENTILE: ["PERCENTILE(array, k)", "Returns the k-th percentile (k from 0 to 1) of the values."],
  "PERCENTILE.EXC": ["PERCENTILE.EXC(array, k)", "Returns the k-th percentile of the values, with k strictly between 0 and 1."],
  "PERCENTILE.INC": ["PERCENTILE.INC(array, k)", "Returns the k-th percentile (k from 0 to 1) of the values."],
  PI: ["PI()", "Returns the number pi (3.14159...)."],
  PMT: ["PMT(rate, nper, pv, [fv], [type])", "Returns the periodic payment for a loan with constant payments and interest rate."],
  POWER: ["POWER(number, power)", "Returns a number raised to a power."],
  PRODUCT: ["PRODUCT(number1, [number2], ...)", "Multiplies all the numbers."],
  PROPER: ["PROPER(text)", "Capitalizes the first letter of each word."],
  PV: ["PV(rate, nper, pmt, [fv], [type])", "Returns the present value of an investment."],
  QUARTILE: ["QUARTILE(array, quart)", "Returns a quartile (0-4) of the values."],
  "QUARTILE.EXC": ["QUARTILE.EXC(array, quart)", "Returns a quartile (1-3) of the values, excluding the median from the halves."],
  "QUARTILE.INC": ["QUARTILE.INC(array, quart)", "Returns a quartile (0-4) of the values."],
  QUOTIENT: ["QUOTIENT(numerator, denominator)", "Returns the integer part of a division."],
  RAND: ["RAND()", "Returns a random number between 0 and 1."],
  RANDBETWEEN: ["RANDBETWEEN(bottom, top)", "Returns a random integer between two numbers."],
  RANK: ["RANK(number, ref, [order])", "Returns the rank of a number in a list (0 = descending, 1 = ascending)."],
  "RANK.EQ": ["RANK.EQ(number, ref, [order])", "Returns the rank of a number in a list (0 = descending, 1 = ascending)."],
  RATE: ["RATE(nper, pmt, pv, [fv], [type], [guess])", "Returns the interest rate per period of a loan or investment."],
  REPLACE: ["REPLACE(old_text, start_num, num_chars, new_text)", "Replaces part of a text at a given position."],
  REPT: ["REPT(text, number_times)", "Repeats text a given number of times."],
  RIGHT: ["RIGHT(text, [num_chars])", "Returns the last characters of a text."],
  ROUND: ["ROUND(number, num_digits)", "Rounds a number to a given number of digits."],
  ROUNDDOWN: ["ROUNDDOWN(number, num_digits)", "Rounds a number down, toward zero."],
  ROUNDUP: ["ROUNDUP(number, num_digits)", "Rounds a number up, away from zero."],
  ROW: ["ROW([reference])", "Returns the row number of a reference."],
  ROWS: ["ROWS(array)", "Returns the number of rows in a range or array."],
  SEARCH: ["SEARCH(find_text, within_text, [start_num])", "Returns the position of one text inside another (not case-sensitive, allows wildcards)."],
  SECOND: ["SECOND(serial_number)", "Returns the second (0-59) of a time."],
  SIGN: ["SIGN(number)", "Returns 1 for positive numbers, -1 for negative numbers and 0 for zero."],
  SMALL: ["SMALL(array, k)", "Returns the k-th smallest number in a set of data."],
  SQRT: ["SQRT(number)", "Returns the square root of a number."],
  STDEV: ["STDEV(number1, [number2], ...)", "Estimates the standard deviation of a sample."],
  "STDEV.P": ["STDEV.P(number1, [number2], ...)", "Returns the standard deviation of an entire population."],
  "STDEV.S": ["STDEV.S(number1, [number2], ...)", "Estimates the standard deviation of a sample."],
  STDEVP: ["STDEVP(number1, [number2], ...)", "Returns the standard deviation of an entire population."],
  SUBSTITUTE: ["SUBSTITUTE(text, old_text, new_text, [instance_num])", "Replaces existing text with new text."],
  SUM: ["SUM(number1, [number2], ...)", "Adds all the numbers."],
  SUMIF: ["SUMIF(range, criteria, [sum_range])", "Adds the cells that meet a condition."],
  SUMIFS: ["SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)", "Adds the cells that meet all the given conditions."],
  SUMPRODUCT: ["SUMPRODUCT(array1, [array2], ...)", "Multiplies matching items of the ranges and adds up the products."],
  SUMSQ: ["SUMSQ(number1, [number2], ...)", "Adds the squares of the numbers."],
  SWITCH: ["SWITCH(expression, value1, result1, [value2, result2], ..., [default])", "Returns the result matching the first value equal to the expression."],
  T: ["T(value)", "Returns the text if the value is text, otherwise empty text."],
  TEXT: ["TEXT(value, format_text)", "Formats a number or date as text using a number format, e.g. \"0.00\" or \"yyyy-mm-dd\"."],
  TEXTJOIN: ["TEXTJOIN(delimiter, ignore_empty, text1, [text2], ...)", "Joins texts and ranges with a delimiter, optionally skipping empty cells."],
  TIME: ["TIME(hour, minute, second)", "Returns the time for the given hour, minute and second."],
  TODAY: ["TODAY()", "Returns today's date."],
  TRIM: ["TRIM(text)", "Removes extra spaces from text."],
  TRUE: ["TRUE()", "Returns the logical value TRUE."],
  TRUNC: ["TRUNC(number, [num_digits])", "Cuts a number down to an integer or a number of digits, toward zero."],
  UPPER: ["UPPER(text)", "Converts text to uppercase."],
  VALUE: ["VALUE(text)", "Converts text that looks like a number into a number."],
  VAR: ["VAR(number1, [number2], ...)", "Estimates the variance of a sample."],
  "VAR.P": ["VAR.P(number1, [number2], ...)", "Returns the variance of an entire population."],
  "VAR.S": ["VAR.S(number1, [number2], ...)", "Estimates the variance of a sample."],
  VARP: ["VARP(number1, [number2], ...)", "Returns the variance of an entire population."],
  VLOOKUP: ["VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])", "Looks for a value in the first column of a table and returns a value from the given column."],
  WEEKDAY: ["WEEKDAY(serial_number, [return_type])", "Returns the day of the week of a date (1 = Sunday by default)."],
  WEEKNUM: ["WEEKNUM(serial_number, [return_type])", "Returns the week number of the year for a date."],
  WORKDAY: ["WORKDAY(start_date, days, [holidays])", "Returns the date a number of working days before or after a date."],
  "WORKDAY.INTL": ["WORKDAY.INTL(start_date, days, [weekend], [holidays])", "Returns the date a number of working days away, with custom weekend days."],
  XLOOKUP: ["XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])", "Looks for a value in a range and returns the matching item from another range."],
  XOR: ["XOR(logical1, [logical2], ...)", "Returns TRUE if an odd number of arguments are TRUE."],
  YEAR: ["YEAR(serial_number)", "Returns the year of a date."],
}).map(([k, [sig, desc]]) => [k, { sig, desc }]));

/**
 * For formula text up to the caret ("=SUM(A1, IF(B1>0,"), find the innermost
 * unclosed function call and the 0-based index of the argument being typed.
 * Returns { name, argIndex } (name upper-case, without _xlfn.) or null.
 */
export function functionArgIndex(textBeforeCaret) {
  const src = String(textBeforeCaret ?? "");
  if (!src.startsWith("=")) return null;
  const stack = []; // { name } for calls, { name: null } for "(" groups, { brace: true } for "{"
  let i = 1;
  while (i < src.length) {
    const ch = src[i];
    if (ch === '"' || ch === "'") {
      // string literal / quoted sheet name: skip to the closing quote ("" or '' escapes)
      let j = i + 1;
      for (; j < src.length; j++) {
        if (src[j] === ch) {
          if (src[j + 1] === ch) { j++; continue; }
          break;
        }
      }
      i = j + 1;
      continue;
    }
    if (/[0-9]/.test(ch)) {
      const m = /^[0-9.]+(?:[eE][+-]?[0-9]+)?/.exec(src.slice(i));
      i += m[0].length;
      continue;
    }
    if (/[A-Za-z_\\]/.test(ch)) {
      const m = /^[A-Za-z_\\][A-Za-z0-9_.]*/.exec(src.slice(i));
      i += m[0].length;
      if (src[i] === "(") { stack.push({ name: normFuncName(m[0]), argIndex: 0 }); i++; }
      continue;
    }
    if (ch === "(") stack.push({ name: null });
    else if (ch === "{") stack.push({ brace: true });
    else if (ch === ")" || ch === "}") stack.pop();
    else if ((ch === "," || ch === ";") && stack.length && stack[stack.length - 1].name) stack[stack.length - 1].argIndex++;
    i++;
  }
  for (let k = stack.length - 1; k >= 0; k--) {
    if (stack[k].name) return { name: stack[k].name, argIndex: stack[k].argIndex };
  }
  return null;
}

/**
 * Evaluate a formula string in the context of a workbook.
 * ctx = { book, sheet?: index (default book.active), row?, col?, now? }.
 * Returns a plain value: number | string | boolean | null | {error:"#..."}.
 * Referenced formula cells are computed on demand (their results are stored).
 */
export function evaluateFormula(formula, ctx = {}) {
  const book = ctx.book || { sheets: [{ name: "Sheet1", cells: {} }] };
  const eng = new Engine(book, { now: ctx.now });
  const si = ctx.sheet ?? book.active ?? 0;
  const parsed = parseFormula(formula);
  if (parsed.error) return { error: "#NAME?" };
  let v;
  try {
    v = eng.ev(parsed.ast, { si, r: ctx.row ?? 0, c: ctx.col ?? 0 });
    if (v instanceof Ref) v = eng.implicit(v, { r: ctx.row ?? 0, c: ctx.col ?? 0 });
    if (Array.isArray(v)) v = v.length && v[0].length ? v[0][0] : null;
  } catch (e) {
    if (e instanceof Unsupported) return { error: "#NAME?" };
    if (e instanceof TooDeep) {
      eng.compute(e.si, e.key, e.cell);
      return evaluateFormula(formula, ctx);
    }
    throw e;
  }
  if (isErr(v)) return { error: v.code };
  return v;
}

/**
 * Recompute every formula cell of the workbook in place.
 * Returns {count, cycles:["Sheet!A1",...], ms}.
 * opts: { now?: Date, changed?: [{sheet:index, addrs:[...]}] }
 */
export function recalcBook(model, opts = {}) {
  const t0 = typeof performance !== "undefined" ? performance.now() : Date.now();
  const sheets = model?.sheets || [];
  if (opts.changed) for (const ch of opts.changed) if (sheets[ch.sheet]) noteCellsChanged(sheets[ch.sheet], ch.addrs);
  const eng = new Engine(model, opts);
  let count = 0;
  sheets.forEach((sheet, si) => {
    const idx = sheetIndex(sheet);
    for (const key of sortedFormulaKeys(idx)) {
      const cell = sheet.cells[key];
      if (!cell || !cell.f) continue;
      eng.compute(si, key, cell);
      count++;
    }
  });
  const t1 = typeof performance !== "undefined" ? performance.now() : Date.now();
  return { count, cycles: [...new Set(eng.cycles)], ms: t1 - t0 };
}

/**
 * Suggest a number format for a newly entered formula (dates for =TODAY(),
 * =DATE(...), =A1+7 where A1 is a date, ...). Returns a format code or null.
 */
export function formulaAutoFormat(model, sheetIndexArg, f) {
  const parsed = parseFormula(f);
  if (parsed.error) return null;
  const sheets = model.sheets || [];
  const byName = new Map(sheets.map((s, i) => [String(s.name).toLowerCase(), i]));
  const fmtOf = (node) => {
    if (!node) return null;
    switch (node.t) {
      case "fn":
        if (["TODAY", "DATE", "EDATE", "EOMONTH", "DATEVALUE"].includes(node.name)) return "yyyy-mm-dd";
        if (node.name === "NOW") return "yyyy-mm-dd hh:mm";
        if (node.name === "TIME") return "h:mm";
        if (["MIN", "MAX", "IF", "IFERROR"].includes(node.name)) {
          for (const a of node.args.slice(node.name === "IF" ? 1 : 0)) { const z = fmtOf(a); if (z) return z; }
        }
        return null;
      case "ref": {
        if (!node.cell) return null;
        const si = node.sheet == null ? sheetIndexArg : byName.get(String(node.sheet).toLowerCase());
        const cell = si == null ? null : sheets[si]?.cells?.[rcToAddr(node.r1, node.c1)];
        if (!cell) return null;
        if (cell.z && isDateFormat(cell.z)) return cell.z;
        if (cell.t === "d") return /T/.test(String(cell.v)) ? "yyyy-mm-dd hh:mm" : "yyyy-mm-dd";
        return null;
      }
      case "bin":
        if (node.op === "+" || node.op === "-") {
          const a = fmtOf(node.a), b = fmtOf(node.b);
          if (a && !b) return isDateOnlyOrDateTimeFormat(a) || node.op === "+" ? a : null;
          if (b && !a && node.op === "+") return b;
        }
        return null;
      default: return null;
    }
  };
  return fmtOf(parsed.ast);
}

// ---------------------------------------------------------------------------
// Reference rewriting
// ---------------------------------------------------------------------------

/** Quote a sheet name for use in a formula when needed. */
export function quoteSheetName(name) {
  const n = String(name);
  if (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(n) && !/^[A-Za-z]{1,3}\d+$/.test(n) && !/^(TRUE|FALSE)$/i.test(n) && !/^R\d*C\d*$/i.test(n)) return n;
  return `'${n.replace(/'/g, "''")}'`;
}

function endText(p, kind) {
  if (kind === "cols") return `${p.ac ? "$" : ""}${indexToCol(p.c)}`;
  if (kind === "rows") return `${p.ar ? "$" : ""}${p.r + 1}`;
  return `${p.ac ? "$" : ""}${indexToCol(p.c)}${p.ar ? "$" : ""}${p.r + 1}`;
}

function refTokenText(sheetText, kind, a, b) {
  return sheetText + endText(a, kind) + (b ? ":" + endText(b, kind) : "");
}

/** Rewrite every reference token of a formula. fn(tok) returns new text or null (unchanged). */
function rewriteRefs(f, fn) {
  const src = String(f);
  if (!src.startsWith("=")) return src;
  let toks;
  try { toks = tokenize(src.slice(1)); } catch { return src; }
  let changed = false;
  const out = toks.map((t) => {
    if (t.k !== "ref") return t.text;
    const r = fn(t);
    if (r == null || r === t.text) return t.text;
    changed = true;
    return r;
  });
  return changed ? "=" + out.join("") : src;
}

/**
 * Shift relative references by (dr, dc), as when copying a formula from one
 * cell to another. References that fall off the sheet become #REF!.
 */
export function translateFormula(f, dr, dc) {
  if (!dr && !dc) return f;
  return rewriteRefs(f, (t) => {
    const mv = (p) => {
      const q = { ...p };
      if (t.kind !== "rows" && !p.ac) q.c = p.c + dc;
      if (t.kind !== "cols" && !p.ar) q.r = p.r + dr;
      return q;
    };
    const a = mv(t.a), b = t.b ? mv(t.b) : null;
    const bad = (p) => (t.kind !== "rows" && (p.c < 0 || p.c >= MAX_COLS)) || (t.kind !== "cols" && (p.r < 0 || p.r >= MAX_ROWS));
    if (bad(a) || (b && bad(b))) return t.sheetText + "#REF!";
    return refTokenText(t.sheetText, t.kind, a, b);
  });
}

/**
 * Adjust references for a structural change on sheet `targetSheet`
 * (axis "row"|"col", insert `count` > 0 at index `at`, or delete -count).
 * `hostSheet` is the name of the sheet the formula lives on.
 */
export function adjustFormulaForShift(f, hostSheet, targetSheet, axis, at, count) {
  const target = String(targetSheet).toLowerCase();
  const host = String(hostSheet).toLowerCase();
  return rewriteRefs(f, (t) => {
    const sheet = t.sheet == null ? host : t.sheet.toLowerCase();
    if (sheet !== target) return null;
    const isRow = axis === "row";
    if (isRow && t.kind === "cols") return null;
    if (!isRow && t.kind === "rows") return null;
    const key = isRow ? "r" : "c";
    const a = { ...t.a }, b = t.b ? { ...t.b } : null;
    const max = isRow ? MAX_ROWS : MAX_COLS;
    if (count > 0) {
      if (a[key] >= at) a[key] += count;
      if (b && b[key] >= at) b[key] += count;
      if (a[key] >= max || (b && b[key] >= max)) return t.sheetText + "#REF!";
    } else {
      const n = -count, end = at + n; // deleted [at, end)
      if (!b) {
        if (a[key] >= at && a[key] < end) return t.sheetText + "#REF!";
        if (a[key] >= end) a[key] -= n;
      } else {
        let lo = Math.min(a[key], b[key]), hi = Math.max(a[key], b[key]);
        if (lo >= at && hi < end) return t.sheetText + "#REF!";
        const nlo = lo < at ? lo : lo >= end ? lo - n : at;
        const nhi = hi < at ? hi : hi >= end ? hi - n : at - 1;
        if (nhi < nlo) return t.sheetText + "#REF!";
        if (a[key] <= b[key]) { a[key] = nlo; b[key] = nhi; } else { a[key] = nhi; b[key] = nlo; }
      }
    }
    return refTokenText(t.sheetText, t.kind, a, b);
  });
}

/** Point references at a renamed sheet. */
export function renameSheetInFormula(f, oldName, newName) {
  const o = String(oldName).toLowerCase();
  return rewriteRefs(f, (t) => {
    if (t.sheet == null || t.sheet.toLowerCase() !== o) return null;
    return refTokenText(quoteSheetName(newName) + "!", t.kind, t.a, t.b);
  });
}

/** References to a deleted sheet become #REF!. */
export function deleteSheetInFormula(f, name) {
  const o = String(name).toLowerCase();
  return rewriteRefs(f, (t) => (t.sheet != null && t.sheet.toLowerCase() === o ? "#REF!" : null));
}

/** Upper-case function names, references and booleans (Excel style). */
export function normalizeFormula(f) {
  const src = String(f);
  if (!src.startsWith("=")) return src;
  let toks;
  try { toks = tokenize(src.slice(1)); } catch { return src; }
  return "=" + toks.map((t) => {
    if (t.k === "func") return t.text.toUpperCase();
    if (t.k === "bool") return t.text.toUpperCase();
    if (t.k === "err") return t.text.toUpperCase();
    if (t.k === "ref") return t.sheetText + t.text.slice(t.sheetText.length).toUpperCase();
    return t.text;
  }).join("");
}

/** Functions Excel stores with an _xlfn. prefix (newer than Excel 2007). */
export const FUTURE_FUNCTIONS = new Set([
  "XLOOKUP", "XMATCH", "IFS", "MAXIFS", "MINIFS", "SWITCH", "TEXTJOIN", "CONCAT", "IFNA", "DAYS",
  "STDEV.S", "STDEV.P", "VAR.S", "VAR.P", "RANK.EQ", "RANK.AVG", "CEILING.MATH", "FLOOR.MATH", "XOR",
  "ISOWEEKNUM", "NUMBERVALUE", "SHEET", "SHEETS", "FORMULATEXT", "UNIQUE", "SEQUENCE", "RANDARRAY",
  "SORTBY", "LET", "LAMBDA", "TEXTBEFORE", "TEXTAFTER", "TEXTSPLIT", "VSTACK", "HSTACK", "TOCOL", "TOROW",
  "CHOOSECOLS", "CHOOSEROWS", "TAKE", "DROP", "EXPAND", "WRAPROWS", "WRAPCOLS", "BITAND", "BITOR",
  "BITXOR", "CONFIDENCE.NORM", "NORM.DIST", "NORM.INV", "PERCENTILE.INC", "PERCENTILE.EXC",
  "QUARTILE.INC", "QUARTILE.EXC", "MODE.SNGL", "AGGREGATE", "ARABIC", "BASE", "DECIMAL",
  "NETWORKDAYS.INTL", "WORKDAY.INTL",
]);
const XLWS_FUNCTIONS = new Set(["FILTER", "SORT"]);

/** Add _xlfn. prefixes for storage in .xlsx files (formula without "="). */
export function addXlfnPrefixes(f) {
  const src = String(f);
  const body = src.replace(/^=/, "");
  let toks;
  try { toks = tokenize(body); } catch { return src; }
  const out = toks.map((t) => {
    if (t.k !== "func") return t.text;
    if (/^_xl/i.test(t.text)) return t.text;
    if (XLWS_FUNCTIONS.has(t.v)) return `_xlfn._xlws.${t.text}`;
    if (FUTURE_FUNCTIONS.has(t.v)) return `_xlfn.${t.text}`;
    return t.text;
  }).join("");
  return (src.startsWith("=") ? "=" : "") + out;
}

/** Remove _xlfn./_xlws. prefixes (outside string literals). */
export function stripXlfnPrefixes(f) {
  const src = String(f);
  if (!/_xl(fn|ws)\./i.test(src)) return src;
  return src.split(/("(?:[^"]|"")*")/).map((part, i) => (i % 2 ? part : part.replace(/_xl(fn|ws)\./gi, ""))).join("");
}

// ---------------------------------------------------------------------------
// Structural operations on a workbook
// ---------------------------------------------------------------------------

function shiftCoord(v, at, count) {
  if (count > 0) return v >= at ? v + count : v;
  const n = -count;
  if (v >= at && v < at + n) return null;
  return v >= at + n ? v - n : v;
}

/**
 * Insert (count > 0) or delete (count < 0) rows/columns on sheet `si`.
 * Shifts cells, merges and column widths and rewrites formulas in all sheets.
 */
export function shiftCells(book, si, axis, at, count) {
  const sheet = book.sheets[si];
  if (!sheet || !count) return;
  const isRow = axis === "row";
  const out = {};
  for (const key of Object.keys(sheet.cells || {})) {
    const { r, c } = rcOfKey(key);
    const nr = isRow ? shiftCoord(r, at, count) : r;
    const nc = isRow ? c : shiftCoord(c, at, count);
    if (nr == null || nc == null) continue;
    if (nr >= MAX_ROWS || nc >= MAX_COLS) continue;
    out[rcToAddr(nr, nc)] = sheet.cells[key];
  }
  invalidateSheet(sheet);
  sheet.cells = out;
  if (sheet.merges && sheet.merges.length) {
    const merges = [];
    for (const ref of sheet.merges) {
      const m = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i.exec(ref);
      if (!m) continue;
      let r1 = +m[2] - 1, r2 = +m[4] - 1, c1 = colToIndex(m[1]), c2 = colToIndex(m[3]);
      if (isRow) {
        const nr = shiftRange(r1, r2, at, count); if (!nr) continue; [r1, r2] = nr;
      } else {
        const nc = shiftRange(c1, c2, at, count); if (!nc) continue; [c1, c2] = nc;
      }
      if (r1 === r2 && c1 === c2) continue;
      merges.push(`${rcToAddr(r1, c1)}:${rcToAddr(r2, c2)}`);
    }
    sheet.merges = merges;
  }
  if (!isRow && sheet.colWidths) {
    const cw = {};
    for (const [col, w] of Object.entries(sheet.colWidths)) {
      const nc = shiftCoord(colToIndex(col), at, count);
      if (nc != null && nc < MAX_COLS) cw[indexToCol(nc)] = w;
    }
    sheet.colWidths = cw;
  }
  if (typeof sheet.autoFilter === "string" && sheet.autoFilter) {
    const af = shiftRangeText(sheet.autoFilter, isRow, at, count);
    if (af) sheet.autoFilter = af; else delete sheet.autoFilter;
  }
  if (sheet.freeze && typeof sheet.freeze === "object") {
    const key = isRow ? "rows" : "cols";
    const frozen = sheet.freeze[key] || 0;
    if (frozen > 0 && at < frozen) {
      // rows/cols inserted or deleted inside the frozen area grow/shrink it
      sheet.freeze[key] = count > 0 ? Math.min(frozen + count, isRow ? MAX_ROWS - 1 : MAX_COLS - 1)
        : frozen - (Math.min(frozen, at - count) - at);
      if (!(sheet.freeze.rows > 0) && !(sheet.freeze.cols > 0)) delete sheet.freeze;
    }
  }
  const target = String(sheet.name).toLowerCase();
  for (const s of book.sheets) {
    if (!Array.isArray(s.charts) || !s.charts.length) continue;
    s.charts = s.charts.filter((ch) => {
      if (!ch || typeof ch.range !== "string") return true;
      const bang = ch.range.lastIndexOf("!");
      const prefix = bang >= 0 ? ch.range.slice(0, bang) : null;
      const onSheet = prefix == null ? s === sheet : prefix.replace(/^'(.*)'$/, "$1").replace(/''/g, "'").toLowerCase() === target;
      if (!onSheet) return true;
      const nr = shiftRangeText(ch.range, isRow, at, count);
      if (!nr) return false; // its data was deleted: drop the chart
      ch.range = nr;
      return true;
    });
  }
  for (const s of book.sheets) {
    for (const key of Object.keys(s.cells || {})) {
      const cell = s.cells[key];
      if (cell && cell.f) cell.f = adjustFormulaForShift(cell.f, s.name, sheet.name, axis, at, count);
    }
  }
}

/**
 * Shift an A1-style range text ("A1:D20", "$A$1:$C$10", "Sheet1!A1:C5", "B3") for a
 * row/column insert/delete. Returns the new text, the input unchanged when it
 * cannot be parsed, or null when the whole range was deleted.
 */
function shiftRangeText(text, isRow, at, count) {
  const bang = text.lastIndexOf("!");
  const prefix = bang >= 0 ? text.slice(0, bang + 1) : "";
  const m = /^\s*(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?::(\$?)([A-Za-z]{1,3})(\$?)(\d+))?\s*$/.exec(text.slice(prefix.length));
  if (!m) return text;
  let r1 = +m[4] - 1, c1 = colToIndex(m[2]);
  let r2 = m[6] ? +m[8] - 1 : r1, c2 = m[6] ? colToIndex(m[6]) : c1;
  const max = isRow ? MAX_ROWS : MAX_COLS;
  const range = isRow ? shiftRange(Math.min(r1, r2), Math.max(r1, r2), at, count)
    : shiftRange(Math.min(c1, c2), Math.max(c1, c2), at, count);
  if (!range || range[0] >= max) return null;
  range[1] = Math.min(range[1], max - 1);
  if (isRow) [r1, r2] = range; else [c1, c2] = range;
  const end = (ac, c, ar, r) => `${ac}${indexToCol(c).toUpperCase()}${ar}${r + 1}`;
  return prefix + end(m[1], c1, m[3], r1) + (m[6] ? ":" + end(m[5], c2, m[7], r2) : "");
}

function shiftRange(lo, hi, at, count) {
  if (count > 0) return [lo >= at ? lo + count : lo, hi >= at ? hi + count : hi];
  const n = -count, end = at + n;
  if (lo >= at && hi < end) return null;
  const nlo = lo < at ? lo : lo >= end ? lo - n : at;
  const nhi = hi < at ? hi : hi >= end ? hi - n : at - 1;
  return nhi < nlo ? null : [nlo, nhi];
}

/** Rename a sheet and update every formula that refers to it. */
export function renameSheet(book, si, newName) {
  const sheet = book.sheets[si];
  const old = sheet.name;
  sheet.name = newName;
  for (const s of book.sheets) {
    for (const cell of Object.values(s.cells || {})) {
      if (cell && cell.f) cell.f = renameSheetInFormula(cell.f, old, newName);
    }
  }
}

/** Remove a sheet; formulas pointing at it become #REF!. Keeps `active` on the same sheet when possible. */
export function deleteSheet(book, si) {
  const sheet = book.sheets[si];
  if (!sheet) return;
  book.sheets.splice(si, 1);
  for (const s of book.sheets) {
    for (const cell of Object.values(s.cells || {})) {
      if (cell && cell.f) cell.f = deleteSheetInFormula(cell.f, sheet.name);
    }
  }
  const active = book.active || 0;
  if (si < active) book.active = active - 1;
  else if (active >= book.sheets.length) book.active = Math.max(0, book.sheets.length - 1);
}
