/**
 * Markdown -> Word model (the frozen WORD CONTRACT, see
 * src/office-backend/src/model.rs). The reverse direction is
 * model-word.js wordModelToMarkdown(); this parser reads everything that
 * writer produces, plus common CommonMark/GFM constructs:
 *
 *   blocks:  ATX + setext headings, paragraphs (soft/hard line breaks),
 *            > quotes, fenced and indented code, - * + and 1. 1) lists with
 *            nesting by indentation, GFM tables (cells may hold <br>),
 *            thematic breaks (--- -> page break), standalone images
 *   inline:  **bold** __bold__ *italic* _italic_ ***both***, ~~strike~~,
 *            `code` (-> Consolas run), [text](url "title"), <autolinks>,
 *            bare http(s) URLs, <u> <sup> <sub> <b> <i> <em> <strong> <s>
 *            <del> <br>, backslash escapes and HTML entities
 *
 * Pure module (no DOM), so it runs in the viewer and under Node tests.
 */

const CODE_FONT = "Consolas";
const MAX_LEVEL = 8;

// ---------------------------------------------------------------------------
// Inline parsing
// ---------------------------------------------------------------------------

const PUNCT = /[!-/:-@[-`{-~ -⁯⸀-⹿]/;
const isWs = (c) => c === undefined || c === "" || /\s/.test(c);
const isPunct = (c) => !!c && PUNCT.test(c);

const ENTITIES = { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'", nbsp: " ", copy: "©", reg: "®", trade: "™", hellip: "…", mdash: "—", ndash: "–", laquo: "«", raquo: "»", euro: "€" };

function decodeEntity(m) {
  const body = m.slice(1, -1);
  if (body[0] === "#") {
    const n = body[1] === "x" || body[1] === "X" ? parseInt(body.slice(2), 16) : parseInt(body.slice(1), 10);
    if (n > 0 && n <= 0x10ffff) { try { return String.fromCodePoint(n); } catch { /* fall through */ } }
    return m;
  }
  return Object.prototype.hasOwnProperty.call(ENTITIES, body) ? ENTITIES[body] : m;
}

const SAFE_LINK = /^(https?:|mailto:|tel:|ftp:|#|\/|\.{0,2}\/|[^:]*$)/i;

/** Find the index of the ")" closing a link destination starting at `i` (after "("). */
function scanLinkTail(s, i) {
  // ( <dest> "title" ) with balanced parentheses in dest
  let j = i;
  while (s[j] === " ") j++;
  let dest = "";
  if (s[j] === "<") {
    const k = s.indexOf(">", j + 1);
    if (k < 0) return null;
    dest = s.slice(j + 1, k);
    j = k + 1;
  } else {
    let depth = 0;
    const start = j;
    while (j < s.length) {
      const c = s[j];
      if (c === "\\" && j + 1 < s.length) { j += 2; continue; }
      if (c === "(") depth++;
      else if (c === ")") { if (depth === 0) break; depth--; }
      else if (c === " " || c === "\n") break;
      j++;
    }
    dest = s.slice(start, j);
  }
  while (s[j] === " " || s[j] === "\n") j++;
  if (s[j] === '"' || s[j] === "'" || s[j] === "(") {
    const close = s[j] === "(" ? ")" : s[j];
    const k = s.indexOf(close, j + 1);
    if (k < 0) return null;
    j = k + 1;
    while (s[j] === " ") j++;
  }
  if (s[j] !== ")") return null;
  return { dest: dest.replace(/\\([!-/:-@[-`{-~])/g, "$1"), end: j + 1 };
}

/** Index of the "]" matching the "[" at `i`, skipping escapes and code spans. */
function scanBracket(s, i) {
  let depth = 0;
  for (let j = i; j < s.length; j++) {
    const c = s[j];
    if (c === "\\") { j++; continue; }
    if (c === "`") {
      let n = 1;
      while (s[j + n] === "`") n++;
      const close = s.indexOf("`".repeat(n), j + n);
      if (close > 0) { j = close + n - 1; continue; }
    }
    if (c === "[") depth++;
    else if (c === "]") { depth--; if (depth === 0) return j; }
  }
  return -1;
}

const INLINE_TAGS = { u: "u", ins: "u", sup: "sup", sub: "sub", b: "b", strong: "b", i: "i", em: "i", s: "s", del: "s", strike: "s" };

/**
 * Parse inline markdown into runs. `base` marks apply to every run.
 * Algorithm: tokenize into text items and delimiter items, match emphasis
 * delimiters CommonMark-style, and apply each match as a mark on every item
 * between the pair (the output is a flat run list, so no tree is needed).
 */
export function parseInline(src, base = {}) {
  const s = String(src || "");
  const items = [];
  const text = (t, marks = {}) => {
    if (!t) return;
    const last = items[items.length - 1];
    if (last && last.kind === "text" && !last.atomic && !Object.keys(marks).length) last.t += t;
    else items.push({ kind: "text", t, marks: { ...marks } });
  };
  const atomic = (runs) => {
    for (const r of runs) items.push({ kind: "text", t: r.t, marks: { ...r, t: undefined }, atomic: true });
  };
  let i = 0;
  while (i < s.length) {
    const c = s[i];
    // backslash escape / hard break
    if (c === "\\") {
      const n = s[i + 1];
      if (n === "\n") { text("\n"); i += 2; continue; }
      if (n && /[!-/:-@[-`{-~]/.test(n)) { text(n); i += 2; continue; }
      text("\\"); i++; continue;
    }
    // line breaks: 2+ trailing spaces = hard break, else soft break (space)
    if (c === "\n") {
      const last = items[items.length - 1];
      if (last && last.kind === "text" && / {2,}$/.test(last.t)) {
        last.t = last.t.replace(/ +$/, "");
        text("\n");
      } else {
        if (last && last.kind === "text") last.t = last.t.replace(/ +$/, "");
        text(" ");
      }
      i++;
      while (s[i] === " ") i++;
      continue;
    }
    // code span
    if (c === "`") {
      let n = 1;
      while (s[i + n] === "`") n++;
      const fence = "`".repeat(n);
      let k = i + n;
      let close = -1;
      while ((k = s.indexOf(fence, k)) >= 0) {
        if (s[k + n] !== "`" && s[k - 1] !== "`") { close = k; break; }
        k += n;
      }
      if (close < 0) { text(fence); i += n; continue; }
      let code = s.slice(i + n, close).replace(/\n/g, " ");
      if (/^ .*[^ ].* $/.test(code)) code = code.slice(1, -1);
      atomic([{ t: code, font: CODE_FONT }]);
      i = close + n;
      continue;
    }
    // images / links
    if ((c === "!" && s[i + 1] === "[") || c === "[") {
      const open = c === "!" ? i + 1 : i;
      const close = scanBracket(s, open);
      if (close > 0 && s[close + 1] === "(") {
        const tail = scanLinkTail(s, close + 2);
        if (tail) {
          const label = s.slice(open + 1, close);
          if (c === "!") {
            // Inline images cannot live inside a run: keep a link to them.
            const alt = label || "image";
            if (/^data:/i.test(tail.dest)) atomic([{ t: `[${alt}]` }]);
            else atomic([{ t: alt, link: tail.dest }]);
          } else {
            const inner = parseInline(label);
            const link = SAFE_LINK.test(tail.dest) ? tail.dest : null;
            atomic(inner.map((r) => (link && !r.link ? { ...r, link } : r)));
          }
          i = tail.end;
          continue;
        }
      }
      text(c); i++; continue;
    }
    // autolinks and inline HTML
    if (c === "<") {
      const auto = /^<((?:https?|ftp):\/\/[^\s<>]+|mailto:[^\s<>]+)>/i.exec(s.slice(i));
      if (auto) { atomic([{ t: auto[1].replace(/^mailto:/i, ""), link: auto[1] }]); i += auto[0].length; continue; }
      const email = /^<([^\s<>@]+@[^\s<>@]+\.[^\s<>@]+)>/.exec(s.slice(i));
      if (email) { atomic([{ t: email[1], link: `mailto:${email[1]}` }]); i += email[0].length; continue; }
      const tag = /^<(\/?)([a-z]+)\s*\/?>/i.exec(s.slice(i));
      if (tag) {
        const name = tag[2].toLowerCase();
        if (name === "br") { text("\n"); i += tag[0].length; continue; }
        if (INLINE_TAGS[name]) {
          items.push({ kind: "tag", mark: INLINE_TAGS[name], open: !tag[1], marks: {} });
          i += tag[0].length;
          continue;
        }
        if (name === "span" || name === "font" || name === "mark") { i += tag[0].length; continue; }
      }
      text("<"); i++; continue;
    }
    // entities
    if (c === "&") {
      const ent = /^&(#\d{1,7}|#x[0-9a-f]{1,6}|[a-z]{2,8});/i.exec(s.slice(i));
      if (ent) { text(decodeEntity(ent[0])); i += ent[0].length; continue; }
      text("&"); i++; continue;
    }
    // bare URLs (GFM autolink extension)
    if ((c === "h" || c === "w") && (i === 0 || /[\s(]/.test(s[i - 1]))) {
      const m = /^(https?:\/\/|www\.)[^\s<]*[^\s<?!.,:*_~')"]/i.exec(s.slice(i));
      if (m) {
        const url = m[0].startsWith("www.") ? `http://${m[0]}` : m[0];
        atomic([{ t: m[0], link: url }]);
        i += m[0].length;
        continue;
      }
    }
    // emphasis delimiter runs
    if (c === "*" || c === "_" || c === "~") {
      let n = 1;
      while (s[i + n] === c) n++;
      const prev = i > 0 ? s[i - 1] : undefined;
      const next = s[i + n];
      const left = !isWs(next) && (!isPunct(next) || isWs(prev) || isPunct(prev));
      const right = !isWs(prev) && (!isPunct(prev) || isWs(next) || isPunct(next));
      let canOpen = left, canClose = right;
      if (c === "_") {
        canOpen = left && (!right || isPunct(prev));
        canClose = right && (!left || isPunct(next));
      }
      if (c === "~" && n < 2) { text("~"); i++; continue; }
      items.push({ kind: "delim", ch: c, count: n, orig: n, canOpen, canClose, marks: {}, active: true });
      i += n;
      continue;
    }
    // plain text up to the next special character
    let j = i + 1;
    while (j < s.length && !/[\\\n`![<&*_~hw]/.test(s[j])) j++;
    text(s.slice(i, j));
    i = j;
  }

  // --- emphasis matching ---
  const markBetween = (a, b, mark) => {
    for (let k = a + 1; k < b; k++) items[k].marks[mark] = true;
  };
  for (let ci = 0; ci < items.length; ci++) {
    const cl = items[ci];
    if (cl.kind !== "delim" || !cl.canClose) continue;
    while (cl.count > 0) {
      let found = -1;
      for (let oi = ci - 1; oi >= 0; oi--) {
        const op = items[oi];
        if (op.kind !== "delim" || !op.active || op.ch !== cl.ch || !op.canOpen || op.count === 0) continue;
        if (cl.ch !== "~" && (op.canClose || cl.canOpen) && (op.orig + cl.orig) % 3 === 0 && !(op.orig % 3 === 0 && cl.orig % 3 === 0)) continue;
        if (cl.ch === "~" && (op.count < 2 || cl.count < 2)) continue;
        found = oi;
        break;
      }
      if (found < 0) break;
      const op = items[found];
      const use = cl.ch === "~" ? 2 : op.count >= 2 && cl.count >= 2 ? 2 : 1;
      markBetween(found, ci, cl.ch === "~" ? "s" : use === 2 ? "b" : "i");
      // delimiters strictly inside can no longer match anything outside
      for (let k = found + 1; k < ci; k++) if (items[k].kind === "delim") items[k].active = false;
      op.count -= use;
      cl.count -= use;
    }
  }
  // --- inline HTML tag pairs ---
  for (let ci = 0; ci < items.length; ci++) {
    const cl = items[ci];
    if (cl.kind !== "tag" || cl.open) continue;
    for (let oi = ci - 1; oi >= 0; oi--) {
      const op = items[oi];
      if (op.kind === "tag" && op.open && !op.used && op.mark === cl.mark) {
        op.used = true;
        markBetween(oi, ci, cl.mark);
        break;
      }
    }
  }

  // --- flatten to runs ---
  const runs = [];
  for (const it of items) {
    let t = "";
    if (it.kind === "text") t = it.t;
    else if (it.kind === "delim") t = it.ch.repeat(it.count);
    else continue; // tags produce no text
    if (!t) continue;
    const r = { t };
    const marks = { ...base, ...it.marks };
    for (const k of ["b", "i", "u", "s", "sup", "sub"]) if (marks[k]) r[k] = true;
    if (marks.font) r.font = marks.font;
    if (marks.link) r.link = marks.link;
    if (marks.color) r.color = marks.color;
    if (marks.hl) r.hl = marks.hl;
    if (marks.size) r.size = marks.size;
    const last = runs[runs.length - 1];
    if (last && sameMarks(last, r)) last.t += t;
    else runs.push(r);
  }
  return runs;
}

function sameMarks(a, b) {
  for (const k of ["b", "i", "u", "s", "sup", "sub", "font", "link", "color", "hl", "size"]) {
    if ((a[k] || undefined) !== (b[k] || undefined)) return false;
  }
  return true;
}

// ---------------------------------------------------------------------------
// Block parsing
// ---------------------------------------------------------------------------

const RE_FENCE = /^ {0,3}(`{3,}|~{3,})[^`]*$/;
const RE_ATX = /^ {0,3}(#{1,6})(?:[ \t]+(.*?))?(?:[ \t]+#+)?[ \t]*$/;
const RE_HR = /^ {0,3}([-*_])(?:[ \t]*\1){2,}[ \t]*$/;
const RE_SETEXT1 = /^ {0,3}=+[ \t]*$/;
const RE_SETEXT2 = /^ {0,3}-+[ \t]*$/;
const RE_QUOTE = /^ {0,3}> ?(.*)$/;
const RE_LIST = /^( *)([-*+]|\d{1,9}[.)])(?:[ \t]+(.*)|[ \t]*)$/;
const RE_TABLE_DELIM = /^ *\|? *:?-+:? *(\| *:?-+:? *)*\|? *$/;
const RE_IMG_LINE = /^ {0,3}!\[((?:\\.|[^\]\\])*)\]\(\s*<?([^\s>)]+)>?(?:\s+"[^"]*")?\s*\)\s*$/;

const expandTabs = (line) => line.replace(/^\t+/, (m) => "    ".repeat(m.length));

function splitTableRow(line) {
  let s = line.trim();
  if (s.startsWith("|")) s = s.slice(1);
  if (s.endsWith("|") && !s.endsWith("\\|")) s = s.slice(0, -1);
  const cells = [];
  let cur = "";
  let inCode = false;
  for (let k = 0; k < s.length; k++) {
    const ch = s[k];
    if (ch === "\\" && s[k + 1] === "|") { cur += "|"; k++; continue; }
    if (ch === "`") inCode = !inCode;
    if (ch === "|" && !inCode) { cells.push(cur.trim()); cur = ""; continue; }
    cur += ch;
  }
  cells.push(cur.trim());
  return cells;
}

function cellBlocks(text) {
  const blocks = [];
  for (const part of String(text).split(/<br\s*\/?>/i)) {
    const img = RE_IMG_LINE.exec(part.trim());
    if (img && /^data:/i.test(img[2])) {
      blocks.push(imgBlock(img[1], img[2]));
      continue;
    }
    blocks.push({ type: "p", runs: parseInline(part.trim()) });
  }
  return blocks.length ? blocks : [{ type: "p", runs: [] }];
}

function imgBlock(alt, src) {
  const b = { type: "img", src };
  const a = String(alt || "").replace(/\\(.)/g, "$1");
  if (a) b.alt = a;
  return b;
}

/**
 * Parse a markdown string into a word model { kind:"word", blocks }.
 */
export function markdownToModel(markdown) {
  const text = String(markdown || "").replace(/^﻿/, "").replace(/\r\n?/g, "\n");
  const lines = text.split("\n").map(expandTabs);
  const blocks = [];

  let para = null;          // { lines: [] }
  let list = null;          // { stack: [{indent, content}], item: {block, lines} }
  let quote = null;         // { lines: [] }

  const flushPara = () => {
    if (!para) return;
    const src = para.lines.join("\n");
    para = null;
    const img = RE_IMG_LINE.exec(src);
    if (img && !src.includes("\n")) {
      if (/^data:/i.test(img[2])) { blocks.push(imgBlock(img[1], img[2])); return; }
    }
    const runs = parseInline(src);
    if (runs.length) blocks.push({ type: "p", runs });
  };
  const flushItem = () => {
    if (!list || !list.item) return;
    list.item.block.runs = parseInline(list.item.lines.join("\n"));
    list.item = null;
  };
  const endList = () => { flushItem(); list = null; };
  const flushQuote = () => {
    if (!quote) return;
    // blank "> " lines separate quote paragraphs
    const paras = quote.lines.join("\n").split(/\n[ \t]*\n/);
    for (const p of paras) {
      const runs = parseInline(p.replace(/^ {0,3}#{1,6}[ \t]+/, ""));
      if (runs.length) blocks.push({ type: "quote", runs });
    }
    quote = null;
  };
  const flushAll = () => { flushPara(); endList(); flushQuote(); };

  let blankBefore = true;
  for (let li = 0; li < lines.length; li++) {
    const line = lines[li];
    const blank = /^\s*$/.test(line);

    // fenced code
    const fence = RE_FENCE.exec(line);
    if (fence && !(fence[1][0] === "`" && /`/.test(line.slice(line.indexOf(fence[1]) + fence[1].length)))) {
      flushAll();
      const indent = line.length - line.trimStart().length;
      const marker = fence[1];
      const body = [];
      let k = li + 1;
      for (; k < lines.length; k++) {
        const l = lines[k];
        const m = /^ {0,3}(`{3,}|~{3,})[ \t]*$/.exec(l);
        if (m && m[1][0] === marker[0] && m[1].length >= marker.length) break;
        body.push(l.replace(new RegExp(`^ {0,${indent}}`), ""));
      }
      blocks.push({ type: "code", runs: body.length ? [{ t: body.join("\n") }] : [] });
      li = k; // skip closing fence (or run to the end)
      blankBefore = false;
      continue;
    }

    if (blank) {
      flushPara();
      flushQuote();
      if (list) list.blank = true;
      blankBefore = true;
      continue;
    }

    // list items (checked before indented code so nested items work)
    const lm = RE_LIST.exec(line);
    const isHr = RE_HR.test(line);
    // Only "1." ordered items and non-empty items may interrupt a paragraph.
    const interruptsPara = !para || (!!lm && !!lm[3] && (!/^\d/.test(lm[2]) || lm[2] === "1." || lm[2] === "1)"));
    if (lm && !isHr && (list || lm[1].length < 4) && interruptsPara) {
      const indent = lm[1].length;
      const content = indent + lm[2].length + 1;
      if (!list) { flushPara(); flushQuote(); list = { stack: [], item: null, blank: false }; }
      // An item indented less than its predecessor's content is a sibling or
      // belongs to an outer list; deeper than the content column is a child.
      flushItem();
      const st = list.stack;
      while (st.length && indent < st[st.length - 1].indent) st.pop();
      if (st.length && indent >= st[st.length - 1].content) st.push({ indent, content });
      else if (st.length) st[st.length - 1] = { indent, content };
      else st.push({ indent, content });
      const level = Math.min(MAX_LEVEL, st.length - 1);
      const block = { type: "list", list: /\d/.test(lm[2]) ? "ol" : "ul", level, runs: [] };
      blocks.push(block);
      let first = lm[3] || "";
      // GFM task list markers stay visible as ballot boxes.
      first = first.replace(/^\[ \]\s+/, "☐ ").replace(/^\[[xX]\]\s+/, "☑ ");
      list.item = { block, lines: [first] };
      list.blank = false;
      blankBefore = false;
      continue;
    }
    if (list) {
      const indent = line.length - line.trimStart().length;
      const top = list.stack[list.stack.length - 1];
      if (list.item && (!list.blank || indent >= top.content) && !RE_ATX.test(line) && !isHr && !RE_QUOTE.test(line)) {
        // continuation (lazy or indented) of the current item
        // after a blank line the continuation is a new paragraph of the
        // item: keep it as a hard line break inside the list block
        if (list.blank) list.item.lines[list.item.lines.length - 1] += "\\";
        list.item.lines.push(line.trim());
        list.blank = false;
        continue;
      }
      endList();
    }

    // indented code block
    if (!para && blankBefore && /^ {4}/.test(line) && !quote) {
      flushAll();
      const body = [];
      let k = li;
      for (; k < lines.length; k++) {
        const l = lines[k];
        if (/^ {4}/.test(l)) body.push(l.slice(4));
        else if (/^\s*$/.test(l)) body.push("");
        else break;
      }
      while (body.length && !body[body.length - 1]) body.pop();
      blocks.push({ type: "code", runs: [{ t: body.join("\n") }] });
      li = k - 1;
      blankBefore = false;
      continue;
    }

    // quotes
    const qm = RE_QUOTE.exec(line);
    if (qm) {
      flushPara();
      if (!quote) quote = { lines: [] };
      quote.lines.push(qm[1]);
      blankBefore = false;
      continue;
    }
    if (quote) {
      // lazy continuation of a quote paragraph
      if (!RE_ATX.test(line) && !isHr) { quote.lines.push(line.trim()); continue; }
      flushQuote();
    }

    // setext headings (underline after a paragraph)
    if (para && (RE_SETEXT1.test(line) || (RE_SETEXT2.test(line) && !para.table))) {
      const src = para.lines.join("\n");
      para = null;
      blocks.push({ type: RE_SETEXT1.test(line) ? "h1" : "h2", runs: parseInline(src) });
      blankBefore = false;
      continue;
    }

    // thematic break -> page break
    if (isHr) {
      flushPara();
      blocks.push({ type: "pagebreak" });
      blankBefore = false;
      continue;
    }

    // ATX headings
    const hm = RE_ATX.exec(line);
    if (hm) {
      flushPara();
      blocks.push({ type: `h${hm[1].length}`, runs: parseInline(hm[2] || "") });
      blankBefore = false;
      continue;
    }

    // GFM tables: header row + delimiter row
    if (!para && line.includes("|") && li + 1 < lines.length && RE_TABLE_DELIM.test(lines[li + 1]) && lines[li + 1].includes("-")) {
      const header = splitTableRow(line);
      const delims = splitTableRow(lines[li + 1]);
      if (delims.length === header.length || line.trim().startsWith("|")) {
        const rows = [header];
        let k = li + 2;
        for (; k < lines.length; k++) {
          const l = lines[k];
          if (/^\s*$/.test(l) || !l.includes("|")) break;
          rows.push(splitTableRow(l));
        }
        const width = Math.max(...rows.map((r) => r.length));
        blocks.push({
          type: "table",
          rows: rows.map((r) => {
            const cells = r.map((c) => ({ blocks: cellBlocks(c) }));
            while (cells.length < width) cells.push({ blocks: [{ type: "p", runs: [] }] });
            return cells;
          }),
        });
        li = k - 1;
        blankBefore = false;
        continue;
      }
    }

    // paragraph text
    if (!para) para = { lines: [] };
    para.lines.push(para.lines.length ? line.replace(/^ +/, "") : line.replace(/^ {0,3}/, ""));
    blankBefore = false;
  }
  flushAll();
  return { kind: "word", blocks: blocks.length ? blocks : [{ type: "p", runs: [] }] };
}
