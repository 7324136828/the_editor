/**
 * Office format detection (filename extension, MIME type, magic bytes and
 * ZIP/OLE container peeking), text encoding helpers, and the static save
 * capability table used by the viewer's Save / Save As logic.
 *
 * Shared by popup, content script, service worker and viewer. Plain ES module
 * with no DOM dependency, so everything here is Node-testable.
 */

export const FORMATS = {
  docx: { kind: "word", label: "Word Document (.docx)", mime: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" },
  doc: { kind: "word", label: "Word 97-2003 Document (.doc)", mime: "application/msword" },
  rtf: { kind: "word", label: "Rich Text Format (.rtf)", mime: "application/rtf" },
  odt: { kind: "word", label: "OpenDocument Text (.odt)", mime: "application/vnd.oasis.opendocument.text" },
  txt: { kind: "word", label: "Plain Text (.txt)", mime: "text/plain" },
  md: { kind: "word", label: "Markdown (.md)", mime: "text/markdown" },
  xlsx: { kind: "excel", label: "Excel Workbook (.xlsx)", mime: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
  xls: { kind: "excel", label: "Excel 97-2003 Workbook (.xls)", mime: "application/vnd.ms-excel" },
  xlsm: { kind: "excel", label: "Excel Macro-Enabled Workbook (.xlsm)", mime: "application/vnd.ms-excel.sheet.macroEnabled.12" },
  csv: { kind: "excel", label: "CSV (.csv)", mime: "text/csv" },
  tsv: { kind: "excel", label: "TSV (.tsv)", mime: "text/tab-separated-values" },
  ods: { kind: "excel", label: "OpenDocument Spreadsheet (.ods)", mime: "application/vnd.oasis.opendocument.spreadsheet" },
  pptx: { kind: "slides", label: "PowerPoint Presentation (.pptx)", mime: "application/vnd.openxmlformats-officedocument.presentationml.presentation" },
  tex: { kind: "latex", label: "LaTeX Document (.tex)", mime: "application/x-tex" },
};

/** Additional extensions that map onto a supported format. */
const EXT_ALIASES = { markdown: "md", mdown: "md", mkd: "md", ltx: "tex", latex: "tex" };

/** MIME types (lower-case, parameters stripped) seen in the wild -> format. */
const MIME_ALIASES = {
  "application/vnd.ms-word": "doc",
  "application/x-msword": "doc",
  "application/vnd.ms-word.document.macroenabled.12": "docx",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.template": "docx",
  "text/rtf": "rtf",
  "application/x-rtf": "rtf",
  "text/richtext": "rtf",
  "text/x-markdown": "md",
  "text/x-web-markdown": "md",
  "application/x-msexcel": "xls",
  "application/x-excel": "xls",
  "application/excel": "xls",
  "application/vnd.ms-excel.sheet.binary.macroenabled.12": null, // .xlsb: not supported
  "application/csv": "csv",
  "text/comma-separated-values": "csv",
  "text/x-csv": "csv",
  "application/x-csv": "csv",
  "text/tsv": "tsv",
  "application/vnd.ms-powerpoint.presentation.macroenabled.12": "pptx",
  "application/vnd.openxmlformats-officedocument.presentationml.slideshow": "pptx",
  "text/x-tex": "tex",
  "application/x-latex": "tex",
  "text/x-latex": "tex",
};

for (const [fmt, meta] of Object.entries(FORMATS)) MIME_ALIASES[meta.mime.toLowerCase()] = fmt;

export const WORD_FORMATS = ["docx", "doc", "rtf", "odt", "md"];
export const EXCEL_FORMATS = ["xlsx", "xls", "xlsm", "csv", "tsv", "ods"];
export const TEXT_FORMATS = ["txt"];
export const SLIDES_FORMATS = ["pptx"];
export const LATEX_FORMATS = ["tex"];

/**
 * What the viewer turns a format into: "word" (rich editor), "excel" (grid)
 * or "text" (plain-text editor). Markdown opens as a formatted Word document.
 */
export function editorKindOf(format) {
  if (TEXT_FORMATS.includes(format)) return "text";
  if (SLIDES_FORMATS.includes(format)) return "slides";
  if (LATEX_FORMATS.includes(format)) return "latex";
  if (EXCEL_FORMATS.includes(format)) return "excel";
  if (WORD_FORMATS.includes(format)) return "word";
  return null;
}

export function extOf(name) {
  const m = /\.([A-Za-z0-9]{1,10})$/.exec(String(name || "").trim());
  return m ? m[1].toLowerCase() : "";
}

export function formatOfName(name) {
  const ext = extOf(name);
  if (!ext) return null;
  if (FORMATS[ext]) return ext;
  return EXT_ALIASES[ext] || null;
}

export function isOfficeName(name) {
  return formatOfName(name) !== null;
}

/** "Text/CSV; charset=utf-8" -> "text/csv". */
export function mimeEssence(mime) {
  return String(mime || "").split(";")[0].trim().toLowerCase();
}

/** Format for a MIME type (parameters ignored), or null. */
export function formatOfMime(mime) {
  const m = mimeEssence(mime);
  if (!m) return null;
  return Object.prototype.hasOwnProperty.call(MIME_ALIASES, m) ? MIME_ALIASES[m] : null;
}

function toU8(bytes) {
  if (!bytes) return null;
  if (bytes instanceof Uint8Array) return bytes;
  if (bytes instanceof ArrayBuffer) return new Uint8Array(bytes);
  if (ArrayBuffer.isView(bytes)) return new Uint8Array(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  return null;
}

/** Detect the container from raw bytes (magic numbers). */
export function sniffFormat(bytes) {
  const u8 = toU8(bytes);
  if (!u8 || u8.length < 4) return null;
  // OLE2 compound document -> .doc or .xls (see sniffOleKind)
  if (u8[0] === 0xd0 && u8[1] === 0xcf && u8[2] === 0x11 && u8[3] === 0xe0) return "ole2";
  // ZIP container -> docx/xlsx/odt/ods (see sniffZipKind)
  if (u8[0] === 0x50 && u8[1] === 0x4b && (u8[2] === 3 || u8[2] === 5 || u8[2] === 7)) return "zip";
  // RTF ("{\rtf")
  if (u8[0] === 0x7b && u8[1] === 0x5c && u8[2] === 0x72 && u8[3] === 0x74) return "rtf";
  return null;
}

function latin1(u8, start, end) {
  let s = "";
  for (let i = start; i < end; i += 0x4000) {
    s += String.fromCharCode.apply(null, u8.subarray(i, Math.min(end, i + 0x4000)));
  }
  return s;
}

/**
 * Tell docx / xlsx / xlsm / odt / ods apart by peeking at the ZIP local file
 * headers in the first KBs (entry names are stored uncompressed there; the
 * OpenDocument "mimetype" entry is stored uncompressed as the first entry).
 * No JSZip needed, so this works in every context.
 */
export function sniffZipKind(bytes, maxScan = 256 * 1024) {
  const u8 = toU8(bytes);
  if (!u8 || sniffFormat(u8) !== "zip") return null;
  const end = Math.min(u8.length, maxScan);
  const names = [];
  // Walk local headers as long as sizes are known up front.
  let p = 0;
  while (p + 30 <= end && u8[p] === 0x50 && u8[p + 1] === 0x4b && u8[p + 2] === 3 && u8[p + 3] === 4) {
    const flags = u8[p + 6] | (u8[p + 7] << 8);
    const csize = (u8[p + 18] | (u8[p + 19] << 8) | (u8[p + 20] << 16) | (u8[p + 21] << 24)) >>> 0;
    const nlen = u8[p + 26] | (u8[p + 27] << 8);
    const xlen = u8[p + 28] | (u8[p + 29] << 8);
    if (p + 30 + nlen > end) break;
    const name = latin1(u8, p + 30, p + 30 + nlen);
    names.push(name);
    if (name === "mimetype") {
      const dataStart = p + 30 + nlen + xlen;
      const mt = latin1(u8, dataStart, Math.min(end, dataStart + Math.min(csize || 80, 120))).trim();
      if (mt.startsWith("application/vnd.oasis.opendocument.text")) return "odt";
      if (mt.startsWith("application/vnd.oasis.opendocument.spreadsheet")) return "ods";
    }
    if (flags & 0x08) break; // sizes in a data descriptor: cannot skip reliably
    p += 30 + nlen + xlen + csize;
  }
  const fromNames = (list) => {
    if (list.some((n) => n.startsWith("word/"))) return "docx";
    if (list.some((n) => n === "xl/vbaProject.bin")) return "xlsm";
    if (list.some((n) => n.startsWith("xl/"))) return "xlsx";
    if (list.some((n) => n.startsWith("ppt/"))) return "pptx";
    return null;
  };
  const byWalk = fromNames(names);
  if (byWalk) return byWalk;
  // Fallback: plain string search of the scanned region (covers streamed
  // ZIPs with data descriptors, where the walk above stops early).
  const s = latin1(u8, 0, end);
  if (/mimetypeapplication\/vnd\.oasis\.opendocument\.text/.test(s)) return "odt";
  if (/mimetypeapplication\/vnd\.oasis\.opendocument\.spreadsheet/.test(s)) return "ods";
  if (s.includes("word/document.xml") || s.includes("word/_rels") || s.includes("word/styles.xml")) return "docx";
  if (s.includes("xl/vbaProject.bin")) return "xlsm";
  if (s.includes("xl/workbook.xml") || s.includes("xl/worksheets/") || s.includes("xl/_rels")) return "xlsx";
  if (s.includes("ppt/presentation.xml") || s.includes("ppt/slides/") || s.includes("ppt/_rels")) return "pptx";
  if (/\x50\x4b\x03\x04[\s\S]{26}word\//.test(s)) return "docx";
  if (/\x50\x4b\x03\x04[\s\S]{26}xl\//.test(s)) return "xlsx";
  if (/\x50\x4b\x03\x04[\s\S]{26}ppt\//.test(s)) return "pptx";
  return null;
}

/**
 * Tell a Word .doc from an Excel .xls OLE2 file by looking for the stream
 * names ("WordDocument" vs "Workbook"/"Book") in UTF-16LE in the directory.
 */
export function sniffOleKind(bytes, maxScan = 8 * 1024 * 1024) {
  const u8 = toU8(bytes);
  if (!u8 || sniffFormat(u8) !== "ole2") return null;
  const end = Math.min(u8.length, maxScan);
  const utf16 = (s) => {
    const out = new Uint8Array(s.length * 2);
    for (let i = 0; i < s.length; i++) out[i * 2] = s.charCodeAt(i);
    return out;
  };
  const has = (needle) => {
    const first = needle[0];
    outer: for (let i = 512; i + needle.length <= end; i++) {
      if (u8[i] !== first) continue;
      for (let j = 1; j < needle.length; j++) if (u8[i + j] !== needle[j]) continue outer;
      return true;
    }
    return false;
  };
  if (has(utf16("WordDocument"))) return "doc";
  if (has(utf16("Workbook")) || has(utf16("Book\0"))) return "xls";
  return null;
}

/**
 * Resolve the effective format for a file.
 * Order: filename extension (corrected when the bytes clearly disagree, e.g.
 * an RTF saved as .doc) -> magic bytes / container peeking -> MIME type ->
 * "looks like text" fallback.
 */
export function resolveFormat(name, mime, bytes) {
  const u8 = toU8(bytes);
  const byExt = formatOfName(name);
  const sniffed = sniffFormat(u8);
  if (byExt) {
    // Common mislabels Word/Excel themselves produce.
    if (byExt === "doc" && sniffed === "rtf") return "rtf";
    if ((byExt === "doc" || byExt === "rtf") && sniffed === "zip" && sniffZipKind(u8) === "docx") return "docx";
    if (byExt === "xls" && sniffed === "zip") {
      const z = sniffZipKind(u8);
      if (z === "xlsx" || z === "xlsm") return z;
    }
    return byExt;
  }
  const byMime = formatOfMime(mime);
  if (sniffed === "rtf") return "rtf";
  if (sniffed === "zip") {
    const z = sniffZipKind(u8);
    if (z) return z;
    // Unknown ZIP: trust a ZIP-based MIME type, otherwise give up.
    return byMime && ["docx", "xlsx", "xlsm", "odt", "ods", "pptx"].includes(byMime) ? byMime : null;
  }
  if (sniffed === "ole2") {
    if (byMime === "doc" || byMime === "xls") return byMime;
    return sniffOleKind(u8);
  }
  if (byMime) {
    // A text MIME must actually look like text; binary MIMEs need magic we
    // did not find above, so do not trust them.
    if (["txt", "md", "csv", "tsv", "tex"].includes(byMime)) return !u8 || looksLikeText(u8) ? byMime : null;
    if (byMime === "rtf") return null;
    if (["docx", "xlsx", "xlsm", "odt", "ods", "doc", "xls", "pptx"].includes(byMime)) {
      // xls exported as HTML/TSV by web apps is common; SheetJS reads those.
      return byMime === "xls" && u8 && looksLikeText(u8) ? "xls" : null;
    }
  }
  if (mimeEssence(mime) === "text/html") return null;
  if (u8 && u8.length && looksLikeText(u8)) return "txt";
  return null;
}

export function looksLikeText(bytes, maxScan = 8192) {
  const u8 = toU8(bytes);
  if (!u8 || !u8.length) return false;
  // UTF-16 with BOM is text even though it contains NUL bytes.
  if ((u8[0] === 0xff && u8[1] === 0xfe) || (u8[0] === 0xfe && u8[1] === 0xff)) return true;
  const n = Math.min(u8.length, maxScan);
  let bad = 0;
  for (let i = 0; i < n; i++) {
    const b = u8[i];
    if (b === 0) return false;
    // allow tab, LF, CR, FF and printable ASCII + UTF-8/ANSI high bytes
    if (b < 0x20 && b !== 0x09 && b !== 0x0a && b !== 0x0d && b !== 0x0c) bad++;
    if (bad > n * 0.05) return false;
  }
  return true;
}

/** Replace the extension of `name` with the one for `format`. */
export function withExtension(name, format) {
  const base = String(name || "").trim().replace(/\.[A-Za-z0-9]{1,10}$/, "") || "Untitled";
  return `${base}.${format}`;
}

/**
 * Excel-safe sheet name: no []:*?/\ characters, no leading/trailing
 * apostrophes, at most 31 characters, never empty.
 */
export function sanitizeSheetName(name, fallback = "Sheet1") {
  let s = String(name || "").replace(/[[\]:*?/\\]/g, "_").replace(/[\u0000-\u001f]/g, "");
  s = s.replace(/^'+|'+$/g, "").trim();
  if (s.length > 31) s = s.slice(0, 31).trim();
  return s || fallback;
}

// ---------------------------------------------------------------------------
// Text encodings (plain text / CSV): detect on open, write back the same way.
// ---------------------------------------------------------------------------

// Windows-1252 bytes 0x80..0x9F (undefined positions map to the C1 control).
const CP1252_HIGH = [
  0x20ac, 0x81, 0x201a, 0x0192, 0x201e, 0x2026, 0x2020, 0x2021, 0x02c6, 0x2030, 0x0160, 0x2039, 0x0152, 0x8d, 0x017d, 0x8f,
  0x90, 0x2018, 0x2019, 0x201c, 0x201d, 0x2022, 0x2013, 0x2014, 0x02dc, 0x2122, 0x0161, 0x203a, 0x0153, 0x9d, 0x017e, 0x0178,
];
const CP1252_REVERSE = new Map(CP1252_HIGH.map((cp, i) => [cp, 0x80 + i]));

function decodeCp1252(u8) {
  let s = "";
  for (let i = 0; i < u8.length; i += 0x4000) {
    const chunk = u8.subarray(i, Math.min(u8.length, i + 0x4000));
    const codes = new Array(chunk.length);
    for (let j = 0; j < chunk.length; j++) {
      const b = chunk[j];
      codes[j] = b >= 0x80 && b <= 0x9f ? CP1252_HIGH[b - 0x80] : b;
    }
    s += String.fromCharCode.apply(null, codes);
  }
  return s;
}

function decodeUtf16(u8, littleEndian) {
  let s = "";
  const n = u8.length - (u8.length % 2);
  const codes = [];
  for (let i = 0; i < n; i += 2) {
    codes.push(littleEndian ? u8[i] | (u8[i + 1] << 8) : (u8[i] << 8) | u8[i + 1]);
    if (codes.length >= 0x4000) { s += String.fromCharCode.apply(null, codes); codes.length = 0; }
  }
  return s + String.fromCharCode.apply(null, codes);
}

/** Guess BOM-less UTF-16 (lots of NULs on even or odd positions). */
function guessUtf16(u8) {
  const n = Math.min(u8.length, 4096) & ~1;
  if (n < 4) return null;
  let evenZero = 0, oddZero = 0;
  for (let i = 0; i < n; i += 2) {
    if (u8[i] === 0) evenZero++;
    if (u8[i + 1] === 0) oddZero++;
  }
  const pairs = n / 2;
  if (oddZero > pairs * 0.4 && evenZero < pairs * 0.05) return "utf-16le";
  if (evenZero > pairs * 0.4 && oddZero < pairs * 0.05) return "utf-16be";
  return null;
}

/**
 * Decode text bytes with encoding detection.
 * @returns {{text:string, encoding:"utf-8"|"utf-16le"|"utf-16be"|"windows-1252", bom:boolean, eol:"\n"|"\r\n"}}
 */
export function decodeText(bytes) {
  const u8 = toU8(bytes) || new Uint8Array(0);
  let encoding = "utf-8";
  let bom = false;
  let text;
  if (u8.length >= 3 && u8[0] === 0xef && u8[1] === 0xbb && u8[2] === 0xbf) {
    bom = true;
    text = new TextDecoder("utf-8").decode(u8.subarray(3));
  } else if (u8.length >= 2 && u8[0] === 0xff && u8[1] === 0xfe) {
    bom = true; encoding = "utf-16le";
    text = decodeUtf16(u8.subarray(2), true);
  } else if (u8.length >= 2 && u8[0] === 0xfe && u8[1] === 0xff) {
    bom = true; encoding = "utf-16be";
    text = decodeUtf16(u8.subarray(2), false);
  } else {
    const g = guessUtf16(u8);
    if (g) {
      encoding = g;
      text = decodeUtf16(u8, g === "utf-16le");
    } else {
      try {
        text = new TextDecoder("utf-8", { fatal: true }).decode(u8);
      } catch {
        encoding = "windows-1252";
        text = decodeCp1252(u8);
      }
    }
  }
  const crlf = (text.match(/\r\n/g) || []).length;
  const lf = (text.match(/\n/g) || []).length;
  const eol = crlf > 0 && crlf * 2 >= lf ? "\r\n" : "\n";
  return { text, encoding, bom, eol };
}

/** True when every character of `text` exists in Windows-1252. */
export function fitsCp1252(text) {
  for (const ch of String(text)) {
    const c = ch.codePointAt(0);
    if (c < 0x80 || (c >= 0xa0 && c <= 0xff)) continue;
    if (!CP1252_REVERSE.has(c)) return false;
  }
  return true;
}

/**
 * Encode text the way it was read (decodeText's result). Line endings are
 * normalized to `opts.eol`. Windows-1252 text that gained characters outside
 * that code page is written as UTF-8 with BOM instead (result.encoding says
 * which encoding was actually used).
 * @returns {{bytes: Uint8Array, encoding: string}}
 */
export function encodeText(text, opts = {}) {
  let s = String(text ?? "").replace(/\r\n?/g, "\n");
  if (opts.eol === "\r\n") s = s.replace(/\n/g, "\r\n");
  let encoding = opts.encoding || "utf-8";
  let bom = !!opts.bom;
  if (encoding === "windows-1252" && !fitsCp1252(s)) { encoding = "utf-8"; bom = true; }
  if (encoding === "windows-1252") {
    const out = new Uint8Array(s.length);
    for (let i = 0; i < s.length; i++) {
      const c = s.charCodeAt(i);
      out[i] = c < 0x80 || (c >= 0xa0 && c <= 0xff) ? c : CP1252_REVERSE.get(c) ?? 0x3f;
    }
    return { bytes: out, encoding };
  }
  if (encoding === "utf-16le" || encoding === "utf-16be") {
    const le = encoding === "utf-16le";
    const out = new Uint8Array(s.length * 2 + 2);
    let o = 0;
    out[o++] = le ? 0xff : 0xfe; out[o++] = le ? 0xfe : 0xff; // always write a BOM
    for (let i = 0; i < s.length; i++) {
      const c = s.charCodeAt(i);
      out[o++] = le ? c & 0xff : c >> 8;
      out[o++] = le ? c >> 8 : c & 0xff;
    }
    return { bytes: out, encoding };
  }
  const body = new TextEncoder().encode(s);
  if (!bom) return { bytes: body, encoding: "utf-8" };
  const out = new Uint8Array(body.length + 3);
  out.set([0xef, 0xbb, 0xbf]);
  out.set(body, 3);
  return { bytes: out, encoding: "utf-8" };
}

/**
 * Guess the delimiter of CSV text: ",", ";" or "\t", whichever splits the
 * first lines (outside quotes) into the most consistent column count > 1.
 */
export function sniffDelimiter(text, fallback = ",") {
  const lines = [];
  let cur = "", inQ = false;
  const s = String(text || "").slice(0, 64 * 1024);
  for (let i = 0; i < s.length && lines.length < 20; i++) {
    const c = s[i];
    if (c === '"') inQ = !inQ;
    if ((c === "\n" || c === "\r") && !inQ) {
      if (cur.trim()) lines.push(cur);
      cur = "";
      continue;
    }
    cur += c;
  }
  if (cur.trim() && lines.length < 20) lines.push(cur);
  if (!lines.length) return fallback;
  let best = fallback, bestScore = 0;
  for (const d of [",", ";", "\t", "|"]) {
    const counts = lines.map((l) => {
      let n = 0, q = false;
      for (const ch of l) { if (ch === '"') q = !q; else if (ch === d && !q) n++; }
      return n;
    });
    const first = counts[0];
    if (!first) continue;
    const consistent = counts.filter((n) => n === first).length / counts.length;
    const score = consistent * 10 + Math.min(first, 50) / 50;
    if (score > bestScore + 1e-9) { best = d; bestScore = score; }
  }
  return best;
}

// ---------------------------------------------------------------------------
// Save capabilities
// ---------------------------------------------------------------------------

/**
 * Every format the viewer can write, per editor kind.
 *  - local: written in the page (no backend) when the named engine is loaded:
 *      "always" | "docx" (local-docx.modelToDocx + JSZip) | "sheetjs"
 *  - backend: may be written by the backend (only when /health lists it)
 *  - lossy(ctx): reason string when this format cannot hold the document
 *      (the doc stays "unsaved" after such an export), else null.
 */
export const SAVE_FORMATS = {
  word: [
    { id: "docx", label: "Word Document (.docx)", local: "docx", backend: true },
    { id: "rtf", label: "Rich Text Format (.rtf)", local: null, backend: true },
    { id: "md", label: "Markdown (.md)", local: "always", backend: true },
    { id: "html", label: "Web Page (.html)", local: "always", backend: true },
    { id: "tex", label: "LaTeX Document (.tex)", local: "latex", backend: true },
    { id: "txt", label: "Plain Text (.txt)", local: "always", backend: true },
  ],
  excel: [
    { id: "xlsx", label: "Excel Workbook (.xlsx)", local: "sheetjs", backend: true },
    { id: "xls", label: "Excel 97-2003 Workbook (.xls)", local: "sheetjs", backend: true },
    { id: "ods", label: "OpenDocument Spreadsheet (.ods)", local: "sheetjs", backend: true },
    { id: "csv", label: "CSV, comma separated (.csv)", local: "always", backend: true },
    { id: "tsv", label: "TSV, tab separated (.tsv)", local: "always", backend: true },
  ],
  text: [
    { id: "txt", label: "Plain Text (.txt)", local: "always", backend: false },
  ],
  slides: [
    { id: "pptx", label: "PowerPoint Presentation (.pptx)", local: "pptx", backend: true },
    { id: "html", label: "Web Page slideshow (.html)", local: "always", backend: false },
    { id: "md", label: "Markdown outline (.md)", local: "always", backend: true },
    { id: "txt", label: "Plain Text outline (.txt)", local: "always", backend: true },
  ],
  latex: [
    { id: "tex", label: "LaTeX Document (.tex)", local: "always", backend: true },
    { id: "pdf", label: "PDF (compiled with LaTeX) (.pdf)", local: null, backend: true },
    { id: "docx", label: "Word Document (.docx)", local: "docx", backend: true },
    { id: "html", label: "Web Page (.html)", local: "always", backend: true },
    { id: "md", label: "Markdown (.md)", local: "always", backend: true },
    { id: "txt", label: "Plain Text (.txt)", local: "always", backend: true },
  ],
};

/** Backend render lists assumed for a pre-capabilities (api < 2) backend. */
export const LEGACY_BACKEND_RENDER = { word: ["docx", "txt"], excel: ["xlsx", "csv", "tsv"], slides: [], latex: [] };

/**
 * Why saving the current document as `format` loses information, or null.
 * ctx: { kind, source (format the doc was opened from), sheets (excel sheet count) }
 */
export function lossyReason(format, ctx = {}) {
  const { kind, source, sheets = 1 } = ctx;
  if (kind === "word") {
    if (format === "txt") return "Plain text keeps only the words: formatting, images, tables and links are lost.";
    if (format === "html") return "A web page can be viewed but not reopened here for editing.";
    if (format === "md" && source !== "md") return "Markdown cannot store colors, fonts, sizes, alignment or highlights.";
    if (format === "tex" && source !== "tex") return "LaTeX export keeps structure, lists, tables, links and basic formatting; colors, fonts and highlights are simplified.";
    return null;
  }
  if (kind === "slides") {
    if (format === "html") return "A web page slideshow can be viewed and presented but not reopened here for editing.";
    if (format === "md" || format === "txt") return "An outline keeps only the slide text and speaker notes: layout, shapes, images and formatting are lost.";
    return null;
  }
  if (kind === "latex") {
    if (format === "pdf") return "A PDF is a finished document: keep the .tex file to make further changes.";
    if (format === "docx" || format === "html" || format === "md" || format === "txt") {
      return `Converting LaTeX to .${format} keeps the text, headings, lists and tables; custom macros, precise layout and complex math are simplified.`;
    }
    return null;
  }
  if (kind === "excel") {
    if (format === "csv" || format === "tsv") {
      const parts = [];
      if (sheets > 1) parts.push(`only the active sheet is saved (${sheets - 1} other sheet${sheets > 2 ? "s are" : " is"} left out)`);
      if (source !== "csv" && source !== "tsv") parts.push("formulas are saved as their values and cell formatting is lost");
      return parts.length ? `${format.toUpperCase()} is plain text: ${parts.join("; ")}.` : null;
    }
    return null;
  }
  return null;
}

/**
 * The formats the Save As dialog may offer for a document. Never returns a
 * format that cannot be written right now.
 *
 * @param {"word"|"excel"|"text"|"slides"|"latex"} kind
 * @param {object} caps
 *   backendRender: string[]  formats the backend renders for this kind (empty when offline)
 *   localDocx: boolean       local .docx writer loaded
 *   sheetjs: boolean         SheetJS loaded
 *   localPptx: boolean       local .pptx writer loaded (lib/local-pptx.js + JSZip)
 *   latex: boolean           lib/latex.js converters loaded (word -> .tex)
 *   source: string           format the document was opened from
 *   sheets: number           sheet count (excel)
 * @returns {{id,label,engines:("local"|"backend")[],lossy:string|null}[]}
 */
export function computeSaveTargets(kind, caps = {}) {
  const list = SAVE_FORMATS[kind] || [];
  const backend = new Set(caps.backendRender || []);
  const out = [];
  for (const f of list) {
    let local = false;
    if (f.local === "always") local = true;
    else if (f.local === "docx") local = !!caps.localDocx;
    else if (f.local === "sheetjs") local = !!caps.sheetjs;
    else if (f.local === "pptx") local = !!caps.localPptx;
    else if (f.local === "latex") local = caps.latex !== false;
    const viaBackend = !!f.backend && backend.has(f.id);
    if (!local && !viaBackend) continue;
    // Engine preference: the backend keeps more fidelity for binary office
    // formats (cell styles, column widths); simple text formats are written
    // locally (instant, no upload).
    // LaTeX -> .docx goes through the same model either way; keep it local.
    const preferBackend = ["docx", "rtf", "xlsx", "xls", "ods", "pptx", "pdf"].includes(f.id) && !(kind === "latex" && f.id === "docx");
    const engines = [];
    if (preferBackend && viaBackend) engines.push("backend");
    if (local) engines.push("local");
    if (!preferBackend && viaBackend) engines.push("backend");
    out.push({ id: f.id, label: f.label, engines, lossy: lossyReason(f.id, { kind, source: caps.source, sheets: caps.sheets }) });
  }
  return out;
}

/** MIME type to write for a saved format. */
export function mimeOf(format) {
  if (format === "html") return "text/html";
  if (format === "pdf") return "application/pdf";
  return (FORMATS[format] && FORMATS[format].mime) || "application/octet-stream";
}

/** File-picker accept descriptor for showSaveFilePicker / showOpenFilePicker. */
export function pickerTypeFor(format) {
  const label = (FORMATS[format] && FORMATS[format].label) || ({ html: "Web Page (.html)", pdf: "PDF Document (.pdf)" }[format]) || format.toUpperCase();
  return { description: label.replace(/\s*\(\.[a-z]+\)$/, ""), accept: { [mimeOf(format)]: [`.${format}`] } };
}

/** All extensions the viewer opens, for <input accept> and open pickers. */
export const OPEN_EXTENSIONS = [
  ".docx", ".doc", ".rtf", ".odt", ".md", ".markdown", ".txt",
  ".xlsx", ".xlsm", ".xls", ".ods", ".csv", ".tsv",
  ".pptx", ".tex", ".ltx",
];

/** Formats the local (backend-less) engines can fully handle. */
export function localCapabilities() {
  return {
    // SheetJS covers spreadsheet read/write; CSV/TSV are native.
    excel: ["xlsx", "xls", "xlsm", "csv", "tsv", "ods"],
    // local-docx reads/writes .docx; markdown and text are native.
    wordView: ["docx"],
    wordEdit: ["docx", "md"],
    nativeText: ["txt", "md"],
    // Need the backend: legacy and OpenDocument text formats.
    backendOnly: ["doc", "rtf", "odt"],
    // local-pptx reads/writes .pptx; LaTeX source is text (PDF needs the backend + a TeX engine).
    slides: ["pptx"],
    latex: ["tex"],
  };
}
