/**
 * Tests for the viewer's pure helpers: format resolution (incl. ZIP/OLE
 * peeking), text encodings, save-capability computation, the backend client
 * (header encoding, Content-Disposition, health capabilities) and the
 * markdown -> word model parser. Run with `npm test`.
 */

import { test } from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import {
  resolveFormat, sniffZipKind, sniffOleKind, formatOfMime, mimeEssence, editorKindOf,
  decodeText, encodeText, fitsCp1252, sanitizeSheetName, sniffDelimiter, withExtension,
  computeSaveTargets, lossyReason, pickerTypeFor, looksLikeText,
} from "../lib/formats.js";
import {
  createBackendClient, BackendError, filenameFromContentDisposition, backendRenderFormats,
  backendParseFormats, isLoopbackBackendUrl, normalizeBackendUrl, parseWarningsHeader,
} from "../lib/backend.js";
import { markdownToModel, parseInline } from "../lib/markdown.js";
import { wordModelToMarkdown, normalizeWordModel } from "../lib/model-word.js";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const EXAMPLE = path.resolve(HERE, "../../example");
const exampleBytes = (name) => new Uint8Array(fs.readFileSync(path.join(EXAMPLE, name)));

// ---------------------------------------------------------------- helpers

/** Minimal "stored" (uncompressed) ZIP local entries, enough for sniffing. */
function zipOf(entries, { dataDescriptor = false } = {}) {
  const parts = [];
  for (const [name, content] of entries) {
    const nameBytes = new TextEncoder().encode(name);
    const data = new TextEncoder().encode(content);
    const h = new Uint8Array(30);
    const dv = new DataView(h.buffer);
    dv.setUint32(0, 0x04034b50, true);
    dv.setUint16(4, 20, true);
    dv.setUint16(6, dataDescriptor ? 0x08 : 0, true);
    dv.setUint32(18, dataDescriptor ? 0 : data.length, true);
    dv.setUint32(22, dataDescriptor ? 0 : data.length, true);
    dv.setUint16(26, nameBytes.length, true);
    parts.push(h, nameBytes, data);
  }
  const len = parts.reduce((n, p) => n + p.length, 0);
  const out = new Uint8Array(len);
  let o = 0;
  for (const p of parts) { out.set(p, o); o += p.length; }
  return out;
}

function oleWith(streamName) {
  const out = new Uint8Array(2048);
  out.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
  for (let i = 0; i < streamName.length; i++) out[1024 + i * 2] = streamName.charCodeAt(i);
  return out;
}

// ---------------------------------------------------------------- formats

test("resolveFormat: extensionless ZIP files are identified by their entries (PL-03)", () => {
  const docx = exampleBytes("quarterly-report.docx");
  const xlsx = exampleBytes("sample-workbook.xlsx");
  assert.equal(resolveFormat("document", "", docx), "docx");
  assert.equal(resolveFormat("document", "", xlsx), "xlsx");
  assert.equal(resolveFormat("document", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", docx), "docx");
  assert.equal(resolveFormat("document", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", xlsx), "xlsx");
  assert.equal(resolveFormat("download", "application/zip", docx), "docx");
  assert.equal(sniffZipKind(docx), "docx");
  assert.equal(sniffZipKind(xlsx), "xlsx");
});

test("resolveFormat: OpenDocument mimetype entry, xlsm, streamed ZIPs", () => {
  assert.equal(resolveFormat("file", "", zipOf([["mimetype", "application/vnd.oasis.opendocument.text"], ["content.xml", "<x/>"]])), "odt");
  assert.equal(resolveFormat("file", "", zipOf([["mimetype", "application/vnd.oasis.opendocument.spreadsheet"]])), "ods");
  assert.equal(sniffZipKind(zipOf([["[Content_Types].xml", "x"], ["xl/workbook.xml", "x"], ["xl/vbaProject.bin", "x"]])), "xlsm");
  // data descriptor: sizes unknown, falls back to scanning names
  assert.equal(sniffZipKind(zipOf([["[Content_Types].xml", "<Types/>"], ["word/document.xml", "<w/>"]], { dataDescriptor: true })), "docx");
  // an unrelated zip without a helpful MIME type stays unknown
  assert.equal(resolveFormat("archive", "", zipOf([["readme.txt", "hi"]])), null);
  assert.equal(resolveFormat("archive", "application/vnd.oasis.opendocument.text", zipOf([["readme.txt", "hi"]])), "odt");
});

test("resolveFormat: OLE2 files are told apart by stream names or MIME", () => {
  assert.equal(sniffOleKind(oleWith("WordDocument")), "doc");
  assert.equal(sniffOleKind(oleWith("Workbook")), "xls");
  assert.equal(resolveFormat("document", "", oleWith("WordDocument")), "doc");
  assert.equal(resolveFormat("document", "", oleWith("Workbook")), "xls");
  assert.equal(resolveFormat("document", "application/msword", oleWith("Unknown")), "doc");
});

test("resolveFormat: MIME parameters are ignored and aliases are known", () => {
  const csv = new TextEncoder().encode("a,b\n1,2\n");
  assert.equal(resolveFormat("document", "text/csv; charset=utf-8", csv), "csv");
  assert.equal(resolveFormat("document", "application/csv", csv), "csv");
  assert.equal(resolveFormat("document", "text/markdown; charset=UTF-8", new TextEncoder().encode("# hi")), "md");
  assert.equal(resolveFormat("document", "", csv), "txt");
  assert.equal(mimeEssence("Text/CSV; charset=utf-8"), "text/csv");
  assert.equal(formatOfMime("application/x-rtf"), "rtf");
  // binary MIME claims without matching magic are not trusted
  assert.equal(resolveFormat("document", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", new Uint8Array([1, 2, 3, 0])), null);
  // HTML (e.g. a login page) is never a document
  assert.equal(resolveFormat("document", "text/html", new TextEncoder().encode("<html></html>")), null);
});

test("resolveFormat: extension wins, but obvious mislabels are corrected", () => {
  assert.equal(resolveFormat("letter.doc", "", new TextEncoder().encode("{\\rtf1 hi}")), "rtf");
  assert.equal(resolveFormat("letter.doc", "", exampleBytes("quarterly-report.docx")), "docx");
  assert.equal(resolveFormat("book.xls", "", exampleBytes("sample-workbook.xlsx")), "xlsx");
  assert.equal(resolveFormat("notes.markdown", "", new TextEncoder().encode("x")), "md");
  assert.equal(resolveFormat("Report.DOCX", "", new Uint8Array(0)), "docx");
});

test("editorKindOf maps formats to editors", () => {
  assert.equal(editorKindOf("docx"), "word");
  assert.equal(editorKindOf("md"), "word");
  assert.equal(editorKindOf("txt"), "text");
  assert.equal(editorKindOf("tsv"), "excel");
  assert.equal(editorKindOf("pdf"), null);
});

test("decodeText detects BOMs, UTF-16 and Windows-1252; encodeText round-trips", () => {
  const u8 = (a) => new Uint8Array(a);
  let d = decodeText(u8([0xef, 0xbb, 0xbf, 0x68, 0x69]));
  assert.deepEqual([d.text, d.encoding, d.bom], ["hi", "utf-8", true]);
  d = decodeText(u8([0xff, 0xfe, 0x68, 0, 0xfc, 0]));
  assert.deepEqual([d.text, d.encoding, d.bom], ["hü", "utf-16le", true]);
  d = decodeText(u8([0xfe, 0xff, 0, 0x68, 0, 0xfc]));
  assert.deepEqual([d.text, d.encoding], ["hü", "utf-16be"]);
  d = decodeText(u8([0x4d, 0xfc, 0x6c, 0x6c, 0x65, 0x72, 0x3b, 0x80])); // "Müller;€" in cp1252
  assert.deepEqual([d.text, d.encoding], ["Müller;€", "windows-1252"]);
  d = decodeText(new TextEncoder().encode("a\r\nb\r\n"));
  assert.equal(d.eol, "\r\n");
  assert.equal(looksLikeText(u8([0xff, 0xfe, 0x68, 0])), true);

  // round trips
  for (const src of [
    u8([0xef, 0xbb, 0xbf, 0x68, 0xc3, 0xa9, 0x0d, 0x0a]),
    u8([0xff, 0xfe, 0x68, 0, 0xfc, 0]),
    u8([0x4d, 0xfc, 0x6c, 0x6c, 0x65, 0x72, 0x0d, 0x0a, 0x80]),
  ]) {
    const dec = decodeText(src);
    const enc = encodeText(dec.text, dec);
    assert.deepEqual([...enc.bytes], [...src]);
  }
  // cp1252 text that gains a character outside the code page becomes UTF-8 + BOM
  assert.equal(fitsCp1252("Müller €"), true);
  assert.equal(fitsCp1252("✓"), false);
  const r = encodeText("ok ✓", { encoding: "windows-1252" });
  assert.equal(r.encoding, "utf-8");
  assert.deepEqual([...r.bytes.subarray(0, 3)], [0xef, 0xbb, 0xbf]);
});

test("sanitizeSheetName and withExtension", () => {
  assert.equal(sanitizeSheetName("Quarterly sales export 2024-01-15 final"), "Quarterly sales export 2024-01-");
  assert.equal(sanitizeSheetName("report [final]"), "report _final_");
  assert.equal(sanitizeSheetName("'quoted'"), "quoted");
  assert.equal(sanitizeSheetName("///"), "___");
  assert.equal(sanitizeSheetName(""), "Sheet1");
  assert.equal(withExtension("report.rtf", "docx"), "report.docx");
  assert.equal(withExtension("Résumé (final)", "md"), "Résumé (final).md");
  assert.equal(withExtension("", "xlsx"), "Untitled.xlsx");
});

test("sniffDelimiter", () => {
  assert.equal(sniffDelimiter("a;b;c\n1;2,5;3\n"), ";");
  assert.equal(sniffDelimiter("a\tb\n1\t2\n"), "\t");
  assert.equal(sniffDelimiter('a,b\n"x;y",2\n'), ",");
  assert.equal(sniffDelimiter("single column\nvalues\n"), ",");
});

// ---------------------------------------------------------------- save capabilities

test("computeSaveTargets never offers a format that cannot be written (WO-10, PL-15, EX-17)", () => {
  const offline = computeSaveTargets("word", { backendRender: [], localDocx: true, source: "docx" });
  assert.deepEqual(offline.map((t) => t.id), ["docx", "md", "html", "tex", "txt"]);
  assert.deepEqual(offline[0].engines, ["local"]);
  const noDocxWriter = computeSaveTargets("word", { backendRender: [], localDocx: false, source: "docx" });
  assert.ok(!noDocxWriter.some((t) => t.id === "docx"));
  const online = computeSaveTargets("word", { backendRender: ["docx", "rtf", "txt", "md", "html"], localDocx: true, source: "rtf" });
  assert.deepEqual(online.map((t) => t.id), ["docx", "rtf", "md", "html", "tex", "txt"]);
  assert.deepEqual(online[0].engines, ["backend", "local"]);
  assert.deepEqual(online.find((t) => t.id === "rtf").engines, ["backend"]);
  assert.deepEqual(online.find((t) => t.id === "md").engines, ["local", "backend"]);

  const excel = computeSaveTargets("excel", { backendRender: [], sheetjs: true, source: "tsv", sheets: 1 });
  assert.deepEqual(excel.map((t) => t.id), ["xlsx", "xls", "ods", "csv", "tsv"]);
  assert.equal(excel.find((t) => t.id === "tsv").lossy, null);
  const noSheetjs = computeSaveTargets("excel", { backendRender: ["xlsx", "csv", "tsv"], sheetjs: false, source: "xlsx", sheets: 2 });
  assert.deepEqual(noSheetjs.map((t) => t.id), ["xlsx", "csv", "tsv"]);
  assert.match(noSheetjs.find((t) => t.id === "csv").lossy, /only the active sheet/);
  assert.deepEqual(computeSaveTargets("text", {}).map((t) => t.id), ["txt"]);
});

test("lossyReason flags exports that cannot hold the document (W-02, EX-18)", () => {
  assert.match(lossyReason("txt", { kind: "word", source: "docx" }), /formatting/);
  assert.match(lossyReason("md", { kind: "word", source: "docx" }), /Markdown/);
  assert.equal(lossyReason("md", { kind: "word", source: "md" }), null);
  assert.equal(lossyReason("docx", { kind: "word", source: "md" }), null);
  assert.equal(lossyReason("csv", { kind: "excel", source: "csv", sheets: 1 }), null);
  assert.match(lossyReason("csv", { kind: "excel", source: "csv", sheets: 3 }), /2 other sheets are left out/);
  assert.match(lossyReason("csv", { kind: "excel", source: "xlsx", sheets: 1 }), /formulas/);
  assert.equal(lossyReason("xlsx", { kind: "excel", source: "csv", sheets: 2 }), null);
});

test("pickerTypeFor builds File System Access accept types", () => {
  assert.deepEqual(pickerTypeFor("docx").accept, { "application/vnd.openxmlformats-officedocument.wordprocessingml.document": [".docx"] });
  assert.deepEqual(pickerTypeFor("html").accept, { "text/html": [".html"] });
});

// ---------------------------------------------------------------- backend client

function mockFetch(handler) {
  const calls = [];
  const fn = async (url, opts) => { calls.push({ url, opts }); return handler(url, opts); };
  fn.calls = calls;
  return fn;
}

test("backend parse percent-encodes X-File-Name (W-M1, PL-06, BA-09)", async () => {
  const f = mockFetch(() => new Response(JSON.stringify({ kind: "word", format: "docx", model: { kind: "word", blocks: [] } }), { status: 200 }));
  const c = createBackendClient("http://127.0.0.1:1", f);
  for (const name of ["Отчёт.docx", "报告.docx", "Bob’s report.docx", "Résumé 100%.docx"]) {
    const res = await c.parse(new Uint8Array([1]), name);
    const header = f.calls[f.calls.length - 1].opts.headers["X-File-Name"];
    assert.equal(header, encodeURIComponent(name));
    assert.ok(/^[\x20-\x7e]+$/.test(header), "header must be ASCII");
    assert.doesNotThrow(() => new Headers({ "X-File-Name": header }));
    assert.deepEqual(res.warnings, []);
  }
});

test("backend render decodes filename* (WO-23, BA-10)", async () => {
  const f = mockFetch(() => new Response(new Uint8Array([1, 2]), {
    status: 200,
    headers: {
      "Content-Disposition": `attachment; filename="R_sum_.docx"; filename*=UTF-8''R%C3%A9sum%C3%A9%20%E6%8A%A5%E5%91%8A.docx`,
      "X-Output-Format": "docx",
    },
  }));
  const c = createBackendClient("http://x", f);
  const r = await c.render({ kind: "word", blocks: [] }, "docx", "Résumé 报告.docx");
  assert.equal(r.filename, "Résumé 报告.docx");
  assert.equal(r.format, "docx");
  assert.equal(JSON.parse(f.calls[0].opts.body).name, "Résumé 报告.docx");
});

test("filenameFromContentDisposition handles plain, quoted, starred and malformed values (PL-22)", () => {
  assert.equal(filenameFromContentDisposition('attachment; filename="100% done.docx"'), "100% done.docx");
  assert.equal(filenameFromContentDisposition("attachment; filename=plain.xlsx"), "plain.xlsx");
  assert.equal(filenameFromContentDisposition("attachment; filename*=UTF-8''%E2%82%AC.csv"), "€.csv");
  assert.equal(filenameFromContentDisposition("attachment; filename*=UTF-8''bad%E2.csv"), "bad%E2.csv");
  assert.equal(filenameFromContentDisposition("attachment; filename*=iso-8859-1'en'M%FCller.txt"), "Müller.txt");
  assert.equal(filenameFromContentDisposition('attachment; filename="a \\"b\\".txt"'), 'a "b".txt');
  assert.equal(filenameFromContentDisposition(""), "");
  assert.equal(filenameFromContentDisposition(null), "");
});

test("health exposes capabilities; legacy backends get the known render set", async () => {
  const v2 = createBackendClient("http://x", mockFetch(() => new Response(JSON.stringify({
    ok: true, name: "office-backend", version: "0.2.0", api: 2,
    parse: ["docx", "rtf"], render: { word: ["docx", "rtf", "txt", "md", "html"], excel: ["xlsx", "csv", "tsv"] },
  }), { status: 200 })));
  const h = await v2.health();
  assert.equal(h.ok, true);
  assert.equal(h.api, 2);
  assert.deepEqual(backendRenderFormats(h, "word"), ["docx", "rtf", "txt", "md", "html"]);
  assert.deepEqual(backendParseFormats(h), ["docx", "rtf"]);

  const legacy = await createBackendClient("http://x", mockFetch(() => new Response(JSON.stringify({ ok: true, name: "office-backend" }), { status: 200 }))).health();
  assert.deepEqual(backendRenderFormats(legacy, "word"), ["docx", "txt"]);
  assert.deepEqual(backendRenderFormats(legacy, "excel"), ["xlsx", "csv", "tsv"]);
  assert.equal(backendParseFormats(legacy), null);
  assert.deepEqual(backendRenderFormats({ ok: false }, "word"), []);
});

test("network failures become actionable BackendErrors; health never throws", async () => {
  const down = createBackendClient("http://127.0.0.1:9", async () => { throw new TypeError("Failed to fetch"); });
  const h = await down.health();
  assert.equal(h.ok, false);
  assert.match(h.error, /Cannot reach the office backend/);
  await assert.rejects(() => down.parse(new Uint8Array([1]), "a.docx"), (e) => e instanceof BackendError && e.network === true && /office-backend running/.test(e.message));
  const http = createBackendClient("http://x", async () => new Response(JSON.stringify({ error: "unsupported file type" }), { status: 415 }));
  await assert.rejects(() => http.parse(new Uint8Array([1]), "a.bin"), (e) => e.status === 415 && e.message === "unsupported file type" && !e.network);
  const big = createBackendClient("http://x", async () => new Response("", { status: 413 }));
  await assert.rejects(() => big.parse(new Uint8Array([1]), "a.docx"), /too large/);
});

test("isLoopbackBackendUrl", () => {
  assert.equal(isLoopbackBackendUrl("http://127.0.0.1:47831"), true);
  assert.equal(isLoopbackBackendUrl("localhost:9000"), true);
  assert.equal(isLoopbackBackendUrl("http://[::1]:47831"), true);
  assert.equal(isLoopbackBackendUrl("https://office.example.com"), false);
  assert.equal(normalizeBackendUrl(" 127.0.0.1:1/ "), "http://127.0.0.1:1");
});

// ---------------------------------------------------------------- markdown

const T = (blocks) => normalizeWordModel({ kind: "word", blocks }).blocks;

test("markdown: headings, paragraphs, breaks and inline marks", () => {
  const m = markdownToModel([
    "# Title",
    "",
    "Setext",
    "======",
    "",
    "Some **bold**, *italic*, ***both***, ~~gone~~, `code` and [a link](https://x.test/a_(b) \"t\").",
    "Soft wrapped line",
    "hard  ",
    "break and backslash\\",
    "break.",
    "",
    "snake_case_word, 2 * 3, a * b, <u>under</u>, H<sub>2</sub>O, x<sup>2</sup> &amp; &#169;",
  ].join("\n"));
  assert.deepEqual(T(m.blocks), T([
    { type: "h1", runs: [{ t: "Title" }] },
    { type: "h1", runs: [{ t: "Setext" }] },
    {
      type: "p",
      runs: [
        { t: "Some " }, { t: "bold", b: true }, { t: ", " }, { t: "italic", i: true }, { t: ", " },
        { t: "both", b: true, i: true }, { t: ", " }, { t: "gone", s: true }, { t: ", " },
        { t: "code", font: "Consolas" }, { t: " and " }, { t: "a link", link: "https://x.test/a_(b)" },
        { t: ". Soft wrapped line hard\nbreak and backslash\nbreak." },
      ],
    },
    {
      type: "p",
      runs: [
        { t: "snake_case_word, 2 * 3, a * b, " }, { t: "under", u: true }, { t: ", H" }, { t: "2", sub: true },
        { t: "O, x" }, { t: "2", sup: true }, { t: " & ©" },
      ],
    },
  ]));
});

test("markdown: nested lists (2 and 4 space indents), ordered lists, quotes, code, tables, breaks", () => {
  const m = markdownToModel([
    "- one",
    "  - two (2 spaces)",
    "    - three",
    "- back",
    "",
    "1. first",
    "    1. nested (4 spaces)",
    "2) second",
    "",
    "> quoted **text**",
    "> continues",
    "",
    "```js",
    "const a = 1;",
    "",
    "  indented",
    "```",
    "",
    "    indented code",
    "",
    "| A | B |",
    "|:--|--:|",
    "| 1 | x<br>y |",
    "| 2 |",
    "",
    "***",
    "",
    "- [ ] task",
  ].join("\n"));
  assert.deepEqual(T(m.blocks), T([
    { type: "list", list: "ul", level: 0, runs: [{ t: "one" }] },
    { type: "list", list: "ul", level: 1, runs: [{ t: "two (2 spaces)" }] },
    { type: "list", list: "ul", level: 2, runs: [{ t: "three" }] },
    { type: "list", list: "ul", level: 0, runs: [{ t: "back" }] },
    { type: "list", list: "ol", level: 0, runs: [{ t: "first" }] },
    { type: "list", list: "ol", level: 1, runs: [{ t: "nested (4 spaces)" }] },
    { type: "list", list: "ol", level: 0, runs: [{ t: "second" }] },
    { type: "quote", runs: [{ t: "quoted " }, { t: "text", b: true }, { t: " continues" }] },
    { type: "code", runs: [{ t: "const a = 1;\n\n  indented" }] },
    { type: "code", runs: [{ t: "indented code" }] },
    {
      type: "table",
      rows: [
        [{ blocks: [{ type: "p", runs: [{ t: "A" }] }] }, { blocks: [{ type: "p", runs: [{ t: "B" }] }] }],
        [{ blocks: [{ type: "p", runs: [{ t: "1" }] }] }, { blocks: [{ type: "p", runs: [{ t: "x" }] }, { type: "p", runs: [{ t: "y" }] }] }],
        [{ blocks: [{ type: "p", runs: [{ t: "2" }] }] }, { blocks: [{ type: "p", runs: [] }] }],
      ],
    },
    { type: "pagebreak" },
    { type: "list", list: "ul", level: 0, runs: [{ t: "☐ task" }] },
  ]));
});

test("markdown: images, autolinks, unsafe links, escapes, empty input", () => {
  const png = "data:image/png;base64,iVBORw0KGgo=";
  const m = markdownToModel(`![logo](${png})\n\nSee <https://a.test/x> or https://b.test/y. Mail <me@x.test>.\n\n[bad](javascript:alert(1)) \\*not em\\* \\# no heading`);
  assert.deepEqual(m.blocks[0], { type: "img", src: png, alt: "logo" });
  assert.deepEqual(T([m.blocks[1]]), T([{
    type: "p",
    runs: [
      { t: "See " }, { t: "https://a.test/x", link: "https://a.test/x" }, { t: " or " },
      { t: "https://b.test/y", link: "https://b.test/y" }, { t: ". Mail " }, { t: "me@x.test", link: "mailto:me@x.test" }, { t: "." },
    ],
  }]));
  assert.deepEqual(T([m.blocks[2]]), T([{ type: "p", runs: [{ t: "bad *not em* # no heading" }] }]));
  assert.deepEqual(markdownToModel("").blocks, [{ type: "p", runs: [] }]);
  assert.deepEqual(markdownToModel("﻿hi\r\n").blocks, [{ type: "p", runs: [{ t: "hi" }] }]);
  assert.deepEqual(parseInline("**unclosed"), [{ t: "**unclosed" }]);
});

test("markdown: wordModelToMarkdown output parses back to the same model", () => {
  const model = {
    kind: "word",
    blocks: [
      { type: "h1", runs: [{ t: "Report " }, { t: "2026", i: true }] },
      { type: "p", runs: [{ t: "Mixed " }, { t: "bold", b: true }, { t: " " }, { t: "strike", s: true }, { t: " " }, { t: "under", u: true }, { t: " and " }, { t: "x", sup: true }, { t: " " }, { t: "mono", font: "Consolas" }] },
      { type: "p", runs: [{ t: "Line one\nline two with a [bracket] and *star* and _under_score" }] },
      { type: "p", runs: [{ t: "Link: " }, { t: "site", link: "https://example.com/a?b=1&c=2" }] },
      { type: "list", list: "ul", level: 0, runs: [{ t: "a" }] },
      { type: "list", list: "ul", level: 1, runs: [{ t: "b " }, { t: "bold", b: true }] },
      { type: "list", list: "ul", level: 2, runs: [{ t: "c" }] },
      { type: "list", list: "ul", level: 0, runs: [{ t: "d" }] },
      { type: "p", runs: [{ t: "between" }] },
      { type: "list", list: "ol", level: 0, runs: [{ t: "one" }] },
      { type: "list", list: "ol", level: 0, runs: [{ t: "two" }] },
      { type: "list", list: "ol", level: 1, runs: [{ t: "two.a" }] },
      { type: "quote", runs: [{ t: "A quote" }] },
      { type: "code", runs: [{ t: "fn main() {\n    println!(\"hi\");\n}" }] },
      { type: "pagebreak" },
      { type: "h2", runs: [{ t: "Table" }] },
      { type: "table", rows: [
        [{ blocks: [{ type: "p", runs: [{ t: "H1" }] }] }, { blocks: [{ type: "p", runs: [{ t: "H|2" }] }] }],
        [{ blocks: [{ type: "p", runs: [{ t: "v1" }] }, { type: "p", runs: [{ t: "v1b" }] }] }, { blocks: [{ type: "p", runs: [{ t: "v2", b: true }] }] }],
      ] },
      { type: "img", src: "data:image/png;base64,AAAA", alt: "pic" },
      { type: "p", runs: [{ t: "# not a heading, 1. not a list" }] },
    ],
  };
  const md = wordModelToMarkdown(model);
  const back = markdownToModel(md);
  assert.deepEqual(normalizeWordModel(back), normalizeWordModel(model), md);
  // and the markdown is stable
  assert.equal(wordModelToMarkdown(back), md);
});

test("backend render returns X-Office-Warnings notes (renamed sheets, single-sheet CSV, ...)", async () => {
  const notes = ['Sheet name "a/b" is not allowed in Excel; it was saved as "a_b".', "Only the active sheet was saved."];
  const f = mockFetch(() => new Response(new Uint8Array([1]), {
    status: 200,
    headers: { "X-Output-Format": "xlsx", "X-Office-Warnings": encodeURIComponent(JSON.stringify(notes)) },
  }));
  const r = await createBackendClient("http://x", f).render({ kind: "excel", sheets: [] }, "xlsx", "b.xlsx");
  assert.deepEqual(r.warnings, notes);
  assert.deepEqual(parseWarningsHeader(null), []);
  assert.deepEqual(parseWarningsHeader("%E0%A4%A"), []); // malformed: ignored
  assert.deepEqual(parseWarningsHeader(encodeURIComponent('{"not":"a list"}')), []);
  assert.deepEqual(parseWarningsHeader(encodeURIComponent('["ok", 3, ""]')), ["ok"]);
});
