/**
 * Word editor features under plain Node (`npm test`): templates, header /
 * footer / document properties / page setup round trips through the
 * in-page .docx writer and reader, the HTML / Markdown writers, and the pure
 * editing helpers (change case, word count, find).
 *
 * UI behaviour (find bar, navigation pane, dialogs…) is covered by
 * tests/e2e/word-features.mjs in a real browser.
 */

import { test } from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import {
  renderWordHtml, wordModelToMarkdown, wordModelToHtmlDocument, wordModelToText, normalizeWordModel,
  changeCase, textStats, findMatches,
} from "../lib/model-word.js";
import { docxToModel, modelToDocx, splitFieldRuns } from "../lib/local-docx.js";
import { WORD_TEMPLATES, wordTemplate } from "../lib/templates-word.js";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const EXT = path.resolve(HERE, "..");

if (!globalThis.JSZip) {
  const code = fs.readFileSync(path.join(EXT, "vendor/jszip.min.js"), "utf8");
  const mod = { exports: {} };
  new Function("module", "exports", code)(mod, mod.exports);
  globalThis.JSZip = mod.exports;
}
const JSZip = globalThis.JSZip;

async function parts(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  const out = {};
  for (const name of Object.keys(zip.files)) if (!zip.files[name].dir) out[name] = await zip.file(name).async("string");
  return out;
}
const runsText = (b) => (b.runs || []).map((r) => r.t).join("");

// ---------------- templates ----------------

test("every template is a valid word model that survives .docx", async () => {
  const ids = WORD_TEMPLATES.map((t) => t.id);
  assert.deepEqual(ids, ["blank", "letter", "resume", "report", "meeting", "memo"]);
  for (const t of WORD_TEMPLATES) {
    assert.ok(t.name && t.description, t.id);
    const m = wordTemplate(t.id, { date: new Date(2026, 0, 15) });
    assert.equal(m.kind, "word", t.id);
    assert.ok(Array.isArray(m.blocks) && m.blocks.length >= 1, t.id);
    for (const b of m.blocks) {
      assert.ok(["p", "h1", "h2", "h3", "list", "table", "pagebreak"].includes(b.type), `${t.id}: ${b.type}`);
      if (b.runs) for (const r of b.runs) assert.equal(typeof r.t, "string");
    }
    assert.ok(renderWordHtml(m).length > 10);
    const back = (await docxToModel(await modelToDocx(m))).model;
    // (A document ending with a table gets one empty paragraph after it in .docx.)
    const sig = (mm) => {
      const l = normalizeWordModel(mm).blocks.map((b) => `${b.type}:${runsText(b)}`);
      while (l.length > 1 && l[l.length - 1] === "p:") l.pop();
      return l;
    };
    assert.deepEqual(sig(back), sig(m), t.id);
    if (m.meta && m.meta.footer) assert.deepEqual(back.meta.footer.runs.map((r) => r.t).join(""), m.meta.footer.runs.map((r) => r.t).join(""));
  }
  const rich = ["letter", "resume", "report", "meeting", "memo"].map((id) => wordTemplate(id));
  assert.ok(rich.every((m) => m.blocks.some((b) => /^h[12]$/.test(b.type)) || m.blocks.length > 10));
  assert.ok(wordTemplate("resume").blocks.some((b) => b.type === "table"));
  assert.ok(wordTemplate("meeting").blocks.some((b) => b.type === "list" && b.list === "ol"));
  assert.ok(wordTemplate("report").blocks.some((b) => b.type === "pagebreak"));
  assert.match(wordTemplate("report").meta.footer.runs[0].t, /\{PAGE\}.*\{NUMPAGES\}/);
});

test("templates: fresh copies, dates, unknown ids", () => {
  const a = wordTemplate("letter", { date: new Date(2026, 0, 15) });
  a.blocks[0].runs[0].t = "changed";
  assert.notEqual(wordTemplate("letter").blocks[0].runs[0].t, "changed");
  assert.ok(wordModelToText(wordTemplate("memo", { date: new Date(2026, 0, 15) })).includes("2026"));
  assert.equal(wordTemplate("nope"), null);
  assert.equal(wordTemplate("__proto__"), null);
  assert.deepEqual(wordTemplate("blank").blocks, [{ type: "p", runs: [] }]);
});

// ---------------- header / footer / properties / page setup (local .docx) ----------------

const HF_MODEL = {
  kind: "word",
  meta: {
    title: "Plan <2027> & more", author: "Ann Author", subject: "Budget",
    page: { size: "A4", orientation: "landscape", width: 16838, height: 11906, margins: { top: 720, right: 1080, bottom: 720, left: 1080, header: 400, footer: 300 } },
    header: { runs: [{ t: "Acme ", b: true }, { t: "site", link: "https://acme.example/?a=1&b=2" }, { t: "\tdraft" }], align: "center" },
    footer: { runs: [{ t: "Page {PAGE} of {NUMPAGES}", i: true, color: "#595959" }], align: "right" },
  },
  blocks: [{ type: "h1", runs: [{ t: "Title" }] }, { type: "p", runs: [{ t: "Body {PAGE} literal" }] }],
};

test("local .docx: header/footer parts with real PAGE / NUMPAGES fields", async () => {
  const p = await parts(await modelToDocx(HF_MODEL));
  const doc = p["word/document.xml"];
  assert.match(doc, /<w:sectPr><w:headerReference w:type="default" r:id="(rId\d+)"\/><w:footerReference w:type="default" r:id="rId\d+"\/><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"\/>/);
  assert.match(doc, /<w:pgMar w:top="720" w:right="1080" w:bottom="720" w:left="1080" w:header="400" w:footer="300"/);
  const rid = /headerReference w:type="default" r:id="(rId\d+)"/.exec(doc)[1];
  assert.match(p["word/_rels/document.xml.rels"], new RegExp(`Id="${rid}" Type="[^"]+/header" Target="header1.xml"`));
  const h = p["word/header1.xml"];
  assert.match(h, /^<\?xml/);
  assert.match(h, /<w:hdr [^>]*><w:p><w:pPr><w:pStyle w:val="Header"\/><w:jc w:val="center"\/><\/w:pPr>/);
  assert.match(h, /<w:hyperlink r:id="rId1"/);
  assert.match(p["word/_rels/header1.xml.rels"], /Target="https:\/\/acme.example\/\?a=1&amp;b=2" TargetMode="External"/);
  assert.match(h, /<w:tab\/>/);
  const f = p["word/footer1.xml"];
  assert.match(f, /<w:ftr /);
  assert.equal((f.match(/w:fldCharType="begin"/g) || []).length, 2);
  assert.match(f, /<w:instrText xml:space="preserve"> PAGE <\/w:instrText>/);
  assert.match(f, /<w:instrText xml:space="preserve"> NUMPAGES <\/w:instrText>/);
  assert.match(f, /<w:r><w:rPr><w:i\/><w:iCs\/><w:color w:val="595959"\/><\/w:rPr><w:fldChar w:fldCharType="begin"\/>/);
  assert.doesNotMatch(f, /\{PAGE\}/);
  const ct = p["[Content_Types].xml"];
  assert.match(ct, /PartName="\/word\/header1.xml" ContentType="application\/vnd.openxmlformats-officedocument.wordprocessingml.header\+xml"/);
  assert.match(ct, /PartName="\/word\/footer1.xml" ContentType="application\/vnd.openxmlformats-officedocument.wordprocessingml.footer\+xml"/);
  assert.match(p["word/styles.xml"], /w:styleId="Header"/);
  // Body text keeps literal braces (tokens are only fields in headers/footers).
  assert.match(doc, /Body \{PAGE\} literal/);
});

test("local .docx: core.xml carries title, creator and subject", async () => {
  const core = (await parts(await modelToDocx(HF_MODEL)))["docProps/core.xml"];
  assert.match(core, /<dc:title>Plan &lt;2027&gt; &amp; more<\/dc:title>/);
  assert.match(core, /<dc:creator>Ann Author<\/dc:creator>/);
  assert.match(core, /<dc:subject>Budget<\/dc:subject>/);
  const none = (await parts(await modelToDocx({ kind: "word", blocks: [] })))["docProps/core.xml"];
  assert.doesNotMatch(none, /dc:title|dc:creator|dc:subject/);
});

test("local .docx: header, footer, properties and page setup round trip", async () => {
  const { model, warnings } = await docxToModel(await modelToDocx(HF_MODEL));
  assert.deepEqual(warnings, []);
  assert.equal(model.meta.title, "Plan <2027> & more");
  assert.equal(model.meta.author, "Ann Author");
  assert.equal(model.meta.subject, "Budget");
  assert.deepEqual(model.meta.page, { size: "A4", orientation: "landscape", width: 16838, height: 11906, margins: { top: 720, right: 1080, bottom: 720, left: 1080, header: 400, footer: 300 } });
  assert.deepEqual(model.meta.header, HF_MODEL.meta.header);
  assert.deepEqual(model.meta.footer, HF_MODEL.meta.footer);
  assert.equal(runsText(model.blocks[1]), "Body {PAGE} literal");
});

test("local .docx: empty or malformed header/footer writes no part", async () => {
  for (const header of [{ runs: [{ t: "  " }] }, { runs: "x" }, "text", null, { runs: [] }]) {
    const p = await parts(await modelToDocx({ kind: "word", meta: { header }, blocks: [{ type: "p", runs: [{ t: "x" }] }] }));
    assert.ok(!p["word/header1.xml"], JSON.stringify(header));
    assert.doesNotMatch(p["word/document.xml"], /headerReference/);
  }
});

test("local .docx: Word-made header/footer (fields, first-page variant) read as tokens", async () => {
  const fixture = path.resolve(EXT, "../src/office-backend/tests/fixtures/word/hf-fields.docx");
  const { model, warnings } = await docxToModel(fs.readFileSync(fixture));
  assert.equal(model.meta.title, "Fixture title");
  assert.equal(model.meta.author, "Fixture Author");
  assert.equal(model.meta.subject, "Header and footer fields");
  assert.deepEqual(model.meta.header, { runs: [{ t: "Acme Corp", b: true }, { t: " — Quarterly report" }], align: "center" });
  assert.deepEqual(model.meta.footer, { runs: [{ t: "Page {PAGE} of {NUMPAGES}" }], align: "right" });
  assert.ok(warnings.some((w) => /first-page/.test(w)), warnings);
  // A PAGE field in the body keeps its displayed result.
  assert.ok(model.blocks.some((b) => runsText(b) === "Body text with a PAGE field: 3"));
});

test("local .docx: header with a table and a field without result part", async () => {
  const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
  const zip = new JSZip();
  zip.file("_rels/.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file("word/_rels/document.xml.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer3.xml"/></Relationships>');
  zip.file("word/footer3.xml", `<w:ftr ${W}><w:tbl><w:tr><w:tc><w:p><w:r><w:t>Left</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t xml:space="preserve">p. </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>PAGE</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:tc></w:tr></w:tbl><w:p/></w:ftr>`);
  zip.file("word/document.xml", `<w:document ${W}><w:body><w:p><w:r><w:t>x</w:t></w:r></w:p><w:sectPr><w:footerReference w:type="default" r:id="rId7"/></w:sectPr></w:body></w:document>`);
  const { model } = await docxToModel(await zip.generateAsync({ type: "uint8array" }));
  assert.deepEqual(model.meta.footer, { runs: [{ t: "Left\tp. {PAGE}" }] });
});

test("splitFieldRuns cuts runs at tokens and keeps formatting", () => {
  assert.deepEqual(splitFieldRuns([{ t: "A{PAGE}/{NUMPAGES}", b: true }, { t: "{X}" }]), [
    { t: "A", b: true }, { t: "", b: true, field: "PAGE" }, { t: "/", b: true }, { t: "", b: true, field: "NUMPAGES" }, { t: "{X}" },
  ]);
});

// ---------------- HTML / Markdown writers ----------------

test("HTML export shows header and footer (page fields read 1) and uses the title", () => {
  const html = wordModelToHtmlDocument(HF_MODEL);
  assert.match(html, /<title>Plan &lt;2027&gt; &amp; more<\/title>/);
  assert.match(html, /<header class="ow-doc-header" style="text-align:center">.*Acme .*site.*<\/header>/s);
  assert.match(html, /<footer class="ow-doc-footer" style="text-align:right">.*Page 1 of 1.*<\/footer>/s);
  assert.doesNotMatch(html.match(/<footer[\s\S]*<\/footer>/)[0], /\{PAGE\}|\{NUMPAGES\}/);
  assert.ok(html.indexOf("ow-doc-header") < html.indexOf("Title</span></h1>"));
  const plain = wordModelToHtmlDocument({ kind: "word", blocks: [{ type: "p", runs: [{ t: "x" }] }] });
  assert.doesNotMatch(plain, /<header|<footer/);
});

test("Markdown and plain text ignore headers and footers", () => {
  const md = wordModelToMarkdown(HF_MODEL);
  assert.doesNotMatch(md, /Acme|Page/);
  assert.match(md, /^# Title/);
  assert.doesNotMatch(wordModelToText(HF_MODEL), /Acme/);
});

// ---------------- editing helpers ----------------

test("changeCase: upper / lower / title / sentence / toggle keep the length", () => {
  const s = "hello WORLD. this is a tEST! ok? 'quoted' ß";
  assert.equal(changeCase(s, "upper"), "HELLO WORLD. THIS IS A TEST! OK? 'QUOTED' ß");
  assert.equal(changeCase(s, "lower"), "hello world. this is a test! ok? 'quoted' ß");
  assert.equal(changeCase(s, "title"), "Hello World. This Is A Test! Ok? 'quoted' ß");
  assert.equal(changeCase(s, "sentence"), "Hello world. This is a test! Ok? 'Quoted' ß");
  assert.equal(changeCase("aBc", "toggle"), "AbC");
  assert.equal(changeCase("x", "bogus"), "x");
  for (const mode of ["upper", "lower", "title", "sentence", "toggle"]) {
    for (const t of [s, "straße ǅ İstanbul 😀 ﬁ", ""]) assert.equal(changeCase(t, mode).length, t.length, `${mode}: ${t}`);
  }
  assert.equal(changeCase("o'neil's café", "title"), "O'neil's Café");
});

test("textStats counts words and characters like Word", () => {
  assert.deepEqual(textStats("Hello  world\nit's 3.5 km — fine"), { words: 6, chars: 30, charsNoSpaces: 24 });
  assert.deepEqual(textStats(""), { words: 0, chars: 0, charsNoSpaces: 0 });
  assert.equal(textStats("emoji 😀 counts as one char").charsNoSpaces, 21);
});

test("findMatches: case, whole words, special characters, limits", () => {
  const t = "The cat scattered cats. Cat! a.b (x) a+b";
  assert.deepEqual(findMatches(t, "cat"), [[4, 7], [9, 12], [18, 21], [24, 27]]);
  assert.deepEqual(findMatches(t, "cat", { wholeWord: true }), [[4, 7], [24, 27]]);
  assert.deepEqual(findMatches(t, "Cat", { matchCase: true }), [[24, 27]]);
  assert.deepEqual(findMatches(t, "a.b"), [[29, 32]]);
  assert.deepEqual(findMatches(t, "(x)"), [[33, 36]]);
  assert.deepEqual(findMatches(t, "a+b"), [[37, 40]]);
  assert.deepEqual(findMatches("aaaa", "aa"), [[0, 2], [2, 4]]);
  assert.deepEqual(findMatches("x".repeat(100), "x", { limit: 3 }).length, 3);
  assert.deepEqual(findMatches("two  spaces", "two spaces"), [[0, 11]]);
  assert.deepEqual(findMatches("", "x"), []);
  assert.deepEqual(findMatches("abc", ""), []);
  assert.deepEqual(findMatches("Ärger ärger", "ärger", { wholeWord: true }), [[0, 5], [6, 11]]);
});
