/**
 * Word model tests that run under plain Node (`npm test`): rendering,
 * text/markdown/html export, normalization and the client-side .docx
 * reader/writer (lib/local-docx.js, using the vendored JSZip).
 *
 * DOM parsing (parseWordDom against real contenteditable/execCommand output)
 * needs a browser: see tests/browser/run.mjs.
 */

import { test } from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import {
  renderWordHtml, wordModelToText, wordModelToMarkdown, wordModelToHtmlDocument,
  normalizeWordModel, colorToHex,
} from "../lib/model-word.js";
import {
  docxToModel, modelToDocx, parseXml, imageSize, docxPreviewAvailable,
} from "../lib/local-docx.js";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const EXT = path.resolve(HERE, "..");

// Load the vendored UMD JSZip as the global the extension pages use.
if (!globalThis.JSZip) {
  const code = fs.readFileSync(path.join(EXT, "vendor/jszip.min.js"), "utf8");
  const mod = { exports: {} };
  new Function("module", "exports", code)(mod, mod.exports);
  globalThis.JSZip = mod.exports;
}
const JSZip = globalThis.JSZip;

// 2x1 red/blue PNG.
const PNG_B64 =
  "iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAIAAAB7QOjdAAAAD0lEQVR4nGP4z8DAwPAfAAcAAf9+CLHQAAAAAElFTkSuQmCC";
const PNG_URL = `data:image/png;base64,${PNG_B64}`;

const RICH = {
  kind: "word",
  meta: { title: "Rich & <Test>" },
  blocks: [
    { type: "h1", runs: [{ t: "Title" }] },
    { type: "p", align: "center", runs: [{ t: "Centered " }, { t: "bold", b: true }, { t: " and " }, { t: "italic", i: true }] },
    {
      type: "p",
      align: "justify",
      runs: [
        { t: "under", u: true }, { t: " " }, { t: "strike", s: true }, { t: " E=mc" }, { t: "2", sup: true },
        { t: " H" }, { t: "2", sub: true }, { t: "O " }, { t: "red", color: "#FF0000" }, { t: " " },
        { t: "hl", hl: "#FFFF00" }, { t: " " }, { t: "custom", hl: "#ABCDEF" }, { t: " " },
        { t: "big", size: 18 }, { t: " " }, { t: "serif", font: "Times New Roman" },
      ],
    },
    { type: "p", runs: [{ t: "See " }, { t: "the site", link: "https://example.com/a?b=1&c=2" }, { t: " or " }, { t: "jump", link: "#sec" }, { t: "." }] },
    { type: "p", runs: [] },
    { type: "list", list: "ul", level: 0, runs: [{ t: "Parent" }] },
    { type: "list", list: "ul", level: 1, runs: [{ t: "Child" }] },
    { type: "list", list: "ol", level: 2, runs: [{ t: "Grandchild ordered" }] },
    { type: "list", list: "ul", level: 0, runs: [{ t: "Next" }] },
    { type: "p", runs: [{ t: "between" }] },
    { type: "list", list: "ol", level: 0, runs: [{ t: "One" }] },
    { type: "list", list: "ol", level: 0, runs: [{ t: "Two" }] },
    { type: "quote", runs: [{ t: "A quotation" }] },
    { type: "code", runs: [{ t: "let x = 1;\n\tindent()" }] },
    {
      type: "table",
      rows: [
        [{ blocks: [{ type: "p", runs: [{ t: "Wide", b: true }] }], colspan: 2 }],
        [
          { blocks: [{ type: "p", runs: [{ t: "a" }] }] },
          { blocks: [{ type: "list", list: "ul", level: 0, runs: [{ t: "in cell" }] }] },
        ],
        [{ blocks: [{ type: "p", runs: [] }] }, { blocks: [{ type: "p", runs: [{ t: "z" }] }] }],
      ],
    },
    { type: "img", src: PNG_URL, alt: "pixels", w: 40, h: 20 },
    { type: "pagebreak" },
    { type: "h2", runs: [{ t: "After break" }] },
    { type: "p", runs: [{ t: "line1\nline2" }] },
  ],
};

// ---------------- rendering ----------------

test("renderWordHtml renders every model feature", () => {
  const html = renderWordHtml(RICH);
  assert.match(html, /<h1><span>Title<\/span><\/h1>/);
  assert.match(html, /<p style="text-align:center">/);
  assert.doesNotMatch(html, /data-align/);
  assert.match(html, /font-weight:bold/);
  assert.match(html, /font-style:italic/);
  assert.match(html, /text-decoration:underline/);
  assert.match(html, /text-decoration:line-through/);
  assert.match(html, /<sup>2<\/sup>/);
  assert.match(html, /<sub>2<\/sub>/);
  assert.match(html, /color:#FF0000/);
  assert.match(html, /background-color:#FFFF00/);
  assert.match(html, /font-size:18pt/);
  assert.match(html, /font-family:'Times New Roman'/);
  assert.match(html, /<a class="ow-link" href="https:\/\/example.com\/a\?b=1&amp;c=2"/);
  assert.match(html, /<blockquote data-ow="quote">/);
  assert.match(html, /<pre>/);
  assert.match(html, /<td colspan="2">/);
  assert.match(html, /<img class="ow-img" src="data:image\/png;base64,[^"]+" alt="pixels" width="40" height="20">/);
  assert.match(html, /<hr class="ow-pagebreak" data-ow="pagebreak"/);
  assert.match(html, /line1<br>line2/);
  // Nested list structure: child list inside the parent item, in order.
  assert.match(html, /<ul><li><span>Parent<\/span><ul><li><span>Child<\/span><ol><li><span>Grandchild ordered<\/span><\/li><\/ol><\/li><\/ul><\/li><li><span>Next<\/span><\/li><\/ul>/);
});

test("links carry no display-only inline styles", () => {
  const html = renderWordHtml({ blocks: [{ type: "p", runs: [{ t: "go", link: "https://x.test" }] }] });
  assert.doesNotMatch(html, /1a5fb4/i);
  assert.doesNotMatch(html, /<a[^>]*style=/);
  assert.match(html, /<a class="ow-link" href="https:\/\/x.test" target="_blank" rel="noopener noreferrer"><span>go<\/span><\/a>/);
});

test("unsafe link schemes are not rendered as live hrefs", () => {
  const html = renderWordHtml({ blocks: [{ type: "p", runs: [{ t: "x", link: "javascript:alert(1)" }] }] });
  assert.doesNotMatch(html, /\shref="javascript/);
  assert.match(html, /data-href="javascript:alert\(1\)"/);
});

test("spaces that HTML would collapse are rendered as nbsp; trailing newline gets a placeholder", () => {
  const html = renderWordHtml({ blocks: [{ type: "p", runs: [{ t: " a  b " }] }, { type: "p", runs: [{ t: "end\n" }] }] });
  assert.match(html, /<p><span>&nbsp;a &nbsp;b&nbsp;<\/span><\/p>/);
  assert.match(html, /<p><span>end<br><\/span><br><\/p>/);
});

test("escapes text and attributes", () => {
  const html = renderWordHtml({ blocks: [{ type: "p", runs: [{ t: "<b>&\"" }] }, { type: "img", src: 'data:x" onerror="y', alt: '"' }] });
  assert.match(html, /&lt;b&gt;&amp;&quot;/);
  assert.doesNotMatch(html, /onerror="y/);
});

test("empty model renders one editable empty paragraph", () => {
  assert.equal(renderWordHtml({ blocks: [] }), "<p><br></p>");
  assert.equal(renderWordHtml(null), "<p><br></p>");
  assert.deepEqual(normalizeWordModel({ blocks: [] }), normalizeWordModel({ blocks: [{ type: "p", runs: [] }] }));
  assert.equal(renderWordHtml({ blocks: [{ type: "p", runs: [] }] }), "<p><br></p>");
});

// ---------------- helpers ----------------

test("colorToHex", () => {
  assert.equal(colorToHex("rgb(255, 0, 16)"), "#FF0010");
  assert.equal(colorToHex("#abc"), "#AABBCC");
  assert.equal(colorToHex("red"), "#FF0000");
  assert.equal(colorToHex("rgba(0, 0, 0, 0)"), null);
  assert.equal(colorToHex("transparent"), null);
  assert.equal(colorToHex("00ff00"), "#00FF00");
});

test("normalizeWordModel canonicalizes equivalent models", () => {
  const a = normalizeWordModel({
    blocks: [
      { type: "p", align: "left", runs: [{ t: "a", color: "#ff0000" }, { t: "b", color: "#FF0000" }, { t: "" }] },
      { type: "table", rows: [[{ blocks: [] }]] },
    ],
  });
  assert.deepEqual(a, {
    kind: "word",
    blocks: [
      { type: "p", runs: [{ t: "ab", color: "#FF0000" }] },
      { type: "table", rows: [[{ blocks: [{ type: "p", runs: [] }] }]] },
    ],
  });
});

// ---------------- text / markdown / html ----------------

test("wordModelToText numbers ordered lists per level", () => {
  const text = wordModelToText(RICH);
  assert.match(text, /^Title\n/);
  assert.match(text, /• Parent\n  • Child\n    1\. Grandchild ordered\n• Next/);
  assert.match(text, /1\. One\n2\. Two/);
  assert.match(text, /Wide\n/);
  assert.match(text, /\[image: pixels\]/);
  assert.match(text, /\f/);
});

test("wordModelToMarkdown", () => {
  const md = wordModelToMarkdown(RICH);
  assert.match(md, /^# Title\n\n/);
  assert.match(md, /Centered \*\*bold\*\* and \*italic\*/);
  assert.match(md, /<u>under<\/u> ~~strike~~ E=mc<sup>2<\/sup>/);
  assert.match(md, /\[the site\]\(https:\/\/example.com\/a\?b=1&c=2\)/);
  assert.match(md, /- Parent\n {4}- Child\n {8}1\. Grandchild ordered\n- Next/);
  assert.match(md, /1\. One\n2\. Two/);
  assert.match(md, /> A quotation/);
  assert.match(md, /```\nlet x = 1;\n\tindent\(\)\n```/);
  assert.match(md, /\| \*\*Wide\*\* \|   \|\n\| --- \| --- \|\n\| a \| in cell \|/);
  assert.match(md, /!\[pixels\]\(data:image\/png;base64,/);
  assert.match(md, /\n---\n/);
  assert.match(md, /## After break/);
  assert.match(md, /line1\\\nline2/);
});

test("wordModelToMarkdown escapes markdown syntax in text", () => {
  const md = wordModelToMarkdown({ blocks: [{ type: "p", runs: [{ t: "# not a heading *x* [y] 1. z" }] }, { type: "p", runs: [{ t: "- dash" }] }] });
  assert.match(md, /^\\# not a heading \\\*x\\\* \\\[y\\\] 1\. z/);
  assert.match(md, /\n\\- dash/);
  // emphasis delimiters never wrap leading/trailing spaces
  const md2 = wordModelToMarkdown({ blocks: [{ type: "p", runs: [{ t: "a" }, { t: " bold ", b: true }, { t: "b" }] }] });
  assert.equal(md2, "a **bold** b\n");
});

test("wordModelToHtmlDocument is a standalone document", () => {
  const html = wordModelToHtmlDocument(RICH, "My <Doc>");
  assert.match(html, /^<!DOCTYPE html>/);
  assert.match(html, /<meta charset="utf-8">/);
  assert.match(html, /<title>My &lt;Doc&gt;<\/title>/);
  assert.match(html, /<article class="ow-page">/);
  assert.match(html, /<h1><span>Title<\/span><\/h1>/);
  assert.doesNotMatch(html, /contenteditable/);
  assert.match(wordModelToHtmlDocument(RICH), /<title>Rich &amp; &lt;Test&gt;<\/title>/);
});

// ---------------- local docx ----------------

test("parseXml handles entities, CDATA, comments and prefixes", () => {
  const doc = parseXml('<?xml version="1.0"?><!-- c --><w:a xmlns:w="u" w:v="1 &amp; 2"><w:t>x &lt; y</w:t><![CDATA[<raw>]]><e/></w:a>');
  const a = doc.children[0];
  assert.equal(a.local, "a");
  assert.equal(a.attrs["w:v"], "1 & 2");
  assert.equal(a.children[0].children[0].text, "x < y");
  assert.equal(a.children[1].text, "<raw>");
  assert.equal(a.children[2].local, "e");
});

test("imageSize sniffs PNG", () => {
  const bytes = Uint8Array.from(atob(PNG_B64), (c) => c.charCodeAt(0));
  assert.deepEqual(imageSize(bytes), { w: 2, h: 1 });
});

test("docxPreviewAvailable is false without the library", () => {
  assert.equal(docxPreviewAvailable(), false);
});

test("modelToDocx writes a complete package", async () => {
  const bytes = await modelToDocx(RICH);
  assert.ok(bytes instanceof Uint8Array);
  const zip = await JSZip.loadAsync(bytes);
  const names = Object.keys(zip.files);
  for (const n of ["[Content_Types].xml", "_rels/.rels", "word/document.xml", "word/_rels/document.xml.rels",
    "word/styles.xml", "word/numbering.xml", "word/settings.xml", "docProps/core.xml", "word/media/image1.png"]) {
    assert.ok(names.includes(n), `missing ${n}`);
  }
  const ct = await zip.file("[Content_Types].xml").async("string");
  assert.match(ct, /Extension="png" ContentType="image\/png"/);
  const doc = await zip.file("word/document.xml").async("string");
  assert.match(doc, /<w:pStyle w:val="Heading1"\/>/);
  assert.match(doc, /<w:numPr><w:ilvl w:val="1"\/><w:numId w:val="1"\/><\/w:numPr>/);
  assert.match(doc, /<w:hyperlink r:id="rId\d+" w:history="1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"\/>/);
  assert.match(doc, /<w:hyperlink w:anchor="sec"/);
  assert.match(doc, /<w:gridSpan w:val="2"\/>/);
  assert.match(doc, /<w:br w:type="page"\/>/);
  assert.match(doc, /<w:highlight w:val="yellow"\/>/);
  assert.match(doc, /<w:shd w:val="clear" w:color="auto" w:fill="ABCDEF"\/>/);
  assert.match(doc, /<wp:extent cx="381000" cy="190500"\/>/);
  assert.match(doc, /<w:sectPr>/);
  const styles = await zip.file("word/styles.xml").async("string");
  for (const id of ["Heading1", "Heading6", "Quote", "Code", "Hyperlink", "ListParagraph", "TableGrid"]) {
    assert.match(styles, new RegExp(`w:styleId="${id}"`));
  }
  const numbering = await zip.file("word/numbering.xml").async("string");
  assert.match(numbering, /w:numFmt w:val="bullet"/);
  assert.match(numbering, /w:numFmt w:val="decimal"/);
  assert.match(numbering, /<w:startOverride w:val="1"\/>/);
  const rels = await zip.file("word/_rels/document.xml.rels").async("string");
  assert.match(rels, /Target="https:\/\/example.com\/a\?b=1&amp;c=2" TargetMode="External"/);
  // Every part is well-formed enough for our own parser and Word's strictness basics.
  for (const n of names.filter((x) => x.endsWith(".xml") || x.endsWith(".rels"))) {
    const s = await zip.file(n).async("string");
    assert.ok(s.startsWith("<?xml"), n);
    assert.doesNotMatch(s, /[\u0000-\u0008\u000B\u000C\u000E-\u001F]/, n);
  }
});

test("modelToDocx -> docxToModel round-trips the model", async () => {
  const bytes = await modelToDocx(RICH);
  const { model, warnings } = await docxToModel(bytes);
  assert.deepEqual(warnings, []);
  assert.equal(model.meta.title, "Rich & <Test>");
  assert.deepEqual(normalizeWordModel(model), normalizeWordModel(RICH));
});

test("modelToDocx strips XML-illegal characters and handles empty / table-final docs", async () => {
  const empty = await docxToModel(await modelToDocx({ kind: "word", blocks: [] }));
  assert.deepEqual(normalizeWordModel(empty.model).blocks, [{ type: "p", runs: [] }]);
  const bad = await docxToModel(await modelToDocx({ blocks: [{ type: "p", runs: [{ t: "a\u0001b\u0000c" }] }] }));
  assert.equal(bad.model.blocks[0].runs[0].t, "abc");
  const tbl = await docxToModel(await modelToDocx({ blocks: [{ type: "table", rows: [[{ blocks: [{ type: "p", runs: [{ t: "x" }] }] }]] }] }));
  assert.equal(tbl.model.blocks[0].type, "table");
  assert.deepEqual(tbl.model.blocks[1], { type: "p", runs: [] }); // Word needs a paragraph after a final table
});

test("modelToDocx keeps a placeholder for non-embeddable images", async () => {
  const { model } = await docxToModel(await modelToDocx({ blocks: [{ type: "img", src: "https://x.test/a.png", alt: "remote" }] }));
  assert.deepEqual(model.blocks[0].runs, [{ t: "[image: remote]", link: "https://x.test/a.png" }]);
});

test("ordered lists restart numbering per list (separate w:num instances)", async () => {
  const bytes = await modelToDocx({
    blocks: [
      { type: "list", list: "ol", level: 0, runs: [{ t: "a" }] },
      { type: "p", runs: [{ t: "gap" }] },
      { type: "list", list: "ol", level: 0, runs: [{ t: "b" }] },
    ],
  });
  const zip = await JSZip.loadAsync(bytes);
  const doc = await zip.file("word/document.xml").async("string");
  const ids = [...doc.matchAll(/<w:numId w:val="(\d+)"\/>/g)].map((m) => m[1]);
  assert.equal(ids.length, 2);
  assert.notEqual(ids[0], ids[1]);
});

// A hand-written package exercising reader features Word produces.
async function buildDocx(bodyXml, { styles = "", numbering = "", rels = "", extraFiles = {} } = {}) {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>');
  zip.file("_rels/.rels", '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="/word/document.xml"/></Relationships>');
  zip.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><w:body>${bodyXml}<w:sectPr/></w:body></w:document>`);
  const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
  zip.file("word/_rels/document.xml.rels", `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rSt" Type="${R}styles" Target="styles.xml"/><Relationship Id="rNum" Type="${R}numbering" Target="numbering.xml"/>${rels}</Relationships>`);
  zip.file("word/styles.xml", `<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:sz w:val="22"/></w:rPr></w:style>${styles}</w:styles>`);
  zip.file("word/numbering.xml", `<?xml version="1.0"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${numbering}</w:numbering>`);
  for (const [k, v] of Object.entries(extraFiles)) zip.file(k, v);
  return zip.generateAsync({ type: "uint8array" });
}

const P = (inner, ppr = "") => `<w:p>${ppr ? `<w:pPr>${ppr}</w:pPr>` : ""}${inner}</w:p>`;
const R = (t, rpr = "") => `<w:r>${rpr ? `<w:rPr>${rpr}</w:rPr>` : ""}<w:t xml:space="preserve">${t}</w:t></w:r>`;

test("docxToModel reads Word features", async () => {
  const R_ = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
  const bytes = await buildDocx(
    [
      P(R("Doc title"), '<w:pStyle w:val="Title"/>'),
      P(R("Outline heading"), '<w:pStyle w:val="MyHead"/>'),
      P(R("Custom bold style"), '<w:pStyle w:val="Emph"/>'),
      P(R("not bold", "<w:b w:val=\"0\"/>"), '<w:pStyle w:val="Emph"/>'),
      P(R("bullet via style"), '<w:pStyle w:val="ListBullet"/>'),
      P(R("numbered L1"), '<w:numPr><w:ilvl w:val="1"/><w:numId w:val="2"/></w:numPr>'),
      P(R("numbering removed"), '<w:pStyle w:val="ListBullet"/><w:numPr><w:numId w:val="0"/></w:numPr>'),
      P(
        `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> HYPERLINK "https://field.test/" </w:instrText></w:r>` +
        `<w:r><w:fldChar w:fldCharType="separate"/></w:r>${R("field link")}<w:r><w:fldChar w:fldCharType="end"/></w:r>` +
        `<w:hyperlink r:id="rL"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>rel link</w:t></w:r></w:hyperlink>` +
        `<w:hyperlink w:anchor="bm">${R("anchor")}</w:hyperlink>`,
        '<w:jc w:val="both"/>'
      ),
      P(
        R("hl", '<w:highlight w:val="green"/>') + R("shd", '<w:shd w:val="clear" w:fill="ABCDEF"/>') +
        R("hidden", "<w:vanish/>") + R("caps", '<w:rFonts w:ascii="Arial"/><w:sz w:val="28"/><w:color w:val="00FF00"/><w:u w:val="double"/><w:strike/><w:vertAlign w:val="superscript"/>') +
        `<w:ins><w:r><w:t>inserted</w:t></w:r></w:ins><w:del><w:r><w:delText>deleted</w:delText></w:r></w:del>` +
        `<w:sdt><w:sdtContent>${R("sdt")}</w:sdtContent></w:sdt><w:smartTag>${R("smart")}</w:smartTag>` +
        `<w:r><w:tab/><w:t>tab</w:t><w:br/><w:t>br</w:t><w:sym w:font="Symbol" w:char="F041"/></w:r>` +
        `<m:oMath><m:r><m:t>x+1</m:t></m:r></m:oMath>`
      ),
      P(`${R("before")}<w:r><w:br w:type="page"/></w:r>${R("after")}`),
      P(R("pbb"), "<w:pageBreakBefore/>"),
      `<w:tbl><w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr>${P(R("span"))}</w:tc></w:tr>` +
        `<w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr>${P(R("v"))}</w:tc><w:tc>${P(R("w"))}</w:tc></w:tr>` +
        `<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr>${P("")}</w:tc><w:tc><w:p/></w:tc></w:tr></w:tbl>`,
      `<w:sdt><w:sdtContent>${P(R("block sdt"))}</w:sdtContent></w:sdt>`,
      P(`<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="952500" cy="476250"/><wp:docPr id="1" name="p" descr="alt text"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:blipFill><a:blip r:embed="rImg"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice><mc:Fallback><w:t>fallback</w:t></mc:Fallback></mc:AlternateContent></w:r><w:r><w:footnoteReference w:id="1"/></w:r>`),
      "<w:p/>",
    ].join(""),
    {
      styles:
        '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/></w:style>' +
        '<w:style w:type="paragraph" w:styleId="MyHead"><w:name w:val="My Head"/><w:pPr><w:outlineLvl w:val="2"/></w:pPr></w:style>' +
        '<w:style w:type="paragraph" w:styleId="Emph"><w:name w:val="Emph"/><w:basedOn w:val="Normal"/><w:rPr><w:b/></w:rPr></w:style>' +
        '<w:style w:type="paragraph" w:styleId="ListBullet"><w:name w:val="List Bullet"/><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:style>' +
        '<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>',
      numbering:
        '<w:abstractNum w:abstractNumId="10"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/></w:lvl></w:abstractNum>' +
        '<w:abstractNum w:abstractNumId="11"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl><w:lvl w:ilvl="1"><w:numFmt w:val="lowerLetter"/></w:lvl></w:abstractNum>' +
        '<w:num w:numId="1"><w:abstractNumId w:val="10"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="11"/></w:num>',
      rels:
        `<Relationship Id="rL" Type="${R_}hyperlink" Target="https://rel.test/x" TargetMode="External"/>` +
        `<Relationship Id="rImg" Type="${R_}image" Target="media/pic.png"/>` +
        `<Relationship Id="rH" Type="${R_}header" Target="header1.xml"/>`,
      extraFiles: { "word/media/pic.png": Uint8Array.from(atob(PNG_B64), (c) => c.charCodeAt(0)) },
    }
  );
  const { model, warnings } = await docxToModel(bytes);
  const b = model.blocks;
  let k = 0;
  assert.deepEqual(b[k++], { type: "h1", runs: [{ t: "Doc title" }] });
  assert.deepEqual(b[k++], { type: "h3", runs: [{ t: "Outline heading" }] });
  assert.deepEqual(b[k++], { type: "p", runs: [{ t: "Custom bold style", b: true }] });
  assert.deepEqual(b[k++], { type: "p", runs: [{ t: "not bold" }] });
  assert.deepEqual(b[k++], { type: "list", list: "ul", level: 0, runs: [{ t: "bullet via style" }] });
  assert.deepEqual(b[k++], { type: "list", list: "ol", level: 1, runs: [{ t: "numbered L1" }] });
  assert.deepEqual(b[k++], { type: "p", runs: [{ t: "numbering removed" }] });
  assert.deepEqual(b[k++], {
    type: "p", align: "justify",
    runs: [{ t: "field link", link: "https://field.test/" }, { t: "rel link", link: "https://rel.test/x" }, { t: "anchor", link: "#bm" }],
  });
  assert.deepEqual(b[k++], {
    type: "p",
    runs: [
      { t: "hl", hl: "#00FF00" }, { t: "shd", hl: "#ABCDEF" },
      { t: "caps", u: true, s: true, sup: true, color: "#00FF00", size: 14, font: "Arial" },
      { t: "insertedsdtsmart\ttab\nbrAx+1" },
    ],
  });
  assert.deepEqual(b[k++], { type: "p", runs: [{ t: "before" }] });
  assert.deepEqual(b[k++], { type: "pagebreak" });
  assert.deepEqual(b[k++], { type: "p", runs: [{ t: "after" }] });
  assert.deepEqual(b[k++], { type: "pagebreak" });
  assert.deepEqual(b[k++], { type: "p", runs: [{ t: "pbb" }] });
  assert.deepEqual(b[k++], {
    type: "table",
    rows: [
      [{ blocks: [{ type: "p", runs: [{ t: "span" }] }], colspan: 2 }],
      [{ blocks: [{ type: "p", runs: [{ t: "v" }] }] }, { blocks: [{ type: "p", runs: [{ t: "w" }] }] }],
      [{ blocks: [{ type: "p", runs: [] }] }, { blocks: [{ type: "p", runs: [] }] }],
    ],
  });
  assert.deepEqual(b[k++], { type: "p", runs: [{ t: "block sdt" }] });
  assert.deepEqual(b[k++], { type: "img", src: PNG_URL, alt: "alt text", w: 100, h: 50 });
  assert.deepEqual(b[k++], { type: "p", runs: [] });
  assert.equal(b.length, k);
  for (const w of ["Tracked changes were accepted.", "Equations were imported as plain text.",
    "Vertically merged table cells were split.", "Footnotes and endnotes were not imported.",
    "Headers and footers were not imported."]) {
    assert.ok(warnings.includes(w), `missing warning: ${w} in ${JSON.stringify(warnings)}`);
  }
});

test("docxToModel rejects non-docx input with a clear error", async () => {
  await assert.rejects(() => docxToModel(new Uint8Array([1, 2, 3, 4])), /Not a valid \.docx/);
  await assert.rejects(() => docxToModel(new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0, 0, 0, 0])), /password-protected or a legacy/);
  const zip = new JSZip();
  zip.file("a.txt", "x");
  await assert.rejects(async () => docxToModel(await zip.generateAsync({ type: "uint8array" })), /main document part is missing/);
});

test("docxToModel reads the example quarterly report", async () => {
  const bytes = fs.readFileSync(path.join(EXT, "../example/quarterly-report.docx"));
  const { model } = await docxToModel(new Uint8Array(bytes));
  const types = model.blocks.map((b) => b.type);
  assert.equal(types[0], "h1");
  assert.ok(types.includes("table"));
  const link = model.blocks[1].runs.find((r) => r.link);
  assert.deepEqual(link, { t: "the dashboard", link: "https://example.com/dashboard" });
  assert.match(wordModelToText(model), /Regional breakdown/);
  // Literal-marker lists and NoSpacing+Consolas code written by older
  // backend versions are recovered, exactly like the backend reader does.
  assert.deepEqual(model.blocks.filter((b) => b.type === "list").map((b) => [b.list, b.level, b.runs[0].t]), [
    ["ul", 0, "Revenue up 12.4% YoY"],
    ["ul", 1, "Enterprise segment drove most of the gain"],
    ["ul", 0, "Churn reduced to 1.8%"],
    ["ol", 0, "Expand into APAC next quarter"],
    ["ol", 0, "Hire two platform engineers"],
  ]);
  assert.deepEqual(model.blocks.find((b) => b.type === "code").runs.map((r) => r.t).join(""), "npm run report:quarterly -- --quarter=Q3");
  assert.ok(!/•/.test(wordModelToText(model).replace(/^\s*• /gm, "")), "no literal bullets left in text runs");
});

test("docxToModel infers undefined style ids (Subtitle, ListBullet2, Strong)", async () => {
  const bytes = await buildDocx(
    P(R("sub"), '<w:pStyle w:val="Subtitle"/>') +
    P(R("bullet"), '<w:pStyle w:val="ListBullet2"/>') +
    P(`<w:r><w:rPr><w:rStyle w:val="Strong"/></w:rPr><w:t>strong</w:t></w:r>`) +
    P(R("normal explicitly"), '<w:pStyle w:val="Normal"/>')
  );
  const { model } = await docxToModel(bytes);
  assert.deepEqual(model.blocks, [
    { type: "h2", runs: [{ t: "sub" }] },
    { type: "list", list: "ul", level: 1, runs: [{ t: "bullet" }] },
    { type: "p", runs: [{ t: "strong", b: true }] },
    { type: "p", runs: [{ t: "normal explicitly" }] }, // Normal's sz is the baseline, not a run size
  ]);
});

test("wordModelToMarkdown never skips list levels (markdown cannot express it)", async () => {
  const { markdownToModel } = await import("../lib/markdown.js");
  const item = (level, t) => ({ type: "list", list: "ul", level, runs: [{ t }] });
  const md = wordModelToMarkdown({ kind: "word", blocks: [item(2, "a"), item(0, "b"), item(2, "c"), item(1, "d")] });
  assert.equal(md.trim(), "- a\n- b\n    - c\n    - d");
  assert.deepEqual(markdownToModel(md).blocks.map((b) => b.level), [0, 0, 1, 1]);
});

test("docxToModel/modelToDocx keep page size, orientation and margins (meta.page, backend shape)", async () => {
  const model = {
    kind: "word",
    meta: { page: { size: "A4", orientation: "landscape", margins: { top: 1000, right: 900, bottom: 1000, left: 900, header: 500, footer: 500 } } },
    blocks: [{ type: "p", runs: [{ t: "wide" }] }],
  };
  const bytes = await modelToDocx(model);
  const zip = await JSZip.loadAsync(bytes);
  const xml = await zip.file("word/document.xml").async("string");
  assert.match(xml, /<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"\/>/);
  assert.match(xml, /<w:pgMar w:top="1000" w:right="900" w:bottom="1000" w:left="900" w:header="500" w:footer="500"/);
  const { model: back } = await docxToModel(bytes);
  assert.deepEqual(back.meta.page, {
    size: "A4", orientation: "landscape", width: 16838, height: 11906,
    margins: { top: 1000, right: 900, bottom: 1000, left: 900, header: 500, footer: 500 },
  });
  // No meta.page: US Letter portrait with 1in margins, as before.
  const plain = await JSZip.loadAsync(await modelToDocx({ kind: "word", blocks: [{ type: "p", runs: [{ t: "x" }] }] }));
  assert.match(await plain.file("word/document.xml").async("string"), /<w:pgSz w:w="12240" w:h="15840"\/>/);
});
