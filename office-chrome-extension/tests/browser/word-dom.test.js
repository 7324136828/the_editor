/**
 * In-browser tests for parseWordDom / renderWordHtml (lib/model-word.js) and
 * the browser-only parts of lib/local-docx.js. Loaded by word-dom.html;
 * driven by run.mjs (headless Edge/Chrome). Results are published on
 * window.__results / window.__done and rendered into #results, so the page
 * can also be opened by hand from a local web server.
 */

import { renderWordHtml, parseWordDom, normalizeWordModel, wordModelToText } from "../../lib/model-word.js";
import { docxToModel, modelToDocx, renderDocxPreview } from "../../lib/local-docx.js";

const ed = document.getElementById("ed");
const results = [];

async function run(name, fn) {
  try {
    await fn();
    results.push({ name, ok: true });
  } catch (e) {
    results.push({ name, ok: false, detail: `${e && e.message ? e.message : e}\n  html: ${ed.innerHTML.slice(0, 600)}` });
  }
}

function eq(actual, expected, msg = "mismatch") {
  const a = JSON.stringify(actual);
  const b = JSON.stringify(expected);
  if (a !== b) throw new Error(`${msg}\n  actual:   ${a}\n  expected: ${b}`);
}

function ok(cond, msg) {
  if (!cond) throw new Error(msg);
}

const N = (blocks) => normalizeWordModel({ blocks }).blocks;
const parsed = () => normalizeWordModel(parseWordDom(ed)).blocks;
const p = (t, extra = {}) => ({ type: "p", runs: t ? [{ t }] : [], ...extra });
const li = (t, level = 0, list = "ul") => ({ type: "list", list, level, runs: t ? [{ t }] : [] });
const textOf = (blocks) => wordModelToText({ blocks });

function setup(separator = "p") {
  document.execCommand("styleWithCSS", false, true);
  document.execCommand("defaultParagraphSeparator", false, separator);
}

function load(modelOrHtml) {
  ed.innerHTML = typeof modelOrHtml === "string" ? modelOrHtml : renderWordHtml(modelOrHtml);
  ed.focus();
}

function exec(cmd, val = null) {
  const r = document.execCommand(cmd, false, val);
  if (!r) throw new Error(`execCommand(${cmd}) returned false`);
}

function findText(s) {
  const w = document.createTreeWalker(ed, NodeFilter.SHOW_TEXT);
  for (let n = w.nextNode(); n; n = w.nextNode()) {
    const i = n.data.indexOf(s);
    if (i >= 0) return { node: n, index: i };
  }
  throw new Error(`text not found: ${s}`);
}

function setRange(r) {
  const sel = getSelection();
  sel.removeAllRanges();
  sel.addRange(r);
}

/** Caret inside the first text node containing s, `offset` chars into s. */
function caret(s, offset = 0) {
  const { node, index } = findText(s);
  const r = document.createRange();
  r.setStart(node, index + offset);
  r.collapse(true);
  setRange(r);
}

/** Select the substring s (within one text node). */
function select(s) {
  const { node, index } = findText(s);
  const r = document.createRange();
  r.setStart(node, index);
  r.setEnd(node, index + s.length);
  setRange(r);
}

/** Select from the start of a to the end of b (may span nodes). */
function selectSpan(a, b) {
  const A = findText(a);
  const B = findText(b);
  const r = document.createRange();
  r.setStart(A.node, A.index);
  r.setEnd(B.node, B.index + b.length);
  setRange(r);
}

function caretEnd() {
  const r = document.createRange();
  r.selectNodeContents(ed);
  r.collapse(false);
  setRange(r);
}

// A 2x1 PNG.
const PNG = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAIAAAB7QOjdAAAAD0lEQVR4nGP4z8DAwPAfAAcAAf9+CLHQAAAAAElFTkSuQmCC";

// ---------------------------------------------------------------------------
// 1. parse(render(model)) is stable
// ---------------------------------------------------------------------------

const STABLE = {
  "every run mark": [
    {
      type: "p",
      runs: [
        { t: "b", b: true }, { t: "i", i: true }, { t: "u", u: true }, { t: "s", s: true }, { t: "us", u: true, s: true },
        { t: "sup", sup: true }, { t: "sub", sub: true }, { t: "red", color: "#FF0000" }, { t: "black", color: "#000000" },
        { t: "hl", hl: "#FFFF00" }, { t: "size", size: 10.5 }, { t: "font", font: "Times New Roman" },
        { t: "all", b: true, i: true, u: true, color: "#00AA00", hl: "#ABCDEF", size: 14, font: "Arial" },
        { t: "link", link: "https://example.com/?a=1&b=2" }, { t: "red link", link: "https://x.test", color: "#FF0000", u: true },
        { t: "anchor", link: "#bm" }, { t: "js", link: "javascript:void(0)" },
      ],
    },
  ],
  "block types and alignment": [
    { type: "h1", runs: [{ t: "H1" }] }, { type: "h2", align: "center", runs: [{ t: "H2" }] }, { type: "h3", runs: [{ t: "H3" }] },
    { type: "h4", runs: [{ t: "H4" }] }, { type: "h5", runs: [{ t: "H5" }] }, { type: "h6", runs: [{ t: "H6" }] },
    { type: "p", align: "right", runs: [{ t: "right" }] }, { type: "p", align: "justify", runs: [{ t: "justify" }] },
    { type: "quote", runs: [{ t: "quote" }] }, { type: "quote", align: "center", runs: [{ t: "centered quote" }] },
    { type: "code", runs: [{ t: "line 1\n\n  indented\tTab" }] },
    { type: "p", runs: [] }, { type: "p", runs: [] }, { type: "pagebreak" }, { type: "pagebreak" }, { type: "p", runs: [{ t: "end" }] },
  ],
  "whitespace": [
    p(" leading"), p("trailing "), p("a  b   c"), p("tab\there"), p("line\nbreak"), p("ends with newline\n"),
    p("\n starts with newline"), { type: "p", runs: [{ t: "x " }, { t: " y", b: true }] },
  ],
  "nested and mixed lists": [
    li("Parent"), li("Child", 1), li("Grandchild", 2, "ol"), li("Next"),
    p("between"),
    li("A", 0, "ol"), li("B", 0, "ol"), li("C", 1), li("D", 2, "ol"), li("E", 0, "ol"),
    li("type change", 0, "ul"), li("", 0), li("after empty"),
    { type: "list", list: "ul", level: 0, runs: [{ t: "linked", link: "https://l.test" }, { t: " bold", b: true }] },
  ],
  "skipped list levels": [li("deep", 2), li("shallow", 0), li("", 0), li("x", 1)],
  "tables": [
    {
      type: "table",
      rows: [
        [{ blocks: [p("wide")], colspan: 3 }],
        [{ blocks: [p("a")] }, { blocks: [p("")] }, { blocks: [li("in cell"), li("nested", 1)] }],
        [{ blocks: [p("two"), p("paras")] }, { blocks: [{ type: "table", rows: [[{ blocks: [p("inner")] }]] }] }, { blocks: [{ type: "img", src: PNG, w: 4, h: 2 }] }],
      ],
    },
    p("after table"),
  ],
  "images": [
    { type: "img", src: PNG, alt: "two px", w: 20, h: 10 }, { type: "img", src: PNG },
    li("item"), { type: "img", src: PNG, alt: "between items" }, li("item 2"),
  ],
};

async function stabilityTests() {
  for (const [name, blocks] of Object.entries(STABLE)) {
    await run(`stable: ${name}`, () => {
      load({ blocks });
      eq(parsed(), N(blocks));
      // And rendering the parsed model again gives the same HTML.
      const html1 = ed.innerHTML;
      load(parseWordDom(ed));
      eq(ed.innerHTML, html1, "second render differs");
    });
  }
}

// ---------------------------------------------------------------------------
// 2. editing with execCommand (what the viewer toolbar does)
// ---------------------------------------------------------------------------

async function editingTests() {
  setup("p");
  const list3 = { blocks: [li("One"), li("Two"), li("Three")] };

  await run("indent list item -> nested level, nothing lost (WO-01)", () => {
    load(list3);
    caret("Two");
    exec("indent");
    eq(parsed(), N([li("One"), li("Two", 1), li("Three")]));
  });

  await run("indent twice then outdent", () => {
    load(list3);
    caret("Two");
    exec("indent");
    caret("Three");
    exec("indent");
    exec("indent");
    eq(parsed(), N([li("One"), li("Two", 1), li("Three", 2)]));
    exec("outdent");
    eq(parsed(), N([li("One"), li("Two", 1), li("Three", 1)]));
  });

  await run("outdent first list item -> paragraph (WO-02)", () => {
    load(list3);
    caret("One");
    exec("outdent");
    eq(parsed(), N([p("One"), li("Two"), li("Three")]));
  });

  await run("outdent last list item -> paragraph (WO-02)", () => {
    load(list3);
    caret("Three");
    exec("outdent");
    eq(parsed(), N([li("One"), li("Two"), p("Three")]));
  });

  await run("outdent nested item", () => {
    load({ blocks: [li("One"), li("Two", 1), li("Three")] });
    caret("Two");
    exec("outdent");
    eq(parsed(), N([li("One"), li("Two"), li("Three")]));
  });

  await run("toggle bullet off on middle item (WO-02)", () => {
    load(list3);
    caret("Two");
    exec("insertUnorderedList");
    eq(parsed(), N([li("One"), p("Two"), li("Three")]));
  });

  await run("bulleted list on a paragraph (WO-03/RU-10)", () => {
    load({ blocks: [p("Before"), p("Make me a bullet"), p("After")] });
    caret("Make me");
    exec("insertUnorderedList");
    eq(parsed(), N([p("Before"), li("Make me a bullet"), p("After")]));
    // Enter + typing adds a second item instead of merging text.
    caret("bullet", 6);
    exec("insertParagraph");
    exec("insertText", "second item");
    eq(parsed(), N([p("Before"), li("Make me a bullet"), li("second item"), p("After")]));
  });

  await run("numbered list over two selected paragraphs (WO-03)", () => {
    load({ blocks: [p("A"), p("B"), p("C")] });
    selectSpan("A", "B");
    exec("insertOrderedList");
    eq(parsed(), N([li("A", 0, "ol"), li("B", 0, "ol"), p("C")]));
  });

  await run("numbered list on a heading keeps the text as an item (WO-03)", () => {
    load({ blocks: [{ type: "h1", runs: [{ t: "Heading" }] }, p("x")] });
    caret("Heading");
    exec("insertOrderedList");
    const b = parsed();
    eq(textOf(b), "1. Heading\nx");
    eq(b[0].type, "list");
  });

  await run("Enter in a list adds an item", () => {
    load({ blocks: [li("first")] });
    caret("first", 5);
    exec("insertParagraph");
    exec("insertText", "second");
    eq(parsed(), N([li("first"), li("second")]));
  });

  await run("typing into an empty document (WO-02)", () => {
    load({ blocks: [] });
    exec("insertText", "Hello world");
    eq(parsed(), N([p("Hello world")]));
    exec("insertParagraph");
    exec("insertText", "Second");
    eq(parsed(), N([p("Hello world"), p("Second")]));
  });

  await run("typing into an empty document with <div> separator (WO-02)", () => {
    setup("div");
    try {
      load("");
      exec("insertText", "Hello world");
      exec("insertParagraph");
      exec("insertText", "Second");
      eq(parsed(), N([p("Hello world"), p("Second")]));
    } finally {
      setup("p");
    }
  });

  await run("select all + delete + type in a doc starting with a list (WO-01)", () => {
    load({ blocks: [li("One"), li("Two"), p("Para")] });
    exec("selectAll");
    exec("delete");
    exec("insertText", "New");
    const b = parsed();
    eq(b.length, 1, "one block");
    eq(textOf(b).replace(/^• /, ""), "New");
  });

  await run("select all + delete + type in a doc ending with a table (WO-02)", () => {
    load({ blocks: [p("Intro"), { type: "table", rows: [[{ blocks: [p("cell")] }]] }] });
    exec("selectAll");
    exec("delete");
    exec("insertText", "New");
    const b = parsed();
    ok(textOf(b).includes("New"), `text lost: ${JSON.stringify(b)}`);
  });

  await run("insert table and type into cells (RU-10)", () => {
    load({ blocks: [p("x")] });
    caretEnd();
    let html = '<table class="ow-table"><tbody>';
    for (let r = 0; r < 2; r++) html += "<tr><td><br></td><td><br></td></tr>";
    exec("insertHTML", html + "</tbody></table>");
    const tds = ed.querySelectorAll("td");
    const r = document.createRange();
    r.setStart(tds[0], 0);
    r.collapse(true);
    setRange(r);
    exec("insertText", "cellA");
    const r2 = document.createRange();
    r2.setStart(tds[3], 0);
    r2.collapse(true);
    setRange(r2);
    exec("insertText", "cellD");
    const t = parsed().find((b) => b.type === "table");
    ok(t, "table lost");
    eq(t.rows[0][0].blocks, N([p("cellA")]));
    eq(t.rows[1][1].blocks, N([p("cellD")]));
    eq(t.rows[0][1].blocks, N([p("")]));
  });

  await run("foreColor (styleWithCSS) -> color", () => {
    load({ blocks: [p("make red here")] });
    select("red");
    exec("foreColor", "#ff0000");
    eq(parsed(), N([{ type: "p", runs: [{ t: "make " }, { t: "red", color: "#FF0000" }, { t: " here" }] }]));
  });

  await run("foreColor black is kept (WO-13)", () => {
    load({ blocks: [p("make black")] });
    select("black");
    exec("foreColor", "#000000");
    eq(parsed(), N([{ type: "p", runs: [{ t: "make " }, { t: "black", color: "#000000" }] }]));
  });

  await run("hiliteColor -> hl", () => {
    load({ blocks: [p("mark this")] });
    select("this");
    exec("hiliteColor", "#ffff00");
    eq(parsed(), N([{ type: "p", runs: [{ t: "mark " }, { t: "this", hl: "#FFFF00" }] }]));
  });

  await run("fontSize 5 -> 18pt", () => {
    load({ blocks: [p("big text")] });
    select("big");
    exec("fontSize", "5");
    eq(parsed(), N([{ type: "p", runs: [{ t: "big", size: 18 }, { t: " text" }] }]));
  });

  await run("fontName -> font", () => {
    load({ blocks: [p("arial text")] });
    select("arial");
    exec("fontName", "Arial");
    eq(parsed(), N([{ type: "p", runs: [{ t: "arial", font: "Arial" }, { t: " text" }] }]));
  });

  await run("bold / italic / underline / strike / superscript toggles", () => {
    load({ blocks: [p("one two three four five")] });
    select("one");
    exec("bold");
    select("two");
    exec("italic");
    select("three");
    exec("underline");
    select("four");
    exec("strikeThrough");
    select("five");
    exec("superscript");
    eq(parsed(), N([{
      type: "p",
      runs: [{ t: "one", b: true }, { t: " " }, { t: "two", i: true }, { t: " " }, { t: "three", u: true }, { t: " " },
        { t: "four", s: true }, { t: " " }, { t: "five", sup: true }],
    }]));
    select("one");
    exec("bold");
    ok(!parsed()[0].runs.some((r) => r.b), "bold not removed");
  });

  await run("justifyRight replaces an existing alignment (WO-08)", () => {
    load({ blocks: [p("aligned", { align: "center" })] });
    caret("aligned");
    exec("justifyRight");
    eq(parsed(), N([p("aligned", { align: "right" })]));
    exec("justifyLeft");
    eq(parsed(), N([p("aligned")]));
    exec("justifyFull");
    eq(parsed(), N([p("aligned", { align: "justify" })]));
  });

  await run("indent on a paragraph keeps it a paragraph (WO-09)", () => {
    load({ blocks: [p("Para"), p("Other")] });
    caret("Para");
    exec("indent");
    eq(parsed(), N([p("Para"), p("Other")]));
    exec("outdent");
    eq(parsed(), N([p("Para"), p("Other")]));
  });

  await run("formatBlock: quote, heading, back to paragraph", () => {
    load({ blocks: [p("text")] });
    caret("text");
    exec("formatBlock", "<blockquote>");
    eq(parsed(), N([{ type: "quote", runs: [{ t: "text" }] }]));
    exec("formatBlock", "<h2>");
    eq(parsed(), N([{ type: "h2", runs: [{ t: "text" }] }]));
    exec("formatBlock", "<p>");
    eq(parsed(), N([p("text")]));
    exec("formatBlock", "<pre>");
    eq(parsed(), N([{ type: "code", runs: [{ t: "text" }] }]));
  });

  await run("createLink then unlink leaves plain text (WO-13/RU-20)", () => {
    load({ blocks: [p("go there now")] });
    select("there");
    exec("createLink", "https://example.com/");
    eq(parsed(), N([{ type: "p", runs: [{ t: "go " }, { t: "there", link: "https://example.com/" }, { t: " now" }] }]));
    select("there");
    exec("unlink");
    eq(parsed(), N([p("go there now")]));
  });

  await run("unchanged link round-trips without color/underline (RU-20)", () => {
    const blocks = [{ type: "p", runs: [{ t: "See " }, { t: "the dashboard", link: "https://example.com/dashboard" }] }];
    load({ blocks });
    eq(parsed(), N(blocks));
    ok(!/style=/.test(ed.querySelector("a").outerHTML), "inline style on <a>");
  });

  await run("insertImage inside a paragraph splits it (WO-13)", () => {
    load({ blocks: [p("XY")] });
    caret("XY", 1);
    exec("insertImage", PNG);
    eq(parsed(), N([p("X"), { type: "img", src: PNG }, p("Y")]));
  });

  await run("insert page break", () => {
    load({ blocks: [p("before")] });
    caretEnd();
    exec("insertHTML", '<hr class="ow-pagebreak">');
    exec("insertText", "after");
    const b = parsed();
    eq(b.map((x) => x.type), ["p", "pagebreak", "p"]);
    eq(textOf(b).replace("\f", "|"), "before\n|\nafter");
  });

  await run("typing double spaces keeps them", () => {
    load({ blocks: [p("")] });
    // One character at a time, like real typing (Chrome inserts &nbsp; as
    // needed; a multi-space insertText string is inserted raw and collapses).
    for (const ch of "a  b  c") exec("insertText", ch);
    eq(parsed(), N([p("a  b  c")]));
  });

  await run("Backspace merging paragraph into heading adds no junk formatting", () => {
    load({ blocks: [{ type: "h1", runs: [{ t: "Title" }] }, p("Body")] });
    caret("Body");
    exec("delete");
    const b = parsed();
    eq(b.length, 1);
    eq(b[0].type, "h1");
    eq(textOf(b), "TitleBody");
    ok(!b[0].runs.some((r) => r.color || r.font), `junk marks: ${JSON.stringify(b[0].runs)}`);
  });

  await run("Backspace merging two styled paragraphs keeps marks", () => {
    load({ blocks: [{ type: "p", runs: [{ t: "red", color: "#FF0000" }] }, { type: "p", runs: [{ t: "bold", b: true }] }] });
    caret("bold");
    exec("delete");
    eq(parsed(), N([{ type: "p", runs: [{ t: "red", color: "#FF0000" }, { t: "bold", b: true }] }]));
  });
}

// ---------------------------------------------------------------------------
// 3. DOM shapes from paste / other editors
// ---------------------------------------------------------------------------

function dom(build) {
  ed.innerHTML = "";
  build(ed);
}

function el(tag, attrs = {}, ...children) {
  const e = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) e.setAttribute(k, v);
  for (const c of children) e.append(typeof c === "string" ? document.createTextNode(c) : c);
  return e;
}

async function domTests() {
  await run("<p> containing a list (built via DOM, as execCommand makes it)", () => {
    dom((root) => root.append(el("p", {}, "intro", el("ol", {}, el("li", {}, "x"), el("li", {}, "y")), "outro")));
    eq(parsed(), N([p("intro"), li("x", 0, "ol"), li("y", 0, "ol"), p("outro")]));
  });

  await run("nested list inside <li> with own text", () => {
    load("<ul><li>a<ul><li>b</li></ul></li><li>c</li></ul>");
    eq(parsed(), N([li("a"), li("b", 1), li("c")]));
  });

  await run("Chrome indent shape: list directly in list", () => {
    load("<ol><li>One</li><ol><li>Two</li><ul><li>Three</li></ul></ol><li>Four</li></ol>");
    eq(parsed(), N([li("One", 0, "ol"), li("Two", 1, "ol"), li("Three", 2), li("Four", 0, "ol")]));
  });

  await run("<li> wrapping only a nested list has no bullet of its own", () => {
    load("<ul><li><ul><li>deep</li></ul></li></ul>");
    eq(parsed(), N([li("deep", 1)]));
  });

  await run("paragraphs inside <li> join into one item", () => {
    load("<ul>\n  <li>\n    <p>para in li</p>\n    <p>second</p>\n  </li>\n</ul>");
    eq(parsed(), N([li("para in li\nsecond")]));
  });

  await run("stray text directly inside <ul>", () => {
    dom((root) => root.append(el("ul", {}, el("li", {}, "a"), "stray", el("li", {}, "b"))));
    eq(parsed(), N([li("a"), li("stray"), li("b")]));
  });

  await run("image inside list item goes after the item", () => {
    load(`<ul><li>text <img src="${PNG}"> more</li></ul>`);
    eq(parsed(), N([li("text more"), { type: "img", src: PNG }]));
  });

  await run("<font> color/face/size (WO-13)", () => {
    load('<p><font color="#ff0000" face="Arial, sans-serif" size="5">red</font> plain</p>');
    eq(parsed(), N([{ type: "p", runs: [{ t: "red", color: "#FF0000", font: "Arial", size: 18 }, { t: " plain" }] }]));
  });

  await run("styled spans: px size, font list, decoration line, background", () => {
    load('<p><span style="font-size: 16px; font-family: &quot;Courier New&quot;, monospace">a</span>' +
      '<span style="text-decoration-line: underline line-through">b</span>' +
      '<span style="background-color: rgb(0, 255, 0)">c</span>' +
      '<span style="font-weight: 700; font-style: italic">d</span>' +
      '<span style="vertical-align: sub">e</span></p>');
    eq(parsed(), N([{
      type: "p",
      runs: [{ t: "a", size: 12, font: "Courier New" }, { t: "b", u: true, s: true }, { t: "c", hl: "#00FF00" },
        { t: "d", b: true, i: true }, { t: "e", sub: true }],
    }]));
  });

  await run("Google Docs style <b style=font-weight:normal> wrapper", () => {
    load('<b style="font-weight:normal"><p>plain <span style="font-weight:700">bold</span></p></b>');
    eq(parsed(), N([{ type: "p", runs: [{ t: "plain " }, { t: "bold", b: true }] }]));
  });

  await run("<mark>, <code>, <strong>, <em>, <del>", () => {
    load("<p><mark>m</mark><code>c</code><strong>s</strong><em>e</em><del>d</del></p>");
    eq(parsed(), N([{
      type: "p",
      runs: [{ t: "m", hl: "#FFFF00" }, { t: "c", font: "Consolas" }, { t: "s", b: true }, { t: "e", i: true }, { t: "d", s: true }],
    }]));
  });

  await run("top-level text and <br> lines", () => {
    load("Hello<br>World<br><br>after blank");
    eq(parsed(), N([p("Hello"), p("World"), p(""), p("after blank")]));
  });

  await run("<div> lines and empty <div><br></div>", () => {
    load("<div>x</div><div><br></div><div>y</div>");
    eq(parsed(), N([p("x"), p(""), p("y")]));
  });

  await run("<br> inside <p> is a soft line break", () => {
    load("<p>a<br>b</p>");
    eq(parsed(), N([p("a\nb")]));
  });

  await run("source whitespace collapses; nbsp becomes a space", () => {
    load("<p>\n   Hello\n   world  </p>\n\n<p>a&nbsp; b</p>");
    eq(parsed(), N([p("Hello world"), p("a  b")]));
  });

  await run("tables: th is bold, bare text in td, colspan", () => {
    load("<table><thead><tr><th colspan=2>H</th></tr></thead><tbody><tr><td>bare</td><td><div>d1</div><div>d2</div></td></tr></tbody></table>");
    eq(parsed(), N([{
      type: "table",
      rows: [[{ blocks: [{ type: "p", runs: [{ t: "H", b: true }] }], colspan: 2 }], [{ blocks: [p("bare")] }, { blocks: [p("d1"), p("d2")] }]],
    }]));
  });

  await run("pasted <blockquote> is a quote; indent blockquote is not (WO-09)", () => {
    load('<blockquote>pasted quote</blockquote><blockquote style="margin: 0 0 0 40px; border: none; padding: 0px;"><p>Para</p></blockquote>');
    eq(parsed(), N([{ type: "quote", runs: [{ t: "pasted quote" }] }, p("Para")]));
  });

  await run("legacy data-align loses to inline text-align (WO-08)", () => {
    load('<p data-align="center" style="text-align: right;">x</p><p data-align="center">y</p><p align="justify">z</p><div style="text-align:center"><p>inherited</p></div>');
    eq(parsed(), N([p("x", { align: "right" }), p("y", { align: "center" }), p("z", { align: "justify" }), p("inherited", { align: "center" })]));
  });

  await run("legacy rendered link colour is not written into the model (RU-20)", () => {
    load('<p><a href="https://x.test" target="_blank" style="color:#1a5fb4;text-decoration:underline">go</a></p>');
    eq(parsed(), N([{ type: "p", runs: [{ t: "go", link: "https://x.test" }] }]));
  });

  await run("page breaks: hr class, CSS break-before; plain hr ignored", () => {
    load('<p>a</p><hr><p>b</p><hr class="ow-pagebreak"><div style="page-break-before: always">c</div>');
    eq(parsed(), N([p("a"), p("b"), { type: "pagebreak" }, { type: "pagebreak" }, p("c")]));
  });

  await run("scripts, styles and hidden elements are ignored", () => {
    load('<p>a<script>x()</script><style>p{}</style><span style="display:none">hidden</span>b</p>');
    eq(parsed(), N([p("ab")]));
  });

  await run("meta is carried over when requested", () => {
    load({ blocks: [p("x")] });
    const m = parseWordDom(ed, { meta: { title: "T", warnings: ["w"] } });
    eq(m.meta, { title: "T" });
  });
}

// ---------------------------------------------------------------------------
// 4. local-docx in the browser
// ---------------------------------------------------------------------------

async function docxTests() {
  const example = new Uint8Array(await (await fetch("../../../example/quarterly-report.docx")).arrayBuffer());

  await run("renderDocxPreview applies document styles (WO-16)", async () => {
    const box = document.getElementById("preview");
    box.innerHTML = "";
    await renderDocxPreview(example, box);
    const styles = box.querySelectorAll("style");
    ok(styles.length > 0, "no <style> elements in the preview container");
    ok([...styles].every((s) => s.isConnected && s.sheet), "style elements not attached");
    const wrapper = box.querySelector(".docx-wrapper");
    ok(wrapper, "no .docx-wrapper");
    ok(getComputedStyle(wrapper).display === "flex", "docx-preview wrapper CSS not applied");
    const section = box.querySelector("section.docx");
    ok(section && getComputedStyle(section).backgroundColor === "rgb(255, 255, 255)", "page CSS not applied");
  });

  await run("docxToModel reads the example document", async () => {
    const { model } = await docxToModel(example);
    eq(model.blocks[0].type, "h1");
    ok(model.blocks.some((b) => b.type === "table"), "no table");
  });

  await run("modelToDocx output renders in docx-preview", async () => {
    const model = {
      kind: "word",
      blocks: [
        { type: "h1", runs: [{ t: "Generated" }] }, li("bullet"), li("sub", 1), li("num", 0, "ol"),
        { type: "table", rows: [[{ blocks: [p("cell")], colspan: 2 }], [{ blocks: [p("a")] }, { blocks: [p("b")] }]] },
        { type: "img", src: PNG, w: 20, h: 10 },
      ],
    };
    const bytes = await modelToDocx(model);
    const box = document.getElementById("preview");
    box.innerHTML = "";
    await renderDocxPreview(bytes, box);
    const text = box.innerText;
    for (const s of ["Generated", "bullet", "sub", "num", "cell"]) ok(text.includes(s), `missing ${s}`);
    ok(box.querySelector("img"), "image not rendered");
    const back = await docxToModel(bytes);
    eq(normalizeWordModel(back.model).blocks, N(model.blocks));
  });

  await run("modelToDocx rasterizes SVG images to PNG", async () => {
    const svg = "data:image/svg+xml;base64," + btoa('<svg xmlns="http://www.w3.org/2000/svg" width="30" height="10"><rect width="30" height="10" fill="red"/></svg>');
    const bytes = await modelToDocx({ blocks: [{ type: "img", src: svg }] });
    const { model } = await docxToModel(bytes);
    eq(model.blocks[0].type, "img");
    ok(model.blocks[0].src.startsWith("data:image/png;base64,"), "not png");
    eq([model.blocks[0].w, model.blocks[0].h], [30, 10]);
  });

  await run("edit in DOM -> model -> docx -> model keeps the edit", async () => {
    setup("p");
    load({ blocks: [li("One"), li("Two")] });
    caret("Two");
    exec("indent");
    const edited = parseWordDom(ed);
    const back = await docxToModel(await modelToDocx(edited));
    eq(normalizeWordModel(back.model).blocks, N([li("One"), li("Two", 1)]));
  });
}

// ---------------------------------------------------------------------------

(async () => {
  try {
    await stabilityTests();
    await editingTests();
    await domTests();
    await docxTests();
  } catch (e) {
    results.push({ name: "harness", ok: false, detail: String(e && e.stack ? e.stack : e) });
  }
  window.__results = results;
  window.__done = true;
  const out = document.getElementById("results");
  out.innerHTML = results
    .map((r) => `<div class="${r.ok ? "pass" : "fail"}">${r.ok ? "PASS" : "FAIL"} ${r.name}${r.ok ? "" : "\n" + r.detail.replace(/</g, "&lt;")}</div>`)
    .join("");
})();
