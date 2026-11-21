/**
 * Slides (PowerPoint) tests that run under plain Node (`npm test`):
 * model helpers (lib/model-slides.js), the built-in .pptx reader/writer
 * (lib/local-pptx.js) on python-pptx generated fixtures, a full round trip of
 * every contract field, package consistency of written files, and plain-language
 * errors for bad input.
 *
 * When a Python with python-pptx is available (env PYTHON, e.g. the project's
 * venv), written files are also opened with python-pptx; otherwise that test is
 * skipped. Regenerate fixtures with tests/fixtures/slides/make_fixtures.py.
 */

import { test } from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";

import {
  newDeck, newSlide, normalizeDeck, slidesToMarkdown, slidesToText, slidesToHtmlDocument, renderSlideHtml,
  TEMPLATES, LAYOUTS, SHAPES, shapePath, resizeDeck, defaultRunSize, textColorFor, hexColor, cloneDeep, slideTitle,
} from "../lib/model-slides.js";
import { pptxToModel, modelToPptx, parseXml, xmlEsc } from "../lib/local-pptx.js";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const EXT = path.resolve(HERE, "..");
const FIX = path.join(HERE, "fixtures/slides");

if (!globalThis.JSZip) {
  const code = fs.readFileSync(path.join(EXT, "vendor/jszip.min.js"), "utf8");
  const mod = { exports: {} };
  new Function("module", "exports", code)(mod, mod.exports);
  globalThis.JSZip = mod.exports;
}
const JSZip = globalThis.JSZip;

const PNG_B64 = "iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAIAAAB7QOjdAAAAD0lEQVR4nGP4z8DAwPAfAAcAAf9+CLHQAAAAAElFTkSuQmCC";
const PNG_URL = `data:image/png;base64,${PNG_B64}`;
const text = (paras) => (paras || []).map((p) => (p.runs || []).map((r) => r.t).join("")).join("\n");
const readFix = (name) => new Uint8Array(fs.readFileSync(path.join(FIX, name)));

// ---------------------------------------------------------------------------
// model helpers
// ---------------------------------------------------------------------------

test("newDeck: one empty title slide, 16:9, Blank theme", () => {
  const d = newDeck();
  assert.equal(d.kind, "slides");
  assert.deepEqual(d.size, { w: 960, h: 540 });
  assert.equal(d.slides.length, 1);
  const s = d.slides[0];
  assert.equal(s.layout, "title");
  assert.deepEqual(s.items.map((i) => i.ph), ["title", "subtitle"]);
  assert.equal(text(s.items[0].paras), "");
  assert.equal(d.theme.font, TEMPLATES[0].theme.font);
});

test("newDeck templates: by id or name; 4:3 variant", () => {
  assert.equal(newDeck("dark").theme.bg, TEMPLATES.find((t) => t.id === "dark").theme.bg);
  assert.equal(newDeck("Modern blue").theme.accent, "#1A56DB");
  assert.equal(newDeck("nonsense").theme.font, TEMPLATES[0].theme.font);
  assert.deepEqual(newDeck("standard 4:3").size, { w: 720, h: 540 });
  for (const t of TEMPLATES) for (const k of ["font", "headFont", "color", "accent", "bg"]) assert.ok(t.theme[k], `${t.id}.${k}`);
});

test("newSlide: every layout puts placeholders inside the slide (16:9 and 4:3)", () => {
  for (const size of [{ w: 960, h: 540 }, { w: 720, h: 540 }]) {
    const deck = { size };
    for (const l of LAYOUTS) {
      const s = newSlide(l.id, deck);
      assert.equal(s.layout, l.id);
      for (const it of s.items) {
        assert.ok(it.x >= 0 && it.y >= 0 && it.x + it.w <= size.w + 0.01 && it.y + it.h <= size.h + 0.01, `${l.id} ${it.ph}`);
        assert.ok(it.paras[0].runs[0].size > 0);
      }
    }
    assert.equal(newSlide("two", deck).items.filter((i) => i.ph === "body").length, 2);
    assert.equal(newSlide("blank", deck).items.length, 0);
    assert.equal(newSlide("content", deck).items[1].paras[0].bullet, "ul");
    assert.equal(newSlide("bogus", deck).layout, "content");
  }
});

test("normalizeDeck: never throws, fixes bad values, keeps unknown items", () => {
  for (const bad of [null, undefined, 5, "x", [], { slides: "nope" }, { size: { w: -1, h: "a" } }, { slides: [null, 3, { items: "x" }] }]) {
    const d = normalizeDeck(bad);
    assert.equal(d.kind, "slides");
    assert.ok(d.slides.length >= 1);
    assert.ok(d.size.w > 0 && d.size.h > 0);
  }
  const d = normalizeDeck({
    theme: { color: "abc", accent: "red", bg: "#00ff00" },
    slides: [{ layout: "weird", bg: "zz", transition: "spin", items: [
      { type: "text", x: "10", y: null, w: -50, h: 20, paras: [{ runs: [{ t: 5, color: "#f00", size: "12" }], align: "middle", level: 12, bullet: "x" }], rot: -90 },
      { type: "shape", shape: "blob", x: 0, y: 0, w: 1, h: 1 },
      { type: "chart", x: 1, y: 2, w: 3, h: 4, data: [1, 2] },
      null,
    ] }],
  });
  assert.equal(d.theme.color, "#AABBCC");
  assert.equal(d.theme.accent, undefined);
  assert.equal(d.theme.bg, "#00FF00");
  const s = d.slides[0];
  assert.equal(s.layout, undefined);
  assert.equal(s.bg, undefined);
  assert.equal(s.transition, undefined);
  const [t, sh, ch] = s.items;
  assert.equal(s.items.length, 3);
  assert.deepEqual([t.x, t.y, t.w, t.rot], [-40, 0, 50, 270]);
  assert.deepEqual(t.paras[0].runs[0], { t: "5", color: "#FF0000", size: 12 });
  assert.equal(t.paras[0].level, 8);
  assert.equal(t.paras[0].align, undefined);
  assert.equal(t.paras[0].bullet, undefined);
  assert.equal(sh.shape, "rect");
  assert.deepEqual(ch.data, [1, 2]);
});

test("outline exports: markdown, text", () => {
  const d = newDeck();
  d.meta.title = "Plan";
  d.slides[0].items[0].paras = [{ runs: [{ t: "Kick-off" }] }];
  d.slides[0].notes = "Say hi\nthen start";
  const s = newSlide("content", d);
  s.items[0].paras = [{ runs: [{ t: "Agenda" }] }];
  s.items[1].paras = [
    { bullet: "ul", runs: [{ t: "One" }] },
    { bullet: "ol", level: 1, runs: [{ t: "Sub a" }] },
    { bullet: "ol", level: 1, runs: [{ t: "Sub b" }] },
  ];
  s.items.push({ type: "table", x: 0, y: 0, w: 100, h: 50, rows: [[{ paras: [{ runs: [{ t: "A" }] }] }, { paras: [{ runs: [{ t: "B|C" }] }] }], [{ paras: [] }, { paras: [{ runs: [{ t: "2" }] }] }]] });
  s.hidden = true;
  d.slides.push(s);
  const md = slidesToMarkdown(d);
  assert.match(md, /^# Plan/);
  assert.match(md, /## 1\. Kick-off/);
  assert.match(md, /> Say hi\n> then start/);
  assert.match(md, /## 2\. Agenda\n\n- One\n  1\. Sub a\n  2\. Sub b/);
  assert.match(md, /\| A \| B\\\|C \|\n\| --- \| --- \|\n\|  \| 2 \|/);
  const txt = slidesToText(d);
  assert.match(txt, /Slide 1: Kick-off/);
  assert.match(txt, /Slide 2: Agenda \(hidden\)/);
  assert.match(txt, /Notes: Say hi/);
  assert.equal(slideTitle(s), "Agenda");
});

test("HTML slideshow export: self-contained, escaped, navigation script, notes not shown", () => {
  const d = newDeck("dark");
  d.slides[0].items[0].paras = [{ runs: [{ t: "<script>alert(1)</script> & co", link: "javascript:alert(1)" }] }];
  d.slides[0].notes = "secret speaker note";
  d.slides.push({ items: [{ type: "image", x: 0, y: 0, w: 10, h: 10, src: PNG_URL }], hidden: true, transition: "fade" });
  const html = slidesToHtmlDocument(d, "My <deck>");
  assert.match(html, /^<!doctype html>/);
  assert.match(html, /<title>My &lt;deck&gt;<\/title>/);
  assert.ok(!html.includes("<script>alert(1)"));
  assert.ok(html.includes("&lt;script&gt;alert(1)&lt;/script&gt; &amp; co"));
  assert.ok(!/\shref="javascript:/i.test(html), "unsafe links are not clickable");
  assert.ok(!html.includes("secret speaker note"));
  assert.ok(html.includes(PNG_URL));
  assert.ok(html.includes('data-hidden="1"'));
  assert.match(html, /ArrowRight/);
  assert.ok(!/<script[^>]+src=/.test(html), "no external scripts");
  assert.ok(!/https?:\/\//.test(html.replace(/https?:\/\/[^"]*"/g, "")), "no remote resources");
});

test("renderer: shapes, bullets numbering, prompts only in edit mode, default sizes", () => {
  const d = newDeck();
  const s = newSlide("content", d);
  s.items[1].paras = [{ bullet: "ol", runs: [{ t: "a" }] }, { bullet: "ol", runs: [{ t: "b" }] }, { bullet: "ol", level: 1, runs: [{ t: "c" }] }];
  const html = renderSlideHtml(s, d, { edit: true });
  assert.match(html, /data-num="1\."[^>]*>.*a/);
  assert.match(html, /data-num="2\."/);
  assert.match(html, /data-level="1"[^>]*data-bullet="ol"[^>]*data-num="1\."/);
  assert.ok(html.includes('data-prompt="Click to add title"'));
  assert.ok(!renderSlideHtml(s, d).includes("data-prompt"));
  for (const sh of SHAPES) assert.match(shapePath(sh.id, 100, 50), /^M/);
  assert.equal(defaultRunSize({ type: "text", ph: "title" }), 44);
  assert.equal(defaultRunSize({ type: "text", ph: "body" }, 1), 24);
  assert.equal(textColorFor({ type: "shape", fill: "#000080" }, d), "#FFFFFF");
  assert.equal(textColorFor({ type: "text" }, d), d.theme.color);
});

test("resizeDeck keeps items inside the new slide and pictures in proportion", () => {
  const d = newDeck();
  d.slides[0].items.push({ type: "image", x: 860, y: 440, w: 100, h: 100, src: PNG_URL });
  resizeDeck(d, 720, 540);
  assert.deepEqual(d.size, { w: 720, h: 540 });
  const img = d.slides[0].items[2];
  assert.equal(img.w, img.h);
  for (const it of d.slides[0].items) assert.ok(it.x >= -0.01 && it.x + it.w <= 720.01, JSON.stringify(it));
});

// ---------------------------------------------------------------------------
// reader: python-pptx fixtures
// ---------------------------------------------------------------------------

test("pptxToModel: python-pptx deck (layouts, placeholders, bullets, notes, meta)", async () => {
  const { model: d, warnings } = await pptxToModel(readFix("basic.pptx"));
  assert.deepEqual(d.size, { w: 720, h: 540 });
  assert.equal(d.slides.length, 7);
  assert.equal(d.meta.title, "Fixture Deck");
  assert.equal(d.meta.author, "Test Author");
  assert.equal(d.theme.font, "Calibri");
  assert.match(d.theme.accent, /^#[0-9A-F]{6}$/);
  assert.deepEqual(d.slides.map((s) => s.layout), ["title", "content", "two", "section", "blank", "blank", "blank"]);
  // title slide: inherited geometry from the layout, sizes from the master
  const [title, sub] = d.slides[0].items;
  assert.equal(title.ph, "title");
  assert.equal(text(title.paras), "Quarterly Review");
  assert.ok(title.w > 500 && title.h > 50, "geometry inherited from the layout");
  assert.equal(title.paras[0].runs[0].size, 44);
  assert.equal(title.paras[0].align, "center");
  assert.equal(sub.ph, "subtitle");
  assert.equal(sub.paras[0].bullet, undefined, "subtitle has no bullet (layout buNone)");
  assert.equal(d.slides[0].notes, "Welcome everyone\nSecond note line");
  assert.equal(d.slides[0].transition, "fade");
  assert.equal(d.slides[1].transition, "push");
  // bullets with levels and level-dependent sizes
  const body = d.slides[1].items[1];
  assert.equal(body.ph, "body");
  assert.deepEqual(body.paras.map((p) => [text([p]), p.level || 0, p.bullet]), [["First point", 0, "ul"], ["Sub point", 1, "ul"], ["Deep point", 2, "ul"], ["Second point", 0, "ul"]]);
  assert.ok(body.paras[0].runs[0].size > body.paras[1].runs[0].size);
  // two content
  assert.deepEqual(d.slides[2].items.map((i) => i.ph), ["title", "body", "body"]);
  assert.ok(d.slides[2].items[2].x > d.slides[2].items[1].x + d.slides[2].items[1].w - 1);
  assert.ok(Array.isArray(warnings));
});

test("pptxToModel: runs, links, shapes, pictures, connectors, tables, groups, background, hidden, chart", async () => {
  const { model: d, warnings } = await pptxToModel(readFix("basic.pptx"));
  const s5 = d.slides[4].items;
  const tb = s5[0];
  assert.equal(tb.type, "text");
  assert.equal(tb.paras[0].align, "center");
  const [r0, r1, r2] = tb.paras[0].runs;
  assert.deepEqual([r0.t, r0.b, r0.size], ["Bold ", true, 28]);
  assert.deepEqual([r1.t, r1.i, r1.color], ["italic ", true, "#C00000"]);
  assert.deepEqual([r2.t, r2.u, r2.font], ["under", true, "Georgia"]);
  assert.equal(tb.paras[1].runs[0].link, "https://example.com/");
  const rect = s5[1];
  assert.deepEqual([rect.type, rect.shape, rect.fill, rect.line, rect.lineW], ["shape", "rect", "#118833", "#000000", 2]);
  assert.equal(text(rect.paras), "Box text");
  assert.equal(rect.paras[0].runs[0].color, "#FFFFFF", "shape style font color (lt1)");
  const oval = s5[2];
  assert.equal(oval.shape, "ellipse");
  assert.equal(oval.fill, undefined, "background fill = no fill");
  assert.equal(s5[3].shape, "rightArrow");
  assert.equal(s5[3].rot, 30);
  assert.equal(s5[4].shape, "star5");
  const pic = s5[5];
  assert.equal(pic.type, "image");
  assert.match(pic.src, /^data:image\/png;base64,/);
  assert.deepEqual([pic.w, pic.h], [144, 108]);
  const line = s5[6];
  assert.deepEqual([line.type, line.shape, line.w, line.h], ["shape", "line", 216, 0]);
  // table with style approximation and explicit cell fill
  const tbl = d.slides[5].items.find((i) => i.type === "table");
  assert.equal(tbl.rows.length, 3);
  assert.equal(text(tbl.rows[2][1].paras), "R2C1");
  assert.equal(tbl.rows[1][1].fill, "#FFEE00");
  assert.equal(tbl.rows[0][0].fill, d.theme.accent, "header row gets the accent color");
  assert.deepEqual(tbl.colW, [192, 192, 192]);
  assert.equal(d.slides[5].bg, "#EEF4FF");
  // group flattened with child offsets, hidden slide, chart placeholder + one warning
  const s7 = d.slides[6];
  assert.equal(s7.hidden, true);
  const [g1, g2, chart] = s7.items;
  assert.deepEqual([g1.x, g1.y, g1.w], [72, 72, 72]);
  assert.deepEqual([g2.shape, g2.x], ["ellipse", 216]);
  assert.equal(text(chart.paras), "[Chart]");
  assert.equal(warnings.filter((w) => /chart/i.test(w)).length, 1);
});

test("pptxToModel: master decorations, autofit, theme tints, AlternateContent, SVG, SmartArt, merged cells, rotated group, picture background", async () => {
  const { model: d, warnings } = await pptxToModel(readFix("complex.pptx"));
  assert.equal(d.slides.length, 4);
  const [s1, s2, s3, s4] = d.slides;
  // master decoration on slides whose layout shows master shapes, first in z-order
  assert.deepEqual([s1.items[0].type, s1.items[0].y, s1.items[0].w], ["shape", 520, 960]);
  assert.ok(!s3.items.some((i) => i.type === "shape"), "Title Only layout hides master shapes");
  assert.equal(s3.bg, "#FFF5E6", "layout background");
  // title with a line break and a slide-number field
  assert.equal(text(s1.items[1].paras), "Results\n1");
  // normAutofit fontScale 50% and lumMod theme color
  const run = s1.items[2].paras[0].runs[0];
  assert.equal(run.size, 16);
  assert.equal(run.color, "#254061");
  // gradient -> first stop, SVG picture keeps its PNG, SmartArt placeholder, merged cells, freeform -> rect
  assert.equal(s2.items[1].fill, "#FF0000");
  assert.match(s2.items[2].src, /^data:image\/png;base64,/);
  assert.equal(text(s2.items[3].paras), "[SmartArt]");
  const tbl = s2.items.find((i) => i.type === "table");
  assert.equal(text(tbl.rows[0][0].paras), "Merged");
  assert.equal(tbl.rows[0].length, 3);
  const ell = s2.items.find((i) => i.shape === "ellipse");
  assert.deepEqual([ell.x, ell.y, ell.w, ell.rot], [625, 325, 50, 90], "group scale + rotation applied to members");
  assert.equal(s2.items.find((i) => i.fill === "#AA00AA").shape, "rect");
  assert.equal(s3.transition, "zoom");
  // picture background becomes a full-slide picture at the back
  assert.deepEqual([s4.items[0].type, s4.items[0].w, s4.items[0].h, s4.items[0].alt], ["image", 960, 540, "Background picture"]);
  for (const re of [/Gradient/, /SmartArt/, /Merged/, /Freeform/, /Vertical/]) assert.equal(warnings.filter((w) => re.test(w)).length, 1, String(re));
  // and it survives a write/read cycle
  const again = await pptxToModel(await modelToPptx(d));
  assert.equal(again.model.slides.length, 4);
  assert.equal(text(again.model.slides[0].items[1].paras), "Results\n1");
  assert.deepEqual(again.warnings, []);
});

test("pptxToModel: 16:9 deck", async () => {
  const { model: d } = await pptxToModel(readFix("wide.pptx"));
  assert.deepEqual(d.size, { w: 960, h: 540 });
  assert.equal(text(d.slides[0].items[0].paras), "Wide deck");
});

test("pptxToModel: plain-language errors for non-pptx input", async () => {
  const cases = [
    [new Uint8Array(0), /empty/i],
    [new TextEncoder().encode("just some text, not a zip"), /not a PowerPoint presentation/i],
    [new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0, 0]), /97-2003/],
    [new Uint8Array([0x50, 0x4b, 3, 4, 1, 2, 3]), /damaged/i],
  ];
  for (const [bytes, re] of cases) await assert.rejects(pptxToModel(bytes), re);
  const word = new JSZip();
  word.file("[Content_Types].xml", "<Types/>");
  word.file("word/document.xml", "<w:document/>");
  await assert.rejects(pptxToModel(await word.generateAsync({ type: "uint8array" })), /Word document/);
  const empty = new JSZip();
  empty.file("hello.txt", "hi");
  await assert.rejects(pptxToModel(await empty.generateAsync({ type: "uint8array" })), /no presentation part/);
});

test("pptxToModel: tolerates missing slides, odd colors and deep nesting", async () => {
  const bytes = await modelToPptx(newDeck());
  const zip = await JSZip.loadAsync(bytes);
  // a broken slide reference and a slide with weird content
  const rels = await zip.file("ppt/_rels/presentation.xml.rels").async("string");
  zip.file("ppt/_rels/presentation.xml.rels", rels.replace("</Relationships>", '<Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/missing.xml"/></Relationships>'));
  const pres = await zip.file("ppt/presentation.xml").async("string");
  zip.file("ppt/presentation.xml", pres.replace("</p:sldIdLst>", '<p:sldId id="300" r:id="rId99"/></p:sldIdLst>'));
  let slide = await zip.file("ppt/slides/slide1.xml").async("string");
  let deep = "";
  for (let i = 0; i < 60; i++) deep += '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="9" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>';
  deep += '<p:sp><p:nvSpPr><p:cNvPr id="10" name="x"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="12700" cy="12700"/></a:xfrm><a:prstGeom prst="cloud"/><a:solidFill><a:schemeClr val="accent2"><a:lumMod val="75000"/></a:schemeClr></a:solidFill></p:spPr></p:sp>';
  for (let i = 0; i < 60; i++) deep += "</p:grpSp>";
  slide = slide.replace("</p:spTree>", deep + "</p:spTree>");
  zip.file("ppt/slides/slide1.xml", slide);
  const { model, warnings } = await pptxToModel(await zip.generateAsync({ type: "uint8array" }));
  assert.equal(model.slides.length, 1);
  assert.ok(warnings.some((w) => /missing/i.test(w)));
});

// ---------------------------------------------------------------------------
// writer
// ---------------------------------------------------------------------------

/** Every contract field, used for the round trip. */
function fullDeck() {
  const d = {
    kind: "slides",
    meta: { title: "Round & trip <deck>", author: "Ana \"Q\" Ödön" },
    size: { w: 720, h: 540 },
    theme: { font: "Georgia", headFont: "Trebuchet MS", color: "#222222", accent: "#C43E1C", bg: "#FAFAF5" },
    slides: [
      {
        layout: "title", transition: "fade", notes: "First note\nsecond line",
        items: [
          { type: "text", ph: "title", x: 90, y: 88.4, w: 540, h: 188, valign: "bottom", paras: [{ align: "center", runs: [{ t: "Hello", size: 54, b: true }] }] },
          { type: "text", ph: "subtitle", x: 90, y: 284, w: 540, h: 130, valign: "top", paras: [{ align: "center", runs: [{ t: "Sub", size: 24, i: true, color: "#336699" }] }] },
        ],
      },
      {
        layout: "content", bg: "#E8F0FF", transition: "push", hidden: true,
        items: [
          { type: "text", ph: "title", x: 50, y: 29, w: 620, h: 104, valign: "middle", paras: [{ runs: [{ t: "Agenda", size: 40 }] }] },
          { type: "text", ph: "body", x: 50, y: 144, w: 620, h: 343, valign: "top", paras: [
            { bullet: "ul", runs: [{ t: "Plain ", size: 24 }, { t: "bold", size: 24, b: true }, { t: " under", size: 24, u: true }] },
            { bullet: "ul", level: 1, runs: [{ t: "strike", size: 20, s: true }, { t: "2", size: 20, sup: true }, { t: "x", size: 20, sub: true }] },
            { bullet: "ol", level: 2, align: "right", runs: [{ t: "link", size: 20, link: "https://example.org/a?b=1&c=2" }] },
            { bullet: "ol", level: 2, align: "justify", runs: [{ t: "line one\nline two", size: 20, font: "Courier New", hl: "#FFFF00" }] },
            { runs: [{ t: "no bullet", size: 18, color: "#00AA00" }] },
          ] },
        ],
      },
      {
        layout: "two", transition: "wipe",
        items: [
          { type: "text", ph: "title", x: 50, y: 29, w: 620, h: 104, valign: "middle", paras: [{ runs: [{ t: "Two", size: 40 }] }] },
          { type: "text", ph: "body", x: 50, y: 144, w: 300, h: 343, valign: "top", paras: [{ bullet: "ul", runs: [{ t: "Left", size: 24 }] }] },
          { type: "text", ph: "body", x: 370, y: 144, w: 300, h: 343, valign: "top", paras: [{ bullet: "ul", runs: [{ t: "Right", size: 24 }] }] },
        ],
      },
      {
        layout: "section", transition: "zoom",
        items: [
          { type: "text", ph: "title", x: 49, y: 134, w: 620, h: 224, valign: "bottom", paras: [{ runs: [{ t: "Part 2", size: 54 }] }] },
          { type: "text", ph: "body", x: 49, y: 361, w: 620, h: 118, valign: "top", paras: [{ runs: [{ t: "Details", size: 24 }] }] },
        ],
      },
      {
        layout: "blank",
        items: [
          { type: "text", x: 20, y: 20, w: 200, h: 60, rot: 15, fill: "#FFF2CC", line: "#BF9000", lineW: 1.5, valign: "middle", paras: [{ runs: [{ t: "Box", size: 18 }] }] },
          ...SHAPES.map((sh, i) => ({
            type: "shape", shape: sh.id, x: 20 + (i % 6) * 110, y: 120 + Math.floor(i / 6) * 120, w: 100, h: sh.id === "line" ? 0 : 80,
            ...(sh.id === "line" ? { line: "#123456", lineW: 3 } : { fill: "#4472C4", line: "#2F5597", lineW: 1 }),
            ...(sh.id === "rect" ? { paras: [{ align: "center", runs: [{ t: "In shape", size: 14, color: "#FFFF00" }] }], valign: "bottom" } : {}),
          })),
          { type: "image", x: 400, y: 380, w: 120, h: 60, rot: 90, src: PNG_URL, alt: "Two pixels" },
          { type: "table", x: 20, y: 380, w: 300, h: 90, colW: [100, 200], rows: [
            [{ fill: "#C43E1C", paras: [{ runs: [{ t: "H1", size: 18, b: true }] }] }, { fill: "#C43E1C", paras: [{ runs: [{ t: "H2", size: 18, b: true }] }] }],
            [{ paras: [{ runs: [{ t: "a", size: 18 }] }] }, { fill: "#EEEEEE", paras: [{ runs: [{ t: "b & <c>", size: 18 }] }] }],
            [{ paras: [{ runs: [] }] }, { paras: [{ runs: [{ t: "end", size: 18 }] }] }],
          ] },
        ],
      },
    ],
  };
  return normalizeDeck(d);
}

/** Comparable form: defaults filled, empty runs dropped, numbers rounded. */
function canon(deck) {
  const d = cloneDeep(deck);
  const rnd = (v) => Math.round(v * 10) / 10;
  const theme = d.theme;
  const para = (p, it) => ({
    align: p.align || "left", level: p.level || 0, bullet: p.bullet || null,
    runs: (p.runs || []).filter((r) => r.t).map((r) => {
      const o = { t: r.t, size: r.size || defaultRunSize(it, p.level || 0) };
      for (const k of ["b", "i", "u", "s", "sup", "sub", "link", "hl"]) if (r[k]) o[k] = r[k];
      if (r.color) o.color = r.color;
      if (r.font) o.font = r.font;
      return o;
    }),
  });
  return {
    meta: { title: d.meta.title, author: d.meta.author }, size: d.size, theme,
    slides: d.slides.map((s) => ({
      layout: s.layout, bg: s.bg || null, notes: s.notes || "", hidden: !!s.hidden, transition: s.transition || "none",
      items: s.items.map((it) => {
        const o = { type: it.type, x: rnd(it.x), y: rnd(it.y), w: rnd(it.w), h: rnd(it.h), rot: rnd(it.rot || 0) };
        if (it.type === "text" || it.type === "shape") {
          o.ph = it.ph || null; o.fill = it.fill || null; o.line = it.line || null; o.lineW = it.line ? it.lineW || 0.75 : null;
          o.valign = it.valign || (it.type === "shape" ? "middle" : it.ph === "title" ? "middle" : "top");
          o.paras = (it.paras || []).map((p) => para(p, it)).filter((p, i, a) => p.runs.length || a.length > 1);
          if (it.type === "shape") { o.shape = it.shape; if (!o.paras.length) delete o.valign; }
        }
        if (it.type === "image") { o.src = it.src; o.alt = it.alt || ""; }
        if (it.type === "table") {
          delete o.h;
          o.colW = (it.colW || []).map(rnd);
          o.rows = it.rows.map((r) => r.map((c) => ({ fill: c.fill || null, paras: c.paras.map((p) => para(p, it)) })));
        }
        return o;
      }),
    })),
  };
}

test("round trip: model -> .pptx -> model keeps every contract field", async () => {
  const src = fullDeck();
  const bytes = await modelToPptx(src);
  const { model: back, warnings } = await pptxToModel(bytes);
  assert.deepEqual(warnings, []);
  const a = canon(src), b = canon(back);
  assert.deepEqual(b.meta, a.meta);
  assert.deepEqual(b.size, a.size);
  assert.deepEqual(b.theme, a.theme);
  assert.equal(b.slides.length, a.slides.length);
  for (let i = 0; i < a.slides.length; i++) {
    const sa = a.slides[i], sb = b.slides[i];
    for (const k of ["layout", "bg", "notes", "hidden", "transition"]) assert.deepEqual(sb[k], sa[k], `slide ${i + 1} ${k}`);
    assert.equal(sb.items.length, sa.items.length, `slide ${i + 1} item count`);
    sa.items.forEach((ia, j) => assert.deepEqual(sb.items[j], ia, `slide ${i + 1} item ${j}`));
  }
  // and a second trip is stable
  const again = await pptxToModel(await modelToPptx(back));
  assert.deepEqual(canon(again.model), b);
});

async function packageProblems(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  const names = new Set(Object.keys(zip.files).filter((n) => !zip.files[n].dir));
  const problems = [];
  const ct = await zip.file("[Content_Types].xml").async("string");
  const defaults = new Set([...ct.matchAll(/<Default Extension="([^"]+)"/g)].map((m) => m[1].toLowerCase()));
  const overrides = new Set([...ct.matchAll(/<Override PartName="\/([^"]+)"/g)].map((m) => m[1]));
  for (const o of overrides) if (!names.has(o)) problems.push(`override without part: ${o}`);
  for (const n of names) {
    if (n === "[Content_Types].xml") continue;
    if (!overrides.has(n) && !defaults.has(n.split(".").pop().toLowerCase())) problems.push(`no content type: ${n}`);
  }
  for (const n of names) {
    if (!n.endsWith(".rels")) continue;
    const base = n.replace(/_rels\/[^/]*\.rels$/, "");
    const xml = await zip.file(n).async("string");
    for (const m of xml.matchAll(/<Relationship ([^>]+)\/>/g)) {
      if (/TargetMode="External"/.test(m[1])) continue;
      const t = /Target="([^"]+)"/.exec(m[1])[1];
      const parts = (t.startsWith("/") ? t.slice(1) : base + t).split("/");
      const out = [];
      for (const p of parts) { if (p === "..") out.pop(); else if (p && p !== ".") out.push(p); }
      if (!names.has(out.join("/"))) problems.push(`${n}: missing target ${t}`);
    }
  }
  for (const n of names) {
    if (!/\.(xml|rels)$/.test(n)) continue;
    const xml = await zip.file(n).async("string");
    // well-formedness smoke check: balanced root, no raw control characters
    if (/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/.test(xml)) problems.push(`control chars in ${n}`);
    const tree = parseXml(xml);
    if (tree.children.filter((c) => c.local).length !== 1) problems.push(`bad root in ${n}`);
  }
  for (const n of names) {
    if (!/^ppt\/slides\/slide\d+\.xml$/.test(n)) continue;
    const ids = [...(await zip.file(n).async("string")).matchAll(/<p:cNvPr id="(\d+)"/g)].map((m) => m[1]);
    if (new Set(ids).size !== ids.length) problems.push(`duplicate shape ids in ${n}`);
  }
  return { problems, names, zip };
}

test("modelToPptx: complete package (content types, rels, ids, notes master, theme)", async () => {
  const bytes = await modelToPptx(fullDeck());
  const { problems, names, zip } = await packageProblems(bytes);
  assert.deepEqual(problems, []);
  for (const p of ["ppt/presentation.xml", "ppt/presProps.xml", "ppt/viewProps.xml", "ppt/tableStyles.xml", "ppt/theme/theme1.xml",
    "ppt/slideMasters/slideMaster1.xml", "ppt/slideLayouts/slideLayout6.xml", "ppt/notesMasters/notesMaster1.xml", "ppt/theme/theme2.xml",
    "docProps/core.xml", "docProps/app.xml", "ppt/media/image1.png"]) assert.ok(names.has(p), p);
  assert.equal(Object.keys(zip.files)[0], "[Content_Types].xml");
  const pres = await zip.file("ppt/presentation.xml").async("string");
  const ids = [...pres.matchAll(/<p:sldId id="(\d+)"/g)].map((m) => +m[1]);
  assert.deepEqual(ids, [256, 257, 258, 259, 260]);
  assert.match(pres, /<p:sldSz cx="9144000" cy="6858000" type="screen4x3"\/>/);
  assert.match(pres, /<p:notesMasterIdLst>/);
  const master = await zip.file("ppt/slideMasters/slideMaster1.xml").async("string");
  assert.ok([...master.matchAll(/<p:sldLayoutId id="(\d+)"/g)].every((m) => +m[1] > 2147483648));
  const theme = await zip.file("ppt/theme/theme1.xml").async("string");
  assert.match(theme, /<a:accent1><a:srgbClr val="C43E1C"\/>/);
  assert.match(theme, /<a:minorFont><a:latin typeface="Georgia"\/>/);
  assert.equal((theme.match(/<a:effectStyle>/g) || []).length, 3);
  const s2 = await zip.file("ppt/slides/slide2.xml").async("string");
  assert.match(s2, /<p:sld [^>]*show="0"/);
  assert.match(s2, /<p:transition spd="med"><p:push dir="u"\/><\/p:transition>/);
  const core = await zip.file("docProps/core.xml").async("string");
  assert.ok(core.includes("<dc:title>Round &amp; trip &lt;deck&gt;</dc:title>"));
});

test("modelToPptx: no notes master without notes; escapes and strips invalid XML characters", async () => {
  const d = newDeck();
  d.slides[0].items[0].paras = [{ runs: [{ t: "A\u0001B\u0007 <&> \"q\" 😀 \uD800" }] }];
  const bytes = await modelToPptx(d);
  const { problems, names, zip } = await packageProblems(bytes);
  assert.deepEqual(problems, []);
  assert.ok(!names.has("ppt/notesMasters/notesMaster1.xml"));
  const xml = await zip.file("ppt/slides/slide1.xml").async("string");
  assert.ok(xml.includes("AB &lt;&amp;&gt; &quot;q&quot; 😀 </a:t>"), xml.slice(xml.indexOf("<a:t>"), xml.indexOf("<a:t>") + 80));
  assert.equal(xmlEsc("a￾b"), "ab");
  const { model } = await pptxToModel(bytes);
  assert.equal(text(model.slides[0].items[0].paras), "AB <&> \"q\" 😀 ");
});

test("modelToPptx: tolerant of partial / odd models", async () => {
  for (const m of [{}, { slides: [] }, { slides: [{ items: [{ type: "image", x: 0, y: 0, w: 10, h: 10, src: "https://x/y.png" }, { type: "table", rows: [] }, { type: "chart" }] }] }]) {
    const bytes = await modelToPptx(m);
    const { problems } = await packageProblems(bytes);
    assert.deepEqual(problems, []);
    const { model } = await pptxToModel(bytes);
    assert.ok(model.slides.length >= 1);
  }
});

// python-pptx (optional): the real-world reader must open what we write
const PY = process.env.PYTHON || "";
test("python-pptx opens written files and reads their content", { skip: !PY && "set PYTHON to a python with python-pptx" }, async () => {
  const bytes = await modelToPptx(fullDeck());
  const tmp = path.join(os.tmpdir(), `slides-test-${process.pid}.pptx`);
  fs.writeFileSync(tmp, bytes);
  try {
    const r = spawnSync(PY, [path.join(FIX, "validate_pptx.py"), tmp], { encoding: "utf8" });
    assert.equal(r.status, 0, r.stdout + r.stderr);
    const v = JSON.parse(r.stdout);
    assert.equal(v.slides.length, 5);
    assert.equal(v.slides[0].notes, "First note\nsecond line");
    assert.equal(v.slides[1].hidden, true);
    assert.deepEqual(v.slides.map((s) => s.layout), ["Title Slide", "Title and Content", "Two Content", "Section Header", "Blank"]);
    assert.ok(v.slides[4].shapes.some((s) => s.table && s.table[1][1] === "b & <c>"));
    assert.ok(v.slides[4].shapes.some((s) => s.image === "image/png"));
    assert.equal(v.title, "Round & trip <deck>");
  } finally {
    fs.rmSync(tmp, { force: true });
  }
});

test("hexColor helper", () => {
  assert.equal(hexColor("#abc"), "#AABBCC");
  assert.equal(hexColor("123456"), "#123456");
  assert.equal(hexColor("red"), undefined);
});
