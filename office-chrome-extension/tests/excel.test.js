/** Tests for lib/model-excel.js and lib/local-xlsx.js (run with `npm test`). */

import { test } from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import vm from "node:vm";
import path from "node:path";
import { fileURLToPath } from "node:url";

import {
  formatValue, formatValueEx, formatGeneral, formatCell, cellDisplay, cellRaw, cellText, cellAlign,
  parseCellInput, cellFromInput, isoToSerial, serialToIso, ymdToSerial, serialToParts,
  csvParse, csvSerialize, sniffDelimiter, decodeTextBytes, encodeText, csvFieldToCell,
  csvToModel, csvBytesToModel, modelSheetToCsv, modelSheetToCsvBytes,
  sanitizeSheetName, validateSheetName, renderSheetHtml, isDateFormat, charsToPx, pxToChars,
} from "../lib/model-excel.js";
import { recalcBook } from "../lib/formula.js";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const EXT = path.resolve(HERE, "..");

// SheetJS is a browser global in the extension; load the vendored build the same way.
if (!globalThis.XLSX) {
  vm.runInThisContext(fs.readFileSync(path.join(EXT, "vendor/xlsx.full.min.js"), "utf8"), { filename: "xlsx.full.min.js" });
}
const { bytesToModel, modelToBytes, xlsxAvailable } = await import("../lib/local-xlsx.js");

const enc = (s) => new TextEncoder().encode(s);
const dec = (u8) => new TextDecoder().decode(u8);

/** Run fn with process.env.TZ temporarily set (Node re-reads TZ on change). */
function withTZ(tz, fn) {
  const old = process.env.TZ;
  process.env.TZ = tz;
  try { return fn(); } finally { if (old == null) delete process.env.TZ; else process.env.TZ = old; }
}

function zipText(bytes, name) {
  const zip = globalThis.XLSX.CFB.read(bytes, { type: "array" });
  const i = zip.FullPaths.findIndex((p) => p.endsWith("/" + name));
  return i < 0 ? null : dec(zip.FileIndex[i].content);
}

// ---------------------------------------------------------------------------
// number formats
// ---------------------------------------------------------------------------

test("format codes render like Excel", () => {
  const cases = [
    [1234.5, '"$"#,##0.00', "$1,234.50"],
    [-1234.5, '"$"#,##0.00', "-$1,234.50"],
    [0.153, "0.0%", "15.3%"],
    [0.5, "0%", "50%"],
    [42, "0", "42"],
    [3.14159, "0.00", "3.14"],
    [1.005, "0.00", "1.01"],
    [1234567.891, "#,##0", "1,234,568"],
    [1234567.891, "#,##0.00", "1,234,567.89"],
    [-1234.567, "#,##0.00;(#,##0.00)", "(1,234.57)"],
    [0, '#,##0.00;(#,##0.00);"-"', "-"],
    [1234567, "#,##0,", "1,235"],
    [0.5, "#.00", ".50"],
    [12345, "0.00E+00", "1.23E+04"],
    [0.000123, "0.00E+00", "1.23E-04"],
    [5551234567, "(000) 000-0000", "(555) 123-4567"],
    [7, "000", "007"],
    [1.25, "# ?/?", "1 1/4"],
    [1234.5, "#,##0.00 [$€-407]", "1,234.50 €"],
    [45306, "yyyy-mm-dd", "2024-01-15"],
    [45306, "m/d/yyyy", "1/15/2024"],
    [45306, "d-mmm-yy", "15-Jan-24"],
    [45306, "dddd, mmmm d, yyyy", "Monday, January 15, 2024"],
    [45306.4375, "yyyy-mm-dd hh:mm", "2024-01-15 10:30"],
    [0.5, "h:mm", "12:00"],
    [0.5625, "h:mm AM/PM", "1:30 PM"],
    [0.4375, "hh:mm:ss", "10:30:00"],
    [1.5, "[h]:mm:ss", "36:00:00"],
    [45306.5, "[$-409]mmm d, yyyy h:mm AM/PM", "Jan 15, 2024 12:00 PM"],
    ["abc", "@", "abc"],
    ["abc", '0;-0;0;"<"@">"', "<abc>"],
    [1234.5, "General", "1234.5"],
  ];
  for (const [v, z, want] of cases) assert.equal(formatValue(v, z), want, `${v} with ${z}`);
  assert.equal(formatValueEx(-5, "0;[Red]-0").color, "#ff0000");
  assert.ok(isDateFormat("yyyy-mm-dd"));
  assert.ok(isDateFormat("h:mm"));
  assert.ok(!isDateFormat("0.00"));
  assert.ok(!isDateFormat('"m"0'));
});

test("General format: 15-digit integers, 11 significant digits otherwise", () => {
  assert.equal(formatGeneral(123456789012345), "123456789012345");
  assert.equal(formatGeneral(0.1 + 0.2), "0.3");
  assert.equal(formatGeneral(1 / 3), "0.33333333333");
  assert.equal(formatGeneral(-2.5), "-2.5");
  assert.equal(formatGeneral(1e20), "1E+20");
  assert.equal(formatGeneral(1.5e-7), "0.00000015");
  assert.equal(formatGeneral(1e-12), "1E-12");
  assert.equal(formatGeneral(0), "0");
});

test("cellDisplay honours formats, errors, booleans and dates", () => {
  assert.equal(cellDisplay({ t: "n", v: 0.25, z: "0%" }), "25%");
  assert.equal(cellDisplay({ t: "d", v: "2024-01-15" }), "2024-01-15");
  assert.equal(cellDisplay({ t: "d", v: "2024-01-15T00:00:00" }), "2024-01-15"); // midnight hidden
  assert.equal(cellDisplay({ t: "d", v: "2024-01-15T10:30:00" }), "2024-01-15 10:30:00");
  assert.equal(cellDisplay({ t: "d", v: "2024-01-15", z: "d-mmm-yy" }), "15-Jan-24");
  assert.equal(cellDisplay({ t: "e", v: "#DIV/0!" }), "#DIV/0!");
  assert.equal(cellDisplay({ t: "b", v: false }), "FALSE");
  assert.equal(cellDisplay({ f: "=1+1", t: "n", v: 2 }), "2");
  assert.equal(cellDisplay({ f: "=A1", t: "f", v: null }), "=A1"); // never calculated
  assert.equal(formatCell({ t: "n", v: -1, z: "0;[Red]-0" }).color, "#ff0000");
  assert.equal(cellAlign({ t: "n", v: 1 }), "right");
  assert.equal(cellAlign({ t: "s", v: "x" }), "left");
  assert.equal(cellAlign({ t: "b", v: true }), "center");
  assert.equal(cellAlign({ t: "n", v: 1, s: { align: "center" } }), "center");
});

// ---------------------------------------------------------------------------
// data entry
// ---------------------------------------------------------------------------

test("parseCellInput: Excel-like entry semantics", () => {
  assert.deepEqual(parseCellInput("'007"), { t: "s", v: "007" });
  assert.deepEqual(parseCellInput("'=not a formula"), { t: "s", v: "=not a formula" });
  assert.deepEqual(parseCellInput("=a1+1"), { f: "=a1+1", t: "f", v: null });
  assert.deepEqual(parseCellInput("42"), { t: "n", v: 42 });
  assert.deepEqual(parseCellInput("-1.5e3"), { t: "n", v: -1500 });
  assert.deepEqual(parseCellInput("1,234.5"), { t: "n", v: 1234.5 });
  assert.deepEqual(parseCellInput("3%"), { t: "n", v: 0.03, z: "0%" });
  assert.deepEqual(parseCellInput("12.5%"), { t: "n", v: 0.125, z: "0.0%" });
  assert.deepEqual(parseCellInput("$1,234.50"), { t: "n", v: 1234.5, z: '"$"#,##0.00' });
  assert.deepEqual(parseCellInput("-$5"), { t: "n", v: -5, z: '"$"#,##0' });
  assert.deepEqual(parseCellInput("(123)"), { t: "n", v: -123 });
  assert.deepEqual(parseCellInput("2024-01-31"), { t: "d", v: "2024-01-31", z: "yyyy-mm-dd" });
  assert.deepEqual(parseCellInput("1/31/2024"), { t: "d", v: "2024-01-31", z: "m/dd/yyyy" });
  assert.deepEqual(parseCellInput("01/05/24"), { t: "d", v: "2024-01-05", z: "mm/dd/yy" });
  assert.deepEqual(parseCellInput("31-Jan-2024"), { t: "d", v: "2024-01-31", z: "d-mmm-yyyy" });
  assert.deepEqual(parseCellInput("2024-01-15 10:30"), { t: "d", v: "2024-01-15T10:30:00", z: "yyyy-mm-dd hh:mm" });
  assert.deepEqual(parseCellInput("10:30"), { t: "n", v: 0.4375, z: "hh:mm" });
  assert.deepEqual(parseCellInput("TRUE"), { t: "b", v: true });
  assert.deepEqual(parseCellInput("false"), { t: "b", v: false });
  assert.deepEqual(parseCellInput("#N/A"), { t: "e", v: "#N/A" });
  assert.deepEqual(parseCellInput("2024-02-30"), { t: "s", v: "2024-02-30" });
  assert.deepEqual(parseCellInput(" hello "), { t: "s", v: " hello " });
  assert.equal(parseCellInput(""), null);
});

test("typed dates and times never shift with the timezone", () => {
  for (const tz of ["UTC", "America/New_York", "Asia/Tokyo", "Pacific/Kiritimati", "America/Adak"]) {
    withTZ(tz, () => {
      assert.equal(parseCellInput("2024-01-15 10:30").v, "2024-01-15T10:30:00", tz);
      assert.equal(parseCellInput("2024-01-15T10:30:00Z").v, "2024-01-15T10:30:00", tz);
      assert.equal(isoToSerial("2024-01-15T10:30:00"), 45306.4375, tz);
      assert.equal(serialToIso(45306.4375), "2024-01-15T10:30:00", tz);
      assert.equal(serialToIso(45306), "2024-01-15", tz);
    });
  }
  assert.equal(ymdToSerial(1900, 1, 1), 1);
  assert.equal(ymdToSerial(1900, 2, 28), 59);
  assert.equal(ymdToSerial(1900, 3, 1), 61);
  assert.equal(serialToIso(60), "1900-02-29"); // Excel's fake leap day
  assert.equal(serialToParts(45306).dow, 1); // Monday
});

test("cellRaw round-trips through parseCellInput (W-04)", () => {
  const cells = [
    { t: "n", v: 1234.5 }, { t: "n", v: 0.1 + 0.2 }, { t: "n", v: 0.125, z: "0.0%" },
    { t: "n", v: 1234.5, z: '"$"#,##0.00' }, { t: "s", v: "00123" }, { t: "s", v: "TRUE" },
    { t: "s", v: "1,234" }, { t: "s", v: "=x" }, { t: "s", v: "'quoted" }, { t: "s", v: "plain" },
    { t: "b", v: true }, { t: "e", v: "#DIV/0!" }, { t: "d", v: "2024-01-15" },
    { t: "d", v: "2024-01-15T10:30:00", z: "yyyy-mm-dd hh:mm" }, { t: "n", v: 0.4375, z: "h:mm" },
    { f: "=SUM(A1:A3)", t: "n", v: 6 },
  ];
  for (const cell of cells) {
    const raw = cellRaw(cell);
    // re-committing the unchanged text keeps the very same cell
    assert.equal(cellFromInput(raw, cell), cell, JSON.stringify(cell));
    // and parsing it yields the same value and type
    const p = parseCellInput(raw);
    if (cell.f) { assert.equal(p.f, cell.f); continue; }
    assert.equal(p.t, cell.t, raw);
    if (cell.t === "n") assert.ok(Math.abs(p.v - cell.v) < 1e-9, raw);
    else assert.equal(p.v, cell.v, raw);
  }
  assert.equal(cellRaw({ t: "s", v: "00123" }), "'00123");
  assert.equal(cellRaw({ t: "n", v: 0.125, z: "0.0%" }), "12.5%");
  assert.equal(cellRaw({ t: "b", v: false }), "FALSE");
});

test("cellFromInput keeps styles and compatible number formats", () => {
  const prev = { t: "n", v: 1, z: '"$"#,##0.00', s: { b: true } };
  assert.deepEqual(cellFromInput("25", prev), { t: "n", v: 25, z: '"$"#,##0.00', s: { b: true } });
  assert.deepEqual(cellFromInput("hello", prev), { t: "s", v: "hello", z: '"$"#,##0.00', s: { b: true } });
  // clearing keeps formatting only
  assert.deepEqual(cellFromInput("", prev), { t: "s", v: null, z: '"$"#,##0.00', s: { b: true } });
  assert.equal(cellFromInput("", { t: "n", v: 1 }), null);
  // a date typed into a date-formatted cell keeps that format
  const d = { t: "d", v: "2024-01-01", z: "d-mmm-yy" };
  assert.deepEqual(cellFromInput("2024-02-03", d), { t: "d", v: "2024-02-03", z: "d-mmm-yy" });
  // Text (@) format stores whatever is typed as text
  assert.deepEqual(cellFromInput("007", { t: "s", v: "x", z: "@" }), { t: "s", v: "007", z: "@" });
});

// ---------------------------------------------------------------------------
// CSV
// ---------------------------------------------------------------------------

test("csvParse follows RFC 4180 (W-02)", () => {
  assert.deepEqual(csvParse('ruler,12" long,3\r\npipe,2" wide,5\r\n'), [["ruler", '12" long', "3"], ["pipe", '2" wide', "5"]]);
  assert.deepEqual(csvParse('a,"b,c","say ""hi"""\n'), [["a", "b,c", 'say "hi"']]);
  assert.deepEqual(csvParse('"multi\nline",x'), [["multi\nline", "x"]]);
  assert.deepEqual(csvParse("a\tb\tc", "\t"), [["a", "b", "c"]]);
  assert.deepEqual(csvParse('=HYPERLINK("x"),y'), [['=HYPERLINK("x")', "y"]]);
  assert.deepEqual(csvParse("a,b\n\nc,d"), [["a", "b"], [""], ["c", "d"]]);
  const rows = [["a\tb", "c"], ['q"', "x,y"]];
  assert.deepEqual(csvParse(csvSerialize(rows, "\t"), "\t"), rows);
  assert.deepEqual(csvParse(csvSerialize(rows, ";"), ";"), rows);
  assert.equal(csvSerialize([["a;b", "c,d"]], ";"), '"a;b";c,d');
});

test("delimiter sniffing and encoding detection (W-05)", () => {
  assert.equal(sniffDelimiter("a,b,c\n1,2,3\n"), ",");
  assert.equal(sniffDelimiter("a;b;c\n1;2,5;3\n"), ";");
  assert.equal(sniffDelimiter("Müller;12,5\n"), ";");
  assert.equal(sniffDelimiter("a\tb\n1\t2\n"), "\t");
  assert.equal(sniffDelimiter('"x,y";z\n"1,2";3\n'), ";");
  assert.equal(decodeTextBytes(new Uint8Array([0xef, 0xbb, 0xbf, 0x61])).encoding, "utf-8");
  assert.equal(decodeTextBytes(new Uint8Array([0xef, 0xbb, 0xbf, 0x61])).bom, true);
  const u16 = decodeTextBytes(new Uint8Array([0xff, 0xfe, 0x61, 0x00, 0xfc, 0x00]));
  assert.deepEqual([u16.text, u16.encoding], ["aü", "utf-16le"]);
  const cp = decodeTextBytes(new Uint8Array([0x4d, 0xfc, 0x6c, 0x6c, 0x65, 0x72, 0x80]));
  assert.deepEqual([cp.text, cp.encoding], ["Müller€", "windows-1252"]);
  assert.deepEqual([...encodeText("Müller€", "windows-1252")], [0x4d, 0xfc, 0x6c, 0x6c, 0x65, 0x72, 0x80]);
  // characters outside Windows-1252 switch to UTF-8 with BOM instead of corrupting
  assert.deepEqual([...encodeText("日", "windows-1252").slice(0, 3)], [0xef, 0xbb, 0xbf]);
});

test("CSV fields convert only when lossless (EX-07, RU-19)", () => {
  assert.deepEqual(csvFieldToCell("00501"), { t: "s", v: "00501" });
  assert.deepEqual(csvFieldToCell("+15551234567"), { t: "s", v: "+15551234567" });
  assert.deepEqual(csvFieldToCell("12345678901234567890"), { t: "s", v: "12345678901234567890" });
  assert.deepEqual(csvFieldToCell("1e5"), { t: "s", v: "1e5" });
  assert.deepEqual(csvFieldToCell("true"), { t: "s", v: "true" });
  assert.deepEqual(csvFieldToCell("=1+1"), { t: "s", v: "=1+1" });
  assert.deepEqual(csvFieldToCell("42"), { t: "n", v: 42 });
  assert.deepEqual(csvFieldToCell("1.50"), { t: "n", v: 1.5, z: "0.00" });
  assert.deepEqual(csvFieldToCell("1,234"), { t: "n", v: 1234, z: "#,##0" });
  assert.deepEqual(csvFieldToCell("50%"), { t: "n", v: 0.5, z: "0%" });
  assert.deepEqual(csvFieldToCell("TRUE"), { t: "b", v: true });
  assert.deepEqual(csvFieldToCell("2024-01-15"), { t: "d", v: "2024-01-15", z: "yyyy-mm-dd" });
  assert.deepEqual(csvFieldToCell("2024-01-15 10:30"), { t: "d", v: "2024-01-15T10:30:00", z: "yyyy-mm-dd hh:mm" });
});

test("open + save of an unedited CSV is byte-identical", () => {
  const files = [
    "00501,\"1,234\",50%,true,2024-01-15,1e5,12345678901234567890,+15551234567,2024-01-15 10:30\r\n",
    "id,name,amount,date\r\n007,\"Smith, J\",1.50,1/5/2024\r\n008,\"Say \"\"hi\"\"\",-3,01/06/2024\r\n",
    "a;b;c\n1,5;x;\"y;z\"\n",
    "\"a\",\"b\"\n\"1\",\"2\"\n",
    "x\ty\n1\t\"a\tb\"\n",
    "no trailing newline,1",
    "﻿with,bom\n1,2\n",
    "  spaced  ,TRUE,FALSE,#N/A,3.14159265358979\n",
  ];
  for (const text of files) {
    const bytes = enc(text);
    const name = text.includes("\t") ? "t.tsv" : "t.csv";
    const model = csvBytesToModel(bytes, name);
    recalcBook(model);
    const out = modelSheetToCsvBytes(model, 0, name.endsWith(".tsv") ? "tsv" : "csv");
    assert.equal(dec(out), dec(bytes), JSON.stringify(text));
    assert.deepEqual([...out], [...bytes]);
  }
  // Windows-1252 file with ';' and decimal commas
  const cp = new Uint8Array([0x4d, 0xfc, 0x6c, 0x6c, 0x65, 0x72, 0x3b, 0x31, 0x32, 0x2c, 0x35, 0x0d, 0x0a]);
  const m = csvBytesToModel(cp, "w.csv");
  assert.equal(m.sheets[0].cells.A1.v, "Müller");
  assert.equal(m.sheets[0].cells.B1.v, "12,5");
  assert.deepEqual([...modelSheetToCsvBytes(m, 0, "csv")], [...cp]);
  // UTF-16LE with BOM
  const u16 = new Uint8Array([0xff, 0xfe, ...[..."a,b\r\n"].flatMap((c) => [c.charCodeAt(0), 0])]);
  assert.deepEqual([...modelSheetToCsvBytes(csvBytesToModel(u16, "u.csv"), 0, "csv")], [...u16]);
});

test("CSV export writes values, keeps positions, TRUE/FALSE (EX-18, RV-M04)", () => {
  const model = {
    kind: "excel", active: 0, sheets: [{
      name: "S", cells: {
        B3: { t: "n", v: 1 }, C3: { t: "n", v: 2 }, D3: { f: "=SUM(B3:C3)", t: "f", v: null },
        B4: { t: "b", v: true }, C4: { f: "=1/0", t: "f", v: null }, D4: { t: "n", v: 0.5, z: "0%" },
      },
    }],
  };
  recalcBook(model);
  assert.equal(modelSheetToCsv(model, 0), ",,,\r\n,,,\r\n,1,2,3\r\n,TRUE,#DIV/0!,50%\r\n");
  assert.equal(modelSheetToCsv(model, 0, "\t"), "\t\t\t\r\n\t\t\t\r\n\t1\t2\t3\r\n\tTRUE\t#DIV/0!\t50%\r\n");
  assert.equal(cellText({ t: "n", v: 123456789012345 }), "123456789012345");
  assert.equal(cellText({ t: "d", v: "2024-01-15" }), "2024-01-15");
});

test("csvToModel sanitizes the sheet name (W-02 sheet names)", () => {
  assert.equal(csvToModel("a\n", "report [final]").sheets[0].name, "report _final_");
  const long = csvBytesToModel(enc("a\n"), "customer_orders_export_2026-11-20_final_version.csv").sheets[0].name;
  assert.ok(long.length <= 31);
  assert.equal(validateSheetName(long), null);
});

test("sheet name validation and sanitizing (EX-21)", () => {
  const sheets = [{ name: "Sheet1" }, { name: "Data" }];
  assert.match(validateSheetName("Q1/Q2", sheets), /cannot contain/);
  assert.match(validateSheetName("sheet1", sheets, 1), /already exists/);
  assert.equal(validateSheetName("sheet1", sheets, 0), null); // renaming itself (case change)
  assert.match(validateSheetName("'quoted", sheets), /apostrophe/);
  assert.match(validateSheetName("x".repeat(32), sheets), /31/);
  assert.match(validateSheetName("  ", sheets), /empty/);
  assert.match(validateSheetName("History", sheets), /reserved/);
  assert.equal(sanitizeSheetName("a/b:c*d?e[f]g\\h"), "a_b_c_d_e_f_g_h");
  assert.equal(sanitizeSheetName("sheet1", sheets), "sheet1 (2)");
  assert.equal(sanitizeSheetName("'x'"), "x");
  assert.equal(sanitizeSheetName(""), "Sheet");
  assert.equal(sanitizeSheetName("y".repeat(40), [{ name: "y".repeat(31) }]).length, 31);
});

test("renderSheetHtml renders merges, formats and alignment", () => {
  const { html } = renderSheetHtml({
    name: "S", merges: ["A1:B2"],
    cells: { A1: { t: "s", v: "Title", s: { b: true } }, C1: { t: "n", v: 0.5, z: "0%" }, B2: { t: "s", v: "hidden" } },
  }, { minRows: 3, minCols: 3 });
  assert.match(html, /data-addr="A1" rowspan="2" colspan="2" style="font-weight:700;">Title</);
  assert.doesNotMatch(html, /data-addr="B2"/);
  assert.match(html, /class="ow-cell ow-num" data-addr="C1" style="text-align:right;">50%</);
  assert.equal(charsToPx(8.43), 64);
  assert.equal(pxToChars(64), 8.43);
});

// ---------------------------------------------------------------------------
// SheetJS round trips (local-xlsx.js)
// ---------------------------------------------------------------------------

function sampleModel() {
  return {
    kind: "excel", active: 1,
    sheets: [
      {
        name: "Data",
        cells: {
          A1: { t: "d", v: "2024-01-15", z: "yyyy-mm-dd" },
          A2: { t: "d", v: "2024-01-15T10:30:00", z: "yyyy-mm-dd hh:mm" },
          A3: { t: "n", v: 0.5, z: "h:mm" },
          A4: { t: "n", v: 0.153, z: "0.0%" },
          A5: { t: "n", v: 1234.5, z: '"$"#,##0.00', s: { b: true, color: "#C00000", bg: "#FFF2CC", align: "center" } },
          A6: { t: "e", v: "#DIV/0!" },
          A7: { t: "s", v: "00123" },
          A8: { f: "=A4*100", t: "f", v: null },
          A9: { f: "=XLOOKUP(2,B1:B3,C1:C3)", t: "f", v: null },
          B1: { t: "n", v: 1 }, B2: { t: "n", v: 2 }, C1: { t: "s", v: "one" }, C2: { t: "s", v: "two" },
          D1: { t: "s", v: "merged", s: { i: true, u: true, wrap: true } },
        },
        colWidths: { A: 18.5, D: 30 },
        merges: ["D1:E2"],
      },
      { name: "Second Sheet", cells: { A1: { f: "=Data!B2*10", t: "f", v: null } } },
      { name: "Secret", hidden: true, cells: { A1: { t: "s", v: "hidden" } } },
    ],
  };
}

test("SheetJS write -> read keeps values, formats, styles, merges, widths (EX-05, EX-08, EX-15, EX-16)", () => {
  assert.ok(xlsxAvailable());
  for (const tz of ["America/New_York", "Asia/Tokyo"]) {
    withTZ(tz, () => {
      let model = sampleModel();
      recalcBook(model);
      for (let pass = 0; pass < 3; pass++) {
        const bytes = modelToBytes(model, "xlsx");
        model = bytesToModel(bytes, "book.xlsx");
        recalcBook(model);
      }
      const c = model.sheets[0].cells;
      assert.deepEqual(c.A1, { t: "d", v: "2024-01-15", z: "yyyy-mm-dd" }, tz);
      assert.deepEqual(c.A2, { t: "d", v: "2024-01-15T10:30:00", z: "yyyy-mm-dd hh:mm" }, tz);
      assert.deepEqual(c.A3, { t: "n", v: 0.5, z: "h:mm" });
      assert.deepEqual(c.A4, { t: "n", v: 0.153, z: "0.0%" });
      assert.deepEqual(c.A5, { t: "n", v: 1234.5, z: '"$"#,##0.00', s: { b: true, color: "#C00000", bg: "#FFF2CC", align: "center" } });
      assert.deepEqual(c.A6, { t: "e", v: "#DIV/0!" });
      assert.deepEqual(c.A7, { t: "s", v: "00123" });
      assert.deepEqual(c.A8, { t: "n", v: 15.3, f: "=A4*100" });
      assert.deepEqual(c.A9, { t: "s", v: "two", f: "=XLOOKUP(2,B1:B3,C1:C3)" });
      assert.deepEqual(c.D1.s, { i: true, u: true, wrap: true });
      assert.deepEqual(model.sheets[0].merges, ["D1:E2"]);
      assert.deepEqual(model.sheets[0].colWidths, { A: 18.5, D: 30 });
      assert.equal(model.sheets[1].cells.A1.v, 20);
      assert.equal(model.sheets[2].hidden, true);
      assert.equal(model.active, 1);
      assert.deepEqual(model.sheets.map((s) => s.name), ["Data", "Second Sheet", "Secret"]);
    });
  }
});

test("SheetJS writer: _xlfn prefixes, fullCalcOnLoad, cached values, formula-only cells (EX-11, EX-24)", () => {
  const model = sampleModel();
  model.sheets[0].cells.B9 = { f: "=SUM(B1:B2)", t: "f", v: null }; // never calculated
  const bytes = modelToBytes(model, "xlsx");
  const sheet1 = zipText(bytes, "xl/worksheets/sheet1.xml");
  assert.match(sheet1, /<f>_xlfn\.XLOOKUP\(2,B1:B3,C1:C3\)<\/f>/);
  assert.match(zipText(bytes, "xl/workbook.xml"), /<calcPr[^>]*fullCalcOnLoad="1"/);
  assert.match(zipText(bytes, "xl/workbook.xml"), /activeTab="1"/);
  assert.match(zipText(bytes, "xl/styles.xml"), /formatCode="&quot;\$&quot;#,##0\.00"/);
  const back = bytesToModel(bytes, "x.xlsx");
  assert.equal(back.sheets[0].cells.B9.f, "=SUM(B1:B2)"); // not dropped on reopen
  recalcBook(back);
  assert.equal(back.sheets[0].cells.B9.v, 3);
});

test("SheetJS reads the bundled sample workbook and recalculates stale totals (EX-06)", () => {
  const bytes = fs.readFileSync(path.join(EXT, "..", "example", "sample-workbook.xlsx"));
  const model = bytesToModel(new Uint8Array(bytes), "sample-workbook.xlsx");
  assert.ok(model.sheets.length >= 1);
  recalcBook(model);
  const c = model.sheets[0].cells;
  assert.equal(c.B5.f, "=SUM(B2:B4)");
  assert.equal(c.B5.v, c.B2.v + c.B3.v + c.B4.v);
  assert.ok(c.B5.v > 0);
  assert.ok(model.sheets[0].colWidths && model.sheets[0].colWidths.A > 0);
});

test("modelToBytes: csv/tsv/xls/ods and sanitized sheet names", () => {
  const model = sampleModel();
  recalcBook(model);
  model.active = 0;
  const csv = dec(modelToBytes(model, "csv"));
  assert.match(csv, /^2024-01-15,1,one,merged\r\n/);
  assert.match(csv, /\r\n15\.3,/);
  const tsv = dec(modelToBytes(model, "tsv"));
  assert.match(tsv, /^2024-01-15\t1\tone\tmerged\r\n/);
  for (const fmt of ["xls", "ods", "xlsm"]) {
    const out = modelToBytes(model, fmt);
    const back = bytesToModel(out, `x.${fmt}`);
    assert.equal(back.sheets[0].cells.B2.v, 2, fmt);
    assert.equal(back.sheets[0].cells.A7.v, "00123", fmt);
  }
  // invalid / case-duplicate names are sanitized instead of failing the save
  const bad = { kind: "excel", active: 0, sheets: [{ name: "Q1/Q2", cells: { A1: { t: "n", v: 1 } } }, { name: "q1_q2", cells: {} }] };
  const back = bytesToModel(modelToBytes(bad, "xlsx"), "b.xlsx");
  assert.deepEqual(back.sheets.map((s) => s.name), ["Q1_Q2", "q1_q2 (2)"]);
  // csv bytes passed to bytesToModel use the native codec
  const m = bytesToModel(enc("a;b\n1;2\n"), "x.csv");
  assert.equal(m.csvMeta.delim, ";");
  assert.equal(m.sheets[0].cells.B2.v, 2);
});

test("CSV save mirrors bare mid-field quotes (inch marks) and stays byte-identical", () => {
  const cases = [
    'id,name,zip,amount,note\r\n1,"Smith, J",00501,1.50,"He said ""hi"""\r\n2,Ann,02134,1e5,12" ruler\r\n',
    "a;b\n5'11\" tall;\"q\"\"\"\n",
    'a,b\n"x""y",z\n', // strict RFC writer: still quoted on save
  ];
  for (const src of cases) {
    const bytes = new TextEncoder().encode(src);
    const model = csvBytesToModel(bytes, "t.csv");
    assert.deepEqual(new TextDecoder().decode(modelSheetToCsvBytes(model, 0, "csv")), src);
  }
  // An edited cell containing a quote is still written readably.
  const m = csvBytesToModel(new TextEncoder().encode('a\n12" x\n'), "t.csv");
  m.sheets[0].cells.A3 = { t: "s", v: '"quoted"' };
  assert.match(new TextDecoder().decode(modelSheetToCsvBytes(m, 0, "csv")), /"""quoted"""/);
});

test("SheetJS path keeps value-less formatted cells (fill / number format on empty cells)", () => {
  const model = {
    kind: "excel", active: 0,
    sheets: [{ name: "S", cells: {
      A1: { t: "s", v: "head" },
      B1: { t: "s", v: null, s: { bg: "#FFF2CC", b: true } },
      C3: { t: "s", v: null, z: "#,##0.00" },
      A5: { t: "n", v: 1 },
      D7: { t: "s", v: null, z: "0%", s: { align: "right" } },
    } }],
  };
  let m = model;
  for (let pass = 0; pass < 2; pass++) m = bytesToModel(modelToBytes(m, "xlsx"), "b.xlsx");
  const c = m.sheets[0].cells;
  assert.deepEqual(c.B1, { t: "s", v: null, s: { b: true, bg: "#FFF2CC" } });
  assert.deepEqual(c.C3, { t: "s", v: null, z: "#,##0.00" });
  assert.deepEqual(c.D7, { t: "s", v: null, z: "0%", s: { align: "right" } });
  assert.deepEqual(c.A5, { t: "n", v: 1 });
  assert.equal(c.A1.v, "head");
});

test("SheetJS column widths use character units like Excel and the backend", async () => {
  const { storedWidthToChars, charsToStoredWidth } = await import("../lib/local-xlsx.js");
  assert.equal(storedWidthToChars(14.7109375), 14);
  assert.equal(storedWidthToChars(9.140625), 8.43);
  assert.equal(storedWidthToChars(charsToStoredWidth(18.5)), 18.5);
  // The bundled sample stores 14.7109375 etc.: the model says 14 / 10 / 12.
  const bytes = fs.readFileSync(path.join(EXT, "../example/sample-workbook.xlsx"));
  const m = bytesToModel(new Uint8Array(bytes), "sample-workbook.xlsx");
  assert.deepEqual(m.sheets[0].colWidths, { A: 14, B: 14, C: 10, D: 12 });
  // ... and a save -> reopen cycle does not grow them.
  const again = bytesToModel(modelToBytes(m, "xlsx"), "again.xlsx");
  assert.deepEqual(again.sheets[0].colWidths, { A: 14, B: 14, C: 10, D: 12 });
});
