/** Tests for lib/formula.js (run with `npm test`). */

import { test } from "node:test";
import assert from "node:assert/strict";

import {
  tokenize, parseFormula, evaluateFormula, recalcBook, translateFormula,
  adjustFormulaForShift, renameSheetInFormula, deleteSheetInFormula, normalizeFormula,
  addXlfnPrefixes, stripXlfnPrefixes, formulaAutoFormat, shiftCells, renameSheet, deleteSheet,
  quoteSheetName, noteCellsChanged, sheetBounds, SUPPORTED_FUNCTIONS,
} from "../lib/formula.js";
import { ymdToSerial } from "../lib/model-excel.js";

const n = (v) => ({ t: "n", v });
const s = (v) => ({ t: "s", v });
const f = (formula) => ({ f: formula, t: "f", v: null });

function book(cells, extra = []) {
  return { kind: "excel", active: 0, sheets: [{ name: "Sheet1", cells }, ...extra] };
}

/** Evaluate a formula against a small default sheet. */
function ev(formula, cells = {}) {
  const b = book({
    A1: n(1), A2: n(2), A3: n(3), A4: n(4), A5: n(5),
    B1: s("apple"), B2: s("banana"), B3: s("cherry"), B4: s("apple"), B5: { t: "b", v: true },
    C1: n(10), C2: n(20), C3: n(30), C4: n(40), C5: n(50),
    D1: { t: "d", v: "2024-01-31" }, D2: { t: "d", v: "2024-03-15T12:00:00" },
    ...cells,
  }, [{ name: "My Sheet", cells: { A1: n(7), B2: s("x") } }]);
  return evaluateFormula(formula, { book: b, now: new Date(2024, 5, 15, 9, 30, 0) });
}

test("tokenizer recognises references, sheets, strings and errors", () => {
  const toks = tokenize("SUM($A$1:B2)+'My Sheet'!C3&\"a\"\"b\"+Data!A:A+1:3+#DIV/0!+LOG10(100)").filter((t) => t.k !== "ws");
  const kinds = toks.map((t) => t.k + (t.kind ? ":" + t.kind : ""));
  assert.deepEqual(kinds, [
    "func", "(", "ref:area", ")", "op", "ref:cell", "op", "str", "op", "ref:cols", "op", "ref:rows", "op", "err", "op", "func", "(", "num", ")",
  ]);
  assert.equal(toks[2].a.ac, true);
  assert.equal(toks[2].b.ac, false);
  assert.equal(toks[5].sheet, "My Sheet");
  assert.equal(toks[7].v, 'a"b');
  assert.equal(toks[15].v, "LOG10");
  assert.equal(tokenize("_xlfn.XLOOKUP(1,A1:A2,B1:B2)")[0].v, "XLOOKUP");
});

test("parser handles precedence like Excel", () => {
  assert.equal(ev("=-2^2"), 4);           // unary minus binds tighter than ^
  assert.equal(ev("=2^3^2"), 64);         // ^ is left-associative
  assert.equal(ev("=1+2*3"), 7);
  assert.equal(ev("=(1+2)*3"), 9);
  assert.equal(ev("=50%*2"), 1);
  assert.equal(ev("=1+2&3"), "33");
  assert.equal(ev("=1+2=3"), true);
  assert.equal(ev("=\"a\"<\"B\""), true); // case-insensitive
  assert.equal(ev("=1<\"a\""), true);     // numbers < text
  assert.equal(ev("=\"z\"<TRUE"), true);  // text < booleans
  assert.ok(parseFormula("=SUM(1,").error);
  assert.ok(parseFormula("=1+").error);
});

test("arithmetic coercion and error propagation", () => {
  assert.equal(ev('="3"+4'), 7);
  assert.equal(ev("=TRUE+1"), 2);
  assert.deepEqual(ev('="abc"+1'), { error: "#VALUE!" });
  assert.deepEqual(ev("=1/0"), { error: "#DIV/0!" });
  assert.deepEqual(ev("=(1/0)+1"), { error: "#DIV/0!" });
  assert.equal(ev("=Z99+1"), 1); // blank is 0
  assert.equal(ev("=0.1+0.2=0.3"), true);
  assert.deepEqual(ev("=NoSuchSheet!A1"), { error: "#REF!" });
  assert.equal(ev("='My Sheet'!A1*2"), 14);
});

test("aggregate functions", () => {
  assert.equal(ev("=SUM(A1:A5)"), 15);
  assert.equal(ev("=SUM(A1:B5)"), 15); // text/booleans in ranges ignored
  assert.equal(ev("=SUM(A:A)"), 15);
  assert.equal(ev("=SUM(1:1)"), 45333); // A1 + C1 + D1 (a date is a number)
  assert.equal(ev("=SUM(A1:A3,10,\"5\")"), 21);
  assert.equal(ev("=AVERAGE(A1:A5)"), 3);
  assert.deepEqual(ev("=AVERAGE(Z1:Z5)"), { error: "#DIV/0!" });
  assert.equal(ev("=MIN(A1:A5,C1:C5)"), 1);
  assert.equal(ev("=MAX(A1:A5,C1:C5)"), 50);
  assert.equal(ev("=COUNT(A1:C5)"), 10);
  assert.equal(ev("=COUNTA(A1:B5)"), 10);
  assert.equal(ev("=COUNTBLANK(A1:A10)"), 5);
  assert.equal(ev("=PRODUCT(A1:A4)"), 24);
  assert.equal(ev("=MEDIAN(A1:A4)"), 2.5);
  assert.equal(Math.round(ev("=STDEV(A1:A5)") * 1e6) / 1e6, 1.581139);
  assert.equal(Math.round(ev("=STDEV.P(A1:A5)") * 1e6) / 1e6, 1.414214);
  assert.equal(ev("=LARGE(C1:C5,2)"), 40);
  assert.equal(ev("=SMALL(C1:C5,2)"), 20);
  assert.equal(ev("=RANK(30,C1:C5)"), 3);
  assert.equal(ev("=RANK(30,C1:C5,1)"), 3);
  assert.equal(ev("=RANK(10,C1:C5)"), 5);
  assert.equal(ev("=SUMPRODUCT(A1:A5,C1:C5)"), 550);
  assert.equal(ev("=SUMPRODUCT((A1:A5>2)*C1:C5)"), 120);
});

test("conditional aggregates", () => {
  assert.equal(ev('=SUMIF(B1:B5,"apple",C1:C5)'), 50);
  assert.equal(ev('=SUMIF(A1:A5,">2")'), 12);
  assert.equal(ev('=SUMIF(B1:B5,"a*",A1:A5)'), 5);
  assert.equal(ev('=COUNTIF(B1:B5,"apple")'), 2);
  assert.equal(ev('=COUNTIF(B1:B5,"<>apple")'), 3);
  assert.equal(ev('=COUNTIF(A1:A5,">=3")'), 3);
  assert.equal(ev('=COUNTIF(A1:A10,"")'), 5);
  assert.equal(ev("=COUNTIF(A1:A5,3)"), 1);
  assert.equal(ev('=COUNTIFS(B1:B5,"apple",A1:A5,">1")'), 1);
  assert.equal(ev('=SUMIFS(C1:C5,B1:B5,"apple",A1:A5,"<5")'), 50);
  assert.equal(ev('=SUMIFS(C1:C5,B1:B5,"apple",A1:A5,">1")'), 40);
  assert.equal(ev('=AVERAGEIF(B1:B5,"apple",C1:C5)'), 25);
  assert.equal(ev('=AVERAGEIFS(C1:C5,A1:A5,">3")'), 45);
  assert.equal(ev('=MAXIFS(C1:C5,B1:B5,"apple")'), 40);
  assert.equal(ev('=MINIFS(C1:C5,B1:B5,"apple")'), 10);
});

test("logic and information functions", () => {
  assert.equal(ev('=IF(A1>0,"pos","neg")'), "pos");
  assert.equal(ev("=IF(A1>5,1)"), false);
  assert.equal(ev("=IF(A1>0,1/0,2)").error, "#DIV/0!");
  assert.equal(ev("=IF(A1<0,1/0,2)"), 2); // lazy
  assert.equal(ev('=IFS(A3>5,"big",A3>2,"mid",TRUE,"small")'), "mid");
  assert.deepEqual(ev("=IFS(FALSE,1)"), { error: "#N/A" });
  assert.equal(ev('=IFERROR(1/0,"oops")'), "oops");
  assert.equal(ev('=IFERROR(5,"oops")'), 5);
  assert.equal(ev('=IFNA(NA(),"na")'), "na");
  assert.equal(ev("=AND(A1>0,A2>0)"), true);
  assert.equal(ev("=OR(A1>5,A2>5)"), false);
  assert.equal(ev("=XOR(TRUE,TRUE,TRUE)"), true);
  assert.equal(ev("=NOT(A1=1)"), false);
  assert.equal(ev("=ISBLANK(Z1)"), true);
  assert.equal(ev("=ISNUMBER(A1)"), true);
  assert.equal(ev("=ISTEXT(B1)"), true);
  assert.equal(ev("=ISERROR(1/0)"), true);
  assert.equal(ev("=ISNA(NA())"), true);
  assert.equal(ev('=SWITCH(A2,1,"one",2,"two","other")'), "two");
  assert.equal(ev('=CHOOSE(3,"a","b","c")'), "c");
  assert.equal(ev("=ROW(C7)"), 7);
  assert.equal(ev("=COLUMNS(A1:D9)"), 4);
});

test("lookup functions", () => {
  assert.equal(ev('=VLOOKUP("cherry",B1:C5,2,FALSE)'), 30);
  assert.equal(ev('=VLOOKUP("CHERRY",B1:C5,2,0)'), 30);
  assert.deepEqual(ev('=VLOOKUP("kiwi",B1:C5,2,FALSE)'), { error: "#N/A" });
  assert.equal(ev("=VLOOKUP(3.5,A1:C5,3)"), 30); // approximate: largest <= 3.5
  assert.deepEqual(ev("=VLOOKUP(1,A1:C5,9,FALSE)"), { error: "#REF!" });
  assert.equal(ev("=HLOOKUP(10,A1:E5,2,FALSE)"), 20);
  assert.deepEqual(ev("=HLOOKUP(2,A1:E5,1,FALSE)"), { error: "#N/A" });
  assert.equal(ev('=MATCH("banana",B1:B5,0)'), 2);
  assert.equal(ev("=MATCH(25,C1:C5,1)"), 2);
  assert.equal(ev("=MATCH(25,{50,40,30,20,10},-1)"), 3);
  assert.equal(ev("=INDEX(C1:C5,4)"), 40);
  assert.equal(ev("=INDEX(A1:C5,2,3)"), 20);
  assert.equal(ev("=SUM(INDEX(A1:C5,0,3))"), 150);
  assert.equal(ev('=INDEX(C1:C5,MATCH("cherry",B1:B5,0))'), 30);
  assert.equal(ev('=XLOOKUP("banana",B1:B5,C1:C5)'), 20);
  assert.equal(ev('=XLOOKUP("kiwi",B1:B5,C1:C5,"none")'), "none");
  assert.equal(ev("=XLOOKUP(35,C1:C5,A1:A5,,-1)"), 3);
  assert.equal(ev("=XLOOKUP(35,C1:C5,A1:A5,,1)"), 4);
  assert.equal(ev('=XLOOKUP("apple",B1:B5,C1:C5,,0,-1)'), 40);
  assert.equal(ev('=XLOOKUP("ch*",B1:B5,C1:C5,,2)'), 30);
});

test("text functions", () => {
  assert.equal(ev('=CONCAT("a",A1:A3,"z")'), "a123z");
  assert.equal(ev('=CONCATENATE("x",1.5,TRUE)'), "x1.5TRUE");
  assert.equal(ev('=TEXTJOIN(", ",TRUE,B1:B3,Z1)'), "apple, banana, cherry");
  assert.equal(ev('=TEXTJOIN("-",FALSE,A1,Z1,A2)'), "1--2");
  assert.equal(ev('=LEFT("hello",2)'), "he");
  assert.equal(ev('=RIGHT("hello",3)'), "llo");
  assert.equal(ev('=MID("hello",2,3)'), "ell");
  assert.equal(ev('=LEN("hello")'), 5);
  assert.equal(ev('=UPPER("abc")'), "ABC");
  assert.equal(ev('=LOWER("ABC")'), "abc");
  assert.equal(ev('=PROPER("hello wORLD")'), "Hello World");
  assert.equal(ev('=TRIM("  a   b  ")'), "a b");
  assert.equal(ev('=SUBSTITUTE("a-b-c","-","+")'), "a+b+c");
  assert.equal(ev('=SUBSTITUTE("a-b-c","-","+",2)'), "a-b+c");
  assert.equal(ev('=FIND("l","hello")'), 3);
  assert.deepEqual(ev('=FIND("L","hello")'), { error: "#VALUE!" });
  assert.equal(ev('=SEARCH("L","hello")'), 3);
  assert.equal(ev('=SEARCH("l?o","hello")'), 3);
  assert.equal(ev('=TEXT(1234.5,"#,##0.00")'), "1,234.50");
  assert.equal(ev('=TEXT(0.256,"0.0%")'), "25.6%");
  assert.equal(ev('=TEXT(D1,"yyyy-mm-dd")'), "2024-01-31");
  assert.equal(ev('=TEXT(D1,"d-mmm-yy")'), "31-Jan-24");
  assert.equal(ev('=VALUE("1,234.5")'), 1234.5);
  assert.deepEqual(ev('=VALUE("abc")'), { error: "#VALUE!" });
  assert.equal(ev('=A1&""'), "1");
  assert.equal(ev("=1/3&\"\""), "0.333333333333333");
});

test("math functions", () => {
  assert.equal(ev("=ROUND(2.345,2)"), 2.35);
  assert.equal(ev("=ROUND(1.005,2)"), 1.01);
  assert.equal(ev("=ROUND(-2.5,0)"), -3);
  assert.equal(ev("=ROUND(1234,-2)"), 1200);
  assert.equal(ev("=ROUNDUP(2.341,2)"), 2.35);
  assert.equal(ev("=ROUNDDOWN(2.349,2)"), 2.34);
  assert.equal(ev("=ROUNDUP(-2.1,0)"), -3);
  assert.equal(ev("=INT(-2.5)"), -3);
  assert.equal(ev("=MOD(-3,2)"), 1);
  assert.equal(ev("=MOD(5.5,2)"), 1.5);
  assert.deepEqual(ev("=MOD(1,0)"), { error: "#DIV/0!" });
  assert.equal(ev("=ABS(-4)"), 4);
  assert.equal(ev("=SQRT(16)"), 4);
  assert.deepEqual(ev("=SQRT(-1)"), { error: "#NUM!" });
  assert.equal(ev("=POWER(2,10)"), 1024);
  assert.equal(Math.round(ev("=PI()") * 1e5), 314159);
  const r = ev("=RAND()");
  assert.ok(r >= 0 && r < 1);
});

test("date functions use naive serials", () => {
  const today = ymdToSerial(2024, 6, 15);
  assert.equal(ev("=TODAY()"), today);
  assert.equal(ev("=NOW()"), today + (9 * 3600 + 30 * 60) / 86400);
  assert.equal(ev("=DATE(2024,1,31)"), 45322);
  assert.equal(ev("=DATE(2024,14,1)"), ymdToSerial(2025, 2, 1)); // month overflow
  assert.equal(ev("=YEAR(D1)"), 2024);
  assert.equal(ev("=MONTH(D1)"), 1);
  assert.equal(ev("=DAY(D1)"), 31);
  assert.equal(ev("=HOUR(D2)"), 12);
  assert.equal(ev("=WEEKDAY(D1)"), 4); // Wednesday
  assert.equal(ev("=EDATE(D1,1)"), ymdToSerial(2024, 2, 29));
  assert.equal(ev("=EOMONTH(D1,1)"), ymdToSerial(2024, 2, 29));
  assert.equal(ev("=EOMONTH(D1,-1)"), ymdToSerial(2023, 12, 31));
  assert.equal(ev('=DATEDIF(DATE(2020,5,10),DATE(2024,3,9),"Y")'), 3);
  assert.equal(ev('=DATEDIF(DATE(2020,5,10),DATE(2024,3,9),"M")'), 45);
  assert.equal(ev('=DATEDIF(DATE(2024,1,1),DATE(2024,3,1),"D")'), 60);
  assert.equal(ev("=D1+1"), 45323);
  assert.equal(ev('=DATEVALUE("2024-01-31")'), 45322);
  assert.equal(ev("=DAYS(D2,D1)"), 44);
  assert.equal(ev("=TIME(12,30,0)"), 0.5208333333333334);
});

test("recalcBook stores computed values and keeps formulas", () => {
  const b = book({
    A1: n(2), A2: n(3), A3: f("=A1*A2"), A4: f("=A3+1"), A5: f("=SUM(A1:A4)"),
    B1: f("=A3&\"x\""), B2: f("=A1>1"), B3: f("=1/0"),
    C1: { t: "d", v: "2024-01-31", f: "=DATE(2024,1,31)" },
  });
  const res = recalcBook(b);
  const c = b.sheets[0].cells;
  assert.equal(res.count, 7);
  assert.deepEqual(c.A3, { f: "=A1*A2", t: "n", v: 6 });
  assert.equal(c.A4.v, 7);
  assert.equal(c.A5.v, 18);
  assert.deepEqual(c.B1, { f: '=A3&"x"', t: "s", v: "6x" });
  assert.deepEqual(c.B2, { f: "=A1>1", t: "b", v: true });
  assert.deepEqual(c.B3, { f: "=1/0", t: "e", v: "#DIV/0!" });
  assert.deepEqual(c.C1, { t: "d", v: "2024-01-31", f: "=DATE(2024,1,31)" }); // date cache stays a date
  // edit and recalc incrementally
  c.A1 = n(10);
  recalcBook(b, { changed: [{ sheet: 0, addrs: ["A1"] }] });
  assert.equal(c.A5.v, 10 + 3 + 30 + 31);
});

test("recalc order: dependencies below, deep chains and cross-sheet", () => {
  const cells = {};
  const N = 5000;
  // chain that points downwards (worst case for row-order evaluation)
  for (let i = 1; i < N; i++) cells[`A${i}`] = f(`=A${i + 1}+1`);
  cells[`A${N}`] = n(1);
  const b = book(cells, [{ name: "Other", cells: { A1: f("=Sheet1!A1*2") } }]);
  const res = recalcBook(b);
  assert.equal(cells.A1.v, N);
  assert.equal(b.sheets[1].cells.A1.v, N * 2);
  assert.deepEqual(res.cycles, []);
});

test("10k formulas recalc quickly", () => {
  const cells = {};
  for (let i = 1; i <= 10000; i++) {
    cells[`A${i}`] = n(i);
    cells[`B${i}`] = f(`=A${i}*2+IF(A${i}>5000,1,0)`);
  }
  cells.C1 = f("=SUM(B1:B10000)");
  const b = book(cells);
  recalcBook(b);
  const t0 = Date.now();
  const res = recalcBook(b);
  const ms = Date.now() - t0;
  assert.equal(res.count, 10001);
  assert.equal(cells.C1.v, 10000 * 10001 + 5000);
  assert.ok(ms < 1500, `recalc took ${ms}ms`);
});

test("circular references become #CIRC!", () => {
  const b = book({ A1: f("=B1+1"), B1: f("=A1+1"), C1: f("=C1"), D1: f("=SUM(A1:B1)") });
  const res = recalcBook(b);
  const c = b.sheets[0].cells;
  assert.equal(c.A1.t, "e");
  assert.equal(c.A1.v, "#CIRC!");
  assert.equal(c.C1.v, "#CIRC!");
  assert.equal(c.D1.v, "#CIRC!");
  assert.ok(res.cycles.length >= 2);
  // long cycle (deeper than the evaluator's recursion limit)
  const cells = {};
  for (let i = 1; i <= 1000; i++) cells[`A${i}`] = f(`=A${i === 1000 ? 1 : i + 1}+1`);
  const b2 = book(cells);
  const r2 = recalcBook(b2);
  assert.ok(r2.cycles.length >= 1);
  assert.equal(cells.A1.t, "e");
});

test("unknown functions keep the cached value, otherwise #NAME?", () => {
  const b = book({ A1: { f: "=WEBSERVICE(\"x\")", t: "n", v: 42 }, A2: f("=FOOBAR(1)"), A3: { f: "=Table1[Col]", t: "s", v: "kept" }, A4: f("=A1+1") });
  recalcBook(b);
  const c = b.sheets[0].cells;
  assert.equal(c.A1.v, 42);
  assert.equal(c.A1.t, "n");
  assert.deepEqual(c.A2, { f: "=FOOBAR(1)", t: "e", v: "#NAME?" });
  assert.equal(c.A3.v, "kept");
  assert.equal(c.A4.v, 43);
  // stable across recalcs
  recalcBook(b);
  assert.equal(c.A1.v, 42);
});

test("translateFormula shifts relative references only", () => {
  assert.equal(translateFormula("=A1+$B$2+B$3+$C4", 2, 1), "=B3+$B$2+C$3+$C6");
  assert.equal(translateFormula("=SUM(A1:B2)*'My Sheet'!C3", 1, 0), "=SUM(A2:B3)*'My Sheet'!C4");
  assert.equal(translateFormula("=SUM(A:A)+SUM(1:2)", 3, 2), "=SUM(C:C)+SUM(4:5)");
  assert.equal(translateFormula("=A1", -1, 0), "=#REF!");
  assert.equal(translateFormula('="A1"&A1', 1, 0), '="A1"&A2');
});

test("adjustFormulaForShift for inserted and deleted rows/columns", () => {
  const adj = (fx, axis, at, count, host = "Sheet1", target = "Sheet1") => adjustFormulaForShift(fx, host, target, axis, at, count);
  assert.equal(adj("=SUM(A1:A10)+A5+A3", "row", 3, 2), "=SUM(A1:A12)+A7+A3");
  assert.equal(adj("=SUM(A1:A10)+A5+A3", "row", 3, -3), "=SUM(A1:A7)+#REF!+A3");
  assert.equal(adj("=SUM(A4:A6)", "row", 3, -3), "=SUM(#REF!)");
  assert.equal(adj("=SUM(B1:D1)+C1", "col", 1, 1), "=SUM(C1:E1)+D1");
  assert.equal(adj("=SUM(B1:D1)", "col", 2, -1), "=SUM(B1:C1)");
  assert.equal(adj("=SUM(A:A)", "row", 0, 5), "=SUM(A:A)");
  assert.equal(adj("=Other!A5+A5", "row", 0, 1, "Sheet1", "Other"), "=Other!A6+A5");
  assert.equal(adj("=$A$5", "row", 0, 1), "=$A$6");
});

test("shiftCells moves cells, merges, widths and rewrites formulas in all sheets", () => {
  const b = book({ A1: n(1), A2: n(2), A3: f("=SUM(A1:A2)"), B5: s("x") }, [{ name: "Other", cells: { A1: f("=Sheet1!A3*2") } }]);
  b.sheets[0].merges = ["B5:C6"];
  b.sheets[0].colWidths = { A: 10, B: 20 };
  shiftCells(b, 0, "row", 1, 2); // insert 2 rows before row 2
  const c = b.sheets[0].cells;
  assert.equal(c.A1.v, 1);
  assert.equal(c.A4.v, 2);
  assert.equal(c.A5.f, "=SUM(A1:A4)");
  assert.equal(c.B7.v, "x");
  assert.deepEqual(b.sheets[0].merges, ["B7:C8"]);
  assert.equal(b.sheets[1].cells.A1.f, "=Sheet1!A5*2");
  shiftCells(b, 0, "col", 0, -1); // delete column A
  assert.equal(b.sheets[0].cells.A7.v, "x");
  assert.deepEqual(b.sheets[0].colWidths, { A: 20 });
  assert.deepEqual(b.sheets[0].merges, ["A7:B8"]);
  assert.equal(b.sheets[1].cells.A1.f, "=Sheet1!#REF!*2");
});

test("sheet rename/delete update references", () => {
  const b = book({ A1: f("='Q1 Data'!A1+'q1 data'!B2") }, [{ name: "Q1 Data", cells: { A1: n(1) } }]);
  renameSheet(b, 1, "Summary");
  assert.equal(b.sheets[0].cells.A1.f, "=Summary!A1+Summary!B2");
  renameSheet(b, 1, "2024 Data");
  assert.equal(b.sheets[0].cells.A1.f, "='2024 Data'!A1+'2024 Data'!B2");
  assert.equal(renameSheetInFormula("=Old!A1", "old", "It's"), "='It''s'!A1");
  assert.equal(deleteSheetInFormula("=Gone!A1+A1", "Gone"), "=#REF!+A1");
  b.active = 1;
  deleteSheet(b, 0);
  assert.equal(b.active, 0);
  assert.equal(b.sheets[0].name, "2024 Data");
  assert.equal(quoteSheetName("Sheet1"), "Sheet1");
  assert.equal(quoteSheetName("A1"), "'A1'");
  assert.equal(quoteSheetName("My Sheet"), "'My Sheet'");
});

test("normalizeFormula, _xlfn prefixes and auto formats", () => {
  assert.equal(normalizeFormula('=sum(a1:b2)+if(true,"abc",c3)'), '=SUM(A1:B2)+IF(TRUE,"abc",C3)');
  assert.equal(normalizeFormula('=sum(A1'), "=SUM(A1"); // tokens are normalised even if incomplete
  assert.equal(normalizeFormula('="unterminated'), '="unterminated'); // untokenizable text is left alone
  assert.equal(addXlfnPrefixes("=XLOOKUP(1,A:A,B:B)+SUM(1)+TEXTJOIN(\",\",TRUE,A1)"),
    '=_xlfn.XLOOKUP(1,A:A,B:B)+SUM(1)+_xlfn.TEXTJOIN(",",TRUE,A1)');
  assert.equal(addXlfnPrefixes("FILTER(A1:A3,B1:B3)"), "_xlfn._xlws.FILTER(A1:A3,B1:B3)");
  assert.equal(stripXlfnPrefixes('_xlfn.XLOOKUP("_xlfn.",A1,B1)'), 'XLOOKUP("_xlfn.",A1,B1)');
  const b = book({ A1: { t: "d", v: "2024-01-31", z: "m/d/yyyy" }, A2: n(5) });
  assert.equal(formulaAutoFormat(b, 0, "=TODAY()"), "yyyy-mm-dd");
  assert.equal(formulaAutoFormat(b, 0, "=NOW()"), "yyyy-mm-dd hh:mm");
  assert.equal(formulaAutoFormat(b, 0, "=A1+7"), "m/d/yyyy");
  assert.equal(formulaAutoFormat(b, 0, "=A2+7"), null);
  assert.ok(SUPPORTED_FUNCTIONS.includes("XLOOKUP"));
  assert.ok(SUPPORTED_FUNCTIONS.length >= 90);
});

test("noteCellsChanged keeps the formula index and bounds current", () => {
  const b = book({ A1: n(1) });
  recalcBook(b);
  const sh = b.sheets[0];
  sh.cells.C9 = f("=A1*3");
  noteCellsChanged(sh, ["C9"]);
  recalcBook(b);
  assert.equal(sh.cells.C9.v, 3);
  assert.deepEqual(sheetBounds(sh), { r2: 8, c2: 2 });
  delete sh.cells.C9;
  noteCellsChanged(sh, ["C9"]);
  assert.equal(recalcBook(b).count, 0);
});
