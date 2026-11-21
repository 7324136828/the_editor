/** Tests for the excel-features stream (run with `npm test`). */

import { test } from "node:test";
import assert from "node:assert/strict";

import {
  evaluateFormula, shiftCells, addXlfnPrefixes, SUPPORTED_FUNCTIONS, FUNCTION_INFO, functionArgIndex,
} from "../lib/formula.js";
import { ymdToSerial } from "../lib/model-excel.js";

// ===========================================================================
// formula: tests (lib/formula.js)
// ===========================================================================

const fn = (v) => ({ t: "n", v });
const fs = (v) => ({ t: "s", v });

/** Evaluate a formula against an optional sheet of cells. */
function fev(formula, cells = {}) {
  const book = { kind: "excel", active: 0, sheets: [{ name: "Sheet1", cells }] };
  return evaluateFormula(formula, { book, now: new Date(2024, 5, 15, 9, 30, 0) });
}
const close = (actual, expected, tol = 1e-8) => {
  assert.equal(typeof actual, "number", `expected a number, got ${JSON.stringify(actual)}`);
  assert.ok(Math.abs(actual - expected) <= tol * Math.max(1, Math.abs(expected)), `${actual} != ${expected}`);
};
const isError = (v, code) => assert.deepEqual(v, { error: code });

test("formula: financial functions (PMT, FV, PV, NPER, RATE, NPV, IRR)", () => {
  close(fev("=PMT(0.05/12,60,10000)"), -188.7123364);
  close(fev("=PMT(0.08/12,10,10000)"), -1037.032089);
  close(fev("=PMT(0.08/12,10,10000,0,1)"), -1030.164327);
  close(fev("=PMT(0,10,1000)"), -100);
  close(fev("=FV(0.06/12,10,-200,-500,1)"), 2581.403374);
  close(fev("=FV(0.12/12,12,-1000)"), 12682.50301);
  close(fev("=PV(0.08/12,240,500)"), -59777.14585);
  close(fev("=PV(0,12,-100)"), 1200);
  close(fev("=NPER(0.12/12,-100,-1000,10000,1)"), 59.6738657);
  close(fev("=NPER(0.12/12,-100,-1000,10000)"), 60.0821229);
  close(fev("=NPER(0.12/12,-100,-1000)"), -9.57859404);
  close(fev("=RATE(4*12,-200,8000)"), 0.0077014725, 1e-7);
  close(fev("=RATE(4*12,-200,8000)*12"), 0.092417669, 1e-7);
  close(fev("=NPV(0.1,-10000,3000,4200,6800)"), 1188.443412);
  // NPV over a range ignores blanks and text
  close(fev("=NPV(0.08,A2:A6)+A1", { A1: fn(-40000), A2: fn(8000), A3: fn(9200), A4: fn(10000), A5: fn(12000), A6: fn(14500) }), 1922.061555);
  close(fev("=IRR({-70000,12000,15000,18000,21000,26000})"), 0.086630948);
  close(fev("=IRR({-70000,12000,15000,18000,21000})"), -0.021244848);
  close(fev("=IRR({-70000,12000,15000},-0.1)"), -0.443506941);
  isError(fev("=IRR({1,2,3})"), "#NUM!");
  isError(fev("=PMT(0.05,0,100)"), "#NUM!");
});

test("formula: NETWORKDAYS / WORKDAY (+ .INTL)", () => {
  assert.equal(fev("=NETWORKDAYS(DATE(2012,10,1),DATE(2013,3,1))"), 110);
  assert.equal(fev("=NETWORKDAYS(DATE(2012,10,1),DATE(2013,3,1),DATE(2012,11,22))"), 109);
  assert.equal(fev("=NETWORKDAYS(A1,A2,A3:A5)", {
    A1: { t: "d", v: "2012-10-01" }, A2: { t: "d", v: "2013-03-01" },
    A3: { t: "d", v: "2012-11-22" }, A4: { t: "d", v: "2012-12-04" }, A5: { t: "d", v: "2013-01-21" },
  }), 107);
  assert.equal(fev("=NETWORKDAYS(DATE(2013,3,1),DATE(2012,10,1))"), -110);
  assert.equal(fev("=NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,1,31))"), 22);
  assert.equal(fev("=NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,2,1),7,{\"2006/1/2\",\"2006/1/16\"})"), 22);
  assert.equal(fev("=NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,2,1),\"0010001\",{\"2006/1/2\",\"2006/1/16\"})"), 20);
  isError(fev("=NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,2,1),\"1111111\")"), "#VALUE!");

  assert.equal(fev("=WORKDAY(DATE(2008,10,1),151)"), 39933); // 2009-04-30
  assert.equal(fev("=WORKDAY(DATE(2008,10,1),151,{\"2008/11/26\",\"2008/12/4\",\"2009/1/21\"})"), ymdToSerial(2009, 5, 5));
  assert.equal(fev("=WORKDAY(DATE(2024,1,6),-1)"), ymdToSerial(2024, 1, 5)); // Saturday -> previous Friday
  assert.equal(fev("=WORKDAY(DATE(2024,1,5),1)"), ymdToSerial(2024, 1, 8)); // Friday -> Monday
  assert.equal(fev("=WORKDAY(DATE(2024,1,6),0)"), ymdToSerial(2024, 1, 6));
  isError(fev("=WORKDAY.INTL(DATE(2012,1,1),30,0)"), "#NUM!");
  assert.equal(fev("=WORKDAY.INTL(DATE(2012,1,1),90,11)"), 41013);
  assert.equal(fev("=WORKDAY.INTL(DATE(2012,1,1),30,17)"), 40944);
});

test("formula: WEEKNUM, ISOWEEKNUM and WEEKDAY return types", () => {
  assert.equal(fev("=WEEKNUM(DATE(2012,3,9))"), 10);
  assert.equal(fev("=WEEKNUM(DATE(2012,3,9),2)"), 11);
  assert.equal(fev("=WEEKNUM(DATE(2012,3,9),21)"), 10);
  assert.equal(fev("=ISOWEEKNUM(DATE(2012,3,9))"), 10);
  assert.equal(fev("=ISOWEEKNUM(DATE(2021,1,1))"), 53);
  assert.equal(fev("=WEEKNUM(DATE(2021,1,1),21)"), 53);
  assert.equal(fev("=WEEKNUM(DATE(2024,1,1))"), 1);
  assert.equal(fev("=WEEKNUM(DATE(2024,12,31))"), 53);
  assert.equal(fev("=WEEKNUM(DATE(2023,1,8),1)"), 2); // Sunday starts week 2
  assert.equal(fev("=WEEKNUM(DATE(2023,1,8),2)"), 2); // Monday-start: Jan 2-8 is week 2
  assert.equal(fev("=WEEKNUM(DATE(2023,1,1),2)"), 1);
  isError(fev("=WEEKNUM(DATE(2023,1,1),5)"), "#NUM!");

  // 2008-02-14 is a Thursday
  const exp = { 1: 5, 2: 4, 3: 3, 11: 4, 12: 3, 13: 2, 14: 1, 15: 7, 16: 6, 17: 5 };
  assert.equal(fev("=WEEKDAY(DATE(2008,2,14))"), 5);
  for (const [t, v] of Object.entries(exp)) assert.equal(fev(`=WEEKDAY(DATE(2008,2,14),${t})`), v, `type ${t}`);
  isError(fev("=WEEKDAY(DATE(2008,2,14),4)"), "#NUM!");
});

test("formula: PERCENTILE, QUARTILE and MODE families", () => {
  close(fev("=PERCENTILE({1,2,3,4},0.3)"), 1.9);
  close(fev("=PERCENTILE.INC({1,3,2,4},0.3)"), 1.9);
  close(fev("=PERCENTILE.EXC({1,2,3,6,6,6,7,8,9},0.25)"), 2.5);
  isError(fev("=PERCENTILE.EXC({1,2,3,6,6,6,7,8,9},0)"), "#NUM!");
  isError(fev("=PERCENTILE.EXC({1,2,3,6,6,6,7,8,9},0.01)"), "#NUM!");
  isError(fev("=PERCENTILE({1,2},1.5)"), "#NUM!");
  close(fev("=QUARTILE({1,2,4,7,8,9,10,12},1)"), 3.5);
  close(fev("=QUARTILE.INC({1,2,4,7,8,9,10,12},3)"), 9.25);
  assert.equal(fev("=QUARTILE({1,2,4,7,8,9,10,12},4)"), 12);
  assert.equal(fev("=QUARTILE.EXC({6,7,15,36,39,40,41,42,43,47,49},1)"), 15);
  assert.equal(fev("=QUARTILE.EXC({6,7,15,36,39,40,41,42,43,47,49},3)"), 43);
  isError(fev("=QUARTILE.EXC({1,2,3},0)"), "#NUM!");
  isError(fev("=QUARTILE({1,2,3},5)"), "#NUM!");
  assert.equal(fev("=MODE({5.6,4,4,3,2,4})"), 4);
  assert.equal(fev("=MODE.SNGL(A1:A6)", { A1: fn(5.6), A2: fn(4), A3: fn(4), A4: fn(3), A5: fn(2), A6: fn(4) }), 4);
  assert.equal(fev("=MODE({1,2,2,1,3})"), 1); // tie: first value in the data
  isError(fev("=MODE({1,2,3})"), "#N/A");
});

test("formula: statistics, ranking and rounding behave like Excel", () => {
  close(fev("=STDEV.S({2,4,4,4,5,5,7,9})"), 2.138089935);
  close(fev("=STDEV.P({2,4,4,4,5,5,7,9})"), 2);
  close(fev("=VAR.S({2,4,4,4,5,5,7,9})"), 4.571428571);
  close(fev("=VAR.P({2,4,4,4,5,5,7,9})"), 4);
  assert.equal(fev("=MEDIAN(1,2,3,4,5,6)"), 3.5);
  assert.equal(fev("=MEDIAN(1,2,3,4,5)"), 3);
  assert.equal(fev("=LARGE({3,5,3,5,4},3)"), 4);
  assert.equal(fev("=SMALL({3,4,5,2,3,4,5,6,4,7},4)"), 4);
  const cells = { A2: fn(7), A3: fn(3.5), A4: fn(3.5), A5: fn(1), A6: fn(2) };
  assert.equal(fev("=RANK(A3,A2:A6,1)", cells), 3);
  assert.equal(fev("=RANK.EQ(A2,A2:A6)", cells), 1);
  assert.equal(fev("=RANK.EQ(A2,A2:A6,0)", cells), 1);
  assert.equal(fev("=RANK.EQ(A2,A2:A6,1)", cells), 5);
  assert.equal(fev("=ROUNDUP(3.2,0)"), 4);
  assert.equal(fev("=ROUNDUP(3.14159,3)"), 3.142);
  assert.equal(fev("=ROUNDUP(-3.14159,1)"), -3.2);
  assert.equal(fev("=ROUNDUP(31415.92654,-2)"), 31500);
  assert.equal(fev("=ROUNDDOWN(3.14159,3)"), 3.141);
  assert.equal(fev("=ROUNDDOWN(-3.14159,1)"), -3.1);
  assert.equal(fev("=ROUNDDOWN(31415.92654,-2)"), 31400);
  assert.equal(fev("=MROUND(10,3)"), 9);
  assert.equal(fev("=MROUND(-10,-3)"), -9);
  close(fev("=MROUND(1.3,0.2)"), 1.4);
  isError(fev("=MROUND(5,-2)"), "#NUM!");
});

test("formula: conditional aggregates and lookups", () => {
  const cells = {};
  const qty = [5, 4, 15, 3, 22, 12, 10, 33];
  const prod = ["Apples", "Apples", "Artichokes", "Artichokes", "Bananas", "Bananas", "Carrots", "Carrots"];
  const who = ["Tom", "Sarah", "Tom", "Sarah", "Tom", "Sarah", "Tom", "Sarah"];
  qty.forEach((q, i) => { cells[`A${i + 2}`] = fn(q); cells[`B${i + 2}`] = fs(prod[i]); cells[`C${i + 2}`] = fs(who[i]); });
  assert.equal(fev('=SUMIFS(A2:A9,B2:B9,"=A*",C2:C9,"Tom")', cells), 20);
  assert.equal(fev('=SUMIFS(A2:A9,B2:B9,"<>Bananas",C2:C9,"Tom")', cells), 30);
  assert.equal(fev('=COUNTIFS(B2:B9,"A*",C2:C9,"Tom")', cells), 2);
  assert.equal(fev('=AVERAGEIFS(A2:A9,B2:B9,"A*",C2:C9,"Tom")', cells), 10);
  assert.equal(fev('=MAXIFS(A2:A9,C2:C9,"Sarah")', cells), 33);
  assert.equal(fev('=MINIFS(A2:A9,C2:C9,"Sarah",A2:A9,">3")', cells), 4);
  assert.equal(fev('=SUMIF(A2:A9,">10")', cells), 82);

  // XLOOKUP: exact, not found, next smaller/larger, wildcard, reverse search
  assert.equal(fev('=XLOOKUP("Bananas",B2:B9,A2:A9)', cells), 22);
  assert.equal(fev('=XLOOKUP("Bananas",B2:B9,A2:A9,0,0,-1)', cells), 12);
  assert.equal(fev('=XLOOKUP("Kiwi",B2:B9,A2:A9,"none")', cells), "none");
  isError(fev('=XLOOKUP("Kiwi",B2:B9,A2:A9)', cells), "#N/A");
  assert.equal(fev('=XLOOKUP("Ca*",B2:B9,A2:A9,,2)', cells), 10);
  assert.equal(fev('=XLOOKUP("Ca*",B2:B9,A2:A9,,2,-1)', cells), 33);
  assert.equal(fev('=XLOOKUP("Ca*",B2:B9,A2:A9)', cells).error, "#N/A"); // no wildcards in exact mode
  assert.equal(fev("=XLOOKUP(25,{10,20,30},{\"a\",\"b\",\"c\"},,-1)"), "b");
  assert.equal(fev("=XLOOKUP(25,{10,20,30},{\"a\",\"b\",\"c\"},,1)"), "c");
  assert.equal(fev("=XLOOKUP(25,{30,10,20},{\"a\",\"b\",\"c\"},,-1)"), "c"); // unsorted data is fine
  assert.equal(fev("=XLOOKUP(35,{10,20,30},{\"a\",\"b\",\"c\"},\"-\",1)"), "-");
});

test("formula: logic and text functions", () => {
  assert.equal(fev('=IFS(1>2,"a",2>1,"b")'), "b");
  isError(fev('=IFS(1>2,"a")'), "#N/A");
  assert.equal(fev('=SWITCH(2,1,"a",2,"b","z")'), "b");
  assert.equal(fev('=SWITCH(9,1,"a","z")'), "z");
  isError(fev('=SWITCH(9,1,"a")'), "#N/A");
  assert.equal(fev('=IFNA(NA(),"x")'), "x");
  isError(fev('=IFNA(1/0,"x")'), "#DIV/0!");
  assert.equal(fev('=TEXTJOIN(", ",TRUE,"a","","b")'), "a, b");
  assert.equal(fev('=TEXTJOIN(", ",FALSE,"a","","b")'), "a, , b");
  const cells = { A1: fs("x"), A3: fs("y"), B1: fn(1), B2: fn(2) };
  assert.equal(fev('=TEXTJOIN("-",TRUE,A1:A3,B1:B2)', cells), "x-y-1-2");
  assert.equal(fev('=TEXTJOIN("-",FALSE,A1:A3)', cells), "x--y");
  isError(fev('=TEXTJOIN("-",TRUE,{"a",#N/A})'), "#N/A");
  assert.equal(fev("=CONCAT(A1:A3,B1:B2)", cells), "xy12");
  assert.equal(fev('=PROPER("this is a TITLE")'), "This Is A Title");
  assert.equal(fev('=PROPER("2-way street")'), "2-Way Street");
  assert.equal(fev('=PROPER("76BudGet")'), "76Budget");
  assert.equal(fev('=SUBSTITUTE("Sales Data","Sales","Cost")'), "Cost Data");
  assert.equal(fev('=SUBSTITUTE("Quarter 1, 2008","1","2",1)'), "Quarter 2, 2008");
  assert.equal(fev('=SUBSTITUTE("Quarter 1, 2011","1","2",3)'), "Quarter 1, 2012");
  assert.equal(fev('=REPT("*-",3)'), "*-*-*-");
  assert.equal(fev('=FIND("M","Miriam McGovern")'), 1);
  assert.equal(fev('=FIND("m","Miriam McGovern")'), 6);
  assert.equal(fev('=FIND("M","Miriam McGovern",3)'), 8);
  isError(fev('=FIND("z","abc")'), "#VALUE!");
  assert.equal(fev('=SEARCH("e","Statements",6)'), 7);
  assert.equal(fev('=SEARCH("margin","Profit Margin")'), 8);
  assert.equal(fev('=SEARCH("p?o*t","My Profit")'), 4);
  assert.equal(fev('=SEARCH("~?","Why?")'), 4);
  assert.equal(fev('=VALUE("$1,000")'), 1000);
  close(fev('=VALUE("16:48:00")-VALUE("12:00:00")'), 0.2);
  assert.equal(fev('=VALUE("12%")'), 0.12);
  isError(fev('=VALUE("abc")'), "#VALUE!");
});

test("formula: TEXT supports common format codes", () => {
  assert.equal(fev('=TEXT(1234.567,"0.00")'), "1234.57");
  assert.equal(fev('=TEXT(1234.567,"#,##0")'), "1,235");
  assert.equal(fev('=TEXT(0.256,"0%")'), "26%");
  assert.equal(fev('=TEXT(0.123,"0.0%")'), "12.3%");
  assert.equal(fev('=TEXT(DATE(2024,3,5),"yyyy-mm-dd")'), "2024-03-05");
  assert.equal(fev('=TEXT(DATE(2024,3,5),"dd/mm/yyyy")'), "05/03/2024");
  assert.equal(fev('=TEXT(DATE(2024,3,5),"mmm d, yyyy")'), "Mar 5, 2024");
  assert.equal(fev('=TEXT(DATE(2024,3,5),"dddd")'), "Tuesday");
  assert.equal(fev('=TEXT(DATE(2024,3,5),"mmmm")'), "March");
  assert.equal(fev('=TEXT(TIME(15,7,0),"h:mm AM/PM")'), "3:07 PM");
  assert.equal(fev('=TEXT(1234.5,"$#,##0.00")'), "$1,234.50");
  assert.equal(fev('=TEXT(5,"000")'), "005");
});

test("formula: date functions (DATEDIF, EDATE, EOMONTH)", () => {
  assert.equal(fev('=DATEDIF(DATE(2001,1,1),DATE(2003,1,1),"Y")'), 2);
  assert.equal(fev('=DATEDIF(DATE(2001,6,1),DATE(2002,8,15),"D")'), 440);
  assert.equal(fev('=DATEDIF(DATE(2001,6,1),DATE(2002,8,15),"YD")'), 75);
  assert.equal(fev('=DATEDIF(DATE(2001,6,1),DATE(2002,8,15),"M")'), 14);
  assert.equal(fev('=DATEDIF(DATE(2001,6,1),DATE(2002,8,15),"YM")'), 2);
  assert.equal(fev('=DATEDIF(DATE(2001,6,1),DATE(2002,8,15),"MD")'), 14);
  assert.equal(fev('=DATEDIF(DATE(2020,5,20),DATE(2024,5,19),"y")'), 3);
  isError(fev('=DATEDIF(DATE(2003,1,1),DATE(2001,1,1),"Y")'), "#NUM!");
  assert.equal(fev("=EOMONTH(DATE(2011,1,1),1)"), 40602);
  assert.equal(fev("=EOMONTH(DATE(2011,1,1),-3)"), 40482);
  assert.equal(fev("=EDATE(DATE(2011,1,15),1)"), 40589);
  assert.equal(fev("=EDATE(DATE(2011,1,31),1)"), 40602);
  assert.equal(fev("=EDATE(DATE(2011,3,31),-1)"), 40602);
});

test("formula: new _xlfn functions get prefixed, classic ones do not", () => {
  assert.equal(addXlfnPrefixes("=NETWORKDAYS.INTL(A1,B1)+WORKDAY.INTL(A1,3)"), "=_xlfn.NETWORKDAYS.INTL(A1,B1)+_xlfn.WORKDAY.INTL(A1,3)");
  assert.equal(addXlfnPrefixes("=PERCENTILE.INC(A1:A9,0.5)+QUARTILE.EXC(A1:A9,1)+MODE.SNGL(A1:A9)+ISOWEEKNUM(A1)"),
    "=_xlfn.PERCENTILE.INC(A1:A9,0.5)+_xlfn.QUARTILE.EXC(A1:A9,1)+_xlfn.MODE.SNGL(A1:A9)+_xlfn.ISOWEEKNUM(A1)");
  const classic = "=NETWORKDAYS(A1,B1)+WORKDAY(A1,2)+WEEKNUM(A1)+PMT(0.1,2,3)+NPV(0.1,A2:A5)+IRR(A2:A5)+PERCENTILE(A2:A5,0.5)+QUARTILE(A2:A5,1)+MODE(A2:A5)+DAYS360(A1,B1)";
  assert.equal(addXlfnPrefixes(classic), classic);
  assert.equal(fev("=_xlfn.PERCENTILE.INC({1,2,3,4},0.3)"), 1.9);
});

test("formula: FUNCTION_INFO describes every supported function", () => {
  for (const name of SUPPORTED_FUNCTIONS) {
    const info = FUNCTION_INFO[name];
    assert.ok(info, `missing FUNCTION_INFO for ${name}`);
    assert.ok(info.sig.startsWith(name + "("), `bad signature for ${name}: ${info.sig}`);
    assert.ok(info.sig.endsWith(")"), `bad signature for ${name}`);
    assert.ok(/^[A-Z].*\.$/.test(info.desc), `bad description for ${name}`);
  }
  for (const name of Object.keys(FUNCTION_INFO)) assert.ok(SUPPORTED_FUNCTIONS.includes(name), `${name} is not supported`);
  for (const name of ["NETWORKDAYS", "WORKDAY", "WEEKNUM", "PMT", "FV", "PV", "NPER", "RATE", "NPV", "IRR",
    "PERCENTILE.INC", "PERCENTILE.EXC", "QUARTILE.EXC", "MODE.SNGL", "ISOWEEKNUM"]) {
    assert.ok(SUPPORTED_FUNCTIONS.includes(name), name);
  }
  assert.equal(FUNCTION_INFO.SUMIFS.sig, "SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)");
});

test("formula: functionArgIndex finds the call and argument under the caret", () => {
  assert.deepEqual(functionArgIndex("=SUM("), { name: "SUM", argIndex: 0 });
  assert.deepEqual(functionArgIndex("=sum(A1,"), { name: "SUM", argIndex: 1 });
  assert.deepEqual(functionArgIndex("=SUMIFS(A:A,B:B,\"x,y\",C:C"), { name: "SUMIFS", argIndex: 3 });
  assert.deepEqual(functionArgIndex("=IF(A1>0,SUM(B1,B2),"), { name: "IF", argIndex: 2 });
  assert.deepEqual(functionArgIndex("=IF(A1>0,SUM(B1,B2"), { name: "SUM", argIndex: 1 });
  assert.deepEqual(functionArgIndex("=VLOOKUP(A1,{1,2;3,4},"), { name: "VLOOKUP", argIndex: 2 });
  assert.deepEqual(functionArgIndex("=VLOOKUP(A1,{1,2;3"), { name: "VLOOKUP", argIndex: 1 });
  assert.deepEqual(functionArgIndex("=SUM((A1+B1)*2,"), { name: "SUM", argIndex: 1 });
  assert.deepEqual(functionArgIndex("=SUM((A1,"), { name: "SUM", argIndex: 0 });
  assert.deepEqual(functionArgIndex("=_xlfn.XLOOKUP(1,A:A,B:B,"), { name: "XLOOKUP", argIndex: 3 });
  assert.deepEqual(functionArgIndex("=PERCENTILE.INC(A1:A9;"), { name: "PERCENTILE.INC", argIndex: 1 });
  assert.deepEqual(functionArgIndex("=SUM('My, (Sheet'!A1,"), { name: "SUM", argIndex: 1 });
  assert.deepEqual(functionArgIndex("=TEXT(A1,\"(0,0\""), { name: "TEXT", argIndex: 1 });
  assert.deepEqual(functionArgIndex("=TEXT(A1,\"(0,"), { name: "TEXT", argIndex: 1 }); // inside an open string
  assert.deepEqual(functionArgIndex("=ROUND(1.5E+3,"), { name: "ROUND", argIndex: 1 });
  assert.equal(functionArgIndex("=SUM(A1)"), null);
  assert.equal(functionArgIndex("=A1+2"), null);
  assert.equal(functionArgIndex("SUM("), null);
  assert.equal(functionArgIndex(""), null);
});

test("formula: shiftCells keeps autoFilter, charts and freeze panes consistent", () => {
  const mk = () => ({
    kind: "excel", active: 0,
    sheets: [
      {
        name: "Data", cells: { A1: fs("h"), B5: fn(1) },
        autoFilter: "A1:D20",
        freeze: { rows: 2, cols: 1 },
        charts: [
          { type: "column", range: "A1:C10", x: 0, y: 0, w: 300, h: 200 },
          { type: "line", range: "E5:F8", x: 0, y: 0, w: 300, h: 200 },
        ],
      },
      { name: "Other", cells: {}, charts: [{ type: "pie", range: "Data!A1:B4" }, { type: "bar", range: "A1:B4" }] },
    ],
  });

  // insert 2 rows before row 1 (index 0): inside the frozen area
  let b = mk();
  shiftCells(b, 0, "row", 0, 2);
  let s = b.sheets[0];
  assert.equal(s.autoFilter, "A3:D22");
  assert.deepEqual(s.freeze, { rows: 4, cols: 1 });
  assert.equal(s.charts[0].range, "A3:C12");
  assert.equal(s.charts[1].range, "E7:F10");
  assert.equal(b.sheets[1].charts[0].range, "Data!A3:B6"); // other sheet referring to Data
  assert.equal(b.sheets[1].charts[1].range, "A1:B4"); // other sheet's own data: unchanged

  // insert rows below the frozen area / inside the filter range
  b = mk();
  shiftCells(b, 0, "row", 2, 3);
  s = b.sheets[0];
  assert.deepEqual(s.freeze, { rows: 2, cols: 1 });
  assert.equal(s.autoFilter, "A1:D23");
  assert.equal(s.charts[0].range, "A1:C13");

  // delete rows 6..8 (index 5, 3 rows): E5:F8 keeps only row 5
  b = mk();
  shiftCells(b, 0, "row", 5, -3);
  s = b.sheets[0];
  assert.equal(s.autoFilter, "A1:D17");
  assert.equal(s.charts[1].range, "E5:F5");

  // delete the chart's whole data range: the chart is removed
  b = mk();
  shiftCells(b, 0, "row", 4, -4); // rows 5..8
  s = b.sheets[0];
  assert.equal(s.charts.length, 1);
  assert.equal(s.charts[0].range, "A1:C6");

  // deleting rows inside the frozen area shrinks it; deleting all frozen rows and cols removes freeze
  b = mk();
  shiftCells(b, 0, "row", 1, -5);
  assert.deepEqual(b.sheets[0].freeze, { rows: 1, cols: 1 });
  shiftCells(b, 0, "row", 0, -1);
  assert.deepEqual(b.sheets[0].freeze, { rows: 0, cols: 1 });
  shiftCells(b, 0, "col", 0, -1);
  assert.equal(b.sheets[0].freeze, undefined);

  // columns: insert inside the frozen columns, delete the whole filter range
  b = mk();
  shiftCells(b, 0, "col", 0, 1);
  s = b.sheets[0];
  assert.deepEqual(s.freeze, { rows: 2, cols: 2 });
  assert.equal(s.autoFilter, "B1:E20");
  assert.equal(s.charts[0].range, "B1:D10");
  assert.equal(s.charts[1].range, "F5:G8");
  b = mk();
  shiftCells(b, 0, "col", 0, -4); // delete A:D
  s = b.sheets[0];
  assert.equal(s.autoFilter, undefined);
  assert.equal("autoFilter" in s, false);
  assert.deepEqual(s.charts.map((c) => c.range), ["A5:B8"]);
  assert.deepEqual(s.freeze, { rows: 2, cols: 0 });
  assert.deepEqual(b.sheets[1].charts.map((c) => c.range), ["A1:B4"]); // Data!A1:B4 chart dropped

  // columns after the frozen area leave it alone; absolute refs keep their $
  b = mk();
  b.sheets[0].autoFilter = "$B$2:$C$9";
  shiftCells(b, 0, "col", 1, 2);
  assert.deepEqual(b.sheets[0].freeze, { rows: 2, cols: 1 });
  assert.equal(b.sheets[0].autoFilter, "$D$2:$E$9");
});

// ===========================================================================
// sort, filter, fill series, charts and the xlsx round trip
// (lib/model-excel.js, lib/local-xlsx.js)
// ===========================================================================

import { readFileSync, existsSync } from "node:fs";
import { runInThisContext } from "node:vm";
import { fileURLToPath as toPath } from "node:url";
import { dirname as dirOf, resolve as resolvePath, join as joinPath } from "node:path";
import {
  currentRegion, detectHeaderRow, sortRange, sortKeyOf, compareSortKeys, filterHiddenRows, filterValues,
  makeFilterTest, fillSeries, chartData, chartSvg, chartAnchorFromXY, chartXYFromAnchor, chartLayout,
  gridColPx, excelColPx, offsetRef,
} from "../lib/model-excel.js";
import { translateFormula, recalcBook } from "../lib/formula.js";

const XF_EXT = resolvePath(dirOf(toPath(import.meta.url)), "..");
if (!globalThis.XLSX) runInThisContext(readFileSync(joinPath(XF_EXT, "vendor/xlsx.full.min.js"), "utf8"), { filename: "xlsx.full.min.js" });
const LX = await import("../lib/local-xlsx.js");

const N = (v) => ({ t: "n", v });
const S = (v) => ({ t: "s", v });
const D = (v) => ({ t: "d", v });
const tr = (f, dr, dc) => translateFormula(f, dr, dc);

test("sort: Excel order (numbers < text < logical < errors < blanks), blanks last both ways", () => {
  const keys = [N(10), S("b"), { t: "b", v: true }, null, S("A"), N(-2), { t: "e", v: "#N/A" }, D("2024-01-01"), S("")].map(sortKeyOf);
  const order = (desc) => keys.map((k, i) => [k, i]).sort((a, b) => compareSortKeys(a[0], b[0], desc) || a[1] - b[1]).map(([, i]) => i);
  assert.deepEqual(order(false), [5, 0, 7, 4, 1, 2, 6, 3, 8]); // -2, 10, date serial 45292, "A", "b", TRUE, #N/A, blank, ""
  assert.deepEqual(order(true), [6, 2, 1, 4, 7, 0, 5, 3, 8]);
  // text compare is case-insensitive like Excel
  assert.equal(compareSortKeys(sortKeyOf(S("apple")), sortKeyOf(S("APPLE"))), 0);
});

test("sort: header row detected, rows move together, relative refs shift, absolute refs stay", () => {
  const sheet = { name: "S", cells: {
    A1: S("Name"), B1: S("Qty"), C1: S("Total"),
    A2: S("pear"), B2: N(3), C2: { t: "n", v: 30, f: "=B2*$E$1" },
    A3: S("apple"), B3: N(1), C3: { t: "n", v: 10, f: "=B3*$E$1" },
    A4: S("fig"), B4: S("n/a"), C4: { t: "n", v: 0, f: "=SUM(B$2:B4)" },
    A5: S("kiwi"), C5: N(5),
    E1: N(10),
  } };
  const rg = currentRegion(sheet, 2, 1);
  assert.deepEqual(rg, { r1: 0, r2: 4, c1: 0, c2: 2 });
  assert.equal(detectHeaderRow(sheet, rg), true);
  const res = sortRange(sheet, rg, [{ col: 1 }], { header: true, translate: tr });
  const after = Object.fromEntries(res.changes.map((c) => [c.addr, c.after]));
  assert.deepEqual([2, 3, 4, 5].map((r) => after[`A${r}`].v), ["apple", "pear", "fig", "kiwi"]);
  assert.equal(after.C2.f, "=B2*$E$1"); // row 3 -> row 2: same-row reference stays consistent
  assert.equal(after.C3.f, "=B3*$E$1");
  assert.equal(after.C4.f, "=SUM(B$2:B4)"); // did not move
  assert.equal(after.B5, null); // blank key last
  // recalculating the sorted sheet gives the right totals
  for (const ch of res.changes) if (ch.after) sheet.cells[ch.addr] = ch.after; else delete sheet.cells[ch.addr];
  recalcBook({ kind: "excel", sheets: [sheet] });
  assert.equal(sheet.cells.C2.v, 10);
  assert.equal(sheet.cells.C3.v, 30);
  // descending, two keys
  const s2 = { cells: { A1: S("k"), B1: S("v"), A2: S("x"), B2: N(1), A3: S("y"), B3: N(2), A4: S("x"), B4: N(3) } };
  const r2 = sortRange(s2, { r1: 0, r2: 3, c1: 0, c2: 1 }, [{ col: 0 }, { col: 1, desc: true }], { header: true });
  assert.deepEqual(r2.changes.filter((c) => c.addr[0] === "B").map((c) => c.after.v), [3, 1, 2]);
});

test("sort: merged cells are refused with a clear message; no header guessed for all-text data", () => {
  const sheet = { cells: { A1: S("a"), A2: S("b"), B1: N(1) }, merges: ["A1:A2"] };
  const res = sortRange(sheet, { r1: 0, r2: 2, c1: 0, c2: 1 }, [{ col: 0 }], {});
  assert.match(res.error, /merged/i);
  assert.equal(detectHeaderRow({ cells: { A1: S("x"), A2: S("y") } }, { r1: 0, r2: 1, c1: 0, c2: 0 }), false);
  assert.equal(detectHeaderRow({ cells: { A1: { t: "s", v: "Name", s: { b: true } }, A2: S("y") } }, { r1: 0, r2: 1, c1: 0, c2: 0 }), true);
});

test("filter: value lists, blanks, number / text conditions, top N and combined columns", () => {
  const sheet = { cells: { A1: S("Region"), B1: S("Units") } };
  const data = [["East", 10], ["West", 50], [null, 30], ["East", 70], ["North", 5], ["west", 90]];
  data.forEach(([a, b], i) => { if (a) sheet.cells[`A${i + 2}`] = S(a); sheet.cells[`B${i + 2}`] = N(b); });
  assert.deepEqual(filterValues(sheet, 0, 1, 6), ["East", "North", "West", "west", ""].filter((v, i, arr) => arr.indexOf(v) === i).sort((a, b) => (a === "" ? 1 : b === "" ? -1 : a.localeCompare(b, undefined, { sensitivity: "base" }))));
  const hid = (crit) => filterHiddenRows(sheet, "A1:B7", crit);
  assert.deepEqual(hid({ 0: { values: ["East"] } }), [2, 3, 5, 6]);
  assert.deepEqual(hid({ 0: { values: [""] } }), [1, 2, 4, 5, 6]);
  assert.deepEqual(hid({ 1: { cond: { op: "gt", a: "40" } } }), [1, 3, 5]);
  assert.deepEqual(hid({ 1: { cond: { op: "between", a: "10", b: "50" } } }), [4, 5, 6]);
  assert.deepEqual(hid({ 1: { cond: { op: "top", a: "2" } } }), [1, 2, 3, 5]);
  assert.deepEqual(hid({ 1: { cond: { op: "above" } } }), [1, 3, 5]); // average 42.5
  assert.deepEqual(hid({ 0: { cond: { op: "begins", a: "we" } } }), [1, 3, 4, 5]);
  assert.deepEqual(hid({ 0: { cond: { op: "contains", a: "st" } } }), [3, 5]);
  assert.deepEqual(hid({ 0: { values: ["East"] }, 1: { cond: { op: "gt", a: 20 } } }), [1, 2, 3, 5, 6]);
  assert.equal(makeFilterTest(null)(N(1)), true);
});

test("fill series: numbers, trends, Ctrl, dates, names, text with numbers, formulas, reverse", () => {
  const vals = (src, n, o) => fillSeries(src, n, o).map((c) => c && c.v);
  assert.deepEqual(vals([N(1), N(2), N(3)], 3), [4, 5, 6]);
  assert.deepEqual(vals([N(2), N(4), N(6)], 2), [8, 10]);
  assert.deepEqual(vals([N(0.1), N(0.2)], 2), [0.3, 0.4]);
  assert.deepEqual(vals([N(5)], 3), [5, 5, 5]);            // single number copies
  assert.deepEqual(vals([N(5)], 3, { ctrl: true }), [6, 7, 8]); // Ctrl counts
  assert.deepEqual(vals([N(1), N(2)], 2, { ctrl: true }), [1, 2]); // Ctrl copies a series
  assert.deepEqual(vals([N(3), N(4)], 3, { reverse: true }), [2, 1, 0]);
  assert.deepEqual(vals([D("2024-02-27")], 3), ["2024-02-28", "2024-02-29", "2024-03-01"]);
  assert.deepEqual(vals([D("2024-01-31"), D("2024-02-29")], 2), ["2024-03-31", "2024-04-30"]);
  assert.deepEqual(vals([D("2024-01-01"), D("2024-01-08")], 2), ["2024-01-15", "2024-01-22"]);
  assert.deepEqual(vals([S("Monday")], 2), ["Tuesday", "Wednesday"]);
  assert.deepEqual(vals([S("Sat")], 3), ["Sun", "Mon", "Tue"]);
  assert.deepEqual(vals([S("Nov")], 3), ["Dec", "Jan", "Feb"]);
  assert.deepEqual(vals([S("MARCH")], 1), ["APRIL"]);
  assert.deepEqual(vals([S("Jan"), S("Apr")], 2), ["Jul", "Oct"]);
  assert.deepEqual(vals([S("Item 1")], 2), ["Item 2", "Item 3"]);
  assert.deepEqual(vals([S("Q1 2024")], 1), ["Q1 2025"]);
  assert.deepEqual(vals([S("Item 1"), S("Item 3")], 2), ["Item 5", "Item 7"]);
  assert.deepEqual(vals([S("ID-007")], 2), ["ID-008", "ID-009"]);
  assert.deepEqual(vals([S("hello")], 2), ["hello", "hello"]);
  assert.deepEqual(vals([S("a"), N(1)], 4), ["a", 1, "a", 1]);
  const f = fillSeries([{ t: "n", v: 2, f: "=A1*$B$1" }], 2, { translate: (x, d) => translateFormula(x, d, 0) });
  assert.deepEqual(f.map((c) => c.f), ["=A2*$B$1", "=A3*$B$1"]);
  const up = fillSeries([{ t: "n", v: 2, f: "=A5+1" }], 1, { reverse: true, translate: (x, d) => translateFormula(x, d, 0) });
  assert.equal(up[0].f, "=A4+1");
  // formats are copied with the values
  assert.deepEqual(fillSeries([{ t: "n", v: 1, z: "0.00", s: { b: true } }, N(2)], 1)[0], { t: "n", v: 3, z: "0.00", s: { b: true } });
});

test("charts: data from range, SVG for every type, geometry helpers are inverse", () => {
  const sheet = { colWidths: { B: 20 }, cells: { A1: S("Month"), B1: S("Sales"), C1: S("Cost"), A2: S("Jan"), B2: N(10), C2: N(4), A3: S("Feb"), B3: N(14), C3: S("x") } };
  const data = chartData(sheet, { range: "A1:C3" });
  assert.deepEqual(data.categories, ["Jan", "Feb"]);
  assert.deepEqual(data.series.map((s) => [s.name, s.values]), [["Sales", [10, 14]], ["Cost", [4, null]]]);
  assert.deepEqual(chartLayout("B2:B5").series, [{ c: 1 }]);
  for (const type of ["column", "bar", "line", "pie", "scatter", "area"]) {
    const svg = chartSvg(sheet, { type, range: "A1:C3", title: "T & <x>" }, 400, 250);
    assert.match(svg, /^<svg [^>]*width="400" height="250"/);
    assert.match(svg, /T &amp; &lt;x&gt;/);
    assert.ok(!/NaN|undefined/.test(svg), `${type}: ${svg.slice(0, 200)}`);
  }
  assert.match(chartSvg(sheet, { type: "column", range: "A1:A3" }, 300, 200), /No numbers to chart/);
  assert.equal(gridColPx(sheet, 0), 88);
  assert.equal(gridColPx(sheet, 1), 145);
  assert.equal(excelColPx(sheet, 0), 64);
  for (const [x, y] of [[0, 0], [100, 50], [300, 1234], [87, 23]]) {
    const a = chartAnchorFromXY(sheet, x, y);
    const back = chartXYFromAnchor(sheet, a.col, a.dx, a.row, a.dy);
    assert.ok(Math.abs(back.x - x) <= 1 && Math.abs(back.y - y) <= 1, `${x},${y} -> ${JSON.stringify(a)} -> ${JSON.stringify(back)}`);
  }
  assert.equal(offsetRef("A1:C3", 2, 1), "B3:D5");
  assert.equal(offsetRef("A1:C3", -1, 0), null);
});

test("local xlsx: freeze, autoFilter and charts survive write -> read", () => {
  const charts = ["column", "bar", "line", "pie", "scatter", "area"].map((type, i) => ({
    type, title: i % 2 ? `${type} chart` : undefined, range: type === "pie" ? "A1:B4" : "A1:C4", x: 300 + i * 7, y: 30 + i * 290, w: 420, h: 260,
  })).map((c) => { if (!c.title) delete c.title; return c; });
  const model = { kind: "excel", active: 0, sheets: [
    { name: "My Data", colWidths: { A: 14 }, freeze: { rows: 2, cols: 1 }, autoFilter: "A1:C4", charts,
      cells: { A1: S("Month"), B1: S("Sales"), C1: S("Cost"), A2: S("Jan"), B2: N(10), C2: N(4), A3: S("Feb"), B3: N(14), C3: N(3), A4: S("Mar"), B4: N(8), C4: N(6) } },
    { name: "Plain", freeze: { rows: 0, cols: 3 }, cells: { A1: N(1) } },
  ] };
  const bytes = LX.modelToBytes(model, "xlsx");
  const back = LX.bytesToModel(bytes, "x.xlsx");
  const sh = back.sheets[0];
  assert.deepEqual(sh.freeze, { rows: 2, cols: 1 });
  assert.equal(sh.autoFilter, "A1:C4");
  assert.deepEqual(back.sheets[1].freeze, { rows: 0, cols: 3 });
  assert.equal(back.sheets[1].autoFilter, undefined);
  assert.equal(sh.charts.length, 6);
  sh.charts.forEach((c, i) => {
    const o = charts[i];
    assert.equal(c.type, o.type);
    assert.equal(c.range, o.range);
    assert.equal(c.title, o.title);
    for (const k of ["x", "y", "w", "h"]) assert.ok(Math.abs(c[k] - o[k]) <= 2, `${o.type}.${k}: ${c[k]} vs ${o[k]}`);
  });
  assert.ok(!(back.warnings || []).some((w) => /chart|drawing|filter|frozen/i.test(w)), JSON.stringify(back.warnings));
  // the package is well formed: content types and relationships for every part
  const zip = globalThis.XLSX.CFB.read(bytes, { type: "array" });
  const text = (name) => new TextDecoder().decode(zip.FileIndex[zip.FullPaths.findIndex((p) => p.endsWith("/" + name))].content);
  const ct = text("[Content_Types].xml");
  for (let i = 1; i <= 6; i++) assert.ok(ct.includes(`/xl/charts/chart${i}.xml`));
  assert.ok(ct.includes("/xl/drawings/drawing1.xml"));
  assert.match(text("xl/worksheets/sheet1.xml"), /<pane xSplit="1" ySplit="2" topLeftCell="B3" activePane="bottomRight" state="frozen"\/>/);
  assert.match(text("xl/worksheets/sheet1.xml"), /<autoFilter ref="A1:C4"\/>[\s\S]*<drawing r:id="rIdDrawing1"\/><\/worksheet>/);
  assert.match(text("xl/charts/chart1.xml"), /<c:f>'My Data'!\$B\$2:\$B\$4<\/c:f>/);
});

test("local xlsx: reads freeze, filter and charts written by openpyxl (backend fixture)", () => {
  const file = resolvePath(XF_EXT, "../src/office-backend/tests/fixtures/excel/charts_freeze_filter.xlsx");
  if (!existsSync(file)) return; // fixture lives in the backend tree
  const m = LX.bytesToModel(new Uint8Array(readFileSync(file)), "charts_freeze_filter.xlsx");
  const sh = m.sheets[0];
  assert.deepEqual(sh.freeze, { rows: 1, cols: 1 });
  assert.equal(sh.autoFilter, "A1:D6");
  assert.deepEqual(sh.charts.map((c) => c.type), ["column", "bar", "line", "pie", "scatter", "area"]);
  for (const c of sh.charts) {
    assert.match(c.range, /^[A-D]1:[B-D]6$/);
    assert.ok(c.w > 100 && c.h > 100 && c.x > 0, JSON.stringify(c));
  }
  assert.ok(!(m.warnings || []).some((w) => /chart/i.test(w)), JSON.stringify(m.warnings));
});

test("local xlsx: unreadable charts are counted in one warning", () => {
  const model = { kind: "excel", sheets: [{ name: "A", cells: { A1: S("h"), B1: S("v"), A2: S("x"), B2: N(1) }, charts: [{ type: "column", range: "A1:B2", x: 0, y: 0, w: 300, h: 200 }] }] };
  const bytes = LX.modelToBytes(model, "xlsx");
  const zip = globalThis.XLSX.CFB.read(bytes, { type: "array" });
  const i = zip.FullPaths.findIndex((p) => p.endsWith("/xl/charts/chart1.xml"));
  const xml = new TextDecoder().decode(zip.FileIndex[i].content).replace(/<c:f>A!/g, "<c:f>Other!");
  zip.FileIndex[i].content = new TextEncoder().encode(xml);
  const patched = globalThis.XLSX.CFB.write(zip, { fileType: "zip", type: "array" });
  const m = LX.bytesToModel(new Uint8Array(patched), "x.xlsx");
  assert.equal(m.sheets[0].charts, undefined);
  assert.ok((m.warnings || []).some((w) => /^1 chart could not be read/.test(w)), JSON.stringify(m.warnings));
});
