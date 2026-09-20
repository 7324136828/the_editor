import assert from 'node:assert/strict';
import test from 'node:test';
import { evaluateFormula, recalculateWorkbook, parseCoordinate, columnName, formatCell } from '../src/services/formulas.ts';

const evaluate = (formula, values = {}) => evaluateFormula(formula, coord => values[coord] ?? null);
const sheet = (id, name, cells) => ({ id, name, cells, rowCount: 10, colCount: 3, columns: ['A', 'B', 'C'].map(key => ({ key, label: key, width: 100, type: 'string' })) });
const book = sheets => ({ id: 'workbook', title: 'Test', activeSheetId: sheets[0].id, sheets, formulasAudit: [] });

test('arithmetic precedence, parentheses, unary signs, percent and exponent', () => {
  assert.equal(evaluate('=2+3*4'), 14);
  assert.equal(evaluate('=(2+3)*4'), 20);
  assert.equal(evaluate('=-2+10/2'), 3);
  assert.equal(evaluate('=2^3^2'), 512);
  assert.equal(evaluate('=200*10%'), 20);
  assert.equal(evaluate('=1.5e2 + .5'), 150.5);
});

test('real aggregates, nested functions, ranges, blank and text cells', () => {
  const values = { A1: 5, A2: 10, A3: 'Notes', B1: 15, B2: 20 };
  assert.equal(evaluate('=SUM(A1:B3)', values), 50);
  assert.equal(evaluate('=AVERAGE(A1:B3)', values), 12.5);
  assert.equal(evaluate('=MIN(A1:B3)', values), 5);
  assert.equal(evaluate('=MAX(A1:B3)', values), 20);
  assert.equal(evaluate('=COUNT(A1:B3)', values), 4);
  assert.equal(evaluate('=SUM(A2:A1,MAX(B1:B2),2)', values), 37);
  assert.equal(evaluate('=ROUND(AVERAGE(A1:A2),0)', values), 8);
  assert.equal(evaluate('=ABS(-5)'), 5);
  assert.equal(evaluate('=$A$1+A2', values), 15);
});

test('formulas reject executable syntax and return spreadsheet errors', () => {
  assert.equal(evaluate('=1/0'), '#DIV/0!');
  assert.equal(evaluate('=AVERAGE(A1:A2)'), '#DIV/0!');
  assert.equal(evaluate('=A1+2', { A1: 'text' }), '#VALUE!');
  assert.equal(evaluate('=NOPE(2)'), '#NAME?');
  assert.equal(evaluate('=SUM(A1:A999999)'), '#NUM!');
  assert.equal(evaluate('=2+'), '#ERROR!');
  assert.equal(evaluate('=globalThis.process.exit()'), '#NAME?');
  assert.equal(evaluate('=constructor.constructor("return 1")()'), '#NAME?');
  assert.equal(evaluate('=1;alert(1)'), '#ERROR!');
  assert.equal(evaluate('=1e308*1e308'), '#NUM!');
});

test('dependency graph recalculates transitively without changing source', () => {
  const original = book([sheet('one', 'Sheet 1', { A1: { value: 3 }, B1: { value: 0, formula: '=A1*2' }, C1: { value: 0, formula: '=SUM(A1:B1)' } })]);
  const initial = recalculateWorkbook(original);
  assert.equal(initial.sheets[0].cells.C1.value, 9);
  assert.equal(original.sheets[0].cells.C1.value, 0);
  const changed = recalculateWorkbook({ ...initial, sheets: [{ ...initial.sheets[0], cells: { ...initial.sheets[0].cells, A1: { value: 7 } } }] });
  assert.equal(changed.sheets[0].cells.B1.value, 14);
  assert.equal(changed.sheets[0].cells.C1.value, 21);
  assert.deepEqual(changed.formulasAudit.find(entry => entry.cellCoord === 'C1').dependencies, ['A1', 'B1']);
});

test('cross-sheet references and cycles are handled deterministically', () => {
  const result = recalculateWorkbook(book([
    sheet('one', 'First', { A1: { value: 2 }, B1: { value: 0, formula: "='Second Sheet'!A1+1" }, C1: { value: 0, formula: '=Missing!A1' } }),
    sheet('two', 'Second Sheet', { A1: { value: 0, formula: '=First!A1*3' }, B1: { value: 0, formula: '=C1' }, C1: { value: 0, formula: '=B1' } }),
  ]));
  assert.equal(result.sheets[0].cells.B1.value, 7);
  assert.equal(result.sheets[0].cells.C1.value, '#REF!');
  assert.equal(result.sheets[1].cells.B1.value, '#CYCLE!');
  assert.equal(result.sheets[1].cells.C1.value, '#CYCLE!');
  const self = recalculateWorkbook(book([sheet('one', 'First', { A1: { value: 0, formula: '=A1+1' } })]));
  assert.equal(self.sheets[0].cells.A1.value, '#CYCLE!');
});

test('coordinates extend beyond Z and numeric formats remain accurate', () => {
  assert.equal(columnName(26), 'AA');
  assert.deepEqual(parseCoordinate('$AA$12'), { col: 26, row: 12 });
  assert.equal(parseCoordinate('A0'), null);
  assert.equal(parseCoordinate('XFE1'), null);
  assert.equal(formatCell({ value: 0, format: { numberFormat: 'currency' } }), '$0.00');
  assert.equal(formatCell({ value: .325, format: { numberFormat: 'percent' } }), '32.5%');
});
