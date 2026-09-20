import type { CellFormat, CellMap, CellValue, SalesPoint } from './types';

const ERROR_CYCLE = '#CYCLE!';
const ERROR_DIV0 = '#DIV/0!';
const ERROR_NA = '#N/A';
const ERROR_VALUE = '#VALUE!';

export function normalizeRef(ref: string): string {
  return ref.replace(/\$/g, '').toUpperCase();
}

type Token =
  | { t: 'num'; v: number }
  | { t: 'str'; v: string }
  | { t: 'ref'; v: string }
  | { t: 'name'; v: string }
  | { t: 'op'; v: string };

function tokenize(src: string): Token[] {
  const tokens: Token[] = [];
  let i = 0;
  while (i < src.length) {
    const c = src[i];
    if (/\s/.test(c)) {
      i++;
      continue;
    }
    if (/[0-9.]/.test(c)) {
      const m = /^[0-9]*\.?[0-9]+/.exec(src.slice(i));
      if (!m) throw new Error('bad number');
      tokens.push({ t: 'num', v: parseFloat(m[0]) });
      i += m[0].length;
      continue;
    }
    if (c === '"') {
      const end = src.indexOf('"', i + 1);
      if (end < 0) throw new Error('unterminated string');
      tokens.push({ t: 'str', v: src.slice(i + 1, end) });
      i = end + 1;
      continue;
    }
    if (c === '$' || /[A-Za-z]/.test(c)) {
      const refMatch = /^\$?[A-Za-z]{1,3}\$?[0-9]+/.exec(src.slice(i));
      if (refMatch) {
        tokens.push({ t: 'ref', v: normalizeRef(refMatch[0]) });
        i += refMatch[0].length;
        continue;
      }
      const nameMatch = /^[A-Za-z_][A-Za-z0-9_.]*/.exec(src.slice(i));
      if (nameMatch) {
        tokens.push({ t: 'name', v: nameMatch[0].toUpperCase() });
        i += nameMatch[0].length;
        continue;
      }
      throw new Error('bad token');
    }
    if ('+-*/(),:'.includes(c)) {
      tokens.push({ t: 'op', v: c });
      i++;
      continue;
    }
    throw new Error('bad token');
  }
  return tokens;
}

type Range = { range: string[] };
type Arg = CellValue | Range;

function isError(v: Arg): v is string {
  return typeof v === 'string' && v.startsWith('#');
}

function colToIndex(col: string): number {
  let n = 0;
  for (const ch of col) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

function indexToCol(idx: number): string {
  let s = '';
  while (idx > 0) {
    const r = (idx - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    idx = Math.floor((idx - 1) / 26);
  }
  return s;
}

function splitRef(ref: string): { col: number; row: number } {
  const m = /^([A-Z]+)([0-9]+)$/.exec(ref);
  if (!m) throw new Error('bad ref');
  return { col: colToIndex(m[1]), row: parseInt(m[2], 10) };
}

function expandRange(a: string, b: string): Range {
  const p = splitRef(a);
  const q = splitRef(b);
  const refs: string[] = [];
  const c1 = Math.min(p.col, q.col);
  const c2 = Math.max(p.col, q.col);
  const r1 = Math.min(p.row, q.row);
  const r2 = Math.max(p.row, q.row);
  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) {
      refs.push(`${indexToCol(c)}${r}`);
    }
  }
  return { range: refs };
}

function rangeMatrix(range: Range): string[][] {
  const first = splitRef(range.range[0]);
  const last = splitRef(range.range[range.range.length - 1]);
  const rows = Math.abs(last.row - first.row) + 1;
  const cols = Math.abs(last.col - first.col) + 1;
  const matrix: string[][] = [];
  for (let r = 0; r < rows; r++) {
    matrix.push(range.range.slice(r * cols, (r + 1) * cols));
  }
  return matrix;
}

class Parser {
  private pos = 0;
  constructor(
    private tokens: Token[],
    private cells: CellMap,
    private seen: Set<string>,
  ) {}

  private peek(): Token | undefined {
    return this.tokens[this.pos];
  }

  private next(): Token | undefined {
    return this.tokens[this.pos++];
  }

  private expectOp(op: string): void {
    const t = this.next();
    if (!t || t.t !== 'op' || t.v !== op) throw new Error(`expected ${op}`);
  }

  parse(): Arg {
    const v = this.parseExpr();
    if (this.pos !== this.tokens.length) throw new Error('trailing tokens');
    return v;
  }

  private parseExpr(): Arg {
    let left = this.parseTerm();
    for (;;) {
      const t = this.peek();
      if (t && t.t === 'op' && (t.v === '+' || t.v === '-')) {
        this.next();
        const right = this.parseTerm();
        left = this.applyArithmetic(left, right, t.v);
        if (isError(left)) return left;
      } else {
        return left;
      }
    }
  }

  private parseTerm(): Arg {
    let left = this.parseFactor();
    for (;;) {
      const t = this.peek();
      if (t && t.t === 'op' && (t.v === '*' || t.v === '/')) {
        this.next();
        const right = this.parseFactor();
        left = this.applyArithmetic(left, right, t.v);
        if (isError(left)) return left;
      } else {
        return left;
      }
    }
  }

  private parseFactor(): Arg {
    const t = this.peek();
    if (t && t.t === 'op' && t.v === '-') {
      this.next();
      const v = this.parseFactor();
      return this.negate(v);
    }
    if (t && t.t === 'op' && t.v === '+') {
      this.next();
      return this.parseFactor();
    }
    return this.parsePrimary();
  }

  private parsePrimary(): Arg {
    const t = this.next();
    if (!t) throw new Error('unexpected end');
    if (t.t === 'num') return t.v;
    if (t.t === 'str') return t.v;
    if (t.t === 'op' && t.v === '(') {
      const v = this.parseExpr();
      this.expectOp(')');
      return v;
    }
    if (t.t === 'ref') {
      const next = this.peek();
      if (next && next.t === 'op' && next.v === ':') {
        this.next();
        const b = this.next();
        if (!b || b.t !== 'ref') throw new Error('bad range');
        return expandRange(t.v, b.v);
      }
      return evaluateCell(t.v, this.cells, this.seen);
    }
    if (t.t === 'name') {
      const next = this.peek();
      if (next && next.t === 'op' && next.v === '(') {
        this.next();
        const args: Arg[] = [];
        const closing = this.peek();
        if (!(closing && closing.t === 'op' && closing.v === ')')) {
          for (;;) {
            args.push(this.parseExpr());
            const sep = this.next();
            if (!sep || sep.t !== 'op' || (sep.v !== ',' && sep.v !== ')')) {
              throw new Error('bad args');
            }
            if (sep.v === ')') break;
          }
        } else {
          this.next();
        }
        return this.callFunction(t.v, args);
      }
      if (t.v === 'TRUE') return 'TRUE';
      if (t.v === 'FALSE') return 'FALSE';
      return ERROR_VALUE;
    }
    throw new Error('unexpected token');
  }

  private toNumber(v: Arg): number | string {
    if (typeof v === 'object') return ERROR_VALUE;
    if (isError(v)) return v;
    if (typeof v === 'number') return v;
    if (v === '') return 0;
    const n = Number(v);
    return Number.isFinite(n) ? n : ERROR_VALUE;
  }

  private negate(v: Arg): Arg {
    const n = this.toNumber(v);
    if (typeof n === 'string') return n;
    return -n;
  }

  private applyArithmetic(a: Arg, b: Arg, op: string): Arg {
    const na = this.toNumber(a);
    if (typeof na === 'string') return na;
    const nb = this.toNumber(b);
    if (typeof nb === 'string') return nb;
    switch (op) {
      case '+':
        return na + nb;
      case '-':
        return na - nb;
      case '*':
        return na * nb;
      case '/':
        return nb === 0 ? ERROR_DIV0 : na / nb;
      default:
        return ERROR_VALUE;
    }
  }

  private flattenValues(args: Arg[]): CellValue[] | string {
    const out: CellValue[] = [];
    for (const arg of args) {
      if (typeof arg === 'object') {
        for (const ref of arg.range) {
          const v = evaluateCell(ref, this.cells, this.seen);
          if (isError(v)) return v;
          out.push(v);
        }
      } else {
        if (isError(arg)) return arg;
        out.push(arg);
      }
    }
    return out;
  }

  private callFunction(name: string, args: Arg[]): Arg {
    switch (name) {
      case 'SUM': {
        const values = this.flattenValues(args);
        if (typeof values === 'string') return values;
        return values.reduce<number>((acc, v) => acc + (typeof v === 'number' ? v : 0), 0);
      }
      case 'SUMIFS': {
        const [sumRange, critRange, criteria] = args;
        if (typeof sumRange !== 'object' || typeof critRange !== 'object') return ERROR_VALUE;
        if (isError(criteria) || typeof criteria === 'object') return ERROR_VALUE;
        if (sumRange.range.length !== critRange.range.length) return ERROR_VALUE;
        let total = 0;
        for (let i = 0; i < critRange.range.length; i++) {
          const cv = evaluateCell(critRange.range[i], this.cells, this.seen);
          if (isError(cv)) return cv;
          if (String(cv) === String(criteria)) {
            const sv = evaluateCell(sumRange.range[i], this.cells, this.seen);
            if (isError(sv)) return sv;
            if (typeof sv === 'number') total += sv;
          }
        }
        return total;
      }
      case 'AVERAGEIF': {
        const [critRange, criteria, avgRange] = args;
        if (typeof critRange !== 'object' || typeof avgRange !== 'object') return ERROR_VALUE;
        if (isError(criteria) || typeof criteria === 'object') return ERROR_VALUE;
        if (critRange.range.length !== avgRange.range.length) return ERROR_VALUE;
        let total = 0;
        let count = 0;
        for (let i = 0; i < critRange.range.length; i++) {
          const cv = evaluateCell(critRange.range[i], this.cells, this.seen);
          if (isError(cv)) return cv;
          if (String(cv) === String(criteria)) {
            const av = evaluateCell(avgRange.range[i], this.cells, this.seen);
            if (isError(av)) return av;
            if (typeof av === 'number') {
              total += av;
              count++;
            }
          }
        }
        return count === 0 ? ERROR_DIV0 : total / count;
      }
      case 'VLOOKUP': {
        const [lookup, table, indexArg] = args;
        if (typeof table !== 'object') return ERROR_VALUE;
        if (isError(lookup) || typeof lookup === 'object') return ERROR_VALUE;
        const index = typeof indexArg === 'number' ? Math.trunc(indexArg) : NaN;
        if (!Number.isFinite(index) || index < 1) return ERROR_VALUE;
        const matrix = rangeMatrix(table);
        if (index > matrix[0].length) return ERROR_VALUE;
        for (const row of matrix) {
          const cv = evaluateCell(row[0], this.cells, this.seen);
          if (isError(cv)) return cv;
          if (String(cv) === String(lookup)) {
            return evaluateCell(row[index - 1], this.cells, this.seen);
          }
        }
        return ERROR_NA;
      }
      case 'LEFT': {
        const [source, countArg] = args;
        if (isError(source) || typeof source === 'object') return ERROR_VALUE;
        const n = typeof countArg === 'number' ? Math.trunc(countArg) : NaN;
        if (!Number.isFinite(n) || n < 0) return ERROR_VALUE;
        return String(source).slice(0, n);
      }
      case 'CONCAT': {
        const values = this.flattenValues(args);
        if (typeof values === 'string') return values;
        return values.map((v) => String(v)).join('');
      }
      default:
        return ERROR_VALUE;
    }
  }
}

export function evaluateCell(ref: string, cells: CellMap, seen: Set<string> = new Set()): CellValue {
  const key = normalizeRef(ref);
  if (seen.has(key)) return ERROR_CYCLE;
  const entry = cells[key];
  if (!entry) return 0;
  if (entry.formula === undefined) return entry.value ?? 0;
  seen.add(key);
  try {
    const src = entry.formula.startsWith('=') ? entry.formula.slice(1) : entry.formula;
    const result = new Parser(tokenize(src), cells, seen).parse();
    if (typeof result === 'object') return ERROR_VALUE;
    return result;
  } catch {
    return ERROR_VALUE;
  } finally {
    seen.delete(key);
  }
}

export function getSalesPoints(cells: CellMap): SalesPoint[] {
  const points: SalesPoint[] = [];
  for (let row = 2; row <= 25; row++) {
    const productId = evaluateCell(`A${row}`, cells);
    const region = evaluateCell(`B${row}`, cells);
    const unitPrice = evaluateCell(`C${row}`, cells);
    const unitsSold = evaluateCell(`D${row}`, cells);
    const grossSales = evaluateCell(`E${row}`, cells);
    const category = evaluateCell(`F${row}`, cells);
    if (
      typeof unitPrice !== 'number' ||
      typeof unitsSold !== 'number' ||
      typeof grossSales !== 'number' ||
      !Number.isFinite(unitPrice) ||
      !Number.isFinite(unitsSold) ||
      !Number.isFinite(grossSales)
    ) {
      continue;
    }
    points.push({
      row,
      productId: String(productId),
      region: String(region),
      unitPrice,
      unitsSold,
      grossSales,
      category: String(category),
    });
  }
  return points;
}

const currencyFormatter = new Intl.NumberFormat('en-US', {
  style: 'currency',
  currency: 'USD',
});
const percentFormatter = new Intl.NumberFormat('en-US', {
  style: 'percent',
  maximumFractionDigits: 0,
});
const generalFormatter = new Intl.NumberFormat('en-US', {
  maximumFractionDigits: 2,
});

export function formatCellValue(value: CellValue, format?: CellFormat['number']): string {
  if (typeof value === 'string') return value;
  if (!Number.isFinite(value)) return String(value);
  switch (format) {
    case 'currency':
      return currencyFormatter.format(value);
    case 'percent':
      return percentFormatter.format(value);
    default:
      return generalFormatter.format(value);
  }
}
