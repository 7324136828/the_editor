import type { CellData, CellValue, ExcelDocumentModel } from "../types/office";

type Result = CellValue | CellValue[];
const fail = (code: string): never => {
  throw new Error(code);
};
const errorValue = (value: unknown): boolean =>
  typeof value === "string" &&
  /^#(REF!|VALUE!|DIV\/0!|NAME\?|CYCLE!|NUM!|ERROR!)$/.test(value);
const number = (value: Result): number => {
  if (Array.isArray(value)) return fail("#VALUE!");
  if (errorValue(value)) return fail(String(value));
  if (value === null || value === "") return 0;
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : fail("#VALUE!");
};

export const columnName = (index: number): string => {
  let name = "";
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26))
    name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  return name;
};

export const parseCoordinate = (
  coord: string,
): { col: number; row: number } | null => {
  const match = /^\$?([A-Z]+)\$?([1-9]\d*)$/i.exec(coord);
  if (!match) return null;
  const col =
    [...match[1].toUpperCase()].reduce(
      (n, char) => n * 26 + char.charCodeAt(0) - 64,
      0,
    ) - 1;
  const row = Number(match[2]);
  return col < 16384 && row <= 1048576 ? { col, row } : null;
};

/** Deliberately small spreadsheet language; formula text is never executable JavaScript. */
export function evaluateFormula(
  formula: string,
  resolve: (coord: string, sheet?: string) => CellValue,
): CellValue {
  try {
    if (formula.length > 8192) return "#ERROR!";
    const source = formula.replace(/^=/, "");
    const tokens: string[] = [];
    let cursor = 0;
    while (cursor < source.length) {
      if (/\s/.test(source[cursor])) {
        cursor++;
        continue;
      }
      const match =
        /^(?:\d*\.?\d+(?:[eE][+-]?\d+)?|\$?[A-Za-z_][A-Za-z_0-9.$]*|'(?:[^']|'')*'|"(?:[^"]|"")*"|[+\-*/^(),:!%])/.exec(
          source.slice(cursor),
        );
      if (!match) return "#ERROR!";
      tokens.push(match[0]);
      cursor += match[0].length;
    }
    let pos = 0;
    let depth = 0;
    const take = () => tokens[pos++];
    const expect = (token: string) => {
      if (take() !== token) fail("#ERROR!");
    };
    const reference = (token: string, sheet?: string): Result => {
      const start = parseCoordinate(token);
      if (!start) return fail("#REF!");
      if (tokens[pos] !== ":")
        return resolve(`${columnName(start.col)}${start.row}`, sheet);
      take();
      const end = parseCoordinate(take() || "");
      if (!end) return fail("#REF!");
      const values: CellValue[] = [];
      const area =
        (Math.abs(start.col - end.col) + 1) *
        (Math.abs(start.row - end.row) + 1);
      if (area > 10000) return fail("#NUM!");
      for (
        let row = Math.min(start.row, end.row);
        row <= Math.max(start.row, end.row);
        row++
      ) {
        for (
          let col = Math.min(start.col, end.col);
          col <= Math.max(start.col, end.col);
          col++
        )
          values.push(resolve(`${columnName(col)}${row}`, sheet));
      }
      return values;
    };
    const primary = (): Result => {
      if (++depth > 128) return fail("#ERROR!");
      try {
        const token = take();
        if (!token) return fail("#ERROR!");
        if (token === "(") {
          const value = expression();
          expect(")");
          return value;
        }
        if (/^(?:\d|\.)/.test(token)) return Number(token);
        if (token.startsWith('"'))
          return token.slice(1, -1).replace(/""/g, '"');
        if (tokens[pos] === "!") {
          take();
          return reference(
            take() || "",
            token.startsWith("'")
              ? token.slice(1, -1).replace(/''/g, "'")
              : token,
          );
        }
        if (tokens[pos] === "(") {
          take();
          const args: Result[] = [];
          if (tokens[pos] !== ")") {
            args.push(expression());
            while (tokens[pos] === ",") {
              take();
              args.push(expression());
            }
          }
          expect(")");
          const flat = args.flat();
          const error = flat.find(errorValue);
          if (error) return fail(String(error));
          const nums = flat.filter(
            (value): value is number => typeof value === "number",
          );
          switch (token.toUpperCase()) {
            case "SUM":
              return nums.reduce((sum, n) => sum + n, 0);
            case "AVERAGE":
              return nums.length
                ? nums.reduce((sum, n) => sum + n, 0) / nums.length
                : fail("#DIV/0!");
            case "MIN":
              return nums.length ? Math.min(...nums) : 0;
            case "MAX":
              return nums.length ? Math.max(...nums) : 0;
            case "COUNT":
              return nums.length;
            case "ABS":
              return args.length === 1
                ? Math.abs(number(args[0]))
                : fail("#VALUE!");
            case "ROUND": {
              if (args.length < 1 || args.length > 2) return fail("#VALUE!");
              const factor = 10 ** number(args[1] ?? 0);
              return Math.round(number(args[0]) * factor) / factor;
            }
            default:
              return fail("#NAME?");
          }
        }
        if (token.toUpperCase() === "TRUE") return true;
        if (token.toUpperCase() === "FALSE") return false;
        return /^[A-Za-z$]+\$?\d+$/.test(token)
          ? reference(token)
          : fail("#NAME?");
      } finally {
        depth--;
      }
    };
    const unary = (): Result => {
      if (tokens[pos] === "+" || tokens[pos] === "-") {
        const sign = take();
        return (sign === "-" ? -1 : 1) * number(unary());
      }
      let value = primary();
      while (tokens[pos] === "%") {
        take();
        value = number(value) / 100;
      }
      return value;
    };
    const power = (): Result => {
      const left = unary();
      return tokens[pos] === "^"
        ? (take(), number(left) ** number(power()))
        : left;
    };
    const term = (): Result => {
      let left = power();
      while (tokens[pos] === "*" || tokens[pos] === "/") {
        const operator = take();
        const right = number(power());
        if (operator === "/" && right === 0) return fail("#DIV/0!");
        left = operator === "*" ? number(left) * right : number(left) / right;
      }
      return left;
    };
    const expression = (): Result => {
      let left = term();
      while (tokens[pos] === "+" || tokens[pos] === "-") {
        const operator = take();
        const right = number(term());
        left = operator === "+" ? number(left) + right : number(left) - right;
      }
      return left;
    };
    const value = expression();
    if (pos !== tokens.length || Array.isArray(value)) return "#ERROR!";
    if (typeof value === "number" && !Number.isFinite(value)) return "#NUM!";
    return value;
  } catch (error) {
    return error instanceof Error && errorValue(error.message)
      ? error.message
      : "#ERROR!";
  }
}

export function formatCell(cell: CellData): string {
  const value = cell.value;
  if (value === null || value === undefined) return "";
  if (typeof value !== "number") return String(value);
  const decimals = Math.max(
    0,
    Math.min(
      10,
      cell.format?.decimalPlaces ??
        (cell.format?.numberFormat === "percent" ? 1 : 2),
    ),
  );
  switch (cell.format?.numberFormat) {
    case "currency":
      return value.toLocaleString("en-US", {
        style: "currency",
        currency: "USD",
        maximumFractionDigits: decimals,
        minimumFractionDigits: decimals,
      });
    case "percent":
      return value.toLocaleString("en-US", {
        style: "percent",
        maximumFractionDigits: decimals,
        minimumFractionDigits: decimals,
      });
    case "number":
      return value.toLocaleString("en-US", {
        maximumFractionDigits: decimals,
        minimumFractionDigits: decimals,
      });
    default:
      return value.toLocaleString("en-US", { maximumFractionDigits: 10 });
  }
}

export function recalculateWorkbook(
  document: ExcelDocumentModel,
): ExcelDocumentModel {
  const sheets = document.sheets.map((sheet) => ({
    ...sheet,
    cells: Object.fromEntries(
      Object.entries(sheet.cells).map(([coord, cell]) => [coord, { ...cell }]),
    ),
  }));
  const cache = new Map<string, CellValue>();
  const visiting = new Set<string>();
  const audit: ExcelDocumentModel["formulasAudit"] = [];
  const resolve = (sheetId: string, coord: string): CellValue => {
    const key = `${sheetId}!${coord}`;
    if (cache.has(key)) return cache.get(key)!;
    if (visiting.has(key)) return "#CYCLE!";
    if (visiting.size > 256) return "#ERROR!";
    const sheet = sheets.find((s) => s.id === sheetId);
    if (!sheet) return "#REF!";
    const cell = sheet.cells[coord];
    if (!cell) return null;
    visiting.add(key);
    const dependencies = new Set<string>();
    const value = cell.formula
      ? evaluateFormula(cell.formula, (reference, name) => {
          const target = name
            ? sheets.find((s) => s.name.toLowerCase() === name.toLowerCase())
            : sheet;
          dependencies.add(name ? `${name}!${reference}` : reference);
          return target ? resolve(target.id, reference) : "#REF!";
        })
      : cell.value;
    visiting.delete(key);
    cache.set(key, value);
    cell.value = value;
    cell.formatted = formatCell(cell);
    if (cell.formula && sheetId === document.activeSheetId)
      audit.push({
        cellCoord: coord,
        formula: cell.formula,
        evaluatedValue: String(value ?? ""),
        dependencies: [...dependencies],
      });
    return value;
  };
  sheets.forEach((sheet) =>
    Object.keys(sheet.cells).forEach((coord) => resolve(sheet.id, coord)),
  );
  return { ...document, sheets, formulasAudit: audit };
}
