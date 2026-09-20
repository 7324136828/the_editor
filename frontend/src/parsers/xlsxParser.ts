import * as XLSX from "xlsx";
import type {
  ExcelDocumentModel,
  ExcelWorksheet,
  ExcelColumn,
  CellData,
  ExcelFormulaAudit,
} from "../types/office";

export async function parseXlsxFile(file: File): Promise<ExcelDocumentModel> {
  if (file.size > 25 * 1024 * 1024)
    throw new Error("Spreadsheet imports are limited to 25 MB.");
  const bytes = new Uint8Array(await file.arrayBuffer());
  const extension = file.name.split(".").pop()?.toLowerCase();
  if (
    (extension === "xlsx" && (bytes[0] !== 80 || bytes[1] !== 75)) ||
    (extension === "xls" && (bytes[0] !== 208 || bytes[1] !== 207))
  )
    throw new Error(
      "The file contents do not match its spreadsheet extension. Use XLSX, XLS, or CSV with the correct extension.",
    );
  let workbook: XLSX.WorkBook;
  try {
    workbook = XLSX.read(bytes, {
      type: "array",
      cellFormula: true,
      cellStyles: true,
      cellNF: true,
    });
  } catch {
    throw new Error(
      "This spreadsheet could not be read. Check that it is a valid, unencrypted XLSX, XLS, or CSV file.",
    );
  }
  if (!workbook.SheetNames.length || workbook.SheetNames.length > 50)
    throw new Error("A workbook must contain between 1 and 50 worksheets.");
  const sheets: ExcelWorksheet[] = [];
  const formulasAudit: ExcelFormulaAudit[] = [];
  for (const [index, sheetName] of workbook.SheetNames.entries()) {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet)
      throw new Error("The workbook contains an unreadable worksheet.");
    const range = XLSX.utils.decode_range(sheet["!ref"] || "A1");
    // Preserve absolute coordinates when the used range begins at e.g. D10.
    const colCount = Math.max(range.e.c + 1, 8);
    const rowCount = Math.max(range.e.r + 1, 20);
    if (
      !Number.isFinite(colCount) ||
      !Number.isFinite(rowCount) ||
      colCount > 100 ||
      rowCount > 1000 ||
      colCount * rowCount > 10000
    )
      throw new Error(
        `Worksheet "${sheetName}" exceeds the editor limit: 1,000 rows, 100 columns, and 10,000 displayed cells. Import a smaller range.`,
      );
    const columns: ExcelColumn[] = Array.from(
      { length: colCount },
      (_, column) => {
        const key = XLSX.utils.encode_col(column);
        const width =
          sheet["!cols"]?.[column]?.wpx ||
          (sheet["!cols"]?.[column]?.wch || 16) * 7;
        return {
          key,
          label: sheet[`${key}1`] ? String(sheet[`${key}1`].v ?? key) : key,
          width: Math.max(45, Math.min(500, width)),
          type: "string",
        };
      },
    );
    const cells: Record<string, CellData> = {};
    if (Object.keys(sheet).length > 10010)
      throw new Error(
        `Worksheet "${sheetName}" exceeds the 10,000 cell import limit.`,
      );
    for (const [coord, raw] of Object.entries(sheet)) {
      if (coord.startsWith("!") || !/^[A-Z]+[1-9][0-9]*$/.test(coord)) continue;
      const position = XLSX.utils.decode_cell(coord);
      if (position.r >= rowCount || position.c >= colCount)
        throw new Error(
          "A worksheet cell falls outside its declared used range.",
        );
      const rawCell = raw as XLSX.CellObject;
      const value =
        rawCell.t === "e" ? rawCell.w || "#ERROR!" : (rawCell.v ?? null);
      const formula = rawCell.f ? `=${rawCell.f}` : undefined;
      const formatCode = typeof rawCell.z === "string" ? rawCell.z : "";
      const numberFormat = /%/.test(formatCode)
        ? "percent"
        : /[$€£¥]/.test(formatCode)
          ? "currency"
          : typeof value === "number"
            ? "number"
            : "general";
      const decimalPlaces = formatCode.match(/\.([0#]+)/)?.[1].length;
      const style = rawCell.s as
        { fgColor?: { rgb?: string }; patternType?: string } | undefined;
      const fill = style?.fgColor?.rgb?.slice(-6);
      cells[coord] = {
        value:
          value instanceof Date
            ? value.toISOString()
            : (value as CellData["value"]),
        formula,
        formatted: rawCell.w || String(value ?? ""),
        format: {
          align: typeof value === "number" ? "right" : "left",
          numberFormat,
          ...(decimalPlaces !== undefined ? { decimalPlaces } : {}),
          ...(fill && style?.patternType !== "none"
            ? { fill: `#${fill}` }
            : {}),
        },
      };
      if (formula)
        formulasAudit.push({
          cellCoord: `${sheetName}!${coord}`,
          formula,
          evaluatedValue: rawCell.w || String(value ?? ""),
          dependencies: [],
        });
    }
    sheets.push({
      id: `sheet-${index + 1}`,
      name: sheetName,
      rowCount,
      colCount,
      columns,
      cells,
    });
  }
  return {
    id: `doc-excel-${crypto.randomUUID()}`,
    title: file.name,
    activeSheetId: sheets[0].id,
    sheets,
    formulasAudit,
  };
}
