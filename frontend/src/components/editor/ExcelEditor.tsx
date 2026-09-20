import React, { useEffect, useMemo, useRef, useState } from "react";
import {
  CellData,
  CellFormat,
  ExcelDocumentModel,
  ExcelWorksheet,
} from "../../types/office";
import { Bold, Italic, Plus } from "lucide-react";
import {
  columnName,
  formatCell,
  parseCoordinate,
  recalculateWorkbook,
} from "../../services/formulas";
import "../../styles/editors.css";

interface ExcelEditorProps {
  document: ExcelDocumentModel;
  onChangeDocument?: (doc: ExcelDocumentModel) => void;
  selectedCoord?: string | null;
  onSelectCell?: (coord: string) => void;
}

export const ExcelEditor: React.FC<ExcelEditorProps> = ({
  document,
  onChangeDocument,
  selectedCoord,
  onSelectCell,
}) => {
  const calculated = useMemo(() => recalculateWorkbook(document), [document]);
  const activeSheet =
    calculated.sheets.find((sheet) => sheet.id === calculated.activeSheetId) ||
    calculated.sheets[0];
  const [activeCoord, setActiveCoord] = useState(
    selectedCoord?.replace(/.*!/, "") || "A1",
  );
  const [editing, setEditing] = useState<"cell" | "formula" | null>(null);
  const [draft, setDraft] = useState("");
  const [notice, setNotice] = useState("");
  const gridRef = useRef<HTMLDivElement>(null);
  const cell = activeSheet?.cells[activeCoord];
  const raw = cell?.formula || (cell?.value == null ? "" : String(cell.value));
  useEffect(() => {
    if (selectedCoord) {
      const coord = selectedCoord.replace(/.*!/, "").toUpperCase();
      if (parseCoordinate(coord)) setActiveCoord(coord);
    }
  }, [selectedCoord]);
  useEffect(() => {
    setEditing(null);
  }, [activeSheet?.id]);
  useEffect(() => {
    gridRef.current
      ?.querySelector(`[data-cell="${activeCoord}"]`)
      ?.scrollIntoView({ block: "nearest", inline: "nearest" });
  }, [activeCoord, activeSheet?.id]);
  const commit = (next: ExcelDocumentModel) =>
    onChangeDocument?.(recalculateWorkbook(next));
  const replaceSheet = (sheet: ExcelWorksheet) =>
    commit({
      ...document,
      sheets: document.sheets.map((s) => (s.id === sheet.id ? sheet : s)),
    });
  const writeCell = (text: string, existing?: CellData): CellData => {
    const numeric = text.trim() !== "" && Number.isFinite(Number(text));
    return {
      ...existing,
      value: text.startsWith("=")
        ? null
        : numeric
          ? Number(text)
          : text.toUpperCase() === "TRUE"
            ? true
            : text.toUpperCase() === "FALSE"
              ? false
              : text,
      formula: text.startsWith("=") ? text : undefined,
      formatted: undefined,
    };
  };
  const changeValue = (text: string, coord = activeCoord) => {
    if (!activeSheet) return;
    replaceSheet({
      ...activeSheet,
      cells: {
        ...activeSheet.cells,
        [coord]: writeCell(text, activeSheet.cells[coord]),
      },
    });
  };
  const updateFormat = (patch: Partial<CellFormat>) => {
    if (!activeSheet) return;
    replaceSheet({
      ...activeSheet,
      cells: {
        ...activeSheet.cells,
        [activeCoord]: {
          ...cell,
          value: cell?.value ?? "",
          format: { ...cell?.format, ...patch },
        },
      },
    });
  };
  const select = (coord: string) => {
    setActiveCoord(coord);
    setEditing(null);
    onSelectCell?.(coord);
  };
  const navigate = (row: number, col: number) => {
    const current = parseCoordinate(activeCoord)!;
    select(
      `${columnName(Math.max(0, Math.min((activeSheet?.columns.length || 1) - 1, current.col + col)))}${Math.max(1, Math.min(activeSheet?.rowCount || 1, current.row + row))}`,
    );
    gridRef.current?.focus();
  };
  const beginEdit = (kind: "cell" | "formula", text = raw) => {
    setDraft(text);
    setEditing(kind);
  };
  const edit = (text: string) => {
    setDraft(text);
    changeValue(text);
  };
  const onEditKey: React.KeyboardEventHandler<HTMLInputElement> = (event) => {
    // Workspace shortcuts such as Save and Undo must reach the application while editing.
    if (event.ctrlKey || event.metaKey) return;
    if (event.key === "Enter" || event.key === "Tab") {
      event.preventDefault();
      navigate(
        event.key === "Enter" ? (event.shiftKey ? -1 : 1) : 0,
        event.key === "Tab" ? (event.shiftKey ? -1 : 1) : 0,
      );
    }
    if (event.key === "Escape") {
      event.preventDefault();
      setEditing(null);
      gridRef.current?.focus();
    }
    event.stopPropagation();
  };
  const onPaste: React.ClipboardEventHandler<HTMLDivElement> = (event) => {
    const text = event.clipboardData.getData("text/plain");
    if (
      !activeSheet ||
      (!text.includes("\t") &&
        !text.includes("\n") &&
        (event.target as HTMLElement).tagName === "INPUT")
    )
      return;
    event.preventDefault();
    const rows = text
      .replace(/\r/g, "")
      .replace(/\n$/, "")
      .split("\n")
      .map((row) => row.split("\t"));
    if (rows.length > 1000 || rows.some((row) => row.length > 100)) {
      setNotice(
        "Paste up to 1,000 rows and 100 columns at a time. No cells were changed.",
      );
      return;
    }
    setNotice("");
    const start = parseCoordinate(activeCoord)!;
    const columnCount = Math.max(
      activeSheet.columns.length,
      start.col + Math.max(...rows.map((row) => row.length)),
    );
    const rowCount = Math.max(
      activeSheet.rowCount,
      start.row + rows.length - 1,
    );
    if (
      columnCount > 100 ||
      rowCount > 1000 ||
      columnCount * rowCount > 10000
    ) {
      setNotice(
        "This editor supports up to 1,000 rows, 100 columns and 10,000 grid cells per sheet. Paste into a smaller range or a new sheet. No cells were changed.",
      );
      return;
    }
    const cells = { ...activeSheet.cells };
    rows.forEach((values, r) =>
      values.forEach((value, c) => {
        const coord = `${columnName(start.col + c)}${start.row + r}`;
        cells[coord] = writeCell(value, cells[coord]);
      }),
    );
    const count = Math.max(
      activeSheet.columns.length,
      start.col + Math.min(100, Math.max(...rows.map((row) => row.length))),
    );
    const columns = Array.from(
      { length: count },
      (_, index) =>
        activeSheet.columns[index] || {
          key: columnName(index),
          label: columnName(index),
          width: 110,
          type: "string" as const,
        },
    );
    replaceSheet({
      ...activeSheet,
      cells,
      columns,
      colCount: count,
      rowCount: Math.max(
        activeSheet.rowCount,
        start.row + Math.min(rows.length, 1000) - 1,
      ),
    });
    setEditing(null);
  };
  const addSheet = () => {
    if (document.sheets.length >= 50) {
      setNotice("This editor supports up to 50 worksheets per workbook.");
      return;
    }
    let number = document.sheets.length + 1;
    while (document.sheets.some((sheet) => sheet.name === `Sheet ${number}`))
      number++;
    const sheet: ExcelWorksheet = {
      id: `sheet-${crypto.randomUUID()}`,
      name: `Sheet ${number}`,
      rowCount: 25,
      colCount: 8,
      columns: Array.from({ length: 8 }, (_, index) => ({
        key: columnName(index),
        label: columnName(index),
        width: 110,
        type: "string",
      })),
      cells: {},
    };
    commit({
      ...document,
      activeSheetId: sheet.id,
      sheets: [...document.sheets, sheet],
    });
    select("A1");
  };
  if (!activeSheet)
    return (
      <div className="office-editor">
        <button className="vscode-btn" onClick={addSheet}>
          Create worksheet
        </button>
      </div>
    );
  return (
    <div className="office-editor">
      <div className="office-toolbar">
        <span className="office-badge excel-badge">XLSX</span>
        <span
          style={{
            minWidth: 38,
            textAlign: "center",
            color: "var(--vscode-editor-fg)",
            fontFamily: "monospace",
          }}
        >
          {activeCoord}
        </span>
        <span style={{ fontStyle: "italic", color: "#4ade80" }}>fx</span>
        <input
          aria-label="Formula bar"
          className="vscode-input"
          style={{
            flex: 1,
            minWidth: 170,
            fontFamily: "var(--vscode-mono-family)",
          }}
          value={editing === "formula" ? draft : raw}
          onFocus={() => beginEdit("formula")}
          onChange={(e) => edit(e.target.value)}
          onBlur={() => setEditing(null)}
          onKeyDown={onEditKey}
          placeholder="Value or formula: =SUM(A1:A10)"
        />
      </div>
      <div className="office-toolbar compact-toolbar">
        <button
          className="vscode-icon-btn"
          aria-label="Bold cell"
          title="Bold cell"
          aria-pressed={!!cell?.format?.bold}
          onClick={() => updateFormat({ bold: !cell?.format?.bold })}
        >
          <Bold size={14} />
        </button>
        <button
          className="vscode-icon-btn"
          aria-label="Italic cell"
          title="Italic cell"
          aria-pressed={!!cell?.format?.italic}
          onClick={() => updateFormat({ italic: !cell?.format?.italic })}
        >
          <Italic size={14} />
        </button>
        <select
          className="vscode-input"
          aria-label="Number format"
          value={cell?.format?.numberFormat || "general"}
          onChange={(e) =>
            updateFormat({
              numberFormat: e.target.value as CellFormat["numberFormat"],
            })
          }
        >
          <option value="general">General</option>
          <option value="number">Number</option>
          <option value="currency">Currency</option>
          <option value="percent">Percent</option>
        </select>
        <button
          className="vscode-btn-secondary"
          title="Add a row (up to 1,000 rows and 10,000 grid cells)"
          disabled={
            activeSheet.rowCount >= 1000 ||
            (activeSheet.rowCount + 1) * activeSheet.columns.length > 10000
          }
          onClick={() =>
            replaceSheet({ ...activeSheet, rowCount: activeSheet.rowCount + 1 })
          }
        >
          + Row
        </button>
        <button
          className="vscode-btn-secondary"
          title="Add a column (up to 100 columns and 10,000 grid cells)"
          disabled={
            activeSheet.columns.length >= 100 ||
            activeSheet.rowCount * (activeSheet.columns.length + 1) > 10000
          }
          onClick={() => {
            const key = columnName(activeSheet.columns.length);
            replaceSheet({
              ...activeSheet,
              colCount: activeSheet.columns.length + 1,
              columns: [
                ...activeSheet.columns,
                { key, label: key, width: 110, type: "string" },
              ],
            });
          }}
        >
          + Column
        </button>
        <span className="toolbar-spacer" />
        <span className="muted">
          Double-click or type to edit · Enter / Tab to move
        </span>
      </div>
      {notice && (
        <div role="status" className="office-toolbar">
          {notice}
          <button
            className="vscode-btn-secondary"
            onClick={() => setNotice("")}
          >
            Dismiss
          </button>
        </div>
      )}
      <div
        ref={gridRef}
        className="sheet-grid-wrap"
        tabIndex={0}
        aria-label="Worksheet grid"
        onPaste={onPaste}
        onCopy={(event) => {
          if ((event.target as HTMLElement).tagName !== "INPUT") {
            event.preventDefault();
            event.clipboardData.setData("text/plain", raw);
          }
        }}
        onKeyDown={(event) => {
          if (editing || event.ctrlKey || event.metaKey || event.altKey) return;
          const offsets: Record<string, [number, number]> = {
            ArrowUp: [-1, 0],
            ArrowDown: [1, 0],
            ArrowLeft: [0, -1],
            ArrowRight: [0, 1],
            Tab: [0, event.shiftKey ? -1 : 1],
            Enter: [event.shiftKey ? -1 : 1, 0],
          };
          if (offsets[event.key]) {
            event.preventDefault();
            navigate(...offsets[event.key]);
          } else if (event.key === "Backspace" || event.key === "Delete") {
            event.preventDefault();
            changeValue("");
          } else if (event.key === "F2") {
            event.preventDefault();
            beginEdit("cell");
          } else if (event.key.length === 1) {
            event.preventDefault();
            beginEdit("cell", event.key);
            changeValue(event.key);
          }
        }}
      >
        <table
          className="sheet-grid"
          aria-label={activeSheet.name}
          style={{
            width: activeSheet.columns.reduce(
              (sum, col) => sum + col.width,
              46,
            ),
          }}
        >
          <colgroup>
            <col style={{ width: 46 }} />
            {activeSheet.columns.map((col) => (
              <col key={col.key} style={{ width: col.width }} />
            ))}
          </colgroup>
          <thead>
            <tr>
              <th className="corner" />
              {activeSheet.columns.map((col) => (
                <th key={col.key} title={col.label}>
                  {col.key}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {Array.from({ length: activeSheet.rowCount }, (_, row) => (
              <tr key={row}>
                <th className="row-label">{row + 1}</th>
                {activeSheet.columns.map((col) => {
                  const coord = `${col.key}${row + 1}`;
                  const data = activeSheet.cells[coord];
                  const selected = coord === activeCoord;
                  return (
                    <td
                      key={coord}
                      data-cell={coord}
                      aria-label={`${coord}: ${data?.formatted || ""}`}
                      className={selected ? "selected-cell" : ""}
                      style={{
                        color:
                          data?.format?.textColor ||
                          (data?.format?.fill &&
                          /^#[def]/i.test(data.format.fill)
                            ? "#182330"
                            : "var(--vscode-editor-fg)"),
                        background: data?.format?.fill,
                        fontWeight: data?.format?.bold ? "bold" : "normal",
                        fontStyle: data?.format?.italic ? "italic" : "normal",
                        textAlign:
                          data?.format?.align ||
                          (typeof data?.value === "number" ? "right" : "left"),
                      }}
                      onClick={() => {
                        select(coord);
                        gridRef.current?.focus();
                      }}
                      onDoubleClick={() => {
                        select(coord);
                        beginEdit(
                          "cell",
                          data?.formula || String(data?.value ?? ""),
                        );
                      }}
                    >
                      {selected && editing === "cell" ? (
                        <input
                          autoFocus
                          aria-label={`Edit ${coord}`}
                          value={draft}
                          onClick={(e) => e.stopPropagation()}
                          onChange={(e) => edit(e.target.value)}
                          onBlur={() => setEditing(null)}
                          onKeyDown={onEditKey}
                        />
                      ) : data ? (
                        formatCell(data)
                      ) : (
                        ""
                      )}
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <div className="sheet-tabs">
        {document.sheets.map((sheet) => (
          <button
            key={sheet.id}
            className={`sheet-tab ${sheet.id === activeSheet.id ? "active" : ""}`}
            onClick={() => {
              commit({ ...document, activeSheetId: sheet.id });
              select("A1");
            }}
          >
            {sheet.name}
          </button>
        ))}
        <button
          className="vscode-icon-btn"
          title="Add worksheet (up to 50 sheets)"
          aria-label="Add worksheet"
          disabled={document.sheets.length >= 50}
          onClick={addSheet}
        >
          <Plus size={14} />
        </button>
        <span className="toolbar-spacer" />
        <span className="muted">
          {activeSheet.rowCount} rows × {activeSheet.columns.length} columns
        </span>
      </div>
    </div>
  );
};
