import React from "react";
import { OfficeDocument } from "../../types/office";
import { columnName } from "../../services/formulas";
export const MinimapView: React.FC<{ activeDoc: OfficeDocument | null }> = ({
  activeDoc: doc,
}) => {
  if (!doc) return <p>No active document.</p>;
  if (doc.type === "word")
    return (
      <div className="document-context">
        <h4>DOCUMENT STRUCTURE</h4>
        {doc.data.paragraphs.map((p, i) => {
          const text =
            p.tableData?.flat().join(" ") ??
            (p.smartArt
              ? [
                  p.smartArt.title,
                  ...p.smartArt.items.map((item) => item.text),
                ].join(" ")
              : undefined) ??
            p.text ??
            p.runs?.map((r) => r.text).join("") ??
            "";
          return (
            <div key={p.id} title={text}>
              <small>
                {i + 1} · {p.type}
              </small>
              <div
                style={{
                  height: p.type.startsWith("heading") ? 14 : 7,
                  width: `${Math.min(100, Math.max(5, text.length / 5))}%`,
                  background: p.type.startsWith("heading")
                    ? "#4175c5"
                    : "var(--vscode-border)",
                  marginTop: 4,
                }}
              />
            </div>
          );
        })}
        <p>
          {doc.data.paragraphs.length} content blocks. Bar length reflects text
          length.
        </p>
      </div>
    );
  if (doc.type === "excel") {
    const sheet =
      doc.data.sheets.find((s) => s.id === doc.data.activeSheetId) ??
      doc.data.sheets[0];
    if (!sheet) return null;
    const cols = Math.min(10, sheet.colCount),
      rows = Math.min(14, sheet.rowCount);
    return (
      <div className="document-context">
        <h4>{sheet.name}</h4>
        <div
          style={{
            display: "grid",
            gridTemplateColumns: `repeat(${cols},1fr)`,
            gap: 3,
          }}
        >
          {Array.from({ length: cols * rows }, (_, i) => {
            const coord = `${columnName(i % cols)}${Math.floor(i / cols) + 1}`,
              cell = sheet.cells[coord],
              filled = cell?.value != null && cell.value !== "";
            return (
              <div
                key={coord}
                title={`${coord}: ${cell?.formula ?? cell?.value ?? "Empty"}`}
                style={{
                  height: 14,
                  background: cell?.formula
                    ? "#25874f"
                    : filled
                      ? "#397966"
                      : "var(--vscode-border)",
                }}
              />
            );
          })}
        </div>
        <p>
          First {rows} rows × {cols} columns. Green marks populated cells;
          bright green marks formulas.
        </p>
      </div>
    );
  }
  if (doc.type === "powerpoint")
    return (
      <div className="document-context">
        <h4>SLIDE OVERVIEW</h4>
        {doc.data.slides.map((s) => (
          <div
            key={s.id}
            style={{
              background: s.background,
              padding: 12,
              color: "white",
              border: "1px solid var(--vscode-border)",
              minHeight: 60,
            }}
          >
            <strong>
              {s.slideNumber}. {s.title}
            </strong>
            <div>
              {s.objects.filter((o) => o.visible !== false).length} visible
              objects · {s.notes ? "notes" : "no notes"}
            </div>
          </div>
        ))}
      </div>
    );
  return (
    <div className="document-context">
      <h4>SOURCE MAP</h4>
      {doc.data.content
        .split("\n")
        .slice(0, 100)
        .map((line, i) => (
          <div
            key={i}
            title={`Line ${i + 1}: ${line}`}
            style={{
              height: 3,
              marginLeft:
                Math.min(40, line.length - line.trimStart().length) * 2,
              width: `${Math.min(100, Math.max(1, line.trim().length))}%`,
              background: line.trim() ? "#5385a5" : "transparent",
            }}
          />
        ))}
      <p>First 100 lines.</p>
    </div>
  );
};
