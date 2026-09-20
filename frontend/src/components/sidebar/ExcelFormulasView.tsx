import React from "react";
import { ExcelFormulaAudit } from "../../types/office";
import { FunctionSquare, Calculator } from "lucide-react";

interface ExcelFormulasViewProps {
  formulas: ExcelFormulaAudit[];
  onSelectFormulaCell?: (coord: string) => void;
}

export const ExcelFormulasView: React.FC<ExcelFormulasViewProps> = ({
  formulas,
  onSelectFormulaCell,
}) => {
  if (!formulas || formulas.length === 0) {
    return (
      <div
        style={{
          padding: "8px 16px",
          color: "var(--vscode-tab-inactive-fg)",
          fontSize: 11,
        }}
      >
        No dynamic formulas detected in this sheet.
      </div>
    );
  }

  return (
    <div style={{ padding: "4px 0" }}>
      {formulas.map((f, idx) => (
        <div
          key={idx}
          className="vscode-tree-item"
          style={{ justifyContent: "space-between", padding: "5px 12px" }}
          onClick={() => onSelectFormulaCell?.(f.cellCoord)}
          title={`Cell ${f.cellCoord}: ${f.formula} -> ${f.evaluatedValue}`}
        >
          <div
            style={{
              display: "flex",
              alignItems: "center",
              gap: 6,
              minWidth: 0,
            }}
          >
            <span
              style={{
                fontFamily: "var(--vscode-mono-family)",
                fontWeight: "bold",
                color: "#107c41",
                fontSize: 11,
                minWidth: 28,
              }}
            >
              {f.cellCoord}
            </span>
            <span
              style={{
                fontFamily: "var(--vscode-mono-family)",
                fontSize: 11,
                color: "var(--vscode-editor-fg)",
                overflow: "hidden",
                textOverflow: "ellipsis",
              }}
            >
              {f.formula}
            </span>
          </div>
          <span
            style={{
              fontSize: 11,
              fontWeight: 600,
              color: "#34d399",
              marginLeft: 6,
            }}
          >
            {f.evaluatedValue}
          </span>
        </div>
      ))}
    </div>
  );
};
