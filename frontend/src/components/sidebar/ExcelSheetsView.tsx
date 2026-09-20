import React from "react";
import { ExcelWorksheet } from "../../types/office";
import { Table, Plus, Check } from "lucide-react";

interface ExcelSheetsViewProps {
  sheets: ExcelWorksheet[];
  activeSheetId: string;
  onSelectSheet: (sheetId: string) => void;
  onAddSheet?: () => void;
}

export const ExcelSheetsView: React.FC<ExcelSheetsViewProps> = ({
  sheets,
  activeSheetId,
  onSelectSheet,
  onAddSheet,
}) => {
  return (
    <div style={{ padding: "4px 0" }}>
      {sheets.map((sheet) => {
        const isActive = sheet.id === activeSheetId;
        return (
          <div
            key={sheet.id}
            className={`vscode-tree-item ${isActive ? "active" : ""}`}
            onClick={() => onSelectSheet(sheet.id)}
            style={{
              borderLeft: isActive
                ? "3px solid #107c41"
                : "3px solid transparent",
              justifyContent: "space-between",
            }}
          >
            <div
              style={{
                display: "flex",
                alignItems: "center",
                gap: 6,
                minWidth: 0,
              }}
            >
              <Table
                size={13}
                color={isActive ? "#34d399" : "var(--vscode-tab-inactive-fg)"}
              />
              <span
                style={{
                  fontWeight: isActive ? 600 : "normal",
                  overflow: "hidden",
                  textOverflow: "ellipsis",
                }}
              >
                {sheet.name}
              </span>
            </div>
            <span
              style={{ fontSize: 10, color: "var(--vscode-tab-inactive-fg)" }}
            >
              {sheet.columns.length} cols × {sheet.rowCount} rows
            </span>
          </div>
        );
      })}

      {onAddSheet && (
        <div
          className="vscode-tree-item"
          style={{ color: "#34d399", marginTop: 4 }}
          onClick={onAddSheet}
        >
          <Plus size={13} />
          <span>Add New Worksheet...</span>
        </div>
      )}
    </div>
  );
};
