import React from "react";
import { ExcelColumn } from "../../types/office";
import { Hash, DollarSign, Type, Calendar } from "lucide-react";

interface ExcelDataFieldsViewProps {
  columns: ExcelColumn[];
}

export const ExcelDataFieldsView: React.FC<ExcelDataFieldsViewProps> = ({
  columns,
}) => {
  return (
    <div style={{ padding: "4px 0" }}>
      {columns.map((col) => {
        return (
          <div
            key={col.key}
            className="vscode-tree-item"
            style={{ justifyContent: "space-between" }}
          >
            <div
              style={{
                display: "flex",
                alignItems: "center",
                gap: 6,
                minWidth: 0,
              }}
            >
              {col.type === "currency" && (
                <DollarSign size={13} color="#107c41" />
              )}
              {col.type === "number" && <Hash size={13} color="#3b82f6" />}
              {col.type === "date" && <Calendar size={13} color="#f59e0b" />}
              {col.type === "string" && <Type size={13} color="#9ca3af" />}
              <span style={{ overflow: "hidden", textOverflow: "ellipsis" }}>
                {col.label}
              </span>
            </div>
            <span
              style={{
                fontSize: 10,
                color: "#6b7280",
                backgroundColor: "rgba(255, 255, 255, 0.06)",
                padding: "1px 5px",
                borderRadius: 3,
                textTransform: "uppercase",
              }}
            >
              Col {col.key}
            </span>
          </div>
        );
      })}
    </div>
  );
};
