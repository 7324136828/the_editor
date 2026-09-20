import React, { useState } from "react";
import { ExcelColumn } from "../../types/office";
import {
  ArrowUpDown,
  ArrowDownAZ,
  ArrowUpAZ,
  Filter,
  RefreshCw,
} from "lucide-react";

interface ExcelFilterSortViewProps {
  columns: ExcelColumn[];
  onSort?: (colKey: string, direction: "asc" | "desc") => void;
}

export const ExcelFilterSortView: React.FC<ExcelFilterSortViewProps> = ({
  columns,
  onSort,
}) => {
  const [selectedCol, setSelectedCol] = useState(columns[0]?.key || "A");

  return (
    <div
      style={{
        padding: "8px 12px",
        display: "flex",
        flexDirection: "column",
        gap: 8,
        fontSize: 11,
      }}
    >
      <div>
        <label
          style={{
            display: "block",
            color: "var(--vscode-tab-inactive-fg)",
            marginBottom: 4,
            fontSize: 10,
          }}
        >
          SELECT COLUMN:
        </label>
        <select
          className="vscode-input"
          value={selectedCol}
          onChange={(e) => setSelectedCol(e.target.value)}
          style={{ width: "100%", height: 26 }}
        >
          {columns.map((c) => (
            <option key={c.key} value={c.key}>
              Column {c.key}: {c.label}
            </option>
          ))}
        </select>
      </div>

      <div style={{ display: "flex", gap: 6 }}>
        <button
          className="vscode-btn-secondary"
          style={{ flex: 1, justifyContent: "center" }}
          onClick={() => onSort?.(selectedCol, "asc")}
          title="Sort Ascending (A-Z / 1-9)"
        >
          <ArrowDownAZ size={13} />
          <span>Sort Asc</span>
        </button>
        <button
          className="vscode-btn-secondary"
          style={{ flex: 1, justifyContent: "center" }}
          onClick={() => onSort?.(selectedCol, "desc")}
          title="Sort Descending (Z-A / 9-1)"
        >
          <ArrowUpAZ size={13} />
          <span>Sort Desc</span>
        </button>
      </div>

      <div
        style={{
          display: "flex",
          alignItems: "center",
          gap: 6,
          color: "var(--vscode-tab-inactive-fg)",
          marginTop: 2,
          borderTop: "1px solid #333",
          paddingTop: 6,
        }}
      >
        <Filter size={12} color="#107c41" />
        <span>Sorting preserves the header row</span>
      </div>
    </div>
  );
};
