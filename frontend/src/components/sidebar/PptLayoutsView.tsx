import React from "react";
import { Layout, Palette } from "lucide-react";

interface PptLayoutsViewProps {
  currentLayout: string;
  themeName: string;
  onChangeLayout?: (
    layout: "title" | "content" | "two-column" | "dashboard" | "blank",
  ) => void;
}

export const PptLayoutsView: React.FC<PptLayoutsViewProps> = ({
  currentLayout,
  themeName,
  onChangeLayout,
}) => {
  const layouts = [
    { id: "title", label: "Title Slide" },
    { id: "content", label: "Title & Content" },
    { id: "two-column", label: "Two Column Comparison" },
    { id: "dashboard", label: "Metrics Dashboard" },
    { id: "blank", label: "Blank Slide" },
  ];

  return (
    <div
      style={{
        padding: "8px 12px",
        display: "flex",
        flexDirection: "column",
        gap: 8,
      }}
    >
      <div
        style={{
          display: "flex",
          alignItems: "center",
          gap: 6,
          color: "var(--vscode-tab-inactive-fg)",
          fontSize: 10,
        }}
      >
        <Palette size={12} color="#c43e1c" />
        <span>
          THEME: <strong>{themeName}</strong>
        </span>
      </div>

      <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
        <span style={{ fontSize: 10, color: "var(--vscode-tab-inactive-fg)" }}>
          SLIDE LAYOUT TEMPLATE:
        </span>
        {layouts.map((l) => (
          <div
            key={l.id}
            onClick={() => onChangeLayout?.(l.id as any)}
            style={{
              padding: "6px 10px",
              fontSize: 11,
              borderRadius: 4,
              backgroundColor:
                currentLayout === l.id ? "#37373d" : "var(--vscode-sidebar-bg)",
              border:
                currentLayout === l.id
                  ? "1px solid #c43e1c"
                  : "1px solid #3c3c3c",
              color:
                currentLayout === l.id
                  ? "var(--vscode-tab-active-fg)"
                  : "var(--vscode-editor-fg)",
              cursor: "pointer",
              display: "flex",
              alignItems: "center",
              justifyContent: "space-between",
            }}
          >
            <span>{l.label}</span>
            {currentLayout === l.id && (
              <span
                style={{ fontSize: 9, color: "#f87171", fontWeight: "bold" }}
              >
                Active
              </span>
            )}
          </div>
        ))}
      </div>
    </div>
  );
};
