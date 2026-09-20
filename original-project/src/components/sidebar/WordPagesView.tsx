import React from "react";
import { WordSection } from "../../types/office";
import { FileText, LayoutTemplate } from "lucide-react";

interface WordPagesViewProps {
  sections: WordSection[];
  onSelectSection: (sectionId: string) => void;
}

export const WordPagesView: React.FC<WordPagesViewProps> = ({
  sections,
  onSelectSection,
}) => {
  return (
    <div
      style={{
        padding: "4px 8px",
        display: "flex",
        flexDirection: "column",
        gap: 6,
      }}
    >
      {sections.map((sec, idx) => (
        <div
          key={sec.id}
          onClick={() => onSelectSection(sec.id)}
          style={{
            display: "flex",
            alignItems: "center",
            gap: 10,
            padding: "6px 10px",
            backgroundColor: "var(--vscode-sidebar-bg)",
            borderRadius: 4,
            cursor: "pointer",
            border: "1px solid #3c3c3c",
          }}
        >
          <div
            style={{
              width: 24,
              height: 32,
              backgroundColor: "var(--vscode-editor-bg)",
              border: "1px solid #555",
              borderRadius: 2,
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              fontSize: 10,
              color: "var(--vscode-tab-inactive-fg)",
            }}
          >
            {sec.pageNumber || idx + 1}
          </div>
          <div style={{ flex: 1, minWidth: 0 }}>
            <div
              style={{
                fontSize: 12,
                fontWeight: 600,
                color: "var(--vscode-editor-fg)",
                whiteSpace: "nowrap",
                overflow: "hidden",
                textOverflow: "ellipsis",
              }}
            >
              {sec.title}
            </div>
            <div
              style={{
                fontSize: 10,
                color: "var(--vscode-tab-inactive-fg)",
                display: "flex",
                gap: 6,
              }}
            >
              <span>Page {sec.pageNumber || idx + 1}</span>
              <span>•</span>
              <span style={{ textTransform: "capitalize" }}>
                {sec.orientation}
              </span>
            </div>
          </div>
        </div>
      ))}
    </div>
  );
};
