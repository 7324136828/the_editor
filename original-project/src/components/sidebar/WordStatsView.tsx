import React from "react";
import { WordDocumentModel } from "../../types/office";
import { Clock, FileText, AlignLeft, Award } from "lucide-react";

interface WordStatsViewProps {
  stats: WordDocumentModel["stats"];
}

export const WordStatsView: React.FC<WordStatsViewProps> = ({ stats }) => {
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
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
        <div
          style={{
            backgroundColor: "var(--vscode-sidebar-bg)",
            padding: 8,
            borderRadius: 4,
            border: "1px solid #3c3c3c",
          }}
        >
          <div style={{ color: "var(--vscode-tab-inactive-fg)", fontSize: 10 }}>
            WORDS
          </div>
          <div style={{ fontSize: 15, fontWeight: "bold", color: "#60a5fa" }}>
            {stats.words.toLocaleString()}
          </div>
        </div>
        <div
          style={{
            backgroundColor: "var(--vscode-sidebar-bg)",
            padding: 8,
            borderRadius: 4,
            border: "1px solid #3c3c3c",
          }}
        >
          <div style={{ color: "var(--vscode-tab-inactive-fg)", fontSize: 10 }}>
            ESTIMATED PAGES
          </div>
          <div style={{ fontSize: 15, fontWeight: "bold", color: "#60a5fa" }}>
            {stats.pages}
          </div>
        </div>
        <div
          style={{
            backgroundColor: "var(--vscode-sidebar-bg)",
            padding: 8,
            borderRadius: 4,
            border: "1px solid #3c3c3c",
          }}
        >
          <div style={{ color: "var(--vscode-tab-inactive-fg)", fontSize: 10 }}>
            CHARACTERS
          </div>
          <div
            style={{
              fontSize: 13,
              fontWeight: 600,
              color: "var(--vscode-editor-fg)",
            }}
          >
            {stats.characters.toLocaleString()}
          </div>
        </div>
        <div
          style={{
            backgroundColor: "var(--vscode-sidebar-bg)",
            padding: 8,
            borderRadius: 4,
            border: "1px solid #3c3c3c",
          }}
        >
          <div style={{ color: "var(--vscode-tab-inactive-fg)", fontSize: 10 }}>
            PARAGRAPHS
          </div>
          <div
            style={{
              fontSize: 13,
              fontWeight: 600,
              color: "var(--vscode-editor-fg)",
            }}
          >
            {stats.paragraphs}
          </div>
        </div>
      </div>

      <div
        style={{
          display: "flex",
          alignItems: "center",
          gap: 6,
          color: "#a0aec0",
          marginTop: 4,
        }}
      >
        <Clock size={12} color="#fbbf24" />
        <span>
          Reading Time: <strong>~{stats.readingTimeMinutes} mins</strong>
        </span>
      </div>

      <div
        style={{
          display: "flex",
          alignItems: "center",
          gap: 6,
          color: "#a0aec0",
        }}
      >
        <Award size={12} color="#34d399" />
        <span>
          Readability: <strong>{stats.readabilityScore}</strong>
        </span>
      </div>
    </div>
  );
};
