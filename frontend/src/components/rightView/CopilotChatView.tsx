import React, { useState } from "react";
import { OfficeDocument } from "../../types/office";
import { DocumentChat } from "../chat/DocumentChat";
import { collectDocumentContext, getDocumentDiagnostics } from "../../plugins";

export const CopilotChatView: React.FC<{
  activeDoc: OfficeDocument | null;
}> = ({ activeDoc }) => {
  const [mode, setMode] = useState<"chat" | "diagnostics">("chat");

  if (!activeDoc) {
    return (
      <div style={{ padding: 16 }}>
        <DocumentChat activeDoc={null} />
      </div>
    );
  }

  const context = collectDocumentContext(activeDoc);
  const diagnostics = getDocumentDiagnostics(activeDoc);

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100%" }}>
      {/* Switch between Chat & Local Diagnostics */}
      <div
        style={{
          display: "flex",
          borderBottom: "1px solid var(--vscode-panel-border, #333)",
          backgroundColor: "var(--vscode-sideBar-bg, #252526)",
        }}
      >
        <button
          className="vscode-btn-secondary"
          style={{
            flex: 1,
            borderRadius: 0,
            border: "none",
            borderBottom: mode === "chat" ? "2px solid #38bdf8" : "none",
            backgroundColor: mode === "chat" ? "var(--vscode-editor-bg, #1e1e1e)" : "transparent",
            color: mode === "chat" ? "#38bdf8" : "var(--vscode-tab-inactive-fg, #888)",
            padding: "8px 12px",
            fontSize: 12,
            fontWeight: 500,
            cursor: "pointer",
          }}
          onClick={() => setMode("chat")}
        >
          AI Chat Assistant
        </button>
        <button
          className="vscode-btn-secondary"
          style={{
            flex: 1,
            borderRadius: 0,
            border: "none",
            borderBottom: mode === "diagnostics" ? "2px solid #38bdf8" : "none",
            backgroundColor: mode === "diagnostics" ? "var(--vscode-editor-bg, #1e1e1e)" : "transparent",
            color: mode === "diagnostics" ? "#38bdf8" : "var(--vscode-tab-inactive-fg, #888)",
            padding: "8px 12px",
            fontSize: 12,
            fontWeight: 500,
            cursor: "pointer",
          }}
          onClick={() => setMode("diagnostics")}
        >
          Document Health
        </button>
      </div>

      <div style={{ flex: 1, overflow: "hidden" }}>
        {mode === "chat" ? (
          <DocumentChat activeDoc={activeDoc} />
        ) : (
          <div style={{ padding: 14, overflowY: "auto", height: "100%" }}>
            <h4 style={{ margin: "0 0 8px 0", fontSize: 13 }}>DIAGNOSTICS & SUMMARY</h4>
            <p style={{ fontSize: 12, color: "var(--vscode-editor-fg, #ccc)" }}>{context.summary}</p>
            <div style={{ marginTop: 12 }}>
              <div style={{ fontSize: 11, fontWeight: 600, color: "#888", marginBottom: 6 }}>OUTLINE</div>
              <pre
                style={{
                  fontSize: 11,
                  backgroundColor: "var(--vscode-editor-bg, #1e1e1e)",
                  padding: 8,
                  borderRadius: 4,
                  overflowX: "auto",
                }}
              >
                {context.outline.join("\n") || "No outline items."}
              </pre>
            </div>
            <div style={{ marginTop: 12 }}>
              <div style={{ fontSize: 11, fontWeight: 600, color: "#888", marginBottom: 6 }}>ISSUES & CHECKS</div>
              <pre
                style={{
                  fontSize: 11,
                  backgroundColor: "var(--vscode-editor-bg, #1e1e1e)",
                  padding: 8,
                  borderRadius: 4,
                  overflowX: "auto",
                }}
              >
                {diagnostics.map((d) => `${d.severity.toUpperCase()} · ${d.location}\n${d.message}`).join("\n\n") ||
                  "No issues found."}
              </pre>
            </div>
          </div>
        )}
      </div>
    </div>
  );
};
