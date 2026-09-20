import React, { useState } from "react";
import { OfficeDocument } from "../../types/office";
import { collectDocumentContext, getDocumentDiagnostics } from "../../plugins";
export const CopilotChatView: React.FC<{
  activeDoc: OfficeDocument | null;
}> = ({ activeDoc }) => {
  const [view, setView] = useState<"outline" | "text" | "checks">("outline");
  if (!activeDoc) return <p>Open a file to inspect its context.</p>;
  const context = collectDocumentContext(activeDoc),
    diagnostics = getDocumentDiagnostics(activeDoc);
  return (
    <div className="document-context">
      <h4>LIVE DOCUMENT CONTEXT</h4>
      <p>{context.summary}</p>
      <p style={{ color: "var(--vscode-tab-inactive-fg)" }}>
        Local, deterministic analysis of the current editor. Use the MCP
        extension to share saved document context with a compatible assistant.
      </p>
      <div style={{ display: "flex", gap: 6 }}>
        {(["outline", "text", "checks"] as const).map((v) => (
          <button
            key={v}
            className="vscode-btn-secondary"
            aria-pressed={v === view}
            onClick={() => setView(v)}
          >
            {v}
          </button>
        ))}
      </div>
      {view === "outline" ? (
        <pre>{context.outline.join("\n") || "No headings or symbols yet."}</pre>
      ) : view === "text" ? (
        <>
          <pre>{context.text || "Empty document"}</pre>
          {context.truncated && <p>Preview limited to 12,000 characters.</p>}
        </>
      ) : (
        <pre>
          {diagnostics
            .map(
              (d) =>
                `${d.severity.toUpperCase()} · ${d.location}\n${d.message}`,
            )
            .join("\n\n") || "No issues found by the available checks."}
        </pre>
      )}
    </div>
  );
};
