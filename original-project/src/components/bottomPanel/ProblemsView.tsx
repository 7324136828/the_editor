import React from "react";
import { ProblemItem } from "../../types/vscode";
import { AlertCircle, AlertTriangle, Info } from "lucide-react";

interface ProblemsViewProps {
  problems: ProblemItem[];
  onSelectProblem?: (p: ProblemItem) => void;
}

export const ProblemsView: React.FC<ProblemsViewProps> = ({
  problems,
  onSelectProblem,
}) => {
  if (!problems || problems.length === 0) {
    return (
      <div
        style={{
          padding: "16px",
          color: "var(--vscode-tab-inactive-fg)",
          fontSize: 12,
          textAlign: "center",
        }}
      >
        No problems or lint warnings detected in workspace documents.
      </div>
    );
  }

  return (
    <div
      style={{ display: "flex", flexDirection: "column", gap: 4, fontSize: 12 }}
    >
      {problems.map((prob) => (
        <div
          key={prob.id}
          className="vscode-tree-item"
          style={{
            justifyContent: "space-between",
            padding: "5px 8px",
            borderRadius: 3,
          }}
          onClick={() => onSelectProblem?.(prob)}
        >
          <div
            style={{
              display: "flex",
              alignItems: "center",
              gap: 8,
              minWidth: 0,
            }}
          >
            {prob.severity === "error" && (
              <AlertCircle size={14} color="#f87171" />
            )}
            {prob.severity === "warning" && (
              <AlertTriangle size={14} color="#fbbf24" />
            )}
            {prob.severity === "info" && <Info size={14} color="#60a5fa" />}
            <span style={{ color: "var(--vscode-tab-active-fg)" }}>
              {prob.message}
            </span>
            <span
              style={{ color: "var(--vscode-tab-inactive-fg)", fontSize: 11 }}
            >
              [{prob.source}]
            </span>
          </div>

          <div
            style={{
              display: "flex",
              alignItems: "center",
              gap: 8,
              fontSize: 11,
              color: "var(--vscode-tab-inactive-fg)",
            }}
          >
            <span>{prob.fileName}</span>
            <span style={{ color: "#60a5fa" }}>{prob.location}</span>
          </div>
        </div>
      ))}
    </div>
  );
};
