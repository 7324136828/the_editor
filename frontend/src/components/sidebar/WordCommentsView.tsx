import React from "react";
import { WordComment } from "../../types/office";
import { MessageSquare, CheckCircle, Circle } from "lucide-react";

interface WordCommentsViewProps {
  comments: WordComment[];
  onToggleResolve: (commentId: string) => void;
}

export const WordCommentsView: React.FC<WordCommentsViewProps> = ({
  comments,
  onToggleResolve,
}) => {
  if (comments.length === 0) {
    return (
      <div
        style={{
          padding: "8px 16px",
          color: "var(--vscode-tab-inactive-fg)",
          fontSize: 11,
        }}
      >
        No review comments in this document.
      </div>
    );
  }

  return (
    <div
      style={{
        padding: "4px 8px",
        display: "flex",
        flexDirection: "column",
        gap: 6,
      }}
    >
      {comments.map((c) => (
        <div
          key={c.id}
          style={{
            backgroundColor: c.resolved
              ? "var(--vscode-editor-bg)"
              : "var(--vscode-sidebar-bg)",
            border: `1px solid ${c.resolved ? "#333333" : "#444444"}`,
            borderRadius: 4,
            padding: 8,
            opacity: c.resolved ? 0.65 : 1,
          }}
        >
          <div
            style={{
              display: "flex",
              alignItems: "center",
              justifyContent: "space-between",
              marginBottom: 4,
            }}
          >
            <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
              <div
                style={{
                  width: 16,
                  height: 16,
                  borderRadius: "50%",
                  backgroundColor: c.avatarColor,
                  color: "#fff",
                  fontSize: 9,
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  fontWeight: "bold",
                }}
              >
                {c.author.charAt(0)}
              </div>
              <span
                style={{
                  fontSize: 11,
                  fontWeight: 600,
                  color: "var(--vscode-editor-fg)",
                }}
              >
                {c.author}
              </span>
            </div>
            <button
              className="vscode-icon-btn"
              style={{ width: 18, height: 18 }}
              title={c.resolved ? "Mark unresolved" : "Mark resolved"}
              onClick={() => onToggleResolve(c.id)}
            >
              {c.resolved ? (
                <CheckCircle size={13} color="#48bb78" />
              ) : (
                <Circle size={13} color="#858585" />
              )}
            </button>
          </div>

          <div
            style={{
              fontSize: 10,
              color: "#3b82f6",
              fontStyle: "italic",
              marginBottom: 4,
              borderLeft: "2px solid #3b82f6",
              paddingLeft: 6,
            }}
          >
            "{c.selectedText}"
          </div>

          <div
            style={{
              fontSize: 11,
              color: "var(--vscode-sidebar-fg)",
              lineHeight: 1.4,
            }}
          >
            {c.comment}
          </div>

          <div style={{ fontSize: 9, color: "#718096", marginTop: 4 }}>
            {c.timestamp}
          </div>
        </div>
      ))}
    </div>
  );
};
