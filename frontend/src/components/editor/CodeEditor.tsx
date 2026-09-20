import React, { useRef } from "react";
import { CodeDocumentModel } from "../../types/office";

interface CodeEditorProps {
  document: CodeDocumentModel;
  onChangeDocument?: (doc: CodeDocumentModel) => void;
}

export const CodeEditor: React.FC<CodeEditorProps> = ({
  document,
  onChangeDocument,
}) => {
  const textarea = useRef<HTMLTextAreaElement>(null);
  const lines = document.content.split("\n");
  const update = (content: string) => {
    const symbols: CodeDocumentModel["symbols"] = [];
    content.split("\n").forEach((line, index) => {
      const match =
        /^\s*(?:export\s+)?(?:default\s+)?(?:async\s+)?(function|def|class|interface|const|let|var)\s+([\w$]+)/.exec(
          line,
        );
      if (match)
        symbols.push({
          name: match[2],
          kind:
            match[1] === "def" || match[1] === "function"
              ? "function"
              : match[1] === "interface"
                ? "interface"
                : match[1] === "class"
                  ? "class"
                  : "variable",
          line: index + 1,
        });
    });
    onChangeDocument?.({ ...document, content, symbols });
  };
  return (
    <div
      style={{
        height: "100%",
        background: "#1e1e1e",
        overflow: "auto",
        fontFamily: "var(--vscode-mono-family)",
        fontSize: 13,
      }}
    >
      <div
        style={{
          minHeight: "100%",
          display: "flex",
          padding: "12px 0",
          boxSizing: "border-box",
        }}
      >
        <div
          aria-hidden="true"
          style={{
            width: 48,
            minWidth: 48,
            color: "#858585",
            textAlign: "right",
            paddingRight: 16,
            userSelect: "none",
            lineHeight: "21px",
          }}
        >
          {lines.map((_, index) => (
            <div key={index}>{index + 1}</div>
          ))}
        </div>
        <textarea
          ref={textarea}
          className="code-editor-input"
          aria-label={`${document.language} editor`}
          spellCheck={false}
          autoCapitalize="off"
          autoCorrect="off"
          wrap="off"
          value={document.content}
          onChange={(e) => update(e.target.value)}
          style={{
            display: "block",
            flex: 1,
            minWidth: 400,
            minHeight: `${Math.max(lines.length + 1, 20) * 21}px`,
            background: "transparent",
            border: 0,
            outline: 0,
            color: "#9cdcfe",
            fontFamily: "inherit",
            fontSize: "inherit",
            lineHeight: "21px",
            resize: "none",
            padding: "0 16px 0 0",
            overflow: "hidden",
            tabSize: 2,
          }}
          onKeyDown={(event) => {
            if (event.key !== "Tab") return;
            event.preventDefault();
            const field = event.currentTarget;
            const start = field.selectionStart;
            const end = field.selectionEnd;
            const value = field.value;
            if (event.shiftKey) {
              const lineStart = value.lastIndexOf("\n", start - 1) + 1;
              const indentation =
                /^(?: {1,2}|\t)/.exec(value.slice(lineStart))?.[0].length || 0;
              update(
                value.slice(0, lineStart) +
                  value.slice(lineStart + indentation),
              );
              requestAnimationFrame(() =>
                field.setSelectionRange(
                  Math.max(lineStart, start - indentation),
                  Math.max(lineStart, end - indentation),
                ),
              );
            } else {
              update(value.slice(0, start) + "  " + value.slice(end));
              requestAnimationFrame(() =>
                field.setSelectionRange(start + 2, start + 2),
              );
            }
          }}
        />
      </div>
    </div>
  );
};
