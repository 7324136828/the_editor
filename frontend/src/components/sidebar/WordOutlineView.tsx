import React from "react";
import { WordHeading } from "../../types/office";
import { Heading1, Heading2, Heading3, Bookmark } from "lucide-react";

interface WordOutlineViewProps {
  headings: WordHeading[];
  onSelectHeading: (headingId: string) => void;
}

export const WordOutlineView: React.FC<WordOutlineViewProps> = ({
  headings,
  onSelectHeading,
}) => {
  if (headings.length === 0) {
    return (
      <div
        style={{
          padding: "8px 16px",
          color: "var(--vscode-tab-inactive-fg)",
          fontSize: 11,
        }}
      >
        No headings detected in this document.
      </div>
    );
  }

  return (
    <div style={{ padding: "2px 0" }}>
      {headings.map((h) => {
        const indent = (h.level - 1) * 14;
        return (
          <div
            key={h.id}
            className="vscode-tree-item"
            style={{ paddingLeft: 12 + indent }}
            onClick={() => onSelectHeading(h.id)}
            title={`${h.text} (Page ${h.pageIndex})`}
          >
            {h.level === 1 && <Heading1 size={13} color="#4175c5" />}
            {h.level === 2 && <Heading2 size={13} color="#60a5fa" />}
            {h.level >= 3 && <Bookmark size={12} color="#93c5fd" />}
            <span style={{ overflow: "hidden", textOverflow: "ellipsis" }}>
              {h.text}
            </span>
            <span
              style={{
                marginLeft: "auto",
                fontSize: 10,
                color: "#6e7681",
                paddingRight: 8,
              }}
            >
              p.{h.pageIndex}
            </span>
          </div>
        );
      })}
    </div>
  );
};
