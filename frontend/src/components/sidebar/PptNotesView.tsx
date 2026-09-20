import React from "react";
import { MessageSquareText } from "lucide-react";

interface PptNotesViewProps {
  notes: string;
  onChangeNotes?: (notes: string) => void;
}

export const PptNotesView: React.FC<PptNotesViewProps> = ({
  notes,
  onChangeNotes,
}) => {
  return (
    <div
      style={{
        padding: "8px 12px",
        display: "flex",
        flexDirection: "column",
        gap: 6,
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
        <MessageSquareText size={12} color="#c43e1c" />
        <span>SPEAKER NOTES:</span>
      </div>

      <textarea
        className="vscode-input"
        rows={6}
        style={{
          width: "100%",
          resize: "vertical",
          fontSize: 11,
          lineHeight: 1.5,
          fontFamily: "inherit",
        }}
        value={notes}
        onChange={(e) => onChangeNotes?.(e.target.value)}
        placeholder="Add speaker notes or presentation cues for this slide..."
      />
    </div>
  );
};
