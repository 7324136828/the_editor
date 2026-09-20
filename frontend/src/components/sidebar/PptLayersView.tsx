import React from "react";
import { PptShapeObject } from "../../types/office";
import { Layers, Type, Square, Eye, EyeOff, BarChart2 } from "lucide-react";

interface PptLayersViewProps {
  objects: PptShapeObject[];
  selectedObjectId: string | null;
  onSelectObject: (id: string) => void;
  onToggleVisibility?: (id: string) => void;
}

export const PptLayersView: React.FC<PptLayersViewProps> = ({
  objects,
  selectedObjectId,
  onSelectObject,
  onToggleVisibility,
}) => {
  if (!objects || objects.length === 0) {
    return (
      <div
        style={{
          padding: "8px 16px",
          color: "var(--vscode-tab-inactive-fg)",
          fontSize: 11,
        }}
      >
        No shapes or elements on this slide.
      </div>
    );
  }

  // Sort objects by z-index descending (top-most first)
  const sorted = [...objects].sort((a, b) => (b.zIndex || 0) - (a.zIndex || 0));

  return (
    <div style={{ padding: "4px 0" }}>
      {sorted.map((obj) => {
        const isSelected = obj.id === selectedObjectId;
        const isHidden = obj.visible === false;

        return (
          <div
            key={obj.id}
            className={`vscode-tree-item ${isSelected ? "active" : ""}`}
            onClick={() => onSelectObject(obj.id)}
            style={{
              justifyContent: "space-between",
              opacity: isHidden ? 0.4 : 1,
            }}
          >
            <div
              style={{
                display: "flex",
                alignItems: "center",
                gap: 6,
                minWidth: 0,
              }}
            >
              {obj.kind === "text" && <Type size={13} color="#60a5fa" />}
              {obj.kind === "shape" && <Square size={13} color="#f97316" />}
              {obj.kind === "metric" && <BarChart2 size={13} color="#34d399" />}
              <span style={{ overflow: "hidden", textOverflow: "ellipsis" }}>
                {obj.text || obj.title || obj.metricLabel || `Object ${obj.id}`}
              </span>
            </div>

            <div
              style={{ display: "flex", alignItems: "center", gap: 4 }}
              onClick={(e) => e.stopPropagation()}
            >
              {onToggleVisibility && (
                <button
                  className="vscode-icon-btn"
                  style={{ width: 18, height: 18 }}
                  onClick={() => onToggleVisibility(obj.id)}
                  title={isHidden ? "Show element" : "Hide element"}
                >
                  {isHidden ? (
                    <EyeOff size={11} color="#858585" />
                  ) : (
                    <Eye size={11} />
                  )}
                </button>
              )}
            </div>
          </div>
        );
      })}
    </div>
  );
};
