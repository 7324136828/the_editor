import React from "react";
import { PptSlide } from "../../types/office";
import { Presentation, Copy, Trash2 } from "lucide-react";

interface PptSlideNavigatorViewProps {
  slides: PptSlide[];
  activeSlideId: string;
  onSelectSlide: (slideId: string) => void;
  onDuplicateSlide?: (slideId: string) => void;
  onDeleteSlide?: (slideId: string) => void;
}

export const PptSlideNavigatorView: React.FC<PptSlideNavigatorViewProps> = ({
  slides,
  activeSlideId,
  onSelectSlide,
  onDuplicateSlide,
  onDeleteSlide,
}) => {
  return (
    <div
      style={{
        padding: "6px 8px",
        display: "flex",
        flexDirection: "column",
        gap: 6,
      }}
    >
      {slides.map((slide) => {
        const isActive = slide.id === activeSlideId;
        return (
          <div
            key={slide.id}
            onClick={() => onSelectSlide(slide.id)}
            style={{
              display: "flex",
              alignItems: "center",
              gap: 8,
              padding: "6px 8px",
              backgroundColor: isActive
                ? "#37373d"
                : "var(--vscode-sidebar-bg)",
              border: isActive ? "1px solid #c43e1c" : "1px solid #333",
              borderRadius: 4,
              cursor: "pointer",
              transition: "background-color 0.15s",
            }}
          >
            {/* Slide Index Badge */}
            <span
              style={{
                fontSize: 10,
                color: isActive ? "#f87171" : "var(--vscode-tab-inactive-fg)",
                fontWeight: "bold",
                minWidth: 16,
                textAlign: "center",
              }}
            >
              {slide.slideNumber}
            </span>

            {/* Miniature Slide Preview */}
            <div
              style={{
                width: 54,
                height: 32,
                borderRadius: 2,
                background: slide.background || "#181b22",
                border: "1px solid rgba(255,255,255,0.15)",
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                overflow: "hidden",
                flexShrink: 0,
              }}
            >
              <Presentation
                size={14}
                color={isActive ? "#f87171" : "#6b7280"}
              />
            </div>

            {/* Slide Title & Objects Count */}
            <div style={{ flex: 1, minWidth: 0 }}>
              <div
                style={{
                  fontSize: 11,
                  fontWeight: isActive ? 600 : "normal",
                  color: isActive
                    ? "var(--vscode-tab-active-fg)"
                    : "var(--vscode-editor-fg)",
                  whiteSpace: "nowrap",
                  overflow: "hidden",
                  textOverflow: "ellipsis",
                }}
              >
                {slide.title}
              </div>
              <div
                style={{ fontSize: 9, color: "var(--vscode-tab-inactive-fg)" }}
              >
                {slide.objects.length} elements • {slide.layout}
              </div>
            </div>

            {/* Slide action buttons */}
            {isActive && (
              <div
                style={{ display: "flex", gap: 2 }}
                onClick={(e) => e.stopPropagation()}
              >
                {onDuplicateSlide && (
                  <button
                    className="vscode-icon-btn"
                    style={{ width: 20, height: 20 }}
                    title="Duplicate slide"
                    onClick={() => onDuplicateSlide(slide.id)}
                  >
                    <Copy size={11} />
                  </button>
                )}
                {onDeleteSlide && slides.length > 1 && (
                  <button
                    className="vscode-icon-btn"
                    style={{ width: 20, height: 20 }}
                    title="Delete slide"
                    onClick={() => onDeleteSlide(slide.id)}
                  >
                    <Trash2 size={11} color="#f87171" />
                  </button>
                )}
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
};
