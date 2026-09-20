import React, { useEffect, useState } from "react";
import { PptDocumentModel } from "../../types/office";
import { ChevronLeft, ChevronRight, X, MessageSquare } from "lucide-react";

interface PresentationModalProps {
  document: PptDocumentModel;
  onClose: () => void;
}

export const PresentationModal: React.FC<PresentationModalProps> = ({
  document,
  onClose,
}) => {
  const [index, setIndex] = useState(
    Math.max(
      0,
      document.slides.findIndex((s) => s.id === document.activeSlideId),
    ),
  );
  const [showNotes, setShowNotes] = useState(false);
  const [viewport, setViewport] = useState({
    width: window.innerWidth,
    height: window.innerHeight,
  });
  const slide = document.slides[index] || document.slides[0];
  useEffect(() => {
    const resize = () =>
      setViewport({ width: window.innerWidth, height: window.innerHeight });
    const keydown = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        event.preventDefault();
        event.stopImmediatePropagation();
        onClose();
      } else if (["ArrowRight", " ", "Enter", "PageDown"].includes(event.key)) {
        event.preventDefault();
        setIndex((i) => Math.min(document.slides.length - 1, i + 1));
      } else if (["ArrowLeft", "PageUp"].includes(event.key)) {
        event.preventDefault();
        setIndex((i) => Math.max(0, i - 1));
      } else if (event.key === "Home") setIndex(0);
      else if (event.key === "End") setIndex(document.slides.length - 1);
    };
    window.addEventListener("resize", resize);
    window.addEventListener("keydown", keydown, true);
    return () => {
      window.removeEventListener("resize", resize);
      window.removeEventListener("keydown", keydown, true);
    };
  }, [document.slides.length, onClose]);
  const scale = Math.max(
    0.15,
    Math.min(
      (viewport.width - 110) / 920,
      (viewport.height - (showNotes ? 210 : 110)) / 518,
    ),
  );
  if (!slide) return null;
  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-label="Slideshow"
      style={{
        position: "fixed",
        inset: 0,
        background: "#0a0d14",
        zIndex: 10000,
        display: "flex",
        flexDirection: "column",
        color: "#fff",
      }}
    >
      <div
        style={{
          minHeight: 44,
          padding: "0 16px",
          display: "flex",
          alignItems: "center",
          gap: 12,
          background: "#0008",
          fontSize: 12,
        }}
      >
        <strong style={{ color: "#e57351" }}>SLIDESHOW</strong>
        <span>{document.title}</span>
        <span style={{ color: "#4fd1c5" }}>
          {index + 1} / {document.slides.length}
        </span>
        <span style={{ flex: 1 }} />
        <button
          className="vscode-btn-secondary"
          onClick={() => setShowNotes(!showNotes)}
        >
          <MessageSquare size={13} /> {showNotes ? "Hide notes" : "Notes"}
        </button>
        <button
          autoFocus
          className="vscode-icon-btn"
          title="Exit slideshow (Esc)"
          aria-label="Exit slideshow"
          onClick={onClose}
        >
          <X size={17} />
        </button>
      </div>
      <div
        style={{
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          flex: 1,
          minHeight: 0,
          gap: 12,
        }}
      >
        <button
          className="vscode-icon-btn"
          aria-label="Previous slide"
          disabled={index === 0}
          onClick={() => setIndex((i) => Math.max(0, i - 1))}
        >
          <ChevronLeft size={24} />
        </button>
        <div
          style={{
            width: 920 * scale,
            height: 518 * scale,
            position: "relative",
            overflow: "hidden",
            background: slide.background,
            boxShadow: "0 20px 60px #0008",
            flexShrink: 0,
          }}
        >
          {slide.objects
            .filter((obj) => obj.visible !== false)
            .map((obj) => (
              <div
                key={obj.id}
                style={{
                  position: "absolute",
                  left: obj.x * scale,
                  top: obj.y * scale,
                  width: obj.width * scale,
                  height: obj.height * scale,
                  background: obj.fill,
                  border: obj.stroke
                    ? `${(obj.strokeWidth || 1) * scale}px solid ${obj.stroke}`
                    : undefined,
                  borderRadius:
                    obj.shapeKind === "circle"
                      ? "50%"
                      : (obj.borderRadius || 0) * scale,
                  color: obj.color,
                  fontSize: obj.fontSize * scale,
                  fontWeight: obj.fontWeight,
                  textAlign: obj.align,
                  zIndex: obj.zIndex,
                  boxSizing: "border-box",
                  padding: 8 * scale,
                  display: "flex",
                  flexDirection: "column",
                  justifyContent: "center",
                  whiteSpace: "pre-wrap",
                  overflowWrap: "anywhere",
                  transform: `rotate(${obj.rotation || 0}deg)`,
                }}
              >
                {obj.title !== undefined && (
                  <strong
                    style={{ fontSize: "1.05em", marginBottom: 4 * scale }}
                  >
                    {obj.title}
                  </strong>
                )}
                {obj.subtitle !== undefined && (
                  <div style={{ fontSize: ".85em", opacity: 0.85 }}>
                    {obj.subtitle}
                  </div>
                )}
                {obj.metricValue !== undefined && (
                  <strong
                    style={{ fontSize: "1.7em", color: document.accentColor }}
                  >
                    {obj.metricValue}
                  </strong>
                )}
                {obj.metricLabel !== undefined && (
                  <div style={{ fontSize: ".85em", opacity: 0.8 }}>
                    {obj.metricLabel}
                  </div>
                )}
                {obj.text !== undefined && <div>{obj.text}</div>}
              </div>
            ))}
        </div>
        <button
          className="vscode-icon-btn"
          aria-label="Next slide"
          disabled={index >= document.slides.length - 1}
          onClick={() =>
            setIndex((i) => Math.min(document.slides.length - 1, i + 1))
          }
        >
          <ChevronRight size={24} />
        </button>
      </div>
      {showNotes && (
        <div
          style={{
            minHeight: 90,
            maxHeight: 170,
            overflow: "auto",
            whiteSpace: "pre-wrap",
            background: "#161920",
            borderTop: "1px solid #333",
            padding: "14px 24px",
            fontSize: 13,
          }}
        >
          <strong style={{ color: "#e57351" }}>Speaker notes: </strong>
          {slide.notes || "No speaker notes for this slide."}
        </div>
      )}
      <div
        style={{ padding: 8, fontSize: 11, textAlign: "center", color: "#778" }}
      >
        Arrow keys to navigate · Esc to exit
      </div>
    </div>
  );
};
