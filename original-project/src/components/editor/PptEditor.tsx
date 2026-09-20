import React, { useEffect, useState } from "react";
import { PptDocumentModel, PptSlide, PptShapeObject } from "../../types/office";
import {
  Play,
  Plus,
  Type,
  Square,
  BarChart2,
  Trash2,
  ZoomIn,
  ZoomOut,
  Copy,
  ArrowUp,
  ArrowDown,
} from "lucide-react";
import { PresentationModal } from "./PresentationModal";
import "../../styles/editors.css";

interface PptEditorProps {
  document: PptDocumentModel;
  onChangeDocument?: (doc: PptDocumentModel) => void;
  selectedObjectId?: string | null;
  onSelectObject?: (id: string) => void;
}

export const PptEditor: React.FC<PptEditorProps> = ({
  document,
  onChangeDocument,
  selectedObjectId,
  onSelectObject,
}) => {
  const [zoom, setZoom] = useState(80);
  const [isPlaying, setIsPlaying] = useState(false);
  const [localSelected, setLocalSelected] = useState<string | null>(
    selectedObjectId || null,
  );
  const [drag, setDrag] = useState<{
    id: string;
    clientX: number;
    clientY: number;
    x: number;
    y: number;
    dx: number;
    dy: number;
  } | null>(null);
  const slide =
    document.slides.find((s) => s.id === document.activeSlideId) ||
    document.slides[0];
  const slideIndex = document.slides.findIndex((s) => s.id === slide?.id);
  const selected = slide?.objects.find((obj) => obj.id === localSelected);
  const scale = zoom / 100;
  useEffect(() => {
    setLocalSelected(selectedObjectId || null);
  }, [selectedObjectId]);
  const select = (id: string) => {
    setLocalSelected(id);
    onSelectObject?.(id);
  };
  const commit = (next: PptDocumentModel) =>
    onChangeDocument?.({
      ...next,
      slides: next.slides.map((s, index) => ({ ...s, slideNumber: index + 1 })),
    });
  const updateSlide = (patch: Partial<PptSlide>) =>
    commit({
      ...document,
      slides: document.slides.map((s) =>
        s.id === slide.id ? { ...s, ...patch } : s,
      ),
    });
  const updateObject = (id: string, patch: Partial<PptShapeObject>) =>
    updateSlide({
      objects: slide.objects.map((obj) =>
        obj.id === id ? { ...obj, ...patch } : obj,
      ),
    });
  const removeObject = () => {
    if (!selected || selected.locked) return;
    updateSlide({
      objects: slide.objects.filter((obj) => obj.id !== selected.id),
    });
    select("");
  };
  const addObject = (kind: "text" | "shape" | "metric") => {
    const obj: PptShapeObject = {
      id: `object-${crypto.randomUUID()}`,
      kind,
      shapeKind:
        kind === "metric"
          ? "metric-box"
          : kind === "shape"
            ? "rounded-rect"
            : undefined,
      x: 80,
      y: 140,
      width: kind === "metric" ? 220 : 300,
      height: kind === "metric" ? 120 : 90,
      fill: kind === "text" ? "transparent" : "#263650",
      color: "#ffffff",
      fontSize: 22,
      zIndex: Math.max(0, ...slide.objects.map((item) => item.zIndex)) + 1,
      borderRadius: kind === "text" ? 0 : 8,
      ...(kind === "metric"
        ? { metricValue: "100", metricLabel: "New metric" }
        : { text: kind === "text" ? "New text box" : "New shape" }),
    };
    updateSlide({ objects: [...slide.objects, obj] });
    select(obj.id);
  };
  const addSlide = (duplicate = false) => {
    const id = `slide-${crypto.randomUUID()}`;
    const next: PptSlide =
      duplicate && slide
        ? {
            ...slide,
            id,
            title: `${slide.title} (copy)`,
            objects: slide.objects.map((obj) => ({
              ...obj,
              id: `object-${crypto.randomUUID()}`,
            })),
          }
        : {
            id,
            slideNumber: document.slides.length + 1,
            title: "Untitled slide",
            layout: "title",
            background: "#181b22",
            notes: "",
            transition: "none",
            objects: [
              {
                id: `object-${crypto.randomUUID()}`,
                kind: "text",
                x: 60,
                y: 50,
                width: 800,
                height: 90,
                text: "New slide",
                fill: "transparent",
                color: "#ffffff",
                fontSize: 36,
                fontWeight: "bold",
                zIndex: 1,
              },
            ],
          };
    const slides = [...document.slides];
    slides.splice(slideIndex + 1, 0, next);
    commit({ ...document, slides, activeSlideId: id });
    select("");
  };
  const moveSlide = (direction: number) => {
    const target = slideIndex + direction;
    if (target < 0 || target >= document.slides.length) return;
    const slides = [...document.slides];
    [slides[slideIndex], slides[target]] = [slides[target], slides[slideIndex]];
    commit({ ...document, slides });
  };
  if (!slide)
    return (
      <div className="office-editor">
        <button className="vscode-btn" onClick={() => addSlide()}>
          Create first slide
        </button>
      </div>
    );
  return (
    <div className="office-editor">
      <div className="office-toolbar">
        <span className="office-badge ppt-badge">SLIDES</span>
        <button className="vscode-btn-secondary" onClick={() => addSlide()}>
          <Plus size={13} /> Slide
        </button>
        <button
          className="vscode-icon-btn"
          title="Duplicate slide"
          aria-label="Duplicate slide"
          onClick={() => addSlide(true)}
        >
          <Copy size={14} />
        </button>
        <button
          className="vscode-icon-btn"
          title="Delete slide"
          aria-label="Delete slide"
          disabled={document.slides.length < 2}
          onClick={() => {
            const slides = document.slides.filter((s) => s.id !== slide.id);
            commit({
              ...document,
              slides,
              activeSlideId: slides[Math.max(0, slideIndex - 1)].id,
            });
            select("");
          }}
        >
          <Trash2 size={14} />
        </button>
        <button
          className="vscode-icon-btn"
          title="Move slide earlier"
          aria-label="Move slide earlier"
          disabled={slideIndex === 0}
          onClick={() => moveSlide(-1)}
        >
          <ArrowUp size={14} />
        </button>
        <button
          className="vscode-icon-btn"
          title="Move slide later"
          aria-label="Move slide later"
          disabled={slideIndex === document.slides.length - 1}
          onClick={() => moveSlide(1)}
        >
          <ArrowDown size={14} />
        </button>
        <button
          className="vscode-btn-secondary"
          onClick={() => addObject("text")}
        >
          <Type size={13} /> Text
        </button>
        <button
          className="vscode-btn-secondary"
          onClick={() => addObject("shape")}
        >
          <Square size={13} /> Shape
        </button>
        <button
          className="vscode-btn-secondary"
          onClick={() => addObject("metric")}
        >
          <BarChart2 size={13} /> Metric
        </button>
        <span className="toolbar-spacer" />
        <button
          className="vscode-icon-btn"
          title="Zoom out"
          aria-label="Zoom out"
          onClick={() => setZoom(Math.max(40, zoom - 10))}
        >
          <ZoomOut size={14} />
        </button>
        <span>{zoom}%</span>
        <button
          className="vscode-icon-btn"
          title="Zoom in"
          aria-label="Zoom in"
          onClick={() => setZoom(Math.min(150, zoom + 10))}
        >
          <ZoomIn size={14} />
        </button>
        <button className="vscode-btn" onClick={() => setIsPlaying(true)}>
          <Play size={13} /> Present
        </button>
      </div>
      <div className="ppt-main">
        <nav className="ppt-filmstrip" aria-label="Slides">
          {document.slides.map((item, index) => (
            <button
              key={item.id}
              className={`ppt-thumb ${item.id === slide.id ? "active" : ""}`}
              onClick={() => {
                commit({ ...document, activeSlideId: item.id });
                select("");
              }}
            >
              <strong>{index + 1}</strong>
              <span>{item.title}</span>
              <span style={{ color: "#888", fontSize: 10 }}>
                {item.objects.length} objects
              </span>
            </button>
          ))}
        </nav>
        <div className="ppt-workspace">
          <div className="office-toolbar compact-toolbar">
            <label htmlFor="slide-title">Slide title</label>
            <input
              id="slide-title"
              className="vscode-input"
              value={slide.title}
              onChange={(e) => updateSlide({ title: e.target.value })}
              style={{ flex: 1 }}
            />
            <label htmlFor="slide-background">Background</label>
            <input
              id="slide-background"
              className="vscode-input"
              value={slide.background}
              onChange={(e) => updateSlide({ background: e.target.value })}
              style={{ width: 135 }}
            />
          </div>
          <div className="ppt-scroll">
            <div
              className="ppt-stage"
              style={{
                width: 920 * scale,
                height: 518 * scale,
                background: slide.background,
              }}
              onClick={() => select("")}
            >
              {slide.objects
                .filter((obj) => obj.visible !== false)
                .map((obj) => {
                  const isSelected = selected?.id === obj.id;
                  const activeDrag = drag?.id === obj.id ? drag : null;
                  return (
                    <div
                      key={obj.id}
                      className={`ppt-object ${isSelected ? "selected" : ""}`}
                      tabIndex={0}
                      aria-label={`Slide object: ${obj.text || obj.title || obj.metricLabel || obj.kind}`}
                      style={{
                        left: (obj.x + (activeDrag?.dx || 0)) * scale,
                        top: (obj.y + (activeDrag?.dy || 0)) * scale,
                        width: obj.width * scale,
                        height: obj.height * scale,
                        background: obj.fill,
                        border: obj.stroke
                          ? `${obj.strokeWidth || 1}px solid ${obj.stroke}`
                          : undefined,
                        borderRadius:
                          obj.shapeKind === "circle"
                            ? "50%"
                            : obj.borderRadius || 0,
                        color: obj.color,
                        fontSize: obj.fontSize * scale,
                        fontWeight: obj.fontWeight,
                        textAlign: obj.align,
                        zIndex: obj.zIndex,
                        transform: `rotate(${obj.rotation || 0}deg)`,
                        padding: 8 * scale,
                      }}
                      onClick={(e) => {
                        e.stopPropagation();
                        select(obj.id);
                      }}
                      onKeyDown={(e) => {
                        if (
                          (e.target as HTMLElement).tagName === "TEXTAREA" ||
                          obj.locked
                        )
                          return;
                        if (e.key === "Delete" || e.key === "Backspace") {
                          e.preventDefault();
                          removeObject();
                        }
                        const moves: Record<string, [number, number]> = {
                          ArrowLeft: [-1, 0],
                          ArrowRight: [1, 0],
                          ArrowUp: [0, -1],
                          ArrowDown: [0, 1],
                        };
                        if (moves[e.key]) {
                          e.preventDefault();
                          const [dx, dy] = moves[e.key];
                          updateObject(obj.id, {
                            x: Math.max(0, obj.x + dx * (e.shiftKey ? 10 : 1)),
                            y: Math.max(0, obj.y + dy * (e.shiftKey ? 10 : 1)),
                          });
                        }
                      }}
                    >
                      {obj.title !== undefined && (
                        <strong style={{ fontSize: "1.05em", marginBottom: 4 }}>
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
                          style={{
                            fontSize: "1.7em",
                            color: document.accentColor,
                          }}
                        >
                          {obj.metricValue}
                        </strong>
                      )}
                      {obj.metricLabel !== undefined && (
                        <div style={{ fontSize: ".85em", opacity: 0.8 }}>
                          {obj.metricLabel}
                        </div>
                      )}
                      {obj.text !== undefined &&
                        (isSelected && !obj.locked ? (
                          <textarea
                            aria-label="Object text"
                            value={obj.text}
                            rows={Math.max(1, obj.text.split("\n").length)}
                            onChange={(e) =>
                              updateObject(obj.id, { text: e.target.value })
                            }
                            onClick={(e) => e.stopPropagation()}
                          />
                        ) : (
                          <div>{obj.text}</div>
                        ))}
                      {isSelected && !obj.locked && (
                        <button
                          type="button"
                          title="Drag to move object"
                          aria-label="Drag to move object"
                          style={{
                            position: "absolute",
                            right: 0,
                            top: 0,
                            cursor: "move",
                            border: 0,
                            color: "white",
                            background: "#007acc",
                            padding: "0 4px",
                            fontSize: 12,
                            touchAction: "none",
                          }}
                          onPointerDown={(e) => {
                            e.stopPropagation();
                            e.currentTarget.setPointerCapture(e.pointerId);
                            setDrag({
                              id: obj.id,
                              clientX: e.clientX,
                              clientY: e.clientY,
                              x: obj.x,
                              y: obj.y,
                              dx: 0,
                              dy: 0,
                            });
                          }}
                          onPointerMove={(e) => {
                            if (drag?.id === obj.id)
                              setDrag({
                                ...drag,
                                dx: (e.clientX - drag.clientX) / scale,
                                dy: (e.clientY - drag.clientY) / scale,
                              });
                          }}
                          onPointerUp={(e) => {
                            if (drag?.id === obj.id) {
                              updateObject(obj.id, {
                                x: Math.round(
                                  Math.max(0, Math.min(900, drag.x + drag.dx)),
                                ),
                                y: Math.round(
                                  Math.max(0, Math.min(500, drag.y + drag.dy)),
                                ),
                              });
                              setDrag(null);
                            }
                            e.currentTarget.releasePointerCapture(e.pointerId);
                          }}
                          onPointerCancel={() => setDrag(null)}
                        >
                          ↔
                        </button>
                      )}
                    </div>
                  );
                })}
            </div>
          </div>
          {selected && (
            <div className="ppt-properties">
              <strong>Selected {selected.kind}</strong>
              {(
                [
                  "x",
                  "y",
                  "width",
                  "height",
                  "fontSize",
                  "rotation",
                  "zIndex",
                ] as const
              ).map((key) => (
                <label key={key}>
                  {
                    {
                      x: "X",
                      y: "Y",
                      width: "Width",
                      height: "Height",
                      fontSize: "Font",
                      rotation: "Angle",
                      zIndex: "Layer",
                    }[key]
                  }
                  <input
                    aria-label={`Object ${key}`}
                    className="vscode-input"
                    type="number"
                    value={selected[key] || 0}
                    disabled={selected.locked}
                    onChange={(e) =>
                      updateObject(selected.id, {
                        [key]: Math.max(
                          ["width", "height", "fontSize"].includes(key) ? 1 : 0,
                          Math.min(2000, Number(e.target.value)),
                        ),
                      })
                    }
                  />
                </label>
              ))}
              <label>
                Color
                <input
                  aria-label="Object text color"
                  type="color"
                  value={
                    /^#[\da-f]{6}$/i.test(selected.color)
                      ? selected.color
                      : "#ffffff"
                  }
                  disabled={selected.locked}
                  onChange={(e) =>
                    updateObject(selected.id, { color: e.target.value })
                  }
                />
              </label>
              <label>
                Fill
                <input
                  aria-label="Object fill"
                  className="vscode-input"
                  style={{ width: 85 }}
                  value={selected.fill}
                  disabled={selected.locked}
                  onChange={(e) =>
                    updateObject(selected.id, { fill: e.target.value })
                  }
                />
              </label>
              <select
                aria-label="Object alignment"
                className="vscode-input"
                value={selected.align || "left"}
                disabled={selected.locked}
                onChange={(e) =>
                  updateObject(selected.id, {
                    align: e.target.value as PptShapeObject["align"],
                  })
                }
              >
                <option>left</option>
                <option>center</option>
                <option>right</option>
              </select>
              <button
                className="vscode-btn-secondary"
                disabled={selected.locked}
                onClick={() =>
                  updateObject(selected.id, {
                    fontWeight:
                      selected.fontWeight === "bold" ? "normal" : "bold",
                  })
                }
              >
                Bold
              </button>
              <button
                className="vscode-btn-secondary"
                onClick={() =>
                  updateObject(selected.id, { locked: !selected.locked })
                }
              >
                {selected.locked ? "Unlock" : "Lock"}
              </button>
              <button
                className="vscode-icon-btn"
                title="Delete selected object"
                aria-label="Delete selected object"
                disabled={selected.locked}
                onClick={removeObject}
              >
                <Trash2 size={14} />
              </button>
              {(
                [
                  "text",
                  "title",
                  "subtitle",
                  "metricValue",
                  "metricLabel",
                ] as const
              )
                .filter((key) => selected[key] !== undefined)
                .map((key) => (
                  <label key={key} style={{ width: "100%" }}>
                    {key}
                    <textarea
                      aria-label={`Selected object ${key}`}
                      className="vscode-input"
                      value={selected[key]}
                      disabled={selected.locked}
                      onChange={(e) =>
                        updateObject(selected.id, { [key]: e.target.value })
                      }
                    />
                  </label>
                ))}
            </div>
          )}
          <div className="ppt-notes">
            <span>Speaker notes</span>
            <textarea
              className="vscode-input"
              aria-label="Speaker notes"
              value={slide.notes}
              onChange={(e) => updateSlide({ notes: e.target.value })}
              placeholder="Add notes for your presentation…"
            />
          </div>
        </div>
      </div>
      {isPlaying && (
        <PresentationModal
          document={document}
          onClose={() => setIsPlaying(false)}
        />
      )}
    </div>
  );
};
