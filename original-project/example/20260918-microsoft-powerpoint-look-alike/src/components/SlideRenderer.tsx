import { useEffect, useRef, useState } from "react";
import { RotateCw } from "lucide-react";
import type { Slide, SlideObject } from "../types";

export interface SlideRendererProps {
  slide: Slide;
  scale?: number;
  interactive?: boolean;
  selectedObjectId?: string | null;
  onSelectObject?: (id: string | null) => void;
  onUpdateObject?: (id: string, patch: Partial<SlideObject>) => void;
  onCommitObject?: (id: string, patch: Partial<SlideObject>) => void;
  onEditingChange?: (editing: boolean) => void;
}

const MIN_SIZE = 8;
const HANDLES: { key: string; dx: number; dy: number; cursor: string }[] = [
  { key: "nw", dx: -1, dy: -1, cursor: "nwse-resize" },
  { key: "n", dx: 0, dy: -1, cursor: "ns-resize" },
  { key: "ne", dx: 1, dy: -1, cursor: "nesw-resize" },
  { key: "e", dx: 1, dy: 0, cursor: "ew-resize" },
  { key: "se", dx: 1, dy: 1, cursor: "nwse-resize" },
  { key: "s", dx: 0, dy: 1, cursor: "ns-resize" },
  { key: "sw", dx: -1, dy: 1, cursor: "nesw-resize" },
  { key: "w", dx: -1, dy: 0, cursor: "ew-resize" },
];

interface Gesture {
  mode: "move" | "resize" | "rotate";
  id: string;
  startClientX: number;
  startClientY: number;
  orig: SlideObject;
  dirX: number;
  dirY: number;
  moved: boolean;
  patch: Partial<SlideObject>;
}

function handlePosition(dx: number, dy: number): React.CSSProperties {
  return {
    left: `${((dx + 1) / 2) * 100}%`,
    top: `${((dy + 1) / 2) * 100}%`,
  };
}

function ObjectBody({ obj }: { obj: SlideObject }) {
  if (obj.kind === "text") {
    return (
      <div
        className="obj-text"
        style={{
          color: obj.color ?? "#1d2b45",
          fontFamily: obj.fontFamily ?? '"Segoe UI", "Segoe UI Variable", system-ui, sans-serif',
          fontSize: obj.fontSize ?? 18,
          fontWeight: obj.fontWeight ?? 400,
          fontStyle: obj.fontStyle ?? "normal",
          textDecoration: obj.textDecoration === "underline" ? "underline" : "none",
          textAlign: obj.align ?? "left",
          lineHeight: obj.lineHeight ?? 1.25,
          letterSpacing: obj.letterSpacing != null ? `${obj.letterSpacing}px` : undefined,
        }}
      >
        {obj.text}
      </div>
    );
  }
  if (obj.kind === "chart") {
    const data = obj.chartData ?? [40, 60, 50, 80, 70];
    const max = Math.max(...data, 1);
    return (
      <div className="obj-chart">
        {data.map((v, i) => (
          <div
            key={i}
            className="obj-chart-bar"
            style={{ height: `${(v / max) * 100}%`, background: obj.fill ?? "var(--accent)" }}
          />
        ))}
      </div>
    );
  }
  if (obj.kind === "line") {
    return (
      <div className="obj-line-box">
        <div
          className="obj-line"
          style={{
            height: Math.max(obj.strokeWidth ?? 2, 1),
            background: obj.stroke ?? "#1d2b45",
          }}
        />
      </div>
    );
  }
  const radius =
    obj.shape === "ellipse" || obj.shape === "ring"
      ? "50%"
      : obj.shape === "pill"
        ? obj.height / 2
        : (obj.radius ?? 0);
  if (obj.shape === "ring") {
    return (
      <div
        className="obj-shape"
        style={{
          borderRadius: radius,
          border: `${obj.strokeWidth ?? 2}px solid ${obj.stroke ?? "var(--accent)"}`,
          background: "transparent",
        }}
      />
    );
  }
  return (
    <div
      className="obj-shape"
      style={{
        borderRadius: radius,
        background: obj.fill ?? "var(--accent)",
        border: obj.stroke ? `${obj.strokeWidth ?? 1}px solid ${obj.stroke}` : undefined,
      }}
    />
  );
}

export function SlideRenderer({
  slide,
  scale = 1,
  interactive = false,
  selectedObjectId = null,
  onSelectObject,
  onUpdateObject,
  onCommitObject,
  onEditingChange,
}: SlideRendererProps) {
  const surfaceRef = useRef<HTMLDivElement>(null);
  const gestureRef = useRef<Gesture | null>(null);
  const cancelEditRef = useRef(false);
  const [editingId, setEditingId] = useState<string | null>(null);
  const [draft, setDraft] = useState("");

  useEffect(() => {
    if (!interactive) setEditingId(null);
  }, [interactive]);

  useEffect(() => {
    return () => onEditingChange?.(false);
  }, [onEditingChange]);

  const endEditing = () => {
    const cancel = cancelEditRef.current;
    cancelEditRef.current = false;
    if (editingId && !cancel) {
      const obj = slide.objects.find((o) => o.id === editingId);
      if (obj && obj.text !== draft) {
        onCommitObject?.(editingId, { text: draft });
      }
    }
    setEditingId(null);
    onEditingChange?.(false);
  };

  const onGestureMove = (e: PointerEvent) => {
    const g = gestureRef.current;
    if (!g) return;
    const dx = (e.clientX - g.startClientX) / scale;
    const dy = (e.clientY - g.startClientY) / scale;
    if (!g.moved && Math.hypot(dx, dy) < 2) return;
    g.moved = true;
    const o = g.orig;
    let patch: Partial<SlideObject> = {};
    if (g.mode === "move") {
      patch = { x: Math.round(o.x + dx), y: Math.round(o.y + dy) };
    } else if (g.mode === "resize") {
      let { x, y, width, height } = o;
      if (g.dirX === 1) width = Math.max(MIN_SIZE, o.width + dx);
      if (g.dirX === -1) {
        width = Math.max(MIN_SIZE, o.width - dx);
        x = o.x + Math.min(dx, o.width - MIN_SIZE);
      }
      if (g.dirY === 1) height = Math.max(MIN_SIZE, o.height + dy);
      if (g.dirY === -1) {
        height = Math.max(MIN_SIZE, o.height - dy);
        y = o.y + Math.min(dy, o.height - MIN_SIZE);
      }
      patch = { x: Math.round(x), y: Math.round(y), width: Math.round(width), height: Math.round(height) };
    } else {
      const rect = surfaceRef.current?.getBoundingClientRect();
      if (!rect) return;
      const cx = rect.left + (o.x + o.width / 2) * scale;
      const cy = rect.top + (o.y + o.height / 2) * scale;
      const angle = (Math.atan2(e.clientY - cy, e.clientX - cx) * 180) / Math.PI + 90;
      patch = { rotation: Math.round(((angle % 360) + 360) % 360) };
    }
    g.patch = patch;
    onUpdateObject?.(g.id, patch);
  };

  const onGestureEnd = () => {
    const g = gestureRef.current;
    gestureRef.current = null;
    window.removeEventListener("pointermove", onGestureMove);
    window.removeEventListener("pointerup", onGestureEnd);
    if (g && g.moved) {
      onCommitObject?.(g.id, g.patch);
    }
  };

  const beginGesture = (e: React.PointerEvent, obj: SlideObject, mode: Gesture["mode"], dirX = 0, dirY = 0) => {
    if (!interactive || editingId) return;
    e.preventDefault();
    e.stopPropagation();
    onSelectObject?.(obj.id);
    gestureRef.current = {
      mode,
      id: obj.id,
      startClientX: e.clientX,
      startClientY: e.clientY,
      orig: { ...obj },
      dirX,
      dirY,
      moved: false,
      patch: {},
    };
    window.addEventListener("pointermove", onGestureMove);
    window.addEventListener("pointerup", onGestureEnd);
  };

  const startEditing = (obj: SlideObject) => {
    if (!interactive || obj.kind !== "text") return;
    setEditingId(obj.id);
    setDraft(obj.text ?? "");
    onSelectObject?.(obj.id);
    onEditingChange?.(true);
  };

  const ordered = [...slide.objects].sort((a, b) => (a.zIndex ?? 0) - (b.zIndex ?? 0));

  return (
    <div className="slide-outer" style={{ width: 960 * scale, height: 540 * scale }}>
      <div
        ref={surfaceRef}
        className={`slide-surface${interactive ? " slide-surface-interactive" : ""}`}
        style={{
          width: 960,
          height: 540,
          transform: `scale(${scale})`,
          background: slide.background,
        }}
        onPointerDown={interactive ? () => onSelectObject?.(null) : undefined}
      >
        {ordered.map((obj) => {
          const selected = interactive && selectedObjectId === obj.id;
          const editing = editingId === obj.id;
          const hitHeight = obj.kind === "line" ? Math.max(obj.height, 14) : obj.height;
          return (
            <div
              key={obj.id}
              className={`canvas-object${selected ? " canvas-object-selected" : ""}${interactive ? " canvas-object-live" : ""}`}
              style={{
                left: obj.x,
                top: obj.kind === "line" ? obj.y - (hitHeight - obj.height) / 2 : obj.y,
                width: obj.width,
                height: hitHeight,
                transform: obj.rotation ? `rotate(${obj.rotation}deg)` : undefined,
                opacity: obj.opacity ?? 1,
                zIndex: obj.zIndex ?? 0,
              }}
              onPointerDown={interactive && !editing ? (e) => beginGesture(e, obj, "move") : undefined}
              onDoubleClick={interactive ? () => startEditing(obj) : undefined}
            >
              <ObjectBody obj={obj} />
              {selected && !editing && (
                <div className="selection-chrome">
                  <div className="selection-rotate-stem" />
                  <button
                    type="button"
                    className="selection-rotate"
                    aria-label="Rotate object"
                    title="Rotate"
                    onPointerDown={(e) => beginGesture(e, obj, "rotate")}
                  >
                    <RotateCw size={11} strokeWidth={2.4} />
                  </button>
                  {HANDLES.map((h) => (
                    <span
                      key={h.key}
                      className="selection-handle"
                      style={{ ...handlePosition(h.dx, h.dy), cursor: h.cursor }}
                      onPointerDown={(e) => beginGesture(e, obj, "resize", h.dx, h.dy)}
                    />
                  ))}
                </div>
              )}
              {editing && (
                <textarea
                  className="obj-editor"
                  value={draft}
                  autoFocus
                  aria-label="Edit text"
                  style={{
                    color: obj.color ?? "#1d2b45",
                    fontFamily: obj.fontFamily ?? '"Segoe UI", system-ui, sans-serif',
                    fontSize: obj.fontSize ?? 18,
                    fontWeight: obj.fontWeight ?? 400,
                    fontStyle: obj.fontStyle ?? "normal",
                    textDecoration: obj.textDecoration === "underline" ? "underline" : "none",
                    textAlign: obj.align ?? "left",
                    lineHeight: obj.lineHeight ?? 1.25,
                    letterSpacing: obj.letterSpacing != null ? `${obj.letterSpacing}px` : undefined,
                  }}
                  onFocus={(e) => e.currentTarget.select()}
                  onChange={(e) => setDraft(e.currentTarget.value)}
                  onBlur={endEditing}
                  onKeyDown={(e) => {
                    if (e.key === "Escape") {
                      cancelEditRef.current = true;
                      e.currentTarget.blur();
                    }
                    e.stopPropagation();
                  }}
                  onPointerDown={(e) => e.stopPropagation()}
                />
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}
