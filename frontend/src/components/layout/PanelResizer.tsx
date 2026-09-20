import React, { useRef, useState } from "react";
import { clamp } from "../../services/panelLayout";
interface PanelResizerProps {
  label: string;
  orientation: "vertical" | "horizontal";
  value: number;
  min: number;
  max: number;
  direction?: 1 | -1;
  onResize: (value: number) => void;
  onReset: () => void;
  controls?: string;
}
export function PanelResizer({
  label,
  orientation,
  value,
  min,
  max,
  direction = 1,
  onResize,
  onReset,
  controls,
}: PanelResizerProps) {
  const drag = useRef<{
    start: number;
    size: number;
    pointerId: number;
  } | null>(null);
  const [dragging, setDragging] = useState(false);
  const end = (element: HTMLElement) => {
    const active = drag.current;
    drag.current = null;
    setDragging(false);
    if (active && element.hasPointerCapture(active.pointerId))
      element.releasePointerCapture(active.pointerId);
  };
  return (
    <div
      role="separator"
      tabIndex={0}
      aria-label={label}
      aria-orientation={orientation}
      aria-valuemin={min}
      aria-valuemax={max}
      aria-valuenow={Math.round(value)}
      aria-valuetext={`${Math.round(value)} pixels`}
      aria-controls={controls}
      className={`panel-resizer ${orientation} ${dragging ? "resizing" : ""}`}
      title={`${label}. Drag or use arrow keys; double-click to reset.`}
      onPointerDown={(e) => {
        if (e.button !== 0 || drag.current) return;
        e.preventDefault();
        e.currentTarget.focus();
        drag.current = {
          start: orientation === "vertical" ? e.clientX : e.clientY,
          size: value,
          pointerId: e.pointerId,
        };
        e.currentTarget.setPointerCapture(e.pointerId);
        setDragging(true);
      }}
      onPointerMove={(e) => {
        const active = drag.current;
        if (!active || active.pointerId !== e.pointerId) return;
        const coordinate = orientation === "vertical" ? e.clientX : e.clientY;
        onResize(
          clamp(
            active.size + (coordinate - active.start) * direction,
            min,
            max,
          ),
        );
      }}
      onPointerUp={(e) => {
        if (drag.current?.pointerId === e.pointerId) end(e.currentTarget);
      }}
      onPointerCancel={(e) => {
        if (drag.current?.pointerId !== e.pointerId) return;
        if (drag.current) onResize(drag.current.size);
        end(e.currentTarget);
      }}
      onLostPointerCapture={(e) => {
        if (drag.current?.pointerId === e.pointerId) {
          drag.current = null;
          setDragging(false);
        }
      }}
      onDoubleClick={onReset}
      onKeyDown={(e) => {
        const increase =
            orientation === "vertical" ? "ArrowRight" : "ArrowDown",
          decrease = orientation === "vertical" ? "ArrowLeft" : "ArrowUp";
        if (e.key === increase || e.key === decrease) {
          e.preventDefault();
          onResize(
            clamp(
              value +
                (e.key === increase ? 1 : -1) *
                  direction *
                  (e.shiftKey ? 40 : 10),
              min,
              max,
            ),
          );
        } else if (e.key === "Home") {
          e.preventDefault();
          onResize(min);
        } else if (e.key === "End") {
          e.preventDefault();
          onResize(max);
        } else if (e.key === "Enter") {
          e.preventDefault();
          onReset();
        } else if (e.key === "Escape" && drag.current) {
          onResize(drag.current.size);
          end(e.currentTarget);
        }
      }}
    />
  );
}
