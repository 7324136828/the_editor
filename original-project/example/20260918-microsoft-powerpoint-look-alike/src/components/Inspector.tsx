import { ArrowDown, ArrowUp, X } from "lucide-react";
import type { SlideObject } from "../types";

export interface DesignIdea {
  id: string;
  name: string;
  background: string;
  ink: string;
  hint: string;
}

export const designIdeas: DesignIdea[] = [
  { id: "focus", name: "Midnight focus", background: "#101a2d", ink: "#ffffff", hint: "Dark stage, light type" },
  { id: "paper", name: "Warm paper", background: "#faf7f3", ink: "#1d2b45", hint: "Soft editorial light" },
  { id: "wash", name: "Accent wash", background: "var(--accent-soft)", ink: "var(--ink)", hint: "Theme-tinted canvas" },
  { id: "slate", name: "Slate split", background: "#eef0f4", ink: "#1d2b45", hint: "Cool neutral panel" },
];

interface InspectorProps {
  object: SlideObject | null;
  accentHex: string;
  onUpdate: (id: string, patch: Partial<SlideObject>) => void;
  onBringForward: () => void;
  onSendBackward: () => void;
  onApplyIdea: (idea: DesignIdea) => void;
  onClose: () => void;
}

const toHex = (value: string | undefined, fallback: string) =>
  value && /^#[0-9a-fA-F]{6}$/.test(value) ? value : fallback;

const ACCENT = "var(--accent)";
const INK = "var(--ink)";

function NumberField({
  label,
  value,
  onChange,
  min,
  max,
}: {
  label: string;
  value: number;
  onChange: (v: number) => void;
  min?: number;
  max?: number;
}) {
  return (
    <label className="insp-num">
      <span>{label}</span>
      <input
        type="number"
        value={Math.round(value)}
        min={min}
        max={max}
        onChange={(e) => {
          const v = Number(e.currentTarget.value);
          if (Number.isFinite(v)) onChange(v);
        }}
      />
    </label>
  );
}

export function Inspector({
  object,
  accentHex,
  onUpdate,
  onBringForward,
  onSendBackward,
  onApplyIdea,
  onClose,
}: InspectorProps) {
  const title = object ? "Format Shape" : "Design Ideas";
  return (
    <aside className="inspector" aria-label={title}>
      <div className="inspector-head">
        <span className="inspector-title">{title}</span>
        <button type="button" className="inspector-close" aria-label="Close pane" title="Close" onClick={onClose}>
          <X size={14} />
        </button>
      </div>
      {object ? (
        <div className="inspector-body">
          <section className="insp-section">
            <span className="insp-label">Fill</span>
            <div className="insp-swatches">
              <button
                type="button"
                className="swatch"
                style={{ background: accentHex }}
                title="Theme accent"
                aria-label="Theme accent fill"
                onClick={() => onUpdate(object.id, { fill: ACCENT })}
              />
              {["#101a2d", "#ffffff", "#8a94a6", "#f2c14e"].map((c) => (
                <button
                  key={c}
                  type="button"
                  className="swatch"
                  style={{ background: c }}
                  title={c}
                  aria-label={`Fill ${c}`}
                  onClick={() => onUpdate(object.id, { fill: c })}
                />
              ))}
              <input
                type="color"
                className="swatch-picker"
                value={toHex(object.fill, accentHex)}
                aria-label="Custom fill color"
                title="Custom fill color"
                onChange={(e) => onUpdate(object.id, { fill: e.currentTarget.value })}
              />
            </div>
          </section>
          <section className="insp-section">
            <span className="insp-label">Text color</span>
            <div className="insp-swatches">
              <button
                type="button"
                className="swatch swatch-ink"
                style={{ background: "#1d2b45" }}
                title="Theme ink"
                aria-label="Theme ink text color"
                onClick={() => onUpdate(object.id, { color: INK })}
              />
              <button
                type="button"
                className="swatch"
                style={{ background: "#ffffff" }}
                title="White"
                aria-label="White text color"
                onClick={() => onUpdate(object.id, { color: "#ffffff" })}
              />
              <button
                type="button"
                className="swatch"
                style={{ background: accentHex }}
                title="Theme accent"
                aria-label="Accent text color"
                onClick={() => onUpdate(object.id, { color: ACCENT })}
              />
              <input
                type="color"
                className="swatch-picker"
                value={toHex(object.color, "#1d2b45")}
                aria-label="Custom text color"
                title="Custom text color"
                onChange={(e) => onUpdate(object.id, { color: e.currentTarget.value })}
              />
            </div>
          </section>
          <section className="insp-section">
            <div className="insp-row">
              <span className="insp-label">Opacity</span>
              <span className="insp-value">{Math.round((object.opacity ?? 1) * 100)}%</span>
            </div>
            <input
              type="range"
              min={5}
              max={100}
              value={Math.round((object.opacity ?? 1) * 100)}
              aria-label="Opacity"
              onChange={(e) => onUpdate(object.id, { opacity: Number(e.currentTarget.value) / 100 })}
            />
          </section>
          <section className="insp-section">
            <span className="insp-label">Position</span>
            <div className="insp-grid">
              <NumberField label="X" value={object.x} onChange={(v) => onUpdate(object.id, { x: v })} />
              <NumberField label="Y" value={object.y} onChange={(v) => onUpdate(object.id, { y: v })} />
              <NumberField label="W" value={object.width} min={8} onChange={(v) => onUpdate(object.id, { width: v })} />
              <NumberField label="H" value={object.height} min={0} onChange={(v) => onUpdate(object.id, { height: v })} />
              <NumberField
                label="Rotate"
                value={object.rotation ?? 0}
                min={0}
                max={360}
                onChange={(v) => onUpdate(object.id, { rotation: v })}
              />
            </div>
          </section>
          <section className="insp-section">
            <span className="insp-label">Arrange</span>
            <div className="insp-arrange">
              <button type="button" className="insp-btn" onClick={onBringForward}>
                <ArrowUp size={13} />
                Bring forward
              </button>
              <button type="button" className="insp-btn" onClick={onSendBackward}>
                <ArrowDown size={13} />
                Send backward
              </button>
            </div>
          </section>
        </div>
      ) : (
        <div className="inspector-body">
          <p className="insp-hint">Design ideas for this slide. Select a card to apply it.</p>
          <div className="idea-grid">
            {designIdeas.map((idea) => (
              <button
                key={idea.id}
                type="button"
                className="idea-card"
                onClick={() => onApplyIdea(idea)}
                title={idea.hint}
              >
                <span className="idea-preview" style={{ background: idea.background }}>
                  <span className="idea-line" style={{ background: accentHex }} />
                  <span className="idea-line idea-line-lg" style={{ background: idea.ink === "var(--ink)" ? "#1d2b45" : idea.ink }} />
                  <span className="idea-line idea-line-sm" style={{ background: idea.ink === "var(--ink)" ? "#1d2b45" : idea.ink, opacity: 0.55 }} />
                </span>
                <span className="idea-name">{idea.name}</span>
              </button>
            ))}
          </div>
        </div>
      )}
    </aside>
  );
}
