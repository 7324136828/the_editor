import { useState } from "react";
import { Copy, Plus, Trash2 } from "lucide-react";
import type { Slide } from "../types";
import { SlideRenderer } from "./SlideRenderer";

export interface SlidePaneProps {
  slides: Slide[];
  activeSlideId: string;
  onSelect: (id: string) => void;
  onAdd: () => void;
  onDuplicate: () => void;
  onDelete: () => void;
  onReorder: (sourceId: string, targetId: string) => void;
}

const THUMB_SCALE = 214 / 960;

export function SlidePane({
  slides,
  activeSlideId,
  onSelect,
  onAdd,
  onDuplicate,
  onDelete,
  onReorder,
}: SlidePaneProps) {
  const [dragOverId, setDragOverId] = useState<string | null>(null);
  const canDelete = slides.length > 1;

  return (
    <aside className="slide-pane" aria-label="Slide thumbnails">
      <div className="slide-pane-scroll">
        <div className="slides-section-row">
          <span>SLIDES</span>
          <span>{slides.length}</span>
        </div>
        <ol className="thumb-list">
          {slides.map((slide, i) => {
            const active = slide.id === activeSlideId;
            return (
              <li
                key={slide.id}
                className={`thumb-item${active ? " thumb-item-active" : ""}${dragOverId === slide.id ? " thumb-item-drop" : ""}`}
                draggable
                onDragStart={(e) => {
                  e.dataTransfer.setData("text/plain", slide.id);
                  e.dataTransfer.effectAllowed = "move";
                }}
                onDragOver={(e) => {
                  e.preventDefault();
                  e.dataTransfer.dropEffect = "move";
                  setDragOverId(slide.id);
                }}
                onDragLeave={() => setDragOverId((cur) => (cur === slide.id ? null : cur))}
                onDrop={(e) => {
                  e.preventDefault();
                  const sourceId = e.dataTransfer.getData("text/plain");
                  setDragOverId(null);
                  if (sourceId && sourceId !== slide.id) onReorder(sourceId, slide.id);
                }}
                onDragEnd={() => setDragOverId(null)}
              >
                <button
                  type="button"
                  className="thumb-card"
                  onClick={() => onSelect(slide.id)}
                  aria-label={`Slide ${i + 1}: ${slide.title}`}
                  aria-current={active}
                >
                  <span className="thumb-number">{i + 1}</span>
                  <span className="thumb-frame">
                    <SlideRenderer slide={slide} scale={THUMB_SCALE} />
                  </span>
                </button>
                {active && (
                  <span className="thumb-actions">
                    <button
                      type="button"
                      className="thumb-action"
                      aria-label="Duplicate slide"
                      title="Duplicate slide"
                      onClick={onDuplicate}
                    >
                      <Copy size={12} />
                    </button>
                    <button
                      type="button"
                      className="thumb-action"
                      aria-label="Delete slide"
                      title={canDelete ? "Delete slide" : "Cannot delete the last slide"}
                      onClick={onDelete}
                      disabled={!canDelete}
                    >
                      <Trash2 size={12} />
                    </button>
                  </span>
                )}
              </li>
            );
          })}
        </ol>
        <button type="button" className="new-slide-btn" onClick={onAdd}>
          <Plus size={14} />
          New slide
        </button>
      </div>
    </aside>
  );
}
