import { useCallback, useEffect, useState } from "react";
import { ChevronLeft, ChevronRight, X } from "lucide-react";
import type { Slide } from "../types";
import { SlideRenderer } from "./SlideRenderer";

interface PresentationViewProps {
  slides: Slide[];
  startIndex: number;
  onExit: (lastIndex: number) => void;
}

export function PresentationView({ slides, startIndex, onExit }: PresentationViewProps) {
  const [index, setIndex] = useState(startIndex);
  const [fit, setFit] = useState(1);

  useEffect(() => {
    const update = () => {
      setFit(Math.min(window.innerWidth / 960, window.innerHeight / 540) * 0.96);
    };
    update();
    window.addEventListener("resize", update);
    return () => window.removeEventListener("resize", update);
  }, []);

  const go = useCallback(
    (next: number) => setIndex(Math.max(0, Math.min(slides.length - 1, next))),
    [slides.length]
  );

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") onExit(index);
      else if (e.key === "ArrowRight" || e.key === "ArrowDown" || e.key === " " || e.key === "PageDown") {
        e.preventDefault();
        go(index + 1);
      } else if (e.key === "ArrowLeft" || e.key === "ArrowUp" || e.key === "PageUp") {
        e.preventDefault();
        go(index - 1);
      } else if (e.key === "Home") {
        e.preventDefault();
        go(0);
      } else if (e.key === "End") {
        e.preventDefault();
        go(slides.length - 1);
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [index, slides.length, go, onExit]);

  const slide = slides[index];

  return (
    <div className="present-overlay" role="dialog" aria-label="Slide show">
      <div className="present-stage" onClick={() => go(index + 1)}>
        <div key={slide.id} className="present-slide">
          <SlideRenderer slide={slide} scale={fit} />
        </div>
      </div>
      <div className="present-hud">
        <span className="present-count">
          Slide {index + 1} of {slides.length}
        </span>
        <span className="present-controls">
          <button
            type="button"
            aria-label="Previous slide"
            title="Previous"
            className="present-btn"
            disabled={index === 0}
            onClick={(e) => {
              e.stopPropagation();
              go(index - 1);
            }}
          >
            <ChevronLeft size={18} />
          </button>
          <button
            type="button"
            aria-label="Next slide"
            title="Next"
            className="present-btn"
            disabled={index === slides.length - 1}
            onClick={(e) => {
              e.stopPropagation();
              go(index + 1);
            }}
          >
            <ChevronRight size={18} />
          </button>
          <button
            type="button"
            aria-label="End slide show"
            title="End show (Esc)"
            className="present-btn"
            onClick={(e) => {
              e.stopPropagation();
              onExit(index);
            }}
          >
            <X size={18} />
          </button>
        </span>
      </div>
    </div>
  );
}
