import { CheckCircle2, LayoutGrid, MessageSquare, MonitorPlay, PanelLeft } from "lucide-react";
import type { ViewMode } from "../types";

interface StatusBarProps {
  slideIndex: number;
  slideCount: number;
  showNotes: boolean;
  onToggleNotes: () => void;
  onComments: () => void;
  viewMode: ViewMode;
  onViewMode: (mode: ViewMode) => void;
  zoom: number;
  onZoom: (zoom: number) => void;
  onPresent: () => void;
}

export function StatusBar({
  slideIndex,
  slideCount,
  showNotes,
  onToggleNotes,
  onComments,
  viewMode,
  onViewMode,
  zoom,
  onZoom,
  onPresent,
}: StatusBarProps) {
  return (
    <footer className="statusbar">
      <div className="statusbar-left">
        <span className="status-item">
          Slide {slideIndex + 1} of {slideCount}
        </span>
        <span className="status-item status-hide-sm">English (United States)</span>
        <span className="status-item status-hide-sm">
          <CheckCircle2 size={12} />
          Accessibility: Good to go
        </span>
      </div>
      <div className="statusbar-right">
        <button
          type="button"
          className={`status-btn${showNotes ? " status-btn-active" : ""}`}
          onClick={onToggleNotes}
          aria-pressed={showNotes}
          title="Notes"
        >
          Notes
        </button>
        <button type="button" className="status-btn" onClick={onComments} title="Comments">
          <MessageSquare size={12} />
          Comments
        </button>
        <span className="status-sep" />
        <button
          type="button"
          className={`status-view${viewMode === "normal" ? " status-view-active" : ""}`}
          aria-label="Normal view"
          title="Normal"
          onClick={() => onViewMode("normal")}
        >
          <PanelLeft size={14} />
        </button>
        <button
          type="button"
          className={`status-view${viewMode === "sorter" ? " status-view-active" : ""}`}
          aria-label="Slide Sorter view"
          title="Slide Sorter"
          onClick={() => onViewMode("sorter")}
        >
          <LayoutGrid size={14} />
        </button>
        <button
          type="button"
          className="status-view"
          aria-label="Slide Show"
          title="Slide Show"
          onClick={onPresent}
        >
          <MonitorPlay size={14} />
        </button>
        <span className="status-sep" />
        <span className="status-zoom-value">{zoom}%</span>
        <input
          type="range"
          className="zoom-slider"
          min={45}
          max={110}
          value={zoom}
          aria-label="Zoom"
          title="Zoom"
          onChange={(e) => onZoom(Number(e.currentTarget.value))}
        />
      </div>
    </footer>
  );
}
