import React from "react";
import { OfficeDocument } from "../../types/office";
import { GitBranch, PanelLeft, PanelBottom, PanelRight } from "lucide-react";

interface StatusBarProps {
  activeDoc: OfficeDocument | null;
  /** @deprecated Developer problem badges are no longer displayed. */
  problemsCount?: number;
  /** @deprecated Developer problem badges are no longer displayed. */
  onOpenProblems?: () => void;
  // Layout toggles
  isSidebarOpen: boolean;
  onToggleSidebar: () => void;
  isBottomPanelOpen: boolean;
  onToggleBottomPanel: () => void;
  isRightViewOpen: boolean;
  onToggleRightView: () => void;
}

export const StatusBar: React.FC<StatusBarProps> = ({
  activeDoc,
  isSidebarOpen,
  onToggleSidebar,
  isBottomPanelOpen,
  onToggleBottomPanel,
  isRightViewOpen,
  onToggleRightView,
}) => {
  const getDocStatusText = () => {
    if (!activeDoc) return "No active document";
    switch (activeDoc.type) {
      case "word":
        return `Word: ${activeDoc.data.stats.words.toLocaleString()} words • ${activeDoc.data.stats.pages} pages`;
      case "excel": {
        const sheet =
          activeDoc.data.sheets.find(
            (s) => s.id === activeDoc.data.activeSheetId,
          ) || activeDoc.data.sheets[0];
        return `Excel: ${sheet?.name} • ${sheet?.columns.length} cols × ${sheet?.rowCount} rows`;
      }
      case "powerpoint": {
        const slide =
          activeDoc.data.slides.find(
            (s) => s.id === activeDoc.data.activeSlideId,
          ) || activeDoc.data.slides[0];
        return `PowerPoint: Slide ${slide?.slideNumber || 1} of ${activeDoc.data.slides.length}`;
      }
      case "code":
        return `${activeDoc.data.language}: ${activeDoc.data.content.split("\n").length} lines`;
    }
  };

  return (
    <div className="vscode-statusbar">
      {/* Left status items */}
      <div className="vscode-statusbar-left">
        <div className="vscode-statusbar-item" title="Local workspace">
          <GitBranch size={12} />
          <span>local workspace</span>
        </div>

        <div className="vscode-statusbar-item" style={{ color: "#e0f2fe" }}>
          <span>{getDocStatusText()}</span>
        </div>
      </div>

      {/* Right status items */}
      <div className="vscode-statusbar-right">
        {activeDoc?.type === "code" && (
          <div className="vscode-statusbar-item">
            <span>UTF-8</span>
          </div>
        )}

        {/* Layout quick toggles */}
        <button
          className="vscode-statusbar-item vscode-icon-btn"
          onClick={onToggleSidebar}
          title="Toggle Left Sidebar"
          aria-label="Toggle left sidebar"
          aria-pressed={isSidebarOpen}
        >
          <PanelLeft size={11} />
        </button>

        <button
          className="vscode-statusbar-item vscode-icon-btn"
          onClick={onToggleBottomPanel}
          title="Toggle Bottom Panel"
          aria-label="Toggle find and replace panel"
          aria-pressed={isBottomPanelOpen}
        >
          <PanelBottom size={11} />
        </button>

        <button
          className="vscode-statusbar-item vscode-icon-btn"
          onClick={onToggleRightView}
          title="Toggle Right View"
          aria-label="Toggle right inspector"
          aria-pressed={isRightViewOpen}
        >
          <PanelRight size={11} />
        </button>
      </div>
    </div>
  );
};
