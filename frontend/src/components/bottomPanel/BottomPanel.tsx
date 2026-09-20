import React from "react";
import { SearchOptions } from "../../services/search";
import {
  BottomPanelTab,
  SearchMatch,
  ProblemItem,
  OutputLogItem,
  DocumentType,
} from "../../types/vscode";
import { OfficeDocument } from "../../types/office";
import { FindInFilesView } from "./FindInFilesView";
import { Search, Maximize2, Minimize2, X } from "lucide-react";

interface BottomPanelProps {
  activeTab?: BottomPanelTab;
  onSelectTab?: (tab: BottomPanelTab) => void;
  onClose: () => void;
  height?: number;
  isMaximized?: boolean;
  onToggleMaximize?: () => void;
  documents: {
    id: string;
    name: string;
    type: DocumentType;
    doc: OfficeDocument;
  }[];
  /** @deprecated Developer panels are no longer displayed. */
  problems?: ProblemItem[];
  /** @deprecated Developer panels are no longer displayed. */
  outputLogs?: OutputLogItem[];
  onNavigateToMatch: (match: SearchMatch) => void;
  onReplaceAll?: (
    query: string,
    replaceWith: string,
    options?: SearchOptions,
  ) => void;
  /** @deprecated The workspace terminal is no longer displayed. */
  onOpenFile?: (fileName: string) => void;
  /** @deprecated The workspace terminal is no longer displayed. */
  onTriggerSlideshow?: () => void;
  /** @deprecated The workspace terminal is no longer displayed. */
  onToggleTheme?: (theme: "dark" | "light") => void;
}

export const BottomPanel: React.FC<BottomPanelProps> = ({
  onClose,
  height,
  isMaximized = false,
  onToggleMaximize,
  documents,
  onNavigateToMatch,
  onReplaceAll,
}) => (
  <section
    className="vscode-bottom-panel"
    aria-label="Find and replace"
    style={height === undefined ? undefined : { height }}
  >
    <div className="vscode-panel-header">
      <div className="vscode-panel-tabs">
        <div className="vscode-panel-tab active">
          <span style={{ display: "flex", alignItems: "center", gap: 6 }}>
            <Search size={13} />
            <span>FIND / REPLACE</span>
          </span>
        </div>
      </div>
      <div className="vscode-panel-actions">
        {onToggleMaximize && (
          <button
            className="vscode-icon-btn"
            onClick={onToggleMaximize}
            title={isMaximized ? "Restore panel height" : "Maximize panel"}
            aria-label={isMaximized ? "Restore panel height" : "Maximize panel"}
          >
            {isMaximized ? <Minimize2 size={13} /> : <Maximize2 size={13} />}
          </button>
        )}
        <button
          className="vscode-icon-btn"
          onClick={onClose}
          title="Close panel (Ctrl+J)"
          aria-label="Close find and replace panel"
        >
          <X size={13} />
        </button>
      </div>
    </div>
    <div className="vscode-panel-content">
      <FindInFilesView
        documents={documents}
        onNavigateToMatch={onNavigateToMatch}
        onReplaceAll={onReplaceAll}
      />
    </div>
  </section>
);
