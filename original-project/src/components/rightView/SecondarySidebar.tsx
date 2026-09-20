import React, { useEffect, useState } from "react";
import { RightViewTab } from "../../types/vscode";
import { OfficeDocument } from "../../types/office";
import { PropertyInspector } from "./PropertyInspector";
import { CopilotChatView } from "./CopilotChatView";
import { MinimapView } from "./MinimapView";
import { FileDetailsView } from "./FileDetailsView";
import {
  Sliders,
  Sparkles,
  Map,
  Info,
  X,
  Settings2,
  StickyNote,
} from "lucide-react";
import type { ViewPreferencesController } from "../../views/useViewPreferences";
import { builtinViews } from "../../views/preferences";
import { CustomViewContent } from "../../views/CustomViewContent";

interface SecondarySidebarProps {
  activeTab: RightViewTab;
  onSelectTab: (tab: RightViewTab) => void;
  onClose: () => void;
  activeDoc: OfficeDocument | null;
  onChangeDoc?: (doc: OfficeDocument) => void;
  selectedPptObjectId?: string | null;
  selectedExcelCoord?: string;
  viewPreferences?: ViewPreferencesController;
  onCustomizeViews?: () => void;
}
const icons: Record<string, React.ReactNode> = {
  inspector: <Sliders size={12} />,
  copilot: <Sparkles size={12} />,
  minimap: <Map size={12} />,
  metadata: <Info size={12} />,
};
export const SecondarySidebar: React.FC<SecondarySidebarProps> = ({
  activeTab,
  onSelectTab,
  onClose,
  activeDoc,
  onChangeDoc,
  selectedPptObjectId,
  selectedExcelCoord,
  viewPreferences,
  onCustomizeViews,
}) => {
  const [customTab, setCustomTab] = useState<string | null>(null);
  useEffect(() => setCustomTab(null), [activeTab, viewPreferences?.profile]);
  const definitions =
    viewPreferences?.availableViews("right") ??
    builtinViews(activeDoc?.type ?? "code", "right");
  const visible =
    viewPreferences?.preferences.right ?? definitions.map((view) => view.id);
  const desired = customTab ?? activeTab;
  const selected = visible.includes(desired) ? desired : visible[0];
  const custom = viewPreferences?.preferences.custom.find(
    (view) => view.id === selected,
  );
  const selectView = (id: string) => {
    if (id.startsWith("custom:")) setCustomTab(id);
    else {
      setCustomTab(null);
      onSelectTab(id as RightViewTab);
    }
  };
  return (
    <div className="vscode-right-view">
      <div className="vscode-right-view-header">
        <div
          className="vscode-right-view-tabs"
          role="tablist"
          aria-label="Right sidebar views"
        >
          {visible.map((id) => {
            const view = definitions.find((definition) => definition.id === id);
            if (!view) return null;
            return (
              <div
                key={id}
                className={`vscode-right-view-tab ${selected === id ? "active" : ""}`}
              >
                <button
                  className="view-tab-select"
                  role="tab"
                  aria-label={view.title}
                  aria-selected={selected === id}
                  aria-controls="right-view-content"
                  title={view.description}
                  onClick={() => selectView(id)}
                >
                  {icons[id] ?? <StickyNote size={12} />}
                  <span>{view.title.toUpperCase()}</span>
                </button>
                {viewPreferences && (
                  <button
                    className="view-tab-hide"
                    aria-label={`Hide ${view.title}`}
                    title={`Hide ${view.title}`}
                    onClick={() =>
                      viewPreferences.setVisible("right", id, false)
                    }
                  >
                    <X size={11} />
                  </button>
                )}
              </div>
            );
          })}
        </div>
        <div className="view-header-actions">
          {onCustomizeViews && (
            <button
              className="vscode-icon-btn"
              onClick={onCustomizeViews}
              title="Customize views"
              aria-label="Customize right sidebar views"
            >
              <Settings2 size={14} />
            </button>
          )}
          <button
            className="vscode-icon-btn"
            onClick={onClose}
            title="Close Right View (Ctrl+Alt+B)"
            aria-label="Close right sidebar"
          >
            <X size={13} />
          </button>
        </div>
      </div>
      <div
        className="vscode-right-view-body"
        id="right-view-content"
        role="tabpanel"
        aria-label={
          definitions.find((view) => view.id === selected)?.title ??
          "Right sidebar"
        }
      >
        {selected === "inspector" && (
          <PropertyInspector
            activeDoc={activeDoc}
            onChangeDoc={onChangeDoc}
            selectedPptObjectId={selectedPptObjectId}
            selectedExcelCoord={selectedExcelCoord}
          />
        )}
        {selected === "copilot" && <CopilotChatView activeDoc={activeDoc} />}
        {selected === "minimap" && <MinimapView activeDoc={activeDoc} />}
        {selected === "metadata" && <FileDetailsView activeDoc={activeDoc} />}
        {custom && (
          <CustomViewContent
            view={custom}
            activeDoc={activeDoc}
            onChangeNotes={(notes) =>
              viewPreferences?.updateNotes(custom.id, notes)
            }
            storageError={viewPreferences?.storageError}
          />
        )}
        {!visible.length && (
          <div className="views-empty">
            <p>No right sidebar views are shown.</p>
            <button className="vscode-btn-secondary" onClick={onCustomizeViews}>
              Add a view
            </button>
          </div>
        )}
      </div>
    </div>
  );
};
