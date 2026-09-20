import React from "react";
import { EditorTab } from "../../types/vscode";
import { OfficeDocument } from "../../types/office";
import { WordEditor } from "./WordEditor";
import { ExcelEditor } from "./ExcelEditor";
import { PptEditor } from "./PptEditor";
import { CodeEditor } from "./CodeEditor";
import {
  FileText,
  Table,
  Presentation,
  FileCode,
  X,
  ChevronRight,
} from "lucide-react";

interface EditorContainerProps {
  tabs: EditorTab[];
  activeTabId: string;
  activeDoc: OfficeDocument | null;
  onSelectTab: (tabId: string) => void;
  onCloseTab: (tabId: string) => void;
  onChangeActiveDoc?: (doc: OfficeDocument) => void;
  // Specific navigation props
  targetWordHeadingId?: string | null;
  targetExcelCoord?: string | null;
  selectedPptObjectId?: string | null;
  onSelectPptObject?: (id: string) => void;
  onSelectExcelCell?: (coord: string) => void;
}

export const EditorContainer: React.FC<EditorContainerProps> = ({
  tabs,
  activeTabId,
  activeDoc,
  onSelectTab,
  onCloseTab,
  onChangeActiveDoc,
  targetWordHeadingId,
  targetExcelCoord,
  selectedPptObjectId,
  onSelectPptObject,
  onSelectExcelCell,
}) => {
  const getTabIcon = (type: string) => {
    switch (type) {
      case "word":
        return <FileText size={14} color="#4175c5" />;
      case "excel":
        return <Table size={14} color="#107c41" />;
      case "powerpoint":
        return <Presentation size={14} color="#c43e1c" />;
      default:
        return <FileCode size={14} color="#007acc" />;
    }
  };

  const activeTab = tabs.find((t) => t.id === activeTabId);

  return (
    <div className="vscode-editor-workspace">
      {/* Tabs Header Bar */}
      <div
        className="vscode-tabs-bar"
        role="tablist"
        aria-label="Open documents"
      >
        {tabs.map((tab) => {
          const isActive = tab.id === activeTabId;
          return (
            <div
              key={tab.id}
              role="tab"
              tabIndex={isActive ? 0 : -1}
              aria-selected={isActive}
              onKeyDown={(e) => {
                const i = tabs.findIndex((t) => t.id === tab.id);
                if (e.key === "ArrowRight" || e.key === "ArrowLeft") {
                  e.preventDefault();
                  onSelectTab(
                    tabs[
                      (i + (e.key === "ArrowRight" ? 1 : tabs.length - 1)) %
                        tabs.length
                    ].id,
                  );
                }
                if (e.key === "Delete") {
                  e.preventDefault();
                  onCloseTab(tab.id);
                }
              }}
              className={`vscode-tab ${tab.docType} ${isActive ? "active" : ""}`}
              onClick={() => onSelectTab(tab.id)}
            >
              {getTabIcon(tab.docType)}
              <span>{tab.title}</span>
              {tab.isDirty && (
                <div
                  style={{
                    width: 6,
                    height: 6,
                    borderRadius: "50%",
                    backgroundColor: "#ffffff",
                  }}
                />
              )}
              <div
                className="vscode-tab-close"
                onClick={(e) => {
                  e.stopPropagation();
                  onCloseTab(tab.id);
                }}
                title="Close"
              >
                <X size={12} />
              </div>
            </div>
          );
        })}
      </div>

      {/* Breadcrumb Navigation Bar */}
      {activeTab && (
        <div className="vscode-breadcrumb-bar">
          <span>workspace</span>
          <ChevronRight size={12} />
          <span>documents</span>
          <ChevronRight size={12} />
          <span
            style={{
              color: "#cccccc",
              display: "flex",
              alignItems: "center",
              gap: 4,
            }}
          >
            {getTabIcon(activeTab.docType)}
            {activeTab.title}
          </span>
        </div>
      )}

      {/* Main Canvas Viewport */}
      <div className="vscode-editor-canvas">
        {!activeDoc ? (
          <div
            style={{
              height: "100%",
              display: "flex",
              flexDirection: "column",
              alignItems: "center",
              justifyContent: "center",
              color: "#858585",
              gap: 12,
            }}
          >
            <div style={{ fontSize: 24, fontWeight: 300, color: "#d4d4d4" }}>
              VS Code Office Studio
            </div>
            <div style={{ fontSize: 13 }}>
              Select a file from the left sidebar or drag-and-drop a Word,
              Excel, or PowerPoint document.
            </div>
          </div>
        ) : activeDoc.type === "word" ? (
          <WordEditor
            key={activeDoc.data.id}
            document={activeDoc.data}
            onChangeDocument={(doc) =>
              onChangeActiveDoc?.({ type: "word", data: doc })
            }
            targetHeadingId={targetWordHeadingId}
          />
        ) : activeDoc.type === "excel" ? (
          <ExcelEditor
            key={activeDoc.data.id}
            onSelectCell={onSelectExcelCell}
            document={activeDoc.data}
            onChangeDocument={(doc) =>
              onChangeActiveDoc?.({ type: "excel", data: doc })
            }
            selectedCoord={targetExcelCoord}
          />
        ) : activeDoc.type === "powerpoint" ? (
          <PptEditor
            key={activeDoc.data.id}
            document={activeDoc.data}
            onChangeDocument={(doc) =>
              onChangeActiveDoc?.({ type: "powerpoint", data: doc })
            }
            selectedObjectId={selectedPptObjectId}
            onSelectObject={onSelectPptObject}
          />
        ) : (
          <CodeEditor
            key={activeDoc.data.id}
            document={activeDoc.data}
            onChangeDocument={(doc) =>
              onChangeActiveDoc?.({ type: "code", data: doc })
            }
          />
        )}
      </div>
    </div>
  );
};
