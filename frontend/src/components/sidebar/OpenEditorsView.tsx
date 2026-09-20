import React from "react";
import { EditorTab } from "../../types/vscode";
import { FileText, Table, Presentation, FileCode, X } from "lucide-react";

interface OpenEditorsViewProps {
  tabs: EditorTab[];
  activeTabId: string;
  onSelectTab: (tabId: string) => void;
  onCloseTab: (tabId: string) => void;
}

export const OpenEditorsView: React.FC<OpenEditorsViewProps> = ({
  tabs,
  activeTabId,
  onSelectTab,
  onCloseTab,
}) => {
  const getIcon = (type: string) => {
    switch (type) {
      case "word":
        return <FileText size={13} color="#4175c5" />;
      case "excel":
        return <Table size={13} color="#107c41" />;
      case "powerpoint":
        return <Presentation size={13} color="#c43e1c" />;
      default:
        return <FileCode size={13} color="#007acc" />;
    }
  };

  return (
    <div style={{ padding: "2px 0" }}>
      {tabs.map((tab) => {
        const isActive = tab.id === activeTabId;
        return (
          <div
            key={tab.id}
            className={`vscode-tree-item ${isActive ? "active" : ""}`}
            onClick={() => onSelectTab(tab.id)}
            style={{ justifyContent: "space-between" }}
          >
            <div
              style={{
                display: "flex",
                alignItems: "center",
                gap: 6,
                minWidth: 0,
              }}
            >
              {getIcon(tab.docType)}
              <span style={{ overflow: "hidden", textOverflow: "ellipsis" }}>
                {tab.title}
              </span>
            </div>

            <button
              className="vscode-icon-btn vscode-tab-close"
              style={{ width: 18, height: 18 }}
              onClick={(e) => {
                e.stopPropagation();
                onCloseTab(tab.id);
              }}
              title="Close editor"
            >
              <X size={12} />
            </button>
          </div>
        );
      })}
    </div>
  );
};
