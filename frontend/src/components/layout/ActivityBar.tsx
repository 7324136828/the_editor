import React from "react";
import { ActivityBarTab, DocumentType } from "../../types/vscode";
import {
  Files,
  History,
  Blocks,
  Settings,
  FileText,
  Table2,
  Presentation,
  FileCode2,
} from "lucide-react";
interface ActivityBarProps {
  activeTab: ActivityBarTab;
  onSelectTab: (tab: ActivityBarTab) => void;
  searchBadgeCount?: number;
  documentType?: DocumentType;
}
export const ActivityBar: React.FC<ActivityBarProps> = ({
  activeTab,
  onSelectTab,
  documentType,
}) => {
  const Icon =
    documentType === "word"
      ? FileText
      : documentType === "excel"
        ? Table2
        : documentType === "powerpoint"
          ? Presentation
          : documentType === "code"
            ? FileCode2
            : Files;
  const labels = {
    word: "Document outline",
    excel: "Workbook explorer",
    powerpoint: "Slide navigator",
    code: "Code explorer",
  };
  const items = [
    {
      id: "explorer",
      icon: Icon,
      label: documentType ? labels[documentType] : "Explorer",
    },
    { id: "source-control", icon: History, label: "Saved version history" },
    {
      id: "extensions",
      icon: Blocks,
      label: "Extensions & MCP (Ctrl+Shift+X)",
    },
  ] as const;
  return (
    <nav className="vscode-activitybar" aria-label="Activity bar">
      <div>
        {items.map(({ id, icon: Item, label }) => (
          <button
            key={id}
            aria-label={label}
            title={label}
            aria-pressed={activeTab === id}
            className={`vscode-activitybar-item ${activeTab === id ? "active" : ""}`}
            onClick={() => onSelectTab(id)}
            data-document-type={id === "explorer" ? documentType : undefined}
          >
            <Item
              size={23}
              color={
                id === "explorer"
                  ? ({
                      word: "#6ca1ee",
                      excel: "#59c78b",
                      powerpoint: "#ed926d",
                      code: "#d9c872",
                    }[documentType!] ?? undefined)
                  : undefined
              }
            />
          </button>
        ))}
      </div>
      <button
        className={`vscode-activitybar-item ${activeTab === "settings" ? "active" : ""}`}
        title="Settings"
        aria-label="Settings"
        onClick={() => onSelectTab("settings")}
      >
        <Settings size={22} />
      </button>
    </nav>
  );
};
