import React, { useEffect, useRef, useState } from "react";
import {
  Command,
  Search,
  PanelLeft,
  PanelBottom,
  PanelRight,
  Maximize2,
} from "lucide-react";
import { BottomPanelTab } from "../../types/vscode";
interface TitleBarProps {
  title: string;
  onOpenCommandPalette: () => void;
  isSidebarOpen: boolean;
  onToggleSidebar: () => void;
  isBottomPanelOpen: boolean;
  onToggleBottomPanel: () => void;
  isRightViewOpen: boolean;
  onToggleRightView: () => void;
  onOpenFilePicker: () => void;
  onOpenSampleWord: () => void;
  onOpenSampleExcel: () => void;
  onOpenSamplePpt: () => void;
  onOpenBottomTab: (tab: BottomPanelTab) => void;
  onTriggerSlideshow: () => void;
  onToggleTheme: () => void;
  onNew?: () => void;
  onSave?: () => void;
  onSaveAll?: () => void;
  onSaveCopy?: () => void;
  onRename?: () => void;
  onExport?: () => void;
  onUndo?: () => void;
  onRedo?: () => void;
  onExtensions?: () => void;
  onPrint?: () => void;
  onCustomizeViews?: () => void;
  onResetLayout?: () => void;
}
export const TitleBar: React.FC<TitleBarProps> = (p) => {
  const [menu, setMenu] = useState<string | null>(null);
  const ref = useRef<HTMLDivElement>(null);
  useEffect(() => {
    const close = (e: MouseEvent) => {
      if (!ref.current?.contains(e.target as Node)) setMenu(null);
    };
    const esc = (e: KeyboardEvent) => {
      if (e.key === "Escape") setMenu(null);
    };
    window.addEventListener("mousedown", close);
    window.addEventListener("keydown", esc);
    return () => {
      window.removeEventListener("mousedown", close);
      window.removeEventListener("keydown", esc);
    };
  }, []);
  type Item = [string, (() => void) | undefined, string?];
  const menus: Record<string, Item[]> = {
    File: [
      ["New file…", p.onNew, "Ctrl+N"],
      ["Open from computer…", p.onOpenFilePicker, "Ctrl+O"],
      ["Save to workspace", p.onSave, "Ctrl+S"],
      ["Save all", p.onSaveAll, "Ctrl+Shift+S"],
      ["Save a copy…", p.onSaveCopy],
      ["Rename…", p.onRename],
      ["Download file", p.onExport],
      ["Print / Save PDF…", p.onPrint, "Ctrl+P"],
    ],
    Edit: [
      ["Undo", p.onUndo, "Ctrl+Z"],
      ["Redo", p.onRedo, "Ctrl+Y"],
      [
        "Find / Replace in files",
        () => p.onOpenBottomTab("search"),
        "Ctrl+Shift+F",
      ],
    ],
    View: [
      ["Command palette", p.onOpenCommandPalette, "Ctrl+Shift+P"],
      ["Toggle sidebar", p.onToggleSidebar, "Ctrl+B"],
      ["Toggle panel", p.onToggleBottomPanel, "Ctrl+J"],
      ["Toggle inspector", p.onToggleRightView, "Ctrl+Alt+B"],
      ["Customize views…", p.onCustomizeViews],
      ["Reset panel sizes", p.onResetLayout],
      ["Extensions & MCP", p.onExtensions],
      ["Switch color theme", p.onToggleTheme],
    ],
    Run: [["Present slides", p.onTriggerSlideshow, "F5"]],
    Help: [
      ["All commands & shortcuts", p.onOpenCommandPalette, "F1"],
      ["Plugin examples", p.onExtensions],
    ],
  };
  return (
    <header className="vscode-titlebar">
      <div className="vscode-titlebar-left">
        <div className="vscode-app-logo">
          <Command size={19} />
          <strong>Code Office</strong>
        </div>
        <div className="vscode-menubar" ref={ref}>
          {Object.entries(menus).map(([name, items]) => (
            <div key={name} style={{ position: "relative" }}>
              <button
                className={`vscode-menu-trigger ${menu === name ? "active" : ""}`}
                onClick={() => setMenu(menu === name ? null : name)}
                aria-expanded={menu === name}
              >
                {name}
              </button>
              {menu === name && (
                <div className="vscode-dropdown-menu" role="menu">
                  {items.map(([label, action, shortcut]) => (
                    <button
                      key={label}
                      role="menuitem"
                      className="vscode-menu-item"
                      disabled={!action}
                      onClick={() => {
                        setMenu(null);
                        action?.();
                      }}
                    >
                      <span>{label}</span>
                      <span className="vscode-menu-shortcut">{shortcut}</span>
                    </button>
                  ))}
                </div>
              )}
            </div>
          ))}
        </div>
      </div>
      <div className="vscode-titlebar-center">
        <button
          className="vscode-search-box-btn"
          onClick={p.onOpenCommandPalette}
          title="Command palette"
        >
          <Search size={13} />
          <span>{p.title}</span>
          <kbd>Ctrl Shift P</kbd>
        </button>
      </div>
      <div className="vscode-titlebar-right">
        <button
          className="vscode-icon-btn"
          title="Toggle primary sidebar"
          aria-label="Toggle primary sidebar"
          aria-pressed={p.isSidebarOpen}
          onClick={p.onToggleSidebar}
        >
          <PanelLeft size={15} />
        </button>
        <button
          className="vscode-icon-btn"
          title="Toggle bottom panel"
          aria-label="Toggle bottom panel"
          aria-pressed={p.isBottomPanelOpen}
          onClick={p.onToggleBottomPanel}
        >
          <PanelBottom size={15} />
        </button>
        <button
          className="vscode-icon-btn"
          title="Toggle inspector"
          aria-label="Toggle inspector"
          aria-pressed={p.isRightViewOpen}
          onClick={p.onToggleRightView}
        >
          <PanelRight size={15} />
        </button>
        <button
          className="vscode-icon-btn"
          title="Toggle fullscreen"
          onClick={() => {
            if (document.fullscreenElement) void document.exitFullscreen();
            else
              void document.documentElement.requestFullscreen().catch(() => {});
          }}
        >
          <Maximize2 size={14} />
        </button>
      </div>
    </header>
  );
};
