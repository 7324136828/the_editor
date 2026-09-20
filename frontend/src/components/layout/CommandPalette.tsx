import React, { useState, useEffect, useRef } from "react";
import { CommandItem } from "../../types/vscode";
import { Search } from "lucide-react";

interface CommandPaletteProps {
  isOpen: boolean;
  onClose: () => void;
  commands: CommandItem[];
}

export const CommandPalette: React.FC<CommandPaletteProps> = ({
  isOpen,
  onClose,
  commands,
}) => {
  const [search, setSearch] = useState("");
  const [selectedIdx, setSelectedIdx] = useState(0);
  const inputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    if (isOpen) {
      setSearch("");
      setSelectedIdx(0);
      setTimeout(() => inputRef.current?.focus(), 50);
    }
  }, [isOpen]);

  const filtered = commands.filter(
    (c) =>
      c.title.toLowerCase().includes(search.toLowerCase()) ||
      (c.category && c.category.toLowerCase().includes(search.toLowerCase())),
  );

  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (!isOpen) return;
      if (e.key === "Escape") {
        onClose();
      } else if (e.key === "ArrowDown") {
        e.preventDefault();
        setSelectedIdx((prev) => (prev + 1) % Math.max(1, filtered.length));
      } else if (e.key === "ArrowUp") {
        e.preventDefault();
        setSelectedIdx(
          (prev) => (prev - 1 + filtered.length) % Math.max(1, filtered.length),
        );
      } else if (e.key === "Enter") {
        e.preventDefault();
        if (filtered[selectedIdx]) {
          filtered[selectedIdx].action();
          onClose();
        }
      }
    };

    window.addEventListener("keydown", handleKeyDown);
    return () => window.removeEventListener("keydown", handleKeyDown);
  }, [isOpen, filtered, selectedIdx, onClose]);

  if (!isOpen) return null;

  return (
    <div className="vscode-modal-overlay" onClick={onClose}>
      <div
        className="vscode-command-palette"
        onClick={(e) => e.stopPropagation()}
      >
        <div
          style={{
            display: "flex",
            alignItems: "center",
            gap: 8,
            padding: "8px 12px",
            borderBottom: "1px solid #333",
          }}
        >
          <Search size={14} color="#007acc" />
          <input
            ref={inputRef}
            className="vscode-input"
            style={{
              flex: 1,
              border: "none",
              background: "transparent",
              outline: "none",
              fontSize: 13,
              color: "var(--vscode-tab-active-fg)",
            }}
            placeholder="Type a command or search files..."
            value={search}
            onChange={(e) => {
              setSearch(e.target.value);
              setSelectedIdx(0);
            }}
          />
        </div>

        <div style={{ maxHeight: 320, overflowY: "auto", padding: "4px 0" }}>
          {filtered.length === 0 ? (
            <div
              style={{
                padding: "12px 16px",
                color: "var(--vscode-tab-inactive-fg)",
                fontSize: 12,
              }}
            >
              No matching commands found.
            </div>
          ) : (
            filtered.map((cmd, idx) => {
              const isSelected = idx === selectedIdx;
              return (
                <div
                  key={cmd.id}
                  className={`vscode-tree-item ${isSelected ? "active" : ""}`}
                  style={{
                    padding: "8px 14px",
                    justifyContent: "space-between",
                  }}
                  onClick={() => {
                    cmd.action();
                    onClose();
                  }}
                  onMouseEnter={() => setSelectedIdx(idx)}
                >
                  <div
                    style={{ display: "flex", alignItems: "center", gap: 6 }}
                  >
                    {cmd.category && (
                      <span
                        style={{
                          color: "var(--vscode-tab-inactive-fg)",
                          fontSize: 11,
                        }}
                      >
                        {cmd.category}:
                      </span>
                    )}
                    <span
                      style={{
                        color: "var(--vscode-tab-active-fg)",
                        fontWeight: isSelected ? 600 : "normal",
                      }}
                    >
                      {cmd.title}
                    </span>
                  </div>

                  {cmd.shortcut && (
                    <span className="vscode-menu-shortcut">{cmd.shortcut}</span>
                  )}
                </div>
              );
            })
          )}
        </div>
      </div>
    </div>
  );
};
