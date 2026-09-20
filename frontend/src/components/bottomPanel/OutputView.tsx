import React, { useState } from "react";
import { OutputLogItem } from "../../types/vscode";

interface OutputViewProps {
  logs: OutputLogItem[];
}

export const OutputView: React.FC<OutputViewProps> = ({ logs }) => {
  const [activeChannel, setActiveChannel] = useState<string>("Office Studio");

  const filteredLogs = logs.filter(
    (l) => activeChannel === "All" || l.channel === activeChannel,
  );

  return (
    <div
      style={{
        height: "100%",
        display: "flex",
        flexDirection: "column",
        fontFamily: "var(--vscode-mono-family)",
        fontSize: 12,
      }}
    >
      {/* Channel selector toolbar */}
      <div
        style={{
          display: "flex",
          alignItems: "center",
          gap: 12,
          paddingBottom: 6,
          borderBottom: "1px solid #2a2a2a",
          marginBottom: 6,
        }}
      >
        <span style={{ color: "var(--vscode-tab-inactive-fg)", fontSize: 11 }}>
          OUTPUT CHANNEL:
        </span>
        <select
          className="vscode-input"
          style={{ height: 22, fontSize: 11 }}
          value={activeChannel}
          onChange={(e) => setActiveChannel(e.target.value)}
        >
          <option value="All">All Channels</option>
          <option value="Office Studio">Office Studio</option>
          <option value="Parser">Parser</option>
          <option value="Formulas">Formulas</option>
          <option value="Build">Build</option>
        </select>
      </div>

      {/* Log lines */}
      <div
        style={{
          flex: 1,
          overflowY: "auto",
          display: "flex",
          flexDirection: "column",
          gap: 3,
        }}
      >
        {filteredLogs.map((log) => (
          <div
            key={log.id}
            style={{ display: "flex", gap: 8, lineHeight: 1.4 }}
          >
            <span style={{ color: "#6e7681", minWidth: 70 }}>
              [{log.timestamp}]
            </span>
            <span
              style={{
                color:
                  log.channel === "Formulas"
                    ? "#107c41"
                    : log.channel === "Parser"
                      ? "#38bdf8"
                      : "var(--vscode-editor-fg)",
                minWidth: 100,
                fontWeight: 600,
              }}
            >
              [{log.channel}]
            </span>
            <span
              style={{
                color:
                  log.level === "error"
                    ? "#f87171"
                    : log.level === "warn"
                      ? "#fbbf24"
                      : "var(--vscode-sidebar-fg)",
              }}
            >
              {log.message}
            </span>
          </div>
        ))}
      </div>
    </div>
  );
};
