import React, { useEffect, useState } from "react";
import {
  Blocks,
  Check,
  Download,
  FileJson,
  Play,
  Radio,
  RefreshCw,
} from "lucide-react";
import type { OfficeDocument } from "../../types/office";
import { collectDocumentContext, defaultPluginHost } from "../../plugins";
import type { DocumentContext, PluginHost } from "../../plugins";
import "./plugins.css";

interface ExtensionsPanelProps {
  activeDoc: OfficeDocument | null;
  onChangeDoc: (document: OfficeDocument) => void;
  onNotify: (message: string) => void;
  host?: PluginHost;
}

export function ExtensionsPanel({
  activeDoc,
  onChangeDoc,
  onNotify,
  host = defaultPluginHost,
}: ExtensionsPanelProps) {
  const [, update] = useState(0);
  const [filter, setFilter] = useState("");
  const [result, setResult] = useState("");
  const [context, setContext] = useState<DocumentContext | null>(null);
  const [connection, setConnection] = useState<
    "idle" | "checking" | "connected" | "error"
  >("idle");
  const [mcpInfo, setMcpInfo] = useState("");

  useEffect(() => host.subscribe(() => update((value) => value + 1)), [host]);
  useEffect(() => {
    setContext(null);
    setResult("");
  }, [activeDoc?.data.id]);

  const execute = (id: string) => {
    if (!activeDoc) return;
    try {
      const output = host.execute(id, activeDoc);
      if (output.document) onChangeDoc(output.document);
      if (output.context) setContext(output.context);
      setResult(output.message);
      onNotify(output.message);
    } catch (error) {
      const message =
        error instanceof Error ? error.message : "Plugin command failed.";
      setResult(message);
      onNotify(message);
    }
  };

  const downloadContext = () => {
    if (!activeDoc) return;
    const snapshot = collectDocumentContext(activeDoc);
    setContext(snapshot);
    const url = URL.createObjectURL(
      new Blob([JSON.stringify(snapshot, null, 2)], {
        type: "application/json",
      }),
    );
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = `${activeDoc.data.title.replace(/\.[^.]+$/, "")}.context.json`;
    anchor.click();
    window.setTimeout(() => URL.revokeObjectURL(url), 1000);
    onNotify("Downloaded context from the current document.");
  };

  const checkMcp = async () => {
    setConnection("checking");
    const controller = new AbortController();
    const timer = window.setTimeout(() => controller.abort(), 8000);
    const request = async (id: number, method: string, params?: unknown) => {
      const response = await fetch("/mcp", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Accept: "application/json, text/event-stream",
          "MCP-Protocol-Version": "2025-11-25",
        },
        body: JSON.stringify({ jsonrpc: "2.0", id, method, params }),
        signal: controller.signal,
      });
      if (!response.ok)
        throw new Error(`Local MCP server returned HTTP ${response.status}.`);
      const payload = await response.json();
      if (payload.error) throw new Error(payload.error.message);
      return payload.result;
    };
    try {
      const server = await request(1, "initialize", {
        protocolVersion: "2025-11-25",
        capabilities: {},
        clientInfo: { name: "office-studio-panel", version: "1.0.0" },
      });
      const initialized = await fetch("/mcp", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Accept: "application/json, text/event-stream",
          "MCP-Protocol-Version": server.protocolVersion,
        },
        body: JSON.stringify({
          jsonrpc: "2.0",
          method: "notifications/initialized",
        }),
        signal: controller.signal,
      });
      if (!initialized.ok)
        throw new Error("MCP initialization was not accepted.");
      const tools = await request(2, "tools/list");
      const resources = await request(3, "resources/list");
      setConnection("connected");
      setMcpInfo(
        `${server.serverInfo.name} · ${tools.tools.length} read-only tools · ${resources.resources.length} saved documents`,
      );
    } catch (error) {
      setConnection("error");
      setMcpInfo(
        error instanceof Error && error.name === "AbortError"
          ? "Connection timed out. Start the local server with npm run dev."
          : error instanceof Error
            ? `${error.message} Start the local server with npm run dev.`
            : "Cannot connect to the local MCP server.",
      );
    } finally {
      window.clearTimeout(timer);
    }
  };

  const plugins = host
    .list()
    .filter((plugin) =>
      `${plugin.name} ${plugin.description}`
        .toLowerCase()
        .includes(filter.toLowerCase()),
    );
  return (
    <div className="studio-extensions">
      <div className="studio-extension-search">
        <input
          className="vscode-input"
          aria-label="Search extensions"
          placeholder="Search installed extensions"
          value={filter}
          onChange={(event) => setFilter(event.target.value)}
        />
      </div>
      <p className="studio-extension-caption">
        LOCAL EXTENSIONS <span>{host.list().length}</span>
      </p>
      {plugins.map((plugin) => (
        <section key={plugin.id} className="studio-extension-card">
          <div className="studio-extension-title">
            <Blocks size={21} />
            <div>
              <strong>{plugin.name}</strong>
              <small>Built-in · v{plugin.version}</small>
            </div>
          </div>
          <p>{plugin.description}</p>
          <div className="studio-extension-controls">
            <span>
              {plugin.permissions.includes("document:write")
                ? "Reads and edits current document"
                : "Reads current document"}
            </span>
            <button
              className="vscode-btn-secondary"
              onClick={() => {
                host.setEnabled(plugin.id, !plugin.enabled);
                if (plugin.enabled && plugin.id === "studio.context")
                  setContext(null);
              }}
              aria-label={`${plugin.enabled ? "Disable" : "Enable"} ${plugin.name}`}
            >
              {plugin.enabled ? (
                <>
                  <Check size={12} /> Enabled
                </>
              ) : (
                "Enable"
              )}
            </button>
          </div>
          {plugin.enabled && (
            <div className="studio-extension-commands">
              {plugin.commands.map((command) => {
                const supported =
                  activeDoc !== null &&
                  command.documentTypes.includes(activeDoc.type);
                return (
                  <button
                    className="vscode-btn-secondary"
                    key={command.id}
                    disabled={!supported}
                    title={
                      supported
                        ? command.title
                        : `Open a ${command.documentTypes.join(" / ")} document`
                    }
                    onClick={() => execute(command.id)}
                  >
                    <Play size={11} />
                    {command.title}
                  </button>
                );
              })}
            </div>
          )}
        </section>
      ))}
      {plugins.length === 0 && (
        <p className="studio-extension-empty">
          No installed extensions match your search.
        </p>
      )}
      {result && (
        <div className="studio-extension-result" role="status">
          {result}
        </div>
      )}
      {context && (
        <section className="studio-extension-card">
          <div className="studio-extension-title">
            <FileJson size={17} />
            <strong>Current context</strong>
          </div>
          <p>{context.summary}</p>
          <p className="studio-extension-hint">
            Snapshot of {context.document.title}.{" "}
            {context.truncated
              ? "Text is limited to 12,000 characters."
              : "Contains the current editor text."}
          </p>
          <details>
            <summary>Inspect JSON</summary>
            <pre>{JSON.stringify(context, null, 2)}</pre>
          </details>
          <button
            className="vscode-btn-secondary"
            onClick={downloadContext}
            disabled={!activeDoc}
          >
            <Download size={12} /> Download fresh snapshot
          </button>
        </section>
      )}
      <section className="studio-extension-card studio-mcp-card">
        <div className="studio-extension-title">
          <Radio size={20} />
          <strong>Workspace MCP</strong>
        </div>
        <p>
          Connect a compatible client to the local workspace. Tools and
          resources expose saved documents; save edits before reading them
          through MCP.
        </p>
        <code>http://127.0.0.1:3001/mcp</code>
        <div className="studio-extension-controls">
          <span className={`studio-mcp-${connection}`}>
            {connection === "connected"
              ? "Connected"
              : connection === "checking"
                ? "Connecting…"
                : connection === "error"
                  ? "Unavailable"
                  : "Not checked"}
          </span>
          <button
            className="vscode-btn-secondary"
            disabled={connection === "checking"}
            onClick={checkMcp}
          >
            <RefreshCw size={12} /> Test connection
          </button>
        </div>
        {mcpInfo && <p role="status">{mcpInfo}</p>}
        <p className="studio-extension-hint">
          Read-only · localhost · no AI model required. JavaScript and Python
          examples are in examples/plugins.
        </p>
      </section>
      <p className="studio-extension-hint studio-extension-footer">
        Extensions are trusted, source-controlled modules. Add plugins through
        the host API; arbitrary uploaded scripts are not executed. See
        docs/plugins.md.
      </p>
    </div>
  );
}
