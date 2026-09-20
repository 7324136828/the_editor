import React, { useState } from "react";
import { OfficeDocument } from "../../types/office";
import { evaluateFormula } from "../../services/formulas";
import { collectDocumentContext } from "../../plugins";
interface TerminalViewProps {
  documents?: { name: string; doc: OfficeDocument }[];
  onOpenFile?: (name: string) => void;
  onTriggerSlideshow?: () => void;
  onToggleTheme?: (theme: "dark" | "light") => void;
}
export const TerminalView: React.FC<TerminalViewProps> = ({
  documents = [],
  onOpenFile,
  onTriggerSlideshow,
  onToggleTheme,
}) => {
  const [input, setInput] = useState(""),
    [lines, setLines] = useState([
      "Code Office workspace terminal. Type help for commands. This is a document command console, not an operating-system shell.",
    ]);
  const run = () => {
    const command = input.trim();
    if (!command) return;
    const [root, ...parts] = command.split(" "),
      arg = parts.join(" ");
    let result = "";
    switch (root.toLowerCase()) {
      case "help":
        result =
          "ls / dir — list workspace files\nopen <name> — open a file\nstats — workspace document statistics\ncalc <expression> — safe arithmetic and SUM, AVERAGE, MIN, MAX, COUNT\ntheme dark | light — change theme\npresent — start the active presentation\nclear — clear output";
        break;
      case "ls":
      case "dir":
        result = documents.map((d) => d.name).join("\n") || "No files";
        break;
      case "open": {
        const match = documents.find(
          (d) => arg && d.name.toLowerCase().includes(arg.toLowerCase()),
        );
        if (match) {
          onOpenFile?.(match.name);
          result = `Opened ${match.name}`;
        } else result = "File not found. Use ls to see available files.";
        break;
      }
      case "stats":
        result = documents
          .map((d) => `${d.name}: ${collectDocumentContext(d.doc).summary}`)
          .join("\n");
        break;
      case "calc":
        result = String(evaluateFormula(arg, () => null));
        break;
      case "theme":
        if (arg === "dark" || arg === "light") {
          onToggleTheme?.(arg);
          result = `Theme: ${arg}`;
        } else result = "Usage: theme dark | light";
        break;
      case "present":
      case "slideshow":
        onTriggerSlideshow?.();
        result = "Presentation requested for the active editor.";
        break;
      case "clear":
      case "cls":
        setLines([]);
        setInput("");
        return;
      default:
        result = `Unknown command: ${root}. Type help.`;
    }
    setLines((old) => [...old.slice(-100), `> ${command}`, result]);
    setInput("");
  };
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
      <pre
        style={{
          flex: 1,
          overflow: "auto",
          whiteSpace: "pre-wrap",
          userSelect: "text",
        }}
      >
        {lines.join("\n")}
      </pre>
      <form
        style={{ display: "flex", gap: 8, paddingTop: 8 }}
        onSubmit={(e) => {
          e.preventDefault();
          run();
        }}
      >
        <span>workspace&gt;</span>
        <input
          aria-label="Workspace command"
          className="vscode-input"
          style={{ flex: 1 }}
          value={input}
          onChange={(e) => setInput(e.target.value)}
          placeholder="help, ls, open, calc…"
        />
        <button className="vscode-btn-secondary" type="submit">
          Run
        </button>
      </form>
    </div>
  );
};
