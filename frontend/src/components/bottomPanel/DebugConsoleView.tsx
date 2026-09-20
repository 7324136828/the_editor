import React, { useState } from "react";
import { evaluateFormula } from "../../services/formulas";
export const DebugConsoleView: React.FC = () => {
  const [expression, setExpression] = useState(""),
    [output, setOutput] = useState<string[]>([
      "Safe formula console: arithmetic, parentheses, SUM, AVERAGE, MIN, MAX, COUNT. Use the workbook for cell references.",
    ]);
  return (
    <div
      style={{
        display: "flex",
        flexDirection: "column",
        height: "100%",
        gap: 8,
      }}
    >
      <pre
        style={{
          flex: 1,
          overflow: "auto",
          whiteSpace: "pre-wrap",
          font: "12px var(--vscode-mono-family)",
          userSelect: "text",
        }}
      >
        {output.join("\n")}
      </pre>
      <form
        style={{ display: "flex", gap: 8 }}
        onSubmit={(e) => {
          e.preventDefault();
          if (expression.trim())
            setOutput((old) => [
              ...old.slice(-100),
              `> ${expression}`,
              String(evaluateFormula(expression, () => null)),
            ]);
          setExpression("");
        }}
      >
        <input
          aria-label="Formula expression"
          className="vscode-input"
          style={{ flex: 1 }}
          value={expression}
          onChange={(e) => setExpression(e.target.value)}
          placeholder="=SUM(10,20)*2"
        />
        <button type="submit" className="vscode-btn-secondary">
          Evaluate
        </button>
      </form>
    </div>
  );
};
