import React from "react";
import { OfficeDocument } from "../../types/office";
import {
  Info,
  FileCheck,
  HardDrive,
  Calendar,
  ShieldCheck,
} from "lucide-react";

interface FileDetailsViewProps {
  activeDoc: OfficeDocument | null;
}

export const FileDetailsView: React.FC<FileDetailsViewProps> = ({
  activeDoc,
}) => {
  if (!activeDoc) {
    return (
      <div
        style={{
          color: "var(--vscode-tab-inactive-fg)",
          fontSize: 12,
          padding: 12,
        }}
      >
        No document active to view details.
      </div>
    );
  }

  return (
    <div
      style={{
        display: "flex",
        flexDirection: "column",
        gap: 12,
        fontSize: 12,
        color: "var(--vscode-sidebar-fg)",
      }}
    >
      <div
        style={{
          display: "flex",
          alignItems: "center",
          gap: 6,
          color: "var(--vscode-tab-inactive-fg)",
          fontSize: 10,
        }}
      >
        <Info size={12} />
        <span>FILE ATTRIBUTES & SPECIFICATION</span>
      </div>

      <div
        style={{
          display: "flex",
          flexDirection: "column",
          gap: 8,
          backgroundColor: "var(--vscode-sidebar-bg)",
          padding: 12,
          borderRadius: 6,
          border: "1px solid #3c3c3c",
        }}
      >
        <div>
          <span
            style={{ fontSize: 10, color: "var(--vscode-tab-inactive-fg)" }}
          >
            FILE NAME:
          </span>
          <div
            style={{ fontWeight: 600, color: "var(--vscode-tab-active-fg)" }}
          >
            {activeDoc.data.title}
          </div>
        </div>

        <div>
          <span
            style={{ fontSize: 10, color: "var(--vscode-tab-inactive-fg)" }}
          >
            FORMAT SCHEMA:
          </span>
          <div style={{ color: "#38bdf8" }}>
            {activeDoc.type === "word"
              ? "Microsoft Word OpenXML (.docx)"
              : activeDoc.type === "excel"
                ? "Microsoft Excel Spreadsheet OpenXML (.xlsx)"
                : activeDoc.type === "powerpoint"
                  ? "Microsoft PowerPoint Presentation OpenXML (.pptx)"
                  : `${activeDoc.type === "code" ? activeDoc.data.language : "Text"} source`}
          </div>
        </div>

        <div>
          <span
            style={{ fontSize: 10, color: "var(--vscode-tab-inactive-fg)" }}
          >
            FILE IDENTIFIER:
          </span>
          <div
            style={{ fontFamily: "var(--vscode-mono-family)", fontSize: 11 }}
          >
            {activeDoc.data.id}
          </div>
        </div>

        <div>
          <span
            style={{ fontSize: 10, color: "var(--vscode-tab-inactive-fg)" }}
          >
            PARSING ENGINE:
          </span>
          <div>Local JavaScript ? JSZip / SheetJS</div>
        </div>

        <div
          style={{
            display: "flex",
            alignItems: "center",
            gap: 6,
            color: "#4ade80",
            fontSize: 11,
            marginTop: 4,
          }}
        >
          <ShieldCheck size={14} />
          <span>Local workspace model ? no macros executed</span>
        </div>
      </div>
    </div>
  );
};
