import React, { useRef } from "react";
import { EditorTab, DocumentType } from "../../types/vscode";
import {
  FileText,
  Table,
  Presentation,
  FileCode,
  Upload,
  Plus,
} from "lucide-react";

interface FileExplorerViewProps {
  files: { id: string; name: string; type: DocumentType; size?: string }[];
  activeFileId: string;
  onOpenFile: (fileId: string) => void;
  onUploadFile: (file: File) => void;
}

export const FileExplorerView: React.FC<FileExplorerViewProps> = ({
  files,
  activeFileId,
  onOpenFile,
  onUploadFile,
}) => {
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFileInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      onUploadFile(file);
    }
    e.target.value = "";
  };

  const getFileIcon = (type: DocumentType) => {
    switch (type) {
      case "word":
        return <FileText size={14} color="#4175c5" />;
      case "excel":
        return <Table size={14} color="#107c41" />;
      case "powerpoint":
        return <Presentation size={14} color="#c43e1c" />;
      case "code":
      default:
        return <FileCode size={14} color="#007acc" />;
    }
  };

  return (
    <div style={{ padding: "2px 0" }}>
      <input
        ref={fileInputRef}
        type="file"
        accept=".docx,.xlsx,.xls,.pptx,.csv,.txt,.md,.json,.ts,.js,.py"
        style={{ display: "none" }}
        onChange={handleFileInputChange}
      />

      <div
        style={{
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
          padding: "4px 12px 6px",
          borderBottom: "1px solid rgba(255,255,255,0.05)",
        }}
      >
        <span
          style={{
            fontSize: 10,
            color: "var(--vscode-tab-inactive-fg)",
            fontWeight: 600,
          }}
        >
          WORKSPACE FILES
        </span>
        <button
          className="vscode-btn-secondary"
          style={{ padding: "2px 6px", fontSize: 10 }}
          onClick={() => fileInputRef.current?.click()}
          title="Upload Word, Excel, PowerPoint or code file from computer"
        >
          <Upload size={10} />
          <span>Load File</span>
        </button>
      </div>

      {files.map((file) => {
        const isActive = file.id === activeFileId;
        return (
          <div
            key={file.id}
            className={`vscode-tree-item ${isActive ? "active" : ""}`}
            onClick={() => onOpenFile(file.id)}
            title={`Open ${file.name}`}
          >
            {getFileIcon(file.type)}
            <span style={{ overflow: "hidden", textOverflow: "ellipsis" }}>
              {file.name}
            </span>
          </div>
        );
      })}
    </div>
  );
};
