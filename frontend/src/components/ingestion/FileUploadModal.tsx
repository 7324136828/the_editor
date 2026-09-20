import React, { useState, useEffect, useRef } from "react";
import {
  Upload,
  FileText,
  FileSpreadsheet,
  Presentation,
  FileArchive,
  X,
  AlertCircle,
  CheckCircle2,
  Loader2,
  Clipboard,
} from "lucide-react";
import { convertDocument } from "../../services/api";

interface FileUploadModalProps {
  isOpen: boolean;
  onClose: () => void;
  onUploadSuccess?: (jobId: string) => void;
  onLoadDirectly?: (file: File) => void;
}

export const FileUploadModal: React.FC<FileUploadModalProps> = ({
  isOpen,
  onClose,
  onUploadSuccess,
  onLoadDirectly,
}) => {
  const [dragActive, setDragActive] = useState(false);
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [uploading, setUploading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [statusMessage, setStatusMessage] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Global Clipboard Paste Listener (Pillar 2)
  useEffect(() => {
    if (!isOpen) return;

    const handlePaste = (e: ClipboardEvent) => {
      const items = e.clipboardData?.items;
      if (!items) return;

      for (let i = 0; i < items.length; i++) {
        const item = items[i];
        if (item.kind === "file") {
          const file = item.getAsFile();
          if (file) {
            setSelectedFile(file);
            setStatusMessage(`Captured file from clipboard: ${file.name}`);
            break;
          }
        }
      }
    };

    window.addEventListener("paste", handlePaste);
    return () => window.removeEventListener("paste", handlePaste);
  }, [isOpen]);

  if (!isOpen) return null;

  const handleDrag = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === "dragenter" || e.type === "dragover") {
      setDragActive(true);
    } else if (e.type === "dragleave") {
      setDragActive(false);
    }
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      const file = e.dataTransfer.files[0];
      setSelectedFile(file);
      setError(null);
    }
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      setSelectedFile(e.target.files[0]);
      setError(null);
    }
  };

  const handleSubmit = async (mode: "convert" | "open") => {
    if (!selectedFile) return;

    if (mode === "open" && onLoadDirectly) {
      onLoadDirectly(selectedFile);
      onClose();
      return;
    }

    setUploading(true);
    setError(null);
    setStatusMessage("Staging file into isolated system temp folder...");

    try {
      const res = await convertDocument(selectedFile);
      setStatusMessage("Job created successfully!");
      if (onUploadSuccess) {
        onUploadSuccess(res.job_id);
      }
      setTimeout(() => {
        onClose();
      }, 800);
    } catch (err: any) {
      setError(err.message || "Failed to process file.");
    } finally {
      setUploading(false);
    }
  };

  return (
    <div
      style={{
        position: "fixed",
        top: 0,
        left: 0,
        right: 0,
        bottom: 0,
        backgroundColor: "rgba(0, 0, 0, 0.65)",
        backdropFilter: "blur(2px)",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        zIndex: 9999,
      }}
      onClick={onClose}
    >
      <div
        style={{
          backgroundColor: "var(--vscode-sideBar-bg, #252526)",
          border: "1px solid var(--vscode-panel-border, #454545)",
          borderRadius: 8,
          width: "90%",
          maxWidth: 520,
          padding: 24,
          color: "var(--vscode-editor-fg, #d4d4d4)",
          fontFamily: "var(--vscode-font-family, system-ui, sans-serif)",
          boxShadow: "0 8px 32px rgba(0,0,0,0.5)",
        }}
        onClick={(e) => e.stopPropagation()}
      >
        {/* Header */}
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
          <h3 style={{ margin: 0, fontSize: 16, display: "flex", alignItems: "center", gap: 8 }}>
            <Upload size={18} style={{ color: "#38bdf8" }} />
            Import / Convert Document
          </h3>
          <button
            onClick={onClose}
            style={{ background: "none", border: "none", color: "#888", cursor: "pointer", padding: 4 }}
          >
            <X size={18} />
          </button>
        </div>

        {/* Dropzone */}
        <div
          onDragEnter={handleDrag}
          onDragLeave={handleDrag}
          onDragOver={handleDrag}
          onDrop={handleDrop}
          onClick={() => fileInputRef.current?.click()}
          style={{
            border: `2px dashed ${dragActive ? "#38bdf8" : "var(--vscode-panel-border, #444)"}`,
            borderRadius: 6,
            padding: "28px 16px",
            textAlign: "center",
            backgroundColor: dragActive ? "rgba(56, 189, 248, 0.08)" : "var(--vscode-editor-bg, #1e1e1e)",
            cursor: "pointer",
            marginBottom: 16,
            transition: "border-color 0.2s, background-color 0.2s",
          }}
        >
          <input
            ref={fileInputRef}
            type="file"
            accept=".docx,.xlsx,.xls,.pptx,.pdf,.csv,.txt,.md,.json"
            onChange={handleFileChange}
            style={{ display: "none" }}
          />

          <div style={{ display: "flex", justifyContent: "center", gap: 8, marginBottom: 12 }}>
            <FileText size={24} style={{ color: "#2b579a" }} />
            <FileSpreadsheet size={24} style={{ color: "#217346" }} />
            <Presentation size={24} style={{ color: "#d24726" }} />
            <FileArchive size={24} style={{ color: "#e37400" }} />
          </div>

          <div style={{ fontWeight: 500, fontSize: 14, marginBottom: 4 }}>
            {selectedFile ? selectedFile.name : "Drag & drop your document here, or click to browse"}
          </div>
          <div style={{ fontSize: 12, color: "var(--vscode-tab-inactive-fg, #888)" }}>
            Supports Word (.docx), Excel (.xlsx, .csv), PowerPoint (.pptx), PDF (.pdf), or text files.
          </div>

          <div
            style={{
              marginTop: 12,
              display: "inline-flex",
              alignItems: "center",
              gap: 4,
              fontSize: 11,
              color: "#60a5fa",
              backgroundColor: "rgba(59, 130, 246, 0.1)",
              padding: "3px 8px",
              borderRadius: 12,
            }}
          >
            <Clipboard size={12} />
            Tip: You can also paste files directly with Ctrl+V
          </div>
        </div>

        {error && (
          <div
            style={{
              padding: "8px 12px",
              backgroundColor: "rgba(239, 68, 68, 0.15)",
              border: "1px solid #ef4444",
              borderRadius: 4,
              color: "#fca5a5",
              fontSize: 12,
              marginBottom: 16,
              display: "flex",
              alignItems: "center",
              gap: 8,
            }}
          >
            <AlertCircle size={14} />
            <span>{error}</span>
          </div>
        )}

        {statusMessage && (
          <div
            style={{
              padding: "8px 12px",
              backgroundColor: "rgba(34, 197, 94, 0.15)",
              border: "1px solid #22c55e",
              borderRadius: 4,
              color: "#86efac",
              fontSize: 12,
              marginBottom: 16,
              display: "flex",
              alignItems: "center",
              gap: 8,
            }}
          >
            <CheckCircle2 size={14} />
            <span>{statusMessage}</span>
          </div>
        )}

        {/* Buttons */}
        <div style={{ display: "flex", justifyContent: "flex-end", gap: 10 }}>
          <button
            onClick={onClose}
            disabled={uploading}
            style={{
              backgroundColor: "transparent",
              color: "var(--vscode-editor-fg, #d4d4d4)",
              border: "1px solid var(--vscode-panel-border, #444)",
              borderRadius: 4,
              padding: "8px 14px",
              fontSize: 13,
              cursor: "pointer",
            }}
          >
            Cancel
          </button>

          {onLoadDirectly && (
            <button
              onClick={() => handleSubmit("open")}
              disabled={!selectedFile || uploading}
              style={{
                backgroundColor: "var(--vscode-button-secondary-bg, #3a3d41)",
                color: "#ffffff",
                border: "none",
                borderRadius: 4,
                padding: "8px 14px",
                fontSize: 13,
                cursor: !selectedFile || uploading ? "not-allowed" : "pointer",
                opacity: !selectedFile || uploading ? 0.6 : 1,
              }}
            >
              Open in Studio
            </button>
          )}

          <button
            onClick={() => handleSubmit("convert")}
            disabled={!selectedFile || uploading}
            style={{
              backgroundColor: "var(--vscode-button-bg, #0e639c)",
              color: "#ffffff",
              border: "none",
              borderRadius: 4,
              padding: "8px 16px",
              fontSize: 13,
              fontWeight: 500,
              cursor: !selectedFile || uploading ? "not-allowed" : "pointer",
              opacity: !selectedFile || uploading ? 0.6 : 1,
              display: "flex",
              alignItems: "center",
              gap: 6,
            }}
          >
            {uploading && <Loader2 size={14} className="animate-spin" style={{ animation: "spin 1s linear infinite" }} />}
            {uploading ? "Processing..." : "Stage & Convert (ZIP)"}
          </button>
        </div>
      </div>
    </div>
  );
};

