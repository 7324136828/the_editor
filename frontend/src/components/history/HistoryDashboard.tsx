import React, { useState, useEffect } from "react";
import {
  History,
  Download,
  Trash2,
  RefreshCw,
  CheckCircle2,
  XCircle,
  AlertCircle,
  Clock,
  FileArchive,
  ArrowRight,
} from "lucide-react";
import { listJobs, discardJob, downloadZipUrl, JobRecord } from "../../services/api";

interface HistoryDashboardProps {
  onBackToStudio?: () => void;
  onOpenUploadModal?: () => void;
}

export const HistoryDashboard: React.FC<HistoryDashboardProps> = ({
  onBackToStudio,
  onOpenUploadModal,
}) => {
  const [jobs, setJobs] = useState<JobRecord[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  const fetchJobs = async () => {
    try {
      const data = await listJobs();
      setJobs(data);
      setError(null);
    } catch (err: any) {
      setError(err.message || "Failed to load conversion jobs history.");
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchJobs();
    const interval = setInterval(fetchJobs, 4000);
    return () => clearInterval(interval);
  }, []);

  const handleDiscard = async (jobId: string) => {
    if (!window.confirm("Are you sure you want to discard this conversion job and purge its temporary files?")) {
      return;
    }
    try {
      await discardJob(jobId);
      await fetchJobs();
    } catch (err: any) {
      alert(`Discard failed: ${err.message}`);
    }
  };

  const formatBytes = (bytes: number) => {
    if (!bytes) return "0 B";
    const k = 1024;
    const sizes = ["B", "KB", "MB", "GB"];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return `${(bytes / Math.pow(k, i)).toFixed(1)} ${sizes[i]}`;
  };

  const getStatusBadge = (status: JobRecord["status"]) => {
    switch (status) {
      case "completed":
        return (
          <span
            style={{
              display: "inline-flex",
              alignItems: "center",
              gap: 4,
              backgroundColor: "rgba(34, 197, 94, 0.15)",
              color: "#4ade80",
              padding: "2px 8px",
              borderRadius: 12,
              fontSize: 12,
              fontWeight: 500,
            }}
          >
            <CheckCircle2 size={12} />
            Completed
          </span>
        );
      case "in_progress":
        return (
          <span
            style={{
              display: "inline-flex",
              alignItems: "center",
              gap: 4,
              backgroundColor: "rgba(59, 130, 246, 0.15)",
              color: "#60a5fa",
              padding: "2px 8px",
              borderRadius: 12,
              fontSize: 12,
              fontWeight: 500,
            }}
          >
            <RefreshCw size={12} className="animate-spin" style={{ animation: "spin 1.5s linear infinite" }} />
            Converting
          </span>
        );
      case "discarded":
        return (
          <span
            style={{
              display: "inline-flex",
              alignItems: "center",
              gap: 4,
              backgroundColor: "rgba(107, 114, 128, 0.15)",
              color: "#9ca3af",
              padding: "2px 8px",
              borderRadius: 12,
              fontSize: 12,
            }}
          >
            <Clock size={12} />
            Discarded
          </span>
        );
      case "failed":
        return (
          <span
            style={{
              display: "inline-flex",
              alignItems: "center",
              gap: 4,
              backgroundColor: "rgba(239, 68, 68, 0.15)",
              color: "#f87171",
              padding: "2px 8px",
              borderRadius: 12,
              fontSize: 12,
            }}
          >
            <XCircle size={12} />
            Failed
          </span>
        );
      default:
        return (
          <span
            style={{
              backgroundColor: "rgba(107, 114, 128, 0.15)",
              color: "#9ca3af",
              padding: "2px 8px",
              borderRadius: 12,
              fontSize: 12,
            }}
          >
            {status}
          </span>
        );
    }
  };

  return (
    <div
      style={{
        flex: 1,
        padding: "24px 32px",
        overflowY: "auto",
        backgroundColor: "var(--vscode-editor-bg, #1e1e1e)",
        color: "var(--vscode-editor-fg, #d4d4d4)",
        fontFamily: "var(--vscode-font-family, system-ui, sans-serif)",
      }}
    >
      {/* Title & Actions */}
      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          marginBottom: 20,
          borderBottom: "1px solid var(--vscode-panel-border, #333)",
          paddingBottom: 16,
        }}
      >
        <div>
          <h2 style={{ margin: "0 0 4px 0", fontSize: 20, display: "flex", alignItems: "center", gap: 10 }}>
            <History size={22} style={{ color: "#38bdf8" }} />
            Historical Conversions & Processing Ledger
          </h2>
          <p style={{ margin: 0, fontSize: 13, color: "var(--vscode-tab-inactive-fg, #888)" }}>
            Persistent record of document conversion jobs with isolated system temp execution & ZIP packaging.
          </p>
        </div>

        <div style={{ display: "flex", gap: 10 }}>
          {onOpenUploadModal && (
            <button
              onClick={onOpenUploadModal}
              style={{
                backgroundColor: "var(--vscode-button-bg, #0e639c)",
                color: "#ffffff",
                border: "none",
                borderRadius: 4,
                padding: "8px 14px",
                fontSize: 13,
                cursor: "pointer",
                display: "flex",
                alignItems: "center",
                gap: 6,
                fontWeight: 500,
              }}
            >
              Upload & Convert Document
            </button>
          )}
          {onBackToStudio && (
            <button
              onClick={onBackToStudio}
              style={{
                backgroundColor: "transparent",
                color: "var(--vscode-editor-fg, #d4d4d4)",
                border: "1px solid var(--vscode-panel-border, #444)",
                borderRadius: 4,
                padding: "8px 14px",
                fontSize: 13,
                cursor: "pointer",
                display: "flex",
                alignItems: "center",
                gap: 6,
              }}
            >
              Back to Studio <ArrowRight size={14} />
            </button>
          )}
        </div>
      </div>

      {error && (
        <div
          style={{
            padding: "10px 14px",
            backgroundColor: "rgba(239, 68, 68, 0.15)",
            border: "1px solid #ef4444",
            borderRadius: 6,
            color: "#fca5a5",
            marginBottom: 16,
            display: "flex",
            alignItems: "center",
            gap: 8,
          }}
        >
          <AlertCircle size={16} />
          <span>{error}</span>
        </div>
      )}

      {loading ? (
        <div style={{ padding: 40, textAlign: "center", color: "var(--vscode-tab-inactive-fg, #888)" }}>
          <RefreshCw size={24} className="animate-spin" style={{ margin: "0 auto 12px auto", display: "block" }} />
          Loading conversion history...
        </div>
      ) : jobs.length === 0 ? (
        <div
          style={{
            padding: "60px 20px",
            textAlign: "center",
            border: "2px dashed var(--vscode-panel-border, #333)",
            borderRadius: 8,
            color: "var(--vscode-tab-inactive-fg, #888)",
          }}
        >
          <FileArchive size={36} style={{ margin: "0 auto 12px auto", opacity: 0.6 }} />
          <h4 style={{ margin: "0 0 6px 0", color: "var(--vscode-editor-fg, #d4d4d4)" }}>No conversions found</h4>
          <p style={{ margin: "0 0 16px 0", fontSize: 13 }}>
            Upload or paste Word, Excel, PowerPoint, or PDF documents to stage and extract content.
          </p>
          {onOpenUploadModal && (
            <button
              onClick={onOpenUploadModal}
              style={{
                backgroundColor: "var(--vscode-button-bg, #0e639c)",
                color: "#ffffff",
                border: "none",
                borderRadius: 4,
                padding: "8px 16px",
                fontSize: 13,
                cursor: "pointer",
              }}
            >
              Upload Document
            </button>
          )}
        </div>
      ) : (
        <div
          style={{
            border: "1px solid var(--vscode-panel-border, #333)",
            borderRadius: 6,
            overflow: "hidden",
          }}
        >
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
            <thead>
              <tr
                style={{
                  backgroundColor: "var(--vscode-sideBar-bg, #252526)",
                  borderBottom: "1px solid var(--vscode-panel-border, #333)",
                  textAlign: "left",
                }}
              >
                <th style={{ padding: "10px 14px", fontWeight: 600 }}>Filename</th>
                <th style={{ padding: "10px 14px", fontWeight: 600 }}>Size</th>
                <th style={{ padding: "10px 14px", fontWeight: 600 }}>Status</th>
                <th style={{ padding: "10px 14px", fontWeight: 600 }}>Created</th>
                <th style={{ padding: "10px 14px", fontWeight: 600 }}>Completed</th>
                <th style={{ padding: "10px 14px", fontWeight: 600, textAlign: "right" }}>Actions</th>
              </tr>
            </thead>
            <tbody>
              {jobs.map((job) => (
                <tr
                  key={job.id}
                  style={{
                    borderBottom: "1px solid var(--vscode-panel-border, #2d2d2d)",
                    backgroundColor: "transparent",
                  }}
                >
                  <td style={{ padding: "12px 14px", fontWeight: 500 }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                      <FileArchive size={16} style={{ color: "#60a5fa" }} />
                      <span>{job.filename}</span>
                    </div>
                  </td>
                  <td style={{ padding: "12px 14px", color: "var(--vscode-tab-inactive-fg, #888)" }}>
                    {formatBytes(job.file_size)}
                  </td>
                  <td style={{ padding: "12px 14px" }}>{getStatusBadge(job.status)}</td>
                  <td style={{ padding: "12px 14px", color: "var(--vscode-tab-inactive-fg, #888)" }}>
                    {new Date(job.created_at).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" })}
                  </td>
                  <td style={{ padding: "12px 14px", color: "var(--vscode-tab-inactive-fg, #888)" }}>
                    {job.completed_at
                      ? new Date(job.completed_at).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" })
                      : "—"}
                  </td>
                  <td style={{ padding: "12px 14px", textAlign: "right" }}>
                    <div style={{ display: "flex", justifyContent: "flex-end", gap: 8 }}>
                      {job.status === "in_progress" && (
                        <button
                          onClick={() => handleDiscard(job.id)}
                          style={{
                            backgroundColor: "rgba(239, 68, 68, 0.15)",
                            color: "#f87171",
                            border: "1px solid rgba(239, 68, 68, 0.3)",
                            borderRadius: 4,
                            padding: "4px 10px",
                            fontSize: 12,
                            cursor: "pointer",
                            display: "flex",
                            alignItems: "center",
                            gap: 4,
                          }}
                        >
                          <Trash2 size={12} /> Discard
                        </button>
                      )}

                      {job.status === "completed" && (
                        <a
                          href={downloadZipUrl(job.id)}
                          download
                          style={{
                            backgroundColor: "var(--vscode-button-bg, #0e639c)",
                            color: "#ffffff",
                            textDecoration: "none",
                            borderRadius: 4,
                            padding: "4px 12px",
                            fontSize: 12,
                            fontWeight: 500,
                            display: "inline-flex",
                            alignItems: "center",
                            gap: 4,
                          }}
                        >
                          <Download size={12} /> Download ZIP
                        </a>
                      )}

                      {(job.status === "failed" || job.status === "discarded") && (
                        <span style={{ fontSize: 12, color: "var(--vscode-tab-inactive-fg, #666)" }}>
                          {job.status}
                        </span>
                      )}
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
};

