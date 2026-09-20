import React, { useEffect, useState } from "react";
import { History, RotateCcw } from "lucide-react";
import {
  getDocumentHistory,
  getDocumentVersion,
  DocumentVersion,
} from "../../services/api";
import { OfficeDocument } from "../../types/office";
import { WorkspaceEntry } from "../../hooks/useWorkspace";
export function HistoryPanel({
  entry,
  onRestore,
  onNotify,
}: {
  entry?: WorkspaceEntry;
  onRestore: (doc: OfficeDocument) => void;
  onNotify: (message: string) => void;
}) {
  const [versions, setVersions] = useState<DocumentVersion[]>([]),
    [error, setError] = useState(""),
    [busy, setBusy] = useState(false);
  useEffect(() => {
    let cancelled = false;
    setVersions([]);
    setError("");
    if (entry?.revision)
      getDocumentHistory(entry.id)
        .then((v) => {
          if (!cancelled) setVersions(v);
        })
        .catch((e) => {
          if (!cancelled) setError(e.message);
        });
    return () => {
      cancelled = true;
    };
  }, [entry?.id, entry?.revision]);
  return (
    <aside className="vscode-sidebar utility-sidebar">
      <div className="vscode-sidebar-title-strip">
        <History size={14} /> VERSION HISTORY
      </div>
      <div className="utility-body">
        <h3>{entry?.name ?? "No active file"}</h3>
        <p>
          Saved workspace revisions. Restore loads a version into the editor;
          Save creates a new revision.
        </p>
        {error && <p role="alert">{error}</p>}
        {!versions.length && !error && (
          <p>Save this file to create its first revision.</p>
        )}
        {versions.map((v) => (
          <div className="history-item" key={v.revision}>
            <strong>Revision {v.revision}</strong>
            <time>{new Date(v.updatedAt).toLocaleString()}</time>
            <button
              disabled={busy}
              className="vscode-btn-secondary"
              onClick={async () => {
                setBusy(true);
                try {
                  const record = await getDocumentVersion(
                    entry!.id,
                    v.revision,
                  );
                  onRestore(record.document);
                  onNotify(
                    `Loaded revision ${v.revision}. Save to keep this restoration.`,
                  );
                } catch (e) {
                  setError((e as Error).message);
                } finally {
                  setBusy(false);
                }
              }}
            >
              <RotateCcw size={12} />
              Restore to editor
            </button>
          </div>
        ))}
      </div>
    </aside>
  );
}
