import React, { useState, useMemo } from "react";
import { SearchOptions, searchDocuments } from "../../services/search";
import { SearchMatch, DocumentType } from "../../types/vscode";
import { OfficeDocument } from "../../types/office";
import {
  Search,
  Replace,
  CaseSensitive,
  WholeWord,
  Regex,
  ChevronRight,
  ChevronDown,
  FileText,
  Table,
  Presentation,
  FileCode,
  CheckCheck,
} from "lucide-react";

interface FindInFilesViewProps {
  documents: {
    id: string;
    name: string;
    type: DocumentType;
    doc: OfficeDocument;
  }[];
  onNavigateToMatch: (match: SearchMatch) => void;
  onReplaceAll?: (
    query: string,
    replaceWith: string,
    options?: SearchOptions,
  ) => void;
}

export const FindInFilesView: React.FC<FindInFilesViewProps> = ({
  documents,
  onNavigateToMatch,
  onReplaceAll,
}) => {
  const [query, setQuery] = useState("");
  const [replaceQuery, setReplaceQuery] = useState("");
  const [showReplace, setShowReplace] = useState(true);
  const [matchCase, setMatchCase] = useState(false);
  const [wholeWord, setWholeWord] = useState(false);
  const [useRegex, setUseRegex] = useState(false);
  const [collapsedFiles, setCollapsedFiles] = useState<Record<string, boolean>>(
    {},
  );

  const [searchError, setSearchError] = useState("");
  const matches = useMemo(() => {
    try {
      return searchDocuments(documents, query, {
        matchCase,
        wholeWord,
        useRegex,
      });
    } catch {
      return [];
    }
  }, [query, matchCase, wholeWord, useRegex, documents]);
  React.useEffect(() => {
    try {
      searchDocuments([], query, { matchCase, wholeWord, useRegex });
      setSearchError("");
    } catch (error) {
      setSearchError((error as Error).message);
    }
  }, [query, matchCase, wholeWord, useRegex]);

  // Group matches by file
  const groupedMatches = useMemo(() => {
    const groups: Record<
      string,
      { fileName: string; fileType: DocumentType; matches: SearchMatch[] }
    > = {};
    matches.forEach((m) => {
      if (!groups[m.fileId]) {
        groups[m.fileId] = {
          fileName: m.fileName,
          fileType: m.fileType,
          matches: [],
        };
      }
      groups[m.fileId].matches.push(m);
    });
    return groups;
  }, [matches]);

  const toggleCollapseFile = (fileId: string) => {
    setCollapsedFiles((prev) => ({ ...prev, [fileId]: !prev[fileId] }));
  };

  const getDocIcon = (type: DocumentType) => {
    switch (type) {
      case "word":
        return <FileText size={13} color="#4175c5" />;
      case "excel":
        return <Table size={13} color="#107c41" />;
      case "powerpoint":
        return <Presentation size={13} color="#c43e1c" />;
      default:
        return <FileCode size={13} color="#007acc" />;
    }
  };

  return (
    <div style={{ display: "flex", height: "100%", gap: 16, fontSize: 12 }}>
      {/* Search Input Controls Sidebar */}
      <div
        style={{
          width: 320,
          minWidth: 260,
          display: "flex",
          flexDirection: "column",
          gap: 6,
          borderRight: "1px solid #333",
          paddingRight: 16,
        }}
      >
        {/* Search Bar */}
        <div
          style={{
            position: "relative",
            display: "flex",
            alignItems: "center",
          }}
        >
          <input
            className="vscode-input"
            style={{ width: "100%", paddingRight: 74, height: 26 }}
            placeholder="Search across all open documents..."
            value={query}
            onChange={(e) => setQuery(e.target.value)}
          />
          <div
            style={{
              position: "absolute",
              right: 4,
              display: "flex",
              alignItems: "center",
              gap: 2,
            }}
          >
            <button
              className={`vscode-icon-btn ${matchCase ? "active" : ""}`}
              style={{ width: 20, height: 20 }}
              onClick={() => setMatchCase(!matchCase)}
              title="Match Case (Alt+C)"
            >
              <CaseSensitive size={12} />
            </button>
            <button
              className={`vscode-icon-btn ${wholeWord ? "active" : ""}`}
              style={{ width: 20, height: 20 }}
              onClick={() => setWholeWord(!wholeWord)}
              title="Match Whole Word (Alt+W)"
            >
              <WholeWord size={12} />
            </button>
            <button
              className={`vscode-icon-btn ${useRegex ? "active" : ""}`}
              style={{ width: 20, height: 20 }}
              onClick={() => setUseRegex(!useRegex)}
              title="Use Regular Expression (Alt+R)"
            >
              <Regex size={12} />
            </button>
          </div>
        </div>

        {/* Replace Bar */}
        {showReplace && (
          <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
            <input
              className="vscode-input"
              style={{ flex: 1, height: 26 }}
              placeholder="Replace with..."
              value={replaceQuery}
              onChange={(e) => setReplaceQuery(e.target.value)}
            />
            <button
              className="vscode-btn-secondary"
              style={{ height: 26, padding: "0 8px", fontSize: 11 }}
              title="Replace All across loaded documents"
              disabled={!matches.length || !!searchError}
              onClick={() =>
                onReplaceAll?.(query, replaceQuery, {
                  matchCase,
                  wholeWord,
                  useRegex,
                })
              }
            >
              <CheckCheck size={12} />
              <span>Replace All</span>
            </button>
          </div>
        )}

        <div
          style={{
            fontSize: 11,
            color: "var(--vscode-tab-inactive-fg)",
            marginTop: 4,
          }}
        >
          {searchError ? (
            <span role="alert">{searchError}</span>
          ) : query.trim() ? (
            <span>
              Found <strong>{matches.length}</strong> results in{" "}
              <strong>{Object.keys(groupedMatches).length}</strong> files
            </span>
          ) : (
            <span>
              Type to search across Word, Excel, PowerPoint, and code files
            </span>
          )}
        </div>
      </div>

      {/* Grouped Match Results Tree */}
      <div style={{ flex: 1, overflowY: "auto", paddingRight: 8 }}>
        {Object.entries(groupedMatches).map(([fileId, group]) => {
          const isCollapsed = collapsedFiles[fileId];
          return (
            <div key={fileId} style={{ marginBottom: 8 }}>
              {/* File Group Header */}
              <div
                onClick={() => toggleCollapseFile(fileId)}
                style={{
                  display: "flex",
                  alignItems: "center",
                  gap: 6,
                  padding: "4px 6px",
                  backgroundColor: "var(--vscode-sidebar-bg)",
                  borderRadius: 3,
                  cursor: "pointer",
                  fontSize: 11,
                  fontWeight: 600,
                  color: "var(--vscode-editor-fg)",
                }}
              >
                {isCollapsed ? (
                  <ChevronRight size={13} />
                ) : (
                  <ChevronDown size={13} />
                )}
                {getDocIcon(group.fileType)}
                <span>{group.fileName}</span>
                <span
                  style={{
                    marginLeft: "auto",
                    fontSize: 10,
                    backgroundColor: "rgba(255, 255, 255, 0.1)",
                    padding: "1px 6px",
                    borderRadius: 10,
                  }}
                >
                  {group.matches.length}
                </span>
              </div>

              {/* Match Items in this file */}
              {!isCollapsed && (
                <div
                  style={{
                    paddingLeft: 16,
                    marginTop: 4,
                    display: "flex",
                    flexDirection: "column",
                    gap: 2,
                  }}
                >
                  {group.matches.map((m) => (
                    <div
                      key={m.id}
                      className="vscode-tree-item"
                      style={{
                        borderRadius: 3,
                        padding: "4px 8px",
                        justifyContent: "space-between",
                      }}
                      onClick={() => onNavigateToMatch(m)}
                      title={`Jump to ${m.location}`}
                    >
                      <div
                        style={{
                          display: "flex",
                          alignItems: "center",
                          gap: 8,
                          minWidth: 0,
                        }}
                      >
                        <span
                          style={{
                            fontSize: 10,
                            color: "#38bdf8",
                            fontFamily: "var(--vscode-mono-family)",
                            minWidth: 90,
                          }}
                        >
                          {m.location}
                        </span>
                        <span
                          style={{
                            overflow: "hidden",
                            textOverflow: "ellipsis",
                            color: "var(--vscode-sidebar-fg)",
                          }}
                        >
                          {m.previewText}
                        </span>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
};
