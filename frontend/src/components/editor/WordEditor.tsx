import React, { useEffect, useLayoutEffect, useRef, useState } from "react";
import { WordDocumentModel, WordParagraph } from "../../types/office";
import {
  Bold,
  Italic,
  AlignLeft,
  AlignCenter,
  AlignRight,
  ZoomIn,
  ZoomOut,
  MessageSquarePlus,
  Plus,
  Table2,
  Trash2,
  ArrowDown,
  ArrowUp,
  GitFork,
  X,
} from "lucide-react";
import "../../styles/editors.css";

interface WordEditorProps {
  document: WordDocumentModel;
  onChangeDocument?: (doc: WordDocumentModel) => void;
  targetHeadingId?: string | null;
}

function EditableText({
  value,
  label,
  style,
  onChange,
  onFocus,
  onKeyDown,
}: {
  value: string;
  label: string;
  style?: React.CSSProperties;
  onChange: (text: string) => void;
  onFocus?: () => void;
  onKeyDown?: React.KeyboardEventHandler<HTMLTextAreaElement>;
}) {
  const ref = useRef<HTMLTextAreaElement>(null);
  useLayoutEffect(() => {
    if (ref.current) {
      ref.current.style.height = "0px";
      ref.current.style.height = `${ref.current.scrollHeight + 2}px`;
    }
  }, [value, style]);
  return (
    <textarea
      ref={ref}
      className="word-editable"
      aria-label={label}
      value={value}
      rows={1}
      style={style}
      onChange={(event) => onChange(event.target.value)}
      onFocus={onFocus}
      onKeyDown={onKeyDown}
      spellCheck
    />
  );
}

export function updateWordMetadata(
  document: WordDocumentModel,
): WordDocumentModel {
  const content = document.paragraphs
    .map((paragraph) =>
      paragraph.type === "table"
        ? (paragraph.tableData || []).flat().join(" ")
        : paragraph.type === "smartart"
          ? [
              paragraph.smartArt?.title ?? "",
              ...(paragraph.smartArt?.items ?? []).map((item) => item.text),
            ].join(" ")
          : (paragraph.text ??
            paragraph.runs?.map((run) => run.text).join("") ??
            ""),
    )
    .join("\n");
  const words = content.trim() ? content.trim().split(/\s+/u).length : 0;
  return {
    ...document,
    modifiedAt: new Date().toISOString(),
    headings: document.paragraphs
      .filter((p) => p.type.startsWith("heading-"))
      .map((p) => ({
        id: p.id,
        level: Number(p.type.slice(-1)) as 1 | 2 | 3,
        text: p.text || "",
        pageIndex: 1,
      })),
    stats: {
      ...document.stats,
      words,
      characters: content.length,
      charactersNoSpaces: content.replace(/\s/gu, "").length,
      paragraphs: document.paragraphs.length,
      pages: Math.max(1, Math.ceil(words / 450)),
      readingTimeMinutes: Math.max(1, Math.ceil(words / 200)),
    },
  };
}

export const WordEditor: React.FC<WordEditorProps> = ({
  document,
  onChangeDocument,
  targetHeadingId,
}) => {
  const [zoom, setZoom] = useState(100);
  const [activeParaId, setActiveParaId] = useState<string | null>(
    document.paragraphs[0]?.id || null,
  );
  const [showComments, setShowComments] = useState(false);
  const [commentDraft, setCommentDraft] = useState("");
  const pageRef = useRef<HTMLDivElement>(null);
  const active = document.paragraphs.find((p) => p.id === activeParaId);
  const commit = (next: WordDocumentModel) =>
    onChangeDocument?.(updateWordMetadata(next));
  const updateParagraph = (id: string, patch: Partial<WordParagraph>) =>
    commit({
      ...document,
      paragraphs: document.paragraphs.map((p) =>
        p.id === id ? { ...p, ...patch } : p,
      ),
    });
  const textOf = (p: WordParagraph) =>
    p.text ?? p.runs?.map((run) => run.text).join("") ?? "";
  const changeText = (p: WordParagraph, text: string) =>
    updateParagraph(p.id, { text, runs: [{ ...p.runs?.[0], text }] });
  const toggleStyle = (key: "bold" | "italic") => {
    if (!active || ["table", "smartart"].includes(active.type)) return;
    const runs = active.runs?.length ? active.runs : [{ text: textOf(active) }];
    const value = !runs.every((run) => run[key]);
    updateParagraph(active.id, {
      runs: runs.map((run) => ({ ...run, [key]: value })),
    });
  };
  const insert = (type: "body" | "table") => {
    const paragraph: WordParagraph = {
      id: `p-${crypto.randomUUID()}`,
      type,
      text: "",
      ...(type === "table"
        ? {
            tableData: [
              ["Column 1", "Column 2", "Column 3"],
              ["", "", ""],
              ["", "", ""],
            ],
          }
        : {}),
    };
    const paragraphs = [...document.paragraphs];
    const index = paragraphs.findIndex((p) => p.id === activeParaId);
    paragraphs.splice(index < 0 ? paragraphs.length : index + 1, 0, paragraph);
    commit({ ...document, paragraphs });
    setActiveParaId(paragraph.id);
    requestAnimationFrame(() => {
      const inserted = [
        ...(pageRef.current?.querySelectorAll("[data-paragraph-id]") || []),
      ].find(
        (element) => element.getAttribute("data-paragraph-id") === paragraph.id,
      );
      inserted?.scrollIntoView({ block: "center", behavior: "smooth" });
      inserted?.querySelector("textarea")?.focus();
    });
  };
  useEffect(() => {
    if (!targetHeadingId) return;
    const heading = document.headings.find((h) => h.id === targetHeadingId);
    const target =
      document.paragraphs.find((p) => p.id === targetHeadingId) ||
      document.paragraphs.find(
        (p) => p.type.startsWith("heading-") && p.text === heading?.text,
      );
    const element = [
      ...(pageRef.current?.querySelectorAll("[data-paragraph-id]") || []),
    ].find((el) => el.getAttribute("data-paragraph-id") === target?.id);
    element?.scrollIntoView({ behavior: "smooth", block: "center" });
    if (target) setActiveParaId(target.id);
  }, [targetHeadingId]);
  const scale = zoom / 100;
  const landscape = document.sections[0]?.orientation === "landscape";
  const margin = { narrow: 36, normal: 72, wide: 108 }[
    document.typography.marginSize
  ];
  return (
    <div className="office-editor">
      <div className="office-toolbar">
        <span className="office-badge word-badge">WORD</span>
        <select
          aria-label="Font family"
          className="vscode-input"
          value={document.typography.fontFamily}
          onChange={(e) =>
            commit({
              ...document,
              typography: {
                ...document.typography,
                fontFamily: e.target.value,
              },
            })
          }
        >
          {[
            "Segoe UI",
            "Arial",
            "Calibri",
            "Georgia",
            "Consolas",
            "Times New Roman",
          ].map((font) => (
            <option key={font}>{font}</option>
          ))}
        </select>
        <input
          aria-label="Font size"
          className="vscode-input toolbar-number"
          type="number"
          min="8"
          max="72"
          value={document.typography.fontSize}
          onChange={(e) =>
            commit({
              ...document,
              typography: {
                ...document.typography,
                fontSize: Math.max(8, Math.min(72, Number(e.target.value))),
              },
            })
          }
        />
        <select
          aria-label="Paragraph style"
          className="vscode-input"
          value={active?.type || "body"}
          disabled={!active || ["table", "smartart"].includes(active.type)}
          onChange={(e) =>
            active &&
            updateParagraph(active.id, {
              type: e.target.value as WordParagraph["type"],
            })
          }
        >
          <option value="body">Body</option>
          <option value="heading-1">Heading 1</option>
          <option value="heading-2">Heading 2</option>
          <option value="heading-3">Heading 3</option>
          <option value="quote">Quote</option>
          <option value="callout">Callout</option>
          <option value="table" hidden>
            Table
          </option>
          <option value="smartart" hidden>
            SmartArt
          </option>
        </select>
        <button
          className="vscode-icon-btn"
          title="Bold paragraph (Ctrl+B)"
          aria-label="Bold paragraph"
          aria-pressed={active?.runs?.every((r) => r.bold) || false}
          disabled={!active || ["table", "smartart"].includes(active.type)}
          onClick={() => toggleStyle("bold")}
        >
          <Bold size={14} />
        </button>
        <button
          className="vscode-icon-btn"
          title="Italic paragraph (Ctrl+I)"
          aria-label="Italic paragraph"
          aria-pressed={active?.runs?.every((r) => r.italic) || false}
          disabled={!active || ["table", "smartart"].includes(active.type)}
          onClick={() => toggleStyle("italic")}
        >
          <Italic size={14} />
        </button>
        {(["left", "center", "right"] as const).map((align, index) => {
          const Icon = [AlignLeft, AlignCenter, AlignRight][index];
          return (
            <button
              key={align}
              className="vscode-icon-btn"
              title={`Align ${align}`}
              aria-label={`Align ${align}`}
              aria-pressed={active?.align === align}
              disabled={!active || active.type === "smartart"}
              onClick={() => active && updateParagraph(active.id, { align })}
            >
              <Icon size={14} />
            </button>
          );
        })}
        <button
          className="vscode-icon-btn"
          title="Insert paragraph after selection"
          aria-label="Add paragraph"
          onClick={() => insert("body")}
        >
          <Plus size={14} />
        </button>
        <button
          className="vscode-icon-btn"
          title="Insert editable table"
          aria-label="Add table"
          onClick={() => insert("table")}
        >
          <Table2 size={14} />
        </button>
        <button
          className="vscode-icon-btn"
          title="Delete selected paragraph, table, or SmartArt"
          aria-label="Delete paragraph"
          disabled={!active}
          onClick={() => {
            commit({
              ...document,
              paragraphs: document.paragraphs.filter(
                (p) => p.id !== activeParaId,
              ),
            });
            setActiveParaId(null);
          }}
        >
          <Trash2 size={14} />
        </button>
        <button
          className="vscode-icon-btn"
          title="Comments"
          aria-label="Comments"
          aria-pressed={showComments}
          onClick={() => setShowComments(!showComments)}
        >
          <MessageSquarePlus size={14} />
        </button>
        <span className="toolbar-spacer" />
        <button
          className="vscode-icon-btn"
          title="Zoom out"
          aria-label="Zoom out"
          onClick={() => setZoom(Math.max(50, zoom - 10))}
        >
          <ZoomOut size={14} />
        </button>
        <span>{zoom}%</span>
        <button
          className="vscode-icon-btn"
          title="Zoom in"
          aria-label="Zoom in"
          onClick={() => setZoom(Math.min(160, zoom + 10))}
        >
          <ZoomIn size={14} />
        </button>
      </div>
      {active?.type === "table" && (
        <div className="office-toolbar compact-toolbar">
          <span>Selected table</span>
          <button
            className="vscode-btn-secondary"
            onClick={() =>
              updateParagraph(active.id, {
                tableData: [
                  ...(active.tableData || []),
                  Array(active.tableData?.[0]?.length || 3).fill(""),
                ],
              })
            }
          >
            Add row
          </button>
          <button
            className="vscode-btn-secondary"
            onClick={() =>
              updateParagraph(active.id, {
                tableData: (active.tableData || [[""]]).map((row) => [
                  ...row,
                  "",
                ]),
              })
            }
          >
            Add column
          </button>
        </div>
      )}
      {active?.type === "smartart" && active.smartArt && (
        <div className="office-toolbar compact-toolbar smartart-toolbar">
          <GitFork size={14} />
          <span>Selected SmartArt</span>
          <label>
            Layout
            <select
              className="vscode-input"
              aria-label="SmartArt layout"
              value={active.smartArt.layout}
              onChange={(event) =>
                updateParagraph(active.id, {
                  smartArt: {
                    ...active.smartArt!,
                    layout: event.target.value as NonNullable<
                      WordParagraph["smartArt"]
                    >["layout"],
                  },
                })
              }
            >
              <option value="process">Process</option>
              <option value="cycle">Cycle</option>
              <option value="hierarchy">Hierarchy</option>
              <option value="pyramid">Pyramid</option>
            </select>
          </label>
          <label>
            Color
            <input
              type="color"
              aria-label="SmartArt accent color"
              value={active.smartArt.accentColor}
              onChange={(event) =>
                updateParagraph(active.id, {
                  smartArt: {
                    ...active.smartArt!,
                    accentColor: event.target.value,
                  },
                })
              }
            />
          </label>
          <button
            className="vscode-btn-secondary"
            disabled={active.smartArt.items.length >= 12}
            onClick={() =>
              updateParagraph(active.id, {
                smartArt: {
                  ...active.smartArt!,
                  items: [
                    ...active.smartArt!.items,
                    { id: `smartart-${crypto.randomUUID()}`, text: "New item" },
                  ],
                },
              })
            }
          >
            <Plus size={12} /> Add item
          </button>
        </div>
      )}
      <div className="word-workspace">
        <div className="word-scroll">
          <div
            ref={pageRef}
            className="word-page"
            style={{
              width: (landscape ? 1056 : 816) * scale,
              minHeight: (landscape ? 816 : 1056) * scale,
              padding: margin * scale,
              fontFamily: document.typography.fontFamily,
              fontSize: document.typography.fontSize * scale,
              lineHeight: document.typography.lineHeight,
              color: document.typography.textColor,
            }}
          >
            <div className="word-page-header">
              <span>{document.title}</span>
              <span>{document.author}</span>
            </div>
            {document.paragraphs.map((p, index) => {
              const selected = p.id === activeParaId;
              return (
                <div
                  key={p.id}
                  data-paragraph-id={p.id}
                  className={`word-block ${p.type} ${selected ? "selected" : ""}`}
                  style={{ textAlign: p.align || "left" }}
                  onClick={() => setActiveParaId(p.id)}
                >
                  {p.type === "table" ? (
                    <table className="word-table">
                      <tbody>
                        {p.tableData?.map((row, r) => (
                          <tr key={r}>
                            {row.map((value, c) => (
                              <td key={c}>
                                <EditableText
                                  value={value}
                                  label={`Table ${index + 1}, row ${r + 1}, column ${c + 1}`}
                                  onFocus={() => setActiveParaId(p.id)}
                                  onChange={(text) =>
                                    updateParagraph(p.id, {
                                      tableData: p.tableData!.map(
                                        (oldRow, ri) =>
                                          oldRow.map((cell, ci) =>
                                            r === ri && c === ci ? text : cell,
                                          ),
                                      ),
                                    })
                                  }
                                />
                              </td>
                            ))}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  ) : p.type === "smartart" && p.smartArt ? (
                    <div
                      className={`word-smartart smartart-${p.smartArt.layout}`}
                      style={
                        {
                          "--smartart-accent": p.smartArt.accentColor,
                        } as React.CSSProperties
                      }
                    >
                      <input
                        className="smartart-title"
                        aria-label={`SmartArt ${index + 1} title`}
                        value={p.smartArt.title}
                        maxLength={120}
                        onFocus={() => setActiveParaId(p.id)}
                        onChange={(event) =>
                          updateParagraph(p.id, {
                            smartArt: {
                              ...p.smartArt!,
                              title: event.target.value,
                            },
                          })
                        }
                      />
                      <div className="smartart-items">
                        {p.smartArt.items.map((item, itemIndex) => (
                          <div className="smartart-item" key={item.id}>
                            <span className="smartart-number">
                              {itemIndex + 1}
                            </span>
                            <textarea
                              aria-label={`SmartArt ${index + 1}, item ${itemIndex + 1}`}
                              value={item.text}
                              maxLength={500}
                              rows={2}
                              onFocus={() => setActiveParaId(p.id)}
                              onChange={(event) =>
                                updateParagraph(p.id, {
                                  smartArt: {
                                    ...p.smartArt!,
                                    items: p.smartArt!.items.map((old) =>
                                      old.id === item.id
                                        ? { ...old, text: event.target.value }
                                        : old,
                                    ),
                                  },
                                })
                              }
                            />
                            <div className="smartart-item-actions">
                              <button
                                type="button"
                                aria-label={`Move SmartArt item ${itemIndex + 1} up`}
                                disabled={itemIndex === 0}
                                onClick={() => {
                                  const items = [...p.smartArt!.items];
                                  [items[itemIndex - 1], items[itemIndex]] = [
                                    items[itemIndex],
                                    items[itemIndex - 1],
                                  ];
                                  updateParagraph(p.id, {
                                    smartArt: { ...p.smartArt!, items },
                                  });
                                }}
                              >
                                <ArrowUp size={12} />
                              </button>
                              <button
                                type="button"
                                aria-label={`Move SmartArt item ${itemIndex + 1} down`}
                                disabled={
                                  itemIndex === p.smartArt!.items.length - 1
                                }
                                onClick={() => {
                                  const items = [...p.smartArt!.items];
                                  [items[itemIndex], items[itemIndex + 1]] = [
                                    items[itemIndex + 1],
                                    items[itemIndex],
                                  ];
                                  updateParagraph(p.id, {
                                    smartArt: { ...p.smartArt!, items },
                                  });
                                }}
                              >
                                <ArrowDown size={12} />
                              </button>
                              <button
                                type="button"
                                aria-label={`Delete SmartArt item ${itemIndex + 1}`}
                                disabled={p.smartArt!.items.length === 1}
                                onClick={() =>
                                  updateParagraph(p.id, {
                                    smartArt: {
                                      ...p.smartArt!,
                                      items: p.smartArt!.items.filter(
                                        (old) => old.id !== item.id,
                                      ),
                                    },
                                  })
                                }
                              >
                                <X size={12} />
                              </button>
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>
                  ) : (
                    <EditableText
                      value={textOf(p)}
                      label={`${p.type.startsWith("heading-") ? "Heading" : "Paragraph"} ${index + 1}`}
                      style={{
                        textAlign: p.align || "left",
                        fontWeight: p.runs?.every((r) => r.bold)
                          ? "bold"
                          : undefined,
                        fontStyle:
                          p.runs?.every((r) => r.italic) || p.type === "quote"
                            ? "italic"
                            : undefined,
                      }}
                      onFocus={() => setActiveParaId(p.id)}
                      onChange={(text) => changeText(p, text)}
                      onKeyDown={(event) => {
                        if (
                          (event.ctrlKey || event.metaKey) &&
                          ["b", "i"].includes(event.key.toLowerCase())
                        ) {
                          event.preventDefault();
                          event.stopPropagation();
                          toggleStyle(
                            event.key.toLowerCase() === "b" ? "bold" : "italic",
                          );
                        }
                      }}
                    />
                  )}
                </div>
              );
            })}
            {!document.paragraphs.length && (
              <button className="vscode-btn" onClick={() => insert("body")}>
                Start writing
              </button>
            )}
            <div className="word-page-footer">
              <span>
                {document.stats.words.toLocaleString()} words ·{" "}
                {document.stats.paragraphs} paragraphs
              </span>
              <span>Letter · {landscape ? "Landscape" : "Portrait"}</span>
            </div>
          </div>
        </div>
        {showComments && (
          <aside className="word-comments">
            <h3>Document comments</h3>
            <p className="muted">Comment on the selected paragraph.</p>
            <textarea
              aria-label="New comment"
              className="vscode-input"
              value={commentDraft}
              onChange={(e) => setCommentDraft(e.target.value)}
              placeholder="Add a comment…"
            />
            <button
              className="vscode-btn"
              disabled={!commentDraft.trim()}
              onClick={() => {
                commit({
                  ...document,
                  comments: [
                    ...document.comments,
                    {
                      id: crypto.randomUUID(),
                      author: "You",
                      avatarColor: "#007acc",
                      timestamp: new Date().toISOString(),
                      selectedText: active ? textOf(active).slice(0, 120) : "",
                      comment: commentDraft.trim(),
                      resolved: false,
                    },
                  ],
                });
                setCommentDraft("");
              }}
            >
              Add comment
            </button>
            {document.comments.map((comment) => (
              <article
                key={comment.id}
                className="word-comment"
                style={{ opacity: comment.resolved ? 0.55 : 1 }}
              >
                <strong>{comment.author}</strong>
                <blockquote>{comment.selectedText}</blockquote>
                <p>{comment.comment}</p>
                <button
                  className="vscode-btn-secondary"
                  onClick={() =>
                    commit({
                      ...document,
                      comments: document.comments.map((c) =>
                        c.id === comment.id
                          ? { ...c, resolved: !c.resolved }
                          : c,
                      ),
                    })
                  }
                >
                  {comment.resolved ? "Reopen" : "Resolve"}
                </button>
              </article>
            ))}
          </aside>
        )}
      </div>
    </div>
  );
};
