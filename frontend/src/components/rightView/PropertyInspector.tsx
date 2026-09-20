import React from "react";
import { OfficeDocument, CellFormat, PptShapeObject } from "../../types/office";
interface PropertyInspectorProps {
  activeDoc: OfficeDocument | null;
  onChangeDoc?: (doc: OfficeDocument) => void;
  selectedPptObjectId?: string | null;
  selectedExcelCoord?: string;
}
export const PropertyInspector: React.FC<PropertyInspectorProps> = ({
  activeDoc: doc,
  onChangeDoc,
  selectedPptObjectId,
  selectedExcelCoord = "A1",
}) => {
  if (!doc) return <p>Select a file to inspect its properties.</p>;
  if (doc.type === "word") {
    const { typography: t } = doc.data;
    const typo = (patch: Partial<typeof t>) =>
      onChangeDoc?.({
        type: "word",
        data: { ...doc.data, typography: { ...t, ...patch } },
      });
    return (
      <div className="inspector-form">
        <h4>DOCUMENT LAYOUT</h4>
        <label>
          Font family
          <select
            className="vscode-input"
            value={t.fontFamily}
            onChange={(e) => typo({ fontFamily: e.target.value })}
          >
            {[
              "Segoe UI",
              "Arial",
              "Calibri",
              "Georgia",
              "Consolas",
              "Times New Roman",
            ].map((f) => (
              <option key={f}>{f}</option>
            ))}
          </select>
        </label>
        <div className="inspector-grid">
          <label>
            Font size
            <input
              type="number"
              min="8"
              max="72"
              className="vscode-input"
              value={t.fontSize}
              onChange={(e) =>
                typo({
                  fontSize: Math.max(8, Math.min(72, Number(e.target.value))),
                })
              }
            />
          </label>
          <label>
            Line height
            <input
              type="number"
              min="1"
              max="3"
              step=".1"
              className="vscode-input"
              value={t.lineHeight}
              onChange={(e) =>
                typo({
                  lineHeight: Math.max(1, Math.min(3, Number(e.target.value))),
                })
              }
            />
          </label>
        </div>
        <label>
          Page margins
          <select
            className="vscode-input"
            value={t.marginSize}
            onChange={(e) =>
              typo({ marginSize: e.target.value as typeof t.marginSize })
            }
          >
            <option value="normal">Normal · 1 inch</option>
            <option value="narrow">Narrow · 0.5 inch</option>
            <option value="wide">Wide · 1.5 inches</option>
          </select>
        </label>
        <label>
          Orientation
          <select
            className="vscode-input"
            value={doc.data.sections[0]?.orientation ?? "portrait"}
            onChange={(e) =>
              onChangeDoc?.({
                type: "word",
                data: {
                  ...doc.data,
                  sections: (doc.data.sections.length
                    ? doc.data.sections
                    : [
                        {
                          id: "section-1",
                          title: "Document",
                          pageNumber: 1,
                          orientation: "portrait",
                        },
                      ]
                  ).map((s) => ({
                    ...s,
                    orientation: e.target.value as "portrait" | "landscape",
                  })),
                },
              })
            }
          >
            <option value="portrait">Portrait</option>
            <option value="landscape">Landscape</option>
          </select>
        </label>
        <label>
          Text color
          <input
            type="color"
            value={t.textColor}
            onChange={(e) => typo({ textColor: e.target.value })}
          />
        </label>
        <h4>PROPERTIES</h4>
        <label>
          Author
          <input
            className="vscode-input"
            value={doc.data.author}
            onChange={(e) =>
              onChangeDoc?.({
                type: "word",
                data: { ...doc.data, author: e.target.value },
              })
            }
          />
        </label>
        <p>
          Page count in statistics is estimated. Use Print to review browser
          pagination.
        </p>
      </div>
    );
  }
  if (doc.type === "excel") {
    const sheet =
      doc.data.sheets.find((s) => s.id === doc.data.activeSheetId) ??
      doc.data.sheets[0];
    if (!sheet) return <p>No worksheet selected.</p>;
    const cell = sheet.cells[selectedExcelCoord],
      fmt = cell?.format ?? {};
    const update = (patch: CellFormat) =>
      onChangeDoc?.({
        type: "excel",
        data: {
          ...doc.data,
          sheets: doc.data.sheets.map((s) =>
            s.id === sheet.id
              ? {
                  ...s,
                  cells: {
                    ...s.cells,
                    [selectedExcelCoord]: {
                      ...cell,
                      value: cell?.value ?? null,
                      formatted: undefined,
                      format: { ...fmt, ...patch },
                    },
                  },
                }
              : s,
          ),
        },
      });
    return (
      <div className="inspector-form">
        <h4>CELL {selectedExcelCoord}</h4>
        <label>
          Number format
          <select
            className="vscode-input"
            value={fmt.numberFormat ?? "general"}
            onChange={(e) =>
              update({
                numberFormat: e.target.value as CellFormat["numberFormat"],
              })
            }
          >
            <option value="general">General</option>
            <option value="currency">Currency ($)</option>
            <option value="percent">Percentage</option>
            <option value="number">Number</option>
          </select>
        </label>
        <label>
          Decimal places
          <input
            type="number"
            min="0"
            max="10"
            className="vscode-input"
            value={fmt.decimalPlaces ?? 2}
            onChange={(e) =>
              update({
                decimalPlaces: Math.max(
                  0,
                  Math.min(10, Number(e.target.value)),
                ),
              })
            }
          />
        </label>
        <label>
          Alignment
          <select
            className="vscode-input"
            value={fmt.align ?? "left"}
            onChange={(e) =>
              update({ align: e.target.value as CellFormat["align"] })
            }
          >
            <option>left</option>
            <option>center</option>
            <option>right</option>
          </select>
        </label>
        <div className="inspector-grid">
          <label>
            Fill
            <input
              type="color"
              value={
                /^#[0-9a-f]{6}$/i.test(fmt.fill ?? "") ? fmt.fill : "#1e1e1e"
              }
              onChange={(e) => update({ fill: e.target.value })}
            />
          </label>
          <label>
            Text color
            <input
              type="color"
              value={fmt.textColor ?? "#d4d4d4"}
              onChange={(e) => update({ textColor: e.target.value })}
            />
          </label>
        </div>
        <button
          className="vscode-btn-secondary"
          onClick={() => update({ fill: undefined, textColor: undefined })}
        >
          Clear colors
        </button>
        <h4>{sheet.name}</h4>
        <p>
          {sheet.rowCount} rows · {sheet.colCount} columns
        </p>
        <p>
          {Object.values(sheet.cells).filter((c) => c.formula).length} formula
          cells
        </p>
        <p>
          Cell styles persist in the workspace. Native XLSX export preserves
          values, formulas and number formats.
        </p>
      </div>
    );
  }
  if (doc.type === "powerpoint") {
    const slide =
        doc.data.slides.find((s) => s.id === doc.data.activeSlideId) ??
        doc.data.slides[0],
      obj = slide?.objects.find((o) => o.id === selectedPptObjectId);
    const patch = (p: Partial<PptShapeObject>) => {
      if (obj)
        onChangeDoc?.({
          type: "powerpoint",
          data: {
            ...doc.data,
            slides: doc.data.slides.map((s) =>
              s.id === slide.id
                ? {
                    ...s,
                    objects: s.objects.map((o) =>
                      o.id === obj.id ? { ...o, ...p } : o,
                    ),
                  }
                : s,
            ),
          },
        });
    };
    return (
      <div className="inspector-form">
        <h4>SLIDE PROPERTIES</h4>
        <label>
          Background
          <input
            type="color"
            value={
              /^#[0-9a-f]{6}$/i.test(slide?.background)
                ? slide.background
                : "#15243b"
            }
            onChange={(e) =>
              onChangeDoc?.({
                type: "powerpoint",
                data: {
                  ...doc.data,
                  slides: doc.data.slides.map((s) =>
                    s.id === slide.id
                      ? { ...s, background: e.target.value }
                      : s,
                  ),
                },
              })
            }
          />
        </label>
        {!obj ? (
          <p>Select an object to edit its geometry and appearance.</p>
        ) : (
          <>
            <h4>SELECTED {obj.kind.toUpperCase()}</h4>
            <div className="inspector-grid">
              {(
                ["x", "y", "width", "height", "fontSize", "rotation"] as const
              ).map((key) => (
                <label key={key}>
                  {key}
                  <input
                    className="vscode-input"
                    type="number"
                    value={obj[key] ?? 0}
                    disabled={obj.locked}
                    onChange={(e) =>
                      patch({
                        [key]: Math.max(
                          key === "width" ||
                            key === "height" ||
                            key === "fontSize"
                            ? 1
                            : -2000,
                          Math.min(4000, Number(e.target.value)),
                        ),
                      })
                    }
                  />
                </label>
              ))}
            </div>
            <label>
              Text
              <textarea
                className="vscode-input"
                value={obj.text ?? obj.title ?? ""}
                disabled={obj.locked}
                onChange={(e) =>
                  patch(
                    obj.text !== undefined
                      ? { text: e.target.value }
                      : { title: e.target.value },
                  )
                }
              />
            </label>
            <div className="inspector-grid">
              <label>
                Fill
                <input
                  type="color"
                  value={
                    /^#[0-9a-f]{6}$/i.test(obj.fill) ? obj.fill : "#15243b"
                  }
                  onChange={(e) => patch({ fill: e.target.value })}
                />
              </label>
              <label>
                Text color
                <input
                  type="color"
                  value={
                    /^#[0-9a-f]{6}$/i.test(obj.color) ? obj.color : "#ffffff"
                  }
                  onChange={(e) => patch({ color: e.target.value })}
                />
              </label>
            </div>
            <button
              className="vscode-btn-secondary"
              onClick={() =>
                patch({
                  zIndex: Math.max(...slide.objects.map((o) => o.zIndex)) + 1,
                })
              }
            >
              Bring to front
            </button>
            <button
              className="vscode-btn-secondary"
              onClick={() =>
                patch({
                  zIndex: Math.min(...slide.objects.map((o) => o.zIndex)) - 1,
                })
              }
            >
              Send to back
            </button>
            <button
              className="vscode-btn-secondary"
              onClick={() => patch({ locked: !obj.locked })}
            >
              {obj.locked ? "Unlock object" : "Lock object"}
            </button>
            <button
              className="vscode-btn-secondary"
              disabled={obj.locked}
              onClick={() =>
                onChangeDoc?.({
                  type: "powerpoint",
                  data: {
                    ...doc.data,
                    slides: doc.data.slides.map((s) =>
                      s.id === slide.id
                        ? {
                            ...s,
                            objects: s.objects.filter((o) => o.id !== obj.id),
                          }
                        : s,
                    ),
                  },
                })
              }
            >
              Delete object
            </button>
          </>
        )}
      </div>
    );
  }
  return (
    <div className="inspector-form">
      <h4>TEXT DOCUMENT</h4>
      <label>
        Language
        <select
          className="vscode-input"
          value={doc.data.language}
          onChange={(e) =>
            onChangeDoc?.({
              type: "code",
              data: {
                ...doc.data,
                language: e.target.value as typeof doc.data.language,
              },
            })
          }
        >
          {["typescript", "javascript", "python", "markdown", "json"].map(
            (l) => (
              <option key={l}>{l}</option>
            ),
          )}
        </select>
      </label>
      <p>
        {doc.data.content.split("\n").length} lines · {doc.data.content.length}{" "}
        characters
      </p>
      <p>
        Plain-text editing with local symbols and JSON diagnostics. Code is
        never executed by opening a file.
      </p>
    </div>
  );
};
