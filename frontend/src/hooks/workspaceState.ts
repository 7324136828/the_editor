import type { OfficeDocument } from "../types/office";
import type { StoredDocument } from "../services/api";

export interface WorkspaceEntry {
  id: string;
  name: string;
  document: OfficeDocument;
  revision: number | null;
  saved: string | null;
  saving?: boolean;
  error?: string;
  updatedAt?: string;
  seeded?: boolean;
}
export interface WorkspaceState {
  entries: Record<string, WorkspaceEntry>;
  openIds: string[];
  activeId: string;
}

const object = (value: unknown): value is Record<string, any> =>
  value !== null && typeof value === "object" && !Array.isArray(value);
const string = (value: unknown): value is string => typeof value === "string";
const number = (value: unknown): value is number =>
  typeof value === "number" && Number.isFinite(value);
const items = (value: unknown, valid: (item: any) => boolean): boolean =>
  Array.isArray(value) && value.every(valid);
const optionalStrings = (value: Record<string, any>, keys: string[]) =>
  keys.every((key) => value[key] === undefined || string(value[key]));
const optionalNumbers = (value: Record<string, any>, keys: string[]) =>
  keys.every((key) => value[key] === undefined || number(value[key]));
const id = (value: unknown): value is string =>
  string(value) && /^[a-zA-Z0-9][a-zA-Z0-9_-]{0,127}$/.test(value);

/** Browser recovery and HTTP responses are data boundaries, not typed TS objects. */
export function isOfficeDocument(value: unknown): value is OfficeDocument {
  if (
    !object(value) ||
    !object(value.data) ||
    !id(value.data.id) ||
    !string(value.data.title)
  )
    return false;
  const data = value.data;
  if (value.type === "word") {
    return (
      ["author", "createdAt", "modifiedAt"].every((key) => string(data[key])) &&
      items(
        data.paragraphs,
        (paragraph) =>
          object(paragraph) &&
          string(paragraph.id) &&
          [
            "heading-1",
            "heading-2",
            "heading-3",
            "body",
            "quote",
            "callout",
            "table",
            "smartart",
          ].includes(paragraph.type) &&
          optionalStrings(paragraph, ["text"]) &&
          (paragraph.runs === undefined ||
            items(paragraph.runs, (run) => object(run) && string(run.text))) &&
          (paragraph.tableData === undefined ||
            items(paragraph.tableData, (row) => items(row, string))) &&
          (paragraph.smartArt === undefined
            ? paragraph.type !== "smartart"
            : paragraph.type === "smartart" &&
              object(paragraph.smartArt) &&
              string(paragraph.smartArt.title) &&
              ["process", "cycle", "hierarchy", "pyramid"].includes(
                paragraph.smartArt.layout,
              ) &&
              /^#[a-f\d]{6}$/i.test(paragraph.smartArt.accentColor) &&
              items(
                paragraph.smartArt.items,
                (item) => object(item) && string(item.id) && string(item.text),
              )),
      ) &&
      items(
        data.headings,
        (heading) =>
          object(heading) &&
          string(heading.id) &&
          string(heading.text) &&
          number(heading.level) &&
          number(heading.pageIndex),
      ) &&
      items(
        data.sections,
        (section) =>
          object(section) &&
          string(section.id) &&
          string(section.title) &&
          number(section.pageNumber) &&
          ["portrait", "landscape"].includes(section.orientation),
      ) &&
      items(
        data.comments,
        (comment) =>
          object(comment) &&
          [
            "id",
            "author",
            "avatarColor",
            "timestamp",
            "selectedText",
            "comment",
          ].every((key) => string(comment[key])) &&
          typeof comment.resolved === "boolean",
      ) &&
      object(data.typography) &&
      string(data.typography.fontFamily) &&
      string(data.typography.textColor) &&
      number(data.typography.fontSize) &&
      data.typography.fontSize > 0 &&
      number(data.typography.lineHeight) &&
      data.typography.lineHeight > 0 &&
      ["normal", "narrow", "wide"].includes(data.typography.marginSize) &&
      object(data.stats) &&
      [
        "words",
        "characters",
        "charactersNoSpaces",
        "paragraphs",
        "pages",
        "readingTimeMinutes",
      ].every((key) => number(data.stats[key])) &&
      string(data.stats.readabilityScore)
    );
  }
  if (value.type === "excel") {
    return (
      string(data.activeSheetId) &&
      items(
        data.sheets,
        (sheet) =>
          object(sheet) &&
          string(sheet.id) &&
          string(sheet.name) &&
          Number.isInteger(sheet.rowCount) &&
          sheet.rowCount >= 1 &&
          sheet.rowCount <= 1000 &&
          Number.isInteger(sheet.colCount) &&
          sheet.colCount >= 1 &&
          sheet.colCount <= 100 &&
          sheet.rowCount * sheet.colCount <= 10000 &&
          items(
            sheet.columns,
            (column) =>
              object(column) &&
              string(column.key) &&
              /^[A-Z]+$/.test(column.key) &&
              string(column.label) &&
              number(column.width) &&
              column.width > 0,
          ) &&
          sheet.columns.length <= 100 &&
          sheet.rowCount * sheet.columns.length <= 10000 &&
          object(sheet.cells) &&
          Object.entries(sheet.cells).every(
            ([coord, cell]) =>
              /^[A-Z]{1,3}[1-9]\d{0,6}$/.test(coord) &&
              object(cell) &&
              (cell.value === null ||
                ["string", "boolean"].includes(typeof cell.value) ||
                number(cell.value)) &&
              optionalStrings(cell, ["formula", "formatted"]) &&
              (cell.format === undefined ||
                (object(cell.format) &&
                  optionalStrings(cell.format, [
                    "align",
                    "fill",
                    "textColor",
                    "numberFormat",
                  ]) &&
                  optionalNumbers(cell.format, ["decimalPlaces"]))),
          ),
      ) &&
      data.sheets.length > 0 &&
      data.sheets.length <= 50 &&
      data.sheets.some((sheet: any) => sheet.id === data.activeSheetId) &&
      items(
        data.formulasAudit,
        (audit) =>
          object(audit) &&
          string(audit.cellCoord) &&
          string(audit.formula) &&
          (string(audit.evaluatedValue) || number(audit.evaluatedValue)) &&
          items(audit.dependencies, string),
      ) &&
      (data.namedRanges === undefined ||
        (object(data.namedRanges) &&
          Object.values(data.namedRanges).every(string)))
    );
  }
  if (value.type === "powerpoint") {
    return (
      string(data.activeSlideId) &&
      string(data.themeName) &&
      string(data.accentColor) &&
      items(
        data.slides,
        (slide) =>
          object(slide) &&
          ["id", "title", "background", "notes"].every((key) =>
            string(slide[key]),
          ) &&
          number(slide.slideNumber) &&
          ["title", "content", "two-column", "dashboard", "blank"].includes(
            slide.layout,
          ) &&
          ["none", "fade", "slide", "zoom"].includes(slide.transition) &&
          items(
            slide.objects,
            (shape) =>
              object(shape) &&
              string(shape.id) &&
              ["text", "shape", "metric", "image", "chart"].includes(
                shape.kind,
              ) &&
              ["x", "y", "width", "height", "fontSize", "zIndex"].every((key) =>
                number(shape[key]),
              ) &&
              string(shape.fill) &&
              string(shape.color) &&
              optionalStrings(shape, [
                "text",
                "title",
                "subtitle",
                "metricValue",
                "metricLabel",
                "stroke",
                "align",
                "shapeKind",
                "fontWeight",
              ]) &&
              optionalNumbers(shape, [
                "rotation",
                "strokeWidth",
                "borderRadius",
              ]),
          ),
      ) &&
      data.slides.length > 0 &&
      data.slides.some((slide: any) => slide.id === data.activeSlideId)
    );
  }
  return (
    value.type === "code" &&
    ["typescript", "javascript", "python", "markdown", "json"].includes(
      data.language,
    ) &&
    string(data.content) &&
    items(
      data.symbols,
      (symbol) =>
        object(symbol) &&
        string(symbol.name) &&
        ["function", "interface", "class", "variable"].includes(symbol.kind) &&
        number(symbol.line),
    )
  );
}

/** Salvage valid entries without allowing corrupt identities or baselines to poison startup. */
export function recoverWorkspace(value: unknown): WorkspaceState | null {
  if (!object(value) || !object(value.entries) || !Array.isArray(value.openIds))
    return null;
  const entries: Record<string, WorkspaceEntry> = Object.create(null);
  for (const [key, candidate] of Object.entries(value.entries)) {
    if (
      !object(candidate) ||
      candidate.id !== key ||
      !isOfficeDocument(candidate.document) ||
      candidate.document.data.id !== key ||
      !string(candidate.name) ||
      (candidate.revision !== null &&
        !(Number.isInteger(candidate.revision) && candidate.revision >= 1))
    )
      continue;
    let saved: string | null = null;
    if (string(candidate.saved)) {
      try {
        const baseline: unknown = JSON.parse(candidate.saved);
        if (
          isOfficeDocument(baseline) &&
          baseline.data.id === key &&
          baseline.type === candidate.document.type
        )
          saved = JSON.stringify(baseline);
      } catch {
        /* Retain current text as a dirty draft if its baseline is damaged. */
      }
    }
    entries[key] = {
      id: key,
      name: candidate.document.data.title,
      document: candidate.document,
      revision: candidate.revision,
      saved,
      saving: false,
      updatedAt: string(candidate.updatedAt) ? candidate.updatedAt : undefined,
      seeded:
        candidate.seeded === true &&
        candidate.revision === null &&
        saved === null,
      error: string(candidate.error) ? candidate.error : undefined,
    };
  }
  if (Object.keys(value.entries).length && !Object.keys(entries).length)
    return null;
  const openIds = [
    ...new Set<string>(
      value.openIds.filter(
        (key: unknown): key is string => string(key) && !!entries[key],
      ),
    ),
  ];
  const activeId =
    string(value.activeId) && openIds.includes(value.activeId)
      ? value.activeId
      : (openIds[0] ?? "");
  return { entries, openIds, activeId };
}

export function isStoredDocument(value: unknown): value is StoredDocument {
  return (
    object(value) &&
    id(value.id) &&
    string(value.name) &&
    string(value.updatedAt) &&
    Number.isInteger(value.revision) &&
    value.revision >= 1 &&
    isOfficeDocument(value.document) &&
    value.document.data.id === value.id
  );
}

/** Never let an older list response overwrite edits or saves begun after that request. */
export function mergeWorkspaceRecords(
  current: WorkspaceState,
  atRequest: WorkspaceState,
  records: readonly StoredDocument[],
  savingIds: ReadonlySet<string> = new Set(),
): WorkspaceState {
  const entries = { ...current.entries };
  for (const record of records) {
    const existing = entries[record.id];
    const original = atRequest.entries[record.id];
    if (existing !== original || existing?.saving || savingIds.has(record.id))
      continue;
    if (existing?.revision != null && record.revision < existing.revision)
      continue;
    if (
      !existing ||
      existing.seeded ||
      existing.saved === JSON.stringify(existing.document)
    ) {
      entries[record.id] = {
        ...record,
        name: record.document.data.title,
        saved: JSON.stringify(record.document),
        saving: false,
        seeded: false,
      };
    }
  }
  return { ...current, entries };
}
