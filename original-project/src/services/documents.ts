import { OfficeDocument } from "../types/office";
import {
  sampleWordDocument,
  sampleExcelDocument,
  samplePptDocument,
  sampleCodeDocument,
} from "../data/sampleDocuments";
import { recalculateWorkbook } from "./formulas";

export const clone = <T>(value: T): T => JSON.parse(JSON.stringify(value));
export const fingerprint = (doc: OfficeDocument) => JSON.stringify(doc);
export const documentExtension = (doc: OfficeDocument) =>
  ({
    word: "docx",
    excel: "xlsx",
    powerpoint: "pptx",
    code:
      doc.type === "code"
        ? {
            typescript: "ts",
            javascript: "js",
            python: "py",
            markdown: "md",
            json: "json",
          }[doc.data.language]
        : "txt",
  })[doc.type];
export const sampleDocuments: OfficeDocument[] = [
  { type: "word", data: sampleWordDocument },
  { type: "excel", data: sampleExcelDocument },
  { type: "powerpoint", data: samplePptDocument },
  { type: "code", data: sampleCodeDocument },
];

export function normalizeDocument(doc: OfficeDocument): OfficeDocument {
  if (doc.type === "excel")
    return { type: "excel", data: recalculateWorkbook(doc.data) };
  if (doc.type !== "word") return doc;
  const text = doc.data.paragraphs
    .map(
      (p) =>
        p.tableData?.flat().join(" ") ??
        (p.smartArt
          ? [
              p.smartArt.title,
              ...p.smartArt.items.map((item) => item.text),
            ].join(" ")
          : undefined) ??
        p.runs?.map((r) => r.text).join("") ??
        p.text ??
        "",
    )
    .join("\n");
  const words = text.trim() ? text.trim().split(/\s+/).length : 0;
  return {
    ...doc,
    data: {
      ...doc.data,
      headings: doc.data.paragraphs
        .filter((p) => p.type.startsWith("heading-"))
        .map((p) => ({
          id: p.id,
          level: Number(p.type.slice(-1)) as 1 | 2 | 3,
          text: p.text ?? "",
          pageIndex: 1,
        })),
      stats: {
        ...doc.data.stats,
        words,
        characters: text.length,
        charactersNoSpaces: text.replace(/\s/g, "").length,
        paragraphs: doc.data.paragraphs.length,
        pages: Math.max(1, Math.ceil(words / 450)),
        readingTimeMinutes: Math.ceil(words / 220),
        readabilityScore: "Not analyzed",
      },
    },
  };
}

export function createDocument(
  type: OfficeDocument["type"],
  name: string,
): OfficeDocument {
  const doc = clone(sampleDocuments.find((d) => d.type === type)!);
  doc.data.id = `doc-${crypto.randomUUID()}`;
  doc.data.title = name;
  if (doc.type === "word") {
    doc.data.author = "Local author";
    doc.data.createdAt = doc.data.modifiedAt = new Date().toISOString();
    doc.data.paragraphs = [
      { id: "p-1", type: "heading-1", text: "Untitled document" },
      { id: "p-2", type: "body", text: "Start writing here." },
    ];
    doc.data.comments = [];
    doc.data.sections = [
      {
        id: "section-1",
        title: "Document",
        pageNumber: 1,
        orientation: "portrait",
      },
    ];
  } else if (doc.type === "excel") {
    doc.data.sheets = [
      {
        id: "sheet-1",
        name: "Sheet1",
        rowCount: 20,
        colCount: 8,
        cells: {},
        columns: Array.from({ length: 8 }, (_, i) => ({
          key: String.fromCharCode(65 + i),
          label: String.fromCharCode(65 + i),
          width: 115,
          type: "string",
        })),
      },
    ];
    doc.data.activeSheetId = "sheet-1";
    doc.data.formulasAudit = [];
    doc.data.namedRanges = {};
  } else if (doc.type === "powerpoint") {
    doc.data.slides = [
      {
        id: "slide-1",
        slideNumber: 1,
        title: "Untitled slide",
        layout: "title",
        background: "#15243b",
        notes: "",
        transition: "none",
        objects: [
          {
            id: "title-1",
            kind: "text",
            text: "Your next great idea",
            x: 80,
            y: 130,
            width: 800,
            height: 160,
            fill: "transparent",
            color: "#ffffff",
            fontSize: 48,
            fontWeight: "bold",
            zIndex: 1,
          },
        ],
      },
    ];
    doc.data.activeSlideId = "slide-1";
  } else {
    doc.data.content = "";
    doc.data.symbols = [];
    const ext = name.split(".").pop();
    doc.data.language =
      ext === "py"
        ? "python"
        : ext === "json"
          ? "json"
          : ext === "ts"
            ? "typescript"
            : ext === "js"
              ? "javascript"
              : "markdown";
  }
  return normalizeDocument(doc);
}
