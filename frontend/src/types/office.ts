// Word Document Types
export interface WordHeading {
  id: string;
  level: 1 | 2 | 3 | 4;
  text: string;
  pageIndex: number;
}

export interface WordComment {
  id: string;
  author: string;
  avatarColor: string;
  timestamp: string;
  selectedText: string;
  comment: string;
  resolved: boolean;
}

export interface WordSection {
  id: string;
  title: string;
  pageNumber: number;
  orientation: "portrait" | "landscape";
}

export type SmartArtLayout = "process" | "cycle" | "hierarchy" | "pyramid";

export interface SmartArtItem {
  id: string;
  text: string;
}

export interface WordSmartArt {
  title: string;
  layout: SmartArtLayout;
  accentColor: string;
  items: SmartArtItem[];
}

export interface WordParagraph {
  id: string;
  type:
    | "heading-1"
    | "heading-2"
    | "heading-3"
    | "body"
    | "quote"
    | "callout"
    | "table"
    | "smartart";
  text?: string;
  runs?: {
    text: string;
    bold?: boolean;
    italic?: boolean;
    highlight?: string;
  }[];
  tableData?: string[][];
  smartArt?: WordSmartArt;
  align?: "left" | "center" | "right" | "justify";
}

export interface WordDocumentModel {
  id: string;
  title: string;
  author: string;
  createdAt: string;
  modifiedAt: string;
  headings: WordHeading[];
  sections: WordSection[];
  paragraphs: WordParagraph[];
  comments: WordComment[];
  stats: {
    words: number;
    characters: number;
    charactersNoSpaces: number;
    paragraphs: number;
    pages: number;
    readingTimeMinutes: number;
    readabilityScore: string;
  };
  typography: {
    fontFamily: string;
    fontSize: number;
    lineHeight: number;
    marginSize: "normal" | "narrow" | "wide";
    textColor: string;
  };
}

// Excel Spreadsheet Types
export type CellValue = string | number | boolean | null;

export interface CellFormat {
  bold?: boolean;
  italic?: boolean;
  align?: "left" | "center" | "right";
  fill?: string;
  textColor?: string;
  numberFormat?: "general" | "currency" | "percent" | "number";
  decimalPlaces?: number;
}

export interface CellData {
  value: CellValue;
  formula?: string;
  formatted?: string;
  format?: CellFormat;
}

export interface ExcelColumn {
  key: string;
  label: string;
  width: number;
  type: "string" | "number" | "currency" | "date";
}

export interface ExcelWorksheet {
  id: string;
  name: string;
  rowCount: number;
  colCount: number;
  cells: Record<string, CellData>; // key like "A1", "B2"
  columns: ExcelColumn[];
}

export interface ExcelFormulaAudit {
  cellCoord: string;
  formula: string;
  evaluatedValue: string | number;
  dependencies: string[];
}

export interface ExcelDocumentModel {
  id: string;
  title: string;
  activeSheetId: string;
  sheets: ExcelWorksheet[];
  namedRanges?: Record<string, string>;
  formulasAudit: ExcelFormulaAudit[];
}

// PowerPoint Presentation Types
export type PptShapeKind =
  "rectangle" | "rounded-rect" | "circle" | "card" | "badge" | "metric-box";

export interface PptShapeObject {
  id: string;
  kind: "text" | "shape" | "metric" | "image" | "chart";
  shapeKind?: PptShapeKind;
  x: number;
  y: number;
  width: number;
  height: number;
  text?: string;
  title?: string;
  subtitle?: string;
  metricValue?: string;
  metricLabel?: string;
  fill: string;
  stroke?: string;
  strokeWidth?: number;
  borderRadius?: number;
  rotation?: number;
  color: string;
  fontSize: number;
  fontWeight?: "normal" | "bold" | "600";
  align?: "left" | "center" | "right";
  zIndex: number;
  locked?: boolean;
  visible?: boolean;
}

export interface PptSlide {
  id: string;
  slideNumber: number;
  title: string;
  layout: "title" | "content" | "two-column" | "dashboard" | "blank";
  background: string;
  objects: PptShapeObject[];
  notes: string;
  transition: "none" | "fade" | "slide" | "zoom";
}

export interface PptDocumentModel {
  id: string;
  title: string;
  activeSlideId: string;
  themeName: string;
  accentColor: string;
  slides: PptSlide[];
}

// Code / Text Document Types
export interface CodeDocumentModel {
  id: string;
  title: string;
  language: "typescript" | "javascript" | "python" | "markdown" | "json";
  content: string;
  symbols: {
    name: string;
    kind: "function" | "interface" | "class" | "variable";
    line: number;
  }[];
}

// Union Office Document
export type OfficeDocument =
  | { type: "word"; data: WordDocumentModel }
  | { type: "excel"; data: ExcelDocumentModel }
  | { type: "powerpoint"; data: PptDocumentModel }
  | { type: "code"; data: CodeDocumentModel };
