export type ActivityBarTab =
  | "explorer"
  | "search"
  | "source-control"
  | "debug"
  | "extensions"
  | "settings";

export type BottomPanelTab =
  "search" | "problems" | "output" | "terminal" | "debug";

export type RightViewTab = "inspector" | "copilot" | "minimap" | "metadata";

export type DocumentType = "word" | "excel" | "powerpoint" | "code";

export interface EditorTab {
  id: string;
  title: string;
  filePath: string;
  docType: DocumentType;
  isDirty?: boolean;
  isActive?: boolean;
}

export interface SearchMatch {
  id: string;
  fileId: string;
  fileName: string;
  fileType: DocumentType;
  location: string; // e.g., "Paragraph 3", "Sheet1!B5", "Slide 2", "Line 42"
  previewText: string;
  matchIndex: number;
  matchLength: number;
  targetRef?: {
    sheetId?: string;
    cellCoord?: string;
    slideIndex?: number;
    elementId?: string;
    headingId?: string;
    lineNumber?: number;
  };
}

export interface ProblemItem {
  id: string;
  severity: "error" | "warning" | "info";
  message: string;
  source: string;
  fileId: string;
  fileName: string;
  location: string;
}

export interface OutputLogItem {
  id: string;
  timestamp: string;
  channel: "Office Studio" | "Parser" | "Formulas" | "Build";
  level: "info" | "warn" | "error";
  message: string;
}

export interface CommandItem {
  id: string;
  title: string;
  category?: string;
  shortcut?: string;
  action: () => void;
}
