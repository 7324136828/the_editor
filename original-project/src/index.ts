// VS Code Office Studio - Node Module Entry Point
export { VSCodeOfficeStudio, default } from "./App";
export type { VSCodeOfficeStudioProps } from "./App";

// Layout & UI Components
export { TitleBar } from "./components/layout/TitleBar";
export { ActivityBar } from "./components/layout/ActivityBar";
export { StatusBar } from "./components/layout/StatusBar";
export { CommandPalette } from "./components/layout/CommandPalette";
export { PrimarySidebar } from "./components/sidebar/PrimarySidebar";
export { AccordionSection } from "./components/sidebar/AccordionSection";
export { EditorContainer } from "./components/editor/EditorContainer";
export { BottomPanel } from "./components/bottomPanel/BottomPanel";
export { SecondarySidebar } from "./components/rightView/SecondarySidebar";
export { PanelResizer } from "./components/layout/PanelResizer";
export { ManageViewsDialog } from "./components/layout/ManageViewsDialog";
export { useViewPreferences } from "./views/useViewPreferences";
export type { ViewPreferencesController } from "./views/useViewPreferences";
export type {
  CustomView,
  ViewArea,
  ViewPreferences,
} from "./views/preferences";

// Editors
export { WordEditor } from "./components/editor/WordEditor";
export { ExcelEditor } from "./components/editor/ExcelEditor";
export { PptEditor } from "./components/editor/PptEditor";
export { CodeEditor } from "./components/editor/CodeEditor";
export { PresentationModal } from "./components/editor/PresentationModal";

// Parsers & Loaders
export { loadOfficeFile } from "./parsers/fileLoader";
export { parseDocxFile } from "./parsers/docxParser";
export { parseXlsxFile } from "./parsers/xlsxParser";
export { parsePptxFile } from "./parsers/pptxParser";

// Sample Documents
export {
  sampleWordDocument,
  sampleExcelDocument,
  samplePptDocument,
  sampleCodeDocument,
} from "./data/sampleDocuments";

// Types
export * from "./types/vscode";
export * from "./types/office";
export * from "./plugins";
export { ExtensionsPanel } from "./components/plugins/ExtensionsPanel";
export { exportDocument, serializeDocument } from "./services/exportDocument";
export { evaluateFormula, recalculateWorkbook } from "./services/formulas";
