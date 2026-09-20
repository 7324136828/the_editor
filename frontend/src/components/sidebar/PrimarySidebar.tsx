import React from "react";
import { EditorTab, DocumentType } from "../../types/vscode";
import { OfficeDocument } from "../../types/office";
import { AccordionSection } from "./AccordionSection";
import { OpenEditorsView } from "./OpenEditorsView";
import { FileExplorerView } from "./FileExplorerView";
import { WordOutlineView } from "./WordOutlineView";
import { WordPagesView } from "./WordPagesView";
import { WordCommentsView } from "./WordCommentsView";
import { WordStatsView } from "./WordStatsView";
import { ExcelSheetsView } from "./ExcelSheetsView";
import { ExcelFormulasView } from "./ExcelFormulasView";
import { ExcelDataFieldsView } from "./ExcelDataFieldsView";
import { ExcelFilterSortView } from "./ExcelFilterSortView";
import { PptSlideNavigatorView } from "./PptSlideNavigatorView";
import { PptLayersView } from "./PptLayersView";
import { PptLayoutsView } from "./PptLayoutsView";
import { PptNotesView } from "./PptNotesView";
import { Settings2 } from "lucide-react";
import type { ViewPreferencesController } from "../../views/useViewPreferences";
import { CustomViewContent } from "../../views/CustomViewContent";

interface PrimarySidebarProps {
  viewPreferences?: ViewPreferencesController;
  onCustomizeViews?: () => void;
  activeDoc: OfficeDocument | null;
  tabs: EditorTab[];
  activeTabId: string;
  files: { id: string; name: string; type: DocumentType }[];
  onSelectTab: (tabId: string) => void;
  onCloseTab: (tabId: string) => void;
  onOpenFile: (fileId: string) => void;
  onUploadFile: (file: File) => void;
  // Word actions
  onSelectHeading?: (headingId: string) => void;
  onToggleCommentResolve?: (commentId: string) => void;
  // Excel actions
  onSelectExcelSheet?: (sheetId: string) => void;
  onAddExcelSheet?: () => void;
  onSelectFormulaCell?: (coord: string) => void;
  onSortExcelColumn?: (colKey: string, dir: "asc" | "desc") => void;
  // PPT actions
  onSelectPptSlide?: (slideId: string) => void;
  onDuplicatePptSlide?: (slideId: string) => void;
  onDeletePptSlide?: (slideId: string) => void;
  selectedPptObjectId?: string | null;
  onSelectPptObject?: (id: string) => void;
  onTogglePptObjectVisibility?: (id: string) => void;
  onChangePptLayout?: (layout: any) => void;
  onChangePptNotes?: (notes: string) => void;
}

export const PrimarySidebar: React.FC<PrimarySidebarProps> = ({
  activeDoc,
  tabs,
  activeTabId,
  files,
  onSelectTab,
  onCloseTab,
  onOpenFile,
  onUploadFile,
  onSelectHeading,
  onToggleCommentResolve,
  onSelectExcelSheet,
  onAddExcelSheet,
  onSelectFormulaCell,
  onSortExcelColumn,
  onSelectPptSlide,
  onDuplicatePptSlide,
  onDeletePptSlide,
  selectedPptObjectId,
  onSelectPptObject,
  onTogglePptObjectVisibility,
  onChangePptLayout,
  onChangePptNotes,
  viewPreferences,
  onCustomizeViews,
}) => {
  const getSidebarHeaderTitle = () => {
    if (!activeDoc) return "EXPLORER";
    switch (activeDoc.type) {
      case "word":
        return "WORD EXPLORER & OUTLINE";
      case "excel":
        return "EXCEL WORKBOOK & AUDITOR";
      case "powerpoint":
        return "SLIDE DECK & OBJECTS";
      case "code":
        return "PROJECT EXPLORER";
    }
  };

  const currentPptSlide =
    activeDoc?.type === "powerpoint"
      ? activeDoc.data.slides.find(
          (s) => s.id === activeDoc.data.activeSlideId,
        ) || activeDoc.data.slides[0]
      : null;

  const sections: React.ReactElement<{ id: string; onHide?: () => void }>[] =
    [];
  sections.push(
    <AccordionSection
      id="open-editors"
      title="OPEN EDITORS"
      badge={tabs.length}
      defaultExpanded={true}
    >
      <OpenEditorsView
        tabs={tabs}
        activeTabId={activeTabId}
        onSelectTab={onSelectTab}
        onCloseTab={onCloseTab}
      />
    </AccordionSection>,
  );
  sections.push(
    <AccordionSection
      id="workspace-files"
      title="WORKSPACE FILES"
      badge={files.length}
      defaultExpanded={true}
    >
      <FileExplorerView
        files={files}
        activeFileId={activeTabId}
        onOpenFile={onOpenFile}
        onUploadFile={onUploadFile}
      />
    </AccordionSection>,
  );
  if (activeDoc?.type === "word")
    sections.push(
      <AccordionSection
        id="word-outline"
        title="DOCUMENT OUTLINE"
        badge={activeDoc.data.headings.length}
        defaultExpanded={true}
      >
        <WordOutlineView
          headings={activeDoc.data.headings}
          onSelectHeading={(id) => onSelectHeading?.(id)}
        />
      </AccordionSection>,
    );
  if (activeDoc?.type === "word")
    sections.push(
      <AccordionSection
        id="word-pages"
        title="PAGES & SECTIONS"
        badge={activeDoc.data.sections.length}
        defaultExpanded={true}
      >
        <WordPagesView
          sections={activeDoc.data.sections}
          onSelectSection={() =>
            onSelectHeading?.(activeDoc.data.paragraphs[0]?.id ?? "")
          }
        />
      </AccordionSection>,
    );
  if (activeDoc?.type === "word")
    sections.push(
      <AccordionSection
        id="word-comments"
        title="REVIEW COMMENTS"
        badge={activeDoc.data.comments.length}
        defaultExpanded={true}
      >
        <WordCommentsView
          comments={activeDoc.data.comments}
          onToggleResolve={(id) => onToggleCommentResolve?.(id)}
        />
      </AccordionSection>,
    );
  if (activeDoc?.type === "word")
    sections.push(
      <AccordionSection
        id="word-stats"
        title="DOCUMENT STATISTICS"
        defaultExpanded={true}
      >
        <WordStatsView stats={activeDoc.data.stats} />
      </AccordionSection>,
    );
  if (activeDoc?.type === "excel")
    sections.push(
      <AccordionSection
        id="excel-sheets"
        title="WORKSHEETS"
        badge={activeDoc.data.sheets.length}
        defaultExpanded={true}
      >
        <ExcelSheetsView
          sheets={activeDoc.data.sheets}
          activeSheetId={activeDoc.data.activeSheetId}
          onSelectSheet={(id) => onSelectExcelSheet?.(id)}
          onAddSheet={onAddExcelSheet}
        />
      </AccordionSection>,
    );
  if (activeDoc?.type === "excel")
    sections.push(
      <AccordionSection
        id="excel-formulas"
        title="FORMULAS & FUNCTIONS"
        badge={activeDoc.data.formulasAudit.length}
        defaultExpanded={true}
      >
        <ExcelFormulasView
          formulas={activeDoc.data.formulasAudit}
          onSelectFormulaCell={onSelectFormulaCell}
        />
      </AccordionSection>,
    );
  if (activeDoc?.type === "excel")
    sections.push(
      <AccordionSection
        id="excel-fields"
        title="DATA COLUMNS & FIELDS"
        badge={
          activeDoc.data.sheets.find(
            (s) => s.id === activeDoc.data.activeSheetId,
          )?.columns.length || 0
        }
        defaultExpanded={true}
      >
        <ExcelDataFieldsView
          columns={
            activeDoc.data.sheets.find(
              (s) => s.id === activeDoc.data.activeSheetId,
            )?.columns || []
          }
        />
      </AccordionSection>,
    );
  if (activeDoc?.type === "excel")
    sections.push(
      <AccordionSection
        id="excel-filter-sort"
        title="SORT ROWS"
        defaultExpanded={true}
      >
        <ExcelFilterSortView
          columns={
            activeDoc.data.sheets.find(
              (s) => s.id === activeDoc.data.activeSheetId,
            )?.columns || []
          }
          onSort={onSortExcelColumn}
        />
      </AccordionSection>,
    );
  if (activeDoc?.type === "powerpoint")
    sections.push(
      <AccordionSection
        id="ppt-slides"
        title="SLIDE NAVIGATOR"
        badge={activeDoc.data.slides.length}
        defaultExpanded={true}
      >
        <PptSlideNavigatorView
          slides={activeDoc.data.slides}
          activeSlideId={activeDoc.data.activeSlideId}
          onSelectSlide={(id) => onSelectPptSlide?.(id)}
          onDuplicateSlide={onDuplicatePptSlide}
          onDeleteSlide={onDeletePptSlide}
        />
      </AccordionSection>,
    );
  if (activeDoc?.type === "powerpoint")
    sections.push(
      <AccordionSection
        id="ppt-layers"
        title="OBJECTS & LAYERS"
        badge={currentPptSlide?.objects.length || 0}
        defaultExpanded={true}
      >
        <PptLayersView
          objects={currentPptSlide?.objects || []}
          selectedObjectId={selectedPptObjectId || null}
          onSelectObject={(id) => onSelectPptObject?.(id)}
          onToggleVisibility={onTogglePptObjectVisibility}
        />
      </AccordionSection>,
    );
  if (activeDoc?.type === "powerpoint")
    sections.push(
      <AccordionSection
        id="ppt-layouts"
        title="LAYOUTS & THEMES"
        defaultExpanded={true}
      >
        <PptLayoutsView
          currentLayout={currentPptSlide?.layout || "content"}
          themeName={activeDoc.data.themeName}
          onChangeLayout={onChangePptLayout}
        />
      </AccordionSection>,
    );
  if (activeDoc?.type === "powerpoint")
    sections.push(
      <AccordionSection
        id="ppt-notes"
        title="SPEAKER NOTES"
        defaultExpanded={true}
      >
        <PptNotesView
          notes={currentPptSlide?.notes || ""}
          onChangeNotes={onChangePptNotes}
        />
      </AccordionSection>,
    );
  if (activeDoc?.type === "code")
    sections.push(
      <AccordionSection
        id="code-symbols"
        title="OUTLINE & SYMBOLS"
        badge={activeDoc.data.symbols.length}
        defaultExpanded={true}
      >
        <div style={{ padding: "4px 0" }}>
          {activeDoc.data.symbols.map((sym, i) => (
            <div
              key={i}
              className="vscode-tree-item"
              style={{ justifyContent: "space-between" }}
            >
              <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                <span
                  style={{
                    fontSize: 10,
                    color: "#38bdf8",
                    fontWeight: "bold",
                  }}
                >
                  {sym.kind.charAt(0).toUpperCase()}
                </span>
                <span>{sym.name}</span>
              </div>
              <span
                style={{
                  fontSize: 10,
                  color: "var(--vscode-tab-inactive-fg)",
                }}
              >
                :{sym.line}
              </span>
            </div>
          ))}
        </div>
      </AccordionSection>,
    );
  const registered = new Map(
    sections.map((section) => [section.props.id, section]),
  );
  const visible = viewPreferences?.preferences.left ?? [...registered.keys()];
  return (
    <div className="vscode-sidebar">
      <div className="vscode-sidebar-title-strip">
        <span>{getSidebarHeaderTitle()}</span>
        {onCustomizeViews && (
          <button
            className="vscode-icon-btn"
            aria-label="Customize explorer views"
            title="Customize views"
            onClick={onCustomizeViews}
          >
            <Settings2 size={14} />
          </button>
        )}
      </div>
      <div className="vscode-sidebar-content">
        {visible.map((id) => {
          const custom = viewPreferences?.preferences.custom.find(
            (view) => view.id === id,
          );
          if (custom)
            return (
              <AccordionSection
                key={id}
                id={id}
                title={custom.title.toUpperCase()}
                onHide={() => viewPreferences?.setVisible("left", id, false)}
              >
                <CustomViewContent
                  view={custom}
                  activeDoc={activeDoc}
                  onChangeNotes={(notes) =>
                    viewPreferences?.updateNotes(id, notes)
                  }
                  storageError={viewPreferences?.storageError}
                />
              </AccordionSection>
            );
          const section = registered.get(id);
          return section
            ? React.cloneElement(section, {
                key: id,
                onHide: viewPreferences
                  ? () => viewPreferences.setVisible("left", id, false)
                  : undefined,
              })
            : null;
        })}
        {!visible.length && (
          <div className="views-empty">
            <p>No explorer views are shown.</p>
            <button className="vscode-btn-secondary" onClick={onCustomizeViews}>
              Add a view
            </button>
          </div>
        )}
      </div>
    </div>
  );
};
