import React, { useEffect, useRef, useState } from "react";
import {
  Save,
  Download,
  Plus,
  Upload,
  Undo2,
  Redo2,
  Circle,
  Check,
  LoaderCircle,
  X,
  Settings2,
} from "lucide-react";
import {
  ActivityBarTab,
  BottomPanelTab,
  RightViewTab,
  EditorTab,
  SearchMatch,
  CommandItem,
} from "./types/vscode";
import { OfficeDocument, PptSlide } from "./types/office";
import { useWorkspace } from "./hooks/useWorkspace";
import { createDocument, clone, documentExtension } from "./services/documents";
import { loadOfficeFile } from "./parsers/fileLoader";
import { exportDocument } from "./services/exportDocument";
import { replaceInDocument, SearchOptions } from "./services/search";
import { recalculateWorkbook } from "./services/formulas";
import { applySlideLayout } from "./services/slides";
import { defaultPluginHost } from "./plugins";
import { usePanelLayout } from "./hooks/usePanelLayout";
import { PanelResizer } from "./components/layout/PanelResizer";
import { useViewPreferences } from "./views/useViewPreferences";
import { ManageViewsDialog } from "./components/layout/ManageViewsDialog";
import { ExtensionsPanel } from "./components/plugins/ExtensionsPanel";
import { TitleBar } from "./components/layout/TitleBar";
import { ActivityBar } from "./components/layout/ActivityBar";
import { StatusBar } from "./components/layout/StatusBar";
import { CommandPalette } from "./components/layout/CommandPalette";
import { WorkspaceDialog } from "./components/layout/WorkspaceDialog";
import { PrimarySidebar } from "./components/sidebar/PrimarySidebar";
import { HistoryPanel } from "./components/sidebar/HistoryPanel";
import { EditorContainer } from "./components/editor/EditorContainer";
import { BottomPanel } from "./components/bottomPanel/BottomPanel";
import { SecondarySidebar } from "./components/rightView/SecondarySidebar";
import { PresentationModal } from "./components/editor/PresentationModal";
import "./styles/vscode.css";
import "./styles/workspace.css";
import "./styles/panels.css";

export interface VSCodeOfficeStudioProps {
  initialTheme?: "dark" | "light";
  defaultActiveFile?: OfficeDocument["type"];
}
export const VSCodeOfficeStudio: React.FC<VSCodeOfficeStudioProps> = ({
  initialTheme = "dark",
  defaultActiveFile = "word",
}) => {
  const [theme, setTheme] = useState<"dark" | "light">(() => {
    try {
      return (
        (localStorage.getItem("office.theme") as "dark" | "light") ||
        initialTheme
      );
    } catch {
      return initialTheme;
    }
  });
  const panelLayout = usePanelLayout();
  const isSidebarOpen = panelLayout.showPrimary,
    setIsSidebarOpen = panelLayout.setPrimaryOpen;
  const isRightViewOpen = panelLayout.showSecondary,
    setIsRightViewOpen = panelLayout.setSecondaryOpen;
  const isBottomPanelOpen = panelLayout.preferences.bottomOpen,
    setIsBottomPanelOpen = panelLayout.setBottomOpen;
  const [isBottomMaximized, setIsBottomMaximized] = useState(false);
  const [isCommandPaletteOpen, setIsCommandPaletteOpen] = useState(false),
    [isSlideshowRunning, setIsSlideshowRunning] = useState(false);
  const [activeActivityTab, setActiveActivityTab] =
      useState<ActivityBarTab>("explorer"),
    [activeRightTab, setActiveRightTab] = useState<RightViewTab>("inspector");
  const [targetWordHeadingId, setTargetWordHeadingId] = useState<string | null>(
      null,
    ),
    [targetExcelCoord, setTargetExcelCoord] = useState<string | null>(null),
    [selectedPptObjectId, setSelectedPptObjectId] = useState<string | null>(
      null,
    ),
    [selectedCell, setSelectedCell] = useState("A1");
  const [notice, setNotice] = useState<{
    message: string;
    error: boolean;
  } | null>(null);
  const notify = (message: string, error = false) => {
    setNotice({ message, error });
  };
  const workspace = useWorkspace(defaultActiveFile, notify);
  const entry = workspace.entries[workspace.activeId],
    activeDoc = entry?.document ?? null;
  const viewPreferences = useViewPreferences(activeDoc?.type ?? "code");
  const [autosave, setAutosave] = useState(() => {
    try {
      return localStorage.getItem("office.autosave") === "true";
    } catch {
      return false;
    }
  });
  const [dialog, setDialog] = useState<
      | "new"
      | "copy"
      | "rename"
      | "close"
      | "replace"
      | "reload"
      | "views"
      | null
    >(null),
    [name, setName] = useState("Untitled.docx"),
    [newType, setNewType] = useState<OfficeDocument["type"]>("word"),
    [closeId, setCloseId] = useState("");
  const [pendingReplace, setPendingReplace] = useState<{
    query: string;
    replacement: string;
    options: SearchOptions;
  } | null>(null);
  const filePickerRef = useRef<HTMLInputElement>(null);
  const [, refreshPlugins] = useState(0);
  useEffect(
    () => defaultPluginHost.subscribe(() => refreshPlugins((v) => v + 1)),
    [],
  );
  useEffect(() => {
    document.body.classList.toggle("theme-light", theme === "light");
    try {
      localStorage.setItem("office.theme", theme);
    } catch {}
  }, [theme]);
  useEffect(() => {
    try {
      localStorage.setItem("office.autosave", String(autosave));
    } catch {}
  }, [autosave]);
  useEffect(() => {
    if (!notice || notice.error) return;
    const timer = setTimeout(() => setNotice(null), 5000);
    return () => clearTimeout(timer);
  }, [notice]);
  useEffect(() => {
    if (!autosave || !workspace.ready || !workspace.connected) return;
    const timer = setTimeout(() => {
      for (const e of Object.values(workspace.entries))
        if (workspace.dirty(e) && !e.saving && !e.error)
          void workspace.save(e.id);
    }, 1500);
    return () => clearTimeout(timer);
  }, [workspace.entries, autosave, workspace.ready, workspace.connected]);
  useEffect(() => {
    setTargetWordHeadingId(null);
    setTargetExcelCoord(null);
    setSelectedPptObjectId(null);
    setSelectedCell("A1");
  }, [workspace.activeId]);
  const update = (doc: OfficeDocument) =>
    workspace.update(
      doc.type === "excel"
        ? { type: "excel", data: recalculateWorkbook(doc.data) }
        : doc,
    );
  const tabs: EditorTab[] = workspace.openIds
    .filter((id) => workspace.entries[id])
    .map((id) => {
      const e = workspace.entries[id];
      return {
        id,
        title: e.name,
        filePath: `workspace/${e.name}`,
        docType: e.document.type,
        isDirty: workspace.dirty(e),
      };
    });
  const files = Object.values(workspace.entries).map((e) => ({
    id: e.id,
    name: e.name,
    type: e.document.type,
  }));
  const docsForSearch = Object.values(workspace.entries).map((e) => ({
    id: e.id,
    name: e.name,
    type: e.document.type,
    doc: e.document,
  }));
  const openBottom = (_tab: BottomPanelTab = "search") => {
    setIsBottomPanelOpen(true);
  };
  const customizeViews = () => setDialog("views");
  const resetPanelSizes = () => {
    panelLayout.resetSizes();
    setIsBottomMaximized(false);
  };
  const primaryResizer = (
    <PanelResizer
      label="Resize primary sidebar"
      orientation="vertical"
      controls="primary-sidebar-pane"
      value={panelLayout.primaryWidth}
      {...panelLayout.bounds.primaryWidth}
      onResize={(value) => panelLayout.resize("primaryWidth", value)}
      onReset={() => panelLayout.resetSize("primaryWidth")}
    />
  );
  const secondaryResizer = (
    <PanelResizer
      label="Resize secondary sidebar"
      orientation="vertical"
      direction={-1}
      controls="secondary-sidebar-pane"
      value={panelLayout.secondaryWidth}
      {...panelLayout.bounds.secondaryWidth}
      onResize={(value) => panelLayout.resize("secondaryWidth", value)}
      onReset={() => panelLayout.resetSize("secondaryWidth")}
    />
  );
  const bottomHeight = isBottomMaximized
    ? panelLayout.bounds.bottomHeight.max
    : panelLayout.bottomHeight;
  const showExtensions = () => {
    setActiveActivityTab("extensions");
    setIsSidebarOpen(true);
  };
  const showNew = () => {
    setName("Untitled.docx");
    setNewType("word");
    setDialog("new");
  };
  const showCopy = () => {
    if (entry) {
      setName(entry.name.replace(/(\.[^.]+)?$/, "-copy$1"));
      setDialog("copy");
    }
  };
  const showRename = () => {
    if (entry) {
      setName(entry.name);
      setDialog("rename");
    }
  };
  const closeTab = (id: string) => {
    if (workspace.entries[id] && workspace.dirty(workspace.entries[id])) {
      setCloseId(id);
      setDialog("close");
    } else workspace.close(id);
  };
  const saveAll = async () => {
    for (const e of Object.values(workspace.entries))
      if (workspace.dirty(e)) await workspace.save(e.id);
  };
  const download = async () => {
    if (!entry) return;
    try {
      await exportDocument(entry.document, entry.name);
      notify(
        `Downloaded ${entry.document.type === "code" ? entry.name : entry.name.replace(/\.[^.]+$/, "") + "." + documentExtension(entry.document)}. Workspace changes are saved separately.`,
      );
    } catch (e) {
      notify((e as Error).message, true);
    }
  };
  const startPresentation = () => {
    if (activeDoc?.type === "powerpoint") setIsSlideshowRunning(true);
    else notify("Open a presentation to start its slideshow.");
  };
  const upload = async (file: File) => {
    try {
      let doc = await loadOfficeFile(file);
      if (doc.type === "excel")
        doc = { type: "excel", data: recalculateWorkbook(doc.data) };
      workspace.add(doc, file.name);
      setActiveActivityTab("explorer");
      notify(
        `Opened ${file.name}${doc.type !== "code" ? ". Basic content imported; advanced Office formatting and embedded media may be omitted." : ""}`,
      );
    } catch (e) {
      notify((e as Error).message, true);
    }
  };
  const navigate = (match: SearchMatch) => {
    const e = workspace.entries[match.fileId];
    if (!e) return;
    workspace.open(e.id);
    if (e.document.type === "excel" && match.targetRef?.sheetId)
      workspace.update(
        {
          type: "excel",
          data: { ...e.document.data, activeSheetId: match.targetRef.sheetId },
        },
        e.id,
      );
    if (e.document.type === "powerpoint" && match.targetRef?.slideIndex) {
      const slide = e.document.data.slides[match.targetRef.slideIndex - 1];
      if (slide)
        workspace.update(
          {
            type: "powerpoint",
            data: { ...e.document.data, activeSlideId: slide.id },
          },
          e.id,
        );
    }
    setTimeout(() => {
      setTargetWordHeadingId(match.targetRef?.headingId ?? null);
      setTargetExcelCoord(match.targetRef?.cellCoord ?? null);
      setSelectedPptObjectId(match.targetRef?.elementId ?? null);
      if (match.targetRef?.lineNumber) {
        const textarea =
          document.querySelector<HTMLTextAreaElement>(".code-editor-input");
        if (textarea) {
          const offset =
            textarea.value
              .split("\n")
              .slice(0, match.targetRef.lineNumber - 1)
              .join("\n").length + (match.targetRef.lineNumber > 1 ? 1 : 0);
          textarea.focus();
          textarea.setSelectionRange(
            offset + match.matchIndex,
            offset + match.matchIndex + match.matchLength,
          );
          textarea.scrollTop = (match.targetRef.lineNumber - 1) * 21;
        }
      }
    }, 0);
  };
  const changeSlide = (fn: (slide: PptSlide) => PptSlide) => {
    if (activeDoc?.type !== "powerpoint") return;
    update({
      type: "powerpoint",
      data: {
        ...activeDoc.data,
        slides: activeDoc.data.slides.map((s) =>
          s.id === activeDoc.data.activeSlideId ? fn(s) : s,
        ),
      },
    });
  };
  const duplicateSlide = (id: string) => {
    if (activeDoc?.type !== "powerpoint") return;
    const index = activeDoc.data.slides.findIndex((s) => s.id === id),
      source = activeDoc.data.slides[index];
    if (!source) return;
    const copied = {
      ...clone(source),
      id: crypto.randomUUID(),
      objects: source.objects.map((o) => ({
        ...clone(o),
        id: crypto.randomUUID(),
      })),
    };
    const slides = [...activeDoc.data.slides];
    slides.splice(index + 1, 0, copied);
    update({
      type: "powerpoint",
      data: {
        ...activeDoc.data,
        activeSlideId: copied.id,
        slides: slides.map((s, i) => ({ ...s, slideNumber: i + 1 })),
      },
    });
  };
  const deleteSlide = (id: string) => {
    if (activeDoc?.type !== "powerpoint") return;
    if (activeDoc.data.slides.length === 1) {
      notify("A presentation must contain at least one slide.");
      return;
    }
    const slides = activeDoc.data.slides
      .filter((s) => s.id !== id)
      .map((s, i) => ({ ...s, slideNumber: i + 1 }));
    update({
      type: "powerpoint",
      data: {
        ...activeDoc.data,
        slides,
        activeSlideId:
          activeDoc.data.activeSlideId === id
            ? slides[0].id
            : activeDoc.data.activeSlideId,
      },
    });
  };
  const sort = (col: string, dir: "asc" | "desc") => {
    if (activeDoc?.type !== "excel") return;
    const sheet = activeDoc.data.sheets.find(
      (s) => s.id === activeDoc.data.activeSheetId,
    )!;
    if (Object.values(sheet.cells).some((c) => c.formula)) {
      notify(
        "Sorting a sheet with formulas is unavailable because relative references would need rebasing. Copy values into a separate sheet first.",
      );
      return;
    }
    const rows = Array.from(
      { length: sheet.rowCount - 1 },
      (_, i) => i + 2,
    ).sort((a, b) => {
      const av = sheet.cells[`${col}${a}`]?.value ?? "",
        bv = sheet.cells[`${col}${b}`]?.value ?? "";
      if (av === "" || bv === "") return av === "" ? (bv === "" ? 0 : 1) : -1;
      return (
        (typeof av === "number" && typeof bv === "number"
          ? av - bv
          : String(av).localeCompare(String(bv), undefined, {
              numeric: true,
            })) * (dir === "asc" ? 1 : -1)
      );
    });
    const cells = { ...sheet.cells };
    rows.forEach((source, i) =>
      sheet.columns.forEach((c) => {
        const value = sheet.cells[`${c.key}${source}`];
        if (value) cells[`${c.key}${i + 2}`] = clone(value);
        else delete cells[`${c.key}${i + 2}`];
      }),
    );
    update({
      type: "excel",
      data: {
        ...activeDoc.data,
        sheets: activeDoc.data.sheets.map((s) =>
          s.id === sheet.id ? { ...s, cells } : s,
        ),
      },
    });
    notify(`Sorted ${sheet.name} by ${col}; row 1 retained as header.`);
  };
  const commands: CommandItem[] = [
    {
      id: "new",
      title: "File: New document, workbook, presentation or code",
      shortcut: "Ctrl+N",
      action: showNew,
    },
    {
      id: "open",
      title: "File: Open from computer",
      shortcut: "Ctrl+O",
      action: () => filePickerRef.current?.click(),
    },
    {
      id: "save",
      title: "File: Save to workspace",
      shortcut: "Ctrl+S",
      action: () => void workspace.save(),
    },
    {
      id: "saveall",
      title: "File: Save all",
      shortcut: "Ctrl+Shift+S",
      action: () => void saveAll(),
    },
    {
      id: "download",
      title: "File: Download native file",
      action: () => void download(),
    },
    { id: "copy", title: "File: Save a copy", action: showCopy },
    { id: "rename", title: "File: Rename", action: showRename },
    {
      id: "undo",
      title: "Edit: Undo",
      shortcut: "Ctrl+Z",
      action: () => workspace.undo(),
    },
    {
      id: "redo",
      title: "Edit: Redo",
      shortcut: "Ctrl+Y",
      action: () => workspace.undo(true),
    },
    {
      id: "sidebar",
      title: "View: Toggle sidebar",
      shortcut: "Ctrl+B",
      action: () => setIsSidebarOpen((v) => !v),
    },
    {
      id: "panel",
      title: "View: Toggle bottom panel",
      shortcut: "Ctrl+J",
      action: () => setIsBottomPanelOpen((v) => !v),
    },
    {
      id: "inspector",
      title: "View: Toggle inspector",
      shortcut: "Ctrl+Alt+B",
      action: () => setIsRightViewOpen((v) => !v),
    },
    {
      id: "search",
      title: "Search: Find and replace in workspace",
      shortcut: "Ctrl+Shift+F",
      action: () => openBottom("search"),
    },
    {
      id: "customize-views",
      title: "View: Customize views",
      action: customizeViews,
    },
    {
      id: "reset-panels",
      title: "View: Reset panel sizes",
      action: resetPanelSizes,
    },
    {
      id: "plugins",
      title: "Extensions: Manage plugins and MCP",
      shortcut: "Ctrl+Shift+X",
      action: showExtensions,
    },
    {
      id: "theme",
      title: "Preferences: Toggle light / dark theme",
      action: () => setTheme((t) => (t === "dark" ? "light" : "dark")),
    },
    {
      id: "slideshow",
      title: "Presentation: Start slideshow",
      shortcut: "F5",
      action: startPresentation,
    },
    ...files.map((f) => ({
      id: `open-${f.id}`,
      title: `Open: ${f.name}`,
      category: "Files",
      action: () => workspace.open(f.id),
    })),
    ...defaultPluginHost.commands(activeDoc).map((c) => ({
      id: c.id,
      title: `Plugin: ${c.title}`,
      category: "Extensions",
      action: () => {
        if (!activeDoc) return;
        try {
          const result = defaultPluginHost.execute(c.id, activeDoc);
          if (result.document) update(result.document);
          notify(result.message);
          if (result.context) {
            setIsRightViewOpen(true);
            viewPreferences.setVisible("right", "copilot", true);
            setActiveRightTab("copilot");
          }
        } catch (e) {
          notify((e as Error).message, true);
        }
      },
    })),
  ];
  useEffect(() => {
    const handler = (e: KeyboardEvent) => {
      const mod = e.ctrlKey || e.metaKey,
        key = e.key.toLowerCase();
      if (dialog || e.defaultPrevented) return;
      if (mod && key === "s") {
        e.preventDefault();
        if (e.shiftKey) void saveAll();
        else void workspace.save();
      } else if (mod && key === "o") {
        e.preventDefault();
        filePickerRef.current?.click();
      } else if (mod && key === "n") {
        e.preventDefault();
        showNew();
      } else if (mod && key === "w") {
        e.preventDefault();
        closeTab(workspace.activeId);
      } else if ((mod && key === "p" && e.shiftKey) || e.key === "F1") {
        e.preventDefault();
        setIsCommandPaletteOpen(true);
      } else if (mod && key === "p") {
        e.preventDefault();
        window.print();
      } else if (mod && key === "b" && e.altKey) {
        e.preventDefault();
        setIsRightViewOpen((v) => !v);
      } else if (
        mod &&
        key === "b" &&
        !(e.target as HTMLElement).isContentEditable
      ) {
        e.preventDefault();
        setIsSidebarOpen((v) => !v);
      } else if (mod && key === "j") {
        e.preventDefault();
        setIsBottomPanelOpen((v) => !v);
      } else if (mod && e.shiftKey && key === "f") {
        e.preventDefault();
        openBottom("search");
      } else if (mod && e.shiftKey && key === "x") {
        e.preventDefault();
        showExtensions();
      } else if (
        mod &&
        (key === "z" || key === "y") &&
        ((e.target as HTMLElement).closest(".vscode-editor-canvas") ||
          e.target === document.body)
      ) {
        e.preventDefault();
        workspace.undo(key === "y" || e.shiftKey);
      } else if (e.key === "F5") {
        e.preventDefault();
        startPresentation();
      }
    };
    window.addEventListener("keydown", handler);
    return () => window.removeEventListener("keydown", handler);
  });
  const fileForm = () => {
    if (!name.trim() || /[\\/:*?"<>|]/.test(name) || name.length > 160) {
      notify(
        "Use a filename of 1–160 characters without path separators or reserved characters.",
        true,
      );
      return;
    }
    if (dialog === "new") {
      const ext = {
        word: "docx",
        excel: "xlsx",
        powerpoint: "pptx",
        code: "md",
      }[newType];
      const fileName = name.includes(".")
        ? name.trim()
        : `${name.trim()}.${ext}`;
      if (
        newType === "code" &&
        !/\.(md|txt|ts|tsx|js|jsx|py|json|css|html|xml|yaml|yml)$/i.test(
          fileName,
        )
      ) {
        notify("Choose a supported text or code extension.", true);
        return;
      }
      if (newType !== "code" && !fileName.toLowerCase().endsWith(`.${ext}`)) {
        notify(`Use .${ext} for this document type.`, true);
        return;
      }
      workspace.add(createDocument(newType, fileName), fileName);
    } else if (entry) {
      const ext = entry.name.split(".").pop();
      if (name.split(".").pop() !== ext) {
        notify("Keep the current extension when renaming or copying.", true);
        return;
      }
      if (dialog === "rename") workspace.rename(entry.id, name.trim());
      else {
        const id = workspace.add(clone(entry.document), name.trim());
        void workspace.save(id);
      }
    }
    setDialog(null);
  };
  return (
    <div
      className={`vscode-app-container ${theme === "light" ? "theme-light" : ""}`}
      onDragOver={(e) => {
        if (e.dataTransfer.types.includes("Files")) e.preventDefault();
      }}
      onDrop={(e) => {
        e.preventDefault();
        for (const file of Array.from(e.dataTransfer.files)) void upload(file);
      }}
    >
      <input
        ref={filePickerRef}
        type="file"
        multiple
        accept=".docx,.xlsx,.xls,.pptx,.csv,.txt,.md,.json,.ts,.js,.py,.html,.css,.yaml,.yml"
        style={{ display: "none" }}
        onChange={(e) => {
          for (const file of Array.from(e.target.files ?? []))
            void upload(file);
          e.target.value = "";
        }}
      />
      <TitleBar
        title={entry?.name ?? "Code Office workspace"}
        onOpenCommandPalette={() => setIsCommandPaletteOpen(true)}
        isSidebarOpen={isSidebarOpen}
        onToggleSidebar={() => setIsSidebarOpen((v) => !v)}
        isBottomPanelOpen={isBottomPanelOpen}
        onToggleBottomPanel={() => setIsBottomPanelOpen((v) => !v)}
        isRightViewOpen={isRightViewOpen}
        onToggleRightView={() => setIsRightViewOpen((v) => !v)}
        onOpenFilePicker={() => filePickerRef.current?.click()}
        onOpenSampleWord={() => {}}
        onOpenSampleExcel={() => {}}
        onOpenSamplePpt={() => {}}
        onOpenBottomTab={openBottom}
        onTriggerSlideshow={startPresentation}
        onToggleTheme={() => setTheme((t) => (t === "dark" ? "light" : "dark"))}
        onNew={showNew}
        onSave={() => void workspace.save()}
        onSaveAll={() => void saveAll()}
        onSaveCopy={showCopy}
        onRename={showRename}
        onExport={() => void download()}
        onUndo={() => workspace.undo()}
        onRedo={() => workspace.undo(true)}
        onExtensions={showExtensions}
        onPrint={() => window.print()}
        onCustomizeViews={customizeViews}
        onResetLayout={resetPanelSizes}
      />
      <div
        className={`vscode-main-workspace ${panelLayout.compact ? "compact-panels" : "docked-panels"}`}
        ref={panelLayout.workspaceRef}
      >
        <ActivityBar
          activeTab={activeActivityTab}
          documentType={activeDoc?.type}
          onSelectTab={(tab) => {
            setActiveActivityTab(tab);
            setIsSidebarOpen(true);
          }}
        />
        {isSidebarOpen && (
          <div
            id="primary-sidebar-pane"
            className="workspace-side-pane primary-pane"
            style={{ width: panelLayout.primaryWidth }}
          >
            {activeActivityTab === "extensions" ? (
              <ExtensionsPanel
                activeDoc={activeDoc}
                onChangeDoc={update}
                onNotify={notify}
              />
            ) : activeActivityTab === "source-control" ? (
              <HistoryPanel
                key={entry?.id}
                entry={entry}
                onRestore={update}
                onNotify={notify}
              />
            ) : activeActivityTab === "settings" ? (
              <aside className="vscode-sidebar utility-sidebar">
                <div className="vscode-sidebar-title-strip">
                  <Settings2 size={14} /> SETTINGS
                </div>
                <div className="utility-body">
                  <h3>Your workspace</h3>
                  <button
                    className="vscode-btn-secondary"
                    onClick={customizeViews}
                  >
                    Customize views
                  </button>
                  <button
                    className="vscode-btn-secondary"
                    onClick={resetPanelSizes}
                  >
                    Reset panel sizes
                  </button>
                  <p>
                    Drag the dividers to resize panels. Arrow keys resize a
                    focused divider; double-click resets its size.
                  </p>
                  {(panelLayout.storageError ||
                    viewPreferences.storageError) && (
                    <p role="alert">
                      Browser preferences could not be saved. This layout works
                      until the page closes.
                    </p>
                  )}
                  <label>
                    Color theme
                    <select
                      className="vscode-input"
                      value={theme}
                      onChange={(e) =>
                        setTheme(e.target.value as "dark" | "light")
                      }
                    >
                      <option value="dark">Dark</option>
                      <option value="light">Light</option>
                    </select>
                  </label>
                  <label className="checkbox-label">
                    <input
                      type="checkbox"
                      checked={autosave}
                      onChange={(e) => setAutosave(e.target.checked)}
                    />
                    Auto Save to local server
                  </label>
                  <p>
                    Changes are saved after 1.5 seconds of inactivity. Browser
                    recovery drafts are kept independently.
                  </p>
                  <p>
                    {workspace.connected
                      ? "Local storage connected."
                      : "Storage unavailable. Run npm run dev to enable disk saves."}
                  </p>
                  <button
                    className="vscode-btn-secondary"
                    onClick={() =>
                      void workspace
                        .refresh()
                        .then((ok) =>
                          notify(
                            ok ? "Storage connected." : "Storage unavailable.",
                            !ok,
                          ),
                        )
                    }
                  >
                    Reconnect storage
                  </button>
                  <button
                    className="vscode-btn-secondary"
                    onClick={showExtensions}
                  >
                    Manage extensions & MCP
                  </button>
                </div>
              </aside>
            ) : (
              <PrimarySidebar
                viewPreferences={viewPreferences}
                onCustomizeViews={customizeViews}
                activeDoc={activeDoc}
                tabs={tabs}
                activeTabId={workspace.activeId}
                files={files}
                onSelectTab={workspace.open}
                onCloseTab={closeTab}
                onOpenFile={workspace.open}
                onUploadFile={upload}
                onSelectHeading={setTargetWordHeadingId}
                onToggleCommentResolve={(id) => {
                  if (activeDoc?.type === "word")
                    update({
                      type: "word",
                      data: {
                        ...activeDoc.data,
                        comments: activeDoc.data.comments.map((c) =>
                          c.id === id ? { ...c, resolved: !c.resolved } : c,
                        ),
                      },
                    });
                }}
                onSelectExcelSheet={(id) => {
                  if (activeDoc?.type === "excel")
                    update({
                      type: "excel",
                      data: { ...activeDoc.data, activeSheetId: id },
                    });
                }}
                onAddExcelSheet={() => {
                  if (activeDoc?.type === "excel") {
                    if (activeDoc.data.sheets.length >= 50) {
                      notify("A workbook supports up to 50 sheets.");
                      return;
                    }
                    const blank = createDocument("excel", "");
                    if (blank.type !== "excel") return;
                    const s = blank.data.sheets[0];
                    s.id = crypto.randomUUID();
                    let index = activeDoc.data.sheets.length + 1;
                    while (
                      activeDoc.data.sheets.some(
                        (sheet) => sheet.name.toLowerCase() === `sheet${index}`,
                      )
                    )
                      index++;
                    s.name = `Sheet${index}`;
                    update({
                      type: "excel",
                      data: {
                        ...activeDoc.data,
                        sheets: [...activeDoc.data.sheets, s],
                        activeSheetId: s.id,
                      },
                    });
                  }
                }}
                onSelectFormulaCell={setTargetExcelCoord}
                onSortExcelColumn={sort}
                onSelectPptSlide={(id) => {
                  if (activeDoc?.type === "powerpoint")
                    update({
                      type: "powerpoint",
                      data: { ...activeDoc.data, activeSlideId: id },
                    });
                }}
                onDuplicatePptSlide={duplicateSlide}
                onDeletePptSlide={deleteSlide}
                selectedPptObjectId={selectedPptObjectId}
                onSelectPptObject={setSelectedPptObjectId}
                onTogglePptObjectVisibility={(id) =>
                  changeSlide((s) => ({
                    ...s,
                    objects: s.objects.map((o) =>
                      o.id === id ? { ...o, visible: o.visible === false } : o,
                    ),
                  }))
                }
                onChangePptLayout={(layout) =>
                  changeSlide((s) => applySlideLayout(s, layout))
                }
                onChangePptNotes={(notes) =>
                  changeSlide((s) => ({ ...s, notes }))
                }
              />
            )}
            {panelLayout.compact && primaryResizer}
          </div>
        )}
        {isSidebarOpen && !panelLayout.compact && primaryResizer}
        <main className="workspace-center">
          <div className="workspace-actionbar">
            <div className="actionbar-group">
              <button
                className="vscode-btn-secondary"
                onClick={showNew}
                title="New file (Ctrl+N)"
              >
                <Plus size={14} />
                New
              </button>
              <button
                className="vscode-icon-btn"
                onClick={() => filePickerRef.current?.click()}
                title="Open file (Ctrl+O)"
              >
                <Upload size={15} />
              </button>
              <button
                className="vscode-icon-btn"
                title="Customize views"
                aria-label="Customize views"
                onClick={customizeViews}
              >
                <Settings2 size={15} />
              </button>
              <span className="toolbar-divider" />
              <button
                className="vscode-btn"
                disabled={!entry || entry.saving || !workspace.ready}
                onClick={() => void workspace.save()}
                title="Save to workspace (Ctrl+S)"
              >
                {entry?.saving ? (
                  <LoaderCircle className="spin" size={14} />
                ) : (
                  <Save size={14} />
                )}
                Save
              </button>
              <button
                className="vscode-btn-secondary"
                disabled={!entry}
                onClick={() => void download()}
                title="Download native file"
              >
                <Download size={14} />
                <span>Download</span>
              </button>
              <button
                className="vscode-icon-btn"
                onClick={() => workspace.undo()}
                title="Undo"
              >
                <Undo2 size={15} />
              </button>
              <button
                className="vscode-icon-btn"
                onClick={() => workspace.undo(true)}
                title="Redo"
              >
                <Redo2 size={15} />
              </button>
            </div>
            <div
              className={`save-state ${entry?.error ? "error" : ""}`}
              role="status"
            >
              {entry?.saving ? (
                <LoaderCircle className="spin" size={12} />
              ) : entry && workspace.dirty(entry) ? (
                <Circle size={10} />
              ) : (
                <Check size={13} />
              )}
              <span>
                {entry?.error
                  ? "Save failed — changes retained"
                  : !workspace.ready
                    ? "Connecting storage…"
                    : entry?.saving
                      ? "Saving…"
                      : entry && workspace.dirty(entry)
                        ? "Unsaved changes"
                        : entry
                          ? `Saved · r${entry.revision}`
                          : "Workspace ready"}
              </span>
            </div>
          </div>
          {entry?.error && (
            <div className="storage-banner" role="alert">
              {entry.error}
              <button onClick={() => void workspace.save()}>Retry</button>
              <button onClick={showCopy}>Save a copy</button>
              {entry.revision !== null && (
                <button onClick={() => setDialog("reload")}>
                  Reload saved
                </button>
              )}
            </div>
          )}
          {workspace.recoveryError && (
            <div className="storage-banner" role="alert">
              Browser recovery storage is full. Save or download your work
              before closing.
            </div>
          )}
          <EditorContainer
            tabs={tabs}
            activeTabId={workspace.activeId}
            activeDoc={activeDoc}
            onSelectTab={workspace.open}
            onCloseTab={closeTab}
            onChangeActiveDoc={update}
            targetWordHeadingId={targetWordHeadingId}
            targetExcelCoord={targetExcelCoord}
            selectedPptObjectId={selectedPptObjectId}
            onSelectPptObject={setSelectedPptObjectId}
            onSelectExcelCell={setSelectedCell}
          />
          {isBottomPanelOpen && (
            <>
              <PanelResizer
                label="Resize bottom panel"
                orientation="horizontal"
                direction={-1}
                controls="bottom-search-pane"
                value={bottomHeight}
                {...panelLayout.bounds.bottomHeight}
                onResize={(value) => {
                  setIsBottomMaximized(false);
                  panelLayout.resize("bottomHeight", value);
                }}
                onReset={() => {
                  setIsBottomMaximized(false);
                  panelLayout.resetSize("bottomHeight");
                }}
              />
              <div
                id="bottom-search-pane"
                className="workspace-bottom-pane"
                style={{ height: bottomHeight }}
              >
                <BottomPanel
                  height={bottomHeight}
                  isMaximized={isBottomMaximized}
                  onToggleMaximize={() => setIsBottomMaximized((v) => !v)}
                  onClose={() => setIsBottomPanelOpen(false)}
                  documents={docsForSearch}
                  onNavigateToMatch={navigate}
                  onReplaceAll={(query, replacement, options = {}) => {
                    setPendingReplace({ query, replacement, options });
                    setDialog("replace");
                  }}
                />
              </div>
            </>
          )}
        </main>
        {isRightViewOpen && !panelLayout.compact && secondaryResizer}
        {isRightViewOpen && (
          <div
            id="secondary-sidebar-pane"
            className="workspace-side-pane secondary-pane"
            style={{ width: panelLayout.secondaryWidth }}
          >
            {panelLayout.compact && secondaryResizer}
            <SecondarySidebar
              viewPreferences={viewPreferences}
              onCustomizeViews={customizeViews}
              activeTab={activeRightTab}
              onSelectTab={setActiveRightTab}
              onClose={() => setIsRightViewOpen(false)}
              activeDoc={activeDoc}
              onChangeDoc={update}
              selectedPptObjectId={selectedPptObjectId}
              selectedExcelCoord={selectedCell}
            />
          </div>
        )}
      </div>
      <StatusBar
        activeDoc={activeDoc}
        isSidebarOpen={isSidebarOpen}
        onToggleSidebar={() => setIsSidebarOpen((v) => !v)}
        isBottomPanelOpen={isBottomPanelOpen}
        onToggleBottomPanel={() => setIsBottomPanelOpen((v) => !v)}
        isRightViewOpen={isRightViewOpen}
        onToggleRightView={() => setIsRightViewOpen((v) => !v)}
      />
      {notice && (
        <div
          className={`workspace-toast ${notice.error ? "error" : ""}`}
          role={notice.error ? "alert" : "status"}
        >
          {notice.message}
          <button
            className="vscode-icon-btn"
            title="Dismiss notification"
            onClick={() => setNotice(null)}
          >
            <X size={14} />
          </button>
        </div>
      )}
      {dialog === "views" && (
        <ManageViewsDialog
          controller={viewPreferences}
          onClose={() => setDialog(null)}
        />
      )}
      <CommandPalette
        isOpen={isCommandPaletteOpen}
        onClose={() => setIsCommandPaletteOpen(false)}
        commands={commands}
      />
      {isSlideshowRunning && activeDoc?.type === "powerpoint" && (
        <PresentationModal
          document={activeDoc.data}
          onClose={() => setIsSlideshowRunning(false)}
        />
      )}
      {(dialog === "new" || dialog === "copy" || dialog === "rename") && (
        <WorkspaceDialog
          title={
            dialog === "new"
              ? "Create a file"
              : dialog === "copy"
                ? "Save a copy"
                : "Rename file"
          }
          onClose={() => setDialog(null)}
        >
          <form
            onSubmit={(e) => {
              e.preventDefault();
              fileForm();
            }}
          >
            {dialog === "new" && (
              <label>
                File type
                <select
                  className="vscode-input"
                  value={newType}
                  onChange={(e) => {
                    const type = e.target.value as OfficeDocument["type"];
                    setNewType(type);
                    setName(
                      `Untitled.${{ word: "docx", excel: "xlsx", powerpoint: "pptx", code: "md" }[type]}`,
                    );
                  }}
                >
                  <option value="word">Word document (.docx)</option>
                  <option value="excel">Excel workbook (.xlsx)</option>
                  <option value="powerpoint">
                    PowerPoint presentation (.pptx)
                  </option>
                  <option value="code">
                    Text / code (.md, .ts, .js, .py, .json)
                  </option>
                </select>
              </label>
            )}
            <label>
              Filename
              <input
                autoFocus
                className="vscode-input"
                aria-label="Filename"
                value={name}
                onChange={(e) => setName(e.target.value)}
              />
            </label>
            <p>
              {dialog === "copy"
                ? "The original file remains unchanged. The copy receives its own identity and saved history."
                : "Files live in this local workspace. Use Download to export a native file."}
            </p>
            <div className="dialog-actions">
              <button
                type="button"
                className="vscode-btn-secondary"
                onClick={() => setDialog(null)}
              >
                Cancel
              </button>
              <button type="submit" className="vscode-btn">
                {dialog === "new"
                  ? "Create file"
                  : dialog === "copy"
                    ? "Save copy"
                    : "Rename"}
              </button>
            </div>
          </form>
        </WorkspaceDialog>
      )}
      {dialog === "close" && (
        <WorkspaceDialog
          title="Save changes before closing?"
          onClose={() => setDialog(null)}
        >
          <p>{workspace.entries[closeId]?.name} has unsaved changes.</p>
          <div className="dialog-actions">
            <button
              className="vscode-btn-secondary"
              onClick={() => setDialog(null)}
            >
              Cancel
            </button>
            <button
              className="vscode-btn-secondary"
              disabled={workspace.entries[closeId]?.saving}
              onClick={() => {
                workspace.discard(closeId);
                setDialog(null);
              }}
            >
              Discard changes
            </button>
            <button
              className="vscode-btn"
              onClick={async () => {
                if (await workspace.saveAndClose(closeId)) {
                  setDialog(null);
                }
              }}
            >
              Save & close
            </button>
          </div>
        </WorkspaceDialog>
      )}
      {dialog === "reload" && entry && (
        <WorkspaceDialog
          title="Reload saved version?"
          onClose={() => setDialog(null)}
        >
          <p>
            Replace unsaved edits in {entry.name} with the latest saved
            workspace revision. Use Save a copy first to keep both versions.
          </p>
          <div className="dialog-actions">
            <button
              className="vscode-btn-secondary"
              onClick={() => setDialog(null)}
            >
              Cancel
            </button>
            <button
              className="vscode-btn"
              onClick={async () => {
                if (await workspace.reloadSaved(entry.id)) setDialog(null);
              }}
            >
              Reload saved
            </button>
          </div>
        </WorkspaceDialog>
      )}
      {dialog === "replace" && pendingReplace && (
        <WorkspaceDialog
          title="Replace across workspace?"
          onClose={() => setDialog(null)}
        >
          <p>
            Replace “{pendingReplace.query}” with “{pendingReplace.replacement}”
            in matching workspace files. Replacement text is literal. Each
            changed file can be undone separately.
          </p>
          <div className="dialog-actions">
            <button
              className="vscode-btn-secondary"
              onClick={() => setDialog(null)}
            >
              Cancel
            </button>
            <button
              className="vscode-btn"
              onClick={() => {
                try {
                  let count = 0;
                  for (const e of Object.values(workspace.entries)) {
                    const result = replaceInDocument(
                      e.document,
                      pendingReplace.query,
                      pendingReplace.replacement,
                      pendingReplace.options,
                    );
                    if (result.count) {
                      workspace.update(result.document, e.id);
                      count += result.count;
                    }
                  }
                  notify(
                    `Replaced ${count} occurrences. Save all to persist the changes.`,
                  );
                  setDialog(null);
                } catch (e) {
                  notify((e as Error).message, true);
                }
              }}
            >
              Replace all
            </button>
          </div>
        </WorkspaceDialog>
      )}
    </div>
  );
};
export default VSCodeOfficeStudio;
