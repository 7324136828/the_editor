import { useCallback, useEffect, useRef, useState } from "react";
import type { RibbonTab, Slide, SlideObject, ThemePreset, ViewMode, WorkspaceActivity } from "./types";
import { initialSlides, themePresets } from "./data/slides";
import { ActivityBar } from "./components/TitleBar";
import { CommandPanel } from "./components/Ribbon";
import { SlidePane } from "./components/SlidePane";
import { SlideRenderer } from "./components/SlideRenderer";
import { Inspector, type DesignIdea } from "./components/Inspector";
import { StatusBar } from "./components/StatusBar";
import { PresentationView } from "./components/PresentationView";

const FILE_NAME = "FY26 Product Strategy";
const DARK_BACKGROUNDS = new Set(["#101a2d", "#1b2946"]);

let idCounter = 0;
const nextId = (prefix: string) => `${prefix}-${Date.now().toString(36)}-${idCounter++}`;

function isDarkBackground(background: string) {
  if (DARK_BACKGROUNDS.has(background)) return true;
  const m = /^#([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(background);
  if (!m) return false;
  const lum = (parseInt(m[1], 16) * 299 + parseInt(m[2], 16) * 587 + parseInt(m[3], 16) * 114) / 1000;
  return lum < 110;
}

function makeBlankSlide(index: number): Slide {
  return {
    id: nextId("slide"),
    title: `Slide ${index}`,
    background: "var(--canvas)",
    transition: "None",
    notes: "",
    objects: [
      {
        id: nextId("obj"),
        kind: "text",
        x: 70,
        y: 56,
        width: 620,
        height: 56,
        text: "Click to add title",
        color: "var(--ink)",
        fontSize: 36,
        fontWeight: 700,
        letterSpacing: -0.3,
        zIndex: 1,
      },
    ],
  };
}

export default function App() {
  const [slides, setSlides] = useState<Slide[]>(initialSlides);
  const [activeSlideId, setActiveSlideId] = useState(initialSlides[0].id);
  const [selectedObjectId, setSelectedObjectId] = useState<string | null>(null);
  const [activeActivity, setActiveActivity] = useState<WorkspaceActivity>("Slides");
  const [theme, setTheme] = useState<ThemePreset>(themePresets[0]);
  const [viewMode, setViewMode] = useState<ViewMode>("normal");
  const [zoom, setZoom] = useState(100);
  const [showNotes, setShowNotes] = useState(false);
  const [inspectorOpen, setInspectorOpen] = useState(false);
  const [presentStart, setPresentStart] = useState<number | null>(null);
  const [editingText, setEditingText] = useState(false);
  const [autoSave, setAutoSave] = useState(true);
  const [toast, setToast] = useState<string | null>(null);
  const [past, setPast] = useState<Slide[][]>([]);
  const [future, setFuture] = useState<Slide[][]>([]);
  const [pasteboardSize, setPasteboardSize] = useState({ w: 900, h: 600 });

  const toastTimer = useRef<number | null>(null);
  const historyBaseRef = useRef<Slide[] | null>(null);
  const pasteboardRef = useRef<HTMLDivElement>(null);

  const activeSlide = slides.find((s) => s.id === activeSlideId) ?? slides[0];
  const activeIndex = slides.indexOf(activeSlide);
  const selectedObject = activeSlide.objects.find((o) => o.id === selectedObjectId) ?? null;

  const showToast = useCallback((msg: string) => {
    setToast(msg);
    if (toastTimer.current) window.clearTimeout(toastTimer.current);
    toastTimer.current = window.setTimeout(() => setToast(null), 2400);
  }, []);

  const applySlides = useCallback(
    (next: Slide[]) => {
      setPast((p) => [...p.slice(-49), slides]);
      setFuture([]);
      setSlides(next);
    },
    [slides]
  );

  const patchSlide = useCallback(
    (id: string, patch: Partial<Slide>) => {
      applySlides(slides.map((s) => (s.id === id ? { ...s, ...patch } : s)));
    },
    [slides, applySlides]
  );

  const updateObject = useCallback(
    (id: string, patch: Partial<SlideObject>) => {
      if (historyBaseRef.current === null) historyBaseRef.current = slides;
      setSlides((cur) =>
        cur.map((s) =>
          s.id === activeSlide.id
            ? { ...s, objects: s.objects.map((o) => (o.id === id ? { ...o, ...patch } : o)) }
            : s
        )
      );
    },
    [slides, activeSlide.id]
  );

  const commitObject = useCallback(
    (id: string, patch: Partial<SlideObject>) => {
      const base = historyBaseRef.current ?? slides;
      historyBaseRef.current = null;
      setPast((p) => [...p.slice(-49), base]);
      setFuture([]);
      setSlides((cur) =>
        cur.map((s) =>
          s.id === activeSlide.id
            ? { ...s, objects: s.objects.map((o) => (o.id === id ? { ...o, ...patch } : o)) }
            : s
        )
      );
    },
    [slides, activeSlide.id]
  );

  const formatObject = useCallback(
    (patch: Partial<SlideObject>) => {
      if (!selectedObjectId) {
        showToast("Select an object on the slide first");
        return;
      }
      commitObject(selectedObjectId, patch);
    },
    [selectedObjectId, commitObject, showToast]
  );

  const undo = useCallback(() => {
    setPast((p) => {
      if (p.length === 0) return p;
      const prev = p[p.length - 1];
      setFuture((f) => [...f, slides]);
      setSlides(prev);
      return p.slice(0, -1);
    });
  }, [slides]);

  const redo = useCallback(() => {
    setFuture((f) => {
      if (f.length === 0) return f;
      const next = f[f.length - 1];
      setPast((p) => [...p, slides]);
      setSlides(next);
      return f.slice(0, -1);
    });
  }, [slides]);

  const addSlide = useCallback(() => {
    const blank = makeBlankSlide(slides.length + 1);
    const next = [...slides];
    next.splice(activeIndex + 1, 0, blank);
    applySlides(next);
    setActiveSlideId(blank.id);
    setSelectedObjectId(null);
    setViewMode("normal");
  }, [slides, activeIndex, applySlides]);

  const duplicateSlide = useCallback(() => {
    const copy: Slide = {
      ...activeSlide,
      id: nextId("slide"),
      objects: activeSlide.objects.map((o) => ({ ...o, id: nextId("obj") })),
    };
    const next = [...slides];
    next.splice(activeIndex + 1, 0, copy);
    applySlides(next);
    setActiveSlideId(copy.id);
    setSelectedObjectId(null);
    showToast("Slide duplicated");
  }, [slides, activeSlide, activeIndex, applySlides, showToast]);

  const deleteSlide = useCallback(() => {
    if (slides.length <= 1) return;
    const next = slides.filter((s) => s.id !== activeSlide.id);
    applySlides(next);
    setActiveSlideId(next[Math.min(activeIndex, next.length - 1)].id);
    setSelectedObjectId(null);
    showToast("Slide deleted");
  }, [slides, activeSlide.id, activeIndex, applySlides, showToast]);

  const reorderSlides = useCallback(
    (sourceId: string, targetId: string) => {
      const from = slides.findIndex((s) => s.id === sourceId);
      const to = slides.findIndex((s) => s.id === targetId);
      if (from < 0 || to < 0 || from === to) return;
      const next = [...slides];
      const [moved] = next.splice(from, 1);
      next.splice(to, 0, moved);
      applySlides(next);
    },
    [slides, applySlides]
  );

  const deleteObject = useCallback(() => {
    if (!selectedObjectId) return;
    applySlides(
      slides.map((s) =>
        s.id === activeSlide.id ? { ...s, objects: s.objects.filter((o) => o.id !== selectedObjectId) } : s
      )
    );
    setSelectedObjectId(null);
    showToast("Object deleted");
  }, [slides, activeSlide.id, selectedObjectId, applySlides, showToast]);

  const insertObject = useCallback(
    (kind: "text" | "rect" | "ellipse" | "line" | "chart") => {
      const dark = isDarkBackground(activeSlide.background);
      const maxZ = Math.max(0, ...activeSlide.objects.map((o) => o.zIndex ?? 0));
      const base: SlideObject = {
        id: nextId("obj"),
        kind: "shape",
        x: 0,
        y: 0,
        width: 100,
        height: 100,
        zIndex: maxZ + 1,
      };
      if (kind === "text") {
        Object.assign(base, {
          kind: "text",
          x: 310,
          y: 244,
          width: 340,
          height: 52,
          text: "Click to add text",
          color: dark ? "#ffffff" : "var(--ink)",
          fontSize: 20,
        });
      } else if (kind === "rect") {
        Object.assign(base, { shape: "rect", x: 380, y: 200, width: 200, height: 140, fill: "var(--accent)", radius: 4 });
      } else if (kind === "ellipse") {
        Object.assign(base, { shape: "ellipse", x: 400, y: 190, width: 160, height: 160, fill: "var(--accent)" });
      } else if (kind === "line") {
        Object.assign(base, { kind: "line", x: 330, y: 268, width: 300, height: 0, stroke: dark ? "#ffffff" : "var(--ink)", strokeWidth: 2.5 });
      } else {
        Object.assign(base, {
          kind: "chart",
          x: 290,
          y: 180,
          width: 380,
          height: 190,
          fill: "var(--accent)",
          chartData: [30, 52, 44, 68, 60, 84],
        });
      }
      applySlides(
        slides.map((s) => (s.id === activeSlide.id ? { ...s, objects: [...s.objects, base] } : s))
      );
      setSelectedObjectId(base.id);
      setViewMode("normal");
      showToast(
        kind === "text"
          ? "Text box inserted"
          : kind === "chart"
            ? "Chart inserted"
            : `${kind === "rect" ? "Rectangle" : kind === "ellipse" ? "Ellipse" : "Line"} inserted`
      );
    },
    [slides, activeSlide, applySlides, showToast]
  );

  const arrange = useCallback(
    (dir: 1 | -1) => {
      if (!selectedObjectId) return;
      const sorted = [...activeSlide.objects].sort((a, b) => (a.zIndex ?? 0) - (b.zIndex ?? 0));
      const idx = sorted.findIndex((o) => o.id === selectedObjectId);
      const swapIdx = idx + dir;
      if (idx < 0 || swapIdx < 0 || swapIdx >= sorted.length) {
        showToast(dir === 1 ? "Already on top" : "Already at the back");
        return;
      }
      const a = sorted[idx];
      const b = sorted[swapIdx];
      const aZ = a.zIndex ?? 0;
      const bZ = b.zIndex ?? 0;
      const nextObjects = activeSlide.objects.map((o) => {
        if (o.id === a.id) return { ...o, zIndex: aZ === bZ ? bZ + dir : bZ };
        if (o.id === b.id) return { ...o, zIndex: aZ === bZ ? aZ : aZ };
        return o;
      });
      applySlides(slides.map((s) => (s.id === activeSlide.id ? { ...s, objects: nextObjects } : s)));
    },
    [slides, activeSlide, selectedObjectId, applySlides, showToast]
  );

  const applyIdea = useCallback(
    (idea: DesignIdea) => {
      const titleObj = activeSlide.objects
        .filter((o) => o.kind === "text")
        .sort((a, b) => (b.fontSize ?? 0) - (a.fontSize ?? 0))[0];
      applySlides(
        slides.map((s) =>
          s.id === activeSlide.id
            ? {
                ...s,
                background: idea.background,
                objects: s.objects.map((o) => (titleObj && o.id === titleObj.id ? { ...o, color: idea.ink } : o)),
              }
            : s
        )
      );
      showToast("Design idea applied");
    },
    [slides, activeSlide, applySlides, showToast]
  );

  const setTransition = useCallback(
    (name: string) => {
      patchSlide(activeSlide.id, { transition: name });
      showToast(name === "None" ? "Transition removed" : `Transition set to ${name}`);
    },
    [activeSlide.id, patchSlide, showToast]
  );

  const updateNotes = useCallback(
    (notes: string) => {
      setSlides((cur) => cur.map((s) => (s.id === activeSlide.id ? { ...s, notes } : s)));
    },
    [activeSlide.id]
  );

  useEffect(() => {
    const el = pasteboardRef.current;
    if (!el) return;
    const ro = new ResizeObserver(() => {
      setPasteboardSize({ w: el.clientWidth, h: el.clientHeight });
    });
    ro.observe(el);
    setPasteboardSize({ w: el.clientWidth, h: el.clientHeight });
    return () => ro.disconnect();
  }, [inspectorOpen, viewMode]);

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      const target = e.target as HTMLElement | null;
      const inField =
        !!target && (target.tagName === "INPUT" || target.tagName === "TEXTAREA" || target.tagName === "SELECT" || target.isContentEditable);
      if ((e.ctrlKey || e.metaKey) && !inField) {
        if (e.key === "z" || e.key === "Z") {
          e.preventDefault();
          if (e.shiftKey) redo();
          else undo();
          return;
        }
        if (e.key === "y" || e.key === "Y") {
          e.preventDefault();
          redo();
          return;
        }
        if (e.key === "s" || e.key === "S") {
          e.preventDefault();
          showToast("Saved to OneDrive");
          return;
        }
        if (e.key === "d" || e.key === "D") {
          e.preventDefault();
          duplicateSlide();
          return;
        }
      }
      if (editingText || inField || presentStart !== null) return;
      if (e.key === "Escape") {
        setSelectedObjectId(null);
        return;
      }
      if ((e.key === "Delete" || e.key === "Backspace") && selectedObjectId) {
        e.preventDefault();
        deleteObject();
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [editingText, presentStart, selectedObjectId, deleteObject, undo, redo, duplicateSlide, showToast]);

  useEffect(() => {
    if (selectedObjectId && !activeSlide.objects.some((o) => o.id === selectedObjectId)) {
      setSelectedObjectId(null);
    }
  }, [activeSlide, selectedObjectId]);

  const fitScale = Math.min((pasteboardSize.w - 96) / 960, (pasteboardSize.h - 96) / 540, 0.92);
  const editorScale = Math.max(0.2, fitScale * (zoom / 100));
  const sorterScale = 0.26;

  return (
    <div
      className="app-shell"
      style={
        {
          "--accent": theme.accent,
          "--accent-soft": theme.accentSoft,
          "--ink": theme.ink,
          "--canvas": theme.canvas,
        } as React.CSSProperties
      }
    >
      <div className="editor-body">
        <ActivityBar active={activeActivity} onChange={setActiveActivity} onToast={showToast} />
        <aside className="left-panel" aria-label={`${activeActivity} panel`}>
          <header className="left-panel-header">
            <span className="left-panel-kicker">{activeActivity.toUpperCase()}</span>
            <strong className="left-panel-title">{FILE_NAME}</strong>
            <span className="left-panel-meta">
              {activeActivity === "Slides" ? `${slides.length} slides` : "Presentation tools"}
            </span>
          </header>
          {activeActivity === "Slides" ? (
            <SlidePane
              slides={slides}
              activeSlideId={activeSlide.id}
              onSelect={(id) => {
                setActiveSlideId(id);
                setSelectedObjectId(null);
                setViewMode("normal");
              }}
              onAdd={addSlide}
              onDuplicate={duplicateSlide}
              onDelete={deleteSlide}
              onReorder={reorderSlides}
            />
          ) : (
            <CommandPanel
              active={activeActivity as RibbonTab}
              selected={selectedObject}
              theme={theme}
              themes={themePresets}
              transition={activeSlide.transition}
              viewMode={viewMode}
              showNotes={showNotes}
              inspectorOpen={inspectorOpen}
              autoSave={autoSave}
              canUndo={past.length > 0}
              canRedo={future.length > 0}
              canDeleteSlide={slides.length > 1}
              onFormat={formatObject}
              onNewSlide={addSlide}
              onDuplicateSlide={duplicateSlide}
              onDeleteSlide={deleteSlide}
              onInsert={insertObject}
              onTheme={(t) => {
                setTheme(t);
                showToast(`Theme: ${t.name}`);
              }}
              onDesigner={() => {
                setInspectorOpen(true);
                setSelectedObjectId(null);
              }}
              onTransition={setTransition}
              onPresent={(fromBeginning) => setPresentStart(fromBeginning ? 0 : activeIndex)}
              onViewMode={setViewMode}
              onToggleNotes={() => setShowNotes((v) => !v)}
              onToggleInspector={() => setInspectorOpen((v) => !v)}
              onToggleAutoSave={() => {
                setAutoSave((v) => !v);
                showToast(autoSave ? "AutoSave off" : "AutoSave on");
              }}
              onUndo={undo}
              onRedo={redo}
              onSave={() => showToast("Saved to OneDrive")}
              onShare={() => showToast("Share link copied")}
              onToast={showToast}
            />
          )}
        </aside>
        <div className="main-column">
          <div className="workspace">
            {viewMode === "normal" ? (
              <main ref={pasteboardRef} className="pasteboard" aria-label="Slide canvas">
                <div className="pasteboard-center">
                  <SlideRenderer
                    slide={activeSlide}
                    scale={editorScale}
                    interactive
                    selectedObjectId={selectedObjectId}
                    onSelectObject={setSelectedObjectId}
                    onUpdateObject={updateObject}
                    onCommitObject={commitObject}
                    onEditingChange={setEditingText}
                  />
                </div>
              </main>
            ) : (
              <main className="sorter" aria-label="Slide sorter">
                <div className="sorter-grid">
                  {slides.map((slide, i) => (
                    <button
                      key={slide.id}
                      type="button"
                      className={`sorter-card${slide.id === activeSlide.id ? " sorter-card-active" : ""}`}
                      onClick={() => {
                        setActiveSlideId(slide.id);
                        setSelectedObjectId(null);
                        setViewMode("normal");
                      }}
                      aria-label={`Slide ${i + 1}: ${slide.title}`}
                    >
                      <span className="sorter-number">{i + 1}</span>
                      <span className="sorter-frame">
                        <SlideRenderer slide={slide} scale={sorterScale} />
                      </span>
                      <span className="sorter-meta">
                        {slide.title} · {slide.transition}
                      </span>
                    </button>
                  ))}
                </div>
              </main>
            )}
            {inspectorOpen && (
              <Inspector
                object={selectedObject}
                accentHex={theme.accent}
                onUpdate={commitObject}
                onBringForward={() => arrange(1)}
                onSendBackward={() => arrange(-1)}
                onApplyIdea={applyIdea}
                onClose={() => setInspectorOpen(false)}
              />
            )}
          </div>
          {showNotes && (
            <div className="notes-strip">
              <label className="notes-label" htmlFor="notes-field">
                Notes
              </label>
              <textarea
                id="notes-field"
                value={activeSlide.notes}
                placeholder="Click to add notes"
                onChange={(e) => updateNotes(e.currentTarget.value)}
              />
            </div>
          )}
        </div>
      </div>
      <StatusBar
        slideIndex={activeIndex}
        slideCount={slides.length}
        showNotes={showNotes}
        onToggleNotes={() => setShowNotes((v) => !v)}
        onComments={() => showToast("No comments yet")}
        viewMode={viewMode}
        onViewMode={setViewMode}
        zoom={zoom}
        onZoom={setZoom}
        onPresent={() => setPresentStart(activeIndex)}
      />
      {presentStart !== null && (
        <PresentationView
          slides={slides}
          startIndex={presentStart}
          onExit={(last) => {
            setActiveSlideId(slides[last].id);
            setPresentStart(null);
          }}
        />
      )}
      {toast && (
        <div className="toast" role="status">
          {toast}
        </div>
      )}
    </div>
  );
}
