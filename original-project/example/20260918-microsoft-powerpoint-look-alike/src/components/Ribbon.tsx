import { useState, type ComponentType, type ReactNode } from "react";
import {
  Accessibility,
  AlignCenter,
  AlignLeft,
  AlignRight,
  AudioLines,
  Baseline,
  Bold,
  Camera,
  ChartColumn,
  Check,
  ChevronRight,
  Circle,
  Clock,
  FileQuestion,
  GitFork,
  Highlighter,
  Image,
  Italic,
  Languages,
  Link2,
  MessageSquare,
  Minus,
  MonitorPlay,
  PenLine,
  Play,
  Plus,
  Presentation,
  Ruler,
  Save,
  Share2,
  Sigma,
  Sparkles,
  Square,
  SquarePlus,
  StickyNote,
  Table2,
  TextCursorInput,
  Timer,
  Trash2,
  Underline,
  Undo2,
  Redo2,
  Video,
  Copy,
  Eraser,
  PanelRight,
  LayoutGrid,
  PaintBucket,
  Frame,
  Zap,
} from "lucide-react";
import type { RibbonTab, SlideObject, ThemePreset, ViewMode } from "../types";

const FONT_FAMILIES = [
  "Segoe UI",
  "Calibri",
  "Arial",
  "Georgia",
  "Times New Roman",
  "Verdana",
  "Trebuchet MS",
  "Courier New",
];

const FONT_SIZES = [10, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 72, 88];

const TRANSITIONS = ["None", "Fade", "Push", "Wipe", "Split", "Morph", "Cut"];

type IconType = ComponentType<{ size?: number | string; strokeWidth?: number | string }>;

interface CommandPanelProps {
  active: RibbonTab;
  selected: SlideObject | null;
  theme: ThemePreset;
  themes: ThemePreset[];
  transition: string;
  viewMode: ViewMode;
  showNotes: boolean;
  inspectorOpen: boolean;
  autoSave: boolean;
  canUndo: boolean;
  canRedo: boolean;
  canDeleteSlide: boolean;
  onFormat: (patch: Partial<SlideObject>) => void;
  onNewSlide: () => void;
  onDuplicateSlide: () => void;
  onDeleteSlide: () => void;
  onInsert: (kind: "text" | "rect" | "ellipse" | "line" | "chart") => void;
  onTheme: (theme: ThemePreset) => void;
  onDesigner: () => void;
  onTransition: (name: string) => void;
  onPresent: (fromBeginning: boolean) => void;
  onViewMode: (mode: ViewMode) => void;
  onToggleNotes: () => void;
  onToggleInspector: () => void;
  onToggleAutoSave: () => void;
  onUndo: () => void;
  onRedo: () => void;
  onSave: () => void;
  onShare: () => void;
  onToast: (msg: string) => void;
}

function Accordion({
  title,
  children,
  defaultOpen = false,
}: {
  title: string;
  children: ReactNode;
  defaultOpen?: boolean;
}) {
  const [open, setOpen] = useState(defaultOpen);
  return (
    <section className={`side-accordion${open ? " side-accordion-open" : ""}`}>
      <button
        type="button"
        className="side-accordion-trigger"
        onClick={() => setOpen((value) => !value)}
        aria-expanded={open}
      >
        <ChevronRight size={14} />
        <span>{title}</span>
      </button>
      {open && <div className="side-accordion-content">{children}</div>}
    </section>
  );
}

function Command({
  icon: Icon,
  label,
  onClick,
  active,
  disabled,
  wide,
  title,
}: {
  icon?: IconType;
  label: string;
  onClick?: () => void;
  active?: boolean;
  disabled?: boolean;
  wide?: boolean;
  title?: string;
}) {
  return (
    <button
      type="button"
      className={`side-command${wide ? " side-command-wide" : ""}${active ? " side-command-active" : ""}`}
      onClick={onClick}
      disabled={disabled}
      aria-label={title ?? label}
      title={title ?? label}
      aria-pressed={active}
    >
      {Icon && <Icon size={14} />}
      <span>{label}</span>
    </button>
  );
}

function Toggle({
  icon: Icon,
  label,
  active,
  onClick,
}: {
  icon: IconType;
  label: string;
  active?: boolean;
  onClick: () => void;
}) {
  return (
    <button
      type="button"
      className={`side-icon-toggle${active ? " side-icon-toggle-active" : ""}`}
      onClick={onClick}
      aria-label={label}
      title={label}
      aria-pressed={active}
    >
      <Icon size={14} />
    </button>
  );
}

export function CommandPanel(props: CommandPanelProps) {
  const {
    active,
    selected,
    theme,
    themes,
    transition,
    viewMode,
    showNotes,
    inspectorOpen,
    autoSave,
    canUndo,
    canRedo,
    canDeleteSlide,
    onFormat,
    onNewSlide,
    onDuplicateSlide,
    onDeleteSlide,
    onInsert,
    onTheme,
    onDesigner,
    onTransition,
    onPresent,
    onViewMode,
    onToggleNotes,
    onToggleInspector,
    onToggleAutoSave,
    onUndo,
    onRedo,
    onSave,
    onShare,
    onToast,
  } = props;

  const soon = (name: string) => () => onToast(`${name} is not available in this build`);
  const sel = selected;
  const bold = (sel?.fontWeight ?? 400) >= 600;
  const italic = sel?.fontStyle === "italic";
  const underline = sel?.textDecoration === "underline";
  const fontSize = sel?.fontSize ?? 18;
  const sizeOptions = FONT_SIZES.includes(fontSize) ? FONT_SIZES : [...FONT_SIZES, fontSize].sort((a, b) => a - b);

  const renderPanel = () => {
    switch (active) {
      case "Home":
        return (
          <>
            <Accordion title="Document" defaultOpen>
              <div className="side-command-grid">
                <Command icon={Save} label="Save" onClick={onSave} />
                <Command icon={Share2} label="Share" onClick={onShare} />
                <Command icon={Undo2} label="Undo" onClick={onUndo} disabled={!canUndo} />
                <Command icon={Redo2} label="Redo" onClick={onRedo} disabled={!canRedo} />
              </div>
              <label className="side-field side-switch-row">
                <span>AutoSave</span>
                <button
                  type="button"
                  role="switch"
                  aria-checked={autoSave}
                  aria-label="AutoSave"
                  title="AutoSave"
                  className={`switch switch-dark${autoSave ? " switch-on" : ""}`}
                  onClick={onToggleAutoSave}
                >
                  <span className="switch-knob" />
                </button>
                <span className="side-field-meta">{autoSave ? "On" : "Off"}</span>
              </label>
            </Accordion>
            <Accordion title="Slides">
              <div className="side-command-grid">
                <Command icon={Plus} label="New slide" onClick={onNewSlide} />
                <Command icon={Copy} label="Duplicate" onClick={onDuplicateSlide} />
                <Command
                  icon={Trash2}
                  label="Delete slide"
                  wide
                  onClick={onDeleteSlide}
                  disabled={!canDeleteSlide}
                />
              </div>
            </Accordion>
            <Accordion title="Text">
              <label className="side-field">
                <span>Font family</span>
                <select
                  className="side-select"
                  value={sel?.fontFamily ?? "Segoe UI"}
                  onChange={(e) => onFormat({ fontFamily: e.currentTarget.value })}
                >
                  {FONT_FAMILIES.map((f) => (
                    <option key={f} value={f}>
                      {f}
                    </option>
                  ))}
                </select>
              </label>
              <label className="side-field">
                <span>Font size</span>
                <select
                  className="side-select"
                  value={fontSize}
                  onChange={(e) => onFormat({ fontSize: Number(e.currentTarget.value) })}
                >
                  {sizeOptions.map((s) => (
                    <option key={s} value={s}>
                      {s}
                    </option>
                  ))}
                </select>
              </label>
              <div className="side-field">
                <span>Style</span>
                <div className="side-toggle-row">
                  <Toggle
                    icon={Bold}
                    label="Bold"
                    active={bold}
                    onClick={() => onFormat({ fontWeight: bold ? 400 : 700 })}
                  />
                  <Toggle
                    icon={Italic}
                    label="Italic"
                    active={italic}
                    onClick={() => onFormat({ fontStyle: italic ? "normal" : "italic" })}
                  />
                  <Toggle
                    icon={Underline}
                    label="Underline"
                    active={underline}
                    onClick={() => onFormat({ textDecoration: underline ? "none" : "underline" })}
                  />
                </div>
              </div>
              <label className="side-field side-color-control">
                <span>
                  <Baseline size={13} /> Text color
                </span>
                <input
                  type="color"
                  value={sel?.color && /^#[0-9a-fA-F]{6}$/.test(sel.color) ? sel.color : "#1d2b45"}
                  aria-label="Text color"
                  onChange={(e) => onFormat({ color: e.currentTarget.value })}
                />
              </label>
            </Accordion>
            <Accordion title="Paragraph">
              <div className="side-field">
                <span>Alignment</span>
                <div className="side-toggle-row">
                  <Toggle
                    icon={AlignLeft}
                    label="Align left"
                    active={(sel?.align ?? "left") === "left"}
                    onClick={() => onFormat({ align: "left" })}
                  />
                  <Toggle
                    icon={AlignCenter}
                    label="Center"
                    active={sel?.align === "center"}
                    onClick={() => onFormat({ align: "center" })}
                  />
                  <Toggle
                    icon={AlignRight}
                    label="Align right"
                    active={sel?.align === "right"}
                    onClick={() => onFormat({ align: "right" })}
                  />
                </div>
              </div>
            </Accordion>
            <Accordion title="Drawing">
              <div className="side-command-grid">
                <Command icon={Square} label="Rectangle" onClick={() => onInsert("rect")} />
                <Command icon={Circle} label="Ellipse" onClick={() => onInsert("ellipse")} />
                <Command icon={Minus} label="Line" onClick={() => onInsert("line")} />
                <Command icon={Sparkles} label="Designer" onClick={onDesigner} />
              </div>
            </Accordion>
          </>
        );
      case "Insert":
        return (
          <>
            <Accordion title="Add to slide" defaultOpen>
              <div className="side-command-grid">
                <Command icon={TextCursorInput} label="Text Box" onClick={() => onInsert("text")} />
                <Command icon={Square} label="Rectangle" onClick={() => onInsert("rect")} />
                <Command icon={Circle} label="Ellipse" onClick={() => onInsert("ellipse")} />
                <Command icon={Minus} label="Line" onClick={() => onInsert("line")} />
                <Command icon={ChartColumn} label="Chart" wide onClick={() => onInsert("chart")} />
              </div>
            </Accordion>
            <Accordion title="Structured content">
              <div className="side-command-grid">
                <Command icon={Table2} label="Table" onClick={soon("Table")} />
                <Command icon={GitFork} label="SmartArt" onClick={soon("SmartArt")} />
                <Command icon={Sigma} label="Equation" wide onClick={soon("Equation")} />
              </div>
            </Accordion>
            <Accordion title="Media & links">
              <div className="side-command-grid">
                <Command icon={Image} label="Pictures" onClick={soon("Pictures")} />
                <Command icon={Video} label="Video" onClick={soon("Video")} />
                <Command icon={AudioLines} label="Audio" onClick={soon("Audio")} />
                <Command icon={Link2} label="Link" onClick={soon("Link")} />
              </div>
            </Accordion>
          </>
        );
      case "Draw":
        return (
          <>
            <Accordion title="Drawing tools" defaultOpen>
              <div className="side-command-grid">
                <Command icon={PenLine} label="Pen" onClick={soon("Pen")} />
                <Command icon={Highlighter} label="Highlighter" onClick={soon("Highlighter")} />
                <Command icon={Eraser} label="Eraser" onClick={soon("Eraser")} />
                <Command icon={Ruler} label="Ruler" onClick={soon("Ruler")} />
              </div>
            </Accordion>
          </>
        );
      case "Design":
        return (
          <>
            <Accordion title="Themes" defaultOpen>
              <div className="side-theme-list">
                {themes.map((t) => (
                  <button
                    key={t.id}
                    type="button"
                    className={`side-theme-card${t.id === theme.id ? " side-theme-card-active" : ""}`}
                    onClick={() => onTheme(t)}
                    aria-label={`Theme ${t.name}`}
                    title={t.name}
                    aria-pressed={t.id === theme.id}
                  >
                    <span className="side-theme-canvas" style={{ background: t.canvas }}>
                      <span className="theme-swatch-bar" style={{ background: t.accent }} />
                      <span className="theme-swatch-ink" style={{ background: t.ink }} />
                    </span>
                    <span className="side-theme-info">
                      <span className="side-theme-name">{t.name}</span>
                      <span className="side-theme-accent" style={{ background: t.accent }} />
                    </span>
                    {t.id === theme.id && <Check size={14} className="side-theme-check" />}
                  </button>
                ))}
              </div>
            </Accordion>
            <Accordion title="Design tools">
              <div className="side-command-grid">
                <Command icon={Sparkles} label="Designer" onClick={onDesigner} />
                <Command icon={Frame} label="Slide Size" onClick={soon("Slide size")} />
                <Command icon={PaintBucket} label="Format Background" wide onClick={soon("Format background")} />
              </div>
            </Accordion>
          </>
        );
      case "Transitions":
        return (
          <>
            <Accordion title="Transition" defaultOpen>
              <div className="side-transition-list">
                {TRANSITIONS.map((t) => (
                  <button
                    key={t}
                    type="button"
                    className={`side-transition-row${transition === t ? " side-transition-row-active" : ""}`}
                    onClick={() => onTransition(t)}
                    aria-pressed={transition === t}
                    title={t}
                  >
                    <span className={`transition-preview transition-preview-${t.toLowerCase()}`} />
                    <span className="side-transition-name">{t}</span>
                    {transition === t && <Check size={14} />}
                  </button>
                ))}
              </div>
            </Accordion>
            <Accordion title="Timing">
              <div className="side-command-grid">
                <Command icon={AudioLines} label="Sound" onClick={soon("Transition sound")} />
                <Command icon={Clock} label="Advance Slide" onClick={soon("Advance slide")} />
              </div>
            </Accordion>
          </>
        );
      case "Animations":
        return (
          <>
            <Accordion title="Entrance effects" defaultOpen>
              <div className="side-command-grid">
                <Command icon={Play} label="Appear" onClick={soon("Appear")} />
                <Command icon={Sparkles} label="Fade" onClick={soon("Fade animation")} />
                <Command icon={Zap} label="Fly In" onClick={soon("Fly in")} />
                <Command icon={SquarePlus} label="Zoom" onClick={soon("Zoom animation")} />
              </div>
            </Accordion>
            <Accordion title="Animation controls">
              <div className="side-command-grid">
                <Command icon={PanelRight} label="Animation Pane" onClick={soon("Animation pane")} />
                <Command icon={Zap} label="Trigger" onClick={soon("Trigger")} />
                <Command icon={Timer} label="Timing" wide onClick={soon("Animation timing")} />
              </div>
            </Accordion>
          </>
        );
      case "Slide Show":
        return (
          <>
            <Accordion title="Start presentation" defaultOpen>
              <div className="side-command-grid">
                <Command icon={Play} label="From Beginning" wide onClick={() => onPresent(true)} />
                <Command icon={MonitorPlay} label="From Current Slide" wide onClick={() => onPresent(false)} />
              </div>
            </Accordion>
            <Accordion title="Presenter tools">
              <div className="side-command-grid">
                <Command icon={Presentation} label="Presenter View" onClick={soon("Presenter view")} />
                <Command icon={Timer} label="Rehearse Timings" onClick={soon("Rehearse timings")} />
              </div>
            </Accordion>
          </>
        );
      case "Record":
        return (
          <>
            <Accordion title="Recording" defaultOpen>
              <div className="side-command-grid">
                <Command icon={Video} label="Record Slide Show" wide onClick={soon("Record slide show")} />
                <Command icon={Camera} label="Cameo" onClick={soon("Cameo")} />
                <Command icon={MonitorPlay} label="Screen Recording" onClick={soon("Screen recording")} />
              </div>
            </Accordion>
            <Accordion title="Export">
              <div className="side-command-grid">
                <Command icon={Video} label="Export to Video" wide onClick={soon("Export to video")} />
              </div>
            </Accordion>
          </>
        );
      case "Review":
        return (
          <>
            <Accordion title="Review" defaultOpen>
              <div className="side-command-grid">
                <Command
                  icon={Accessibility}
                  label="Accessibility checker"
                  wide
                  onClick={soon("Accessibility checker")}
                />
                <Command icon={MessageSquare} label="Comments" onClick={() => onToast("No comments yet")} />
                <Command icon={GitFork} label="Compare" onClick={soon("Compare")} />
                <Command icon={Languages} label="Translate" wide onClick={soon("Translate")} />
              </div>
            </Accordion>
          </>
        );
      case "View":
        return (
          <>
            <Accordion title="Presentation views" defaultOpen>
              <div className="side-command-grid">
                <Command
                  icon={PanelRight}
                  label="Normal"
                  active={viewMode === "normal"}
                  onClick={() => onViewMode("normal")}
                />
                <Command
                  icon={LayoutGrid}
                  label="Slide Sorter"
                  active={viewMode === "sorter"}
                  onClick={() => onViewMode("sorter")}
                />
              </div>
            </Accordion>
            <Accordion title="Panels">
              <div className="side-command-grid">
                <Command icon={StickyNote} label="Notes" active={showNotes} onClick={onToggleNotes} />
                <Command
                  icon={PanelRight}
                  label="Inspector"
                  active={inspectorOpen}
                  onClick={onToggleInspector}
                />
              </div>
            </Accordion>
            <Accordion title="Zoom">
              <p className="side-info-text">Zoom controls live in the status bar.</p>
            </Accordion>
          </>
        );
      case "Help":
      default:
        return (
          <>
            <Accordion title="Help" defaultOpen>
              <div className="side-command-grid">
                <Command icon={FileQuestion} label="Help" onClick={soon("Help")} />
                <Command icon={MessageSquare} label="Search Help" onClick={soon("Search help")} />
                <Command icon={MessageSquare} label="Feedback" wide onClick={soon("Feedback")} />
              </div>
            </Accordion>
          </>
        );
    }
  };

  return <div className="command-panel">{renderPanel()}</div>;
}
