import {
  Accessibility,
  FileQuestion,
  House,
  LayoutGrid,
  Palette,
  PanelLeft,
  PenLine,
  Presentation,
  Replace,
  Settings,
  Sparkles,
  SquarePlus,
  Video,
} from "lucide-react";
import type { WorkspaceActivity } from "../types";

interface ActivityBarProps {
  active: WorkspaceActivity;
  onChange: (activity: WorkspaceActivity) => void;
  onToast: (message: string) => void;
}

const ACTIVITIES: { id: WorkspaceActivity; label: string; icon: typeof House }[] = [
  { id: "Slides", label: "Slides", icon: LayoutGrid },
  { id: "Home", label: "Home", icon: House },
  { id: "Insert", label: "Insert", icon: SquarePlus },
  { id: "Draw", label: "Draw", icon: PenLine },
  { id: "Design", label: "Design", icon: Palette },
  { id: "Transitions", label: "Transitions", icon: Replace },
  { id: "Animations", label: "Animations", icon: Sparkles },
  { id: "Slide Show", label: "Slide Show", icon: Presentation },
  { id: "Record", label: "Record", icon: Video },
  { id: "Review", label: "Review", icon: Accessibility },
  { id: "View", label: "View", icon: PanelLeft },
  { id: "Help", label: "Help", icon: FileQuestion },
];

export function ActivityBar({ active, onChange, onToast }: ActivityBarProps) {
  return (
    <nav className="activity-bar" aria-label="Workspace activities">
      <button
        type="button"
        className="activity-tile"
        aria-label="PowerPoint Studio — Slides"
        title="PowerPoint Studio"
        onClick={() => onChange("Slides")}
      >
        <span className="activity-tile-mark" aria-hidden="true">
          P
        </span>
      </button>
      {ACTIVITIES.map(({ id, label, icon: Icon }) => (
        <button
          key={id}
          type="button"
          className={`activity-btn${active === id ? " activity-btn-active" : ""}`}
          aria-label={label}
          title={label}
          aria-pressed={active === id}
          onClick={() => onChange(id)}
        >
          <Icon size={20} strokeWidth={1.6} />
        </button>
      ))}
      <div className="activity-spacer" />
      <button
        type="button"
        className="activity-btn"
        aria-label="Account: Zach N."
        title="Zach N."
        aria-pressed={false}
        onClick={() => onToast("Signed in as Zach N.")}
      >
        <span className="activity-avatar-mark">ZN</span>
      </button>
      <button
        type="button"
        className="activity-btn"
        aria-label="Settings"
        title="Settings"
        aria-pressed={false}
        onClick={() => onToast("Settings are not available in this build")}
      >
        <Settings size={19} strokeWidth={1.6} />
      </button>
    </nav>
  );
}
