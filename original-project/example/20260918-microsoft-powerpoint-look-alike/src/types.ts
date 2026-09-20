export type ObjectKind = "text" | "shape" | "line" | "chart";
export type ShapeKind = "rect" | "ellipse" | "pill" | "ring";

export interface SlideObject {
  id: string;
  kind: ObjectKind;
  shape?: ShapeKind;
  x: number;
  y: number;
  width: number;
  height: number;
  text?: string;
  fill?: string;
  stroke?: string;
  strokeWidth?: number;
  opacity?: number;
  rotation?: number;
  radius?: number;
  color?: string;
  fontFamily?: string;
  fontSize?: number;
  fontWeight?: number;
  fontStyle?: "normal" | "italic";
  textDecoration?: "none" | "underline";
  align?: "left" | "center" | "right";
  lineHeight?: number;
  letterSpacing?: number;
  zIndex?: number;
  chartData?: number[];
}

export interface Slide {
  id: string;
  title: string;
  background: string;
  objects: SlideObject[];
  notes: string;
  transition: string;
}

export interface ThemePreset {
  id: string;
  name: string;
  accent: string;
  accentSoft: string;
  ink: string;
  canvas: string;
}

export type RibbonTab =
  | "Home"
  | "Insert"
  | "Draw"
  | "Design"
  | "Transitions"
  | "Animations"
  | "Slide Show"
  | "Record"
  | "Review"
  | "View"
  | "Help";

export type ViewMode = "normal" | "sorter";

export type WorkspaceActivity = "Slides" | RibbonTab;
