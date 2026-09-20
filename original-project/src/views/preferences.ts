import type { DocumentType } from "../types/vscode";

export type ViewArea = "left" | "right";
export type CustomViewKind = "notes" | "context";
export interface CustomView {
  id: string;
  title: string;
  kind: CustomViewKind;
  area: ViewArea;
  notes: string;
}
export interface ViewDefinition {
  id: string;
  title: string;
  description: string;
}
export interface ViewPreferences {
  left: string[];
  right: string[];
  custom: CustomView[];
}
export type ViewProfiles = Record<DocumentType, ViewPreferences>;
export const VIEW_STORAGE_KEY = "office-studio.views.v1";
export const PROFILE_LABELS: Record<DocumentType, string> = {
  word: "Word",
  excel: "Excel",
  powerpoint: "PowerPoint",
  code: "Code & text",
};
const shared: ViewDefinition[] = [
  {
    id: "open-editors",
    title: "Open editors",
    description: "Switch between open documents.",
  },
  {
    id: "workspace-files",
    title: "Workspace files",
    description: "Open and import documents.",
  },
];
const leftViews: Record<DocumentType, ViewDefinition[]> = {
  word: [
    {
      id: "word-outline",
      title: "Document outline",
      description: "Navigate document headings.",
    },
    {
      id: "word-pages",
      title: "Pages & sections",
      description: "Browse document sections.",
    },
    {
      id: "word-comments",
      title: "Review comments",
      description: "Read and resolve comments.",
    },
    {
      id: "word-stats",
      title: "Document statistics",
      description: "Word count and reading statistics.",
    },
  ],
  excel: [
    {
      id: "excel-sheets",
      title: "Worksheets",
      description: "Select and add worksheets.",
    },
    {
      id: "excel-formulas",
      title: "Formulas & functions",
      description: "Inspect formulas and dependencies.",
    },
    {
      id: "excel-fields",
      title: "Data columns & fields",
      description: "Inspect worksheet columns.",
    },
    {
      id: "excel-filter-sort",
      title: "Sort rows",
      description: "Sort the active worksheet.",
    },
  ],
  powerpoint: [
    {
      id: "ppt-slides",
      title: "Slide navigator",
      description: "Select and organize slides.",
    },
    {
      id: "ppt-layers",
      title: "Objects & layers",
      description: "Select objects and control visibility.",
    },
    {
      id: "ppt-layouts",
      title: "Layouts & themes",
      description: "Change slide layouts.",
    },
    {
      id: "ppt-notes",
      title: "Speaker notes",
      description: "Edit the current slide's notes.",
    },
  ],
  code: [
    {
      id: "code-symbols",
      title: "Outline & symbols",
      description: "Browse functions and symbols.",
    },
  ],
};
const rightViews: ViewDefinition[] = [
  {
    id: "inspector",
    title: "Inspector",
    description: "Edit document and selection properties.",
  },
  {
    id: "copilot",
    title: "Context",
    description: "Explore local document text and checks.",
  },
  { id: "minimap", title: "Map", description: "See the document at a glance." },
  { id: "metadata", title: "Info", description: "Read file attributes." },
];
export function builtinViews(
  profile: DocumentType,
  area: ViewArea,
): ViewDefinition[] {
  return area === "left" ? [...shared, ...leftViews[profile]] : [...rightViews];
}
export function defaultViewPreferences(profile: DocumentType): ViewPreferences {
  return {
    left: builtinViews(profile, "left").map((v) => v.id),
    right: rightViews.map((v) => v.id),
    custom: [],
  };
}
export function availableViews(
  profile: DocumentType,
  preferences: ViewPreferences,
  area: ViewArea,
): ViewDefinition[] {
  return [
    ...builtinViews(profile, area),
    ...preferences.custom
      .filter((v) => v.area === area)
      .map((v) => ({
        id: v.id,
        title: v.title,
        description:
          v.kind === "notes"
            ? "Personal notes saved in this browser."
            : "Live context from the current document.",
      })),
  ];
}
export function normalizeViewPreferences(
  value: unknown,
  profile: DocumentType,
): ViewPreferences {
  const defaults = defaultViewPreferences(profile);
  if (!value || typeof value !== "object" || Array.isArray(value))
    return defaults;
  const input = value as Record<string, unknown>;
  const custom: CustomView[] = [];
  if (Array.isArray(input.custom))
    for (const entry of input.custom.slice(0, 24)) {
      if (
        !entry ||
        typeof entry !== "object" ||
        typeof entry.id !== "string" ||
        !/^custom:[a-zA-Z0-9-]{1,80}$/.test(entry.id) ||
        custom.some((v) => v.id === entry.id) ||
        typeof entry.title !== "string" ||
        !entry.title.trim() ||
        !["notes", "context"].includes(entry.kind) ||
        !["left", "right"].includes(entry.area)
      )
        continue;
      custom.push({
        id: entry.id,
        title: entry.title.trim().slice(0, 60),
        kind: entry.kind,
        area: entry.area,
        notes:
          typeof entry.notes === "string" ? entry.notes.slice(0, 100000) : "",
      });
    }
  const result = { ...defaults, custom };
  for (const area of ["left", "right"] as const) {
    const allowed = new Set(
      availableViews(profile, result, area).map((v) => v.id),
    );
    result[area] = Array.isArray(input[area])
      ? [
          ...new Set(
            (input[area] as unknown[]).filter(
              (id): id is string => typeof id === "string" && allowed.has(id),
            ),
          ),
        ]
      : defaults[area];
  }
  return result;
}
export function readViewProfiles(serialized: string | null): ViewProfiles {
  let input: Record<string, unknown> = {};
  try {
    const parsed = JSON.parse(serialized || "null");
    if (
      parsed?.version === 1 &&
      parsed.profiles &&
      typeof parsed.profiles === "object"
    )
      input = parsed.profiles;
  } catch {
    /* Recover safely from obsolete or invalid browser preferences. */
  }
  return Object.fromEntries(
    Object.keys(PROFILE_LABELS).map((profile) => [
      profile,
      normalizeViewPreferences(input[profile], profile as DocumentType),
    ]),
  ) as ViewProfiles;
}
export function setViewVisible(
  preferences: ViewPreferences,
  area: ViewArea,
  id: string,
  visible: boolean,
): ViewPreferences {
  return {
    ...preferences,
    [area]: visible
      ? [...new Set([...preferences[area], id])]
      : preferences[area].filter((v) => v !== id),
  };
}
export function moveView(
  preferences: ViewPreferences,
  area: ViewArea,
  id: string,
  direction: -1 | 1,
): ViewPreferences {
  const order = [...preferences[area]],
    from = order.indexOf(id),
    to = from + direction;
  if (from < 0 || to < 0 || to >= order.length) return preferences;
  [order[from], order[to]] = [order[to], order[from]];
  return { ...preferences, [area]: order };
}
