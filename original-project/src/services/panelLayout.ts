export type PanelSizeKey = "primaryWidth" | "secondaryWidth" | "bottomHeight";
export interface PanelLayoutPreferences {
  primaryWidth: number;
  secondaryWidth: number;
  bottomHeight: number;
  primaryOpen: boolean;
  secondaryOpen: boolean;
  bottomOpen: boolean;
}
export const PANEL_LAYOUT_KEY = "office-studio.panels.v1";
export const DEFAULT_PANEL_LAYOUT: PanelLayoutPreferences = {
  primaryWidth: 260,
  secondaryWidth: 285,
  bottomHeight: 230,
  primaryOpen: true,
  secondaryOpen: true,
  bottomOpen: false,
};
export const clamp = (n: number, min: number, max: number) =>
  Math.min(max, Math.max(min, n));
export function parsePanelLayout(value: unknown): PanelLayoutPreferences {
  const defaults = { ...DEFAULT_PANEL_LAYOUT };
  if (!value || typeof value !== "object" || Array.isArray(value))
    return defaults;
  const candidate = value as Record<string, unknown>;
  for (const key of [
    "primaryWidth",
    "secondaryWidth",
    "bottomHeight",
  ] as const) {
    if (typeof candidate[key] === "number" && Number.isFinite(candidate[key]))
      defaults[key] = clamp(
        candidate[key],
        key === "bottomHeight" ? 80 : 100,
        key === "bottomHeight" ? 1200 : 640,
      );
  }
  for (const key of ["primaryOpen", "secondaryOpen", "bottomOpen"] as const)
    if (typeof candidate[key] === "boolean") defaults[key] = candidate[key];
  return defaults;
}

/** Fit visible docks while preserving a usable editor; saved preferred sizes remain intact. */
export function fitPanelLayout(
  p: PanelLayoutPreferences,
  width: number,
  height: number,
) {
  const compact = width < 900;
  const showPrimary = p.primaryOpen;
  const showSecondary = p.secondaryOpen && (!compact || !showPrimary);
  const primaryMin = Math.floor(Math.min(180, Math.max(100, width - 96)));
  const secondaryMin = Math.floor(Math.min(220, Math.max(100, width - 96)));
  const overlayMax = Math.floor(
    Math.max(primaryMin, Math.min(640, width - 72)),
  );
  const dockBudget = Math.floor(
    Math.max(
      400,
      width - 48 - 320 - (Number(showPrimary) + Number(showSecondary)) * 6,
    ),
  );
  let primaryWidth = clamp(
    Math.round(p.primaryWidth),
    primaryMin,
    compact ? overlayMax : 640,
  );
  let secondaryWidth = clamp(
    Math.round(p.secondaryWidth),
    secondaryMin,
    compact ? overlayMax : 640,
  );
  if (
    !compact &&
    showPrimary &&
    showSecondary &&
    primaryWidth + secondaryWidth > dockBudget
  ) {
    const extra = Math.max(
      1,
      primaryWidth - primaryMin + secondaryWidth - secondaryMin,
    );
    const scale = Math.max(0, (dockBudget - primaryMin - secondaryMin) / extra);
    primaryWidth = Math.round(primaryMin + (primaryWidth - primaryMin) * scale);
    secondaryWidth = dockBudget - primaryWidth;
  }
  const primaryMax = compact
    ? overlayMax
    : showPrimary
      ? Math.min(640, dockBudget - (showSecondary ? secondaryWidth : 0))
      : 640;
  const secondaryMax = compact
    ? overlayMax
    : showSecondary
      ? Math.min(640, dockBudget - (showPrimary ? primaryWidth : 0))
      : 640;
  primaryWidth = clamp(primaryWidth, primaryMin, primaryMax);
  secondaryWidth = clamp(secondaryWidth, secondaryMin, secondaryMax);
  const bottomMin = Math.floor(Math.min(140, Math.max(80, height * 0.35))),
    bottomMax = Math.floor(Math.max(bottomMin, height - 260));
  return {
    compact,
    showPrimary,
    showSecondary,
    primaryWidth: Math.round(primaryWidth),
    secondaryWidth: Math.round(secondaryWidth),
    bottomHeight: Math.round(clamp(p.bottomHeight, bottomMin, bottomMax)),
    bounds: {
      primaryWidth: {
        min: Math.round(primaryMin),
        max: Math.floor(primaryMax),
      },
      secondaryWidth: {
        min: Math.round(secondaryMin),
        max: Math.floor(secondaryMax),
      },
      bottomHeight: { min: Math.round(bottomMin), max: Math.floor(bottomMax) },
    },
  };
}
