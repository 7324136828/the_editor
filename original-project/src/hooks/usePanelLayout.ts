import {
  useEffect,
  useRef,
  useState,
  type Dispatch,
  type SetStateAction,
} from "react";
import {
  PANEL_LAYOUT_KEY,
  DEFAULT_PANEL_LAYOUT,
  parsePanelLayout,
  fitPanelLayout,
  clamp,
  type PanelSizeKey,
} from "../services/panelLayout";
export function usePanelLayout() {
  const [preferences, setPreferences] = useState(() => {
    try {
      return parsePanelLayout(
        JSON.parse(localStorage.getItem(PANEL_LAYOUT_KEY) || "null"),
      );
    } catch {
      return { ...DEFAULT_PANEL_LAYOUT };
    }
  });
  const [viewport, setViewport] = useState({ width: 1280, height: 760 });
  const [storageError, setStorageError] = useState(false);
  const workspaceRef = useRef<HTMLDivElement>(null);
  useEffect(() => {
    const element = workspaceRef.current;
    if (!element) return;
    const measure = () => {
      setViewport({ width: element.clientWidth, height: element.clientHeight });
    };
    measure();
    const observer = new ResizeObserver(measure);
    observer.observe(element);
    return () => observer.disconnect();
  }, []);
  useEffect(() => {
    try {
      localStorage.setItem(PANEL_LAYOUT_KEY, JSON.stringify(preferences));
      setStorageError(false);
    } catch {
      setStorageError(true);
    }
  }, [preferences]);
  const fit = fitPanelLayout(preferences, viewport.width, viewport.height);
  const setPrimaryOpen: Dispatch<SetStateAction<boolean>> = (value) =>
    setPreferences((old) => {
      const current = fitPanelLayout(old, viewport.width, viewport.height);
      const open =
        typeof value === "function" ? value(current.showPrimary) : value;
      return {
        ...old,
        primaryOpen: open,
        ...(open && current.compact ? { secondaryOpen: false } : {}),
      };
    });
  const setSecondaryOpen: Dispatch<SetStateAction<boolean>> = (value) =>
    setPreferences((old) => {
      const current = fitPanelLayout(old, viewport.width, viewport.height);
      const open =
        typeof value === "function" ? value(current.showSecondary) : value;
      return {
        ...old,
        secondaryOpen: open,
        ...(open && current.compact ? { primaryOpen: false } : {}),
      };
    });
  const setBottomOpen: Dispatch<SetStateAction<boolean>> = (value) =>
    setPreferences((old) => ({
      ...old,
      bottomOpen: typeof value === "function" ? value(old.bottomOpen) : value,
    }));
  const resize = (key: PanelSizeKey, value: number) => {
    if (!Number.isFinite(value)) return;
    setPreferences((old) => {
      const current = fitPanelLayout(old, viewport.width, viewport.height);
      return {
        ...old,
        ...(key === "bottomHeight"
          ? {}
          : {
              ...(current.showPrimary
                ? { primaryWidth: current.primaryWidth }
                : {}),
              ...(current.showSecondary
                ? { secondaryWidth: current.secondaryWidth }
                : {}),
            }),
        [key]: clamp(value, current.bounds[key].min, current.bounds[key].max),
      };
    });
  };
  const resetSize = (key: PanelSizeKey) =>
    resize(key, DEFAULT_PANEL_LAYOUT[key]);
  const resetSizes = () =>
    setPreferences((old) => ({
      ...old,
      primaryWidth: DEFAULT_PANEL_LAYOUT.primaryWidth,
      secondaryWidth: DEFAULT_PANEL_LAYOUT.secondaryWidth,
      bottomHeight: DEFAULT_PANEL_LAYOUT.bottomHeight,
    }));
  return {
    ...fit,
    preferences,
    workspaceRef,
    storageError,
    resize,
    resetSize,
    resetSizes,
    setPrimaryOpen,
    setSecondaryOpen,
    setBottomOpen,
  };
}
