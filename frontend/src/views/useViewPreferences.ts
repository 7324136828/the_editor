import { useEffect, useState } from "react";
import type { DocumentType } from "../types/vscode";
import {
  availableViews,
  defaultViewPreferences,
  moveView,
  normalizeViewPreferences,
  readViewProfiles,
  setViewVisible,
  VIEW_STORAGE_KEY,
} from "./preferences";
import type {
  CustomView,
  ViewArea,
  ViewPreferences,
  ViewProfiles,
} from "./preferences";

function initialProfiles(): ViewProfiles {
  try {
    return readViewProfiles(localStorage.getItem(VIEW_STORAGE_KEY));
  } catch {
    return readViewProfiles(null);
  }
}
export function useViewPreferences(profile: DocumentType) {
  const [profiles, setProfiles] = useState(initialProfiles);
  const [storageError, setStorageError] = useState(false);
  useEffect(() => {
    try {
      localStorage.setItem(
        VIEW_STORAGE_KEY,
        JSON.stringify({ version: 1, profiles }),
      );
      setStorageError(false);
    } catch {
      setStorageError(true);
    }
  }, [profiles]);
  const preferences = profiles[profile];
  const update = (change: (current: ViewPreferences) => ViewPreferences) => {
    setProfiles((current) => ({
      ...current,
      [profile]: normalizeViewPreferences(change(current[profile]), profile),
    }));
  };
  return {
    profile,
    preferences,
    storageError,
    availableViews: (area: ViewArea) =>
      availableViews(profile, preferences, area),
    setVisible: (area: ViewArea, id: string, visible: boolean) =>
      update((current) => setViewVisible(current, area, id, visible)),
    move: (area: ViewArea, id: string, direction: -1 | 1) =>
      update((current) => moveView(current, area, id, direction)),
    addView: (
      input: Pick<CustomView, "title" | "kind" | "area">,
    ): string | null => {
      if (!input.title.trim() || preferences.custom.length >= 24) return null;
      const id = `custom:${globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(36).slice(2)}`}`;
      update((current) => ({
        ...setViewVisible(current, input.area, id, true),
        custom: [
          ...current.custom,
          { ...input, id, title: input.title.trim(), notes: "" },
        ],
      }));
      return id;
    },
    updateNotes: (id: string, notes: string) =>
      update((current) => ({
        ...current,
        custom: current.custom.map((view) =>
          view.id === id ? { ...view, notes } : view,
        ),
      })),
    removeCustomView: (id: string) =>
      update((current) => ({
        left: current.left.filter((view) => view !== id),
        right: current.right.filter((view) => view !== id),
        custom: current.custom.filter((view) => view.id !== id),
      })),
    resetViews: () =>
      update((current) => ({
        ...defaultViewPreferences(profile),
        custom: current.custom,
      })),
  };
}
export type ViewPreferencesController = ReturnType<typeof useViewPreferences>;
