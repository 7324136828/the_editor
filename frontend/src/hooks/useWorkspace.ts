import { useEffect, useRef, useState } from "react";
import { OfficeDocument } from "../types/office";
import { listDocuments, saveDocument, ApiError } from "../services/api";
import {
  clone,
  fingerprint,
  normalizeDocument,
  sampleDocuments,
} from "../services/documents";
import {
  recoverWorkspace,
  isStoredDocument,
  mergeWorkspaceRecords,
} from "./workspaceState";
import type { WorkspaceEntry, WorkspaceState } from "./workspaceState";
export type { WorkspaceEntry } from "./workspaceState";
const CACHE_KEY = "office-studio.workspace.v2";
function initialState(defaultType: string): WorkspaceState {
  try {
    const cached = JSON.parse(localStorage.getItem(CACHE_KEY) || "null");
    const recovered = recoverWorkspace(cached);
    if (recovered) return recovered;
  } catch {
    /* Invalid recovery data must not prevent startup. */
  }
  const entries: Record<string, WorkspaceEntry> = {};
  for (const sample of sampleDocuments) {
    const document = normalizeDocument(clone(sample));
    entries[document.data.id] = {
      id: document.data.id,
      name: document.data.title,
      document,
      revision: null,
      saved: null,
      seeded: true,
    };
  }
  return {
    entries,
    openIds: Object.keys(entries),
    activeId:
      Object.values(entries).find((e) => e.document.type === defaultType)?.id ??
      Object.keys(entries)[0] ??
      "",
  };
}

export function useWorkspace(
  defaultType: string,
  notify: (message: string, error?: boolean) => void,
) {
  const [state, setState] = useState(() => initialState(defaultType));
  const stateRef = useRef(state);
  const [ready, setReady] = useState(false);
  const [connected, setConnected] = useState(false);
  const [recoveryError, setRecoveryError] = useState(false);
  const histories = useRef<
    Record<
      string,
      { past: OfficeDocument[]; future: OfficeDocument[]; time: number }
    >
  >({});
  const inFlight = useRef(new Map<string, Promise<boolean>>());
  const refreshGeneration = useRef(0);
  const mutate = (fn: (old: WorkspaceState) => WorkspaceState) => {
    const next = fn(stateRef.current);
    stateRef.current = next;
    setState(next);
  };
  const refresh = async () => {
    const generation = ++refreshGeneration.current;
    const atRequest = stateRef.current;
    try {
      const records = await listDocuments();
      if (generation !== refreshGeneration.current) return false;
      if (!Array.isArray(records) || !records.every(isStoredDocument))
        throw new Error(
          "The storage server returned an invalid workspace. Existing editor drafts were retained.",
        );
      mutate((old) =>
        mergeWorkspaceRecords(
          old,
          atRequest,
          records,
          new Set(inFlight.current.keys()),
        ),
      );
      setConnected(true);
      return true;
    } catch (error) {
      if (generation !== refreshGeneration.current) return false;
      setConnected(false);
      if (error instanceof Error && error.message.includes("invalid workspace"))
        notify(error.message, true);
      return false;
    } finally {
      if (generation === refreshGeneration.current) setReady(true);
    }
  };
  useEffect(() => {
    void refresh();
  }, []);
  useEffect(() => {
    try {
      localStorage.setItem(CACHE_KEY, JSON.stringify(state));
      setRecoveryError(false);
    } catch {
      setRecoveryError(true);
    }
  }, [state]);
  const dirty = (entry: WorkspaceEntry) =>
    entry.saved !== fingerprint(entry.document);
  useEffect(() => {
    const unload = (e: BeforeUnloadEvent) => {
      if (Object.values(stateRef.current.entries).some(dirty)) {
        e.preventDefault();
        e.returnValue = "";
      }
    };
    window.addEventListener("beforeunload", unload);
    return () => window.removeEventListener("beforeunload", unload);
  }, []);
  const open = (id: string) =>
    mutate((old) =>
      old.entries[id]
        ? {
            ...old,
            openIds: old.openIds.includes(id)
              ? old.openIds
              : [...old.openIds, id],
            activeId: id,
          }
        : old,
    );
  const close = (id: string) =>
    mutate((old) => {
      const openIds = old.openIds.filter((i) => i !== id);
      return {
        ...old,
        openIds,
        activeId:
          old.activeId === id
            ? (openIds[openIds.length - 1] ?? "")
            : old.activeId,
      };
    });
  const add = (document: OfficeDocument, name: string) => {
    const id = `doc-${crypto.randomUUID()}`;
    const model = normalizeDocument({
      ...document,
      data: { ...document.data, id, title: name },
    } as OfficeDocument);
    mutate((old) => ({
      entries: {
        ...old.entries,
        [id]: { id, name, document: model, revision: null, saved: null },
      },
      openIds: [...old.openIds, id],
      activeId: id,
    }));
    return id;
  };
  const update = (
    document: OfficeDocument,
    id = document.data.id,
    record = true,
  ) => {
    const existing = stateRef.current.entries[id];
    if (!existing) return;
    if (document.data.id !== id || document.type !== existing.document.type) {
      notify("Cannot apply changes to a different document.", true);
      return;
    }
    const normalized = normalizeDocument(document);
    if (fingerprint(existing.document) === fingerprint(normalized)) return;
    if (record) {
      const history = (histories.current[id] ??= {
        past: [],
        future: [],
        time: 0,
      });
      if (Date.now() - history.time > 600 || !history.past.length)
        history.past = [...history.past.slice(-99), clone(existing.document)];
      history.future = [];
      history.time = Date.now();
    }
    mutate((old) => ({
      ...old,
      entries: {
        ...old.entries,
        [id]: {
          ...old.entries[id],
          document: normalized,
          name: normalized.data.title,
          seeded: false,
          error: undefined,
        },
      },
    }));
  };
  const undo = (redo = false) => {
    const id = stateRef.current.activeId,
      entry = stateRef.current.entries[id],
      history = histories.current[id];
    if (!entry || !history) return;
    const source = redo ? history.future : history.past,
      dest = redo ? history.past : history.future;
    const document = source.pop();
    if (!document) return;
    dest.push(clone(entry.document));
    history.time = 0;
    update(document, id, false);
  };
  const save = (id = stateRef.current.activeId): Promise<boolean> => {
    if (!ready) {
      notify("Wait for workspace storage to finish loading before saving.");
      return Promise.resolve(false);
    }
    const pending = inFlight.current.get(id);
    if (pending) return pending;
    const entry = stateRef.current.entries[id];
    if (!entry) return Promise.resolve(false);
    if (!dirty(entry) && entry.revision !== null) return Promise.resolve(true);
    const captured = clone(entry.document);
    if (histories.current[id]) histories.current[id].time = 0;
    mutate((old) => ({
      ...old,
      entries: {
        ...old.entries,
        [id]: { ...old.entries[id], saving: true, error: undefined },
      },
    }));
    const task = (async () => {
      try {
        const record = await saveDocument(
          id,
          entry.name,
          captured,
          entry.revision,
        );
        mutate((old) =>
          !old.entries[id]
            ? old
            : {
                ...old,
                entries: {
                  ...old.entries,
                  [id]: {
                    ...old.entries[id],
                    saved: fingerprint(captured),
                    revision: record.revision,
                    updatedAt: record.updatedAt,
                    saving: false,
                    seeded: false,
                    error: undefined,
                  },
                },
              },
        );
        setConnected(true);
        notify(`Saved ${entry.name} · revision ${record.revision}`);
        return true;
      } catch (error) {
        const message = error instanceof Error ? error.message : "Save failed";
        if (
          error instanceof ApiError &&
          (error.status === 0 || error.status >= 500)
        )
          setConnected(false);
        mutate((old) =>
          !old.entries[id]
            ? old
            : {
                ...old,
                entries: {
                  ...old.entries,
                  [id]: { ...old.entries[id], saving: false, error: message },
                },
              },
        );
        notify(message, true);
        return false;
      } finally {
        inFlight.current.delete(id);
      }
    })();
    inFlight.current.set(id, task);
    return task;
  };
  const rename = (id: string, name: string) => {
    const entry = stateRef.current.entries[id];
    if (!entry) return;
    update(
      {
        ...entry.document,
        data: { ...entry.document.data, title: name },
      } as OfficeDocument,
      id,
    );
    mutate((old) => ({
      ...old,
      entries: { ...old.entries, [id]: { ...old.entries[id], name } },
    }));
  };
  const discard = (id: string) => {
    const entry = stateRef.current.entries[id];
    if (entry?.saving) {
      notify("Wait for the current save before discarding changes.");
      return;
    }
    if (entry?.saved) update(JSON.parse(entry.saved), id, false);
    else
      mutate((old) => {
        const entries = { ...old.entries };
        delete entries[id];
        return { ...old, entries };
      });
    close(id);
  };
  const saveAndClose = async (id: string) => {
    if (!(await save(id))) return false;
    if (
      stateRef.current.entries[id] &&
      dirty(stateRef.current.entries[id]) &&
      !(await save(id))
    )
      return false;
    if (stateRef.current.entries[id] && dirty(stateRef.current.entries[id])) {
      notify("New edits arrived while saving. Save again before closing.");
      return false;
    }
    close(id);
    return true;
  };
  const reloadSaved = async (id: string) => {
    const previous = stateRef.current.entries[id];
    if (!previous || previous.saving) return false;
    try {
      const records = await listDocuments(),
        record = records.find((r) => r.id === id);
      if (!record || !isStoredDocument(record))
        throw new Error("No saved version is available for this file.");
      if (stateRef.current.entries[id] !== previous)
        throw new Error(
          "The document changed while loading. Try again to replace the latest edits.",
        );
      mutate((old) => ({
        ...old,
        entries: {
          ...old.entries,
          [id]: {
            ...record,
            saved: fingerprint(record.document),
            saving: false,
            seeded: false,
          },
        },
      }));
      delete histories.current[id];
      setConnected(true);
      notify(`Reloaded ${record.name} from revision ${record.revision}.`);
      return true;
    } catch (error) {
      notify((error as Error).message, true);
      return false;
    }
  };
  return {
    ...state,
    ready,
    connected,
    recoveryError,
    dirty,
    refresh,
    open,
    close,
    add,
    update,
    undo,
    save,
    saveAndClose,
    rename,
    discard,
    reloadSaved,
  };
}
