import type { OfficeDocument } from "../types/office";

export type PluginPermission = "document:read" | "document:write";

export interface DocumentContext {
  schemaVersion: 1;
  source: "current-editor";
  document: { id: string; title: string; type: OfficeDocument["type"] };
  summary: string;
  statistics: Record<string, number>;
  outline: string[];
  text: string;
  truncated: boolean;
}

export interface PluginCommandResult {
  message: string;
  document?: OfficeDocument;
  context?: DocumentContext;
}

export interface PluginCommand {
  id: string;
  title: string;
  documentTypes: readonly OfficeDocument["type"][];
  run: (document: OfficeDocument) => PluginCommandResult;
}

/** Source-controlled plugins are trusted application code, not a sandbox. */
export interface StudioPlugin {
  id: string;
  name: string;
  version: string;
  description: string;
  permissions: readonly PluginPermission[];
  commands: readonly PluginCommand[];
}

export interface PluginStorage {
  getItem(key: string): string | null;
  setItem(key: string, value: string): void;
}

export interface InstalledPlugin extends StudioPlugin {
  enabled: boolean;
}

export interface PluginHost {
  register(plugin: StudioPlugin): () => void;
  list(): InstalledPlugin[];
  setEnabled(id: string, enabled: boolean): void;
  commands(
    document: OfficeDocument | null,
  ): Array<PluginCommand & { pluginId: string }>;
  execute(id: string, document: OfficeDocument): PluginCommandResult;
  subscribe(listener: () => void): () => void;
}
