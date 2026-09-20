import type { OfficeDocument } from "../types/office";
import type { PluginHost, PluginStorage, StudioPlugin } from "./types";

const STORAGE_KEY = "office-studio.plugins.v1";

function browserStorage(): PluginStorage | undefined {
  try {
    return typeof window === "undefined" ? undefined : window.localStorage;
  } catch {
    return undefined;
  }
}

export function createPluginHost(
  options: {
    plugins?: readonly StudioPlugin[];
    storage?: PluginStorage | null;
  } = {},
): PluginHost {
  const registry = new Map<string, StudioPlugin>();
  const listeners = new Set<() => void>();
  const storage =
    options.storage === null
      ? undefined
      : (options.storage ?? browserStorage());
  let enabled: Record<string, boolean> = Object.create(null);
  try {
    const saved = JSON.parse(storage?.getItem(STORAGE_KEY) ?? "{}");
    if (saved && typeof saved === "object" && !Array.isArray(saved)) {
      for (const [id, value] of Object.entries(saved))
        if (typeof value === "boolean") enabled[id] = value;
    }
  } catch {
    /* Invalid or unavailable storage must not prevent editing. */
  }

  const emit = () => listeners.forEach((listener) => listener());
  const isEnabled = (id: string) => enabled[id] !== false;
  const host: PluginHost = {
    register(plugin) {
      if (!/^[a-z0-9][a-z0-9.-]+$/.test(plugin.id))
        throw new Error(
          "Plugin id must contain lowercase letters, numbers, dots, or hyphens.",
        );
      if (registry.has(plugin.id))
        throw new Error(`Plugin ${plugin.id} is already registered.`);
      if (!plugin.permissions.includes("document:read"))
        throw new Error("Plugins must declare document:read.");
      const ids = new Set(
        [...registry.values()].flatMap((item) =>
          item.commands.map((command) => command.id),
        ),
      );
      for (const command of plugin.commands) {
        if (!command.id.startsWith(`${plugin.id}.`))
          throw new Error(
            `Command ${command.id} must start with ${plugin.id}.`,
          );
        if (ids.has(command.id))
          throw new Error(`Duplicate command ${command.id}.`);
        ids.add(command.id);
      }
      registry.set(plugin.id, plugin);
      emit();
      return () => {
        registry.delete(plugin.id);
        emit();
      };
    },
    list: () =>
      [...registry.values()].map((plugin) => ({
        ...plugin,
        enabled: isEnabled(plugin.id),
      })),
    setEnabled(id, value) {
      if (!registry.has(id)) throw new Error(`Unknown plugin ${id}.`);
      enabled[id] = value;
      try {
        storage?.setItem(STORAGE_KEY, JSON.stringify(enabled));
      } catch {
        /* Session state still applies. */
      }
      emit();
    },
    commands(document) {
      if (!document) return [];
      return [...registry.values()]
        .filter((plugin) => isEnabled(plugin.id))
        .flatMap((plugin) =>
          plugin.commands
            .filter((command) => command.documentTypes.includes(document.type))
            .map((command) => ({ ...command, pluginId: plugin.id })),
        );
    },
    execute(id, document) {
      const plugin = [...registry.values()].find((item) =>
        item.commands.some((command) => command.id === id),
      );
      if (!plugin) throw new Error(`Unknown plugin command ${id}.`);
      if (!isEnabled(plugin.id)) throw new Error(`${plugin.name} is disabled.`);
      const command = plugin.commands.find((item) => item.id === id)!;
      if (!command.documentTypes.includes(document.type))
        throw new Error(
          `${command.title} does not support ${document.type} documents.`,
        );
      // Clone before executing so a failed or read-only plugin cannot mutate editor state.
      const copy = JSON.parse(JSON.stringify(document)) as OfficeDocument;
      const result = command.run(copy);
      if (!result || typeof result.message !== "string")
        throw new Error("Plugin returned an invalid command result.");
      if (result.document) {
        if (!plugin.permissions.includes("document:write"))
          throw new Error(
            `${plugin.name} does not have document:write permission.`,
          );
        if (
          result.document.type !== document.type ||
          result.document.data.id !== document.data.id
        ) {
          throw new Error(
            "A plugin may update the current document, but cannot replace its identity or type.",
          );
        }
      }
      return result;
    },
    subscribe(listener) {
      listeners.add(listener);
      return () => {
        listeners.delete(listener);
      };
    },
  };
  options.plugins?.forEach((plugin) => host.register(plugin));
  return host;
}
