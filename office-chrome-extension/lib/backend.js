/**
 * Client for the local Rust office backend (office-backend), HTTP API v2:
 *
 *   GET  /health      -> { ok, name, version, api, parse:[fmt], render:{word:[fmt], excel:[fmt]} }
 *   POST /api/parse   octet-stream body, X-File-Name: encodeURIComponent(name)
 *                     -> { kind, format, model, warnings:[string] }
 *   POST /api/render  JSON { kind, model, format, name } -> file bytes
 *                     (Content-Disposition with filename*=UTF-8''…, X-Output-Format)
 *   errors            JSON { error } with a 4xx/5xx status
 *
 * Pure fetch-based; settings are injected so this module is Node-testable.
 */

import { LEGACY_BACKEND_RENDER } from "./formats.js";

export const DEFAULT_BACKEND_URL = "http://127.0.0.1:47831";
export const DEFAULT_SETTINGS = {
  backend_enabled: true,
  backend_url: DEFAULT_BACKEND_URL,
  intercept_downloads: true,
  // Opt-in page integrations (see content/policy.js integrationSettings).
  intercept_text: false,
  intercept_drops: false,
};

export function normalizeBackendUrl(url) {
  let u = String(url || DEFAULT_BACKEND_URL).trim();
  if (!u) u = DEFAULT_BACKEND_URL;
  if (!/^https?:\/\//i.test(u)) u = `http://${u}`;
  return u.replace(/\/+$/, "");
}

/**
 * True when the backend URL points at this machine. Documents are uploaded
 * to the backend for parsing, so a remote URL deserves a visible warning.
 */
export function isLoopbackBackendUrl(url) {
  try {
    const h = new URL(normalizeBackendUrl(url)).hostname.toLowerCase();
    return h === "127.0.0.1" || h === "localhost" || h === "[::1]" || h === "::1" || /^127\.\d+\.\d+\.\d+$/.test(h);
  } catch {
    return false;
  }
}

/** Load settings from chrome.storage.local (extension contexts only). */
export async function loadSettings() {
  if (typeof chrome === "undefined" || !chrome.storage?.local) {
    return { ...DEFAULT_SETTINGS };
  }
  const data = await chrome.storage.local.get("settings");
  return { ...DEFAULT_SETTINGS, ...(data.settings || {}) };
}

export async function saveSettings(settings) {
  if (typeof chrome === "undefined" || !chrome.storage?.local) return;
  await chrome.storage.local.set({ settings });
}

/**
 * File name from a Content-Disposition header: RFC 5987 filename*= wins
 * (percent-decoded; a malformed encoding is kept literally), then the plain
 * filename= value (never percent-decoded). Returns "" when absent.
 */
export function filenameFromContentDisposition(cd) {
  const s = String(cd || "");
  const star = /filename\*\s*=\s*([^;]+)/i.exec(s);
  if (star) {
    let v = star[1].trim().replace(/^"(.*)"$/, "$1");
    const m = /^([\w!#$%&+^`{}~-]*)'[^']*'(.*)$/.exec(v);
    const charset = m ? m[1].toLowerCase() : "utf-8";
    if (m) v = m[2];
    try {
      if (charset === "utf-8" || charset === "") return decodeURIComponent(v);
      // iso-8859-1: percent bytes are code points
      return v.replace(/%([0-9a-f]{2})/gi, (_, h) => String.fromCharCode(parseInt(h, 16)));
    } catch {
      return v;
    }
  }
  const plain = /filename\s*=\s*("((?:\\.|[^"\\])*)"|[^;]+)/i.exec(s);
  if (plain) {
    const v = plain[2] != null ? plain[2].replace(/\\(.)/g, "$1") : plain[1].trim();
    return v.trim();
  }
  return "";
}

/** The backend's render formats for a kind, from a /health result. */
export function backendRenderFormats(health, kind) {
  if (!health || !health.ok) return [];
  if (health.render && Array.isArray(health.render[kind])) return health.render[kind].map(String);
  // Pre-v2 backends did not advertise capabilities; they rendered these.
  return LEGACY_BACKEND_RENDER[kind] ? [...LEGACY_BACKEND_RENDER[kind]] : [];
}

/** The backend's parse formats from a /health result (null = unknown, try anything). */
export function backendParseFormats(health) {
  if (!health || !health.ok) return [];
  return Array.isArray(health.parse) ? health.parse.map(String) : null;
}

export function createBackendClient(baseUrl, fetchImpl = (...a) => fetch(...a)) {
  const base = normalizeBackendUrl(baseUrl);

  async function req(path, { method = "GET", headers = {}, body, timeoutMs = 30000 } = {}) {
    const ctrl = new AbortController();
    const timer = setTimeout(() => ctrl.abort(), timeoutMs);
    try {
      return await fetchImpl(`${base}${path}`, { method, headers, body, signal: ctrl.signal });
    } catch (e) {
      // Network-level failure: make the message actionable and flag it so
      // callers can tell "backend unreachable" from "backend said no".
      const timeout = e && e.name === "AbortError";
      const err = new BackendError(0, timeout
        ? `The office backend at ${base} did not answer in time.`
        : `Cannot reach the office backend at ${base}. Is office-backend running?`);
      err.network = true;
      err.cause = e;
      throw err;
    } finally {
      clearTimeout(timer);
    }
  }

  return {
    base,

    /** /health JSON plus ok:true, or { ok:false, error }. Never throws. */
    async health(timeoutMs = 2500) {
      try {
        const res = await req("/health", { timeoutMs });
        if (!res.ok) return { ok: false, error: `HTTP ${res.status}` };
        const json = await res.json();
        if (!json || typeof json !== "object") return { ok: false, error: "unexpected /health response" };
        return { ...json, ok: json.ok !== false };
      } catch (e) {
        return { ok: false, error: String((e && e.message) || e) };
      }
    },

    /**
     * Parse file bytes into the document model.
     * bytes: Uint8Array | ArrayBuffer, name: file name for format detection
     * (sent percent-encoded, so any Unicode name is safe in the header).
     * Returns { kind, format, model, warnings } or throws BackendError.
     */
    async parse(bytes, name, timeoutMs = 120000) {
      const buf = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
      const res = await req("/api/parse", {
        method: "POST",
        headers: {
          "Content-Type": "application/octet-stream",
          "X-File-Name": encodeURIComponent(name || "document"),
        },
        body: buf,
        timeoutMs,
      });
      if (!res.ok) throw new BackendError(res.status, await safeError(res));
      const json = await res.json();
      if (!Array.isArray(json.warnings)) json.warnings = [];
      return json;
    },

    /**
     * Render a document model back to file bytes.
     * Returns { blob, filename, format }. filename prefers filename*=.
     */
    async render(model, format, name, timeoutMs = 120000) {
      const res = await req("/api/render", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ kind: model.kind, model, format, name: name || "document" }),
        timeoutMs,
      });
      if (!res.ok) throw new BackendError(res.status, await safeError(res));
      const filename = filenameFromContentDisposition(res.headers.get("Content-Disposition")) || name || `document.${format}`;
      const blob = await res.blob();
      const outFormat = res.headers.get("X-Output-Format") || format;
      return { blob, filename, format: outFormat, warnings: parseWarningsHeader(res.headers.get("X-Office-Warnings")) };
    },
  };
}

/**
 * X-Office-Warnings: encodeURIComponent(JSON array of strings) with render
 * notes (renamed sheets, truncated text, only the active sheet in CSV, ...).
 * Malformed values are ignored.
 */
export function parseWarningsHeader(value) {
  if (!value) return [];
  try {
    const list = JSON.parse(decodeURIComponent(value));
    return Array.isArray(list) ? list.filter((w) => typeof w === "string" && w) : [];
  } catch {
    return [];
  }
}

export class BackendError extends Error {
  constructor(status, detail) {
    super(detail || `Backend error ${status}`);
    this.name = "BackendError";
    this.status = status;
    this.detail = detail;
    this.network = false;
  }
}

async function safeError(res) {
  let text = "";
  try { text = await res.text(); } catch { /* ignore */ }
  try {
    const j = JSON.parse(text);
    if (j && (j.error || j.detail)) return String(j.error || j.detail);
  } catch { /* not JSON */ }
  if (res.status === 413) return "The file is too large for the backend (limit 100 MB).";
  return text.slice(0, 300) || `HTTP ${res.status}`;
}

export function isBackendError(e) {
  return e instanceof BackendError || (e && e.name === "BackendError");
}
