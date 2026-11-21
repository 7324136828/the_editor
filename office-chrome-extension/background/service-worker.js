/**
 * Office Chrome Extension — background service worker (MV3).
 *
 * Responsibilities
 *  - Web integration (only when the user granted the optional http/https
 *    host permission and the matching setting is on):
 *      * registers the link/drop content script,
 *      * opens Office documents from link clicks (content script asks us),
 *      * opens Office downloads in the editor instead of saving them.
 *  - "Open in Office Editor" context menu on document links.
 *  - Relays dropped-file chunks from content scripts to the offscreen
 *    document, which owns OPFS writes and does the fetching.
 *
 * Download interception never races the real download: we hook
 * downloads.onDeterminingFilename, which holds the download before its file
 * name is decided (so it cannot finish), fetch the file ourselves, and only
 * then cancel it. When anything goes wrong we release the hold and the
 * browser's own download proceeds untouched.
 */

import "../content/policy.js";
import { loadSettings } from "../lib/backend.js";

const P = globalThis.OfficeExtPolicy;
const VIEWER = "viewer/viewer.html";
const CONTENT_SCRIPT_ID = "office-integration";
const MENU_LINK = "office-open-link";
const MENU_PAGE = "office-open-page";

// ---------------------------------------------------------------- settings

let settingsCache = null;

async function getSettings() {
  if (settingsCache) return settingsCache;
  try {
    settingsCache = await loadSettings();
  } catch {
    settingsCache = {};
  }
  return settingsCache;
}

async function integration() {
  return P.integrationSettings(await getSettings());
}

async function hasWebAccess() {
  try {
    return await chrome.permissions.contains({ origins: P.WEB_ORIGINS });
  } catch {
    return false;
  }
}

async function hasFileAccess() {
  try {
    return await chrome.extension.isAllowedFileSchemeAccess();
  } catch {
    return false;
  }
}

// ---------------------------------------------------------------- tabs

function viewerUrl(params = "") {
  return chrome.runtime.getURL(`${VIEWER}${params ? "?" + params : ""}`);
}

/** Open a viewer tab next to `tab` (the page the request came from), if known. */
async function openViewerTab(params, tab) {
  const props = { url: viewerUrl(params), active: true };
  let anchor = tab;
  if (!anchor) {
    try {
      [anchor] = await chrome.tabs.query({ active: true, lastFocusedWindow: true });
    } catch {
      anchor = null;
    }
  }
  if (anchor && typeof anchor.id === "number" && anchor.id >= 0) {
    props.openerTabId = anchor.id;
    props.index = anchor.index + 1;
    props.windowId = anchor.windowId;
  }
  try {
    return await chrome.tabs.create(props);
  } catch {
    return chrome.tabs.create({ url: props.url, active: true });
  }
}

// Never let a page (or a burst of downloads) open a storm of tabs.
const openLimiter = P.createRateLimiter({ max: 6, windowMs: 15000, perKeyMs: 1500 });

// ---------------------------------------------------------------- bypass list
// URLs the user asked to download normally (Alt-click, right-click → "Save
// link as…", or a failed open we fell back from). Stored in session storage
// so a service-worker restart does not lose them.

const BYPASS_TTL_MS = 2 * 60 * 1000;
let bypass = null; // Map url -> expiry

async function loadBypass() {
  if (bypass) return bypass;
  bypass = new Map();
  try {
    const d = await chrome.storage.session.get("bypass");
    for (const [u, exp] of Object.entries(d.bypass || {})) bypass.set(u, exp);
  } catch {
    /* session storage unavailable */
  }
  return bypass;
}

async function persistBypass() {
  const now = Date.now();
  const obj = {};
  for (const [u, exp] of bypass) if (exp > now) obj[u] = exp;
  try {
    await chrome.storage.session.set({ bypass: obj });
  } catch {
    /* ignore */
  }
}

async function addBypass(url) {
  if (!url) return;
  await loadBypass();
  bypass.set(String(url), Date.now() + BYPASS_TTL_MS);
  if (bypass.size > 200) bypass.delete(bypass.keys().next().value);
  await persistBypass();
}

/** True (and consumes the entry) when `url` should not be intercepted. */
function takeBypass(url) {
  if (!bypass) return false;
  const exp = bypass.get(url);
  if (!exp) return false;
  bypass.delete(url);
  persistBypass();
  return exp > Date.now();
}

// ---------------------------------------------------------------- offscreen

let offscreenCreating = null;

async function ensureOffscreen() {
  if (await chrome.offscreen.hasDocument?.()) return;
  if (!offscreenCreating) {
    offscreenCreating = chrome.offscreen
      .createDocument({
        url: "offscreen/offscreen.html",
        reasons: ["LOCAL_STORAGE", "BLOBS"],
        justification:
          "Fetch Office documents the user chose to open and stage them in the origin-private file system for the editor tab.",
      })
      .catch((e) => {
        // A concurrent creation already succeeded.
        if (!/single offscreen/i.test(String(e && e.message))) throw e;
      })
      .finally(() => {
        offscreenCreating = null;
      });
  }
  await offscreenCreating;
}

async function offscreenCall(type, payload = {}) {
  await ensureOffscreen();
  const res = await chrome.runtime.sendMessage({ target: "office-offscreen", type, ...payload });
  if (!res) throw new Error("the background page did not answer");
  return res;
}

/** Fetch `url` in the offscreen document and stage it; { ok, id, name } or { ok:false, error }. */
async function fetchToHandoff(url, nameHint) {
  try {
    return await offscreenCall("office-fetch-store", { url, nameHint: nameHint || "" });
  } catch (e) {
    return { ok: false, error: String((e && e.message) || e) };
  }
}

// ---------------------------------------------------------------- content script registration

/**
 * The content script only runs when it has something to do: web access was
 * granted and link interception or drop handling is enabled.
 */
async function syncContentScript() {
  const s = await integration();
  const want = (s.intercept_downloads || s.intercept_drops) && (await hasWebAccess());
  let registered = [];
  try {
    registered = await chrome.scripting.getRegisteredContentScripts({ ids: [CONTENT_SCRIPT_ID] });
  } catch {
    registered = [];
  }
  if (want && !registered.length) {
    try {
      await chrome.scripting.registerContentScripts([
        {
          id: CONTENT_SCRIPT_ID,
          matches: ["http://*/*", "https://*/*"],
          js: ["content/policy.js", "content/content.js"],
          runAt: "document_idle",
          allFrames: false,
          persistAcrossSessions: true,
        },
      ]);
    } catch (e) {
      if (!/duplicate/i.test(String(e && e.message))) console.warn("[office-ext] register content script:", e);
    }
    await injectIntoOpenTabs();
  } else if (!want && registered.length) {
    try {
      await chrome.scripting.unregisterContentScripts({ ids: [CONTENT_SCRIPT_ID] });
    } catch {
      /* already gone */
    }
  }
}

/**
 * Pages that were open before the script was registered (or before an
 * extension update orphaned the old one) get it injected now, so the
 * feature works without reloading every tab.
 */
async function injectIntoOpenTabs() {
  let tabs = [];
  try {
    tabs = await chrome.tabs.query({ url: ["http://*/*", "https://*/*"] });
  } catch {
    return;
  }
  for (const t of tabs.slice(0, 200)) {
    if (t.discarded || typeof t.id !== "number") continue;
    chrome.scripting
      .executeScript({ target: { tabId: t.id }, files: ["content/policy.js", "content/content.js"] })
      .catch(() => {}); // restricted pages (web store, error pages) refuse
  }
}

// ---------------------------------------------------------------- context menu

function setupContextMenus() {
  chrome.contextMenus.removeAll(() => {
    void chrome.runtime.lastError;
    chrome.contextMenus.create(
      {
        id: MENU_LINK,
        title: "Open in Office Editor",
        contexts: ["link"],
        targetUrlPatterns: P.linkMenuPatterns(),
      },
      () => {
        if (chrome.runtime.lastError) console.warn("[office-ext] context menu:", chrome.runtime.lastError.message);
      }
    );
    chrome.contextMenus.create(
      {
        id: MENU_PAGE,
        title: "Open this file in Office Editor",
        contexts: ["page"],
        documentUrlPatterns: P.pageMenuPatterns(),
      },
      () => {
        if (chrome.runtime.lastError) console.warn("[office-ext] context menu:", chrome.runtime.lastError.message);
      }
    );
  });
}

/** Explicit user request: open `url` in the editor whatever the settings. */
async function openUrlExplicitly(url, tab) {
  const isFile = /^file:/i.test(url);
  const canFetch = isFile ? await hasFileAccess() : await hasWebAccess();
  if (canFetch) {
    const res = await fetchToHandoff(url, "");
    if (res.ok) {
      await openViewerTab(`handoff=${encodeURIComponent(res.id)}`, tab);
      return;
    }
  }
  // Let the viewer fetch it and explain any failure (sign-in page, CORS…).
  await openViewerTab(`url=${encodeURIComponent(url)}`, tab);
}

chrome.contextMenus.onClicked.addListener((info, tab) => {
  const url = info.menuItemId === MENU_LINK ? info.linkUrl : info.menuItemId === MENU_PAGE ? info.pageUrl : "";
  if (!url) return;
  openUrlExplicitly(url, tab).catch((e) => console.warn("[office-ext] open from menu failed:", e));
});

// ---------------------------------------------------------------- messages

function isFromOwnContentScript(sender) {
  return sender && sender.id === chrome.runtime.id && sender.tab && typeof sender.tab.id === "number" &&
    /^https?:/i.test(String(sender.url || sender.tab.url || "https:"));
}

/** Drop relays in progress: token -> { tabId, uploadId, name, expiresAt }. */
const drops = new Map();

function newToken() {
  const a = new Uint8Array(16);
  crypto.getRandomValues(a);
  return [...a].map((b) => b.toString(16).padStart(2, "0")).join("");
}

const CONTENT_HANDLERS = {
  /** A plain left click on a document link (content script already called preventDefault). */
  async "office-open-url"(msg, sender) {
    const url = String(msg.url || "");
    const s = await integration();
    // Safety net: the content script checks these too.
    if (!P.linkEligibility({ href: url }, s) || !(await hasWebAccess())) {
      await addBypass(url);
      return { ok: false, fallback: true, error: "interception is off" };
    }
    if (!openLimiter.allow(sender.tab.id)) {
      await addBypass(url);
      return { ok: false, fallback: true, error: "too many requests" };
    }
    const res = await fetchToHandoff(url, "");
    if (!res.ok) {
      // Not a document (sign-in page, HTML viewer, 404 …): navigate normally,
      // and do not intercept the download that navigation may start.
      await addBypass(url);
      return { ok: false, fallback: true, error: res.error };
    }
    await openViewerTab(`handoff=${encodeURIComponent(res.id)}`, sender.tab);
    return { ok: true };
  },

  /** Alt-click / right-click on a document link: the user wants the real file. */
  async "office-bypass-url"(msg) {
    await addBypass(String(msg.url || ""));
    return { ok: true };
  },

  async "office-drop-begin"(msg, sender) {
    const s = await integration();
    if (!s.intercept_drops) return { ok: false, error: "Opening dropped files is turned off in the settings." };
    const size = Number(msg.size);
    if (!(size > 0)) return { ok: false, error: "The file is empty." };
    if (size > P.MAX_HANDOFF_BYTES) {
      return { ok: false, error: `The file is larger than ${Math.round(P.MAX_HANDOFF_BYTES / 1048576)} MB.` };
    }
    if (!openLimiter.allow(`drop:${sender.tab.id}:${msg.name}`)) return { ok: false, error: "Too many files at once." };
    const res = await offscreenCall("office-upload-begin", { name: String(msg.name || "document"), mime: String(msg.mime || ""), size });
    if (!res.ok) return res;
    const token = newToken();
    drops.set(token, { tabId: sender.tab.id, uploadId: res.uploadId, expiresAt: Date.now() + 5 * 60 * 1000 });
    return { ok: true, token };
  },

  async "office-drop-chunk"(msg, sender) {
    const d = drops.get(msg.token);
    if (!d || d.tabId !== sender.tab.id || d.expiresAt < Date.now()) return { ok: false, error: "upload expired" };
    return offscreenCall("office-upload-chunk", { uploadId: d.uploadId, seq: msg.seq, b64: msg.b64 });
  },

  async "office-drop-end"(msg, sender) {
    const d = drops.get(msg.token);
    if (!d || d.tabId !== sender.tab.id) return { ok: false, error: "upload expired" };
    drops.delete(msg.token);
    const res = await offscreenCall("office-upload-end", { uploadId: d.uploadId });
    if (!res.ok) return res;
    await openViewerTab(`handoff=${encodeURIComponent(res.id)}`, sender.tab);
    return { ok: true };
  },

  async "office-drop-abort"(msg, sender) {
    const d = drops.get(msg.token);
    if (!d || d.tabId !== sender.tab.id) return { ok: true };
    drops.delete(msg.token);
    await offscreenCall("office-upload-abort", { uploadId: d.uploadId }).catch(() => {});
    return { ok: true };
  },
};

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (!msg || typeof msg !== "object") return false;
  if (msg.target === "office-offscreen") return false; // addressed to the offscreen document
  const handler = CONTENT_HANDLERS[msg.type];
  if (!handler) return false;
  // Only our own content scripts on web pages may use these.
  if (!isFromOwnContentScript(sender)) return false;
  handler(msg, sender)
    .then((r) => sendResponse(r))
    .catch((e) => sendResponse({ ok: false, fallback: true, error: String((e && e.message) || e) }));
  return true; // async response
});

// ---------------------------------------------------------------- download interception

/**
 * Called for every download before its target file name is decided. While
 * we have not called suggest(), the download cannot complete, so we can
 * fetch the file ourselves and cancel it only once the editor has the bytes.
 */
chrome.downloads.onDeterminingFilename.addListener((item, suggest) => {
  // Cheap synchronous pre-filter: leave obviously irrelevant downloads alone
  // without holding them at all.
  const url = String(item.finalUrl || item.url || "");
  if (item.byExtensionId === chrome.runtime.id) return;
  if (!/^(https?|file):/i.test(url)) return;
  const nameExt = P.extOf(String(item.filename || "").split(/[\\/]/).pop()) || P.extOf(P.nameFromUrl(url));
  if (!P.ALL_EXTS.includes(nameExt)) return;

  let released = false;
  const release = () => {
    if (released) return;
    released = true;
    try {
      suggest();
    } catch {
      /* download already gone */
    }
  };
  handleDownload(item)
    .then((taken) => {
      if (!taken) release();
    })
    .catch((e) => {
      console.warn("[office-ext] download interception failed:", e);
      release();
    });
  return true; // we will call suggest() asynchronously
});

/** Returns true when the download was taken over (cancelled + opened in the editor). */
async function handleDownload(item) {
  const s = await integration();
  if (!s.intercept_downloads) return false;
  await loadBypass();
  const url = String(item.finalUrl || item.url || "");
  const isFile = /^file:/i.test(url);
  const access = isFile ? await hasFileAccess() : await hasWebAccess();
  if (!access) return false;
  const verdict = P.downloadEligibility(item, s, {
    extensionId: chrome.runtime.id,
    fileAccess: isFile && access,
    isBypassed: (u) => takeBypass(u) || (item.url !== u && takeBypass(item.url)),
  });
  if (!verdict.intercept) return false;
  if (!openLimiter.allow(null)) return false;

  const res = await fetchToHandoff(verdict.url, verdict.name);
  if (!res.ok) return false; // the browser's own download goes ahead

  // The editor has the bytes: cancel the (still unnamed, hence unfinished)
  // download and remove its "cancelled" entry from the downloads list.
  let cancelled = false;
  try {
    await chrome.downloads.cancel(item.id);
    const [now] = await chrome.downloads.search({ id: item.id });
    cancelled = !now || now.state === "interrupted";
  } catch {
    cancelled = false;
  }
  if (!cancelled) {
    // Could not stop it: keep the user's download, drop our staged copy.
    offscreenCall("office-delete-handoff", { id: res.id }).catch(() => {});
    return false;
  }
  try {
    await chrome.downloads.erase({ id: item.id });
  } catch {
    /* ignore */
  }
  await openViewerTab(`handoff=${encodeURIComponent(res.id)}`, null);
  return true;
}

// ---------------------------------------------------------------- lifecycle

chrome.storage.onChanged.addListener((changes, area) => {
  if (area === "local" && changes.settings) {
    settingsCache = null; // settings take effect immediately
    syncContentScript().catch(() => {});
  }
});

chrome.permissions.onAdded.addListener(() => syncContentScript().catch(() => {}));
chrome.permissions.onRemoved.addListener(() => syncContentScript().catch(() => {}));

chrome.runtime.onInstalled.addListener(() => {
  setupContextMenus();
  syncContentScript().catch(() => {});
  ensureOffscreen()
    .then(() => offscreenCall("office-sweep"))
    .catch(() => {});
});

chrome.runtime.onStartup.addListener(() => {
  syncContentScript().catch(() => {});
  // Remove handoffs nobody consumed (age-based; in-flight ones are kept).
  ensureOffscreen()
    .then(() => offscreenCall("office-sweep"))
    .catch(() => {});
});
