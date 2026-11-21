/**
 * Toolbar popup: new document / spreadsheet, open a file (in the editor
 * tab — its picker keeps a file handle so Save writes back in place, and a
 * native file dialog would close this popup anyway), open from URL, recent
 * files, backend status, one-click web-integration opt-in, settings.
 */

import { listRecent, deleteRecent, clearRecents } from "../lib/handoff.js";
import { createBackendClient, loadSettings } from "../lib/backend.js";
import { formatOfName, editorKindOf } from "../lib/formats.js";

const P = globalThis.OfficeExtPolicy;
const $ = (id) => document.getElementById(id);
const els = {
  dot: $("backend-dot"), backendText: $("backend-text"), backendHelp: $("backend-help"),
  card: $("integration-card"), btnEnableWeb: $("btn-enable-web"), btnDismissWeb: $("btn-dismiss-web"),
  btnNewWord: $("btn-new-word"), btnNewExcel: $("btn-new-excel"),
  btnNewSlides: $("btn-new-slides"), btnNewLatex: $("btn-new-latex"),
  btnOpen: $("btn-open"), btnUrl: $("btn-url"), urlForm: $("url-form"), urlInput: $("url-input"), urlError: $("url-error"),
  btnSettings: $("btn-settings"), recentList: $("recent-list"), btnClear: $("btn-clear-recents"),
  status: $("status"),
};

function viewerUrl(params) {
  return chrome.runtime.getURL(`viewer/viewer.html${params ? `?${params}` : ""}`);
}

async function openViewer(params) {
  try {
    await chrome.tabs.create({ url: viewerUrl(params) });
    window.close();
  } catch (e) {
    setStatus(`Could not open the editor: ${e.message || e}`, true);
  }
}

function setStatus(msg, isError = false) {
  els.status.textContent = msg || "";
  els.status.className = isError ? "error" : "";
}

// ------------------------------------------------------------- backend status

async function refreshBackend() {
  const settings = await loadSettings();
  els.backendHelp.hidden = true;
  if (!settings.backend_enabled) {
    els.dot.className = "dot";
    els.backendText.textContent = "Backend: off (built-in engines only)";
    return;
  }
  els.dot.className = "dot checking";
  els.backendText.textContent = "Backend: checking…";
  const client = createBackendClient(settings.backend_url);
  const h = await client.health(2500);
  if (h.ok) {
    els.dot.className = "dot ok";
    els.backendText.textContent = `Backend: online${h.version ? ` (v${h.version})` : ""}`;
    els.backendText.title = client.base;
  } else {
    els.dot.className = "dot fail";
    els.backendText.textContent = "Backend: offline — .doc/.rtf/.odt need it";
    els.backendText.title = `${client.base}: ${h.error || "unreachable"}`;
    els.backendHelp.hidden = false;
  }
}

// ------------------------------------------------------------- web integration opt-in

async function refreshIntegrationCard() {
  const [settings, granted, local] = await Promise.all([
    loadSettings(),
    chrome.permissions.contains({ origins: P.WEB_ORIGINS }).catch(() => false),
    chrome.storage.local.get("integration_prompt_dismissed"),
  ]);
  const s = P.integrationSettings(settings);
  els.card.hidden = granted || !s.intercept_downloads || !!local.integration_prompt_dismissed;
}

async function enableWebIntegration() {
  try {
    // permissions.request needs this click's user gesture. If the browser
    // closes the popup while it asks, the grant still applies: the service
    // worker reacts to permissions.onAdded.
    const ok = await chrome.permissions.request({ origins: P.WEB_ORIGINS });
    setStatus(ok ? "Document links will now open in the editor." : "Permission not granted.", !ok);
  } catch (e) {
    setStatus(`Could not request permission: ${e.message || e}`, true);
  }
  refreshIntegrationCard();
}

// ------------------------------------------------------------- open from URL

function toggleUrlForm(show) {
  const open = show ?? els.urlForm.hidden;
  els.urlForm.hidden = !open;
  els.btnUrl.setAttribute("aria-expanded", String(open));
  if (open) els.urlInput.focus();
}

function showUrlError(msg) {
  els.urlError.textContent = msg || "";
  els.urlError.hidden = !msg;
  els.urlInput.setAttribute("aria-invalid", msg ? "true" : "false");
}

async function submitUrl(e) {
  e.preventDefault();
  const v = P.validateUserUrl(els.urlInput.value);
  if (!v.ok) {
    showUrlError(v.error);
    return;
  }
  showUrlError("");
  await openViewer(`url=${encodeURIComponent(v.url)}`);
}

// ------------------------------------------------------------- recents

function fmtSize(n) {
  if (!n) return "";
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`;
  return `${(n / 1024 / 1024).toFixed(1)} MB`;
}

function iconFor(name) {
  const fmt = formatOfName(name);
  const kind = fmt ? editorKindOf(fmt) : null;
  if (kind === "excel") return { kind: "excel", label: "X" };
  if (kind === "slides") return { kind: "slides", label: "P" };
  if (kind === "latex") return { kind: "latex", label: "T" };
  if (kind === "word") return { kind: "word", label: "W" };
  return { kind: "", label: "T" };
}

async function renderRecents() {
  const items = await listRecent(); // newest first
  els.recentList.textContent = "";
  els.btnClear.hidden = !items.length;
  if (!items.length) {
    const li = document.createElement("li");
    li.className = "muted";
    li.textContent = "No recent files";
    els.recentList.appendChild(li);
    return;
  }
  const now = Date.now();
  for (const it of items) {
    const li = document.createElement("li");
    li.className = "recent-item";

    const open = document.createElement("button");
    open.className = "recent-open";
    open.title = `Open ${it.name}`;
    const ic = iconFor(it.name);
    const icon = document.createElement("span");
    icon.className = `ficon ${ic.kind}`;
    icon.textContent = ic.label;
    icon.setAttribute("aria-hidden", "true");
    const text = document.createElement("span");
    text.className = "ftext";
    const fname = document.createElement("span");
    fname.className = "fname";
    fname.textContent = it.name;
    const meta = document.createElement("span");
    meta.className = "meta";
    const created = it.created ? new Date(it.created) : null;
    meta.textContent = [fmtSize(it.size), P.relativeTime(it.created, now)].filter(Boolean).join(" · ");
    if (created) meta.title = created.toLocaleString();
    text.append(fname, meta);
    open.append(icon, text);
    open.addEventListener("click", () => openViewer(`recent=${encodeURIComponent(it.id)}`));

    const del = document.createElement("button");
    del.className = "del";
    del.textContent = "×";
    del.title = `Remove ${it.name} from recent files`;
    del.setAttribute("aria-label", del.title);
    del.addEventListener("click", async () => {
      await deleteRecent(it.id);
      await renderRecents();
      els.recentList.querySelector("button")?.focus();
    });

    li.append(open, del);
    els.recentList.appendChild(li);
  }
}

// ------------------------------------------------------------- init

function openBackendHelp(e) {
  e.preventDefault();
  chrome.tabs.create({ url: chrome.runtime.getURL("options/options.html#backend") });
  window.close();
}

async function init() {
  els.btnNewWord.addEventListener("click", () => openViewer("new=word"));
  els.btnNewExcel.addEventListener("click", () => openViewer("new=excel"));
  els.btnNewSlides.addEventListener("click", () => openViewer("new=slides"));
  els.btnNewLatex.addEventListener("click", () => openViewer("new=latex"));
  // The editor tab shows its Open dialog / drop zone (see viewer ?open=1).
  els.btnOpen.addEventListener("click", () => openViewer("open=1"));
  els.btnUrl.addEventListener("click", () => toggleUrlForm());
  els.urlForm.addEventListener("submit", submitUrl);
  els.urlInput.addEventListener("input", () => showUrlError(""));
  els.btnSettings.addEventListener("click", () => {
    chrome.runtime.openOptionsPage();
    window.close();
  });
  els.backendHelp.addEventListener("click", openBackendHelp);
  els.btnEnableWeb.addEventListener("click", enableWebIntegration);
  els.btnDismissWeb.addEventListener("click", async () => {
    await chrome.storage.local.set({ integration_prompt_dismissed: true });
    els.card.hidden = true;
  });
  els.btnClear.addEventListener("click", async () => {
    await clearRecents();
    await renderRecents();
    setStatus("Recent files cleared.");
  });
  await Promise.all([
    refreshBackend().catch((e) => setStatus(String(e.message || e), true)),
    renderRecents().catch((e) => setStatus(`Recent files unavailable: ${e.message || e}`, true)),
    refreshIntegrationCard().catch(() => {}),
  ]);
}

init();
