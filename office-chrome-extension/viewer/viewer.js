/**
 * Office Editor viewer: the editor shell.
 *
 *  - open pipeline: picker (File System Access when available), drag & drop,
 *    paste, URL, handoff, recents, OS launchQueue, New. A file is parsed
 *    into local variables FIRST and only then swapped in, so a failed open
 *    never touches the current document; every path asks before discarding
 *    unsaved changes.
 *  - Word: always-editable rich editor over the word model (backend parser,
 *    or the client-side .docx reader when offline), formatting toolbar,
 *    optional docx-preview "Page view". Markdown opens as a formatted doc.
 *  - Excel: ExcelView grid (viewer/excel.js) over the excel model.
 *  - Text: plain-text editor with encoding detection.
 *  - Slides: SlidesView (viewer/slides.js) over the slides model (backend
 *    pptx reader, or lib/local-pptx.js offline); slideshow mode.
 *  - LaTeX: LatexView (viewer/latex.js): source editor + live preview over
 *    the .tex text; PDF through the backend when a TeX engine is installed.
 *  - Save: Ctrl+S writes back in place (File System Access handle) when the
 *    format is writable, otherwise Save As (name + format list computed from
 *    backend /health capabilities and local writers; lossy exports warn and
 *    keep the document marked unsaved).
 *  - Backend health is re-checked periodically, on focus and on settings change.
 */

import {
  FORMATS, resolveFormat, formatOfName, editorKindOf, withExtension, mimeOf, decodeText,
  encodeText, computeSaveTargets, pickerTypeFor, OPEN_EXTENSIONS, sniffDelimiter, mimeEssence,
  localCapabilities,
} from "../lib/formats.js";
import {
  createBackendClient, loadSettings, isLoopbackBackendUrl, backendRenderFormats,
  backendParseFormats, filenameFromContentDisposition, isBackendError,
} from "../lib/backend.js";
import { getHandoff, deleteHandoff, sweepHandoffs, putHandoff, addRecent, updateRecent, getRecent } from "../lib/handoff.js";
import * as WordModel from "../lib/model-word.js";
import * as ExcelModel from "../lib/model-excel.js";
import * as LocalXlsx from "../lib/local-xlsx.js";
import * as LocalDocx from "../lib/local-docx.js";
import { markdownToModel } from "../lib/markdown.js";
import { WORD_TEMPLATES, wordTemplate } from "../lib/templates-word.js";
import * as SlidesModel from "../lib/model-slides.js";
import * as LocalPptx from "../lib/local-pptx.js";
import * as Latex from "../lib/latex.js";
import { ExcelView } from "./excel.js";
import { SlidesView } from "./slides.js";
import { LatexView } from "./latex.js";

const { renderWordHtml, parseWordDom, wordModelToText } = WordModel;

const $ = (id) => document.getElementById(id);
const els = {
  toolbar: $("toolbar"),
  name: $("doc-name"), kind: $("doc-kind"), dirty: $("dirty-dot"),
  backendStatus: $("backend-status"), backendDot: $("backend-dot"), backendLabel: $("backend-label"),
  status: $("status-msg"), stats: $("doc-stats"),
  infoBar: $("info-bar"), infoIcon: $("info-icon"), infoText: $("info-text"), infoList: $("info-list"),
  infoActions: $("info-actions"), infoClose: $("info-close"),
  wordBar: $("word-bar"), excelBar: $("excel-bar"),
  stage: $("stage"), wordView: $("word-view"), wordPreview: $("word-preview"),
  excelView: $("excel-view"), textView: $("text-view"), printText: $("print-text"),
  slidesView: $("slides-view"), latexView: $("latex-view"),
  empty: $("empty-state"), dropHint: $("drop-hint"), busy: $("busy"), busyText: $("busy-text"),
  btnNew: $("btn-new"), menuNew: $("menu-new"), btnOpen: $("btn-open"), btnSave: $("btn-save"),
  btnSaveAs: $("btn-saveas"), btnPreview: $("btn-preview"), btnPrint: $("btn-print"),
  btnMore: $("btn-more"), menuMore: $("menu-more"),
  miOpenUrl: $("mi-open-url"), miDownloadOriginal: $("mi-download-original"),
  miShortcuts: $("mi-shortcuts"), miSettings: $("mi-settings"),
  emptyNewWord: $("empty-new-word"), emptyNewExcel: $("empty-new-excel"),
  emptyNewSlides: $("empty-new-slides"), emptyNewLatex: $("empty-new-latex"),
  btnOpen2: $("btn-open2"), btnOpenUrl: $("btn-open-url"), backendNote: $("backend-note"),
  fileInput: $("file-input"), imageInput: $("image-input"),
  dialogModal: $("dialog-modal"), dlgTitle: $("dlg-title"), dlgText: $("dlg-text"), dlgActions: $("dlg-actions"),
  saveasModal: $("saveas-modal"), saForm: $("saveas-form"), saName: $("sa-name"), saFormat: $("sa-format"),
  saReason: $("sa-reason"), saLossy: $("sa-lossy"), saWhere: $("sa-where"), saCancel: $("sa-cancel"),
  linkModal: $("link-modal"), lnForm: $("link-form"), lnTitle: $("ln-title"), lnText: $("ln-text"),
  lnUrl: $("ln-url"), lnError: $("ln-error"), lnRemove: $("ln-remove"), lnCancel: $("ln-cancel"),
  urlModal: $("url-modal"), urlForm: $("url-form"), urlInput: $("url-input"), urlError: $("url-error"),
  urlCancel: $("url-cancel"),
  shortcutsModal: $("shortcuts-modal"), ksClose: $("ks-close"),
  toast: $("toast"),
  fbStyle: $("fb-style"), fbFont: $("fb-font"), fbSize: $("fb-size"), fbColor: $("fb-color"),
  fbColorBar: $("fb-color-bar"), fbHighlight: $("fb-highlight"), fbIndent: $("fb-indent"),
  fbOutdent: $("fb-outdent"), fbTable: $("fb-table"), tablePicker: $("table-picker"),
  fbTableMenu: $("fb-table-menu"), menuTable: $("menu-table"),
  tpGrid: $("tp-grid"), tpLabel: $("tp-label"), tableTools: $("table-tools"),
  // word features
  wordSheet: $("word-sheet"), wordHeader: $("word-header"), wordFooter: $("word-footer"),
  hfTools: $("hf-tools"), hfToolsLabel: $("hf-tools-label"),
  fbInsert: $("fb-insert"), menuInsert: $("menu-insert"), fbFormat: $("fb-format"), menuFormat: $("menu-format"),
  fbLayout: $("fb-layout"), menuLayout: $("menu-layout"), fbView: $("fb-view"), menuView: $("menu-view"),
  fbPainter: $("fb-painter"), miNav: $("mi-nav"), miNewTemplate: $("mi-new-template"),
  miPageSetup: $("mi-page-setup"), miProperties: $("mi-properties"), emptyNewTemplate: $("empty-new-template"),
  navPane: $("nav-pane"), navList: $("nav-list"), navEmpty: $("nav-empty"), navClose: $("nav-close"),
  findBar: $("find-bar"), findInput: $("find-input"), findCount: $("find-count"), findCase: $("find-case"),
  findWord: $("find-word"), findPrev: $("find-prev"), findNext: $("find-next"), findClose: $("find-close"),
  findToggleReplace: $("find-toggle-replace"), findReplaceRow: $("find-replace-row"), replaceInput: $("replace-input"),
  replaceOne: $("replace-one"), replaceAll: $("replace-all"),
  zoomCtl: $("zoom-ctl"), zoomOut: $("zoom-out"), zoomIn: $("zoom-in"), zoomRange: $("zoom-range"), zoomLabel: $("zoom-label"),
  wcModal: $("wc-modal"), wcBody: $("wc-body"), wcClose: $("wc-close"),
  symModal: $("sym-modal"), symTabs: $("sym-tabs"), symGrid: $("sym-grid"), symPreview: $("sym-preview"),
  symName: $("sym-name"), symClose: $("sym-close"),
  dtModal: $("dt-modal"), dtForm: $("dt-form"), dtList: $("dt-list"), dtCancel: $("dt-cancel"),
  psModal: $("ps-modal"), psForm: $("ps-form"), psSize: $("ps-size"), psPortrait: $("ps-portrait"),
  psLandscape: $("ps-landscape"), psPreset: $("ps-preset"), psTop: $("ps-top"), psBottom: $("ps-bottom"),
  psLeft: $("ps-left"), psRight: $("ps-right"), psUnits: $("ps-units"), psPreview: $("ps-preview"),
  psError: $("ps-error"), psCancel: $("ps-cancel"),
  propsModal: $("props-modal"), propsForm: $("props-form"), propsTitle: $("props-doc-title"),
  propsAuthor: $("props-author"), propsSubject: $("props-subject"), propsCancel: $("props-cancel"),
  tplModal: $("tpl-modal"), tplGrid: $("tpl-grid"), tplOk: $("tpl-ok"), tplCancel: $("tpl-cancel"),
};

const APP = "Office Editor";
const RECENT_MAX_BYTES = 32 * 1024 * 1024;
const HEALTH_INTERVAL_MS = 15000;
const MAX_IMAGE_WIDTH = 624; // px: 6.5in text width at 96 dpi
const KIND_LABEL = { word: "Word", excel: "Excel", text: "Text", slides: "PowerPoint", latex: "LaTeX" };
const NEW_KINDS = ["word", "excel", "slides", "latex"];

/**
 * The open document (null when nothing is open).
 * {
 *   kind: "word"|"excel"|"text", format: source format, name, mime,
 *   bytes: original bytes (or null for new docs), handle: FileSystemFileHandle|null,
 *   model (word/excel), text/encoding/bom/eol (text),
 *   previewOnly: word shown read-only through docx-preview (no model),
 *   saveTarget: {format, name, handle} after a successful non-lossy save,
 *   recentId, engine, isNew
 * }
 */
let doc = null;
let dirty = false;
let excelView = null;
let slidesView = null;            // SlidesView while a presentation is open
let latexView = null;             // LatexView while a LaTeX document is open
let pageView = false;             // word: showing docx-preview of the current model
let busyCount = 0;
let toastTimer = null;
let statsTimer = null;

const backend = {
  settings: null,
  client: null,
  health: { ok: false, checking: true },
  timer: null,
  checking: null,
};

// ---------------------------------------------------------------------------
// small UI helpers
// ---------------------------------------------------------------------------

function toast(msg, type = "") {
  els.toast.textContent = msg;
  els.toast.className = "toast" + (type ? ` ${type}` : "");
  els.toast.hidden = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { els.toast.hidden = true; }, type === "error" ? 7000 : 3500);
}

function setStatus(msg, isError = false) {
  els.status.textContent = msg || "";
  els.status.className = isError ? "error" : "";
  els.status.title = msg || "";
}

function setBusy(on, text = "Working…") {
  busyCount = Math.max(0, busyCount + (on ? 1 : -1));
  els.busyText.textContent = text;
  els.busy.hidden = busyCount === 0;
}

function errMsg(e) {
  return String((e && e.message) || e || "Unknown error");
}

function updateTitle() {
  const name = doc ? doc.name : "";
  document.title = doc ? `${dirty ? "• " : ""}${name} — ${APP}` : APP;
  els.name.textContent = doc ? name : APP;
  els.name.title = doc ? name : "";
  els.dirty.hidden = !dirty;
}

function setDirty(d) {
  const next = !!d && !!doc;
  if (next === dirty) return;
  dirty = next;
  updateTitle();
}

/** Info bar. opts: { type:"info"|"warn"|"error", text, items:[string], actions:[{label, onClick, primary}] } */
function showInfo({ type = "info", text = "", items = [], actions = [] } = {}) {
  els.infoBar.className = `info-bar ${type}`;
  els.infoIcon.textContent = type === "info" ? "i" : "!";
  els.infoText.textContent = text;
  els.infoList.innerHTML = "";
  const MAX = 6;
  for (const it of items.slice(0, MAX)) {
    const li = document.createElement("li");
    li.textContent = it;
    els.infoList.appendChild(li);
  }
  if (items.length > MAX) {
    const li = document.createElement("li");
    li.textContent = `…and ${items.length - MAX} more`;
    els.infoList.appendChild(li);
  }
  els.infoActions.innerHTML = "";
  for (const a of actions) {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "tb" + (a.primary ? " primary" : "");
    b.textContent = a.label;
    b.addEventListener("click", () => { hideInfo(); a.onClick(); });
    els.infoActions.appendChild(b);
  }
  els.infoBar.hidden = false;
}

function hideInfo() { els.infoBar.hidden = true; }

function openTab(url) {
  if (typeof chrome !== "undefined" && chrome.tabs?.create) chrome.tabs.create({ url, active: true });
  else window.open(url, "_blank", "noopener");
}

function viewerUrl(query) {
  const path = `viewer/viewer.html${query ? `?${query}` : ""}`;
  if (typeof chrome !== "undefined" && chrome.runtime?.getURL) return chrome.runtime.getURL(path);
  return `viewer.html${query ? `?${query}` : ""}`;
}

/** Keep the address bar in sync so a reload reopens the same document. */
function setUrlParams(params) {
  const q = new URLSearchParams(params).toString();
  history.replaceState(null, "", location.pathname + (q ? `?${q}` : ""));
}

// ---------------------------------------------------------------------------
// modal infrastructure: focus trap, Escape, backdrop click, focus restore
// ---------------------------------------------------------------------------

const modalStack = [];

function openModal(modal, { initialFocus, onCancel } = {}) {
  closeMenus();
  const prev = document.activeElement;
  const entry = { modal, onCancel, prev };
  modalStack.push(entry);
  modal.hidden = false;
  const focusables = () => [...modal.querySelectorAll("button, input, select, textarea, [tabindex]:not([tabindex='-1'])")]
    .filter((n) => !n.disabled && !n.hidden && n.offsetParent !== null);
  entry.keydown = (e) => {
    if (e.key === "Escape") {
      e.preventDefault();
      e.stopPropagation();
      if (entry.onCancel) entry.onCancel();
      else closeModal(modal);
    } else if (e.key === "Tab") {
      const f = focusables();
      if (!f.length) return;
      const first = f[0], last = f[f.length - 1];
      if (e.shiftKey && document.activeElement === first) { e.preventDefault(); last.focus(); }
      else if (!e.shiftKey && document.activeElement === last) { e.preventDefault(); first.focus(); }
    }
    // Global shortcuts must not act on the document behind a dialog.
    if ((e.ctrlKey || e.metaKey) && /^[sop]$/i.test(e.key)) e.preventDefault();
    e.stopPropagation();
  };
  entry.mousedown = (e) => {
    if (e.target === modal) {
      e.preventDefault();
      if (entry.onCancel) entry.onCancel();
      else closeModal(modal);
    }
  };
  modal.addEventListener("keydown", entry.keydown);
  modal.addEventListener("mousedown", entry.mousedown);
  const target = initialFocus || focusables()[0];
  if (target) setTimeout(() => { target.focus(); if (target.select && target.tagName === "INPUT") target.select(); }, 0);
}

function closeModal(modal) {
  const i = modalStack.findIndex((m) => m.modal === modal);
  if (i < 0) { modal.hidden = true; return; }
  const [entry] = modalStack.splice(i, 1);
  modal.removeEventListener("keydown", entry.keydown);
  modal.removeEventListener("mousedown", entry.mousedown);
  modal.hidden = true;
  if (entry.prev && entry.prev.isConnected && typeof entry.prev.focus === "function") {
    try { entry.prev.focus({ preventScroll: true }); } catch { /* ignore */ }
  }
}

const anyModalOpen = () => modalStack.length > 0;

/**
 * Generic dialog. buttons: [{id, label, primary?, danger?}] (right to left
 * order is the given order). Resolves with the clicked id, or cancelId on
 * Escape / backdrop click.
 */
function askDialog({ title, message, buttons, cancelId = "cancel", focusId }) {
  return new Promise((resolve) => {
    els.dlgTitle.textContent = title;
    els.dlgText.textContent = message || "";
    els.dlgActions.innerHTML = "";
    let focusBtn = null;
    const done = (id) => { closeModal(els.dialogModal); resolve(id); };
    for (const b of buttons) {
      if (b.spacer) {
        const s = document.createElement("span");
        s.className = "spacer";
        els.dlgActions.appendChild(s);
        continue;
      }
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "tb" + (b.primary ? " primary" : "") + (b.danger ? " danger-text" : "");
      btn.textContent = b.label;
      btn.addEventListener("click", () => done(b.id));
      els.dlgActions.appendChild(btn);
      if (b.id === focusId || (!focusId && b.primary)) focusBtn = btn;
    }
    openModal(els.dialogModal, { initialFocus: focusBtn, onCancel: () => done(cancelId) });
  });
}

// ---------------------------------------------------------------------------
// menus / popovers
// ---------------------------------------------------------------------------

function toggleMenu(btn, menu, force) {
  const open = force ?? menu.hidden;
  closeMenus(menu);
  menu.hidden = !open;
  btn.setAttribute("aria-expanded", String(open));
  if (open) {
    const first = menu.querySelector("button:not(:disabled)");
    if (first) setTimeout(() => first.focus(), 0);
  }
}

/** Word-bar menus: [opener, menu]. Their commands act on the editor selection. */
const WORD_MENUS = [[els.fbInsert, els.menuInsert], [els.fbFormat, els.menuFormat], [els.fbLayout, els.menuLayout], [els.fbView, els.menuView]];

function closeMenus(except) {
  for (const [btn, menu] of [[els.btnNew, els.menuNew], [els.btnMore, els.menuMore], [els.fbTable, els.tablePicker], [els.fbTableMenu, els.menuTable], ...WORD_MENUS]) {
    if (menu === except) continue;
    if (!menu.hidden) {
      menu.hidden = true;
      btn.setAttribute("aria-expanded", "false");
    }
  }
}

function bindMenus() {
  els.btnNew.addEventListener("click", (e) => { e.stopPropagation(); toggleMenu(els.btnNew, els.menuNew); });
  els.btnMore.addEventListener("click", (e) => { e.stopPropagation(); toggleMenu(els.btnMore, els.menuMore); });
  els.menuNew.addEventListener("click", (e) => {
    const b = e.target.closest("button[data-new]");
    if (!b) return;
    closeMenus();
    newDocument(b.dataset.new);
  });
  els.miOpenUrl.addEventListener("click", () => { closeMenus(); promptOpenUrl(); });
  els.miDownloadOriginal.addEventListener("click", () => { closeMenus(); downloadOriginal(); });
  els.miShortcuts.addEventListener("click", () => { closeMenus(); showShortcuts(); });
  els.miNewTemplate.addEventListener("click", () => { closeMenus(); openTemplates(); });
  els.miPageSetup.addEventListener("click", () => { closeMenus(); openPageSetup(); });
  els.miProperties.addEventListener("click", () => { closeMenus(); openProperties(); });
  els.miSettings.addEventListener("click", () => {
    closeMenus();
    if (typeof chrome !== "undefined" && chrome.runtime?.openOptionsPage) chrome.runtime.openOptionsPage();
    else openTab("../options/options.html");
  });
  document.addEventListener("mousedown", (e) => {
    if (!e.target.closest(".menu-wrap")) closeMenus();
  });
  const menuButton = new Map([[els.menuNew, els.btnNew], [els.menuMore, els.btnMore], [els.menuTable, els.fbTableMenu],
    ...WORD_MENUS.map(([b, m]) => [m, b])]);
  const wordMenus = new Set([els.menuTable, ...WORD_MENUS.map(([, m]) => m)]);
  for (const menu of menuButton.keys()) {
    menu.addEventListener("keydown", (e) => {
      const items = [...menu.querySelectorAll("button:not(:disabled)")];
      const i = items.indexOf(document.activeElement);
      if (e.key === "ArrowDown") { e.preventDefault(); items[(i + 1) % items.length]?.focus(); }
      else if (e.key === "ArrowUp") { e.preventDefault(); items[(i - 1 + items.length) % items.length]?.focus(); }
      else if (e.key === "Escape") {
        e.preventDefault();
        e.stopPropagation();
        closeMenus();
        if (wordMenus.has(menu)) restoreWordFocus();
        else menuButton.get(menu).focus();
      }
    });
  }
}

function showShortcuts() {
  openModal(els.shortcutsModal, { initialFocus: els.ksClose });
}

// ---------------------------------------------------------------------------
// backend status (health re-checked periodically, on focus, on settings change)
// ---------------------------------------------------------------------------

function backendEnabled() {
  return !!(backend.settings && backend.settings.backend_enabled);
}

function backendUp() {
  return backendEnabled() && !!backend.client && !!backend.health.ok;
}

/** Can the backend parse `format` right now? */
function backendCanParse(format) {
  if (!backendUp()) return false;
  const list = backendParseFormats(backend.health);
  return list == null || list.includes(format);
}

async function checkBackend({ announce = false } = {}) {
  if (backend.checking) return backend.checking;
  backend.checking = (async () => {
    if (!backend.settings) backend.settings = await loadSettings();
    const wasOk = backendUp();
    if (!backendEnabled()) {
      backend.client = null;
      backend.health = { ok: false, disabled: true };
    } else {
      const url = backend.settings.backend_url;
      if (!backend.client || backend.client.base !== createBackendClient(url).base) backend.client = createBackendClient(url);
      if (!backend.health.ok) renderBackendStatus(true);
      backend.health = await backend.client.health();
    }
    renderBackendStatus(false);
    const isOk = backendUp();
    if (announce) {
      toast(isOk ? "Backend online" : backendEnabled() ? "Backend offline" : "Backend is disabled in Settings", isOk ? "ok" : "");
    }
    if (!wasOk && isOk && doc && doc.previewOnly && doc.bytes) {
      showInfo({
        type: "info",
        text: "The office backend is now available, so this document can be opened for editing.",
        actions: [{ label: "Reopen for editing", primary: true, onClick: () => reopenCurrent() }],
      });
    }
  })();
  try { await backend.checking; } finally { backend.checking = null; }
}

function renderBackendStatus(checking) {
  const dot = els.backendDot;
  let label, tip, cls;
  if (checking) {
    cls = "checking"; label = "Backend…"; tip = "Checking the office backend…";
  } else if (!backendEnabled()) {
    cls = "off"; label = "Offline mode";
    tip = "The office backend is disabled in Settings. Word .docx, Markdown, Excel and text files still open and save; " +
      ".doc, .rtf and .odt need the backend.";
  } else if (backend.health.ok) {
    const h = backend.health;
    cls = "ok"; label = "Backend";
    tip = `Backend online: ${h.name || "office-backend"} ${h.version || ""} at ${backend.client.base}.`;
    if (!isLoopbackBackendUrl(backend.client.base)) {
      cls = "warn"; label = "Remote backend";
      tip += "\nWarning: this backend is not on this computer, so opened documents are uploaded to it.";
    }
  } else {
    cls = "fail"; label = "Backend offline";
    tip = `Backend offline (${backend.health.error || "unreachable"}) at ${backend.client ? backend.client.base : "?"}.\n` +
      "Word .docx, Markdown, Excel and text files still work; .doc, .rtf and .odt need office-backend running.";
  }
  dot.className = `backend-dot ${cls}`;
  els.backendLabel.textContent = label;
  els.backendStatus.title = `${tip}\nClick to check again.`;
  els.backendStatus.setAttribute("aria-label", `${label}. ${tip}`);
  updateBackendNote();
  // The LaTeX view enables "Compile PDF" only while a TeX-capable backend is up.
  if (latexView && typeof latexView.refreshCompile === "function") {
    try { latexView.refreshCompile(); } catch (e) { console.warn("refreshCompile failed:", e); }
  }
}

function updateBackendNote() {
  if (!els.backendNote) return;
  if (backendUp()) {
    els.backendNote.textContent = "Office backend connected: every supported format opens and saves with full fidelity.";
  } else {
    els.backendNote.textContent =
      "Works offline with .docx, .md, .xlsx, .xls, .ods, .csv and .txt. " +
      "Opening .doc, .rtf and .odt needs the office-backend helper" +
      (backendEnabled() ? " (currently not running)." : " (disabled in Settings).");
  }
}

function startHealthPolling() {
  clearInterval(backend.timer);
  backend.timer = setInterval(() => {
    if (document.visibilityState === "visible") checkBackend();
  }, HEALTH_INTERVAL_MS);
  window.addEventListener("focus", () => checkBackend());
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "visible") checkBackend();
  });
  if (typeof chrome !== "undefined" && chrome.storage?.onChanged) {
    chrome.storage.onChanged.addListener((changes, area) => {
      if (area !== "local" || !changes.settings) return;
      loadSettings().then((s) => {
        backend.settings = s;
        backend.client = null;
        backend.health = { ok: false };
        checkBackend();
      });
    });
  }
  els.backendStatus.addEventListener("click", () => checkBackend({ announce: true }));
}

/** A backend call failed at the network level: refresh the status now. */
function noteBackendFailure(e) {
  if (e && e.network) {
    backend.health = { ok: false, error: "unreachable" };
    renderBackendStatus(false);
    checkBackend();
  }
}

// ---------------------------------------------------------------------------
// open pipeline
// ---------------------------------------------------------------------------

/**
 * Ask what to do with unsaved changes before replacing the document.
 * opts.offerNewTab: also offer opening the incoming file in a new tab.
 * Resolves true (go ahead), false (cancel) or "newtab".
 */
async function confirmDiscard(opts = {}) {
  if (!doc || !dirty) return true;
  const buttons = [];
  if (opts.offerNewTab) buttons.push({ id: "newtab", label: "Open in new tab" }, { spacer: true });
  buttons.push(
    { id: "cancel", label: "Cancel" },
    { id: "discard", label: "Don't save", danger: true },
    { id: "save", label: "Save", primary: true },
  );
  const choice = await askDialog({
    title: "Save changes?",
    message: `“${doc.name}” has unsaved changes. Save them before ${opts.action || "continuing"}?`,
    buttons,
  });
  if (choice === "save") return await saveDefault();
  if (choice === "discard") return true;
  if (choice === "newtab") return "newtab";
  return false;
}

/** Store bytes as a handoff and open them in another viewer tab. */
async function openInNewTab(name, bytes, mime = "") {
  try {
    const id = await putHandoff(name, bytes, mime);
    if (typeof chrome !== "undefined" && chrome.tabs?.create) {
      await chrome.tabs.create({ url: viewerUrl(`handoff=${encodeURIComponent(id)}`), active: false });
    } else {
      window.open(viewerUrl(`handoff=${encodeURIComponent(id)}`), "_blank");
    }
    return true;
  } catch (e) {
    toast(`Could not open ${name} in a new tab: ${errMsg(e)}`, "error");
    return false;
  }
}

/**
 * Open a list of sources ({file?, handle?, name?, bytes?, mime?}): the first
 * replaces the current document (after the unsaved-changes guard), the
 * others open in new tabs.
 */
async function openSources(sources, { action = "opening another file" } = {}) {
  const list = sources.filter(Boolean);
  if (!list.length) return false;
  const [first, ...rest] = list;
  const materialize = async (s) => {
    if (s.bytes) return s;
    const file = s.file || (s.handle ? await s.handle.getFile() : null);
    if (!file) throw new Error("No file");
    return { name: file.name, bytes: new Uint8Array(await file.arrayBuffer()), mime: file.type || "", handle: s.handle || null };
  };
  for (const s of rest) {
    try {
      const m = await materialize(s);
      await openInNewTab(m.name, m.bytes, m.mime);
    } catch (e) {
      toast(`Could not open a file: ${errMsg(e)}`, "error");
    }
  }
  let src;
  try {
    src = await materialize(first);
  } catch (e) {
    toast(`Could not read the file: ${errMsg(e)}`, "error");
    return false;
  }
  return openSource(src, { action });
}

/**
 * Open one source {name, bytes, mime, handle?, origin?}. Parses first; the
 * current document is only replaced when the parse succeeded.
 */
async function openSource(src, { skipGuard = false, action = "opening another file" } = {}) {
  if (!skipGuard) {
    const ok = await confirmDiscard({ offerNewTab: !!src.bytes, action });
    if (ok === "newtab") return openInNewTab(src.name, src.bytes, src.mime);
    if (!ok) return false;
  }
  setBusy(true, `Opening ${src.name}…`);
  setStatus(`Opening ${src.name}…`);
  try {
    const loaded = await loadDocument(src);
    await commitDocument(loaded);
    return true;
  } catch (e) {
    console.error(e);
    const msg = errMsg(e);
    // "Could not open x.docx: it is damaged…" or, for a full sentence from
    // the reader, "Could not open x.docx. This is not a valid…".
    const line = `Could not open ${src.name}${/^[A-Z]/.test(msg) ? ". " : ": "}${msg}`;
    setStatus(line, true);
    if (e && e.htmlPage && src.url) {
      showInfo({
        type: "error",
        text: `“${src.url}” returned a web page instead of a document — the site may require you to sign in first.`,
        actions: [{ label: "Open link in a normal tab", onClick: () => openTab(src.url) }],
      });
    } else {
      toast(line, "error");
    }
    if (!doc) showEmpty();
    return false;
  } finally {
    setBusy(false);
  }
}

/**
 * Parse a source into a detached document object (no global state touched).
 */
async function loadDocument(src) {
  const bytes = src.bytes instanceof Uint8Array ? src.bytes : new Uint8Array(src.bytes || []);
  const format = resolveFormat(src.name, src.mime, bytes);
  if (!format) {
    const ext = (/\.([^.]+)$/.exec(src.name || "") || [])[1];
    throw new Error(ext
      ? `.${ext} files are not supported. Supported: Word (.docx .doc .rtf .odt), Markdown, text, Excel (.xlsx .xls .xlsm .ods), CSV/TSV.`
      : "This file type was not recognized.");
  }
  let name = src.name || "document";
  if (formatOfName(name) !== format) {
    // Extensionless or mislabeled (e.g. RTF saved as .doc): name the doc by its real format.
    name = formatOfName(name) ? withExtension(name, format) : `${name}.${format}`;
  }
  const kind = editorKindOf(format);
  const base = { kind, format, name, mime: src.mime || "", bytes, handle: src.handle || null, warnings: [], engine: "" };
  if (kind === "text") {
    const dec = decodeText(bytes);
    const warnings = dec.encoding === "utf-8" ? [] : [`Read as ${encodingLabel(dec.encoding)}; it will be saved the same way.`];
    return { ...base, text: dec.text, encoding: dec.encoding, bom: dec.bom, eol: dec.eol, warnings, engine: "text" };
  }
  if (kind === "excel") return loadExcel(base, bytes);
  if (kind === "slides") return loadSlides(base, bytes);
  if (kind === "latex") {
    const dec = decodeText(bytes);
    const warnings = dec.encoding === "utf-8" ? [] : [`Read as ${encodingLabel(dec.encoding)}; it will be saved the same way.`];
    return { ...base, text: dec.text, encoding: dec.encoding, bom: dec.bom, eol: dec.eol, warnings, engine: "text" };
  }
  return loadWord(base, bytes);
}

async function loadSlides(d, bytes) {
  const errors = [];
  if (backendCanParse(d.format)) {
    try {
      const res = await backend.client.parse(bytes, d.name);
      if (res.model && Array.isArray(res.model.slides)) {
        return finishSlides({ ...d, engine: "backend" }, res.model, collectWarnings(res.warnings));
      }
      errors.push("backend returned no presentation");
    } catch (e) {
      noteBackendFailure(e);
      errors.push(`backend: ${errMsg(e)}`);
    }
  }
  if (typeof LocalPptx.pptxToModel === "function" && typeof JSZip !== "undefined") {
    try {
      const res = await LocalPptx.pptxToModel(bytes);
      return finishSlides({ ...d, engine: "local" }, res.model, collectWarnings(res.warnings));
    } catch (e) {
      errors.push(`built-in reader: ${errMsg(e)}`);
    }
  }
  if (errors.length) console.warn(`readers failed for ${d.name}:`, errors);
  const why = errors.length ? friendlyReason(errors[0]) : "no presentation reader available";
  if (/^(This|The) /.test(why) && /damaged|not a valid|not an? /i.test(why)) throw new Error(why);
  throw new Error(`the presentation could not be read (${why})`);
}

function finishSlides(d, model, warnings) {
  if (!model || typeof model !== "object") throw new Error("Empty presentation");
  const deck = typeof SlidesModel.normalizeDeck === "function" ? SlidesModel.normalizeDeck(model) : model;
  deck.kind = "slides";
  return { ...d, model: deck, warnings: collectWarnings(warnings, deck.meta && deck.meta.warnings) };
}

function encodingLabel(enc) {
  return { "utf-8": "UTF-8", "utf-16le": "UTF-16", "utf-16be": "UTF-16 (big-endian)", "windows-1252": "Windows-1252 (ANSI)" }[enc] || enc;
}

/**
 * Every import note for a parsed model: response warnings, model.warnings,
 * meta.warnings, and the lossy flag/note (e.g. a .doc read as plain text).
 */
function modelWarnings(model, extra) {
  const meta = (model && model.meta) || {};
  const lossy = meta.lossy ? [meta.note || "This file was imported with reduced fidelity: some formatting or content may be missing."] : [];
  return collectWarnings(lossy, extra, model && model.warnings, meta.warnings);
}

function collectWarnings(...lists) {
  const out = [];
  for (const l of lists) for (const w of l || []) if (w && !out.includes(String(w))) out.push(String(w));
  return out;
}

async function loadExcel(d, bytes) {
  const errors = [];
  if (d.format === "csv" || d.format === "tsv") {
    let model;
    if (typeof ExcelModel.csvBytesToModel === "function") {
      model = ExcelModel.csvBytesToModel(bytes, d.name, d.format === "tsv" ? "\t" : undefined);
    } else {
      const dec = decodeText(bytes);
      const delim = d.format === "tsv" ? "\t" : sniffDelimiter(dec.text);
      model = ExcelModel.csvToModel(dec.text, d.name.replace(/\.[^.]+$/, ""), delim);
    }
    return finishExcel({ ...d, engine: "csv" }, model, model.warnings);
  }
  if (backendCanParse(d.format)) {
    try {
      const res = await backend.client.parse(bytes, d.name);
      if (res.model && res.model.sheets) {
        return finishExcel({ ...d, engine: "backend" }, res.model, collectWarnings(res.warnings, res.model.warnings));
      }
      errors.push("backend returned no workbook");
    } catch (e) {
      noteBackendFailure(e);
      errors.push(`backend: ${errMsg(e)}`);
    }
  }
  if (LocalXlsx.xlsxAvailable()) {
    try {
      const model = LocalXlsx.bytesToModel(bytes, d.name);
      return finishExcel({ ...d, engine: "sheetjs" }, model, model.warnings);
    } catch (e) {
      errors.push(`built-in reader: ${errMsg(e)}`);
    }
  }
  if (errors.length) console.warn(`readers failed for ${d.name}:`, errors);
  const why = errors.length ? friendlyReason(errors[0]) : "no spreadsheet reader available";
  if (/^(This|The) /.test(why) && /damaged|not a valid|not an? /i.test(why)) throw new Error(why);
  throw new Error(`the workbook could not be read (${why})`);
}

function finishExcel(d, model, warnings) {
  if (!model || typeof model !== "object") throw new Error("Empty workbook");
  if (d.format === "xlsm") warnings = collectWarnings(warnings, ["Macros (VBA) are not kept: saving creates a macro-free .xlsx workbook."]);
  model.kind = "excel";
  if (!Array.isArray(model.sheets) || !model.sheets.length) model.sheets = [{ name: "Sheet1", cells: {} }];
  for (const s of model.sheets) if (!s.cells) s.cells = {};
  if (!(model.active >= 0 && model.active < model.sheets.length)) model.active = 0;
  return { ...d, model, warnings: collectWarnings(warnings) };
}

async function loadWord(d, bytes) {
  if (d.format === "md") {
    const dec = decodeText(bytes);
    const model = markdownToModel(dec.text);
    return { ...d, model, encoding: dec.encoding, bom: dec.bom, eol: dec.eol, engine: "markdown" };
  }
  const errors = [];
  const tryBackend = async () => {
    if (!backendCanParse(d.format)) return null;
    try {
      const res = await backend.client.parse(bytes, d.name);
      if (!res.model || !Array.isArray(res.model.blocks)) throw new Error("backend returned no document");
      return { model: res.model, warnings: modelWarnings(res.model, res.warnings), engine: "backend" };
    } catch (e) {
      noteBackendFailure(e);
      errors.push(`backend: ${errMsg(e)}`);
      return null;
    }
  };
  const tryLocal = async () => {
    if (d.format !== "docx" || typeof LocalDocx.docxToModel !== "function" || typeof JSZip === "undefined") return null;
    try {
      const res = await LocalDocx.docxToModel(bytes);
      return { model: res.model, warnings: modelWarnings(res.model, res.warnings), engine: "local" };
    } catch (e) {
      errors.push(`built-in reader: ${errMsg(e)}`);
      return null;
    }
  };
  // The backend is the reference reader; the in-page reader covers offline use
  // and backend failures.
  const got = (await tryBackend()) || (await tryLocal());
  if (got) {
    const model = got.model;
    model.kind = "word";
    if (!Array.isArray(model.blocks)) model.blocks = [];
    return { ...d, model, warnings: got.warnings, engine: got.engine };
  }
  if (d.format === "docx" && LocalDocx.docxPreviewAvailable()) {
    // Last resort: a read-only rendering. Render it now, off-screen, so a
    // damaged file fails here instead of after the current document is gone.
    const container = document.createElement("div");
    try {
      await LocalDocx.renderDocxPreview(bytes, container);
      if (!container.textContent.trim() && !container.querySelector("img, section")) throw new Error("nothing to show");
      return {
        ...d,
        previewOnly: true,
        previewNode: container,
        engine: "preview",
        warnings: collectWarnings([`This document could not be converted for editing (${errors.join("; ")}). It is shown read-only.`]),
      };
    } catch (e) {
      errors.push(`preview: ${errMsg(e)}`);
    }
  }
  if (["doc", "rtf", "odt"].includes(d.format) && !backendUp()) {
    throw new Error(backendEnabled()
      ? `Opening .${d.format} files needs the office-backend helper, which is not running. Start office-backend and try again, or save the file as .docx in your word processor.`
      : `Opening .${d.format} files needs the office-backend helper, which is turned off. Enable it in the extension's Settings, or save the file as .docx in your word processor.`);
  }
  // Every reader's reason goes to the console; the user sees the first one
  // (the backend's, which is the most specific) without the reader names.
  if (errors.length) console.warn(`readers failed for ${d.name}:`, errors);
  const why = errors.length ? friendlyReason(errors[0]) : "";
  // Readers that already explain the problem in a full sentence (the backend
  // does) are shown as they are; terse library errors get a plain summary.
  if (/^(This|The) /.test(why) && /damaged|not a valid|not an? /i.test(why)) throw new Error(why);
  throw new Error(`it is damaged or is not a real .${d.format} file${why ? ` (${why})` : ""}`);
}

/**
 * A reader's error as a short reason for people: drops the "reader:" prefix
 * and replaces low-level ZIP/XML library messages (JSZip's "Can't find end of
 * central directory … see https://…") with plain words.
 */
function friendlyReason(err) {
  const s = String(err || "").replace(/^[a-z -]+:\s*/i, "").replace(/\.$/, "").trim();
  if (/central directory|is this a zip|corrupted zip|end of data reached|invalid (?:zip|signature)/i.test(s)) return "it is not a readable ZIP package";
  if (/<parsererror|XML Parsing Error|not well-formed/i.test(s)) return "its XML content is damaged";
  return s.length > 200 ? `${s.slice(0, 200)}…` : s;
}

/** Swap a loaded document in: tear down the old views, render the new one. */
async function commitDocument(loaded, { isNew = false } = {}) {
  teardownViews();
  hideInfo();
  // An error about a previous attempt is stale once a document opens.
  if (els.toast.classList.contains("error")) { clearTimeout(toastTimer); els.toast.hidden = true; }
  doc = {
    ...loaded,
    saveTarget: null,
    recentId: null,
    isNew,
  };
  pageView = false;
  dirty = false;
  if (doc.kind === "word") {
    if (doc.previewOnly) await renderPreviewOnly();
    else renderWordEditor();
  } else if (doc.kind === "excel") {
    renderExcel();
  } else if (doc.kind === "slides") {
    renderSlides();
  } else if (doc.kind === "latex") {
    renderLatex();
  } else {
    renderText();
  }
  els.kind.textContent = KIND_LABEL[doc.kind] + (doc.format ? ` · ${doc.format.toUpperCase()}` : "");
  els.kind.className = `badge ${doc.kind}`;
  updateTitle();
  updateButtons();
  updateStats();
  const engineNote = {
    sheetjs: "Opened with the built-in spreadsheet reader (formats, styles, merges, column widths, freeze panes, filters and charts are kept; images are not).",
    local: doc.kind === "slides" ? "Opened with the built-in PowerPoint reader." : "Opened with the built-in Word reader.",
    preview: "Read-only preview.",
  }[doc.engine];
  setStatus(isNew ? "New document" : engineNote || `Opened ${doc.name}`);
  if (doc.warnings && doc.warnings.length) {
    showInfo({
      type: doc.previewOnly ? "warn" : "info",
      text: doc.previewOnly
        ? "Read-only: this document could not be converted for editing."
        : `Opened with ${doc.warnings.length === 1 ? "a note" : `${doc.warnings.length} notes`} — some content may look different:`,
      items: doc.previewOnly ? doc.warnings.concat(["You can still print it, export it, or download the original file."]) : doc.warnings,
      actions: doc.previewOnly && doc.bytes ? [{ label: "Download original", onClick: () => downloadOriginal() }] : [],
    });
  }
  if (!isNew && doc.bytes && doc.bytes.length && doc.bytes.length <= RECENT_MAX_BYTES) {
    rememberRecent(doc.name, doc.bytes, doc.mime || mimeOf(doc.format));
  } else if (isNew) {
    setUrlParams({ new: doc.kind });
  } else {
    setUrlParams({});
  }
}

function rememberRecent(name, bytes, mime, { saved = false } = {}) {
  const current = doc;
  // Both replace an existing entry with the same name; updateRecent is the
  // "these are the saved bytes of that document" spelling.
  (saved ? updateRecent : addRecent)(name, bytes, mime).then((id) => {
    if (doc !== current || !id) return;
    doc.recentId = id;
    setUrlParams({ recent: id }); // a reload reopens this document
  }).catch((e) => console.warn("recent store failed:", e));
}

function teardownViews() {
  destroyExcelView();
  destroyAppView("slides");
  destroyAppView("latex");
  resetWordFeatures();
  els.wordView.contentEditable = "false";
  els.wordView.innerHTML = "";
  els.wordPreview.innerHTML = "";
  els.textView.value = "";
  els.printText.textContent = "";
  els.stage.classList.remove("excel-mode", "text-mode", "app-mode");
}

function showOnly(which) {
  els.wordView.hidden = which !== "word";
  els.wordSheet.hidden = which !== "word";
  if (which !== "word") { closeFind({ restore: false }); deactivateBand(); }
  els.wordPreview.hidden = which !== "preview";
  els.excelView.hidden = which !== "excel";
  els.excelBar.hidden = which !== "excel";
  els.textView.hidden = which !== "text";
  els.slidesView.hidden = which !== "slides";
  els.latexView.hidden = which !== "latex";
  els.wordBar.hidden = which !== "word";
  els.empty.hidden = which !== "empty";
  els.stage.classList.toggle("excel-mode", which === "excel");
  els.stage.classList.toggle("text-mode", which === "text");
  // Slides and LaTeX views lay themselves out inside a full-bleed stage.
  els.stage.classList.toggle("app-mode", which === "slides" || which === "latex");
  updateNavVisibility();
  els.zoomCtl.hidden = which !== "word";
}

function showEmpty() {
  showOnly("empty");
  updateButtons();
  updateBackendNote();
}

function updateButtons() {
  const has = !!doc;
  els.btnSave.disabled = !has;
  els.btnSaveAs.disabled = !has;
  els.btnPreview.hidden = !(has && doc.kind === "word" && !doc.previewOnly);
  els.btnPreview.setAttribute("aria-pressed", String(pageView));
  els.btnPreview.textContent = pageView ? "Back to editing" : "Page view";
  els.miDownloadOriginal.disabled = !(has && doc.bytes && doc.bytes.length && !doc.isNew);
  els.btnSave.title = has && doc.previewOnly ? "This document is read-only: export it or download the original" : "Save (Ctrl+S)";
  const editableWord = wordEditable();
  els.miPageSetup.disabled = !editableWord;
  els.miProperties.disabled = !editableWord;
  els.stats.disabled = !editableWord;
  els.stats.title = editableWord ? "Word count (Ctrl+Shift+G)" : "";
}

/** A Word document is open in the editor (not read-only, not the page view). */
function wordEditable() {
  return !!doc && doc.kind === "word" && !doc.previewOnly && !pageView;
}

/** Reopen the current document from its original bytes (e.g. once the backend is up). */
async function reopenCurrent() {
  if (!doc || !doc.bytes) return;
  const { name, bytes, mime, handle } = doc;
  await openSource({ name, bytes, mime, handle }, { action: "reopening the document" });
}

async function newDocument(kind, opts = {}) {
  const ok = await confirmDiscard({ action: "creating a new document" });
  if (ok !== true) return false;
  if (kind === "slides") {
    await commitDocument({
      kind: "slides", format: "pptx", name: "Untitled.pptx", mime: "", bytes: null, handle: null, warnings: [], engine: "new",
      model: SlidesModel.newDeck(opts.template),
    }, { isNew: true });
    return true;
  }
  if (kind === "latex") {
    await commitDocument({
      kind: "latex", format: "tex", name: "Untitled.tex", mime: "", bytes: null, handle: null, warnings: [], engine: "new",
      text: Latex.newLatexDocument(opts.template), encoding: "utf-8", bom: false, eol: "\n",
    }, { isNew: true });
    return true;
  }
  if (kind === "excel") {
    await commitDocument({
      kind: "excel", format: "xlsx", name: "Untitled.xlsx", mime: "", bytes: null, handle: null, warnings: [], engine: "new",
      model: { kind: "excel", active: 0, sheets: [{ name: "Sheet1", cells: {} }] },
    }, { isNew: true });
  } else {
    // Templates (lib/templates-word.js); unknown ids fall back to a blank page.
    const tpl = opts.template ? wordTemplate(String(opts.template)) : null;
    await commitDocument({
      kind: "word", format: "docx", name: "Untitled.docx", mime: "", bytes: null, handle: null, warnings: [], engine: "new",
      model: tpl || { kind: "word", blocks: [{ type: "p", runs: [] }] },
    }, { isNew: true });
    if (tpl && opts.template !== "blank") {
      setUrlParams({ new: "word", template: opts.template });
      const info = WORD_TEMPLATES.find((t) => t.id === opts.template);
      setStatus(`New document from the “${info ? info.name : opts.template}” template — replace the [bracketed] text.`);
    }
  }
  if (doc.kind === "word") placeCaretAtStart();
  return true;
}

// ---------------------------------------------------------------------------
// text documents
// ---------------------------------------------------------------------------

function renderText() {
  els.textView.value = doc.text;
  showOnly("text");
  els.textView.focus();
  els.textView.setSelectionRange(0, 0);
  els.textView.scrollTop = 0;
}

// ---------------------------------------------------------------------------
// slides and LaTeX documents (self-contained views)
// ---------------------------------------------------------------------------

function viewConfirm() {
  return async (message, { title = "Are you sure?", ok = "OK" } = {}) => {
    const id = await askDialog({
      title, message,
      buttons: [{ id: "cancel", label: "Cancel" }, { id: "ok", label: ok, primary: true }],
      focusId: "cancel",
    });
    return id === "ok";
  };
}

function destroyAppView(which) {
  const view = which === "slides" ? slidesView : latexView;
  if (view) {
    try { view.destroy(); } catch (e) { console.warn(`${which} view destroy failed:`, e); }
  }
  if (which === "slides") slidesView = null; else latexView = null;
  const root = which === "slides" ? els.slidesView : els.latexView;
  if (root) root.innerHTML = "";
}

function renderSlides() {
  showOnly("slides");
  slidesView = new SlidesView({
    root: els.slidesView,
    getModel: () => (doc && doc.kind === "slides" ? doc.model : null),
    onStatus: (m) => setStatus(m),
    onDirty: (d) => setDirty(d !== false),
    onStats: () => updateStats(),
    confirm: viewConfirm(),
  });
  slidesView.render();
  try { slidesView.focus(); } catch { /* ignore */ }
}

function latexCanCompile() {
  return backendUp() && backendRenderFormats(backend.health, "latex").includes("pdf");
}

/** Compile LaTeX source to a PDF Blob through the backend (throws with the TeX error). */
async function compileLatex(source) {
  if (!backendUp()) throw new Error("PDF output needs the office-backend helper, which is not running.");
  if (!latexCanCompile()) throw new Error("PDF output needs a LaTeX engine (tectonic, pdflatex, xelatex or lualatex) installed where office-backend runs.");
  const res = await backend.client.render({ kind: "latex", source }, "pdf", withExtension((doc && doc.name) || "document.tex", "pdf"));
  return res.blob;
}

function renderLatex() {
  showOnly("latex");
  latexView = new LatexView({
    root: els.latexView,
    text: doc.text || "",
    onStatus: (m) => setStatus(m),
    onDirty: (d) => setDirty(d !== false),
    onStats: () => scheduleStats(),
    confirm: viewConfirm(),
    compile: compileLatex,
    canCompile: latexCanCompile,
  });
  latexView.render();
  try { latexView.focus(); } catch { /* ignore */ }
}

// ---------------------------------------------------------------------------
// excel documents
// ---------------------------------------------------------------------------

// Elements of #excel-bar that belong to viewer.html; anything else in the bar
// was created by an ExcelView and is removed with it.
const STATIC_EXCEL_BAR = new Set([...els.excelBar.children]);

function destroyExcelView() {
  if (!excelView) return;
  try { excelView.destroy(); } catch (e) { console.warn("ExcelView.destroy failed:", e); }
  excelView = null;
  // Belt and braces against listener leaks on shared elements: replace them
  // with fresh clones (ids are kept, ExcelView looks them up by id).
  for (const id of ["name-box", "formula-bar", "btn-add-sheet", "sheet-tabs"]) {
    const el = $(id);
    if (!el) continue;
    const clone = el.cloneNode(false);
    el.replaceWith(clone);
    STATIC_EXCEL_BAR.delete(el);
    STATIC_EXCEL_BAR.add(clone);
  }
  for (const child of [...els.excelBar.children]) if (!STATIC_EXCEL_BAR.has(child)) child.remove();
  els.excelView.innerHTML = "";
}

function renderExcel() {
  showOnly("excel");
  excelView = new ExcelView({
    root: els.excelView,
    getModel: () => (doc && doc.kind === "excel" ? doc.model : null),
    onStatus: (m) => setStatus(m),
    onDirty: (d) => setDirty(d !== false),
    onStats: () => updateStats(),
    confirm: async (message, { title = "Are you sure?", ok = "OK" } = {}) => {
      const id = await askDialog({
        title, message,
        buttons: [{ id: "cancel", label: "Cancel" }, { id: "ok", label: ok, primary: true }],
        focusId: "cancel",
      });
      return id === "ok";
    },
  });
  excelView.render();
  // Put the keyboard in the grid right away (public focus() when available).
  const focusGrid = excelView.focus || excelView._focusGrid;
  if (typeof focusGrid === "function") {
    try { focusGrid.call(excelView); } catch { /* ignore */ }
  }
}

/** Flush an in-progress cell edit into the model (before save / print). */
function commitExcelEdit() {
  if (!excelView) return;
  try {
    if (typeof excelView.commitEdit === "function") excelView.commitEdit();
  } catch (e) {
    console.warn("commitEdit failed:", e);
  }
}

// ---------------------------------------------------------------------------
// word documents: editor
// ---------------------------------------------------------------------------

let savedRange = null;     // last selection inside the editor (toolbar widgets steal focus)
let pendingFontSizePt = null;

function renderWordEditor() {
  const html = renderWordHtml(doc.model);
  els.wordView.innerHTML = html && html.trim() ? html : "<p><br></p>";
  ensureEditableTail();
  els.wordView.contentEditable = "true";
  els.wordView.spellcheck = true;
  try {
    document.execCommand("styleWithCSS", false, true);
    document.execCommand("defaultParagraphSeparator", false, "p");
  } catch { /* old engines */ }
  renderHeaderFooter();
  applyPageSetup();
  applyZoom();
  showOnly("word");
  savedRange = null;
  els.stage.scrollTop = 0;
  placeCaretAtStart();
  if (navOpen) buildNav();
}

/** Make sure the caret can always get below a trailing table / image / page break. */
function ensureEditableTail() {
  const last = els.wordView.lastElementChild;
  if (!last || /^(TABLE|HR)$/.test(last.tagName)) {
    const p = document.createElement("p");
    p.appendChild(document.createElement("br"));
    els.wordView.appendChild(p);
  }
}

function placeCaretAtStart() {
  const first = els.wordView.querySelector("p, h1, h2, h3, h4, h5, h6, li, td, pre, blockquote") || els.wordView;
  const r = document.createRange();
  r.setStart(first, 0);
  r.collapse(true);
  const sel = getSelection();
  sel.removeAllRanges();
  sel.addRange(r);
  els.wordView.focus({ preventScroll: true });
}

function currentWordModel() {
  if (!doc || doc.kind !== "word") return null;
  if (doc.previewOnly) return null;
  const meta = doc.model && doc.model.meta ? { ...doc.model.meta } : undefined;
  const model = parseWordDom(els.wordView, meta ? { meta } : {});
  doc.model = model;
  return model;
}

function selectionInEditor() {
  const sel = getSelection();
  if (!sel || !sel.rangeCount) return null;
  const r = sel.getRangeAt(0);
  return els.wordView.contains(r.commonAncestorContainer) ? r : null;
}

/**
 * Put the caret back into the editor before running a command. The live
 * selection wins when it is still inside the editor (toolbar buttons do not
 * take focus); otherwise the last remembered range is restored (selects,
 * the color picker and dialogs do take focus).
 */
function restoreSelection(preferSaved = false) {
  const editorFocused = els.wordView.contains(document.activeElement);
  const live = editorFocused ? selectionInEditor() : null;
  const valid = (r) => r && els.wordView.contains(r.commonAncestorContainer);
  const range = !preferSaved && live ? live.cloneRange() : valid(savedRange) ? savedRange : live;
  els.wordView.focus({ preventScroll: true });
  if (range && els.wordView.contains(range.commonAncestorContainer)) {
    const sel = getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
  }
}

function closestInEditor(node, selector) {
  let el = node && node.nodeType === 1 ? node : node && node.parentElement;
  while (el && el !== els.wordView) {
    if (el.matches(selector)) return el;
    el = el.parentElement;
  }
  return null;
}

function anchorNode() {
  const r = selectionInEditor() || savedRange;
  return r ? r.startContainer : null;
}

// Commands that restructure blocks but should leave the caret where it was.
// (Chrome's list commands inside a <p> move the caret to the start of the line.)
const CARET_PRESERVING = new Set([
  "insertUnorderedList", "insertOrderedList", "indent", "outdent", "formatBlock",
  "justifyLeft", "justifyCenter", "justifyRight", "justifyFull",
]);


/** Run an execCommand on the editor (selection restored first). */
function exec(cmd, value = null) {
  // Editing a header/footer band: run the command there (inline formatting only).
  if (hfActive && hfActive.isConnected) return execInBand(cmd, value);
  restoreSelection();
  const sel = getSelection();
  // A collapsed caret on an empty line has no text position of its own (its
  // offset equals the end of the previous line), and Chrome already leaves it
  // in the right place there, so only remember positions on lines with text.
  const emptyLine = sel.rangeCount && sel.isCollapsed && !(closestBlock(sel.anchorNode)?.textContent ?? "x");
  const before = CARET_PRESERVING.has(cmd) && sel.rangeCount && !emptyLine
    ? {
      a: textOffsetOf(sel.anchorNode, sel.anchorOffset), f: textOffsetOf(sel.focusNode, sel.focusOffset),
      aStart: atBlockStart(sel.anchorNode, sel.anchorOffset), fStart: atBlockStart(sel.focusNode, sel.focusOffset),
      text: els.wordView.textContent,
    }
    : null;
  let ok = false;
  try { ok = document.execCommand(cmd, false, value); } catch (e) { console.warn(cmd, e); }
  if (cmd === "insertUnorderedList" || cmd === "insertOrderedList" || cmd === "indent" || cmd === "outdent") {
    unwrapListsFromParagraphs();
  }
  // Restore by character position (the command may have replaced the text
  // nodes), as long as the text itself did not change.
  if (before && before.a != null && before.f != null && els.wordView.textContent === before.text) {
    const a = pointAtTextOffset(before.a, before.aStart);
    const f = pointAtTextOffset(before.f, before.fStart);
    if (a && f) {
      try { sel.setBaseAndExtent(a.node, a.offset, f.node, f.offset); } catch { /* keep the command's caret */ }
    }
  }
  afterEdit();
  return ok;
}

/** Character offset of a DOM point from the start of the editor's text. */
function textOffsetOf(node, offset) {
  if (!node || !els.wordView.contains(node)) return null;
  const r = document.createRange();
  r.setStart(els.wordView, 0);
  try { r.setEnd(node, offset); } catch { return null; }
  return r.toString().length;
}

/** True when (node, offset) is at the very start of its block's text. */
function atBlockStart(node, offset) {
  const block = closestBlock(node);
  if (!block) return false;
  const start = textOffsetOf(block, 0);
  return start != null && start === textOffsetOf(node, offset);
}

/**
 * DOM point (in a text node when possible) at a character offset of the
 * editor's text. `lineStart`: the point was at the start of a line, so at a
 * boundary between two blocks it belongs to the later one.
 */
function pointAtTextOffset(target, lineStart = false) {
  const walker = document.createTreeWalker(els.wordView, NodeFilter.SHOW_TEXT);
  let pos = 0;
  let last = null;
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    const len = n.data.length;
    if (target <= pos + len) {
      // At a node boundary prefer the start of the next node (the caret
      // after "line" then Enter should stay on the same line).
      if (target === pos + len) {
        const next = walker.nextNode();
        if (next && (lineStart || closestBlock(next) === closestBlock(n))) return { node: next, offset: 0 };
        return { node: n, offset: len };
      }
      return { node: n, offset: target - pos };
    }
    pos += len;
    last = n;
  }
  return last ? { node: last, offset: last.data.length } : null;
}

function closestBlock(node) {
  return closestInEditor(node, "p, li, h1, h2, h3, h4, h5, h6, pre, blockquote, td, th, div");
}

/**
 * Chrome turns a paragraph into a list as <p><ul>…</ul></p> (invalid nesting
 * that also confuses later Enter/Backspace). Lift such lists out of their
 * paragraph when the paragraph holds nothing else.
 */
function unwrapListsFromParagraphs() {
  for (const list of els.wordView.querySelectorAll("p > ul, p > ol, h1 > ul, h1 > ol, h2 > ul, h2 > ol, h3 > ul, h3 > ol, h4 > ul, h4 > ol")) {
    const block = list.parentElement;
    if (!block || !block.isConnected) continue;
    const others = [...block.childNodes].filter((n) => n !== list && !(n.nodeType === 3 && !n.data.trim()) && !(n.nodeName === "BR") && !/^(UL|OL)$/.test(n.nodeName));
    if (others.length) continue;
    block.replaceWith(...[...block.childNodes].filter((n) => /^(UL|OL)$/.test(n.nodeName)));
  }
}

function afterEdit() {
  { const r = selectionInEditor(); if (r) savedRange = r.cloneRange(); }
  scheduleToolbarState();
  scheduleStats();
}

function markWordDirty() {
  setDirty(true);
  scheduleStats();
}

function applyBlockStyle(style) {
  const tag = { p: "p", h1: "h1", h2: "h2", h3: "h3", h4: "h4", quote: "blockquote", code: "pre" }[style] || "p";
  exec("formatBlock", `<${tag}>`);
  const r = selectionInEditor();
  if (r) {
    // Mark real quotes so they are not mistaken for an indent wrapper.
    const bq = closestInEditor(r.startContainer, "blockquote");
    if (tag === "blockquote" && bq) {
      bq.setAttribute("data-ow", "quote");
      bq.removeAttribute("style");
    }
  }
}

/** execCommand('fontSize') only knows 1..7: apply 7, then rewrite to the real pt size. */
function applyFontSize(pt) {
  if (!(pt > 0)) return;
  pendingFontSizePt = pt;
  exec("fontSize", "7");
  fixFontSizeMarkers();
}

function fixFontSizeMarkers() {
  if (!pendingFontSizePt) return;
  const pt = pendingFontSizePt;
  let found = false;
  for (const el of els.wordSheet.querySelectorAll('font[size="7"], span[style*="xxx-large"], span[style*="-webkit-xxx-large"]')) {
    found = true;
    el.style.fontSize = `${pt}pt`;
    if (el.tagName === "FONT") el.removeAttribute("size");
  }
  return found;
}

function applyHighlight(value) {
  if (value === "none") exec("hiliteColor", "transparent");
  else if (value) exec("hiliteColor", value);
}

function inListItem() {
  return !!closestInEditor(anchorNode(), "li");
}

function changeListLevel(deeper) {
  if (!inListItem()) {
    toast("Indent changes the level of list items. Use the list buttons to start a list.");
    return;
  }
  exec(deeper ? "indent" : "outdent");
  // Chrome may wrap a list in an indent <blockquote style="margin…">; unwrap it
  // so lists stay lists (the parser would treat it as a container anyway).
  for (const bq of els.wordView.querySelectorAll('blockquote[style*="margin"]')) {
    if (bq.children.length && [...bq.children].every((c) => /^(UL|OL)$/.test(c.tagName))) {
      bq.replaceWith(...bq.childNodes);
    }
  }
  markWordDirty();
}

// --- tables -----------------------------------------------------------------

function buildTableHtml(rows, cols) {
  let html = '<table class="ow-table"><tbody>';
  for (let r = 0; r < rows; r++) {
    html += "<tr>";
    for (let c = 0; c < cols; c++) html += `<td${r === 0 && c === 0 ? ' data-ow-new="1"' : ""}><p><br></p></td>`;
    html += "</tr>";
  }
  return html + "</tbody></table><p><br></p>";
}

function insertTable(rows, cols) {
  exec("insertHTML", buildTableHtml(rows, cols));
  // Put the caret into the first cell of the new table.
  const first = els.wordView.querySelector("td[data-ow-new]");
  if (first) {
    first.removeAttribute("data-ow-new");
    els.wordView.focus({ preventScroll: true });
    caretInto(first);
  }
  ensureEditableTail();
  markWordDirty();
  scheduleToolbarState();
}

function currentCell() {
  return closestInEditor(anchorNode(), "td, th");
}

function newCell(like) {
  const td = document.createElement(like && like.tagName === "TH" ? "th" : "td");
  td.innerHTML = "<p><br></p>";
  return td;
}

function caretInto(el) {
  const target = el.querySelector("p, h1, h2, h3, h4, h5, h6, li") || el;
  const r = document.createRange();
  r.selectNodeContents(target);
  r.collapse(true);
  const sel = getSelection();
  sel.removeAllRanges();
  sel.addRange(r);
  savedRange = r;
}

/**
 * Row/column/table operations. They are computed on a copy of the table and
 * swapped in with execCommand("insertHTML") so they join the editor's native
 * undo stack: Ctrl+Z after "Insert row" or even "Delete table" brings the old
 * table back, in order with typing and formatting.
 */
function tableOp(op) {
  const td = currentCell();
  if (!td) { toast("Put the cursor in a table cell first."); return; }
  const table = td.closest("table");
  const rIdx = td.parentElement.rowIndex;
  const cIdx = [...td.parentElement.cells].indexOf(td);
  if (op === "del-table") { replaceTable(table, null); return; }
  const copy = table.cloneNode(true);
  const rows = [...copy.rows];
  const tr = rows[rIdx];
  let target = null; // [row, cell] in the new table
  switch (op) {
    case "row-above":
    case "row-below": {
      const nr = document.createElement("tr");
      for (const c of tr.cells) {
        const cell = newCell(c);
        if (c.colSpan > 1) cell.colSpan = c.colSpan;
        nr.appendChild(cell);
      }
      tr.parentElement.insertBefore(nr, op === "row-above" ? tr : tr.nextSibling);
      target = [op === "row-above" ? rIdx : rIdx + 1, Math.min(cIdx, nr.cells.length - 1)];
      break;
    }
    case "col-left":
    case "col-right":
      for (const row of rows) {
        const ref = row.cells[Math.min(cIdx, row.cells.length - 1)];
        const cell = newCell(ref);
        if (!ref) row.appendChild(cell);
        else row.insertBefore(cell, op === "col-left" ? ref : ref.nextSibling);
      }
      target = [rIdx, op === "col-left" ? cIdx : cIdx + 1];
      break;
    case "del-row":
      if (rows.length <= 1) { replaceTable(table, null); return; }
      tr.remove();
      {
        const ri = Math.min(rIdx, rows.length - 2);
        target = [ri, Math.min(cIdx, copy.rows[ri].cells.length - 1)];
      }
      break;
    case "del-col": {
      const maxCols = Math.max(...rows.map((r) => r.cells.length));
      if (maxCols <= 1) { replaceTable(table, null); return; }
      for (const row of rows) {
        const c = row.cells[Math.min(cIdx, row.cells.length - 1)];
        if (c && row.cells.length > 1) c.remove();
      }
      target = [rIdx, Math.min(cIdx, copy.rows[rIdx].cells.length - 1)];
      break;
    }
    default:
      return;
  }
  replaceTable(table, copy, target);
}

/**
 * Replace `table` with `next` (null: remove it, leaving an empty paragraph)
 * through the undo stack, then put the caret into cell `target` [row, cell].
 */
function replaceTable(table, next, target = null) {
  const all = () => [...els.wordView.querySelectorAll("table")];
  const index = all().indexOf(table);
  const expectRows = next ? next.rows.length : -1;
  els.wordView.focus({ preventScroll: true });
  const range = document.createRange();
  range.selectNode(table);
  const sel = getSelection();
  sel.removeAllRanges();
  sel.addRange(range);
  let ok = false;
  try {
    ok = document.execCommand("insertHTML", false, next ? next.outerHTML : "<p><br></p>");
  } catch (e) {
    console.warn("table edit", e);
  }
  let placed = next ? all()[index] : null;
  if (next && (!ok || !placed || placed.rows.length !== expectRows)) {
    // The browser refused or mangled the edit: apply it directly (not undoable).
    if (table.isConnected) table.replaceWith(next);
    placed = next;
  } else if (!next && (!ok || table.isConnected)) {
    const p = document.createElement("p");
    p.appendChild(document.createElement("br"));
    if (table.isConnected) table.replaceWith(p);
  }
  if (placed && target) {
    const row = placed.rows[target[0]];
    const cell = row && row.cells[Math.max(0, Math.min(target[1], row.cells.length - 1))];
    if (cell) caretInto(cell);
  }
  ensureEditableTail();
  markWordDirty();
  afterEdit();
}

function moveToCell(td, backwards) {
  const table = td.closest("table");
  const cells = [...table.querySelectorAll("td, th")].filter((c) => c.closest("table") === table);
  const i = cells.indexOf(td);
  const target = cells[i + (backwards ? -1 : 1)];
  if (!target && !backwards) {
    // Tab in the last cell adds a row, like Word (undoable like the menu).
    const copy = table.cloneNode(true);
    const last = copy.rows[copy.rows.length - 1];
    const nr = document.createElement("tr");
    for (const c of last.cells) {
      const cell = newCell(c);
      if (c.colSpan > 1) cell.colSpan = c.colSpan;
      nr.appendChild(cell);
    }
    last.parentElement.appendChild(nr);
    replaceTable(table, copy, [copy.rows.length - 1, 0]);
    return;
  }
  if (target) {
    const r = document.createRange();
    r.selectNodeContents(target);
    const sel = getSelection();
    sel.removeAllRanges();
    sel.addRange(r);
    savedRange = r;
  }
}

// --- links --------------------------------------------------------------------

/** Normalize and validate a link address. Returns {url} or {error}. */
function normalizeLinkUrl(input) {
  const s = String(input || "").trim();
  if (!s) return { error: "Enter an address." };
  if (/^#[\w.:-]*$/.test(s)) return { url: s };
  const m = /^([a-z][a-z0-9+.-]*):/i.exec(s);
  if (m) {
    const scheme = m[1].toLowerCase();
    if (!["http", "https", "mailto", "tel", "ftp"].includes(scheme)) {
      return { error: `“${scheme}:” links are not allowed. Use http, https, mailto or tel.` };
    }
    if ((scheme === "http" || scheme === "https") && !/^https?:\/\/[^\s/]+/i.test(s)) return { error: "That web address is incomplete." };
    return { url: s };
  }
  if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s)) return { url: `mailto:${s}` };
  if (/^[^\s/]+\.[^\s/]{2,}(\/\S*)?$/.test(s)) return { url: `https://${s}` };
  return { error: "That does not look like a web address (e.g. https://example.com)." };
}

function openLinkDialog() {
  if (!doc || doc.kind !== "word" || pageView) return;
  const range = selectionInEditor() || savedRange;
  if (!range) placeCaretAtStart();
  { const r = selectionInEditor(); if (r) savedRange = r.cloneRange(); }
  const a = closestInEditor(anchorNode(), "a");
  const selectedText = savedRange ? savedRange.toString() : "";
  els.lnTitle.textContent = a ? "Edit link" : "Insert link";
  els.lnText.value = a ? a.textContent : selectedText;
  els.lnUrl.value = a ? (a.getAttribute("data-href") || a.getAttribute("href") || "") : "";
  els.lnError.hidden = true;
  els.lnRemove.hidden = !a;
  // Text only editable when inserting at a caret or editing an existing link.
  els.lnText.disabled = !a && !!selectedText && !savedRange.collapsed && /\n/.test(selectedText);
  linkTarget = a;
  openModal(els.linkModal, { initialFocus: a || selectedText ? els.lnUrl : els.lnText });
}

let linkTarget = null;

function applyLink() {
  const { url, error } = normalizeLinkUrl(els.lnUrl.value);
  if (error) {
    els.lnError.textContent = error;
    els.lnError.hidden = false;
    els.lnUrl.focus();
    return;
  }
  const text = els.lnText.value;
  closeModal(els.linkModal);
  restoreSelection(true);
  if (linkTarget && linkTarget.isConnected) {
    linkTarget.setAttribute("href", url);
    linkTarget.removeAttribute("data-href");
    linkTarget.classList.add("ow-link");
    if (text && text !== linkTarget.textContent) linkTarget.textContent = text;
  } else {
    const r = selectionInEditor();
    if (!r || r.collapsed || (text && text !== r.toString())) {
      const label = text || url.replace(/^mailto:/, "");
      exec("insertHTML", `<a class="ow-link" href="${WordModel.escapeHtml(url)}">${WordModel.escapeHtml(label)}</a>&nbsp;`);
    } else {
      exec("createLink", url);
      for (const a of els.wordView.querySelectorAll(`a[href="${CSS.escape(url)}"]`)) a.classList.add("ow-link");
    }
  }
  linkTarget = null;
  markWordDirty();
}

function removeLink() {
  closeModal(els.linkModal);
  restoreSelection(true);
  if (linkTarget && linkTarget.isConnected) {
    linkTarget.replaceWith(...linkTarget.childNodes);
  } else {
    exec("unlink");
  }
  linkTarget = null;
  markWordDirty();
}

// --- images -------------------------------------------------------------------

function readAsDataUrl(file) {
  return new Promise((resolve, reject) => {
    const fr = new FileReader();
    fr.onload = () => resolve(String(fr.result));
    fr.onerror = () => reject(fr.error || new Error("read failed"));
    fr.readAsDataURL(file);
  });
}

function imageSize(src) {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => resolve({ w: img.naturalWidth, h: img.naturalHeight });
    img.onerror = () => resolve({ w: 0, h: 0 });
    img.src = src;
  });
}

async function insertImageFile(file) {
  if (!file || !/^image\//.test(file.type)) { toast("That file is not a picture.", "error"); return; }
  if (file.size > 15 * 1024 * 1024) { toast("Pictures larger than 15 MB cannot be inserted.", "error"); return; }
  try {
    const src = await readAsDataUrl(file);
    let { w, h } = await imageSize(src);
    const maxW = textWidthPx() || MAX_IMAGE_WIDTH;
    if (w > maxW) { h = Math.round((h * maxW) / w); w = maxW; }
    const dims = w && h ? ` width="${w}" height="${h}"` : "";
    const alt = WordModel.escapeHtml(file.name.replace(/\.[^.]+$/, ""));
    exec("insertHTML", `<img class="ow-img" src="${src}" alt="${alt}"${dims}>`);
    markWordDirty();
  } catch (e) {
    toast(`Could not insert the picture: ${errMsg(e)}`, "error");
  }
}

// --- toolbar state ------------------------------------------------------------

let toolbarRaf = 0;
function scheduleToolbarState() {
  if (toolbarRaf) return;
  toolbarRaf = requestAnimationFrame(() => { toolbarRaf = 0; updateToolbarState(); });
}

function updateToolbarState() {
  if (!doc || doc.kind !== "word" || els.wordBar.hidden) return;
  const r = selectionInEditor();
  if (!r) return;
  for (const btn of els.wordBar.querySelectorAll("[data-state]")) {
    let on = false;
    try { on = document.queryCommandState(btn.dataset.state); } catch { /* ignore */ }
    btn.classList.toggle("active", on);
    btn.setAttribute("aria-pressed", String(on));
  }
  const node = r.startContainer.nodeType === 1 ? r.startContainer : r.startContainer.parentElement;
  const block = closestInEditor(node, "h1, h2, h3, h4, h5, h6, blockquote, pre, p, li");
  const tag = block ? block.tagName.toLowerCase() : "p";
  const style = { blockquote: "quote", pre: "code", li: "p", h5: "h4", h6: "h4" }[tag] || tag;
  if (els.fbStyle.value !== style) els.fbStyle.value = style;
  if (node && node.nodeType === 1) {
    const cs = getComputedStyle(node);
    const family = (cs.fontFamily || "").split(",")[0].replace(/["']/g, "").trim();
    const match = [...els.fbFont.options].find((o) => o.value && o.value.toLowerCase() === family.toLowerCase());
    els.fbFont.value = match ? match.value : "";
    const pt = Math.round((parseFloat(cs.fontSize) * 72 / 96) * 2) / 2;
    const sizeMatch = [...els.fbSize.options].find((o) => Number(o.value) === pt);
    els.fbSize.value = sizeMatch ? sizeMatch.value : "";
  }
  const inList = !!closestInEditor(node, "li");
  els.fbIndent.disabled = !inList;
  els.fbOutdent.disabled = !inList;
  els.tableTools.hidden = !closestInEditor(node, "td, th");
  // In a table the "Table" menu takes the place of narrow-screen extras.
  els.wordBar.classList.toggle("in-table", !els.tableTools.hidden);
}

function scheduleStats() {
  clearTimeout(statsTimer);
  statsTimer = setTimeout(updateStats, 300);
}

function countWords(text) {
  // Same rules as the Word count dialog.
  return WordModel.textStats(text).words;
}

function updateStats() {
  if (!doc) { els.stats.textContent = ""; return; }
  if (doc.kind === "word") {
    const text = doc.previewOnly || pageView ? els.wordPreview.innerText || "" : els.wordView.innerText || "";
    const words = countWords(text);
    const chars = text.replace(/\s/g, "").length;
    els.stats.textContent = `${words.toLocaleString()} word${words === 1 ? "" : "s"} · ${chars.toLocaleString()} character${chars === 1 ? "" : "s"}${doc.previewOnly ? " · read-only" : ""}`;
    if (!doc.previewOnly && !pageView) refreshFieldChips();
  } else if (doc.kind === "excel" && doc.model) {
    const m = doc.model;
    const sheet = m.sheets[m.active || 0];
    // ExcelView keeps the used range incrementally; scanning a multi-million
    // cell sheet on every keystroke would stall typing.
    const st = excelView && typeof excelView.getStats === "function" ? excelView.getStats() : null;
    if (st) {
      els.stats.textContent = `${st.sheet} · ${st.ref || "empty"} · ${st.sheets} sheet${st.sheets === 1 ? "" : "s"}`;
    } else if (sheet) {
      const range = ExcelModel.usedRange(sheet.cells || {});
      const dims = range
        ? `${ExcelModel.indexToCol(range.c1)}${range.r1 + 1}:${ExcelModel.indexToCol(range.c2)}${range.r2 + 1}`
        : "empty";
      const count = Object.keys(sheet.cells || {}).length;
      els.stats.textContent = `${sheet.name} · ${dims} · ${count.toLocaleString()} cells · ${m.sheets.length} sheet${m.sheets.length === 1 ? "" : "s"}`;
    }
  } else if (doc.kind === "slides" && doc.model) {
    const st = slidesView && typeof slidesView.getStats === "function" ? slidesView.getStats() : null;
    const n = (doc.model.slides || []).length;
    els.stats.textContent = st && st.slide ? `Slide ${st.slide} of ${n}` : `${n} slide${n === 1 ? "" : "s"}`;
  } else if (doc.kind === "latex") {
    const t = latexView ? latexView.getText() : doc.text || "";
    const lines = t ? t.split("\n").length : 0;
    els.stats.textContent = `${lines.toLocaleString()} lines · ${countWords(t.replace(/\\[a-zA-Z@]+\*?/g, " ")).toLocaleString()} words · ${encodingLabel(doc.encoding)}`;
  } else if (doc.kind === "text") {
    const t = els.textView.value;
    const lines = t ? t.split("\n").length : 0;
    els.stats.textContent = `${lines.toLocaleString()} lines · ${countWords(t).toLocaleString()} words · ${t.length.toLocaleString()} characters · ${encodingLabel(doc.encoding)}`;
  }
}

// --- bind the word editor -------------------------------------------------------

function bindWordEditor() {
  const ed = els.wordView;
  ed.addEventListener("input", () => {
    if (pendingFontSizePt) fixFontSizeMarkers();
    const r = selectionInEditor();
    if (r) savedRange = r.cloneRange();
    markWordDirty();
    scheduleToolbarState();
    scheduleNav();
    scheduleFindRefresh();
  });
  // Format painter: apply to the next selection made with the mouse or keyboard.
  ed.addEventListener("mouseup", () => { if (painter) setTimeout(applyPainterIfSelection, 0); });
  ed.addEventListener("keyup", (e) => { if (painter && e.shiftKey) applyPainterIfSelection(); });
  for (const type of ["keyup", "mouseup"]) {
    ed.addEventListener(type, () => {
      const r = selectionInEditor();
      if (r) savedRange = r.cloneRange();
    });
  }
  document.addEventListener("selectionchange", () => {
    const r = selectionInEditor();
    if (r) {
      savedRange = r.cloneRange();
      scheduleToolbarState();
    }
  });
  ed.addEventListener("keydown", (e) => {
    if (e.isComposing) return;
    const mod = e.ctrlKey || e.metaKey;
    if (e.key === "Tab" && !mod && !e.altKey) {
      const td = currentCell();
      if (inListItem()) {
        e.preventDefault();
        changeListLevel(!e.shiftKey);
      } else if (td) {
        e.preventDefault();
        moveToCell(td, e.shiftKey);
      } else if (!e.shiftKey) {
        e.preventDefault();
        exec("insertHTML", '<span class="ow-tab" style="white-space:pre">\t</span>');
      }
      return;
    }
    if (mod && !e.shiftKey && !e.altKey && e.key.toLowerCase() === "k") {
      e.preventDefault();
      openLinkDialog();
    } else if (mod && e.key === "Enter") {
      e.preventDefault();
      insertPageBreak();
    } else if (mod && !e.altKey && e.key.toLowerCase() === "y" && !e.shiftKey) {
      e.preventDefault();
      exec("redo");
    }
  });
  // Ctrl/Cmd+click opens links (plain clicks place the caret, like Word).
  ed.addEventListener("click", (e) => {
    const a = e.target.closest && e.target.closest("a");
    if (!a || !ed.contains(a)) return;
    e.preventDefault();
    if (e.ctrlKey || e.metaKey) {
      const href = a.getAttribute("data-href") || a.getAttribute("href") || "";
      const { url } = normalizeLinkUrl(href);
      if (url && /^(https?|mailto|tel|ftp):/i.test(url)) openTab(url);
    }
  });
  // Pasting a picture (screenshot) inserts it as an image.
  ed.addEventListener("paste", (e) => {
    const dt = e.clipboardData;
    if (!dt) return;
    const files = [...(dt.files || [])];
    const hasText = [...dt.types].some((t) => t === "text/html" || t === "text/plain");
    const img = files.find((f) => /^image\//.test(f.type));
    if (img && !hasText) {
      e.preventDefault();
      e.stopPropagation();
      insertImageFile(img);
    }
  });
  // Images dropped into the editor are inserted, not opened.
  ed.addEventListener("drop", (e) => {
    const files = [...(e.dataTransfer?.files || [])];
    const img = files.find((f) => /^image\//.test(f.type));
    if (img && files.every((f) => /^image\//.test(f.type))) {
      e.preventDefault();
      e.stopPropagation();
      const pos = document.caretRangeFromPoint ? document.caretRangeFromPoint(e.clientX, e.clientY) : null;
      if (pos) { const s = getSelection(); s.removeAllRanges(); s.addRange(pos); savedRange = pos; }
      insertImageFile(img);
    }
  });
}

function insertPageBreak() {
  exec("insertHTML", '<hr class="ow-pagebreak" data-ow="pagebreak" contenteditable="false"><p><br></p>');
  markWordDirty();
}

function bindWordBar() {
  const bar = els.wordBar;
  // Keep the editor selection when clicking toolbar buttons.
  bar.addEventListener("mousedown", (e) => {
    if (e.target.closest("button")) e.preventDefault();
  });
  bar.addEventListener("click", (e) => {
    const btn = e.target.closest("button");
    if (!btn || !doc || doc.kind !== "word" || pageView) return;
    const menuPair = WORD_MENUS.find(([b]) => b === btn);
    if (menuPair) { e.stopPropagation(); toggleMenu(menuPair[0], menuPair[1]); return; }
    const inMenu = !!btn.closest(".menu");
    if (inMenu) closeMenus();
    if (btn.dataset.cmd) {
      exec(btn.dataset.cmd);
      if (btn.dataset.cmd !== "undo" && btn.dataset.cmd !== "redo") markWordDirty();
      if (inMenu) restoreWordFocus();
      return;
    }
    if (btn.dataset.case) { applyChangeCase(btn.dataset.case); return; }
    if (btn.dataset.action) { runWordAction(btn.dataset.action); return; }
    // Block-level tools do not apply inside a header/footer band.
    if (hfActive && (btn.dataset.table || ["fb-indent", "fb-outdent", "fb-link", "fb-image", "fb-table", "fb-table-menu", "fb-pagebreak"].includes(btn.id))) {
      toast("Headers and footers hold one line of formatted text; tables, lists, pictures and links go in the document body.");
      return;
    }
    if (btn.dataset.table) { closeMenus(); restoreSelection(true); tableOp(btn.dataset.table); return; }
    switch (btn.id) {
      case "fb-painter": togglePainter(e.detail >= 2); break;
      case "fb-indent": changeListLevel(true); break;
      case "fb-outdent": changeListLevel(false); break;
      case "fb-link": openLinkDialog(); break;
      case "fb-image": els.imageInput.click(); break;
      case "fb-table": toggleTablePicker(); break;
      case "fb-table-menu": toggleMenu(els.fbTableMenu, els.menuTable); break;
      case "fb-pagebreak": insertPageBreak(); break;
      case "fb-clear":
        exec("removeFormat");
        markWordDirty();
        break;
      default: break;
    }
  });
  els.fbStyle.addEventListener("change", () => { applyBlockStyle(els.fbStyle.value); markWordDirty(); });
  els.fbFont.addEventListener("change", () => {
    if (els.fbFont.value) { exec("fontName", els.fbFont.value); markWordDirty(); }
  });
  els.fbSize.addEventListener("change", () => {
    const pt = Number(els.fbSize.value);
    if (pt > 0) { applyFontSize(pt); markWordDirty(); }
  });
  els.fbColor.addEventListener("input", () => { els.fbColorBar.style.background = els.fbColor.value; });
  els.fbColor.addEventListener("change", () => {
    els.fbColorBar.style.background = els.fbColor.value;
    exec("foreColor", els.fbColor.value);
    markWordDirty();
  });
  els.fbHighlight.addEventListener("change", () => {
    const v = els.fbHighlight.value;
    els.fbHighlight.value = "";
    if (v) { applyHighlight(v); markWordDirty(); }
  });
  els.imageInput.addEventListener("change", () => {
    const f = els.imageInput.files && els.imageInput.files[0];
    els.imageInput.value = "";
    if (f) insertImageFile(f);
  });
  // table size picker
  const ROWS = 8, COLS = 10;
  for (let r = 1; r <= ROWS; r++) {
    for (let c = 1; c <= COLS; c++) {
      const b = document.createElement("button");
      b.type = "button";
      b.dataset.r = r;
      b.dataset.c = c;
      b.setAttribute("aria-label", `${r} by ${c} table`);
      els.tpGrid.appendChild(b);
    }
  }
  const highlight = (r, c) => {
    for (const b of els.tpGrid.children) b.classList.toggle("on", b.dataset.r <= r && b.dataset.c <= c);
    els.tpLabel.textContent = r && c ? `${r} × ${c} table` : "Insert table";
  };
  els.tpGrid.addEventListener("mouseover", (e) => {
    const b = e.target.closest("button");
    if (b) highlight(Number(b.dataset.r), Number(b.dataset.c));
  });
  els.tpGrid.addEventListener("focusin", (e) => {
    const b = e.target.closest("button");
    if (b) highlight(Number(b.dataset.r), Number(b.dataset.c));
  });
  els.tpGrid.addEventListener("mousedown", (e) => e.preventDefault());
  els.tpGrid.addEventListener("click", (e) => {
    const b = e.target.closest("button");
    if (!b) return;
    closeMenus();
    insertTable(Number(b.dataset.r), Number(b.dataset.c));
  });
  els.tpGrid.addEventListener("keydown", (e) => {
    const b = e.target.closest("button");
    if (!b) return;
    let r = Number(b.dataset.r), c = Number(b.dataset.c);
    if (e.key === "ArrowRight") c = Math.min(COLS, c + 1);
    else if (e.key === "ArrowLeft") c = Math.max(1, c - 1);
    else if (e.key === "ArrowDown") r = Math.min(ROWS, r + 1);
    else if (e.key === "ArrowUp") r = Math.max(1, r - 1);
    else if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); closeMenus(); restoreSelection(); return; }
    else return;
    e.preventDefault();
    els.tpGrid.querySelector(`button[data-r="${r}"][data-c="${c}"]`).focus();
  });
}

function toggleTablePicker() {
  const open = els.tablePicker.hidden;
  closeMenus(els.tablePicker);
  els.tablePicker.hidden = !open;
  els.fbTable.setAttribute("aria-expanded", String(open));
  if (open) {
    const first = els.tpGrid.querySelector('button[data-r="2"][data-c="3"]');
    if (first) first.focus();
  }
}

// ---------------------------------------------------------------------------
// word documents: page setup, zoom, header/footer, find & replace,
// navigation pane, word count, symbols, dates, change case, format painter,
// document properties, templates
// ---------------------------------------------------------------------------

const LS = {
  zoom: "officeEditor.word.zoom",
  nav: "officeEditor.word.nav",
  units: "officeEditor.word.units",
  symbols: "officeEditor.word.recentSymbols",
  dateFormat: "officeEditor.word.dateFormat",
};
function lsGet(key) { try { return localStorage.getItem(key); } catch { return null; } }
function lsSet(key, v) { try { localStorage.setItem(key, v); } catch { /* storage disabled */ } }

function wordMeta() {
  if (!doc || !doc.model) return {};
  if (!doc.model.meta || typeof doc.model.meta !== "object") doc.model.meta = {};
  return doc.model.meta;
}

// --- page setup (meta.page, twips; same rules as both .docx writers) -----------

const PAGE_SIZES_TW = { Letter: [12240, 15840], A4: [11906, 16838], Legal: [12240, 20160], A3: [16838, 23811], A5: [8391, 11906], Executive: [10440, 15120] };
const DEFAULT_MARGINS_TW = { top: 1440, right: 1440, bottom: 1440, left: 1440, header: 720, footer: 720 };
const MARGIN_PRESETS = {
  normal: { top: 1440, bottom: 1440, left: 1440, right: 1440 },
  narrow: { top: 720, bottom: 720, left: 720, right: 720 },
  moderate: { top: 1440, bottom: 1440, left: 1080, right: 1080 },
  wide: { top: 1440, bottom: 1440, left: 2880, right: 2880 },
};

function sizeNameOf(w, h) {
  const [a, b] = w <= h ? [w, h] : [h, w];
  for (const [name, [x, y]] of Object.entries(PAGE_SIZES_TW)) if (Math.abs(a - x) <= 60 && Math.abs(b - y) <= 60) return name;
  return "custom";
}

/** meta.page resolved to concrete twips. */
function pageSetup(meta) {
  const p = (meta && typeof meta.page === "object" && meta.page) || {};
  let [w, h] = PAGE_SIZES_TW[p.size] || PAGE_SIZES_TW.Letter;
  const inRange = (v, lo, hi) => typeof v === "number" && v >= lo && v <= hi;
  if (inRange(p.width, 1440, 31680) && inRange(p.height, 1440, 31680)) { w = Math.round(p.width); h = Math.round(p.height); }
  if (p.orientation === "landscape" && w < h) [w, h] = [h, w];
  if (p.orientation === "portrait" && w > h) [w, h] = [h, w];
  const m = { ...DEFAULT_MARGINS_TW };
  if (p.margins && typeof p.margins === "object") {
    for (const k of Object.keys(m)) if (inRange(p.margins[k], 0, 14400)) m[k] = Math.round(p.margins[k]);
  }
  if (m.left + m.right + 1440 > w) { m.left = m.right = Math.min(1440, Math.floor(w / 4)); }
  if (m.top + m.bottom + 1440 > h) { m.top = m.bottom = Math.min(1440, Math.floor(h / 4)); }
  return { size: sizeNameOf(w, h), w, h, landscape: w > h, m };
}

const twPx = (tw) => tw / 15; // 1440 twips = 96 CSS px

function textWidthPx() {
  if (!doc || doc.kind !== "word") return 0;
  const ps = pageSetup(wordMeta());
  return Math.max(48, Math.floor(twPx(ps.w - ps.m.left - ps.m.right)));
}

/** The on-screen page follows the page setup. */
function applyPageSetup() {
  if (!doc || doc.kind !== "word") return;
  const ps = pageSetup(wordMeta());
  const s = els.wordSheet.style;
  s.setProperty("--pw", `${twPx(ps.w)}px`);
  s.setProperty("--ph", `${twPx(ps.h)}px`);
  s.setProperty("--mt", `${twPx(ps.m.top)}px`);
  s.setProperty("--mb", `${twPx(ps.m.bottom)}px`);
  s.setProperty("--ml", `${twPx(ps.m.left)}px`);
  s.setProperty("--mr", `${twPx(ps.m.right)}px`);
  s.setProperty("--hd", `${twPx(Math.min(ps.m.header, ps.m.top))}px`);
  s.setProperty("--fd", `${twPx(Math.min(ps.m.footer, ps.m.bottom))}px`);
}

/** Inches for en-US-like regions, centimeters elsewhere (the user can switch). */
function defaultUnits() {
  const saved = lsGet(LS.units);
  if (saved === "in" || saved === "cm") return saved;
  let region = "";
  try { region = new Intl.Locale(navigator.language || "en-US").maximize().region || ""; } catch { /* old engine */ }
  return ["US", "LR", "MM", "PR"].includes(region) ? "in" : "cm";
}
const twToUnit = (tw, u) => (u === "in" ? tw / 1440 : (tw / 1440) * 2.54);
const unitToTw = (v, u) => Math.round(u === "in" ? v * 1440 : (v / 2.54) * 1440);
const fmtUnit = (tw, u) => String(Math.round(twToUnit(tw, u) * 100) / 100);

let psState = null; // { units, w, h } while the page setup dialog is open

function openPageSetup() {
  if (!wordEditable()) return;
  const meta = wordMeta();
  const ps = pageSetup(meta);
  const units = defaultUnits();
  const known = (meta.page && PAGE_SIZES_TW[meta.page.size]) ? meta.page.size : ps.size;
  const customOpt = els.psSize.querySelector('option[value="custom"]');
  customOpt.hidden = known !== "custom";
  els.psSize.value = known;
  els.psPortrait.checked = !ps.landscape;
  els.psLandscape.checked = ps.landscape;
  els.psUnits.value = units;
  psState = { units, w: Math.min(ps.w, ps.h), h: Math.max(ps.w, ps.h), header: ps.m.header, footer: ps.m.footer };
  setMarginInputs(ps.m, units);
  syncPreset();
  els.psError.hidden = true;
  updatePsPreview();
  openModal(els.psModal, { initialFocus: els.psSize });
}

function setMarginInputs(m, units) {
  for (const [k, el] of [["top", els.psTop], ["bottom", els.psBottom], ["left", els.psLeft], ["right", els.psRight]]) {
    el.value = fmtUnit(m[k], units);
    el.step = units === "in" ? "0.05" : "0.1";
  }
}

function readMarginInputs() {
  const out = {};
  for (const [k, el] of [["top", els.psTop], ["bottom", els.psBottom], ["left", els.psLeft], ["right", els.psRight]]) {
    const v = parseFloat(String(el.value).replace(",", "."));
    out[k] = Number.isFinite(v) ? unitToTw(v, psState.units) : NaN;
  }
  return out;
}

function syncPreset() {
  const m = readMarginInputs();
  const hit = Object.entries(MARGIN_PRESETS).find(([, p]) => ["top", "bottom", "left", "right"].every((k) => Math.abs(p[k] - m[k]) <= 15));
  els.psPreset.value = hit ? hit[0] : "custom";
}

function psPaper() {
  const size = els.psSize.value;
  let [w, h] = PAGE_SIZES_TW[size] || [psState.w, psState.h];
  if (w > h) [w, h] = [h, w];
  return els.psLandscape.checked ? [h, w] : [w, h];
}

function updatePsPreview() {
  if (!psState) return;
  const [w, h] = psPaper();
  const m = readMarginInputs();
  const scale = 104 / Math.max(w, h);
  const pv = els.psPreview;
  pv.style.width = `${Math.round(w * scale)}px`;
  pv.style.height = `${Math.round(h * scale)}px`;
  const t = pv.firstElementChild;
  const ok = (v) => (Number.isFinite(v) && v >= 0 ? v : 0);
  t.style.left = `${ok(m.left) * scale}px`;
  t.style.right = `${ok(m.right) * scale}px`;
  t.style.top = `${ok(m.top) * scale}px`;
  t.style.bottom = `${ok(m.bottom) * scale}px`;
}

function submitPageSetup() {
  const [w, h] = psPaper();
  const m = readMarginInputs();
  const u = psState.units === "in" ? "in" : "cm";
  let err = "";
  if (Object.values(m).some((v) => !Number.isFinite(v) || v < 0)) err = "Margins must be numbers of 0 or more.";
  else if (Object.values(m).some((v) => v > 14400)) err = `Margins can be at most ${fmtUnit(14400, u)} ${u}.`;
  else if (m.left + m.right > w - 1440) err = `Left + right margins leave less than ${fmtUnit(1440, u)} ${u} for text on this paper.`;
  else if (m.top + m.bottom > h - 1440) err = `Top + bottom margins leave less than ${fmtUnit(1440, u)} ${u} for text on this paper.`;
  if (err) {
    els.psError.textContent = err;
    els.psError.hidden = false;
    return;
  }
  const meta = wordMeta();
  const size = els.psSize.value;
  const prev = JSON.stringify(meta.page || null);
  meta.page = {
    size,
    orientation: w > h ? "landscape" : "portrait",
    width: w,
    height: h,
    margins: { top: m.top, right: m.right, bottom: m.bottom, left: m.left, header: psState.header, footer: psState.footer },
  };
  closeModal(els.psModal);
  psState = null;
  applyPageSetup();
  if (JSON.stringify(meta.page) !== prev) {
    markWordDirty();
    setStatus(`Page: ${size === "custom" ? "custom size" : size}, ${meta.page.orientation}.`);
  }
  restoreSelection(true);
}

// --- zoom -------------------------------------------------------------------------

const ZOOM_MIN = 50;
const ZOOM_MAX = 200;
let wordZoom = (() => {
  const z = Number(lsGet(LS.zoom));
  return z >= ZOOM_MIN && z <= ZOOM_MAX ? Math.round(z) : 100;
})();

function applyZoom() {
  els.wordSheet.style.setProperty("--zoom", String(wordZoom / 100));
  els.zoomRange.value = String(wordZoom);
  els.zoomLabel.textContent = `${wordZoom}%`;
  els.zoomRange.setAttribute("aria-valuetext", `${wordZoom}%`);
}

function setZoom(z) {
  const next = Math.max(ZOOM_MIN, Math.min(ZOOM_MAX, Math.round(z)));
  if (next === wordZoom) return;
  // Keep the same part of the document in view.
  const st = els.stage;
  const ratio = st.scrollHeight > st.clientHeight ? (st.scrollTop + st.clientHeight / 2) / st.scrollHeight : 0;
  wordZoom = next;
  lsSet(LS.zoom, String(wordZoom));
  applyZoom();
  if (ratio) st.scrollTop = ratio * st.scrollHeight - st.clientHeight / 2;
  setStatus(`Zoom ${wordZoom}%`);
}
const zoomStep = (dir) => setZoom(dir > 0 ? Math.floor(wordZoom / 10) * 10 + 10 : Math.ceil(wordZoom / 10) * 10 - 10);

// --- header / footer bands ---------------------------------------------------------

let hfActive = null; // the band being edited
let hfRange = null;  // last selection inside it
const HF_FIELDS = { PAGE: "Page number", NUMPAGES: "Total number of pages" };
const HF_COMMANDS = new Set([
  "bold", "italic", "underline", "strikeThrough", "superscript", "subscript", "foreColor", "hiliteColor",
  "fontName", "fontSize", "justifyLeft", "justifyCenter", "justifyRight", "justifyFull", "removeFormat",
  "undo", "redo", "insertText",
]);

const bandKind = (band) => (band === els.wordHeader ? "header" : "footer");

function chipEl(field) {
  const s = document.createElement("span");
  s.className = "ow-field";
  s.contentEditable = "false";
  s.dataset.field = field;
  s.title = `${HF_FIELDS[field]} — filled in on every printed or saved page`;
  s.textContent = field === "PAGE" ? "1" : String(lastPageEstimate);
  return s;
}

/**
 * The text node right after a chip; a zero-width space is added when there is
 * none (otherwise Chrome may put the caret, and typed text, somewhere else).
 * readBand() drops the zero-width spaces again.
 */
const ZWSP = "​";
function caretSpotAfter(chip) {
  const next = chip.nextSibling;
  if (next && next.nodeType === 3 && next.data) return next;
  const t = document.createTextNode(ZWSP);
  chip.after(t);
  return t;
}

/** Turn "{PAGE}" / "{NUMPAGES}" text inside `root` into field chips. */
function chipify(root) {
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  const hits = [];
  for (let n = walker.nextNode(); n; n = walker.nextNode()) if (/\{(PAGE|NUMPAGES)\}/.test(n.data)) hits.push(n);
  for (const n of hits) {
    const frag = document.createDocumentFragment();
    for (const part of n.data.split(/(\{PAGE\}|\{NUMPAGES\})/)) {
      if (!part) continue;
      if (part === "{PAGE}" || part === "{NUMPAGES}") frag.appendChild(chipEl(part.slice(1, -1)));
      else frag.appendChild(document.createTextNode(part));
    }
    n.replaceWith(frag);
  }
  for (const chip of root.querySelectorAll(".ow-field")) caretSpotAfter(chip);
}

function renderBand(band, hf) {
  band.innerHTML = "";
  if (!hf || !Array.isArray(hf.runs) || !hf.runs.some((r) => r && r.t)) return;
  band.innerHTML = renderWordHtml({ blocks: [{ type: "p", runs: hf.runs, align: hf.align }] });
  chipify(band);
}

function renderHeaderFooter() {
  const meta = doc && doc.model ? doc.model.meta || {} : {};
  renderBand(els.wordHeader, meta.header);
  renderBand(els.wordFooter, meta.footer);
}

/** A band's DOM -> { runs, align } (null when empty). */
function readBand(band) {
  const clone = band.cloneNode(true);
  for (const chip of clone.querySelectorAll(".ow-field[data-field]")) {
    // Text that ended up inside a chip (it should not) is kept after the field.
    const extra = chip.textContent.replace(/^\d+/, "");
    chip.replaceWith(document.createTextNode(`{${chip.dataset.field}}${extra}`));
  }
  const tw = document.createTreeWalker(clone, NodeFilter.SHOW_TEXT);
  for (let n = tw.nextNode(); n; n = tw.nextNode()) if (n.data.includes(ZWSP)) n.data = n.data.replaceAll(ZWSP, "");
  const cs = getComputedStyle(band);
  const blocks = WordModel.parseBlocks(clone, {
    defaults: { color: WordModel.colorToHex(cs.color), sizePt: Math.round(parseFloat(cs.fontSize) * 0.75 * 2) / 2 },
  });
  const runs = [];
  let align;
  for (const b of blocks) {
    if (!Array.isArray(b.runs)) continue;
    if (!align && b.align) align = b.align;
    if (runs.length) runs.push({ t: "\n" });
    for (const r of b.runs) if (r && r.t) runs.push({ ...r });
  }
  while (runs.length && /^\n+$/.test(runs[runs.length - 1].t)) runs.pop();
  if (!runs.some((r) => r.t.trim())) return null;
  const merged = [];
  const same = (a, b) => ["b", "i", "u", "s", "sup", "sub", "color", "size", "font", "link", "hl"].every((k) => (a[k] || undefined) === (b[k] || undefined));
  for (const r of runs) {
    const last = merged[merged.length - 1];
    if (last && same(last, r)) last.t += r.t;
    else merged.push(r);
  }
  const hf = { runs: merged };
  if (align && align !== "left") hf.align = align;
  return hf;
}

function syncBand(band) {
  if (!doc || doc.kind !== "word" || !doc.model) return;
  const meta = wordMeta();
  const kind = bandKind(band);
  const hf = readBand(band);
  const before = JSON.stringify(meta[kind] || null);
  if (hf) meta[kind] = hf;
  else delete meta[kind];
  if (JSON.stringify(hf || null) !== before) markWordDirty();
}

function activateBand(band, { caretAtEnd = false } = {}) {
  if (!wordEditable()) return;
  if (hfActive && hfActive !== band) deactivateBand();
  hfActive = band;
  band.classList.add("hf-active");
  els.wordSheet.classList.add("hf-editing");
  els.hfToolsLabel.textContent = band === els.wordHeader ? "Header" : "Footer";
  els.hfTools.hidden = false;
  // Typed text must land in a paragraph (the band is a flex column: loose
  // text and field chips directly inside it would each become a line).
  if (!band.firstElementChild || !/^(P|DIV)$/.test(band.firstElementChild.tagName)) {
    const p = document.createElement("p");
    while (band.firstChild) p.appendChild(band.firstChild);
    if (!p.firstChild) p.appendChild(document.createElement("br"));
    band.appendChild(p);
    caretAtEnd = true;
  }
  positionHfTools();
  if (document.activeElement !== band || caretAtEnd) {
    band.focus({ preventScroll: false });
    if (caretAtEnd) {
      const r = document.createRange();
      const last = band.lastElementChild || band;
      if (!last.textContent) r.setStart(last, 0);
      else { r.selectNodeContents(last); r.collapse(false); }
      r.collapse(true);
      const sel = getSelection();
      sel.removeAllRanges();
      sel.addRange(r);
      hfRange = r.cloneRange();
    }
  }
  band.scrollIntoView({ block: "nearest" });
  setStatus(`Editing the ${bandKind(band)}: it repeats on every page. Esc or Close returns to the document.`);
}

function positionHfTools() {
  if (!hfActive) return;
  const band = hfActive;
  const tools = els.hfTools;
  // Coordinates inside the (zoomed) sheet are unzoomed layout pixels.
  if (band === els.wordHeader) tools.style.top = `${band.offsetTop + band.offsetHeight + 4}px`;
  else tools.style.top = `${Math.max(0, band.offsetTop - 40)}px`;
}

function deactivateBand() {
  if (!hfActive) { els.hfTools.hidden = true; return; }
  const band = hfActive;
  hfActive = null;
  hfRange = null;
  band.classList.remove("hf-active");
  els.wordSheet.classList.remove("hf-editing");
  els.hfTools.hidden = true;
  if (doc && doc.kind === "word") {
    syncBand(band);
    // Re-render so typed "{PAGE}" text and stray markup become clean fields.
    renderBand(band, wordMeta()[bandKind(band)]);
  }
}

function closeBand() {
  deactivateBand();
  setStatus("");
  restoreSelection(true);
}

/** Put the selection back where the user was: header/footer band or document. */
function restoreWordFocus() {
  if (hfActive && hfActive.isConnected) {
    hfActive.focus({ preventScroll: true });
    if (hfRange && hfActive.contains(hfRange.commonAncestorContainer)) {
      const sel = getSelection();
      sel.removeAllRanges();
      sel.addRange(hfRange);
    }
    return;
  }
  restoreSelection(true);
}

function execInBand(cmd, value) {
  const band = hfActive;
  if (!HF_COMMANDS.has(cmd)) {
    toast("Headers and footers hold one line of formatted text; tables, lists, pictures and links go in the document body.");
    return false;
  }
  restoreWordFocus();
  let ok = false;
  try { ok = document.execCommand(cmd, false, value); } catch (e) { console.warn(cmd, e); }
  if (pendingFontSizePt) fixFontSizeMarkers();
  syncBand(band);
  return ok;
}

function insertField(field) {
  if (!hfActive) return;
  const band = hfActive;
  restoreWordFocus();
  const sel = getSelection();
  if (!sel.rangeCount || !band.contains(sel.getRangeAt(0).commonAncestorContainer)) {
    const r = document.createRange();
    r.selectNodeContents(band);
    r.collapse(false);
    sel.removeAllRanges();
    sel.addRange(r);
  }
  // Inserted directly (insertHTML would leave the caret inside the chip).
  const r = sel.getRangeAt(0);
  r.deleteContents();
  const chip = chipEl(field);
  r.insertNode(chip);
  const caret = document.createRange();
  const after = caretSpotAfter(chip);
  caret.setStart(after, after.data.startsWith(ZWSP) ? 1 : 0);
  caret.collapse(true);
  sel.removeAllRanges();
  sel.addRange(caret);
  hfRange = caret.cloneRange();
  syncBand(band);
}

/** Insert > Page numbers: "Page {PAGE} of {NUMPAGES}" in the footer. */
function addPageNumbers() {
  if (!wordEditable()) return;
  deactivateBand();
  const meta = wordMeta();
  const cur = meta.footer && Array.isArray(meta.footer.runs) ? meta.footer : null;
  if (cur && cur.runs.some((r) => /\{PAGE\}/.test(r.t || ""))) {
    toast("The footer already shows page numbers.");
    activateBand(els.wordFooter, { caretAtEnd: true });
    return;
  }
  const runs = cur ? [...cur.runs, { t: "    " }] : [];
  runs.push({ t: "Page {PAGE} of {NUMPAGES}" });
  meta.footer = { runs, align: cur ? cur.align : "center" };
  if (!meta.footer.align) delete meta.footer.align;
  renderBand(els.wordFooter, meta.footer);
  markWordDirty();
  els.wordFooter.scrollIntoView({ block: "nearest" });
  toast("Page numbers (“Page X of Y”) added to the footer. Click the footer to change them.", "ok");
}

let lastPageEstimate = 1;
function refreshFieldChips() {
  const chips = els.wordSheet.querySelectorAll('.ow-field[data-field="NUMPAGES"]');
  if (!chips.length) return;
  lastPageEstimate = estimatePages();
  for (const c of chips) if (c.textContent !== String(lastPageEstimate)) c.textContent = String(lastPageEstimate);
}

// --- print: @page size, margins and running header/footer ---------------------------

function cssString(s) {
  return `"${String(s).replace(/\\/g, "\\\\").replace(/"/g, '\\"').replace(/\n/g, "\\A ").replace(/[\u0000-\u001f]/g, " ")}"`;
}

function hfPrintContent(hf) {
  const text = (hf.runs || []).map((r) => String(r.t || "")).join("");
  const parts = text.split(/(\{PAGE\}|\{NUMPAGES\})/).filter(Boolean)
    .map((p) => (p === "{PAGE}" ? "counter(page)" : p === "{NUMPAGES}" ? "counter(pages)" : cssString(p)));
  return parts.join(" ");
}

function updateWordPrintStyle() {
  let style = document.getElementById("ow-print-page");
  if (!style) {
    style = document.createElement("style");
    style.id = "ow-print-page";
    style.media = "print";
    document.head.appendChild(style);
  }
  if (!doc || doc.kind !== "word" || doc.previewOnly || pageView) { style.textContent = ""; return; }
  const meta = wordMeta();
  const ps = pageSetup(meta);
  const inch = (tw) => `${(tw / 1440).toFixed(3)}in`;
  let boxes = "";
  for (const [kind, edge] of [["header", "top"], ["footer", "bottom"]]) {
    const hf = meta[kind];
    if (!hf || !Array.isArray(hf.runs) || !hf.runs.some((r) => r && r.t && String(r.t).trim())) continue;
    const pos = hf.align === "center" ? "center" : hf.align === "right" ? "right" : "left";
    const box = pos === "center" ? `${edge}-center` : pos === "right" ? `${edge}-right` : `${edge}-left`;
    boxes += `@${box} { content: ${hfPrintContent(hf)}; white-space: pre; font: 10pt Calibri, Carlito, Arial, sans-serif; color: #333; }`;
  }
  style.textContent = `@page { size: ${inch(ps.w)} ${inch(ps.h)}; margin: ${inch(ps.m.top)} ${inch(ps.m.right)} ${inch(ps.m.bottom)} ${inch(ps.m.left)}; ${boxes} }`;
}

// --- find & replace ----------------------------------------------------------------

const find = { matches: [], index: -1, timer: 0, query: "" };
const hasHighlights = () => typeof CSS !== "undefined" && CSS.highlights && typeof Highlight === "function";

function findOptions() {
  return {
    matchCase: els.findCase.getAttribute("aria-pressed") === "true",
    wholeWord: els.findWord.getAttribute("aria-pressed") === "true",
    limit: 5000,
  };
}

/** Text of the editor grouped by block, with the matches mapped back to DOM ranges. */
function computeFindMatches(query, opts) {
  const out = [];
  if (!query) return out;
  const walker = document.createTreeWalker(els.wordView, NodeFilter.SHOW_TEXT);
  let block = null;
  let segs = [];
  let text = "";
  const flush = () => {
    if (!text) return;
    for (const [s, e] of WordModel.findMatches(text, query, { ...opts, limit: opts.limit - out.length })) {
      const a = locate(segs, s, false);
      const b = locate(segs, e, true);
      if (!a || !b) continue;
      const r = document.createRange();
      r.setStart(a.node, a.offset);
      r.setEnd(b.node, b.offset);
      out.push(r);
    }
  };
  for (let n = walker.nextNode(); n && out.length < opts.limit; n = walker.nextNode()) {
    const blk = closestBlock(n) || els.wordView;
    if (blk !== block) { flush(); block = blk; segs = []; text = ""; }
    segs.push({ node: n, start: text.length });
    // Non-breaking spaces (used for runs of spaces) match typed spaces.
    text += n.data.replace(/ /g, " ");
  }
  flush();
  return out;
}

function locate(segs, pos, isEnd) {
  for (let i = segs.length - 1; i >= 0; i--) {
    const s = segs[i];
    if (pos > s.start || (pos === s.start && !isEnd) || (pos === s.start && i === 0)) {
      return { node: s.node, offset: Math.min(pos - s.start, s.node.data.length) };
    }
  }
  return null;
}

function paintFindHighlights() {
  if (!hasHighlights()) return;
  if (!find.matches.length || els.findBar.hidden) {
    CSS.highlights.delete("ow-find");
    CSS.highlights.delete("ow-find-current");
    return;
  }
  CSS.highlights.set("ow-find", new Highlight(...find.matches));
  const cur = find.matches[find.index];
  if (cur) {
    const h = new Highlight(cur);
    h.priority = 1;
    CSS.highlights.set("ow-find-current", h);
  } else CSS.highlights.delete("ow-find-current");
}

function updateFindCount() {
  const n = find.matches.length;
  const q = els.findInput.value;
  els.findCount.textContent = !q ? "" : !n ? "No results" : `${find.index + 1} of ${n >= 5000 ? "5000+" : n.toLocaleString()}`;
  els.findInput.classList.toggle("no-match", !!q && !n);
  for (const b of [els.findPrev, els.findNext, els.replaceOne, els.replaceAll]) b.disabled = !n;
}

/** Recompute matches; the current one is the first at/after `from` (a Range) or the caret. */
function runFind({ from = null, keepIndex = false } = {}) {
  const q = els.findInput.value;
  const prevIndex = find.index;
  find.query = q;
  find.matches = computeFindMatches(q, findOptions());
  if (!find.matches.length) find.index = -1;
  else if (keepIndex) find.index = Math.min(Math.max(0, prevIndex), find.matches.length - 1);
  else {
    const ref = from || (savedRange && els.wordView.contains(savedRange.startContainer) ? savedRange : null);
    let i = 0;
    if (ref) {
      i = find.matches.findIndex((m) => m.compareBoundaryPoints(Range.START_TO_START, ref) >= 0);
      if (i < 0) i = 0;
    }
    find.index = i;
  }
  paintFindHighlights();
  updateFindCount();
  if (find.index >= 0 && !keepIndex) revealMatch(find.matches[find.index]);
}

function revealMatch(range) {
  if (!range) return;
  const rect = range.getBoundingClientRect();
  const st = els.stage.getBoundingClientRect();
  const margin = 60;
  if (rect.top < st.top + margin || rect.bottom > st.bottom - margin) {
    els.stage.scrollTop += rect.top - st.top - st.height / 3;
  }
  if (!hasHighlights()) {
    // Fallback: show the match as the selection.
    const sel = getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
  }
}

function gotoMatch(delta) {
  if (!find.matches.length) { runFind(); return; }
  if (find.matches.some((m) => !m.startContainer.isConnected)) runFind({ keepIndex: true });
  const n = find.matches.length;
  if (!n) return;
  find.index = (find.index + delta + n) % n;
  paintFindHighlights();
  updateFindCount();
  revealMatch(find.matches[find.index]);
}

function scheduleFindRefresh() {
  if (els.findBar.hidden) return;
  clearTimeout(find.timer);
  find.timer = setTimeout(() => runFind({ keepIndex: true }), 250);
}

function setReplaceVisible(on) {
  els.findReplaceRow.hidden = !on;
  els.findToggleReplace.setAttribute("aria-expanded", String(on));
  els.findToggleReplace.textContent = on ? "▾" : "▸";
  const label = on ? "Hide replace" : "Show replace (Ctrl+H)";
  els.findToggleReplace.title = label;
  els.findToggleReplace.setAttribute("aria-label", on ? "Hide replace" : "Show replace");
}

function openFind(replace = false) {
  if (!wordEditable()) return;
  if (hfActive) deactivateBand();
  { const r = selectionInEditor(); if (r) savedRange = r.cloneRange(); }
  const wasOpen = !els.findBar.hidden;
  els.findBar.hidden = false;
  if (replace) setReplaceVisible(true);
  else if (!wasOpen) setReplaceVisible(false);
  const r = savedRange && els.wordView.contains(savedRange.commonAncestorContainer) ? savedRange : null;
  const picked = r && !r.collapsed ? r.toString() : "";
  if (picked && picked.length <= 200 && !/[\n\r]/.test(picked)) els.findInput.value = picked;
  const target = replace && els.findInput.value ? els.replaceInput : els.findInput;
  target.focus();
  target.select();
  runFind();
}

function closeFind({ restore = true } = {}) {
  if (els.findBar.hidden) return;
  clearTimeout(find.timer);
  const cur = find.matches[find.index];
  els.findBar.hidden = true;
  find.matches = [];
  find.index = -1;
  paintFindHighlights();
  if (!restore || !doc || doc.kind !== "word") return;
  els.wordView.focus({ preventScroll: true });
  const sel = getSelection();
  if (cur && cur.startContainer.isConnected && els.wordView.contains(cur.startContainer)) {
    sel.removeAllRanges();
    sel.addRange(cur);
    savedRange = cur.cloneRange();
  } else restoreSelection(true);
}

/** Replace `range` through the editor's undo stack. */
function replaceRange(range, text) {
  els.wordView.focus({ preventScroll: true });
  const sel = getSelection();
  sel.removeAllRanges();
  sel.addRange(range);
  let ok = false;
  try { ok = text ? document.execCommand("insertText", false, text) : document.execCommand("delete"); } catch { ok = false; }
  if (!ok) {
    range.deleteContents();
    if (text) range.insertNode(document.createTextNode(text));
  }
}

function replaceCurrent() {
  if (!find.matches.length) return;
  const cur = find.matches[find.index];
  if (!cur || !cur.startContainer.isConnected) { runFind({ keepIndex: true }); return; }
  const after = cur.cloneRange();
  replaceRange(cur, els.replaceInput.value);
  const sel = getSelection();
  const at = sel.rangeCount ? sel.getRangeAt(0).cloneRange() : after;
  markWordDirty();
  runFind({ from: at });
  els.replaceInput.focus();
}

function replaceAllMatches() {
  const matches = computeFindMatches(els.findInput.value, { ...findOptions(), limit: 100000 });
  if (!matches.length) { updateFindCount(); return; }
  const text = els.replaceInput.value;
  // Last to first: earlier ranges stay valid while later text changes.
  for (let i = matches.length - 1; i >= 0; i--) {
    if (matches[i].startContainer.isConnected) replaceRange(matches[i], text);
  }
  markWordDirty();
  runFind();
  const n = matches.length;
  toast(`Replaced ${n.toLocaleString()} match${n === 1 ? "" : "es"}.`, "ok");
  setStatus(`Replaced ${n.toLocaleString()} match${n === 1 ? "" : "es"} (Ctrl+Z undoes them one by one).`);
  els.replaceInput.focus();
}

// --- navigation pane -------------------------------------------------------------

let navOpen = lsGet(LS.nav) === "1";
let navHeadings = [];
let navTimer = 0;
let navRaf = 0;
let navPin = null; // { i, scrollTop } after a click in the pane

function updateNavVisibility() {
  const show = navOpen && wordEditable();
  els.navPane.hidden = !show;
  els.miNav.setAttribute("aria-checked", String(navOpen));
}

function toggleNav(force) {
  navOpen = typeof force === "boolean" ? force : !navOpen;
  lsSet(LS.nav, navOpen ? "1" : "0");
  updateNavVisibility();
  if (navOpen && wordEditable()) {
    buildNav();
    const cur = els.navList.querySelector(".nav-item.current") || els.navList.querySelector(".nav-item");
    if (cur) cur.focus();
    else els.navClose.focus();
  } else if (!navOpen && wordEditable()) restoreWordFocus();
}

function scheduleNav() {
  if (els.navPane.hidden) return;
  clearTimeout(navTimer);
  navTimer = setTimeout(buildNav, 400);
}

function buildNav() {
  if (els.navPane.hidden) return;
  const hs = [...els.wordView.querySelectorAll("h1, h2, h3, h4, h5, h6")].slice(0, 3000);
  navHeadings = hs;
  const frag = document.createDocumentFragment();
  hs.forEach((h, i) => {
    const b = document.createElement("button");
    b.type = "button";
    const level = Number(h.tagName[1]);
    b.className = `nav-item l${level}`;
    b.dataset.i = String(i);
    b.setAttribute("role", "listitem");
    const text = (h.textContent || "").replace(/\s+/g, " ").trim();
    b.textContent = text || "(empty heading)";
    b.title = text;
    b.setAttribute("aria-label", `Heading ${level}: ${text || "empty"}`);
    frag.appendChild(b);
  });
  els.navList.replaceChildren(frag);
  els.navEmpty.hidden = hs.length > 0;
  updateNavCurrent();
}

function updateNavCurrent() {
  if (els.navPane.hidden || !navHeadings.length) return;
  const top = els.stage.getBoundingClientRect().top + 90;
  let cur = 0;
  for (let i = 0; i < navHeadings.length; i++) {
    if (navHeadings[i].getBoundingClientRect().top <= top) cur = i;
    else break;
  }
  // A heading picked in the pane stays current until the user scrolls away
  // (headings near the end cannot reach the top of the window).
  if (navPin && navPin.i < navHeadings.length && Math.abs(els.stage.scrollTop - navPin.scrollTop) < 8) cur = navPin.i;
  else navPin = null;
  const items = els.navList.children;
  for (let i = 0; i < items.length; i++) {
    const on = i === cur;
    if (items[i].classList.contains("current") !== on) {
      items[i].classList.toggle("current", on);
      if (on) {
        items[i].setAttribute("aria-current", "location");
        const lr = els.navList.getBoundingClientRect();
        const ir = items[i].getBoundingClientRect();
        if (ir.top < lr.top || ir.bottom > lr.bottom) els.navList.scrollTop += ir.top - lr.top - lr.height / 3;
      } else items[i].removeAttribute("aria-current");
    }
  }
}

function goToHeading(i) {
  let h = navHeadings[i];
  if (!h || !h.isConnected) { buildNav(); h = navHeadings[i]; }
  if (!h) return;
  const st = els.stage.getBoundingClientRect();
  els.stage.scrollTop += h.getBoundingClientRect().top - st.top - 16;
  const r = document.createRange();
  r.selectNodeContents(h);
  r.collapse(true);
  els.wordView.focus({ preventScroll: true });
  const sel = getSelection();
  sel.removeAllRanges();
  sel.addRange(r);
  savedRange = r.cloneRange();
  navPin = { i, scrollTop: els.stage.scrollTop };
  updateNavCurrent();
}

// --- word count --------------------------------------------------------------------

/** Rendered line count of the leaf blocks (null when too big to measure quickly). */
function layoutLines(range = null) {
  const tops = (rects) => new Set([...rects].filter((r) => r.width > 0 || r.height > 0).map((r) => Math.round(r.top))).size;
  if (range) return tops(range.getClientRects());
  const blocks = els.wordView.querySelectorAll("p, li, h1, h2, h3, h4, h5, h6, pre, blockquote");
  if (blocks.length > 6000) return null;
  let n = 0;
  const r = document.createRange();
  for (const b of blocks) {
    if (b.querySelector("p, li, h1, h2, h3, h4, h5, h6, pre, blockquote, table")) continue;
    r.selectNodeContents(b);
    n += Math.max(1, tops(r.getClientRects()));
  }
  return n;
}

/** Estimated page count from the laid-out body and the page setup. */
function estimatePages() {
  if (!doc || doc.kind !== "word" || els.wordView.hidden) return 1;
  const ps = pageSetup(wordMeta());
  const bodyH = Math.max(96, twPx(ps.h - ps.m.top - ps.m.bottom));
  const kids = [...els.wordView.children];
  if (!kids.length) return 1;
  let pages = 0;
  let segStart = kids[0].offsetTop;
  const flush = (end) => { pages += Math.max(1, Math.ceil((end - segStart) / bodyH - 0.03)); };
  for (const k of kids) {
    if (k.matches("hr.ow-pagebreak")) { flush(k.offsetTop); segStart = k.offsetTop + k.offsetHeight; }
  }
  const last = kids[kids.length - 1];
  flush(last.offsetTop + last.offsetHeight);
  return Math.max(1, pages);
}

/** Text of a range with line breaks between blocks (without touching the selection). */
function rangeText(r) {
  const box = document.createElement("div");
  box.className = "ow-page";
  box.setAttribute("aria-hidden", "true");
  box.style.cssText = "position:fixed;left:-20000px;top:0;width:700px;opacity:0;pointer-events:none;";
  box.appendChild(r.cloneContents());
  document.body.appendChild(box);
  const t = box.innerText || "";
  box.remove();
  return t;
}

function wordStats() {
  const text = els.wordView.innerText || "";
  const s = WordModel.textStats(text);
  const model = currentWordModel();
  let paragraphs = 0;
  const countBlocks = (blocks) => {
    for (const b of blocks || []) {
      if (Array.isArray(b.runs)) { if (b.runs.some((r) => r.t && r.t.trim())) paragraphs++; }
      else if (b.type === "table") for (const row of b.rows || []) for (const c of row) countBlocks(c.blocks);
    }
  };
  countBlocks(model ? model.blocks : []);
  let lines = layoutLines();
  if (lines == null) lines = text.split("\n").filter((l) => l.trim()).length;
  const out = { pages: estimatePages(), words: s.words, chars: s.chars, charsNoSpaces: s.charsNoSpaces, paragraphs, lines, selection: null };
  const r = savedRange && els.wordView.contains(savedRange.commonAncestorContainer) && !savedRange.collapsed ? savedRange : null;
  if (r) {
    const st = rangeText(r);
    const ss = WordModel.textStats(st);
    out.selection = {
      words: ss.words, chars: ss.chars, charsNoSpaces: ss.charsNoSpaces,
      paragraphs: st.split(/\n+/).filter((l) => l.trim()).length,
      lines: layoutLines(r) || 1,
      pages: null,
    };
  }
  return out;
}

function openWordCount() {
  if (!wordEditable()) return;
  { const r = selectionInEditor(); if (r) savedRange = r.cloneRange(); }
  const st = wordStats();
  const rows = [
    ["Pages (estimate)", "pages"], ["Words", "words"], ["Characters (no spaces)", "charsNoSpaces"],
    ["Characters (with spaces)", "chars"], ["Paragraphs", "paragraphs"], ["Lines (estimate)", "lines"],
  ];
  const fmt = (v) => (v == null ? "—" : Number(v).toLocaleString());
  els.wcBody.replaceChildren(...rows.map(([label, key]) => {
    const tr = document.createElement("tr");
    const th = document.createElement("td");
    th.textContent = label;
    const sel = document.createElement("td");
    sel.className = "wc-sel";
    sel.textContent = st.selection ? fmt(st.selection[key]) : "";
    const all = document.createElement("td");
    all.textContent = fmt(st[key]);
    all.dataset.key = key;
    tr.append(th, sel, all);
    return tr;
  }));
  const table = els.wcBody.closest("table");
  table.classList.toggle("no-sel", !st.selection);
  table.querySelector("#wc-col-sel").classList.add("wc-sel");
  openModal(els.wcModal, { initialFocus: els.wcClose });
}

// --- insert symbol / date & time ---------------------------------------------------

const SYMBOL_GROUPS = [
  ["Common", "©®™§¶†‡•…–—‘’“”«»‹›°±×÷½¼¾‰№✓✗★☆♥☐☑☒→←"],
  ["Arrows", "←→↑↓↔↕⇐⇒⇑⇓⇔↖↗↘↙↩↪⟵⟶⟷➜➔▲▼◀▶"],
  ["Math", "≈≠≤≥∞√∛∑∏∫∂∆∇∈∉⊂⊃⊆⊇∪∩∧∨¬∀∃∅∝∠⊥∥≡≅∼·⋅′″¹²³⁴ⁿ₀₁₂₃µπ"],
  ["Currency", "$€£¥¢₹₽₩₪₺₫₴₦₱฿₿¤ƒ"],
  ["Greek", "αβγδεζηθικλμνξοπρστυφχψωΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩ"],
  ["Latin", "àáâãäåæçèéêëìíîïñòóôõöøœßùúûüýÿÀÁÂÄÅÆÇÈÉÊËÎÏÑÓÔÖØŒÙÚÛÜ¿¡"],
  ["Shapes", "■□▪▫●○◆◇◊▲△▼▽►◄♠♣♦♪♫☀☁☂☎✉✂✎✔✖❤"],
];
const SYMBOL_NAMES = {
  "©": "Copyright", "®": "Registered", "™": "Trademark", "§": "Section", "¶": "Pilcrow (paragraph)", "†": "Dagger",
  "‡": "Double dagger", "•": "Bullet", "…": "Ellipsis", "–": "En dash", "—": "Em dash", "°": "Degree", "±": "Plus-minus",
  "×": "Multiplication", "÷": "Division", "½": "One half", "¼": "One quarter", "¾": "Three quarters", "‰": "Per mille",
  "№": "Numero", "✓": "Check mark", "✗": "Ballot X", "★": "Black star", "☆": "White star", "♥": "Heart",
  "☐": "Ballot box", "☑": "Ballot box with check", "☒": "Ballot box with X", "←": "Left arrow", "→": "Right arrow",
  "↑": "Up arrow", "↓": "Down arrow", "↔": "Left-right arrow", "⇒": "Rightwards double arrow", "⇔": "Left-right double arrow",
  "≈": "Almost equal", "≠": "Not equal", "≤": "Less than or equal", "≥": "Greater than or equal", "∞": "Infinity",
  "√": "Square root", "∑": "Summation", "∏": "Product", "∫": "Integral", "∂": "Partial differential", "∆": "Increment",
  "∇": "Nabla", "∈": "Element of", "∉": "Not an element of", "∀": "For all", "∃": "There exists", "∅": "Empty set",
  "€": "Euro", "£": "Pound", "¥": "Yen", "¢": "Cent", "₹": "Indian rupee", "₽": "Ruble", "₩": "Won", "₿": "Bitcoin",
  "µ": "Micro", "π": "Pi", "α": "Alpha", "β": "Beta", "γ": "Gamma", "δ": "Delta", "λ": "Lambda", "Ω": "Omega",
  "Σ": "Sigma", "‘": "Left single quote", "’": "Right single quote", "“": "Left double quote", "”": "Right double quote",
  "«": "Left guillemet", "»": "Right guillemet", "¿": "Inverted question mark", "¡": "Inverted exclamation mark",
};
let symGroup = null;

function recentSymbols() {
  try {
    const v = JSON.parse(lsGet(LS.symbols) || "[]");
    return Array.isArray(v) ? v.filter((s) => typeof s === "string" && s.length <= 4).slice(0, 20) : [];
  } catch { return []; }
}

function symbolLabel(ch) {
  const cp = ch.codePointAt(0).toString(16).toUpperCase().padStart(4, "0");
  return `${SYMBOL_NAMES[ch] || "Symbol"} (U+${cp})`;
}

function openSymbols() {
  if (!wordEditable()) return;
  { const r = selectionInEditor(); if (r) savedRange = r.cloneRange(); }
  const recent = recentSymbols();
  const groups = recent.length ? [["Recently used", recent.join("")], ...SYMBOL_GROUPS] : SYMBOL_GROUPS;
  els.symTabs.replaceChildren(...groups.map(([name], i) => {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "sym-tab";
    b.setAttribute("role", "tab");
    b.textContent = name;
    b.dataset.i = String(i);
    return b;
  }));
  els.symTabs._groups = groups;
  showSymbolGroup(symGroup != null && symGroup < groups.length && symGroup > 0 ? symGroup : 0);
  els.symName.textContent = "Click a symbol to insert it. Arrow keys move, Enter inserts.";
  els.symPreview.textContent = "";
  openModal(els.symModal, { initialFocus: els.symGrid.querySelector(".sym-cell") || els.symClose });
}

function showSymbolGroup(i) {
  const groups = els.symTabs._groups || SYMBOL_GROUPS;
  symGroup = i;
  for (const t of els.symTabs.children) t.setAttribute("aria-selected", String(Number(t.dataset.i) === i));
  els.symGrid.replaceChildren(...[...groups[i][1]].map((ch) => {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "sym-cell";
    b.textContent = ch;
    b.dataset.ch = ch;
    b.title = symbolLabel(ch);
    b.setAttribute("aria-label", symbolLabel(ch));
    return b;
  }));
}

/** Insert plain text at the caret of the document or of the header/footer being edited. */
function insertTextAtCaret(text) {
  if (hfActive) { execInBand("insertText", text); return; }
  restoreSelection(true);
  exec("insertText", text);
  markWordDirty();
}

function insertSymbol(ch) {
  insertTextAtCaret(ch);
  const recent = [ch, ...recentSymbols().filter((s) => s !== ch)].slice(0, 20);
  lsSet(LS.symbols, JSON.stringify(recent));
  els.symPreview.textContent = ch;
  els.symName.textContent = `Inserted ${symbolLabel(ch)}. Pick another, or Close.`;
}

function dateFormats(now = new Date()) {
  const f = (o) => { try { return new Intl.DateTimeFormat(undefined, o).format(now); } catch { return now.toLocaleString(); } };
  const iso = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-${String(now.getDate()).padStart(2, "0")}`;
  const list = [
    f({ dateStyle: "short" }), f({ dateStyle: "long" }), f({ dateStyle: "full" }), f({ dateStyle: "medium" }), iso,
    f({ timeStyle: "short" }), f({ dateStyle: "short", timeStyle: "short" }), f({ dateStyle: "long", timeStyle: "short" }),
  ];
  return [...new Set(list)];
}

function openDateTime() {
  if (!wordEditable()) return;
  { const r = selectionInEditor(); if (r) savedRange = r.cloneRange(); }
  const formats = dateFormats();
  const last = Number(lsGet(LS.dateFormat)) || 0;
  els.dtList.replaceChildren(...formats.map((s, i) => {
    const lab = document.createElement("label");
    lab.className = "dt-option";
    const inp = document.createElement("input");
    inp.type = "radio";
    inp.name = "dt-format";
    inp.value = s;
    inp.dataset.i = String(i);
    if (i === (last < formats.length ? last : 0)) inp.checked = true;
    lab.append(inp, document.createTextNode(s));
    return lab;
  }));
  openModal(els.dtModal, { initialFocus: els.dtList.querySelector("input:checked") });
}

function submitDateTime() {
  const inp = els.dtList.querySelector("input:checked");
  closeModal(els.dtModal);
  if (!inp) return;
  lsSet(LS.dateFormat, inp.dataset.i);
  insertTextAtCaret(inp.value);
}

// --- change case ---------------------------------------------------------------------

function applyChangeCase(mode) {
  if (!wordEditable()) return;
  const root = hfActive || els.wordView;
  restoreWordFocus();
  const sel = getSelection();
  const r = sel.rangeCount ? sel.getRangeAt(0) : null;
  if (!r || r.collapsed || !root.contains(r.commonAncestorContainer)) {
    toast("Select the text whose case you want to change first.");
    return;
  }
  const segs = [];
  const walker = document.createTreeWalker(r.commonAncestorContainer.nodeType === 3 ? r.commonAncestorContainer.parentNode : r.commonAncestorContainer, NodeFilter.SHOW_TEXT);
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    if (!r.intersectsNode(n)) continue;
    const start = n === r.startContainer ? r.startOffset : 0;
    const end = n === r.endContainer ? r.endOffset : n.data.length;
    if (end > start) segs.push({ node: n, start, end });
  }
  const full = segs.map((s) => s.node.data.slice(s.start, s.end)).join("");
  const next = WordModel.changeCase(full, mode);
  if (next === full) { setStatus("The case is already like that."); return; }
  const a = root === els.wordView ? textOffsetOf(r.startContainer, r.startOffset) : null;
  const f = root === els.wordView ? textOffsetOf(r.endContainer, r.endOffset) : null;
  let pos = full.length;
  for (let i = segs.length - 1; i >= 0; i--) {
    const s = segs[i];
    const len = s.end - s.start;
    pos -= len;
    const piece = next.slice(pos, pos + len);
    if (piece === s.node.data.slice(s.start, s.end)) continue;
    // Replace in place (formatting of the text node is kept); the editor's
    // undo stack records it when execCommand accepts the edit.
    const rr = document.createRange();
    rr.setStart(s.node, s.start);
    rr.setEnd(s.node, s.end);
    sel.removeAllRanges();
    sel.addRange(rr);
    const parent = s.node.parentNode;
    let ok = false;
    try { ok = document.execCommand("insertText", false, piece); } catch { ok = false; }
    if (!ok || (parent && !parent.isConnected)) { if (s.node.isConnected) s.node.replaceData(s.start, len, piece); }
  }
  if (root === els.wordView && a != null && f != null) {
    const pa = pointAtTextOffset(a, true);
    const pf = pointAtTextOffset(f);
    if (pa && pf) { try { sel.setBaseAndExtent(pa.node, pa.offset, pf.node, pf.offset); } catch { /* keep */ } }
    afterEdit();
    markWordDirty();
  } else if (hfActive) syncBand(hfActive);
}

// --- format painter ----------------------------------------------------------------

let painter = null; // { fmt, sticky }

function captureFormat() {
  restoreSelection();
  const sel = getSelection();
  if (!sel.rangeCount) return null;
  const node = sel.getRangeAt(0).startContainer;
  const el = node.nodeType === 1 ? node : node.parentElement;
  if (!el || !els.wordView.contains(el)) return null;
  const q = (c) => { try { return document.queryCommandState(c); } catch { return false; } };
  const cs = getComputedStyle(el);
  let hl = null;
  for (let e = el; e && e !== els.wordView; e = e.parentElement) {
    const bg = getComputedStyle(e).backgroundColor;
    if (bg && !/^(transparent|rgba\(0, 0, 0, 0\))$/.test(bg)) { hl = WordModel.colorToHex(bg); break; }
    if (/^(P|LI|H[1-6]|TD|TH|PRE|BLOCKQUOTE)$/.test(e.tagName)) break;
  }
  return {
    b: q("bold"), i: q("italic"), u: q("underline"), s: q("strikeThrough"), sup: q("superscript"), sub: q("subscript"),
    color: WordModel.colorToHex(cs.color), hl,
    font: (cs.fontFamily || "").split(",")[0].replace(/["']/g, "").trim(),
    size: Math.round(parseFloat(cs.fontSize) * 0.75 * 2) / 2,
  };
}

function togglePainter(sticky = false) {
  if (!wordEditable() || hfActive) { if (hfActive) toast("The format painter works in the document body."); return; }
  if (painter && sticky) { painter.sticky = true; updatePainterUi(); return; }
  if (painter) { stopPainter(); return; }
  const fmt = captureFormat();
  if (!fmt) { toast("Click into text whose formatting you want to copy first."); return; }
  painter = { fmt, sticky };
  updatePainterUi();
  restoreSelection(true);
}

function stopPainter() {
  if (!painter) return;
  painter = null;
  updatePainterUi();
  setStatus("");
}

function updatePainterUi() {
  const on = !!painter;
  els.fbPainter.setAttribute("aria-pressed", String(on));
  els.fbPainter.classList.toggle("active", on);
  els.wordView.classList.toggle("painting", on);
  if (on) setStatus(painter.sticky ? "Format painter is on: select text to apply the copied formatting. Esc stops it." : "Format painter: select the text to apply the copied formatting to.");
}

function applyPainterIfSelection() {
  if (!painter) return;
  const r = selectionInEditor();
  if (!r || r.collapsed) return;
  savedRange = r.cloneRange();
  const f = painter.fmt;
  const q = (c) => { try { return document.queryCommandState(c); } catch { return false; } };
  exec("removeFormat");
  for (const [k, cmd] of [["b", "bold"], ["i", "italic"], ["u", "underline"], ["s", "strikeThrough"], ["sup", "superscript"], ["sub", "subscript"]]) {
    if (!!f[k] !== q(cmd)) exec(cmd);
  }
  const node = selectionInEditor()?.startContainer;
  const el = node && (node.nodeType === 1 ? node : node.parentElement);
  const cs = el ? getComputedStyle(el) : null;
  if (cs && f.color && WordModel.colorToHex(cs.color) !== f.color) exec("foreColor", f.color);
  if (f.hl) exec("hiliteColor", f.hl);
  if (cs && f.font && (cs.fontFamily || "").split(",")[0].replace(/["']/g, "").trim() !== f.font) exec("fontName", f.font);
  if (cs && f.size > 0 && Math.abs(parseFloat(cs.fontSize) * 0.75 - f.size) > 0.25) applyFontSize(f.size);
  markWordDirty();
  if (!painter.sticky) stopPainter();
}

// --- document properties --------------------------------------------------------------

function openProperties() {
  if (!wordEditable()) return;
  const meta = wordMeta();
  els.propsTitle.value = typeof meta.title === "string" ? meta.title : "";
  els.propsAuthor.value = typeof meta.author === "string" ? meta.author : "";
  els.propsSubject.value = typeof meta.subject === "string" ? meta.subject : "";
  openModal(els.propsModal, { initialFocus: els.propsTitle });
}

function submitProperties() {
  const meta = wordMeta();
  let changed = false;
  for (const [key, el] of [["title", els.propsTitle], ["author", els.propsAuthor], ["subject", els.propsSubject]]) {
    const v = el.value.trim();
    const before = typeof meta[key] === "string" ? meta[key] : "";
    if (v === before) continue;
    changed = true;
    if (v) meta[key] = v;
    else delete meta[key];
  }
  closeModal(els.propsModal);
  if (changed) {
    markWordDirty();
    setStatus("Document properties updated (saved with the document).");
  }
  restoreSelection(true);
}

// --- templates ------------------------------------------------------------------------

let tplSelected = "blank";

function openTemplates() {
  const items = WORD_TEMPLATES.map((t) => {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "tpl-item";
    b.setAttribute("role", "option");
    b.dataset.id = t.id;
    b.setAttribute("aria-selected", String(t.id === tplSelected));
    const thumb = document.createElement("div");
    thumb.className = "tpl-thumb";
    thumb.setAttribute("aria-hidden", "true");
    const page = document.createElement("div");
    page.className = "ow-page";
    page.innerHTML = renderWordHtml(wordTemplate(t.id));
    thumb.appendChild(page);
    const name = document.createElement("span");
    name.className = "tpl-name";
    name.textContent = t.name;
    const desc = document.createElement("span");
    desc.className = "tpl-desc";
    desc.textContent = t.description;
    b.append(thumb, name, desc);
    return b;
  });
  els.tplGrid.replaceChildren(...items);
  openModal(els.tplModal, { initialFocus: els.tplGrid.querySelector('[aria-selected="true"]') || items[0] });
}

function selectTemplate(id) {
  tplSelected = id;
  for (const b of els.tplGrid.children) b.setAttribute("aria-selected", String(b.dataset.id === id));
}

async function createFromTemplate(id = tplSelected) {
  closeModal(els.tplModal);
  await newDocument("word", { template: id });
}

// --- actions & shortcuts ----------------------------------------------------------------

function runWordAction(action) {
  if (!wordEditable()) return;
  switch (action) {
    case "link": if (hfActive) toast("Links go in the document body."); else openLinkDialog(); break;
    case "image": if (hfActive) toast("Pictures go in the document body."); else els.imageInput.click(); break;
    case "symbol": openSymbols(); break;
    case "datetime": openDateTime(); break;
    case "pagebreak": if (hfActive) toast("Page breaks go in the document body."); else insertPageBreak(); break;
    case "edit-header": activateBand(els.wordHeader, { caretAtEnd: true }); break;
    case "edit-footer": activateBand(els.wordFooter, { caretAtEnd: true }); break;
    case "page-numbers": addPageNumbers(); break;
    case "painter": togglePainter(false); break;
    case "clear": exec("removeFormat"); markWordDirty(); restoreWordFocus(); break;
    case "page-setup": openPageSetup(); break;
    case "properties": openProperties(); break;
    case "nav": toggleNav(); break;
    case "find": openFind(false); break;
    case "replace": openFind(true); break;
    case "wordcount": openWordCount(); break;
    case "zoom-in": zoomStep(1); restoreWordFocus(); break;
    case "zoom-out": zoomStep(-1); restoreWordFocus(); break;
    case "zoom-reset": setZoom(100); restoreWordFocus(); break;
    default: break;
  }
}

/** Global Word shortcuts. Returns true when handled. */
function handleWordShortcut(e, mod) {
  const k = e.key.toLowerCase();
  if (!mod) {
    if (e.key === "F3" && !els.findBar.hidden) { gotoMatch(e.shiftKey ? -1 : 1); return true; }
    return false;
  }
  if (k === "f" && !e.shiftKey) { openFind(false); return true; }
  if (k === "h" && !e.shiftKey) { openFind(true); return true; }
  if (k === "f" && e.shiftKey) { toggleNav(); return true; }
  if (k === "g" && e.shiftKey) { openWordCount(); return true; }
  if (!e.shiftKey || k === "+") {
    if (k === "=" || k === "+") { zoomStep(1); return true; }
    if (k === "-" || k === "_") { zoomStep(-1); return true; }
    if (k === "0") { setZoom(100); return true; }
  }
  return false;
}

function resetWordFeatures() {
  closeFind({ restore: false });
  if (hfActive) { hfActive.classList.remove("hf-active"); hfActive = null; }
  hfRange = null;
  els.wordSheet.classList.remove("hf-editing");
  els.hfTools.hidden = true;
  els.wordHeader.innerHTML = "";
  els.wordFooter.innerHTML = "";
  painter = null;
  updatePainterUi();
  navHeadings = [];
  els.navList.replaceChildren();
}

function bindWordFeatures() {
  applyZoom();
  // --- header / footer bands
  for (const band of [els.wordHeader, els.wordFooter]) {
    band.addEventListener("focus", () => { if (hfActive !== band) activateBand(band); });
    band.addEventListener("input", () => { syncBand(band); positionHfTools(); });
    band.addEventListener("keydown", (e) => {
      if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); closeBand(); }
      else if (e.key === "Tab" && !e.ctrlKey && !e.altKey && !e.metaKey) {
        e.preventDefault();
        if (!e.shiftKey) document.execCommand("insertText", false, "\t");
      } else if ((e.ctrlKey || e.metaKey) && (e.key === "Enter" || e.key.toLowerCase() === "k")) e.preventDefault();
    });
    // Paste plain text only (no tables / pictures in a one-paragraph band).
    band.addEventListener("paste", (e) => {
      e.preventDefault();
      e.stopPropagation();
      const t = e.clipboardData ? e.clipboardData.getData("text/plain") : "";
      if (t) document.execCommand("insertText", false, t.replace(/\r\n?/g, "\n"));
    });
    band.addEventListener("drop", (e) => { e.preventDefault(); e.stopPropagation(); });
  }
  document.addEventListener("selectionchange", () => {
    if (!hfActive) return;
    const sel = getSelection();
    if (sel.rangeCount && hfActive.contains(sel.getRangeAt(0).commonAncestorContainer)) hfRange = sel.getRangeAt(0).cloneRange();
  });
  // Clicking into the document body ends header/footer editing.
  els.wordView.addEventListener("focus", () => { if (hfActive) { deactivateBand(); setStatus(""); } });
  els.wordView.addEventListener("mousedown", () => { if (hfActive) deactivateBand(); });
  els.hfTools.addEventListener("mousedown", (e) => e.preventDefault());
  els.hfTools.addEventListener("click", (e) => {
    const b = e.target.closest("button[data-hf]");
    if (!b) return;
    if (b.dataset.hf === "close") closeBand();
    else insertField(b.dataset.hf);
  });

  // --- zoom
  els.zoomIn.addEventListener("click", () => zoomStep(1));
  els.zoomOut.addEventListener("click", () => zoomStep(-1));
  els.zoomLabel.addEventListener("click", () => setZoom(100));
  els.zoomRange.addEventListener("input", () => setZoom(Number(els.zoomRange.value)));
  els.stage.addEventListener("wheel", (e) => {
    if (!(e.ctrlKey || e.metaKey) || !wordEditable()) return;
    e.preventDefault();
    if (e.deltaY) zoomStep(e.deltaY < 0 ? 1 : -1);
  }, { passive: false });

  // --- stats -> word count
  els.stats.addEventListener("mousedown", (e) => e.preventDefault());
  els.stats.addEventListener("click", () => openWordCount());
  els.wcClose.addEventListener("click", () => { closeModal(els.wcModal); restoreSelection(true); });

  // --- find bar
  els.findInput.addEventListener("input", () => runFind());
  for (const b of [els.findCase, els.findWord]) {
    b.addEventListener("click", () => {
      b.setAttribute("aria-pressed", String(b.getAttribute("aria-pressed") !== "true"));
      runFind();
      els.findInput.focus();
    });
  }
  els.findNext.addEventListener("click", () => gotoMatch(1));
  els.findPrev.addEventListener("click", () => gotoMatch(-1));
  els.findClose.addEventListener("click", () => closeFind());
  els.findToggleReplace.addEventListener("click", () => {
    setReplaceVisible(els.findReplaceRow.hidden);
    (els.findReplaceRow.hidden ? els.findInput : els.replaceInput).focus();
  });
  els.replaceOne.addEventListener("click", () => replaceCurrent());
  els.replaceAll.addEventListener("click", () => replaceAllMatches());
  els.findBar.addEventListener("keydown", (e) => {
    if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); closeFind(); return; }
    if (e.key === "Enter" && !e.isComposing) {
      if (e.target === els.findInput) { e.preventDefault(); gotoMatch(e.shiftKey ? -1 : 1); }
      else if (e.target === els.replaceInput) {
        e.preventDefault();
        if (e.ctrlKey || e.metaKey) replaceAllMatches(); else replaceCurrent();
      }
    }
  });

  // --- navigation pane
  els.navClose.addEventListener("click", () => toggleNav(false));
  els.navList.addEventListener("click", (e) => {
    const b = e.target.closest(".nav-item");
    if (b) goToHeading(Number(b.dataset.i));
  });
  els.navList.addEventListener("keydown", (e) => {
    const items = [...els.navList.children];
    const i = items.indexOf(document.activeElement);
    if (e.key === "ArrowDown") { e.preventDefault(); items[Math.min(items.length - 1, i + 1)]?.focus(); }
    else if (e.key === "ArrowUp") { e.preventDefault(); items[Math.max(0, i - 1)]?.focus(); }
    else if (e.key === "Home") { e.preventDefault(); items[0]?.focus(); }
    else if (e.key === "End") { e.preventDefault(); items[items.length - 1]?.focus(); }
    else if (e.key === "Escape") { e.preventDefault(); restoreWordFocus(); }
  });
  els.stage.addEventListener("scroll", () => {
    if (els.navPane.hidden || navRaf) return;
    navRaf = requestAnimationFrame(() => { navRaf = 0; updateNavCurrent(); });
  }, { passive: true });

  // --- symbols
  els.symTabs.addEventListener("click", (e) => {
    const t = e.target.closest(".sym-tab");
    if (t) { showSymbolGroup(Number(t.dataset.i)); els.symGrid.querySelector(".sym-cell")?.focus(); }
  });
  els.symGrid.addEventListener("click", (e) => {
    const c = e.target.closest(".sym-cell");
    if (!c) return;
    insertSymbol(c.dataset.ch);
    c.focus();
  });
  els.symGrid.addEventListener("focusin", (e) => {
    const c = e.target.closest(".sym-cell");
    if (c) { els.symPreview.textContent = c.dataset.ch; els.symName.textContent = symbolLabel(c.dataset.ch); }
  });
  els.symGrid.addEventListener("keydown", (e) => {
    const cells = [...els.symGrid.children];
    const i = cells.indexOf(document.activeElement);
    if (i < 0) return;
    const cols = Math.max(1, cells.filter((c) => c.offsetTop === cells[0].offsetTop).length);
    const move = { ArrowRight: 1, ArrowLeft: -1, ArrowDown: cols, ArrowUp: -cols }[e.key];
    if (move) { e.preventDefault(); cells[Math.max(0, Math.min(cells.length - 1, i + move))].focus(); }
  });
  els.symClose.addEventListener("click", () => { closeModal(els.symModal); restoreWordFocus(); });

  // --- date & time
  els.dtForm.addEventListener("submit", (e) => { e.preventDefault(); submitDateTime(); });
  els.dtList.addEventListener("dblclick", (e) => { if (e.target.closest(".dt-option")) submitDateTime(); });
  els.dtCancel.addEventListener("click", () => { closeModal(els.dtModal); restoreWordFocus(); });

  // --- page setup
  els.psForm.addEventListener("submit", (e) => { e.preventDefault(); submitPageSetup(); });
  els.psCancel.addEventListener("click", () => { closeModal(els.psModal); psState = null; restoreSelection(true); });
  els.psPreset.addEventListener("change", () => {
    const p = MARGIN_PRESETS[els.psPreset.value];
    if (p) setMarginInputs(p, psState.units);
    updatePsPreview();
  });
  for (const el of [els.psTop, els.psBottom, els.psLeft, els.psRight]) {
    el.addEventListener("input", () => { els.psError.hidden = true; syncPreset(); updatePsPreview(); });
  }
  for (const el of [els.psSize, els.psPortrait, els.psLandscape]) el.addEventListener("change", updatePsPreview);
  els.psUnits.addEventListener("change", () => {
    const m = readMarginInputs();
    psState.units = els.psUnits.value;
    lsSet(LS.units, psState.units);
    setMarginInputs(m, psState.units);
  });

  // --- properties
  els.propsForm.addEventListener("submit", (e) => { e.preventDefault(); submitProperties(); });
  els.propsCancel.addEventListener("click", () => { closeModal(els.propsModal); restoreSelection(true); });

  // --- templates
  els.tplGrid.addEventListener("click", (e) => {
    const b = e.target.closest(".tpl-item");
    if (b) selectTemplate(b.dataset.id);
  });
  els.tplGrid.addEventListener("dblclick", (e) => {
    const b = e.target.closest(".tpl-item");
    if (b) createFromTemplate(b.dataset.id);
  });
  els.tplGrid.addEventListener("keydown", (e) => {
    const items = [...els.tplGrid.children];
    const i = items.indexOf(document.activeElement);
    if (i < 0) return;
    const cols = Math.max(1, items.filter((c) => c.offsetTop === items[0].offsetTop).length);
    const move = { ArrowRight: 1, ArrowLeft: -1, ArrowDown: cols, ArrowUp: -cols }[e.key];
    if (move) {
      e.preventDefault();
      const t = items[Math.max(0, Math.min(items.length - 1, i + move))];
      t.focus();
      selectTemplate(t.dataset.id);
    } else if (e.key === "Enter") {
      e.preventDefault();
      createFromTemplate(items[i].dataset.id);
    }
  });
  els.tplGrid.addEventListener("focusin", (e) => {
    const b = e.target.closest(".tpl-item");
    if (b) selectTemplate(b.dataset.id);
  });
  els.tplOk.addEventListener("click", () => createFromTemplate());
  els.tplCancel.addEventListener("click", () => closeModal(els.tplModal));
}

// --- page view (docx-preview of the current model) -------------------------------

async function renderPreviewOnly() {
  showOnly("preview");
  els.wordPreview.innerHTML = "";
  if (doc.previewNode) {
    els.wordPreview.append(...doc.previewNode.childNodes);
    delete doc.previewNode;
  } else {
    await LocalDocx.renderDocxPreview(doc.bytes, els.wordPreview);
  }
}

async function togglePageView() {
  if (!doc || doc.kind !== "word" || doc.previewOnly) return;
  if (pageView) {
    pageView = false;
    els.wordPreview.innerHTML = "";
    showOnly("word");
    updateButtons();
    updateStats();
    restoreSelection();
    return;
  }
  if (!LocalDocx.docxPreviewAvailable()) { toast("Page view is not available (docx-preview not loaded).", "error"); return; }
  setBusy(true, "Laying out pages…");
  try {
    let bytes = null;
    if (typeof LocalDocx.modelToDocx === "function" && typeof JSZip !== "undefined") {
      bytes = await LocalDocx.modelToDocx(currentWordModel());
    } else if (doc.format === "docx" && !dirty && doc.bytes) {
      bytes = doc.bytes;
    }
    if (!bytes) throw new Error("no .docx writer available");
    els.wordPreview.innerHTML = "";
    await LocalDocx.renderDocxPreview(bytes, els.wordPreview);
    pageView = true;
    showOnly("preview");
    updateButtons();
  } catch (e) {
    toast(`Page view failed: ${errMsg(e)}`, "error");
  } finally {
    setBusy(false);
  }
}

// ---------------------------------------------------------------------------
// save
// ---------------------------------------------------------------------------

function flushEdits() {
  if (!doc) return;
  if (doc.kind === "excel") commitExcelEdit();
  if (doc.kind === "text") doc.text = els.textView.value;
  if (doc.kind === "slides" && slidesView && typeof slidesView.commitEdit === "function") {
    try { slidesView.commitEdit(); } catch (e) { console.warn("slides commitEdit failed:", e); }
  }
  if (doc.kind === "latex" && latexView) doc.text = latexView.getText();
}

function sheetCount() {
  return doc && doc.kind === "excel" && doc.model ? doc.model.sheets.length : 1;
}

/** Formats the Save As list may offer right now (never one that will fail). */
function currentSaveTargets() {
  if (!doc) return [];
  if (doc.previewOnly) {
    const list = [];
    if (doc.bytes) list.push({ id: "original", label: `Original file (.${doc.format})`, engines: ["local"], lossy: null, original: true });
    list.push({ id: "html", label: "Web Page (.html)", engines: ["local"], lossy: "A web page export keeps the look but cannot be edited here." });
    list.push({ id: "txt", label: "Plain Text (.txt)", engines: ["local"], lossy: "Plain text keeps only the words." });
    return list;
  }
  return computeSaveTargets(doc.kind, {
    backendRender: backendUp() ? backendRenderFormats(backend.health, doc.kind) : [],
    localDocx: typeof LocalDocx.modelToDocx === "function" && typeof JSZip !== "undefined",
    sheetjs: LocalXlsx.xlsxAvailable(),
    localPptx: typeof LocalPptx.modelToPptx === "function" && typeof JSZip !== "undefined",
    latex: typeof Latex.wordModelToLatex === "function",
    source: doc.format,
    sheets: sheetCount(),
  });
}

/** Notes from the last writer (backend X-Office-Warnings), shown after saving. */
let lastSaveNotes = [];

/** Produce the file for `target` ({id, engines}). Returns a Blob. */
async function serialize(target) {
  lastSaveNotes = [];
  const format = target.id;
  if (target.original) return new Blob([doc.bytes], { type: doc.mime || mimeOf(doc.format) });
  if (doc.previewOnly) {
    if (format === "html") {
      const html = `<!doctype html><html><head><meta charset="utf-8"><title>${WordModel.escapeHtml(doc.name)}</title></head><body>${els.wordPreview.innerHTML}</body></html>`;
      return new Blob([html], { type: "text/html" });
    }
    return new Blob([(els.wordPreview.innerText || "") + "\n"], { type: "text/plain" });
  }
  if (doc.kind === "text") {
    const { bytes, encoding } = encodeText(els.textView.value, { encoding: doc.encoding, bom: doc.bom, eol: doc.eol });
    if (encoding !== doc.encoding) doc.pendingEncoding = encoding;
    return new Blob([bytes], { type: "text/plain" });
  }
  const errors = [];
  for (const engine of target.engines) {
    try {
      if (engine === "backend") {
        const model = doc.kind === "word" ? currentWordModel()
          : doc.kind === "latex" ? { kind: "latex", source: doc.text || "" }
          : doc.model;
        const res = await backend.client.render(model, format, withExtension(doc.name, format));
        if (errors.length) console.info("saved with backend after local failure:", errors);
        lastSaveNotes = res.warnings || [];
        return res.blob;
      }
      const blob = await serializeLocal(format);
      if (errors.length) {
        toast(`The backend could not write this file (${errors[0]}); it was saved with the built-in writer instead.`);
      }
      return blob;
    } catch (e) {
      if (isBackendError(e)) noteBackendFailure(e);
      console.warn(`${engine} writer failed for ${format}:`, e);
      errors.push(errMsg(e));
    }
  }
  throw new Error(errors.join("; ") || `No writer available for .${format}`);
}

async function serializeLocal(format) {
  if (doc.kind === "word") {
    const model = currentWordModel();
    switch (format) {
      case "docx": {
        const bytes = await LocalDocx.modelToDocx(model);
        return new Blob([bytes], { type: mimeOf("docx") });
      }
      case "md": {
        const md = WordModel.wordModelToMarkdown(model);
        const { bytes } = encodeText(md, doc.format === "md" ? { encoding: doc.encoding, bom: doc.bom, eol: doc.eol } : {});
        return new Blob([bytes], { type: "text/markdown" });
      }
      case "html": {
        // Document properties' title wins over the file name.
        const title = (model.meta && typeof model.meta.title === "string" && model.meta.title.trim()) || doc.name.replace(/\.[^.]+$/, "");
        return new Blob([WordModel.wordModelToHtmlDocument(model, title)], { type: "text/html" });
      }
      case "txt":
        return new Blob([wordModelToText(model).replace(/\f/g, "\n") + "\n"], { type: "text/plain" });
      case "tex":
        return new Blob([new TextEncoder().encode(Latex.wordModelToLatex(model))], { type: mimeOf("tex") });
      default:
        throw new Error(`.${format} needs the backend`);
    }
  }
  if (doc.kind === "slides") {
    const model = doc.model;
    const title = doc.name.replace(/\.[^.]+$/, "");
    switch (format) {
      case "pptx":
        return new Blob([await LocalPptx.modelToPptx(model)], { type: mimeOf("pptx") });
      case "html":
        return new Blob([SlidesModel.slidesToHtmlDocument(model, title)], { type: "text/html" });
      case "md":
        return new Blob([new TextEncoder().encode(SlidesModel.slidesToMarkdown(model))], { type: "text/markdown" });
      case "txt":
        return new Blob([new TextEncoder().encode(SlidesModel.slidesToText(model))], { type: "text/plain" });
      default:
        throw new Error(`.${format} needs the backend`);
    }
  }
  if (doc.kind === "latex") {
    const source = doc.text || "";
    const title = doc.name.replace(/\.[^.]+$/, "");
    switch (format) {
      case "tex": {
        const { bytes, encoding } = encodeText(source, { encoding: doc.encoding, bom: doc.bom, eol: doc.eol });
        if (encoding !== doc.encoding) doc.pendingEncoding = encoding;
        return new Blob([bytes], { type: mimeOf("tex") });
      }
      case "html":
        return new Blob([Latex.latexToHtmlDocument(source, title)], { type: "text/html" });
      case "docx":
        return new Blob([await LocalDocx.modelToDocx(Latex.latexToWordModel(source))], { type: mimeOf("docx") });
      case "md":
        return new Blob([new TextEncoder().encode(WordModel.wordModelToMarkdown(Latex.latexToWordModel(source)))], { type: "text/markdown" });
      case "txt":
        return new Blob([wordModelToText(Latex.latexToWordModel(source)).replace(/\f/g, "\n") + "\n"], { type: "text/plain" });
      default:
        throw new Error(`.${format} needs the backend`);
    }
  }
  if (doc.kind === "excel") {
    const model = doc.model;
    if (format === "csv" || format === "tsv") {
      let bytes;
      if (typeof ExcelModel.modelSheetToCsvBytes === "function") {
        bytes = ExcelModel.modelSheetToCsvBytes(model, model.active || 0, format);
      } else {
        bytes = new TextEncoder().encode(ExcelModel.modelSheetToCsv(model, model.active || 0, format === "tsv" ? "\t" : ","));
      }
      return new Blob([bytes], { type: mimeOf(format) });
    }
    if (!LocalXlsx.xlsxAvailable()) throw new Error("the built-in spreadsheet writer is not loaded");
    const bytes = LocalXlsx.modelToBytes(model, format);
    return new Blob([bytes], { type: mimeOf(format) });
  }
  throw new Error("nothing to save");
}

function safeDownloadName(name) {
  return String(name || "document").replace(/[\\/:*?"<>|\u0000-\u001f]/g, "_").replace(/^\.+/, "").trim() || "document";
}

/**
 * Download fallback (no File System Access). Uses chrome.downloads so the
 * result is known (completed / cancelled); `overwrite` replaces the previous
 * copy on repeated Ctrl+S instead of piling up "name (1).docx".
 */
async function downloadBlob(blob, filename, { overwrite = false } = {}) {
  const url = URL.createObjectURL(blob);
  const cleanup = () => setTimeout(() => URL.revokeObjectURL(url), 60000);
  if (typeof chrome !== "undefined" && chrome.downloads?.download) {
    try {
      const id = await chrome.downloads.download({
        url, filename: safeDownloadName(filename), conflictAction: overwrite ? "overwrite" : "uniquify", saveAs: false,
      });
      await new Promise((resolve, reject) => {
        const onChanged = (delta) => {
          if (delta.id !== id) return;
          if (delta.state && delta.state.current === "complete") { done(); resolve(); }
          else if (delta.state && delta.state.current === "interrupted") {
            done();
            const err = new Error(delta.error && delta.error.current === "USER_CANCELED" ? "cancelled" : `download failed (${(delta.error && delta.error.current) || "interrupted"})`);
            err.cancelled = delta.error && delta.error.current === "USER_CANCELED";
            reject(err);
          }
        };
        const timer = setTimeout(() => { done(); resolve(); }, 30000);
        const done = () => { clearTimeout(timer); chrome.downloads.onChanged.removeListener(onChanged); };
        chrome.downloads.onChanged.addListener(onChanged);
        chrome.downloads.search({ id }).then((items) => {
          const it = items && items[0];
          if (it && it.state === "complete") { done(); resolve(); }
        }).catch(() => {});
      });
      cleanup();
      return;
    } catch (e) {
      if (e && e.cancelled) { cleanup(); throw e; }
      console.warn("chrome.downloads failed, using a link download:", e);
    }
  }
  const a = document.createElement("a");
  a.href = url;
  a.download = safeDownloadName(filename);
  document.body.appendChild(a);
  a.click();
  a.remove();
  cleanup();
}

async function writeToHandle(handle, blob) {
  if (handle.queryPermission) {
    let perm = await handle.queryPermission({ mode: "readwrite" });
    if (perm !== "granted" && handle.requestPermission) perm = await handle.requestPermission({ mode: "readwrite" });
    if (perm !== "granted") {
      const err = new Error("permission to write the file was not granted");
      err.permission = true;
      throw err;
    }
  }
  const w = await handle.createWritable();
  try {
    await w.write(blob);
    await w.close();
  } catch (e) {
    try { await w.abort(); } catch { /* ignore */ }
    throw e;
  }
}

/**
 * Write the document. target: an entry from currentSaveTargets().
 * dest: { name, handle?, overwrite? }. Returns true when written.
 */
async function writeDocument(target, dest) {
  const lossy = !!target.lossy;
  setBusy(true, `Saving ${dest.name}…`);
  setStatus(`Saving ${dest.name}…`);
  try {
    const blob = await serialize(target);
    const notes = lastSaveNotes;
    if (dest.handle) await writeToHandle(dest.handle, blob);
    else await downloadBlob(blob, dest.name, { overwrite: !!dest.overwrite });
    const savedName = dest.handle ? dest.handle.name : dest.name;
    const showNotes = () => {
      if (notes.length) showInfo({ type: "warn", text: `Saved ${savedName} with ${notes.length === 1 ? "a note" : `${notes.length} notes`}:`, items: notes });
    };
    if (lossy || target.original) {
      setStatus(`Exported ${savedName}${dirty ? " — the document still has unsaved changes" : ""}`);
      toast(`Exported ${savedName}`, "ok");
      showNotes();
      return true;
    }
    // A full-fidelity save: this file is now the document.
    if (target.id !== doc.format) hideInfo(); // import notes described the old file
    doc.name = savedName;
    doc.format = target.id;
    doc.handle = dest.handle || null;
    doc.saveTarget = { format: target.id, name: savedName, handle: dest.handle || null };
    doc.isNew = false;
    let note = "";
    if ((doc.kind === "text" || doc.kind === "latex") && doc.pendingEncoding) {
      note = ` as ${encodingLabel(doc.pendingEncoding)} (the text now contains characters ${encodingLabel(doc.encoding)} cannot store)`;
      doc.encoding = doc.pendingEncoding;
      doc.bom = true;
      delete doc.pendingEncoding;
      hideInfo();
    }
    const bytes = new Uint8Array(await blob.arrayBuffer());
    doc.bytes = bytes;
    doc.mime = mimeOf(target.id);
    els.kind.textContent = KIND_LABEL[doc.kind] + ` · ${doc.format.toUpperCase()}`;
    setDirty(false);
    updateTitle();
    updateButtons();
    setStatus(`Saved ${savedName}${note}${dest.handle ? "" : " to your Downloads folder"}`);
    toast(`Saved ${savedName}${note}`, "ok");
    showNotes();
    if (bytes.length <= RECENT_MAX_BYTES) rememberRecent(savedName, bytes, doc.mime, { saved: true });
    return true;
  } catch (e) {
    if (e && (e.cancelled || e.name === "AbortError")) {
      setStatus("Save cancelled");
      return false;
    }
    console.error(e);
    const msg = errMsg(e);
    setStatus(`Save failed: ${msg}`, true);
    toast(`Save failed: ${msg}`, "error");
    return false;
  } finally {
    setBusy(false);
  }
}

/**
 * Ctrl+S. Writes back in place when the document came from (or was saved
 * to) a file handle in a writable format; otherwise behaves as Save As.
 * Resolves true when the document was saved.
 */
async function saveDefault() {
  if (!doc) return false;
  flushEdits();
  if (doc.previewOnly) {
    return openSaveAs({ reason: "This document is read-only here. You can download the original file or export a copy." });
  }
  const targets = currentSaveTargets();
  const st = doc.saveTarget;
  if (st) {
    const t = targets.find((x) => x.id === st.format);
    if (t) return writeDocument(t, { name: st.name, handle: st.handle, overwrite: !st.handle });
    return openSaveAs({ reason: `.${st.format} cannot be written right now (the backend is offline). Choose another format.` });
  }
  const own = targets.find((x) => x.id === doc.format && !x.lossy);
  if (own && doc.handle) {
    const ok = await writeDocument(own, { name: doc.name, handle: doc.handle });
    if (ok) return true;
    // e.g. permission denied: fall back to Save As
    if (!doc.saveTarget) return openSaveAs({ reason: "The original file could not be written. Save a copy instead." });
    return false;
  }
  if (!own) {
    const label = (FORMATS[doc.format] && FORMATS[doc.format].label) || doc.format.toUpperCase();
    return openSaveAs({
      reason: doc.format === "xlsm"
        ? "Macro-enabled workbooks are saved without their macros. Choose a format for the copy."
        : `${label} files cannot be written${["rtf"].includes(doc.format) ? " without the backend" : ""}. Choose a format to save in.`,
    });
  }
  return openSaveAs({});
}

/** Default Save As format for the current document. */
function defaultSaveFormat(targets) {
  const ids = targets.map((t) => t.id);
  const prefer = [];
  if (doc.saveTarget) prefer.push(doc.saveTarget.format);
  prefer.push(doc.format);
  if (doc.previewOnly) prefer.unshift("original");
  if (doc.kind === "word") prefer.push("docx");
  if (doc.kind === "excel") prefer.push("xlsx");
  if (doc.kind === "slides") prefer.push("pptx");
  if (doc.kind === "latex") prefer.push("tex");
  return prefer.find((f) => ids.includes(f)) || ids[0];
}

let saveAsResolve = null;

/**
 * Name to suggest for a never-saved document: its first line of text (like
 * Word does), trimmed and made safe for a file name. Null when not applicable.
 */
function suggestedNewName() {
  if (!doc || !doc.isNew || doc.saveTarget) return null;
  if (doc.kind === "slides") {
    const first = doc.model && doc.model.slides && doc.model.slides[0];
    let name = first && typeof SlidesModel.slideTitle === "function" ? SlidesModel.slideTitle(first) : "";
    name = String(name || "").replace(/[\\/:*?"<>|\u0000-\u001f]/g, "").replace(/[. ]+$/, "").trim();
    if (name.length > 60) name = name.slice(0, 60).replace(/\s+\S*$/, "").trim() || name.slice(0, 60);
    return name ? `${name}.pptx` : null;
  }
  if (doc.kind !== "word") return null;
  const model = currentWordModel();
  for (const b of (model && model.blocks) || []) {
    const text = (b.runs || []).map((r) => r.t || "").join("").replace(/\s+/g, " ").trim();
    if (!text) continue;
    let name = text.replace(/[\\/:*?"<>|\u0000-\u001f]/g, "").replace(/[. ]+$/, "").trim();
    if (name.length > 60) name = name.slice(0, 60).replace(/\s+\S*$/, "").trim() || name.slice(0, 60);
    return name ? `${name}.docx` : null;
  }
  return null;
}

/**
 * Save As dialog. Resolves true when something was written.
 * opts.reason: explanation shown on top (why Save became Save As).
 */
function openSaveAs({ reason = "" } = {}) {
  if (!doc) return Promise.resolve(false);
  flushEdits();
  const targets = currentSaveTargets();
  if (!targets.length) { toast("No format can be written right now.", "error"); return Promise.resolve(false); }
  els.saFormat.innerHTML = "";
  for (const t of targets) {
    const o = document.createElement("option");
    o.value = t.id;
    o.textContent = t.label + (t.lossy ? " — export" : "");
    els.saFormat.appendChild(o);
  }
  const def = defaultSaveFormat(targets);
  els.saFormat.value = def;
  els.saName.value = def === "original" ? doc.name : withExtension((doc.saveTarget && doc.saveTarget.name) || suggestedNewName() || doc.name, def);
  els.saReason.textContent = reason;
  els.saReason.hidden = !reason;
  els.saWhere.textContent = typeof window.showSaveFilePicker === "function"
    ? "Next you will choose where to save the file."
    : "The file will be saved to your Downloads folder.";
  updateSaveAsLossy();
  return new Promise((resolve) => {
    saveAsResolve = resolve;
    openModal(els.saveasModal, {
      initialFocus: els.saName,
      onCancel: () => finishSaveAs(false),
    });
    // select the base name only, like native dialogs
    setTimeout(() => {
      const v = els.saName.value;
      const dot = v.lastIndexOf(".");
      els.saName.setSelectionRange(0, dot > 0 ? dot : v.length);
    }, 0);
  });
}

function finishSaveAs(result) {
  closeModal(els.saveasModal);
  const r = saveAsResolve;
  saveAsResolve = null;
  if (r) r(result);
}

function updateSaveAsLossy() {
  const t = currentSaveTargets().find((x) => x.id === els.saFormat.value);
  const msg = t && t.lossy ? `${t.lossy} Your document stays open with its unsaved changes.` : "";
  els.saLossy.textContent = msg;
  els.saLossy.hidden = !msg;
  // A lossy format is an export: make the confirm button say so.
  $("sa-ok").textContent = t && (t.lossy || t.original) ? (t.original ? "Download" : "Export") : "Save";
  els.saveasModal.querySelector("h3").textContent = t && t.lossy ? "Export a copy" : "Save As";
}

async function submitSaveAs() {
  const format = els.saFormat.value;
  const target = currentSaveTargets().find((x) => x.id === format);
  let name = els.saName.value.trim();
  if (!target) { finishSaveAs(false); return; }
  if (!name) { els.saName.focus(); return; }
  if (!target.original) name = withExtension(name, format);
  // Take the picker while the click still counts as a user gesture.
  closeModal(els.saveasModal);
  let handle = null;
  if (typeof window.showSaveFilePicker === "function") {
    try {
      const opts = { suggestedName: name };
      const pf = target.original ? doc.format : format;
      if (FORMATS[pf] || pf === "html") opts.types = [pickerTypeFor(pf)];
      handle = await window.showSaveFilePicker(opts);
    } catch (e) {
      if (e && e.name === "AbortError") { finishSaveAs(false); return; }
      console.warn("showSaveFilePicker failed, downloading instead:", e);
      handle = null;
    }
  }
  const ok = await writeDocument(target, { name: handle ? handle.name : name, handle });
  finishSaveAs(ok);
}

function downloadOriginal() {
  if (!doc || !doc.bytes || doc.isNew) return;
  downloadBlob(new Blob([doc.bytes], { type: doc.mime || mimeOf(doc.format) }), doc.name)
    .then(() => toast(`Downloaded ${doc.name}`, "ok"))
    .catch((e) => { if (!e.cancelled) toast(`Download failed: ${errMsg(e)}`, "error"); });
}

// ---------------------------------------------------------------------------
// entry points: picker, dnd, paste, url, handoff, recents, launchQueue
// ---------------------------------------------------------------------------

async function pickAndOpen() {
  if (typeof window.showOpenFilePicker === "function") {
    let handles;
    try {
      handles = await window.showOpenFilePicker({
        multiple: true,
        types: [
          { description: "Documents and spreadsheets", accept: { "application/octet-stream": OPEN_EXTENSIONS } },
        ],
        excludeAcceptAllOption: false,
      });
    } catch (e) {
      if (e && e.name === "AbortError") return;
      console.warn("showOpenFilePicker failed, using <input type=file>:", e);
      els.fileInput.click();
      return;
    }
    await openSources(handles.map((handle) => ({ handle })));
    return;
  }
  els.fileInput.click();
}

function fileLooksOpenable(file) {
  if (formatOfName(file.name)) return true;
  const t = mimeEssence(file.type);
  return /word|excel|spreadsheet|opendocument|csv|rtf|markdown|^text\/plain$/.test(t);
}

async function promptOpenUrl() {
  els.urlError.hidden = true;
  els.urlInput.value = "";
  openModal(els.urlModal, { initialFocus: els.urlInput });
}

async function submitOpenUrl() {
  const raw = els.urlInput.value.trim();
  let url;
  try {
    url = new URL(/^[a-z][a-z0-9+.-]*:/i.test(raw) ? raw : `https://${raw}`);
  } catch {
    url = null;
  }
  if (!url || !/^https?:$/.test(url.protocol) || !url.hostname) {
    els.urlError.textContent = url && url.protocol === "file:"
      ? "Local files cannot be opened by address. Use Open or drag the file here instead."
      : "Enter a web address that starts with http:// or https://";
    els.urlError.hidden = false;
    els.urlInput.focus();
    return;
  }
  closeModal(els.urlModal);
  await openFromUrl(url.href);
}

async function openFromUrl(url, { skipGuard = false } = {}) {
  if (!skipGuard) {
    const ok = await confirmDiscard({ action: "opening the link" });
    if (ok !== true) return false;
  }
  setBusy(true, "Downloading…");
  setStatus(`Downloading ${url}…`);
  let src;
  try {
    if (/^file:/i.test(url)) {
      const bytes = await readFileUrl(url);
      let name = "";
      try { name = decodeURIComponent(new URL(url).pathname.split("/").pop() || ""); } catch { /* keep "" */ }
      src = { name: name || "document", bytes, mime: "", url, origin: "url" };
      setBusy(false);
      return openSource(src, { skipGuard: true });
    }
    const res = await fetch(url, { credentials: "include" });
    if (!res.ok) throw new Error(`the server answered HTTP ${res.status}${res.status === 401 || res.status === 403 ? " (sign-in required?)" : ""}`);
    const ct = res.headers.get("Content-Type") || "";
    let name = filenameFromContentDisposition(res.headers.get("Content-Disposition"));
    if (!name) {
      try { name = decodeURIComponent(new URL(res.url || url).pathname.split("/").pop() || ""); } catch { name = new URL(url).pathname.split("/").pop() || ""; }
    }
    const bytes = new Uint8Array(await res.arrayBuffer());
    src = { name: name || "document", bytes, mime: ct, url, origin: "url" };
    // A login/redirect page served instead of the file.
    if (looksLikeHtmlPage(bytes)) {
      setBusy(false);
      setStatus(`Could not open ${url}: the link returned a web page, not a document`, true);
      showInfo({
        type: "error",
        text: "This link returned a web page instead of a document — the site may require you to sign in first.",
        actions: [{ label: "Open link in a normal tab", primary: true, onClick: () => openTab(url) }],
      });
      if (!doc) showEmpty();
      return false;
    }
  } catch (e) {
    setBusy(false);
    setStatus(`Could not download ${url}: ${errMsg(e)}`, true);
    const actions = [{ label: "Open link in a normal tab", onClick: () => openTab(url) }];
    // Website access is an optional permission: without it a cross-origin
    // fetch fails like a network error. Offer to grant it for this site.
    const origins = await missingSiteAccess(url);
    if (origins) {
      actions.unshift({
        label: `Allow access to ${new URL(url).host}`,
        primary: true,
        onClick: async () => {
          let granted = false;
          try { granted = await chrome.permissions.request({ origins }); } catch (err) { console.warn(err); }
          if (granted) openFromUrl(url, { skipGuard: true });
          else toast("Access was not granted.");
        },
      });
    }
    showInfo({
      type: "error",
      text: origins
        ? `Could not download the file: the extension does not have access to ${new URL(url).host} yet.`
        : `Could not download the file: ${errMsg(e)}.`,
      actions,
    });
    if (!doc) showEmpty();
    return false;
  }
  setBusy(false);
  const ok = await openSource(src, { skipGuard: true });
  if (ok && doc && !doc.recentId) setUrlParams({ url });
  return ok;
}

/** Origins to request when `url` is http(s) and the extension lacks host access to it; else null. */
async function missingSiteAccess(url) {
  try {
    const u = new URL(url);
    if (!/^https?:$/.test(u.protocol) || typeof chrome === "undefined" || !chrome.permissions?.contains) return null;
    const origins = [`${u.protocol}//${u.hostname}/*`];
    return (await chrome.permissions.contains({ origins })) ? null : origins;
  } catch {
    return null;
  }
}

/** Read a file: URL (fetch() does not support file:, XHR does when file access is allowed). */
function readFileUrl(url) {
  return new Promise((resolve, reject) => {
    const xhr = new XMLHttpRequest();
    xhr.open("GET", url);
    xhr.responseType = "arraybuffer";
    xhr.onload = () => (xhr.response ? resolve(new Uint8Array(xhr.response)) : reject(new Error("empty file")));
    xhr.onerror = () => reject(new Error("the file could not be read (allow access to file URLs in the extension's details page)"));
    xhr.send();
  });
}

/** An HTML document (login page, error page) rather than a file. */
function looksLikeHtmlPage(bytes) {
  const head = new TextDecoder().decode(bytes.subarray(0, 1024)).replace(/^﻿/, "");
  return /^\s*(<!--[\s\S]*?-->\s*)*<(!doctype\s+html|html|head|body)\b/i.test(head);
}

function bindEntryPoints() {
  els.btnOpen.addEventListener("click", () => pickAndOpen());
  els.btnOpen2.addEventListener("click", () => pickAndOpen());
  els.btnOpenUrl.addEventListener("click", () => promptOpenUrl());
  els.emptyNewWord.addEventListener("click", () => newDocument("word"));
  els.emptyNewExcel.addEventListener("click", () => newDocument("excel"));
  els.emptyNewSlides.addEventListener("click", () => newDocument("slides"));
  els.emptyNewLatex.addEventListener("click", () => newDocument("latex"));
  els.emptyNewTemplate.addEventListener("click", () => openTemplates());
  els.fileInput.addEventListener("change", async () => {
    const files = [...(els.fileInput.files || [])];
    els.fileInput.value = "";
    if (files.length) await openSources(files.map((file) => ({ file })));
  });
  els.btnSave.addEventListener("click", () => saveDefault());
  els.btnSaveAs.addEventListener("click", () => openSaveAs({}));
  els.btnPreview.addEventListener("click", () => togglePageView());
  els.btnPrint.addEventListener("click", () => printDocument());
  els.infoClose.addEventListener("click", hideInfo);
  els.textView.addEventListener("input", () => { setDirty(true); scheduleStats(); });

  // Save As form
  els.saForm.addEventListener("submit", (e) => { e.preventDefault(); submitSaveAs(); });
  els.saCancel.addEventListener("click", () => finishSaveAs(false));
  els.saFormat.addEventListener("change", () => {
    const f = els.saFormat.value;
    if (f !== "original") els.saName.value = withExtension(els.saName.value || doc.name, f);
    updateSaveAsLossy();
  });
  // Link form
  els.lnForm.addEventListener("submit", (e) => { e.preventDefault(); applyLink(); });
  els.lnCancel.addEventListener("click", () => { closeModal(els.linkModal); restoreSelection(true); });
  els.lnRemove.addEventListener("click", removeLink);
  els.lnUrl.addEventListener("input", () => { els.lnError.hidden = true; });
  // URL form
  els.urlForm.addEventListener("submit", (e) => { e.preventDefault(); submitOpenUrl(); });
  els.urlCancel.addEventListener("click", () => closeModal(els.urlModal));
  els.ksClose.addEventListener("click", () => closeModal(els.shortcutsModal));

  // Drag & drop: only FILE drags are intercepted, so text drag-and-drop
  // inside the editor, the text view and the grid keeps working.
  const isFileDrag = (e) => !!e.dataTransfer && [...(e.dataTransfer.types || [])].includes("Files");
  let dragDepth = 0;
  window.addEventListener("dragenter", (e) => {
    if (!isFileDrag(e) || anyModalOpen()) return;
    dragDepth++;
    if (!(doc && doc.kind === "word" && els.wordView.contains(e.target))) els.dropHint.hidden = false;
  });
  window.addEventListener("dragleave", (e) => {
    if (!isFileDrag(e)) return;
    if (--dragDepth <= 0) { dragDepth = 0; els.dropHint.hidden = true; }
  });
  window.addEventListener("dragover", (e) => {
    if (!isFileDrag(e)) return;
    e.preventDefault();
    e.dataTransfer.dropEffect = "copy";
    // Hide the overlay while hovering the Word page (images drop into the text there).
    if (doc && doc.kind === "word" && !pageView) els.dropHint.hidden = els.wordView.contains(e.target);
  });
  window.addEventListener("drop", async (e) => {
    dragDepth = 0;
    els.dropHint.hidden = true;
    if (!isFileDrag(e) || anyModalOpen()) return;
    e.preventDefault();
    const items = [...(e.dataTransfer.items || [])].filter((it) => it.kind === "file");
    const files = [...(e.dataTransfer.files || [])];
    // Prefer FileSystemHandles (writable in place) when the browser offers them.
    let sources = [];
    if (items.length && typeof items[0].getAsFileSystemHandle === "function") {
      const handles = await Promise.all(items.map((it) => it.getAsFileSystemHandle().catch(() => null)));
      sources = handles.map((h, i) => (h && h.kind === "file" ? { handle: h, file: files[i] } : { file: files[i] }));
    } else {
      sources = files.map((file) => ({ file }));
    }
    const openable = sources.filter((s) => s.file && fileLooksOpenable(s.file));
    if (!openable.length) {
      if (files.length) toast("That file type is not supported. Drop a Word, Excel, CSV, Markdown or text file.", "error");
      return;
    }
    if (openable.length < sources.length) toast(`${sources.length - openable.length} unsupported file(s) skipped`);
    await openSources(openable, { action: "opening the dropped file" });
  });

  // Paste files (from the OS clipboard) to open them.
  window.addEventListener("paste", async (e) => {
    const files = [...(e.clipboardData?.files || [])].filter(fileLooksOpenable);
    if (!files.length) return;
    e.preventDefault();
    await openSources(files.map((file) => ({ file })), { action: "opening the pasted file" });
  });

  window.addEventListener("keydown", onGlobalKeydown);
  window.addEventListener("beforeunload", (e) => {
    if (dirty) { e.preventDefault(); e.returnValue = ""; }
  });
  window.addEventListener("beforeprint", fillPrintText);
}

function fillPrintText() {
  updateWordPrintStyle();
  if (doc && doc.kind === "text") els.printText.textContent = els.textView.value;
  if (doc && doc.kind === "excel") commitExcelEdit();
  const view = doc && doc.kind === "slides" ? slidesView : doc && doc.kind === "latex" ? latexView : null;
  if (view && typeof view.beforePrint === "function") {
    try { view.beforePrint(); } catch (e) { console.warn("beforePrint failed:", e); }
  }
}

function printDocument() {
  fillPrintText();
  window.print();
}

function isTypingTarget(t) {
  if (!t || t.nodeType !== 1) return false;
  return t.isContentEditable || /^(INPUT|TEXTAREA|SELECT)$/.test(t.tagName);
}

function onGlobalKeydown(e) {
  if (anyModalOpen()) return;
  if (e.key === "Escape") {
    if (!els.menuNew.hidden || !els.menuMore.hidden || !els.tablePicker.hidden) { closeMenus(); e.preventDefault(); return; }
    if (painter) { stopPainter(); e.preventDefault(); return; }
    if (!els.findBar.hidden) { closeFind(); e.preventDefault(); return; }
    if (pageView) { togglePageView(); e.preventDefault(); }
    return;
  }
  const mod = e.ctrlKey || e.metaKey;
  // Word tools (only while a Word document is being edited; other views
  // keep their own Ctrl+F etc.).
  if (wordEditable() && !e.altKey && handleWordShortcut(e, mod)) {
    e.preventDefault();
    return;
  }
  if (!mod) {
    // "?" outside text fields (and outside the grid, where typing edits a cell)
    if (e.key === "F1" || (e.key === "?" && !isTypingTarget(e.target) && !els.excelView.contains(e.target)
      && !els.slidesView.contains(e.target) && !els.latexView.contains(e.target))) {
      e.preventDefault();
      showShortcuts();
    }
    return;
  }
  if (e.altKey) return;
  const k = e.key.toLowerCase();
  if (k === "s") {
    e.preventDefault();
    if (!doc) return;
    if (e.shiftKey) openSaveAs({});
    else saveDefault();
  } else if (k === "o" && !e.shiftKey) {
    e.preventDefault();
    pickAndOpen();
  } else if (k === "p" && !e.shiftKey) {
    e.preventDefault();
    printDocument();
  } else if (k === "/" ) {
    e.preventDefault();
    showShortcuts();
  }
}

async function consumeParams() {
  const q = new URLSearchParams(location.search);
  const handoff = q.get("handoff");
  const url = q.get("url");
  const recent = q.get("recent");
  const nw = q.get("new");
  if (handoff) {
    let keep = false;
    try {
      const h = await getHandoff(handoff);
      // Files too large for recents keep their handoff (the sweep removes it
      // after HANDOFF_MAX_AGE_MS), so reloading this tab still works for a while.
      keep = !!(h.bytes && h.bytes.length > RECENT_MAX_BYTES);
      const ok = await openSource({ name: h.name, bytes: h.bytes, mime: h.mime, origin: "handoff" }, { skipGuard: true });
      if (ok && keep) setUrlParams({ handoff });
      else keep = false;
    } catch (e) {
      setStatus(`Could not open the file: ${errMsg(e)}`, true);
      showInfo({ type: "error", text: `The file could not be handed to the editor (${errMsg(e)}). Open it again from the extension.` });
      setUrlParams({});
    } finally {
      if (!keep) deleteHandoff(handoff).catch(() => {});
    }
  } else if (url) {
    await openFromUrl(url, { skipGuard: true });
  } else if (recent) {
    try {
      const r = await getRecent(recent);
      await openSource({ name: r.name, bytes: r.bytes, mime: r.mime, origin: "recent" }, { skipGuard: true });
    } catch (e) {
      setStatus(`That recent file is no longer available (${errMsg(e)}).`, true);
      setUrlParams({});
    }
  } else if (NEW_KINDS.includes(nw)) {
    await newDocument(nw, { template: q.get("template") || undefined });
  } else if (q.get("open")) {
    // From the popup's "Open file…": a fresh tab has no user activation, so the
    // picker cannot be shown automatically. Point at the Open button instead.
    setUrlParams({});
    showEmpty();
    els.btnOpen2.classList.add("attention");
    els.btnOpen2.textContent = "Choose a file to open…";
    els.btnOpen2.addEventListener("click", () => {
      els.btnOpen2.classList.remove("attention");
      els.btnOpen2.textContent = "Open a file…";
    }, { once: true });
    setStatus("Click “Choose a file to open…” or press Ctrl+O.");
    setTimeout(() => els.btnOpen2.focus(), 0);
  }
}

function bindLaunchQueue() {
  // OS "Open with" (file_handlers): handles are kept, so Save writes back in place.
  if (!("launchQueue" in window)) return;
  window.launchQueue.setConsumer(async (params) => {
    const handles = [...(params.files || [])].filter((h) => h && h.kind === "file");
    if (handles.length) await openSources(handles.map((handle) => ({ handle })), { action: "opening the file" });
  });
}

// ---------------------------------------------------------------------------
// test hook (headless UI tests drive the viewer through this)
// ---------------------------------------------------------------------------

window.__officeViewer = {
  get doc() { return doc; },
  get dirty() { return dirty; },
  get backend() { return { ok: backendUp(), health: backend.health }; },
  get excelView() { return excelView; },
  openBytes: (name, bytes, mime = "", opts = {}) =>
    openSource({ name, bytes: bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes), mime, handle: opts.handle || null }, opts),
  newDocument,
  saveDefault,
  openSaveAs,
  currentSaveTargets,
  currentWordModel,
  checkBackend,
  localCapabilities,
  // word features (tests)
  get zoom() { return wordZoom; },
  get findState() { return { open: !els.findBar.hidden, count: find.matches.length, index: find.index }; },
  estimatePages,
  wordStats,
};

// ---------------------------------------------------------------------------
// init
// ---------------------------------------------------------------------------

async function init() {
  backend.settings = await loadSettings();
  bindMenus();
  bindEntryPoints();
  bindWordEditor();
  bindWordBar();
  bindWordFeatures();
  bindLaunchQueue();
  startHealthPolling();
  renderBackendStatus(true);
  updateButtons();
  updateTitle();
  await checkBackend();
  await consumeParams();
  if (!doc) showEmpty();
  // Clean up stale handoffs later (never while other tabs may still be
  // writing theirs during startup).
  setTimeout(() => { sweepHandoffs().catch(() => {}); }, 60000);
}

init();
