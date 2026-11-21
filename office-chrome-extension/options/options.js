/**
 * Options page. Every control saves immediately into chrome.storage.local
 * ("settings"); the service worker, content scripts and open editor tabs
 * listen to storage.onChanged, so changes apply live.
 */

import { loadSettings, saveSettings, createBackendClient, DEFAULT_SETTINGS } from "../lib/backend.js";
import { clearRecents } from "../lib/handoff.js";

const P = globalThis.OfficeExtPolicy;
const $ = (id) => document.getElementById(id);
const els = {
  saveStatus: $("save-status"),
  enabled: $("backend-enabled"),
  url: $("backend-url"),
  urlError: $("url-error"),
  urlWarning: $("url-warning"),
  test: $("btn-test"),
  testResult: $("test-result"),
  howto: $("howto"),
  permStatus: $("perm-status"),
  btnPerm: $("btn-perm"),
  intercept: $("intercept-downloads"),
  interceptText: $("intercept-text"),
  interceptDrops: $("intercept-drops"),
  fileStatus: $("file-status"),
  btnFileAccess: $("btn-file-access"),
  btnClearRecents: $("btn-clear-recents"),
};

let statusTimer = 0;
function flash(msg, isError = false) {
  els.saveStatus.textContent = msg;
  els.saveStatus.className = `save-status${isError ? " error" : ""}`;
  clearTimeout(statusTimer);
  statusTimer = setTimeout(() => { els.saveStatus.textContent = ""; }, isError ? 6000 : 2500);
}

/** Merge `patch` into the stored settings (never drops keys other pages own). */
async function patchSettings(patch) {
  const current = await loadSettings();
  await saveSettings({ ...current, ...patch });
  flash("Saved ✓");
}

// ------------------------------------------------------------------ backend

function showUrlProblem(v) {
  els.urlError.hidden = v.ok;
  els.urlError.textContent = v.ok ? "" : v.error;
  els.url.setAttribute("aria-invalid", v.ok ? "false" : "true");
  els.urlWarning.hidden = !v.ok || v.loopback;
}

async function saveUrl() {
  const v = P.validateBackendUrl(els.url.value);
  showUrlProblem(v);
  if (!v.ok) return false;
  const current = await loadSettings();
  if (current.backend_url !== v.url) {
    els.url.value = v.url;
    await patchSettings({ backend_url: v.url });
    runTest(); // tell the user straight away whether it answers
  }
  return true;
}

function listText(arr) {
  return Array.isArray(arr) && arr.length ? arr.join(", ") : "—";
}

async function runTest() {
  const v = P.validateBackendUrl(els.url.value);
  showUrlProblem(v);
  if (!v.ok) return;
  els.testResult.hidden = false;
  els.testResult.className = "test-result";
  els.testResult.textContent = "Testing…";
  els.test.disabled = true;
  const client = createBackendClient(v.url);
  const h = await client.health(4000);
  els.test.disabled = false;
  els.testResult.textContent = "";
  if (!h.ok) {
    els.testResult.className = "test-result fail";
    const p = document.createElement("div");
    p.textContent = `Not reachable at ${client.base}: ${String(h.error || "no answer").replace(/[.?!]+$/, "")}.`;
    const p2 = document.createElement("div");
    p2.textContent = "Start office-backend (see “How to start the backend” below), then click Test connection.";
    els.testResult.append(p, p2);
    els.howto.open = true;
    return;
  }
  els.testResult.className = "test-result ok";
  const head = document.createElement("strong");
  head.textContent = `Connected to ${h.name || "office-backend"}${h.version ? ` ${h.version}` : ""}`;
  const dl = document.createElement("dl");
  const rows = [
    ["Address", client.base],
    ["API", h.api ? `v${h.api}` : "v1 (old backend — please update it for full support)"],
    ["Opens", listText(h.parse)],
    ["Saves Word as", listText(h.render && h.render.word)],
    ["Saves Excel as", listText(h.render && h.render.excel)],
  ];
  for (const [k, val] of rows) {
    const dt = document.createElement("dt");
    dt.textContent = k;
    const dd = document.createElement("dd");
    dd.textContent = val;
    dl.append(dt, dd);
  }
  els.testResult.append(head, dl);
}

// ------------------------------------------------------------------ web permission

async function hasWebAccess() {
  try {
    return await chrome.permissions.contains({ origins: P.WEB_ORIGINS });
  } catch {
    return false;
  }
}

/** Cached so a toggle can call permissions.request synchronously in its gesture. */
let webGranted = false;

async function refreshPermission() {
  const granted = await hasWebAccess();
  webGranted = granted;
  els.permStatus.textContent = granted ? "Website access: allowed" : "Website access: not allowed";
  els.permStatus.className = `pill ${granted ? "ok" : "off"}`;
  els.btnPerm.hidden = false;
  els.btnPerm.textContent = granted ? "Revoke" : "Allow";
  els.btnPerm.dataset.granted = String(granted);
  return granted;
}

/** Ask for website access (needs the current click as user gesture). */
async function requestWebAccess() {
  try {
    const ok = await chrome.permissions.request({ origins: P.WEB_ORIGINS });
    if (!ok) flash("Website access was not granted — links and downloads will not be opened in the editor.", true);
    return ok;
  } catch (e) {
    flash(`Could not request access: ${e.message || e}`, true);
    return false;
  } finally {
    refreshPermission();
  }
}

async function onPermButton() {
  if (els.btnPerm.dataset.granted === "true") {
    await chrome.permissions.remove({ origins: P.WEB_ORIGINS }).catch(() => false);
    await refreshPermission();
    flash("Website access removed.");
  } else {
    await requestWebAccess();
  }
}

async function onToggle(key, checkbox) {
  const on = checkbox.checked;
  // Turning a website feature on without access: ask right away (this
  // change event is the user gesture the permission prompt needs).
  if (on && !webGranted) {
    await requestWebAccess();
  }
  await patchSettings({ [key]: on });
  syncDependentControls();
}

function syncDependentControls() {
  els.interceptText.disabled = !els.intercept.checked;
}

// ------------------------------------------------------------------ local files

async function refreshFileAccess() {
  let allowed = false;
  try {
    allowed = await chrome.extension.isAllowedFileSchemeAccess();
  } catch {
    allowed = false;
  }
  els.fileStatus.textContent = allowed ? "Access to file URLs: allowed" : "Access to file URLs: off";
  els.fileStatus.className = `pill ${allowed ? "ok" : ""}`;
  els.btnFileAccess.hidden = allowed;
}

// ------------------------------------------------------------------ init

async function init() {
  const s = await loadSettings();
  const integ = P.integrationSettings(s);
  els.enabled.checked = s.backend_enabled !== false;
  els.url.value = s.backend_url || DEFAULT_SETTINGS.backend_url;
  els.intercept.checked = integ.intercept_downloads;
  els.interceptText.checked = integ.intercept_text;
  els.interceptDrops.checked = integ.intercept_drops;
  syncDependentControls();
  showUrlProblem(P.validateBackendUrl(els.url.value));

  els.enabled.addEventListener("change", () => patchSettings({ backend_enabled: els.enabled.checked }));
  els.url.addEventListener("change", () => saveUrl());
  els.url.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      saveUrl();
    }
  });
  els.url.addEventListener("input", () => {
    els.urlError.hidden = true;
    els.url.removeAttribute("aria-invalid");
  });
  els.test.addEventListener("click", async () => {
    if (await saveUrl()) runTest();
  });
  els.btnPerm.addEventListener("click", onPermButton);
  els.intercept.addEventListener("change", () => onToggle("intercept_downloads", els.intercept));
  els.interceptText.addEventListener("change", () => onToggle("intercept_text", els.interceptText));
  els.interceptDrops.addEventListener("change", () => onToggle("intercept_drops", els.interceptDrops));
  els.btnFileAccess.addEventListener("click", () => {
    chrome.tabs.create({ url: `chrome://extensions/?id=${chrome.runtime.id}` });
  });
  els.btnClearRecents.addEventListener("click", async () => {
    await clearRecents();
    flash("Recent files cleared.");
  });

  // Keep the page truthful when settings/permissions change elsewhere.
  chrome.permissions.onAdded.addListener(refreshPermission);
  chrome.permissions.onRemoved.addListener(refreshPermission);
  document.addEventListener("visibilitychange", () => {
    if (!document.hidden) {
      refreshPermission();
      refreshFileAccess();
    }
  });

  if (location.hash === "#backend") els.howto.open = true;
  await Promise.all([refreshPermission(), refreshFileAccess()]);
  if (els.enabled.checked) runTest();
}

init();
