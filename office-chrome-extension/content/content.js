/**
 * Content script (classic script; content/policy.js is injected first).
 *
 * Registered by the service worker only after the user granted web access,
 * and it stays out of the way unless the user's settings ask for it:
 *
 *  Link clicks (setting "intercept_downloads", on by default)
 *    Only a trusted, plain left click (no Ctrl/Shift/Alt/Meta, primary
 *    button) on a same-tab link to an Office document (.docx .doc .rtf .odt
 *    .xlsx .xls .xlsm .ods .pptx; .txt/.md/.csv/.tsv/.tex only with "intercept_text"),
 *    without a download attribute, that the page itself did not handle.
 *    We listen in the bubble phase on window and never stop propagation, so
 *    page handlers and SPA routers see every click first. The service worker
 *    fetches the file; if it is not really a document (sign-in page, HTML
 *    viewer, error) or anything fails, we navigate normally.
 *    Modifier/middle clicks and right-clicks ("Save link as…") tell the
 *    service worker to let the resulting download through untouched.
 *
 *  Dropped files (setting "intercept_drops", OFF by default)
 *    Only when the drop lands outside page drop zones / editable elements
 *    and the page did not handle the drop itself (checked in the bubble
 *    phase via defaultPrevented), so uploads in Gmail, Drive, GitHub etc.
 *    keep working.
 *
 * After an extension update/reload this script is orphaned (chrome.runtime
 * is gone); it then does nothing and the browser behaves normally.
 */
(() => {
  "use strict";
  const P = globalThis.OfficeExtPolicy;
  if (!P) return;

  const alive = () => {
    try {
      return !!(chrome.runtime && chrome.runtime.id);
    } catch {
      return false;
    }
  };

  // Re-injection into a page that already runs a live copy: keep the old one.
  const prev = globalThis.__officeExtContent;
  if (prev && prev.alive()) return;
  globalThis.__officeExtContent = { alive };

  // ------------------------------------------------------------ settings

  /** null until loaded: nothing is intercepted before we know the settings. */
  let settings = null;

  try {
    chrome.storage.local.get("settings").then(
      (d) => { settings = P.integrationSettings(d && d.settings); },
      () => { settings = null; }
    );
    chrome.storage.onChanged.addListener((changes, area) => {
      if (area === "local" && changes.settings) settings = P.integrationSettings(changes.settings.newValue);
    });
  } catch {
    return; // no extension context at all
  }

  async function send(msg) {
    if (!alive()) throw new Error("The extension was updated or disabled. Reload this page and try again.");
    return chrome.runtime.sendMessage(msg);
  }

  // ------------------------------------------------------------ toast

  let toastHost = null;
  let toastBox = null;
  let toastTimer = 0;

  function toast(text, { kind = "info", ms = 4000 } = {}) {
    if (!toastHost) {
      toastHost = document.createElement("office-ext-toast");
      toastHost.style.cssText = "all:initial;position:fixed;right:16px;bottom:16px;z-index:2147483647;";
      const shadow = toastHost.attachShadow({ mode: "closed" });
      const style = document.createElement("style");
      style.textContent =
        ".t{font:13px/1.4 system-ui,-apple-system,'Segoe UI',Roboto,sans-serif;color:#fff;background:#1f2937;" +
        "padding:10px 14px;border-radius:8px;box-shadow:0 6px 24px rgba(0,0,0,.25);max-width:360px}" +
        ".t.error{background:#b42318}.t.ok{background:#166534}";
      toastBox = document.createElement("div");
      toastBox.className = "t";
      toastBox.setAttribute("role", "status");
      shadow.append(style, toastBox);
    }
    toastBox.textContent = text;
    toastBox.className = `t ${kind}`;
    if (!toastHost.isConnected) document.documentElement.appendChild(toastHost);
    clearTimeout(toastTimer);
    if (ms > 0) toastTimer = setTimeout(hideToast, ms);
  }

  function hideToast() {
    clearTimeout(toastTimer);
    if (toastHost && toastHost.isConnected) toastHost.remove();
  }

  // ------------------------------------------------------------ links

  /** The <a>/<area> the event happened on (works inside open shadow roots). */
  function linkFromEvent(e) {
    const path = typeof e.composedPath === "function" ? e.composedPath() : [];
    for (const el of path) {
      if (!el || !el.tagName) continue;
      const tag = String(el.tagName).toUpperCase();
      if ((tag === "A" || tag === "AREA") && typeof el.href === "string" && el.href) return el;
    }
    return null;
  }

  function describeLink(a) {
    return {
      href: a.href,
      target: a.getAttribute("target") || "",
      download: a.hasAttribute("download") ? a.getAttribute("download") : undefined,
      baseUrl: document.baseURI,
    };
  }

  /** Same eligibility, ignoring target/download: used to mark "user wants the real file". */
  function eligibleIgnoringIntent(a) {
    if (!settings) return null;
    return P.linkEligibility({ href: a.href, baseUrl: document.baseURI }, settings);
  }

  function markBypass(a) {
    const url = eligibleIgnoringIntent(a);
    if (url && alive()) send({ type: "office-bypass-url", url }).catch(() => {});
  }

  async function openLink(url) {
    const name = P.nameFromUrl(url) || "document";
    toast(`Opening ${name} in Office Editor…`, { ms: 0 });
    let res = null;
    try {
      res = await send({ type: "office-open-url", url });
    } catch {
      res = null;
    }
    if (res && res.ok) {
      hideToast();
      return;
    }
    // Not openable here: behave exactly like an untouched click.
    hideToast();
    location.assign(url);
  }

  function onClick(e) {
    if (!settings || !e.isTrusted || e.defaultPrevented) return;
    const a = linkFromEvent(e);
    if (!a) return;
    if (!P.isPlainClick(e)) {
      // Ctrl/Shift/Alt/Meta click: the user chose the browser's behaviour
      // (new tab/window, Alt = download); keep the resulting download.
      markBypass(a);
      return;
    }
    const url = P.linkEligibility(describeLink(a), settings);
    if (!url) return;
    if (!alive()) return; // orphaned script: let the browser navigate
    e.preventDefault(); // NOT stopPropagation: the page still sees the click
    openLink(url);
  }

  function onAuxClick(e) {
    // Middle click opens a background tab: normal behaviour, keep the file.
    if (!settings || !e.isTrusted || e.button !== 1) return;
    const a = linkFromEvent(e);
    if (a) markBypass(a);
  }

  function onContextMenu(e) {
    // Right-click may lead to "Save link as…": let that download through.
    // ("Open in Office Editor" from our own menu item is unaffected.)
    if (!settings || !e.isTrusted) return;
    const a = linkFromEvent(e);
    if (a) markBypass(a);
  }

  window.addEventListener("click", onClick, false);
  window.addEventListener("auxclick", onAuxClick, false);
  window.addEventListener("contextmenu", onContextMenu, true);

  // ------------------------------------------------------------ drops

  let overlayEl = null;
  let overlayTimer = 0;

  function showOverlay() {
    if (!overlayEl) {
      overlayEl = document.createElement("office-ext-drop");
      overlayEl.style.cssText =
        "all:initial;position:fixed;inset:12px;z-index:2147483646;pointer-events:none;" +
        "display:flex;align-items:center;justify-content:center;" +
        "border:3px dashed #1a56db;border-radius:12px;background:rgba(26,86,219,.10);" +
        "color:#1a56db;font:600 18px/1.4 system-ui,-apple-system,'Segoe UI',Roboto,sans-serif;";
      overlayEl.textContent = "Drop to open in Office Editor";
    }
    if (!overlayEl.isConnected) document.documentElement.appendChild(overlayEl);
    clearTimeout(overlayTimer);
    // dragover repeats every ~50 ms while hovering; hide shortly after it stops.
    overlayTimer = setTimeout(hideOverlay, 250);
  }

  function hideOverlay() {
    clearTimeout(overlayTimer);
    if (overlayEl && overlayEl.isConnected) overlayEl.remove();
  }

  function dropAllowedHere(e) {
    if (!settings || !settings.intercept_drops || !e.isTrusted || e.defaultPrevented) return false;
    const dt = e.dataTransfer;
    if (!dt || !dt.types || ![...dt.types].includes("Files")) return false;
    const path = typeof e.composedPath === "function" ? e.composedPath() : [];
    return !P.isPageDropTarget(path);
  }

  function onDragOver(e) {
    if (!dropAllowedHere(e)) {
      hideOverlay();
      return;
    }
    const items = [...(e.dataTransfer.items || [])].map((i) => ({ kind: i.kind, type: i.type }));
    if (!P.dragHasOfficeFile(items, settings)) return;
    e.preventDefault(); // make the page body a drop target for this drag only
    e.dataTransfer.dropEffect = "copy";
    showOverlay();
  }

  function onDrop(e) {
    hideOverlay();
    if (!dropAllowedHere(e)) return;
    const files = [...(e.dataTransfer.files || [])].filter((f) => P.fileIsOpenable(f, settings));
    if (!files.length || !alive()) return;
    e.preventDefault();
    sendFiles(files);
  }

  function bytesToBase64(buf) {
    const bytes = new Uint8Array(buf);
    let bin = "";
    const CHUNK = 0x8000;
    for (let i = 0; i < bytes.length; i += CHUNK) {
      bin += String.fromCharCode.apply(null, bytes.subarray(i, i + CHUNK));
    }
    return btoa(bin);
  }

  async function sendFile(f) {
    const mb = Math.round(P.MAX_HANDOFF_BYTES / 1048576);
    if (f.size > P.MAX_HANDOFF_BYTES) throw new Error(`${f.name} is larger than ${mb} MB.`);
    if (!f.size) throw new Error(`${f.name} is empty.`);
    const begin = await send({ type: "office-drop-begin", name: f.name, mime: f.type || "", size: f.size });
    if (!begin || !begin.ok) throw new Error((begin && begin.error) || "The editor did not accept the file.");
    const token = begin.token;
    try {
      let seq = 0;
      for (let off = 0; off < f.size; off += P.DROP_CHUNK_BYTES, seq++) {
        const buf = await f.slice(off, off + P.DROP_CHUNK_BYTES).arrayBuffer();
        const r = await send({ type: "office-drop-chunk", token, seq, b64: bytesToBase64(buf) });
        if (!r || !r.ok) throw new Error((r && r.error) || "transfer failed");
        if (f.size > P.DROP_CHUNK_BYTES) {
          const pct = Math.min(100, Math.round(((off + buf.byteLength) / f.size) * 100));
          toast(`Opening ${f.name}… ${pct}%`, { ms: 0 });
        }
      }
      const end = await send({ type: "office-drop-end", token });
      if (!end || !end.ok) throw new Error((end && end.error) || "transfer failed");
    } catch (err) {
      send({ type: "office-drop-abort", token }).catch(() => {});
      throw err;
    }
  }

  async function sendFiles(files) {
    for (const f of files) {
      toast(`Opening ${f.name} in Office Editor…`, { ms: 0 });
      try {
        await sendFile(f);
        toast(`Opened ${f.name} in a new tab`, { kind: "ok", ms: 2500 });
      } catch (err) {
        toast(`Could not open ${f.name}: ${(err && err.message) || err}`, { kind: "error", ms: 6000 });
      }
    }
  }

  window.addEventListener("dragover", onDragOver, false);
  window.addEventListener("drop", onDrop, false);
  window.addEventListener("dragend", hideOverlay, true);
})();
