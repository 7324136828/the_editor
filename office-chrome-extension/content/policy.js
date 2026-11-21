/**
 * Browser-integration policy shared by the content script, the service
 * worker, the offscreen document, the popup and the options page.
 *
 * Deliberately a dependency-free *classic* script that only assigns
 * `globalThis.OfficeExtPolicy`, so it can be
 *   - injected as a content script in front of content.js,
 *   - side-effect imported by ES modules (`import "../content/policy.js"`),
 *   - loaded by Node tests.
 * Everything here is pure (no chrome.* calls, no DOM access beyond plain
 * property reads on objects that are passed in).
 */
(function (root) {
  "use strict";
  if (root.OfficeExtPolicy) return; // injected twice into the same world

  /** Binary Office formats: link clicks / downloads are opened by default. */
  const OFFICE_EXTS = ["docx", "doc", "rtf", "odt", "xlsx", "xls", "xlsm", "ods", "pptx"];
  /**
   * Text formats: usually *web pages* (GitHub README.md, robots.txt, inline
   * CSV previews), so they are only intercepted when the user opts in.
   */
  const TEXT_EXTS = ["txt", "md", "markdown", "csv", "tsv", "tex"];
  const ALL_EXTS = OFFICE_EXTS.concat(TEXT_EXTS);

  /** Hard cap for anything handed to the viewer through the extension. */
  const MAX_HANDOFF_BYTES = 100 * 1024 * 1024;
  /** Raw bytes per drop-relay message (base64 makes it ~1.33x on the wire). */
  const DROP_CHUNK_BYTES = 4 * 1024 * 1024;

  /**
   * Integration settings live in the same `settings` object as the backend
   * settings (chrome.storage.local). Keys owned by the integration layer:
   *   intercept_downloads  open Office links/downloads in the editor (default on,
   *                        but only effective once the optional web-site
   *                        permission has been granted)
   *   intercept_text       also intercept .txt/.md/.csv/.tsv (default off)
   *   intercept_drops      open Office files dropped onto web pages (default off)
   */
  const INTEGRATION_DEFAULTS = Object.freeze({
    intercept_downloads: true,
    intercept_text: false,
    intercept_drops: false,
  });

  /** Optional host permissions that power interception (see manifest.json). */
  const WEB_ORIGINS = ["http://*/*", "https://*/*"];

  function integrationSettings(raw) {
    const s = raw && typeof raw === "object" ? raw : {};
    return {
      intercept_downloads: s.intercept_downloads !== undefined ? !!s.intercept_downloads : INTEGRATION_DEFAULTS.intercept_downloads,
      intercept_text: !!s.intercept_text,
      intercept_drops: !!s.intercept_drops,
    };
  }

  /** Lower-case extension of a file name or URL path ("" when none). */
  function extOf(name) {
    const m = /\.([A-Za-z0-9]{1,10})$/.exec(String(name || "").trim());
    return m ? m[1].toLowerCase() : "";
  }

  /** Last path segment of a URL, percent-decoded when valid. */
  function nameFromUrl(url) {
    let path;
    try {
      path = new URL(url).pathname;
    } catch {
      path = String(url || "").split(/[?#]/)[0];
    }
    const last = path.split("/").pop() || "";
    try {
      return decodeURIComponent(last);
    } catch {
      return last;
    }
  }

  /** Which extensions are eligible for interception under these settings. */
  function interceptableExt(ext, settings) {
    const s = integrationSettings(settings);
    if (OFFICE_EXTS.includes(ext)) return true;
    return s.intercept_text && TEXT_EXTS.includes(ext);
  }

  /**
   * Decide whether a click on a link should be opened in the editor.
   * `link` is a plain description: { href, target, download, baseUrl }.
   * Returns the absolute URL to open, or null to leave the click alone.
   */
  function linkEligibility(link, settings) {
    const s = integrationSettings(settings);
    if (!s.intercept_downloads || !link) return null;
    // The page explicitly asked for a download: let the browser do it (the
    // download interceptor still applies its own rules to the result).
    if (link.download !== undefined && link.download !== null && link.download !== false) return null;
    const target = String(link.target || "").trim().toLowerCase();
    if (target && target !== "_self") return null;
    let u;
    try {
      u = new URL(String(link.href || ""), link.baseUrl || undefined);
    } catch {
      return null;
    }
    if (u.protocol !== "http:" && u.protocol !== "https:") return null;
    const ext = extOf(nameFromUrl(u.href));
    if (!interceptableExt(ext, s)) return null;
    return u.href;
  }

  /** A plain primary-button click with no modifier keys (user wants "open"). */
  function isPlainClick(e) {
    return !!e && e.button === 0 && !e.metaKey && !e.ctrlKey && !e.shiftKey && !e.altKey;
  }

  /**
   * Decide whether a browser download should be opened in the editor.
   * `item` is a chrome.downloads.DownloadItem (as seen in
   * onDeterminingFilename). Returns { intercept, reason, name, url }.
   */
  function downloadEligibility(item, settings, opts) {
    const o = opts || {};
    const s = integrationSettings(settings);
    const no = (reason) => ({ intercept: false, reason });
    if (!item) return no("no item");
    if (!s.intercept_downloads) return no("disabled");
    if (o.extensionId && item.byExtensionId && item.byExtensionId === o.extensionId) return no("own download");
    const url = String(item.finalUrl || item.url || "");
    const scheme = (/^([a-z][a-z0-9+.-]*):/i.exec(url) || [])[1];
    const sch = scheme ? scheme.toLowerCase() : "";
    if (sch === "file") {
      if (!o.fileAccess) return no("no file access");
    } else if (sch !== "http" && sch !== "https") {
      return no("scheme"); // blob:, data:, chrome-extension:, ftp: …
    }
    if (o.isBypassed && o.isBypassed(url)) return no("bypassed");
    const mime = String(item.mime || "").split(";")[0].trim().toLowerCase();
    if (mime === "text/html" || mime === "application/xhtml+xml") return no("html");
    const suggested = String(item.filename || "").split(/[\\/]/).pop();
    const fromUrl = nameFromUrl(url);
    const name = extOf(suggested) ? suggested : fromUrl;
    const ext = extOf(name);
    if (!interceptableExt(ext, s)) return no("type");
    if (typeof item.fileSize === "number" && item.fileSize > MAX_HANDOFF_BYTES) return no("too large");
    if (typeof item.totalBytes === "number" && item.totalBytes > MAX_HANDOFF_BYTES) return no("too large");
    return { intercept: true, reason: "ok", name: name || "document", url };
  }

  function asciiHead(u8, n) {
    let s = "";
    const len = Math.min(u8.length, n);
    for (let i = 0; i < len; i++) s += String.fromCharCode(u8[i]);
    return s;
  }

  /** An HTML document (login page, error page, GitHub file view) rather than a file. */
  function looksLikeHtml(u8) {
    let head = asciiHead(u8, 1024).replace(/^\xEF\xBB\xBF/, "");
    return /^\s*(<!--[\s\S]*?-->\s*)*<(!doctype\s+html|html|head|body)\b/i.test(head);
  }

  /**
   * Check that fetched bytes really are the document the URL promised, so a
   * sign-in page or an HTML file viewer is never opened as a "document".
   * Returns { ok:true } or { ok:false, reason }.
   */
  function validateFetched(name, mime, bytes) {
    const u8 = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes || []);
    if (!u8.length) return { ok: false, reason: "the server sent an empty file" };
    if (u8.length > MAX_HANDOFF_BYTES) return { ok: false, reason: "the file is too large" };
    const essence = String(mime || "").split(";")[0].trim().toLowerCase();
    const ext = extOf(name);
    const isZip = u8.length >= 4 && u8[0] === 0x50 && u8[1] === 0x4b && (u8[2] === 3 || u8[2] === 5 || u8[2] === 7);
    const isOle = u8.length >= 4 && u8[0] === 0xd0 && u8[1] === 0xcf && u8[2] === 0x11 && u8[3] === 0xe0;
    const isRtf = asciiHead(u8, 5) === "{\\rtf";
    if (essence === "text/html" || essence === "application/xhtml+xml" || looksLikeHtml(u8)) {
      // .xls "exported" as an HTML table is common and SheetJS reads it.
      if (ext === "xls" && /<table/i.test(asciiHead(u8, 64 * 1024))) return { ok: true };
      return { ok: false, reason: "the link returned a web page, not a document" };
    }
    switch (ext) {
      case "docx": case "xlsx": case "xlsm": case "odt": case "ods": case "pptx":
        return isZip ? { ok: true } : { ok: false, reason: `not a valid .${ext} file` };
      case "doc":
        return isOle || isRtf || isZip ? { ok: true } : { ok: false, reason: "not a valid .doc file" };
      case "xls":
        // Real BIFF, or the CSV/TSV/XML "xls" files web apps like to export.
        return isOle || isZip || !u8.subarray(0, 4096).includes(0) ? { ok: true } : { ok: false, reason: "not a valid .xls file" };
      case "rtf":
        return isRtf ? { ok: true } : { ok: false, reason: "not a valid .rtf file" };
      default:
        // Text formats: must not be binary.
        if (TEXT_EXTS.includes(ext)) {
          const scan = u8.subarray(0, 8192);
          const utf16 = (u8[0] === 0xff && u8[1] === 0xfe) || (u8[0] === 0xfe && u8[1] === 0xff);
          return utf16 || !scan.includes(0) ? { ok: true } : { ok: false, reason: "not a text file" };
        }
        return isZip || isOle || isRtf ? { ok: true } : { ok: false, reason: "unsupported file type" };
    }
  }

  /**
   * Validate a URL typed by the user (popup "Open from URL").
   * Returns { ok:true, url } or { ok:false, error }.
   */
  function validateUserUrl(input) {
    let s = String(input || "").trim();
    if (!s) return { ok: false, error: "Enter the address of a document." };
    if (/^file:/i.test(s)) {
      return { ok: false, error: "Local files cannot be opened by address. Use “Open file…” or drag the file onto the editor tab." };
    }
    const hasScheme = /^[a-z][a-z0-9+.-]*:\/\//i.test(s) || /^(javascript|data|blob|mailto|about|chrome|edge|view-source):/i.test(s);
    if (!hasScheme) s = /^(localhost|127\.|\[::1\])/i.test(s) ? `http://${s}` : `https://${s}`;
    let u;
    try {
      u = new URL(s);
    } catch {
      return { ok: false, error: "That is not a valid web address." };
    }
    if (u.protocol !== "http:" && u.protocol !== "https:") {
      return { ok: false, error: `Only http:// and https:// addresses are supported (got ${u.protocol}).` };
    }
    // Without an explicit scheme, a dot-less word is almost certainly a typo
    // (intranet hosts work when typed as http://host/...).
    if (!u.hostname || (!hasScheme && !u.hostname.includes(".") && u.hostname !== "localhost" && !/^\[.*\]$/.test(u.hostname))) {
      return { ok: false, error: "That address has no valid host name." };
    }
    if (!hasScheme && u.pathname === "/" && ALL_EXTS.includes(extOf(u.hostname))) {
      return { ok: false, error: "That looks like a file name. Enter the full web address (https://…), or use “Open file…” for files on this computer." };
    }
    return { ok: true, url: u.href };
  }

  /**
   * Validate the backend address typed in Options. Accepts "host:port" or a
   * full http(s) URL. Returns { ok:true, url, loopback } (url normalised,
   * without trailing slash) or { ok:false, error }.
   */
  function validateBackendUrl(input) {
    let s = String(input || "").trim();
    if (!s) return { ok: false, error: "Enter the backend address, e.g. http://127.0.0.1:47831" };
    if (!/^[a-z][a-z0-9+.-]*:\/\//i.test(s)) s = `http://${s}`;
    let u;
    try {
      u = new URL(s);
    } catch {
      return { ok: false, error: "That is not a valid address. Example: http://127.0.0.1:47831" };
    }
    if (u.protocol !== "http:" && u.protocol !== "https:") return { ok: false, error: "The address must start with http:// or https://" };
    if (!u.hostname) return { ok: false, error: "The address has no host name." };
    if (u.username || u.password) return { ok: false, error: "Do not put a user name or password in the address." };
    if (u.search || u.hash) return { ok: false, error: "The address must not contain ? or #." };
    const host = u.hostname.toLowerCase();
    const loopback = host === "localhost" || host === "[::1]" || host === "::1" || /^127\.\d+\.\d+\.\d+$/.test(host);
    const url = (u.origin + u.pathname).replace(/\/+$/, "");
    return { ok: true, url, loopback };
  }

  /**
   * Small sliding-window rate limiter: at most `max` events per `windowMs`
   * overall, and at most one per `perKeyMs` for the same key (e.g. tab id).
   */
  function createRateLimiter(opts) {
    const o = Object.assign({ max: 5, windowMs: 10000, perKeyMs: 1500 }, opts || {});
    let events = [];
    const lastByKey = new Map();
    return {
      allow(key, now) {
        const t = typeof now === "number" ? now : Date.now();
        events = events.filter((x) => t - x < o.windowMs);
        if (events.length >= o.max) return false;
        if (key !== undefined && key !== null) {
          const last = lastByKey.get(key);
          if (last !== undefined && t - last < o.perKeyMs) return false;
          lastByKey.set(key, t);
          if (lastByKey.size > 500) lastByKey.delete(lastByKey.keys().next().value);
        }
        events.push(t);
        return true;
      },
    };
  }

  /** Office-looking MIME types visible in dataTransfer.items during a drag. */
  const OFFICE_DRAG_MIME = /officedocument|msword|ms-excel|opendocument|application\/rtf|text\/rtf/i;
  const TEXT_DRAG_MIME = /^(text\/csv|text\/tab-separated-values|text\/markdown|text\/plain|application\/csv)$/i;

  /**
   * True when a drag carries at least one file the editor can open, judged
   * from dataTransfer.items (names are not readable until the drop).
   * `items` is an array of { kind, type }.
   */
  function dragHasOfficeFile(items, settings) {
    const s = integrationSettings(settings);
    for (const it of items || []) {
      if (!it || it.kind !== "file") continue;
      const t = String(it.type || "");
      if (OFFICE_DRAG_MIME.test(t)) return true;
      if (s.intercept_text && TEXT_DRAG_MIME.test(t)) return true;
    }
    return false;
  }

  /** True when a dropped File should be opened in the editor. */
  function fileIsOpenable(file, settings) {
    if (!file) return false;
    return interceptableExt(extOf(file.name), settings);
  }

  /**
   * True for drop targets that belong to the page: form fields, editors and
   * elements that look like upload zones. `chain` is the event's composed
   * path, as plain objects ({ tagName, isContentEditable, type, ... }).
   */
  function isPageDropTarget(chain) {
    for (const el of chain || []) {
      if (!el || !el.tagName) continue;
      const tag = String(el.tagName).toUpperCase();
      if (tag === "INPUT" || tag === "TEXTAREA" || tag === "SELECT") return true;
      if (el.isContentEditable) return true;
      const attrs = [el.id, el.className, el.role, el.dropzone]
        .map((x) => (typeof x === "string" ? x : ""))
        .join(" ");
      if (/drop.?zone|dropzone|upload|attach/i.test(attrs)) return true;
    }
    return false;
  }

  /** "just now", "5 min ago", "3 h ago", "yesterday", "4 days ago", or a date. */
  function relativeTime(ts, now) {
    if (!ts) return "";
    const t = typeof now === "number" ? now : Date.now();
    const diff = Math.max(0, t - ts);
    const min = 60 * 1000;
    const hour = 60 * min;
    const day = 24 * hour;
    if (diff < min) return "just now";
    if (diff < hour) return `${Math.floor(diff / min)} min ago`;
    if (diff < day) return `${Math.floor(diff / hour)} h ago`;
    if (diff < 2 * day) return "yesterday";
    if (diff < 7 * day) return `${Math.floor(diff / day)} days ago`;
    return new Date(ts).toLocaleDateString(undefined, { month: "short", day: "numeric", year: diff > 300 * day ? "numeric" : undefined });
  }

  /** Context-menu targetUrlPatterns for links to openable documents. */
  function linkMenuPatterns() {
    const out = [];
    for (const ext of ALL_EXTS) {
      for (const e of [ext, ext.toUpperCase()]) {
        out.push(`*://*/*.${e}`, `*://*/*.${e}?*`, `file:///*.${e}`);
      }
    }
    return out;
  }

  /** documentUrlPatterns for "open this page" on text documents shown inline. */
  function pageMenuPatterns() {
    const out = [];
    for (const ext of TEXT_EXTS) out.push(`*://*/*.${ext}`, `*://*/*.${ext}?*`, `file:///*.${ext}`);
    return out;
  }

  root.OfficeExtPolicy = Object.freeze({
    OFFICE_EXTS,
    TEXT_EXTS,
    ALL_EXTS,
    MAX_HANDOFF_BYTES,
    DROP_CHUNK_BYTES,
    INTEGRATION_DEFAULTS,
    WEB_ORIGINS,
    integrationSettings,
    extOf,
    nameFromUrl,
    interceptableExt,
    linkEligibility,
    isPlainClick,
    downloadEligibility,
    looksLikeHtml,
    validateFetched,
    validateUserUrl,
    validateBackendUrl,
    createRateLimiter,
    dragHasOfficeFile,
    fileIsOpenable,
    isPageDropTarget,
    relativeTime,
    linkMenuPatterns,
    pageMenuPatterns,
  });
})(typeof globalThis !== "undefined" ? globalThis : self);
