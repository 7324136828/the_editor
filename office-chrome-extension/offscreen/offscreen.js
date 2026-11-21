/**
 * Offscreen document: owns writes to the OPFS handoff store for bytes that
 * do not originate in an extension page.
 *
 *  - office-fetch-store   fetch a URL itself (http/https with the user's
 *                          cookies, or file:// via XHR when file access is
 *                          allowed), validate that it really is a document,
 *                          and stream it into a handoff. The bytes never
 *                          travel through extension messaging.
 *  - office-upload-*      receive a dropped file from a content script in
 *                          base64 chunks (relayed by the service worker) and
 *                          append them to a streamed handoff.
 *  - office-sweep         delete stale handoffs.
 *
 * Only the service worker talks to this document (sender has no tab).
 */

import "../content/policy.js";
import { beginHandoff, sweepHandoffs, deleteHandoff } from "../lib/handoff.js";
import { filenameFromContentDisposition } from "../lib/backend.js";

const P = globalThis.OfficeExtPolicy;
const FETCH_TIMEOUT_MS = 120000;

function b64ToBytes(b64) {
  const bin = atob(b64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return bytes;
}

class FetchError extends Error {
  constructor(message, reason) {
    super(message);
    this.reason = reason; // "network" | "http" | "invalid" | "too-large"
  }
}

/** Read a local file:// URL (fetch() does not support the file scheme). */
function readFileUrl(url, maxBytes) {
  return new Promise((resolve, reject) => {
    const xhr = new XMLHttpRequest();
    xhr.open("GET", url);
    xhr.responseType = "arraybuffer";
    xhr.timeout = FETCH_TIMEOUT_MS;
    xhr.onload = () => {
      const buf = xhr.response;
      if (!buf || !buf.byteLength) return reject(new FetchError("could not read the local file", "network"));
      if (buf.byteLength > maxBytes) return reject(new FetchError("the file is too large", "too-large"));
      resolve({ bytes: new Uint8Array(buf), mime: "", name: "" });
    };
    xhr.onerror = () => reject(new FetchError("could not read the local file (is “Allow access to file URLs” enabled?)", "network"));
    xhr.ontimeout = () => reject(new FetchError("reading the file timed out", "network"));
    xhr.send();
  });
}

/** GET an http(s) URL with the user's cookies, enforcing the size cap while streaming. */
async function readHttpUrl(url, maxBytes) {
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), FETCH_TIMEOUT_MS);
  try {
    let res;
    try {
      res = await fetch(url, { credentials: "include", redirect: "follow", signal: ctrl.signal });
    } catch (e) {
      throw new FetchError(`could not download the file (${e && e.message || e})`, "network");
    }
    if (!res.ok) throw new FetchError(`the server answered HTTP ${res.status}`, "http");
    const len = Number(res.headers.get("Content-Length") || 0);
    if (len > maxBytes) throw new FetchError("the file is too large", "too-large");
    const chunks = [];
    let size = 0;
    if (res.body && res.body.getReader) {
      const reader = res.body.getReader();
      for (;;) {
        const { value, done } = await reader.read();
        if (done) break;
        size += value.byteLength;
        if (size > maxBytes) {
          ctrl.abort();
          throw new FetchError("the file is too large", "too-large");
        }
        chunks.push(value);
      }
    } else {
      const buf = new Uint8Array(await res.arrayBuffer());
      size = buf.byteLength;
      if (size > maxBytes) throw new FetchError("the file is too large", "too-large");
      chunks.push(buf);
    }
    const bytes = new Uint8Array(size);
    let off = 0;
    for (const c of chunks) {
      bytes.set(c, off);
      off += c.byteLength;
    }
    return {
      bytes,
      mime: res.headers.get("Content-Type") || "",
      name: filenameFromContentDisposition(res.headers.get("Content-Disposition")),
      finalUrl: res.url || url,
    };
  } finally {
    clearTimeout(timer);
  }
}

async function fetchAndStore({ url, nameHint, maxBytes }) {
  const cap = Math.min(Number(maxBytes) || P.MAX_HANDOFF_BYTES, P.MAX_HANDOFF_BYTES);
  const isFile = /^file:/i.test(url);
  if (!isFile && !/^https?:/i.test(url)) throw new FetchError("unsupported address", "invalid");
  const got = isFile ? await readFileUrl(url, cap) : await readHttpUrl(url, cap);
  // Name: Content-Disposition, then the caller's hint (browser-suggested
  // download name), then the final URL's last path segment.
  let name = got.name || nameHint || P.nameFromUrl(got.finalUrl || url) || "document";
  if (!P.extOf(name) && nameHint && P.extOf(nameHint)) name = nameHint;
  if (!P.extOf(name)) {
    const fromUrl = P.nameFromUrl(url);
    if (P.extOf(fromUrl)) name = fromUrl;
  }
  const verdict = P.validateFetched(name, got.mime, got.bytes);
  if (!verdict.ok) throw new FetchError(verdict.reason, "invalid");
  const w = await beginHandoff(name, got.mime, cap);
  try {
    await w.write(got.bytes);
    const id = await w.close();
    return { id, name, size: got.bytes.byteLength };
  } catch (e) {
    await w.abort();
    throw e;
  }
}

// ---------------- chunked uploads (content-script drops) ----------------

/** uploadId -> { writer, nextSeq, expected, timer } */
const uploads = new Map();
const UPLOAD_IDLE_MS = 60000;

function touchUpload(uploadId) {
  const u = uploads.get(uploadId);
  if (!u) return;
  clearTimeout(u.timer);
  u.timer = setTimeout(() => {
    uploads.delete(uploadId);
    u.writer.abort().catch(() => {});
  }, UPLOAD_IDLE_MS);
}

async function uploadBegin({ name, mime, size }) {
  if (!(Number(size) >= 0) || Number(size) > P.MAX_HANDOFF_BYTES) {
    throw new FetchError(`the file is larger than ${Math.round(P.MAX_HANDOFF_BYTES / 1024 / 1024)} MB`, "too-large");
  }
  const writer = await beginHandoff(name || "document", mime || "", P.MAX_HANDOFF_BYTES);
  const uploadId = writer.id;
  uploads.set(uploadId, { writer, nextSeq: 0, expected: Number(size) });
  touchUpload(uploadId);
  return { uploadId };
}

async function uploadChunk({ uploadId, seq, b64 }) {
  const u = uploads.get(uploadId);
  if (!u) throw new Error("unknown or expired upload");
  if (seq !== u.nextSeq) {
    uploads.delete(uploadId);
    clearTimeout(u.timer);
    await u.writer.abort();
    throw new Error("upload chunks arrived out of order");
  }
  u.nextSeq++;
  touchUpload(uploadId);
  await u.writer.write(b64ToBytes(String(b64 || "")));
  return {};
}

async function uploadEnd({ uploadId }) {
  const u = uploads.get(uploadId);
  if (!u) throw new Error("unknown or expired upload");
  uploads.delete(uploadId);
  clearTimeout(u.timer);
  if (u.writer.size !== u.expected) {
    await u.writer.abort();
    throw new Error("the file was not transferred completely");
  }
  const id = await u.writer.close();
  return { id };
}

async function uploadAbort({ uploadId }) {
  const u = uploads.get(uploadId);
  if (!u) return {};
  uploads.delete(uploadId);
  clearTimeout(u.timer);
  await u.writer.abort();
  return {};
}

const HANDLERS = {
  "office-fetch-store": fetchAndStore,
  "office-upload-begin": uploadBegin,
  "office-upload-chunk": uploadChunk,
  "office-upload-end": uploadEnd,
  "office-upload-abort": uploadAbort,
  "office-sweep": async () => ({ removed: await sweepHandoffs() }),
  "office-delete-handoff": async ({ id }) => {
    await deleteHandoff(String(id || ""));
    return {};
  },
};

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (!msg || msg.target !== "office-offscreen") return false;
  // Only the extension's own service worker may drive this document.
  if (sender.id !== chrome.runtime.id || sender.tab) return false;
  const handler = HANDLERS[msg.type];
  if (!handler) return false;
  (async () => {
    try {
      const out = await handler(msg);
      sendResponse({ ok: true, ...out });
    } catch (e) {
      sendResponse({ ok: false, error: String((e && e.message) || e), reason: (e && e.reason) || "error" });
    }
  })();
  return true;
});

// Housekeeping on load (age-based: never touches in-flight entries).
sweepHandoffs().catch(() => {});
