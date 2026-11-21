/**
 * OPFS-based file stores shared by all extension pages (popup, options,
 * viewer, offscreen share one origin-private filesystem):
 *
 *  - "office_handoff": transient handoffs between contexts (popup, content-
 *    script drops, intercepted links/downloads -> viewer). The viewer reads a
 *    handoff, moves the document into recents and deletes the handoff.
 *    Entries that are never consumed are swept after HANDOFF_MAX_AGE_MS.
 *  - "office_recent": persistent recent-files store so the popup can reopen
 *    documents. One entry per file name (re-opening or saving replaces it),
 *    newest first, capped by count and total size.
 *
 * On-disk layout per entry:  <dir>/<id>/meta.json + <dir>/<id>/data
 *   id         `${Date.now().toString(36)}-${random}` (creation time is
 *              recoverable from the id even before meta.json is readable)
 *   meta.json  { name, mime, size, created, complete }
 * meta.json is written FIRST with complete:false and rewritten with
 * complete:true once all bytes are on disk, so readers and the sweeper can
 * tell an in-flight entry from a finished one.
 *
 * Content scripts cannot reach OPFS; they stream bytes to the offscreen
 * document (via the service worker), which uses beginHandoff() below.
 */

const HANDOFF_DIR = "office_handoff";
const RECENT_DIR = "office_recent";

/** Unconsumed handoffs older than this are deleted by sweepHandoffs(). */
export const HANDOFF_MAX_AGE_MS = 10 * 60 * 1000;
export const MAX_RECENT = 10;
/** Recents are trimmed (oldest first) beyond this many bytes in total. */
export const MAX_RECENT_TOTAL_BYTES = 200 * 1024 * 1024;
/** Largest single handoff accepted by beginHandoff(). */
export const MAX_HANDOFF_BYTES = 100 * 1024 * 1024;

const ID_RE = /^[a-z0-9]+-[a-z0-9]+$/i;

let rootProvider = () => navigator.storage.getDirectory();

/**
 * Test hook: replace the OPFS root with an in-memory implementation
 * (pass null to restore navigator.storage.getDirectory()).
 */
export function __setStorageRootForTests(provider) {
  rootProvider = provider || (() => navigator.storage.getDirectory());
}

async function storeDir(name, create = true) {
  const root = await rootProvider();
  return root.getDirectoryHandle(name, { create });
}

export function newStoreId(now = Date.now()) {
  return `${now.toString(36)}-${Math.random().toString(36).slice(2, 10) || "0"}`;
}

/** Creation time encoded in a store id (0 when the id is not one of ours). */
export function idTimestamp(id) {
  const m = /^([a-z0-9]+)-/i.exec(String(id || ""));
  if (!m) return 0;
  const t = parseInt(m[1], 36);
  return Number.isFinite(t) ? t : 0;
}

async function writeFile(dirHandle, fileName, data) {
  const fh = await dirHandle.getFileHandle(fileName, { create: true });
  const wh = await fh.createWritable();
  try {
    await wh.write(data);
    await wh.close();
  } catch (e) {
    try { await wh.abort?.(); } catch { /* ignore */ }
    throw e;
  }
}

async function readMeta(sub) {
  const metaFile = await (await sub.getFileHandle("meta.json")).getFile();
  return JSON.parse(await metaFile.text());
}

function toBytes(bytes) {
  if (bytes instanceof Uint8Array) return bytes;
  if (bytes instanceof ArrayBuffer) return new Uint8Array(bytes);
  if (ArrayBuffer.isView(bytes)) return new Uint8Array(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  return new Uint8Array(bytes || []);
}

/** Create an entry with an incomplete meta.json; returns { id, sub, meta }. */
async function createEntry(dirName, name, mime) {
  const now = Date.now();
  const id = newStoreId(now);
  const dir = await storeDir(dirName, true);
  const sub = await dir.getDirectoryHandle(id, { create: true });
  const meta = { name: sanitizeName(name), mime: String(mime || ""), size: 0, created: now, complete: false };
  await writeFile(sub, "meta.json", JSON.stringify(meta));
  return { id, sub, meta };
}

async function putIn(dirName, name, bytes, mime = "") {
  const data = toBytes(bytes);
  const { id, sub, meta } = await createEntry(dirName, name, mime);
  try {
    await writeFile(sub, "data", data);
    meta.size = data.byteLength;
    meta.complete = true;
    await writeFile(sub, "meta.json", JSON.stringify(meta));
  } catch (e) {
    await deleteIn(dirName, id);
    throw e;
  }
  return id;
}

async function getIn(dirName, id) {
  if (!ID_RE.test(String(id))) throw new Error("bad store id");
  const dir = await storeDir(dirName, false);
  const sub = await dir.getDirectoryHandle(id);
  const meta = await readMeta(sub);
  if (meta.complete === false) throw new Error("the file is still being transferred");
  const dataFile = await (await sub.getFileHandle("data")).getFile();
  const buf = await dataFile.arrayBuffer();
  return {
    id,
    name: meta.name,
    mime: meta.mime || "",
    size: meta.size ?? buf.byteLength,
    created: meta.created || idTimestamp(id),
    bytes: new Uint8Array(buf),
  };
}

async function deleteIn(dirName, id) {
  try {
    const dir = await storeDir(dirName, false);
    await dir.removeEntry(id, { recursive: true });
  } catch {
    /* already gone */
  }
}

/** Finished entries, newest first (in-flight and unreadable ones skipped). */
async function listIn(dirName) {
  let dir;
  try {
    dir = await storeDir(dirName, false);
  } catch {
    return [];
  }
  const out = [];
  for await (const [name, handle] of dir.entries()) {
    if (handle && handle.kind && handle.kind !== "directory") continue;
    try {
      const sub = await dir.getDirectoryHandle(name);
      const meta = await readMeta(sub);
      if (meta.complete === false) continue;
      out.push({
        id: name,
        name: meta.name,
        mime: meta.mime || "",
        size: meta.size || 0,
        created: meta.created || idTimestamp(name),
      });
    } catch {
      /* skip unreadable entries */
    }
  }
  out.sort((a, b) => b.created - a.created || (a.id < b.id ? 1 : -1));
  return out;
}

// ---------------- transient handoffs ----------------

export const putHandoff = (name, bytes, mime) => putIn(HANDOFF_DIR, name, bytes, mime);
export const getHandoff = (id) => getIn(HANDOFF_DIR, id);
export const deleteHandoff = (id) => deleteIn(HANDOFF_DIR, id);
export const listHandoffs = () => listIn(HANDOFF_DIR);

/**
 * Streamed handoff for large transfers (chunks arriving over messaging, or
 * a fetch body). Usage:
 *   const w = await beginHandoff(name, mime);
 *   await w.write(chunk); ...; const id = await w.close();   // or w.abort()
 * The entry is invisible to readers until close() marks it complete.
 */
export async function beginHandoff(name, mime = "", maxBytes = MAX_HANDOFF_BYTES) {
  const { id, sub, meta } = await createEntry(HANDOFF_DIR, name, mime);
  let writable;
  try {
    writable = await (await sub.getFileHandle("data", { create: true })).createWritable();
  } catch (e) {
    await deleteIn(HANDOFF_DIR, id);
    throw e;
  }
  let size = 0;
  let done = false;
  const fail = async () => {
    done = true;
    try { await writable.abort?.(); } catch { /* ignore */ }
    await deleteIn(HANDOFF_DIR, id);
  };
  return {
    id,
    get size() { return size; },
    async write(chunk) {
      if (done) throw new Error("handoff already finished");
      const data = toBytes(chunk);
      if (size + data.byteLength > maxBytes) {
        await fail();
        throw new Error(`the file is larger than ${Math.round(maxBytes / 1024 / 1024)} MB`);
      }
      try {
        await writable.write(data);
      } catch (e) {
        await fail();
        throw e;
      }
      size += data.byteLength;
    },
    async close() {
      if (done) throw new Error("handoff already finished");
      done = true;
      try {
        await writable.close();
        meta.size = size;
        meta.complete = true;
        await writeFile(sub, "meta.json", JSON.stringify(meta));
      } catch (e) {
        await deleteIn(HANDOFF_DIR, id);
        throw e;
      }
      return id;
    },
    async abort() {
      if (!done) await fail();
    },
  };
}

/**
 * Delete handoffs that were never consumed. Only the entry's AGE matters
 * (from its id, falling back to meta.created): an entry that is still being
 * written — even one whose meta.json is not readable yet — is never touched
 * while it is younger than maxAgeMs.
 */
export async function sweepHandoffs(maxAgeMs = HANDOFF_MAX_AGE_MS, now = Date.now()) {
  const dir = await storeDir(HANDOFF_DIR, false).catch(() => null);
  if (!dir) return 0;
  const names = [];
  for await (const [name] of dir.entries()) names.push(name);
  let removed = 0;
  for (const name of names) {
    let created = idTimestamp(name);
    if (!created) {
      try {
        created = (await readMeta(await dir.getDirectoryHandle(name))).created || 0;
      } catch {
        created = 0; // foreign/broken entry with no recoverable age: remove it
      }
    }
    if (!created || now - created > maxAgeMs) {
      await deleteIn(HANDOFF_DIR, name);
      removed++;
    }
  }
  return removed;
}

// ---------------- recent files ----------------

export const putRecent = (name, bytes, mime) => putIn(RECENT_DIR, name, bytes, mime);
export const getRecent = (id) => getIn(RECENT_DIR, id);
export const deleteRecent = (id) => deleteIn(RECENT_DIR, id);
export const listRecent = () => listIn(RECENT_DIR);

function sameRecentName(a, b) {
  return String(a || "").toLowerCase() === String(b || "").toLowerCase();
}

/**
 * Record a file as recent. An existing entry with the same name (case-
 * insensitive) is replaced, so re-opening a file or saving it again keeps a
 * single, up-to-date entry. Keeps at most MAX_RECENT entries and
 * MAX_RECENT_TOTAL_BYTES in total (newest win). Returns the new id.
 */
export async function addRecent(name, bytes, mime = "") {
  const id = await putIn(RECENT_DIR, name, bytes, mime);
  const all = await listIn(RECENT_DIR);
  const mine = all.find((e) => e.id === id);
  const myName = mine ? mine.name : sanitizeName(name);
  const myCreated = mine ? mine.created : Date.now();
  const keep = [];
  for (const e of all) {
    if (e.id === id) {
      keep.push(e);
      continue;
    }
    // Only replace OLDER entries with the same name, so two concurrent
    // writers can never delete each other's entry (the newest survives).
    const older = e.created < myCreated || (e.created === myCreated && e.id < id);
    if (sameRecentName(e.name, myName) && older) await deleteIn(RECENT_DIR, e.id);
    else keep.push(e);
  }
  let total = 0;
  for (let i = 0; i < keep.length; i++) {
    const e = keep[i];
    total += e.size || 0;
    const overCount = i >= MAX_RECENT;
    const overSize = total > MAX_RECENT_TOTAL_BYTES && e.id !== id;
    if (overCount || overSize) await deleteIn(RECENT_DIR, e.id);
  }
  return id;
}

/**
 * Store the last SAVED version of a document under its recent entry (the
 * viewer calls this after a successful save). Same semantics as addRecent:
 * the previous entry with that name is replaced and moves to the top.
 */
export const updateRecent = (name, bytes, mime = "") => addRecent(name, bytes, mime);

/** Remove every recent file. */
export async function clearRecents() {
  for (const e of await listIn(RECENT_DIR)) await deleteIn(RECENT_DIR, e.id);
  // Also drop anything unreadable left behind.
  try {
    const dir = await storeDir(RECENT_DIR, false);
    const names = [];
    for await (const [n] of dir.entries()) names.push(n);
    for (const n of names) await deleteIn(RECENT_DIR, n);
  } catch {
    /* no store yet */
  }
}

export function sanitizeName(name) {
  const s = String(name || "file");
  let out = "";
  const forbidden = "\\/:*?\"<>|";
  for (const ch of s) {
    const code = ch.codePointAt(0);
    out += (code < 32 || forbidden.includes(ch)) ? "_" : ch;
  }
  out = out.replace(/\s+/g, " ").trim();
  if (out.length > 120) {
    // Keep the extension when shortening, so the format stays recognisable.
    const m = /(\.[A-Za-z0-9]{1,10})$/.exec(out);
    const ext = m ? m[1] : "";
    out = out.slice(0, 120 - ext.length).trimEnd() + ext;
  }
  return out || "file";
}
