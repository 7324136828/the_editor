// Tests for the browser-integration plumbing: content/policy.js (pure
// decisions), content/content.js (click/drop handling, run in a vm with a
// fake DOM + chrome), lib/handoff.js (OPFS stores, with an in-memory OPFS).

import { test } from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import vm from "node:vm";
import { fileURLToPath } from "node:url";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(HERE, "..");
const EXAMPLE = path.join(ROOT, "..", "example");

await import("../content/policy.js");
const P = globalThis.OfficeExtPolicy;

const handoff = await import("../lib/handoff.js");

// ---------------------------------------------------------------------------
// in-memory OPFS
// ---------------------------------------------------------------------------

function notFound(name) {
  const e = new Error(`${name} not found`);
  e.name = "NotFoundError";
  return e;
}

class MemFileHandle {
  constructor(name) {
    this.kind = "file";
    this.name = name;
    this.data = new Uint8Array(0);
    this.lastModified = Date.now();
  }
  async getFile() {
    const d = this.data;
    return {
      size: d.length,
      lastModified: this.lastModified,
      text: async () => new TextDecoder().decode(d),
      arrayBuffer: async () => d.slice().buffer,
    };
  }
  async createWritable() {
    // Like real OPFS: writes go to a swap buffer, visible only after close().
    const parts = [];
    return {
      write: async (x) => {
        parts.push(typeof x === "string" ? new TextEncoder().encode(x) : new Uint8Array(x.buffer ? x.buffer.slice(x.byteOffset, x.byteOffset + x.byteLength) : x));
      },
      close: async () => {
        const n = parts.reduce((a, p) => a + p.length, 0);
        const out = new Uint8Array(n);
        let o = 0;
        for (const p of parts) { out.set(p, o); o += p.length; }
        this.data = out;
        this.lastModified = Date.now();
      },
      abort: async () => {},
    };
  }
}

class MemDirHandle {
  constructor(name = "") {
    this.kind = "directory";
    this.name = name;
    this.children = new Map();
  }
  async getDirectoryHandle(name, { create = false } = {}) {
    let c = this.children.get(name);
    if (!c) {
      if (!create) throw notFound(name);
      c = new MemDirHandle(name);
      this.children.set(name, c);
    }
    if (c.kind !== "directory") throw new TypeError("not a directory");
    return c;
  }
  async getFileHandle(name, { create = false } = {}) {
    let c = this.children.get(name);
    if (!c) {
      if (!create) throw notFound(name);
      c = new MemFileHandle(name);
      this.children.set(name, c);
    }
    if (c.kind !== "file") throw new TypeError("not a file");
    return c;
  }
  async removeEntry(name) {
    if (!this.children.delete(name)) throw notFound(name);
  }
  async *entries() {
    for (const [k, v] of [...this.children]) yield [k, v];
  }
}

function freshStore() {
  const root = new MemDirHandle();
  handoff.__setStorageRootForTests(async () => root);
  return root;
}

// ---------------------------------------------------------------------------
// policy: links
// ---------------------------------------------------------------------------

const ON = { intercept_downloads: true };

test("linkEligibility: binary Office links are intercepted by default", () => {
  for (const ext of ["docx", "doc", "rtf", "odt", "xlsx", "xls", "xlsm", "ods", "DOCX"]) {
    assert.equal(P.linkEligibility({ href: `https://a.test/f/report.${ext}` }, ON), `https://a.test/f/report.${ext}`, ext);
  }
  assert.equal(P.linkEligibility({ href: "https://a.test/r.docx?download=1#x" }, ON), "https://a.test/r.docx?download=1#x");
  assert.equal(P.linkEligibility({ href: "files/r%20q.xlsx", baseUrl: "https://a.test/dir/page" }, ON), "https://a.test/dir/files/r%20q.xlsx");
});

test("linkEligibility: text formats only when opted in (GitHub README.md stays a page)", () => {
  for (const href of ["https://github.com/o/r/blob/main/README.md", "https://a.test/robots.txt", "https://a.test/d.csv", "https://a.test/d.tsv"]) {
    assert.equal(P.linkEligibility({ href }, ON), null, href);
    assert.equal(P.linkEligibility({ href }, { ...ON, intercept_text: true }), href, href);
  }
});

test("linkEligibility: respects the setting, download attribute, target and scheme", () => {
  const href = "https://a.test/r.docx";
  assert.equal(P.linkEligibility({ href }, { intercept_downloads: false }), null);
  assert.equal(P.linkEligibility({ href }, {}), href, "default is on");
  assert.equal(P.linkEligibility({ href, download: "" }, ON), null, "download attribute");
  assert.equal(P.linkEligibility({ href, download: "x.docx" }, ON), null);
  assert.equal(P.linkEligibility({ href, target: "_blank" }, ON), null);
  assert.equal(P.linkEligibility({ href, target: "_self" }, ON), href);
  assert.equal(P.linkEligibility({ href: "ftp://a.test/r.docx" }, ON), null);
  assert.equal(P.linkEligibility({ href: "javascript:void('x.docx')" }, ON), null);
  assert.equal(P.linkEligibility({ href: "https://a.test/report.docx/edit" }, ON), null);
  assert.equal(P.linkEligibility({ href: "https://a.test/page.html" }, ON), null);
  assert.equal(P.linkEligibility({ href: "::::" }, ON), null);
});

test("isPlainClick: only unmodified primary clicks", () => {
  const base = { button: 0, ctrlKey: false, metaKey: false, shiftKey: false, altKey: false };
  assert.equal(P.isPlainClick(base), true);
  for (const k of ["ctrlKey", "metaKey", "shiftKey", "altKey"]) assert.equal(P.isPlainClick({ ...base, [k]: true }), false, k);
  assert.equal(P.isPlainClick({ ...base, button: 1 }), false);
});

test("nameFromUrl decodes the last segment and tolerates bad escapes", () => {
  assert.equal(P.nameFromUrl("https://a.test/x/Q3%20Report.docx?x=1"), "Q3 Report.docx");
  assert.equal(P.nameFromUrl("https://a.test/x/bad%E0%A4.docx"), "bad%E0%A4.docx");
  assert.equal(P.nameFromUrl("https://a.test/"), "");
});

// ---------------------------------------------------------------------------
// policy: downloads
// ---------------------------------------------------------------------------

test("downloadEligibility: intercepts office downloads, names from the suggested filename", () => {
  const v = P.downloadEligibility({ id: 1, url: "https://a.test/get?id=5", filename: "Budget 2025.xlsx", mime: "application/octet-stream" }, ON, { extensionId: "me" });
  assert.equal(v.intercept, true);
  assert.equal(v.name, "Budget 2025.xlsx");
  const w = P.downloadEligibility({ id: 2, url: "https://a.test/files/q.docx", filename: "", mime: "" }, ON, {});
  assert.equal(w.intercept, true);
  assert.equal(w.name, "q.docx");
  assert.equal(w.url, "https://a.test/files/q.docx");
});

test("downloadEligibility: never touches own, blob/data, html, bypassed, disabled or huge downloads", () => {
  const item = { id: 1, url: "https://a.test/q.docx", filename: "q.docx", mime: "" };
  assert.equal(P.downloadEligibility(item, { intercept_downloads: false }, {}).intercept, false);
  assert.equal(P.downloadEligibility({ ...item, byExtensionId: "me" }, ON, { extensionId: "me" }).reason, "own download");
  assert.equal(P.downloadEligibility({ ...item, url: "blob:chrome-extension://me/123" }, ON, {}).reason, "scheme");
  assert.equal(P.downloadEligibility({ ...item, url: "data:application/octet-stream;base64,AA" }, ON, {}).reason, "scheme");
  assert.equal(P.downloadEligibility({ ...item, mime: "text/html; charset=utf-8" }, ON, {}).reason, "html");
  assert.equal(P.downloadEligibility(item, ON, { isBypassed: (u) => u === item.url }).reason, "bypassed");
  assert.equal(P.downloadEligibility({ ...item, totalBytes: P.MAX_HANDOFF_BYTES + 1 }, ON, {}).reason, "too large");
  assert.equal(P.downloadEligibility({ ...item, filename: "q.csv", url: "https://a.test/q.csv" }, ON, {}).reason, "type");
  assert.equal(P.downloadEligibility({ ...item, filename: "q.csv", url: "https://a.test/q.csv" }, { ...ON, intercept_text: true }, {}).intercept, true);
  assert.equal(P.downloadEligibility({ ...item, filename: "setup.exe", url: "https://a.test/setup.exe" }, ON, {}).reason, "type");
});

test("downloadEligibility: file:// only with file access", () => {
  const item = { id: 1, url: "file:///C:/Users/me/report.docx", filename: "report.docx" };
  assert.equal(P.downloadEligibility(item, ON, { fileAccess: false }).intercept, false);
  assert.equal(P.downloadEligibility(item, ON, { fileAccess: true }).intercept, true);
});

// ---------------------------------------------------------------------------
// policy: fetched-content validation
// ---------------------------------------------------------------------------

const enc = (s) => new TextEncoder().encode(s);

test("validateFetched accepts real documents", () => {
  const docx = fs.readFileSync(path.join(EXAMPLE, "quarterly-report.docx"));
  const xlsx = fs.readFileSync(path.join(EXAMPLE, "sample-workbook.xlsx"));
  assert.deepEqual(P.validateFetched("quarterly-report.docx", "application/octet-stream", new Uint8Array(docx)), { ok: true });
  assert.deepEqual(P.validateFetched("x.xlsx", "", new Uint8Array(xlsx)), { ok: true });
  assert.equal(P.validateFetched("a.rtf", "application/rtf", enc("{\\rtf1 hi}")).ok, true);
  assert.equal(P.validateFetched("a.csv", "text/csv; charset=utf-8", enc("a,b\n1,2\n")).ok, true);
  assert.equal(P.validateFetched("a.xls", "", enc("a\tb\n1\t2\n")).ok, true, "tab-separated .xls export");
  assert.equal(P.validateFetched("a.xls", "text/html", enc("<html><body><table><tr><td>1</td></tr></table></body></html>")).ok, true);
});

test("validateFetched rejects sign-in pages, HTML viewers and mismatched bytes", () => {
  const login = enc("<!DOCTYPE html><html><head><title>Sign in</title></head><body>Sign in</body></html>");
  assert.equal(P.validateFetched("report.docx", "application/octet-stream", login).ok, false);
  assert.equal(P.validateFetched("README.md", "text/html; charset=utf-8", enc("<div>readme</div>")).ok, false);
  assert.equal(P.validateFetched("README.md", "text/plain", enc("\uFEFF<!-- c --><html>")).ok, false);
  assert.equal(P.validateFetched("report.docx", "", enc("not a zip at all")).ok, false);
  assert.equal(P.validateFetched("data.csv", "text/csv", new Uint8Array([1, 0, 2, 0])).ok, false);
  assert.equal(P.validateFetched("a.docx", "", new Uint8Array()).ok, false);
});

// ---------------------------------------------------------------------------
// policy: URL validation, rate limiting, drops, relative time, menus
// ---------------------------------------------------------------------------

test("validateUserUrl: http(s) only, helpful errors", () => {
  assert.deepEqual(P.validateUserUrl(" https://a.test/x.docx "), { ok: true, url: "https://a.test/x.docx" });
  assert.equal(P.validateUserUrl("example.com/a.xlsx").url, "https://example.com/a.xlsx");
  assert.equal(P.validateUserUrl("localhost:8000/a.docx").url, "http://localhost:8000/a.docx");
  assert.equal(P.validateUserUrl("http://intranet/a.docx").ok, true);
  assert.match(P.validateUserUrl("file:///C:/a.docx").error, /Open file/);
  assert.equal(P.validateUserUrl("javascript:alert(1)").ok, false);
  assert.equal(P.validateUserUrl("").ok, false);
  assert.equal(P.validateUserUrl("foo").ok, false);
  assert.match(P.validateUserUrl("report.docx").error, /file name/);
});

test("validateBackendUrl: normalises and flags non-loopback hosts", () => {
  assert.deepEqual(P.validateBackendUrl("127.0.0.1:47831"), { ok: true, url: "http://127.0.0.1:47831", loopback: true });
  assert.deepEqual(P.validateBackendUrl("http://localhost:5000/"), { ok: true, url: "http://localhost:5000", loopback: true });
  assert.equal(P.validateBackendUrl("http://[::1]:47831").loopback, true);
  assert.deepEqual(P.validateBackendUrl("https://docs.example.com"), { ok: true, url: "https://docs.example.com", loopback: false });
  assert.equal(P.validateBackendUrl("ftp://x").ok, false);
  assert.equal(P.validateBackendUrl("http://u:p@127.0.0.1:1").ok, false);
  assert.equal(P.validateBackendUrl("http://127.0.0.1:1/?x").ok, false);
  assert.equal(P.validateBackendUrl("").ok, false);
  assert.equal(P.validateBackendUrl("http://exa mple").ok, false);
});

test("createRateLimiter: per-key spacing and a global window", () => {
  const rl = P.createRateLimiter({ max: 3, windowMs: 10000, perKeyMs: 2000 });
  assert.equal(rl.allow(1, 0), true);
  assert.equal(rl.allow(1, 500), false, "same tab too soon");
  assert.equal(rl.allow(2, 600), true);
  assert.equal(rl.allow(3, 700), true);
  assert.equal(rl.allow(4, 800), false, "global cap");
  assert.equal(rl.allow(1, 10001), true, "window slid");
  assert.equal(rl.allow(null, 10002), false, "window full again");
  assert.equal(rl.allow(null, 10601), true, "keyless events only count globally");
});

test("dragHasOfficeFile / fileIsOpenable / isPageDropTarget", () => {
  const docxMime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
  assert.equal(P.dragHasOfficeFile([{ kind: "file", type: docxMime }], {}), true);
  assert.equal(P.dragHasOfficeFile([{ kind: "file", type: "image/png" }], {}), false);
  assert.equal(P.dragHasOfficeFile([{ kind: "string", type: docxMime }], {}), false);
  assert.equal(P.dragHasOfficeFile([{ kind: "file", type: "text/csv" }], {}), false);
  assert.equal(P.dragHasOfficeFile([{ kind: "file", type: "text/csv" }], { intercept_text: true }), true);
  assert.equal(P.fileIsOpenable({ name: "a.xlsx" }, {}), true);
  assert.equal(P.fileIsOpenable({ name: "a.png" }, {}), false);
  assert.equal(P.fileIsOpenable({ name: "a.md" }, {}), false);
  assert.equal(P.isPageDropTarget([{ tagName: "TEXTAREA" }]), true);
  assert.equal(P.isPageDropTarget([{ tagName: "DIV", isContentEditable: true }]), true);
  assert.equal(P.isPageDropTarget([{ tagName: "SPAN" }, { tagName: "DIV", className: "file-dropzone" }]), true);
  assert.equal(P.isPageDropTarget([{ tagName: "DIV", id: "upload-area" }]), true);
  assert.equal(P.isPageDropTarget([{ tagName: "P" }, { tagName: "BODY" }, { tagName: "HTML" }, {}]), false);
  assert.equal(P.isPageDropTarget([{ tagName: "svg", className: { baseVal: "x" } }]), false);
});

test("relativeTime", () => {
  const now = Date.UTC(2026, 10, 21, 12);
  assert.equal(P.relativeTime(now - 10 * 1000, now), "just now");
  assert.equal(P.relativeTime(now - 5 * 60 * 1000, now), "5 min ago");
  assert.equal(P.relativeTime(now - 3 * 3600 * 1000, now), "3 h ago");
  assert.equal(P.relativeTime(now - 30 * 3600 * 1000, now), "yesterday");
  assert.equal(P.relativeTime(now - 4 * 86400 * 1000, now), "4 days ago");
  assert.ok(P.relativeTime(now - 40 * 86400 * 1000, now).length > 0);
  assert.equal(P.relativeTime(0, now), "");
});

test("context-menu patterns are well-formed match patterns", () => {
  const re = /^(\*|https?|file):\/\/(\*|[^/*]*)?\/.*$/;
  for (const p of [...P.linkMenuPatterns(), ...P.pageMenuPatterns()]) assert.match(p, re, p);
  assert.ok(P.linkMenuPatterns().includes("*://*/*.docx?*"));
});

// ---------------------------------------------------------------------------
// handoff / recents (in-memory OPFS)
// ---------------------------------------------------------------------------

test("handoff: put/get/delete round trip; ids carry their creation time", async () => {
  freshStore();
  const before = Date.now();
  const id = await handoff.putHandoff("Report.docx", new Uint8Array([1, 2, 3]), "application/x");
  assert.match(id, /^[a-z0-9]+-[a-z0-9]+$/);
  assert.ok(handoff.idTimestamp(id) >= before && handoff.idTimestamp(id) <= Date.now());
  const h = await handoff.getHandoff(id);
  assert.equal(h.name, "Report.docx");
  assert.equal(h.mime, "application/x");
  assert.deepEqual([...h.bytes], [1, 2, 3]);
  await handoff.deleteHandoff(id);
  await assert.rejects(handoff.getHandoff(id));
  await assert.rejects(handoff.getHandoff("../evil"), /bad store id/);
});

test("handoff: streamed writes are invisible until complete and survive a sweep", async () => {
  const root = freshStore();
  const w = await handoff.beginHandoff("big.xlsx", "");
  await w.write(new Uint8Array([1, 2]));
  // In flight: not listed, not readable, NOT swept (this was the PL-20 bug).
  assert.equal((await handoff.listHandoffs()).length, 0);
  await assert.rejects(handoff.getHandoff(w.id), /still being transferred/);
  assert.equal(await handoff.sweepHandoffs(), 0);
  // Even with meta.json missing entirely (crash between mkdir and meta write).
  const dir = await root.getDirectoryHandle("office_handoff");
  await dir.getDirectoryHandle(handoff.newStoreId(), { create: true });
  assert.equal(await handoff.sweepHandoffs(), 0);
  await w.write(new Uint8Array([3]));
  const id = await w.close();
  const h = await handoff.getHandoff(id);
  assert.deepEqual([...h.bytes], [1, 2, 3]);
  assert.equal(h.size, 3);
  // Old entries go once they are past the max age.
  assert.equal(await handoff.sweepHandoffs(handoff.HANDOFF_MAX_AGE_MS, Date.now() + handoff.HANDOFF_MAX_AGE_MS + 1000), 2);
  assert.equal(dir.children.size, 0);
});

test("handoff: sweep removes foreign entries without a recoverable age", async () => {
  const root = freshStore();
  const dir = await root.getDirectoryHandle("office_handoff", { create: true });
  await dir.getDirectoryHandle("garbage", { create: true });
  await handoff.putHandoff("a.docx", new Uint8Array([1]), "");
  assert.equal(await handoff.sweepHandoffs(), 1);
  assert.equal(dir.children.size, 1);
});

test("handoff: size cap aborts and cleans up", async () => {
  const root = freshStore();
  const w = await handoff.beginHandoff("x.docx", "", 4);
  await w.write(new Uint8Array(3));
  await assert.rejects(w.write(new Uint8Array(3)), /larger than/);
  const dir = await root.getDirectoryHandle("office_handoff");
  assert.equal(dir.children.size, 0);
  await assert.rejects(w.close(), /already finished/);
});

test("recents: same name replaces the old entry (case-insensitive), newest first", async () => {
  freshStore();
  await handoff.addRecent("a.docx", new Uint8Array([1]), "");
  await new Promise((r) => setTimeout(r, 3));
  await handoff.addRecent("b.xlsx", new Uint8Array([2]), "");
  await new Promise((r) => setTimeout(r, 3));
  const id3 = await handoff.addRecent("A.docx", new Uint8Array([9, 9]), "");
  const list = await handoff.listRecent();
  assert.deepEqual(list.map((e) => e.name), ["A.docx", "b.xlsx"]);
  assert.equal(list[0].id, id3);
  assert.deepEqual([...(await handoff.getRecent(id3)).bytes], [9, 9], "stores the newest (saved) bytes");
  // updateRecent is the viewer's "after save" entry point: same semantics.
  await new Promise((r) => setTimeout(r, 3));
  const id4 = await handoff.updateRecent("b.xlsx", new Uint8Array([7]), "");
  const after = await handoff.listRecent();
  assert.deepEqual(after.map((e) => e.name), ["b.xlsx", "A.docx"]);
  assert.equal(after[0].id, id4);
});

test("recents: capped by count and by total size", async () => {
  freshStore();
  for (let i = 0; i < handoff.MAX_RECENT + 3; i++) {
    await handoff.addRecent(`f${i}.docx`, new Uint8Array([i]), "");
    await new Promise((r) => setTimeout(r, 2));
  }
  const list = await handoff.listRecent();
  assert.equal(list.length, handoff.MAX_RECENT);
  assert.equal(list[0].name, `f${handoff.MAX_RECENT + 2}.docx`);
  assert.ok(!list.some((e) => e.name === "f0.docx"));
  await handoff.clearRecents();
  assert.equal((await handoff.listRecent()).length, 0);
});

test("recents: concurrent adds of the same name keep exactly one entry", async () => {
  freshStore();
  await Promise.all([
    handoff.addRecent("same.docx", new Uint8Array([1]), ""),
    handoff.addRecent("same.docx", new Uint8Array([2]), ""),
    handoff.addRecent("same.docx", new Uint8Array([3]), ""),
  ]);
  const list = await handoff.listRecent();
  assert.equal(list.length, 1);
});

test("sanitizeName keeps the extension when shortening", () => {
  assert.equal(handoff.sanitizeName("a/b:c?.docx"), "a_b_c_.docx");
  const long = "x".repeat(200) + ".xlsx";
  const s = handoff.sanitizeName(long);
  assert.equal(s.length, 120);
  assert.ok(s.endsWith(".xlsx"));
  assert.equal(handoff.sanitizeName(""), "file");
});

// ---------------------------------------------------------------------------
// content script behaviour (vm with a fake DOM + chrome)
// ---------------------------------------------------------------------------

function fakeElement(tag = "div") {
  return {
    tagName: tag.toUpperCase(),
    style: {},
    isConnected: false,
    attachShadow() { return { append() {} }; },
    append() {},
    setAttribute() {},
    remove() { this.isConnected = false; },
  };
}

async function loadContentScript({ settings = {}, reply = async () => ({ ok: true }), runtimeId = "ext-id" } = {}) {
  const listeners = {};
  const sent = [];
  const assigned = [];
  const chrome = {
    runtime: {
      id: runtimeId,
      sendMessage: async (m) => {
        sent.push(m);
        return reply(m);
      },
    },
    storage: {
      local: { get: async () => ({ settings }) },
      onChanged: { addListener: (fn) => { listeners.__storage = fn; } },
    },
  };
  const ctx = {
    chrome,
    window: { addEventListener: (t, fn, cap) => { (listeners[t] ||= []).push({ fn, cap }); } },
    document: {
      baseURI: "https://site.test/docs/",
      documentElement: { appendChild(el) { el.isConnected = true; } },
      createElement: fakeElement,
    },
    location: { href: "https://site.test/docs/", assign: (u) => assigned.push(u) },
    setTimeout, clearTimeout, btoa, atob, URL, console, Promise, Uint8Array, Math,
  };
  vm.createContext(ctx);
  vm.runInContext(fs.readFileSync(path.join(ROOT, "content", "policy.js"), "utf8"), ctx);
  vm.runInContext(fs.readFileSync(path.join(ROOT, "content", "content.js"), "utf8"), ctx);
  await new Promise((r) => setTimeout(r, 0)); // settings load
  const fire = (type, ev) => { for (const l of listeners[type] || []) l.fn(ev); return ev; };
  return { ctx, chrome, listeners, sent, assigned, fire };
}

function link(href, attrs = {}) {
  return {
    tagName: "A",
    href,
    getAttribute: (n) => (n in attrs ? attrs[n] : null),
    hasAttribute: (n) => n in attrs,
  };
}

function clickOn(a, extra = {}) {
  return {
    type: "click", isTrusted: true, defaultPrevented: false, button: 0,
    ctrlKey: false, metaKey: false, shiftKey: false, altKey: false,
    prevented: false, stopped: false,
    composedPath: () => [{ tagName: "SPAN" }, a, { tagName: "BODY" }],
    preventDefault() { this.defaultPrevented = true; this.prevented = true; },
    stopPropagation() { this.stopped = true; },
    ...extra,
  };
}

const tick = () => new Promise((r) => setTimeout(r, 5));

test("content: plain click on a .docx link opens it (bubble phase, no stopPropagation)", async () => {
  const env = await loadContentScript({ settings: { intercept_downloads: true } });
  assert.equal(env.listeners.click[0].cap, false, "bubble phase");
  const ev = env.fire("click", clickOn(link("https://files.test/q3.docx")));
  assert.equal(ev.prevented, true);
  assert.equal(ev.stopped, false);
  await tick();
  assert.deepEqual(env.sent.map((m) => m.type), ["office-open-url"]);
  assert.equal(env.sent[0].url, "https://files.test/q3.docx");
  assert.deepEqual(env.assigned, []);
});

test("content: falls back to normal navigation when the editor cannot open it", async () => {
  const env = await loadContentScript({ reply: async () => ({ ok: false, fallback: true }) });
  env.fire("click", clickOn(link("https://files.test/q3.docx")));
  await tick();
  assert.deepEqual(env.assigned, ["https://files.test/q3.docx"]);
  const env2 = await loadContentScript({ reply: async () => { throw new Error("Extension context invalidated."); } });
  env2.fire("click", clickOn(link("https://files.test/q3.docx")));
  await tick();
  assert.deepEqual(env2.assigned, ["https://files.test/q3.docx"]);
});

test("content: leaves clicks alone when off, untrusted, handled, modified, or not a document", async () => {
  const off = await loadContentScript({ settings: { intercept_downloads: false } });
  assert.equal(off.fire("click", clickOn(link("https://f.test/a.docx"))).prevented, false);

  const env = await loadContentScript({ settings: {} });
  const a = link("https://f.test/a.docx");
  assert.equal(env.fire("click", clickOn(a, { isTrusted: false })).prevented, false, "synthetic click");
  assert.equal(env.fire("click", clickOn(a, { defaultPrevented: true })).prevented, false, "page handled it");
  assert.equal(env.fire("click", clickOn(link("https://f.test/a.docx", { download: "" }))).prevented, false);
  assert.equal(env.fire("click", clickOn(link("https://f.test/a.docx", { target: "_blank" }))).prevented, false);
  assert.equal(env.fire("click", clickOn(link("https://github.com/o/r/blob/main/README.md"))).prevented, false);
  assert.equal(env.fire("click", clickOn(link("https://f.test/page"))).prevented, false);
  const alt = env.fire("click", clickOn(a, { altKey: true }));
  assert.equal(alt.prevented, false, "Alt-click = download");
  await tick();
  assert.deepEqual(env.sent.map((m) => m.type), ["office-bypass-url"], "Alt-click keeps the real download");
  env.fire("contextmenu", { isTrusted: true, composedPath: () => [a] });
  env.fire("auxclick", { isTrusted: true, button: 1, composedPath: () => [a] });
  await tick();
  assert.deepEqual(env.sent.map((m) => m.type), ["office-bypass-url", "office-bypass-url", "office-bypass-url"]);
});

test("content: settings changes apply live", async () => {
  const env = await loadContentScript({ settings: { intercept_downloads: true } });
  env.listeners.__storage({ settings: { newValue: { intercept_downloads: false } } }, "local");
  assert.equal(env.fire("click", clickOn(link("https://f.test/a.docx"))).prevented, false);
  env.listeners.__storage({ settings: { newValue: { intercept_downloads: true, intercept_text: true } } }, "local");
  assert.equal(env.fire("click", clickOn(link("https://f.test/notes.md"))).prevented, true);
});

test("content: an orphaned script (extension updated) lets the browser navigate", async () => {
  const env = await loadContentScript({ settings: {} });
  delete env.chrome.runtime.id;
  const ev = env.fire("click", clickOn(link("https://f.test/a.docx")));
  assert.equal(ev.prevented, false);
});

function dragEvent(type, { files = [], items = [], path = [{ tagName: "BODY" }], prevented = false, trusted = true } = {}) {
  return {
    type, isTrusted: trusted, defaultPrevented: prevented, prevented: false,
    dataTransfer: { types: ["Files"], items, files, dropEffect: "none" },
    composedPath: () => path,
    preventDefault() { this.defaultPrevented = true; this.prevented = true; },
  };
}

function fakeFile(name, bytes) {
  const data = new Uint8Array(bytes);
  return {
    name, type: "", size: data.length,
    slice: (a, b) => ({ arrayBuffer: async () => data.slice(a, b).buffer }),
  };
}

const DOCX_ITEM = { kind: "file", type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" };

test("content: drops are ignored unless opted in (uploads keep working)", async () => {
  const env = await loadContentScript({ settings: {} });
  assert.equal(env.fire("dragover", dragEvent("dragover", { items: [DOCX_ITEM] })).prevented, false);
  assert.equal(env.fire("drop", dragEvent("drop", { files: [fakeFile("a.docx", [1])] })).prevented, false);
  await tick();
  assert.equal(env.sent.length, 0);
});

test("content: opted-in drops respect page drop zones and the page's own handling", async () => {
  const env = await loadContentScript({ settings: { intercept_drops: true } });
  const f = () => [fakeFile("a.docx", [1, 2, 3])];
  assert.equal(env.fire("dragover", dragEvent("dragover", { items: [DOCX_ITEM], path: [{ tagName: "TEXTAREA" }] })).prevented, false);
  assert.equal(env.fire("dragover", dragEvent("dragover", { items: [{ kind: "file", type: "image/png" }] })).prevented, false);
  assert.equal(env.fire("drop", dragEvent("drop", { files: f(), prevented: true })).prevented, false, "page handled it: untouched");
  assert.equal(env.fire("drop", dragEvent("drop", { files: f(), trusted: false })).prevented, false, "synthetic drop");
  assert.equal(env.fire("drop", dragEvent("drop", { files: f(), path: [{ tagName: "DIV", isContentEditable: true }] })).prevented, false);
  await tick();
  assert.equal(env.sent.length, 0);
  assert.equal(env.listeners.drop[0].cap, false, "bubble phase");
});

test("content: an accepted drop streams the file in chunks to the service worker", async () => {
  const env = await loadContentScript({
    settings: { intercept_drops: true },
    reply: async (m) => (m.type === "office-drop-begin" ? { ok: true, token: "t1" } : { ok: true }),
  });
  assert.equal(env.fire("dragover", dragEvent("dragover", { items: [DOCX_ITEM] })).prevented, true);
  const ev = env.fire("drop", dragEvent("drop", { files: [fakeFile("a.docx", [80, 75, 3, 4]), fakeFile("pic.png", [1])] }));
  assert.equal(ev.prevented, true);
  await tick();
  assert.deepEqual(env.sent.map((m) => m.type), ["office-drop-begin", "office-drop-chunk", "office-drop-end"]);
  assert.equal(env.sent[0].name, "a.docx");
  assert.equal(env.sent[0].size, 4);
  assert.equal(env.sent[1].seq, 0);
  assert.deepEqual([...Buffer.from(env.sent[1].b64, "base64")], [80, 75, 3, 4]);
});

// ---------------------------------------------------------------------------
// static checks on the manifest / worker
// ---------------------------------------------------------------------------

test("manifest: minimal permissions, optional web access, version 0.3.0", () => {
  const m = JSON.parse(fs.readFileSync(path.join(ROOT, "manifest.json"), "utf8"));
  assert.equal(m.version, "0.3.0");
  assert.ok(!m.permissions.includes("tabs"));
  assert.ok(m.permissions.includes("contextMenus"));
  assert.ok(!(m.host_permissions || []).some((h) => /^\*|^https?:\/\/\*/.test(h)), "no blanket host access at install");
  assert.deepEqual(m.optional_host_permissions, P.WEB_ORIGINS);
  assert.equal(m.content_scripts, undefined, "content script is registered only after opt-in");
  for (const f of ["content/policy.js", "content/content.js", "background/service-worker.js", "offscreen/offscreen.html", "popup/popup.html", "options/options.html"]) {
    assert.ok(fs.existsSync(path.join(ROOT, f)), f);
  }
});

test("service worker never re-intercepts its own downloads and uses onDeterminingFilename", () => {
  const src = fs.readFileSync(path.join(ROOT, "background", "service-worker.js"), "utf8");
  assert.match(src, /onDeterminingFilename/);
  assert.match(src, /byExtensionId === chrome\.runtime\.id/);
  assert.doesNotMatch(src, /downloads\.onCreated/);
});
