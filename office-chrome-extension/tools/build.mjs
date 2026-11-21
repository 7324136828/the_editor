#!/usr/bin/env node
/**
 * Cross-platform packager for the extension (no npm dependencies; uses the
 * vendored JSZip). Ships the same runtime files as build.ps1.
 *
 *   node tools/build.mjs                  dist/extension.zip   (Chrome Web Store upload / sideload)
 *   node tools/build.mjs --crx            + dist/extension.crx (signed CRX3 packed extension)
 *   node tools/build.mjs --verify FILE    check a .crx signature and print its extension ID
 *
 * Options:
 *   --out DIR     output directory (default: dist)
 *   --key FILE    PEM private key used to sign the .crx. Default: the
 *                 CRX_PRIVATE_KEY environment variable (PEM text), else key.pem
 *                 in this folder. If none exists, a new key is generated into
 *                 key.pem (git-ignored, never packaged). Keep that key: the
 *                 extension ID is derived from it.
 *
 * Also writes DIR/build-info.json (name, version, extension ID, SHA-256s).
 */

import fs from "node:fs";
import path from "node:path";
import crypto from "node:crypto";
import { fileURLToPath } from "node:url";
import { packCrx, verifyCrx, generatePrivateKeyPem, loadPrivateKey } from "./crx.mjs";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const JSZip = loadVendoredJSZip();

/** vendor/jszip.min.js is a UMD bundle; package.json "type": "module" stops require() from treating it as CommonJS. */
function loadVendoredJSZip() {
  const source = fs.readFileSync(path.join(ROOT, "vendor", "jszip.min.js"), "utf8");
  const module = { exports: {} };
  new Function("module", "exports", "define", source)(module, module.exports, undefined);
  if (typeof module.exports !== "function") throw new Error("vendor/jszip.min.js did not export JSZip");
  return module.exports;
}

/** Runtime files only; tests, tools and package.json stay out (same list as build.ps1). */
export const PACKAGE_ITEMS = ["manifest.json", "icons", "background", "content", "lib", "offscreen", "options", "popup", "viewer", "vendor"];
const JUNK = new Set([".DS_Store", "Thumbs.db", "desktop.ini"]);
// Fixed timestamp so identical sources produce an identical zip.
const ZIP_DATE = new Date(Date.UTC(2020, 0, 1, 0, 0, 0));

function parseArgs(argv) {
  const opts = { out: "dist", crx: false, key: null, verify: null };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    const next = () => {
      if (i + 1 >= argv.length) throw new Error(`${a} needs a value`);
      return argv[++i];
    };
    if (a === "--crx") opts.crx = true;
    else if (a === "--out") opts.out = next();
    else if (a === "--key") opts.key = next();
    else if (a === "--verify") opts.verify = next();
    else if (a === "-h" || a === "--help") opts.help = true;
    else throw new Error(`unknown argument: ${a}`);
  }
  return opts;
}

function listFiles(rel) {
  const abs = path.join(ROOT, rel);
  if (!fs.existsSync(abs)) return [];
  if (fs.statSync(abs).isFile()) return [rel];
  const out = [];
  for (const entry of fs.readdirSync(abs, { withFileTypes: true })) {
    if (JUNK.has(entry.name)) continue;
    const child = path.posix.join(rel, entry.name);
    if (entry.isDirectory()) out.push(...listFiles(child));
    else if (entry.isFile()) out.push(child);
  }
  return out;
}

/** Files the manifest points at; every one must be in the package. */
export function manifestReferences(manifest) {
  const refs = new Set();
  const addIcons = (icons) => {
    if (icons && typeof icons === "object") Object.values(icons).forEach((v) => refs.add(v));
    else if (typeof icons === "string") refs.add(icons);
  };
  addIcons(manifest.icons);
  addIcons(manifest.action?.default_icon);
  if (manifest.action?.default_popup) refs.add(manifest.action.default_popup);
  if (manifest.options_ui?.page) refs.add(manifest.options_ui.page);
  if (manifest.options_page) refs.add(manifest.options_page);
  if (manifest.background?.service_worker) refs.add(manifest.background.service_worker);
  for (const h of manifest.file_handlers || []) if (h.action) refs.add(h.action);
  for (const cs of manifest.content_scripts || []) [...(cs.js || []), ...(cs.css || [])].forEach((f) => refs.add(f));
  for (const war of manifest.web_accessible_resources || [])
    for (const r of war.resources || []) if (!r.includes("*")) refs.add(r);
  return [...refs].map((r) => r.replace(/^\//, ""));
}

export function readManifest() {
  const manifest = JSON.parse(fs.readFileSync(path.join(ROOT, "manifest.json"), "utf8"));
  if (manifest.manifest_version !== 3) throw new Error(`manifest_version must be 3, got ${manifest.manifest_version}`);
  if (!/^\d+(\.\d+){0,3}$/.test(manifest.version || "")) throw new Error(`invalid manifest version "${manifest.version}"`);
  const pkg = JSON.parse(fs.readFileSync(path.join(ROOT, "package.json"), "utf8"));
  if (pkg.version !== manifest.version)
    throw new Error(`package.json version ${pkg.version} does not match manifest.json version ${manifest.version}`);
  return manifest;
}

export async function buildZip() {
  const manifest = readManifest();
  const files = PACKAGE_ITEMS.flatMap(listFiles).sort();
  if (!files.length) throw new Error("nothing to package");
  const missing = manifestReferences(manifest).filter((r) => !files.includes(r));
  if (missing.length) throw new Error(`manifest references files missing from the package: ${missing.join(", ")}`);
  const zip = new JSZip();
  for (const f of files) zip.file(f, fs.readFileSync(path.join(ROOT, f)), { date: ZIP_DATE, createFolders: false });
  const bytes = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE", compressionOptions: { level: 9 } });
  return { manifest, files, bytes };
}

function resolveKey(keyPath) {
  if (keyPath) return { pem: fs.readFileSync(path.resolve(keyPath), "utf8"), source: keyPath };
  const env = process.env.CRX_PRIVATE_KEY;
  if (env && env.trim()) return { pem: env.replace(/\\n/g, "\n"), source: "CRX_PRIVATE_KEY" };
  const local = path.join(ROOT, "key.pem");
  if (fs.existsSync(local)) return { pem: fs.readFileSync(local, "utf8"), source: local };
  const pem = generatePrivateKeyPem();
  fs.writeFileSync(local, pem, { mode: 0o600 });
  console.warn(`No signing key given: generated ${local}. Keep it to publish updates under the same extension ID.`);
  return { pem, source: `${local} (new)` };
}

const sha256 = (buf) => crypto.createHash("sha256").update(buf).digest("hex");
const kb = (n) => `${Math.round(n / 1024)} KB`;

async function main() {
  const opts = parseArgs(process.argv.slice(2));
  if (opts.help) {
    console.log(fs.readFileSync(fileURLToPath(import.meta.url), "utf8").split("*/")[0]);
    return;
  }
  if (opts.verify) {
    const { id, zip } = verifyCrx(fs.readFileSync(opts.verify));
    console.log(`${opts.verify}: valid CRX3, extension ID ${id}, ${kb(zip.length)} zip payload`);
    return;
  }

  const out = path.resolve(ROOT, opts.out);
  fs.mkdirSync(out, { recursive: true });
  const { manifest, files, bytes } = await buildZip();
  const zipPath = path.join(out, "extension.zip");
  fs.writeFileSync(zipPath, bytes);
  console.log(`Wrote ${zipPath} (${files.length} files, ${kb(bytes.length)})`);

  const info = { name: manifest.name, version: manifest.version, zip: { file: "extension.zip", sha256: sha256(bytes) } };
  const crxPath = path.join(out, "extension.crx");
  if (opts.crx) {
    const { pem, source } = resolveKey(opts.key);
    const { crx, id } = packCrx(bytes, loadPrivateKey(pem));
    verifyCrx(crx); // never ship a file Chrome would reject
    fs.writeFileSync(crxPath, crx);
    info.id = id;
    info.crx = { file: "extension.crx", sha256: sha256(crx) };
    console.log(`Wrote ${crxPath} (${kb(crx.length)}), signed with ${source}`);
    console.log(`Extension ID: ${id}`);
  } else if (fs.existsSync(crxPath)) {
    fs.rmSync(crxPath); // don't leave a stale .crx next to a fresh zip
  }
  fs.writeFileSync(path.join(out, "build-info.json"), JSON.stringify(info, null, 2) + "\n");
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  main().catch((err) => {
    console.error(`build failed: ${err.message}`);
    process.exit(1);
  });
}
