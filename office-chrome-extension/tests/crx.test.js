// Packaging: CRX3 signing/verification (tools/crx.mjs) and the zip builder (tools/build.mjs).
import test from "node:test";
import assert from "node:assert/strict";
import crypto from "node:crypto";
import {
  packCrx,
  verifyCrx,
  generatePrivateKeyPem,
  loadPrivateKey,
  extensionIdFromPrivateKey,
  extensionIdFromPublicKey,
  manifestKeyFromPrivateKey,
} from "../tools/crx.mjs";
import { buildZip, manifestReferences, PACKAGE_ITEMS } from "../tools/build.mjs";

const PEM = generatePrivateKeyPem();
const ZIP = Buffer.concat([Buffer.from([0x50, 0x4b, 0x03, 0x04]), Buffer.from("fake zip payload")]);

test("packCrx writes a CRX3 file that verifies and round-trips the zip", () => {
  const { crx, id } = packCrx(ZIP, PEM);
  assert.equal(crx.subarray(0, 4).toString("latin1"), "Cr24");
  assert.equal(crx.readUInt32LE(4), 3);
  assert.match(id, /^[a-p]{32}$/);
  const v = verifyCrx(crx);
  assert.equal(v.id, id);
  assert.deepEqual(v.zip, ZIP);
  assert.equal(id, extensionIdFromPrivateKey(PEM));
});

test("extension ID is the a-p encoding of the first 16 bytes of SHA-256(SPKI)", () => {
  const der = Buffer.from(manifestKeyFromPrivateKey(PEM), "base64");
  const hex = crypto.createHash("sha256").update(der).digest("hex").slice(0, 32);
  const expected = [...hex].map((h) => "abcdefghijklmnop"[parseInt(h, 16)]).join("");
  assert.equal(extensionIdFromPublicKey(der), expected);
});

test("PKCS#1 and PKCS#8 PEM keys give the same extension ID", () => {
  const pkcs1 = loadPrivateKey(PEM).export({ type: "pkcs1", format: "pem" });
  assert.match(pkcs1, /BEGIN RSA PRIVATE KEY/);
  assert.equal(extensionIdFromPrivateKey(pkcs1), extensionIdFromPrivateKey(PEM));
});

test("verifyCrx rejects tampered payloads, headers and non-CRX input", () => {
  const { crx } = packCrx(ZIP, PEM);
  const payload = Buffer.from(crx);
  payload[payload.length - 1] ^= 0xff;
  assert.throws(() => verifyCrx(payload), /signature does not verify/);

  const header = Buffer.from(crx);
  header[crx.length - ZIP.length - 1] ^= 0x01; // last byte of the signed crx_id
  assert.throws(() => verifyCrx(header), /signature does not verify|crx_id/);

  const v2 = Buffer.from(crx);
  v2.writeUInt32LE(2, 4);
  assert.throws(() => verifyCrx(v2), /unsupported CRX version/);
  assert.throws(() => verifyCrx(ZIP), /bad magic/);
  const short = Buffer.from(crx.subarray(0, 12));
  assert.throws(() => verifyCrx(short), /header length exceeds/);
});

test("non-RSA keys are refused", () => {
  const { privateKey } = crypto.generateKeyPairSync("ec", { namedCurve: "P-256" });
  assert.throws(() => loadPrivateKey(privateKey.export({ type: "pkcs8", format: "pem" })), /must be RSA/);
});

test("manifestReferences lists every file the manifest needs", () => {
  const refs = manifestReferences({
    icons: { 16: "icons/a.png" },
    action: { default_popup: "popup/p.html", default_icon: "icons/b.png" },
    options_ui: { page: "options/o.html" },
    background: { service_worker: "background/sw.js" },
    file_handlers: [{ action: "/viewer/v.html" }],
    content_scripts: [{ js: ["content/c.js"], css: ["content/c.css"] }],
    web_accessible_resources: [{ resources: ["viewer/*", "lib/x.js"] }],
  });
  assert.deepEqual(refs.sort(), [
    "background/sw.js", "content/c.css", "content/c.js", "icons/a.png", "icons/b.png",
    "lib/x.js", "options/o.html", "popup/p.html", "viewer/v.html",
  ]);
});

test("buildZip packages only runtime files, deterministically, with everything the manifest references", async () => {
  const first = await buildZip();
  const second = await buildZip();
  assert.ok(first.bytes.equals(second.bytes), "zip output is reproducible");
  assert.ok(first.files.includes("manifest.json"));
  for (const f of first.files) assert.ok(PACKAGE_ITEMS.includes(f.split("/")[0]), `unexpected file in package: ${f}`);
  assert.ok(!first.files.some((f) => f.startsWith("tests/") || f.startsWith("tools/") || f === "package.json"));
  for (const ref of manifestReferences(first.manifest)) assert.ok(first.files.includes(ref), `missing ${ref}`);
  const { crx } = packCrx(first.bytes, PEM);
  assert.deepEqual(verifyCrx(crx).zip, first.bytes);
});
