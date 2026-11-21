/**
 * CRX3 packing and verification (the "packed extension" format Chrome writes
 * for "Pack extension"), implemented with node:crypto only.
 *
 * File layout:
 *   "Cr24" | uint32le version (3) | uint32le header length | CrxFileHeader | zip
 *
 * CrxFileHeader (protobuf, components/crx_file/crx3.proto):
 *   repeated AsymmetricKeyProof sha256_with_rsa = 2;   // {1: public_key (DER SPKI), 2: signature}
 *   bytes signed_header_data                  = 10000; // SignedData {1: crx_id (16 bytes)}
 *
 * The RSA PKCS#1 v1.5 SHA-256 signature covers
 *   "CRX3 SignedData\0" | uint32le len(signed_header_data) | signed_header_data | zip
 *
 * The extension ID is the first 16 bytes of SHA-256(public key DER), written
 * as 32 characters a-p (one per nibble).
 */

import crypto from "node:crypto";

const MAGIC = Buffer.from("Cr24", "latin1");
const SIGNATURE_CONTEXT = Buffer.from("CRX3 SignedData\x00", "latin1");
const FIELD_SHA256_WITH_RSA = 2;
const FIELD_SIGNED_HEADER_DATA = 10000;

function varint(n) {
  const out = [];
  while (n > 0x7f) {
    out.push((n & 0x7f) | 0x80);
    n = Math.floor(n / 128);
  }
  out.push(n);
  return Buffer.from(out);
}

function lengthDelimited(field, bytes) {
  return Buffer.concat([varint(field * 8 + 2), varint(bytes.length), bytes]);
}

function uint32le(n) {
  const b = Buffer.alloc(4);
  b.writeUInt32LE(n);
  return b;
}

/** Parse the length-delimited (wire type 2) fields of a protobuf message. */
function parseMessage(buf) {
  const fields = [];
  let i = 0;
  const readVarint = () => {
    let result = 0;
    let mul = 1;
    for (;;) {
      if (i >= buf.length) throw new Error("truncated protobuf varint");
      const byte = buf[i++];
      result += (byte & 0x7f) * mul;
      if (!(byte & 0x80)) return result;
      mul *= 128;
      if (mul > 2 ** 49) throw new Error("protobuf varint too long");
    }
  };
  while (i < buf.length) {
    const tag = readVarint();
    const field = Math.floor(tag / 8);
    const wire = tag % 8;
    if (wire === 0) readVarint();
    else if (wire === 1) i += 8;
    else if (wire === 5) i += 4;
    else if (wire === 2) {
      const len = readVarint();
      if (i + len > buf.length) throw new Error("truncated protobuf field");
      fields.push({ field, bytes: buf.subarray(i, i + len) });
      i += len;
    } else throw new Error(`unsupported protobuf wire type ${wire}`);
    if (i > buf.length) throw new Error("truncated protobuf message");
  }
  return fields;
}

/** Load a PEM private key (PKCS#1 "RSA PRIVATE KEY" or PKCS#8 "PRIVATE KEY"). */
export function loadPrivateKey(pem) {
  const key = crypto.createPrivateKey({ key: pem, format: "pem" });
  if (key.asymmetricKeyType !== "rsa") throw new Error(`CRX signing key must be RSA, got ${key.asymmetricKeyType}`);
  return key;
}

/** New 2048-bit RSA key as PKCS#8 PEM (the format Chrome writes for "Pack extension"). */
export function generatePrivateKeyPem() {
  const { privateKey } = crypto.generateKeyPairSync("rsa", { modulusLength: 2048 });
  return privateKey.export({ type: "pkcs8", format: "pem" });
}

export function publicKeyDer(privateKey) {
  return crypto.createPublicKey(privateKey).export({ type: "spki", format: "der" });
}

function crxIdBytes(publicDer) {
  return crypto.createHash("sha256").update(publicDer).digest().subarray(0, 16);
}

/** Chrome extension ID (32 chars a-p) for a DER SubjectPublicKeyInfo. */
export function extensionIdFromPublicKey(publicDer) {
  return [...crxIdBytes(publicDer).toString("hex")]
    .map((h) => String.fromCharCode(97 + parseInt(h, 16)))
    .join("");
}

/** Extension ID for a private key (PEM string or KeyObject). */
export function extensionIdFromPrivateKey(privateKey) {
  const key = typeof privateKey === "string" ? loadPrivateKey(privateKey) : privateKey;
  return extensionIdFromPublicKey(publicKeyDer(key));
}

/** The manifest "key" value (base64 DER public key) that pins the extension ID. */
export function manifestKeyFromPrivateKey(privateKey) {
  const key = typeof privateKey === "string" ? loadPrivateKey(privateKey) : privateKey;
  return publicKeyDer(key).toString("base64");
}

function signedPayload(signedHeaderData, zip) {
  return Buffer.concat([SIGNATURE_CONTEXT, uint32le(signedHeaderData.length), signedHeaderData, zip]);
}

/**
 * Build a CRX3 file from zip bytes and an RSA private key.
 * Returns { crx: Buffer, id: string }.
 */
export function packCrx(zipBytes, privateKey) {
  const key = typeof privateKey === "string" ? loadPrivateKey(privateKey) : privateKey;
  const zip = Buffer.from(zipBytes);
  const pub = publicKeyDer(key);
  const signedHeaderData = lengthDelimited(1, crxIdBytes(pub));
  const signature = crypto.sign("sha256", signedPayload(signedHeaderData, zip), {
    key,
    padding: crypto.constants.RSA_PKCS1_PADDING,
  });
  const proof = Buffer.concat([lengthDelimited(1, pub), lengthDelimited(2, signature)]);
  const header = Buffer.concat([
    lengthDelimited(FIELD_SHA256_WITH_RSA, proof),
    lengthDelimited(FIELD_SIGNED_HEADER_DATA, signedHeaderData),
  ]);
  const crx = Buffer.concat([MAGIC, uint32le(3), uint32le(header.length), header, zip]);
  return { crx, id: extensionIdFromPublicKey(pub) };
}

/**
 * Parse and verify a CRX3 file the way Chrome does: the header must carry a
 * signed crx_id, and an RSA proof whose key hashes to that id must verify.
 * Returns { id, zip, publicKey }; throws on anything invalid.
 */
export function verifyCrx(crxBytes) {
  const crx = Buffer.from(crxBytes);
  if (crx.length < 12 || !crx.subarray(0, 4).equals(MAGIC)) throw new Error("not a CRX file (bad magic)");
  const version = crx.readUInt32LE(4);
  if (version !== 3) throw new Error(`unsupported CRX version ${version} (expected 3)`);
  const headerLen = crx.readUInt32LE(8);
  if (12 + headerLen > crx.length) throw new Error("CRX header length exceeds file size");
  const header = crx.subarray(12, 12 + headerLen);
  const zip = crx.subarray(12 + headerLen);
  if (zip.length < 4 || zip.readUInt32LE(0) !== 0x04034b50) throw new Error("CRX payload is not a zip archive");

  const fields = parseMessage(header);
  const signedHeaderData = fields.find((f) => f.field === FIELD_SIGNED_HEADER_DATA)?.bytes;
  if (!signedHeaderData) throw new Error("CRX header has no signed_header_data");
  const crxId = parseMessage(signedHeaderData).find((f) => f.field === 1)?.bytes;
  if (!crxId || crxId.length !== 16) throw new Error("CRX signed_header_data has no 16-byte crx_id");

  const payload = signedPayload(signedHeaderData, zip);
  const proofs = fields.filter((f) => f.field === FIELD_SHA256_WITH_RSA);
  if (!proofs.length) throw new Error("CRX header has no sha256_with_rsa proof");
  let developerKey = null;
  for (const proof of proofs) {
    const parts = parseMessage(proof.bytes);
    const pub = parts.find((p) => p.field === 1)?.bytes;
    const sig = parts.find((p) => p.field === 2)?.bytes;
    if (!pub || !sig) throw new Error("CRX proof is missing its public key or signature");
    const keyObj = crypto.createPublicKey({ key: Buffer.from(pub), format: "der", type: "spki" });
    const ok = crypto.verify("sha256", payload, { key: keyObj, padding: crypto.constants.RSA_PKCS1_PADDING }, sig);
    if (!ok) throw new Error("CRX signature does not verify");
    if (crxIdBytes(pub).equals(crxId)) developerKey = Buffer.from(pub);
  }
  if (!developerKey) throw new Error("no CRX proof matches the declared crx_id");
  return { id: extensionIdFromPublicKey(developerKey), zip: Buffer.from(zip), publicKey: developerKey };
}
