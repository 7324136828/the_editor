import JSZip from "jszip";

export const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
export const P = "http://schemas.openxmlformats.org/presentationml/2006/main";
export const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
export const R =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
export const elements = (
  parent: Document | Element,
  ns: string,
  local: string,
): Element[] => Array.from(parent.getElementsByTagNameNS(ns, local));
export const first = (
  parent: Document | Element | undefined | null,
  ns: string,
  local: string,
): Element | undefined => (parent ? elements(parent, ns, local)[0] : undefined);
export const children = (parent: Element): Element[] =>
  Array.from(parent.childNodes).filter(
    (node) => node.nodeType === 1,
  ) as Element[];
export const attr = (
  element: Element | undefined | null,
  ns: string,
  name: string,
): string => element?.getAttributeNS(ns, name) || "";
export const numeric = (
  value: string | null | undefined,
  fallback: number,
): number =>
  value !== undefined &&
  value !== null &&
  value !== "" &&
  Number.isFinite(Number(value))
    ? Number(value)
    : fallback;

export async function officePackage(file: File): Promise<JSZip> {
  if (file.size > 25 * 1024 * 1024)
    throw new Error("Office imports are limited to 25 MB.");
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(await file.arrayBuffer());
  } catch {
    throw new Error(
      "This file is not a readable Office Open XML package. Legacy .doc/.ppt and encrypted documents are not supported.",
    );
  }
  if (!zip.file("[Content_Types].xml") || Object.keys(zip.files).length > 10000)
    throw new Error("This Office package is invalid or too complex to import.");
  return zip;
}

export async function xmlPart(
  zip: JSZip,
  name: string,
  required = false,
): Promise<Document | undefined> {
  const part = zip.file(name);
  if (!part) {
    if (required)
      throw new Error(`The document is missing its required ${name} part.`);
    return undefined;
  }
  const text = await part.async("text");
  if (text.length > 20 * 1024 * 1024)
    throw new Error(`The ${name} part is too large to import.`);
  if (/<!DOCTYPE|<!ENTITY/i.test(text))
    throw new Error("Office XML with entity declarations is not supported.");
  const document = new DOMParser().parseFromString(text, "application/xml");
  if (
    document.getElementsByTagName("parsererror").length ||
    !document.documentElement
  )
    throw new Error(`The ${name} part contains invalid XML.`);
  return document;
}

export function resolvePart(source: string, target: string): string {
  if (/^[a-z]+:/i.test(target) || target.includes("\\"))
    throw new Error("External document relationships cannot be imported.");
  const parts = target.startsWith("/") ? [] : source.split("/").slice(0, -1);
  for (const part of target.split("/")) {
    if (!part || part === ".") continue;
    if (part === "..") {
      if (!parts.length)
        throw new Error("Invalid Office package relationship.");
      parts.pop();
    } else parts.push(part);
  }
  return parts.join("/");
}
