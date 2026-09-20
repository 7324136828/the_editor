import type {
  WordDocumentModel,
  WordHeading,
  WordParagraph,
  WordComment,
} from "../types/office";
import {
  W,
  elements,
  first,
  children,
  attr,
  numeric,
  officePackage,
  xmlPart,
} from "./openXml";

function textOf(element: Element): string {
  return children(element)
    .map((child) => {
      if (child.namespaceURI === W && child.localName === "t")
        return child.textContent || "";
      if (child.namespaceURI === W && child.localName === "tab") return "\t";
      if (child.namespaceURI === W && ["br", "cr"].includes(child.localName))
        return "\n";
      return textOf(child);
    })
    .join("");
}
const enabled = (element: Element | undefined) =>
  !!element && !["0", "false", "off"].includes(attr(element, W, "val"));

export async function parseDocxFile(file: File): Promise<WordDocumentModel> {
  const zip = await officePackage(file);
  const document = (await xmlPart(zip, "word/document.xml", true))!;
  const body = first(document, W, "body");
  if (!body) throw new Error("The Word document has no document body.");
  const bodyChildren = children(body);
  if (bodyChildren.length > 10000)
    throw new Error(
      "This document exceeds the 10,000 paragraph/table import limit.",
    );
  const paragraphs: WordParagraph[] = [];
  const headings: WordHeading[] = [];
  const commentAnchors = new Map<string, string>();
  let page = 1;
  for (const [index, element] of bodyChildren.entries()) {
    if (element.namespaceURI !== W) continue;
    if (element.localName === "tbl") {
      const rows = children(element)
        .filter((child) => child.localName === "tr")
        .map((row) =>
          children(row)
            .filter((child) => child.localName === "tc")
            .map((cell) => elements(cell, W, "p").map(textOf).join("\n")),
        );
      if (rows.length > 1000 || rows.some((row) => row.length > 100))
        throw new Error("Tables are limited to 1,000 rows and 100 columns.");
      paragraphs.push({ id: `p-${index}`, type: "table", tableData: rows });
    } else if (element.localName === "p") {
      page += elements(element, W, "br").filter(
        (br) => attr(br, W, "type") === "page",
      ).length;
      const text = textOf(element);
      const properties = first(element, W, "pPr");
      const style = attr(first(properties, W, "pStyle"), W, "val");
      const headingLevel = /^(?:heading\s*)?([1-3])$/i.exec(style)?.[1];
      const alignment = attr(first(properties, W, "jc"), W, "val");
      const align =
        alignment === "both"
          ? "justify"
          : ["left", "center", "right", "justify"].includes(alignment)
            ? (alignment as WordParagraph["align"])
            : undefined;
      const runs = elements(element, W, "r")
        .map((run) => {
          const properties = first(run, W, "rPr");
          const fill = attr(first(properties, W, "shd"), W, "fill");
          return {
            text: textOf(run),
            bold: enabled(first(properties, W, "b")),
            italic: enabled(first(properties, W, "i")),
            ...(fill && /^[a-f\d]{6}$/i.test(fill)
              ? { highlight: `#${fill}` }
              : {}),
          };
        })
        .filter((run) => run.text.length);
      const type = headingLevel
        ? (`heading-${headingLevel}` as WordParagraph["type"])
        : first(properties, W, "pBdr")
          ? "quote"
          : first(properties, W, "shd")
            ? "callout"
            : "body";
      paragraphs.push({ id: `p-${index}`, type, text, runs, align });
      if (headingLevel)
        headings.push({
          id: `p-${index}`,
          level: Number(headingLevel) as 1 | 2 | 3,
          text,
          pageIndex: page,
        });
      for (const reference of [
        ...elements(element, W, "commentReference"),
        ...elements(element, W, "commentRangeStart"),
      ])
        commentAnchors.set(attr(reference, W, "id"), text);
    }
  }
  const smartArtXml = await xmlPart(zip, "customXml/officeStudioSmartArt.xml");
  if (smartArtXml) {
    try {
      const metadata = JSON.parse(
        smartArtXml.documentElement.textContent || "",
      );
      if (metadata?.version === 1 && Array.isArray(metadata.blocks)) {
        for (const block of metadata.blocks.slice(0, 1000)) {
          const diagram = block?.smartArt;
          if (
            !Number.isInteger(block?.index) ||
            block.index < 0 ||
            block.index >= paragraphs.length ||
            typeof block.id !== "string" ||
            !diagram ||
            typeof diagram.title !== "string" ||
            !["process", "cycle", "hierarchy", "pyramid"].includes(
              diagram.layout,
            ) ||
            !/^#[a-f\d]{6}$/i.test(diagram.accentColor) ||
            !Array.isArray(diagram.items) ||
            diagram.items.length < 1 ||
            diagram.items.length > 12 ||
            diagram.items.some(
              (item: unknown) =>
                !item ||
                typeof item !== "object" ||
                typeof (item as { id?: unknown }).id !== "string" ||
                typeof (item as { text?: unknown }).text !== "string",
            )
          )
            continue;
          paragraphs[block.index] = {
            id: block.id,
            type: "smartart",
            smartArt: diagram,
          };
        }
      }
    } catch {
      // Invalid optional application metadata does not make the DOCX unreadable.
    }
  }
  if (!paragraphs.length)
    paragraphs.push({ id: "p-empty", type: "body", text: "" });
  const styleDocument = await xmlPart(zip, "word/styles.xml");
  const defaults = styleDocument
    ? first(styleDocument, W, "docDefaults")
    : undefined;
  const sections = elements(document, W, "sectPr");
  const section = sections[sections.length - 1];
  const pageSize = first(section, W, "pgSz");
  const landscape =
    attr(pageSize, W, "orient") === "landscape" ||
    numeric(attr(pageSize, W, "w"), 12240) >
      numeric(attr(pageSize, W, "h"), 15840);
  const margin = numeric(attr(first(section, W, "pgMar"), W, "left"), 1440);
  const fontSize = numeric(attr(first(defaults, W, "sz"), W, "val"), 21) / 1.5;
  const lineHeight =
    numeric(attr(first(defaults, W, "spacing"), W, "line"), 384) / 240;
  const commentsXml = await xmlPart(zip, "word/comments.xml");
  const comments: WordComment[] = commentsXml
    ? elements(commentsXml, W, "comment").map((comment) => ({
        id: `comment-${attr(comment, W, "id")}`,
        author: attr(comment, W, "author") || "Unknown author",
        avatarColor: "#007acc",
        timestamp: attr(comment, W, "date") || "",
        selectedText: commentAnchors.get(attr(comment, W, "id")) || "",
        comment: elements(comment, W, "p").map(textOf).join("\n"),
        resolved: false,
      }))
    : [];
  const allText = paragraphs
    .map((paragraph) =>
      paragraph.type === "table"
        ? paragraph.tableData?.map((row) => row.join(" ")).join("\n") || ""
        : paragraph.type === "smartart"
          ? [
              paragraph.smartArt?.title ?? "",
              ...(paragraph.smartArt?.items ?? []).map((item) => item.text),
            ].join(" ")
          : paragraph.text || "",
    )
    .join("\n");
  const words = allText.trim() ? allText.trim().split(/\s+/).length : 0;
  const pages = Math.max(page, Math.ceil(words / 350), 1);
  const textColor = attr(first(defaults, W, "color"), W, "val");
  return {
    id: `doc-word-${crypto.randomUUID()}`,
    title: file.name,
    author: "Imported document",
    createdAt: new Date().toISOString(),
    modifiedAt: new Date().toISOString(),
    paragraphs,
    headings,
    comments,
    sections: [
      {
        id: "section-1",
        title: "Main document",
        pageNumber: 1,
        orientation: landscape ? "landscape" : "portrait",
      },
    ],
    stats: {
      words,
      characters: allText.length,
      charactersNoSpaces: allText.replace(/\s/g, "").length,
      paragraphs: paragraphs.length,
      pages,
      readingTimeMinutes: Math.ceil(words / 220),
      readabilityScore: "Not evaluated",
    },
    typography: {
      fontFamily: attr(first(defaults, W, "rFonts"), W, "ascii") || "Aptos",
      fontSize: Math.max(6, Math.min(96, fontSize)),
      lineHeight: Math.max(0.8, Math.min(3, lineHeight)),
      marginSize: margin <= 900 ? "narrow" : margin >= 1800 ? "wide" : "normal",
      textColor: /^[a-f\d]{6}$/i.test(textColor) ? `#${textColor}` : "#24292f",
    },
  };
}
