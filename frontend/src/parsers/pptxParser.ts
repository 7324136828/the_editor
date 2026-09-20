import type {
  PptDocumentModel,
  PptSlide,
  PptShapeObject,
} from "../types/office";
import {
  P,
  A,
  R,
  elements,
  first,
  attr,
  numeric,
  officePackage,
  xmlPart,
  resolvePart,
} from "./openXml";

const relationshipNS =
  "http://schemas.openxmlformats.org/package/2006/relationships";
const drawingColor = (
  element: Element | undefined,
  fallback: string,
): string => {
  const rgb = first(element, A, "srgbClr")?.getAttribute("val");
  return rgb && /^[a-f\d]{6}$/i.test(rgb) ? `#${rgb}` : fallback;
};
const textOf = (element: Element): string =>
  elements(element, A, "p")
    .map((paragraph) =>
      elements(paragraph, A, "t")
        .map((text) => text.textContent || "")
        .join(""),
    )
    .join("\n");

export async function parsePptxFile(file: File): Promise<PptDocumentModel> {
  const zip = await officePackage(file);
  const presentation = (await xmlPart(zip, "ppt/presentation.xml", true))!;
  const relationships = (await xmlPart(
    zip,
    "ppt/_rels/presentation.xml.rels",
    true,
  ))!;
  const slideIds = elements(presentation, P, "sldId");
  if (!slideIds.length) throw new Error("The presentation contains no slides.");
  if (slideIds.length > 200)
    throw new Error("Presentations are limited to 200 slides.");
  const slideSize = first(presentation, P, "sldSz");
  const xScale = 880 / numeric(slideSize?.getAttribute("cx"), 8382000);
  const yScale = 495 / numeric(slideSize?.getAttribute("cy"), 4714875);
  const relationMap = new Map(
    elements(relationships, relationshipNS, "Relationship").map((item) => [
      item.getAttribute("Id"),
      item,
    ]),
  );
  const slides: PptSlide[] = [];
  for (const [index, slideId] of slideIds.entries()) {
    const relationship = relationMap.get(attr(slideId, R, "id"));
    if (!relationship || relationship.getAttribute("TargetMode") === "External")
      throw new Error("The presentation has an invalid slide relationship.");
    const slidePath = resolvePart(
      "ppt/presentation.xml",
      relationship.getAttribute("Target") || "",
    );
    const slideXml = (await xmlPart(zip, slidePath, true))!;
    const commonSlide = first(slideXml, P, "cSld");
    if (!commonSlide)
      throw new Error("A presentation slide has no content definition.");
    const shapeElements = elements(commonSlide, P, "sp");
    if (shapeElements.length > 1000)
      throw new Error("Slides are limited to 1,000 editable objects.");
    const objects: PptShapeObject[] = shapeElements.map((shape, shapeIndex) => {
      const properties = first(shape, P, "spPr");
      const transform = first(properties, A, "xfrm");
      const offset = first(transform, A, "off");
      const extent = first(transform, A, "ext");
      const textBody = first(shape, P, "txBody");
      const text = textBody ? textOf(textBody) : "";
      const runProperties =
        first(textBody, A, "rPr") || first(textBody, A, "defRPr");
      const geometry = first(properties, A, "prstGeom")?.getAttribute("prst");
      const textBox = first(shape, P, "cNvSpPr")?.getAttribute("txBox") === "1";
      const line = first(properties, A, "ln");
      const alignment = first(textBody, A, "pPr")?.getAttribute("algn");
      const directFill = properties
        ? (Array.from(properties.childNodes).find(
            (node) =>
              node.nodeType === 1 &&
              (node as Element).localName === "solidFill",
          ) as Element | undefined)
        : undefined;
      return {
        id: `object-${index + 1}-${shapeIndex + 1}`,
        kind: textBox || !geometry ? "text" : "shape",
        shapeKind:
          geometry === "ellipse"
            ? "circle"
            : geometry === "roundRect"
              ? "rounded-rect"
              : "rectangle",
        x: numeric(offset?.getAttribute("x"), 0) * xScale,
        y: numeric(offset?.getAttribute("y"), 0) * yScale,
        width: Math.max(
          1,
          numeric(extent?.getAttribute("cx"), 2857500) * xScale,
        ),
        height: Math.max(
          1,
          numeric(extent?.getAttribute("cy"), 952500) * yScale,
        ),
        text,
        fill: directFill ? drawingColor(directFill, "#ffffff") : "transparent",
        stroke:
          line && !first(line, A, "noFill")
            ? drawingColor(line, "#64748b")
            : undefined,
        strokeWidth: numeric(line?.getAttribute("w"), 9525) * xScale,
        color: drawingColor(runProperties, "#24292f"),
        fontSize: Math.max(
          6,
          Math.min(120, numeric(runProperties?.getAttribute("sz"), 1800) / 75),
        ),
        fontWeight:
          runProperties?.getAttribute("b") === "1" ? "bold" : "normal",
        align:
          alignment === "ctr" ? "center" : alignment === "r" ? "right" : "left",
        zIndex: shapeIndex + 1,
        rotation: numeric(transform?.getAttribute("rot"), 0) / 60000,
      };
    });
    const directory = slidePath.slice(0, slidePath.lastIndexOf("/"));
    const filename = slidePath.slice(slidePath.lastIndexOf("/") + 1);
    const slideRelationships = await xmlPart(
      zip,
      `${directory}/_rels/${filename}.rels`,
    );
    const notesRelation = slideRelationships
      ? elements(slideRelationships, relationshipNS, "Relationship").find(
          (item) =>
            item.getAttribute("Type")?.endsWith("/notesSlide") &&
            item.getAttribute("TargetMode") !== "External",
        )
      : undefined;
    let notes = "";
    if (notesRelation) {
      const notesXml = await xmlPart(
        zip,
        resolvePart(slidePath, notesRelation.getAttribute("Target") || ""),
      );
      if (notesXml)
        notes = elements(notesXml, P, "sp")
          .filter(
            (shape) => first(shape, P, "ph")?.getAttribute("type") === "body",
          )
          .map(textOf)
          .join("\n");
    }
    const background = drawingColor(first(commonSlide, P, "bg"), "#ffffff");
    const transition = first(slideXml, P, "transition");
    slides.push({
      id: `slide-${index + 1}`,
      slideNumber: index + 1,
      title:
        commonSlide.getAttribute("name") ||
        objects.find((object) => object.text)?.text?.split("\n")[0] ||
        `Slide ${index + 1}`,
      layout: index === 0 ? "title" : "content",
      background,
      objects,
      notes,
      transition: first(transition, P, "fade")
        ? "fade"
        : first(transition, P, "push")
          ? "slide"
          : first(transition, P, "zoom")
            ? "zoom"
            : "none",
    });
  }
  const theme = await xmlPart(zip, "ppt/theme/theme1.xml");
  return {
    id: `doc-ppt-${crypto.randomUUID()}`,
    title: file.name,
    activeSlideId: slides[0].id,
    themeName: theme?.documentElement.getAttribute("name") || "Imported theme",
    accentColor: drawingColor(
      theme ? first(theme, A, "accent1") : undefined,
      "#007acc",
    ),
    slides,
  };
}
