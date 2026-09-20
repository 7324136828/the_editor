import type { PptSlide } from "../types/office";

/** Reflow existing content into a layout without deleting objects or their text. */
export function applySlideLayout(
  slide: PptSlide,
  layout: PptSlide["layout"],
): PptSlide {
  if (!slide.objects.length || layout === "blank") return { ...slide, layout };
  const [heading, ...body] = slide.objects;
  if (layout === "title") {
    return {
      ...slide,
      layout,
      objects: [
        {
          ...heading,
          x: 70,
          y: 85,
          width: 780,
          height: 115,
          align: "center",
          fontSize: Math.max(32, heading.fontSize),
        },
        ...body.map((object, index) => ({
          ...object,
          x: 150,
          y: 220 + index * (245 / Math.max(1, body.length)),
          width: 620,
          height: Math.max(25, 220 / Math.max(1, body.length)),
          align: "center" as const,
        })),
      ],
    };
  }
  const columns =
    layout === "two-column"
      ? 2
      : layout === "dashboard"
        ? body.length > 4
          ? 3
          : 2
        : 1;
  const rows = Math.max(1, Math.ceil(body.length / columns));
  const width = (800 - (columns - 1) * 20) / columns;
  const height = (330 - (rows - 1) * 15) / rows;
  return {
    ...slide,
    layout,
    objects: [
      { ...heading, x: 60, y: 35, width: 800, height: 75, align: "left" },
      ...body.map((object, index) => ({
        ...object,
        x: 60 + (index % columns) * (width + 20),
        y: 135 + Math.floor(index / columns) * (height + 15),
        width,
        height: Math.max(20, height),
        fontSize: Math.min(object.fontSize, rows > 4 ? 14 : 24),
      })),
    ],
  };
}
