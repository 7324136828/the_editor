import type {
  SmartArtLayout,
  WordParagraph,
  WordSmartArt,
} from "../types/office";
import type { StudioPlugin } from "./types";

const templates: Record<
  SmartArtLayout,
  { title: string; steps: string[]; color: string }
> = {
  process: {
    title: "Process",
    steps: ["Plan", "Build", "Review", "Deliver"],
    color: "#2b6cb0",
  },
  cycle: {
    title: "Continuous cycle",
    steps: ["Discover", "Design", "Create", "Learn"],
    color: "#0f766e",
  },
  hierarchy: {
    title: "Team hierarchy",
    steps: ["Leadership", "Product", "Engineering", "Operations"],
    color: "#7c3aed",
  },
  pyramid: {
    title: "Priority pyramid",
    steps: ["Vision", "Strategy", "Initiatives", "Actions"],
    color: "#c2410c",
  },
};

export function createSmartArtBlock(
  layout: SmartArtLayout,
  id = `p-${crypto.randomUUID()}`,
): WordParagraph {
  const template = templates[layout];
  const smartArt: WordSmartArt = {
    title: template.title,
    layout,
    accentColor: template.color,
    items: template.steps.map((text, index) => ({
      id: `${id}-item-${index + 1}`,
      text,
    })),
  };
  return { id, type: "smartart", smartArt };
}

export const smartArtPlugin: StudioPlugin = {
  id: "studio.smartart",
  name: "SmartArt for Word",
  version: "1.0.0",
  description:
    "Insert editable process, cycle, hierarchy, and pyramid diagrams into Word documents.",
  permissions: ["document:read", "document:write"],
  commands: (Object.keys(templates) as SmartArtLayout[]).map((layout) => ({
    id: `studio.smartart.insert-${layout}`,
    title: `Insert ${layout} SmartArt`,
    documentTypes: ["word"] as const,
    run(document) {
      if (document.type !== "word")
        throw new Error("Open a Word document to insert SmartArt.");
      const block = createSmartArtBlock(layout);
      return {
        message: `Inserted an editable ${layout} SmartArt diagram. Save to keep it.`,
        document: {
          type: "word" as const,
          data: {
            ...document.data,
            modifiedAt: new Date().toISOString(),
            paragraphs: [...document.data.paragraphs, block],
          },
        },
      };
    },
  })),
};
