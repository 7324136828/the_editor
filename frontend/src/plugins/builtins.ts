import type { StudioPlugin } from "./types";
import type { WordDocumentModel } from "../types/office";
import { collectDocumentContext, getDocumentDiagnostics } from "./context";
import { createPluginHost } from "./host";
import { smartArtPlugin } from "./smartArt";

const presets: Array<{
  id: string;
  title: string;
  typography: Partial<WordDocumentModel["typography"]>;
  orientation: "portrait" | "landscape";
}> = [
  {
    id: "business",
    title: "Business document",
    typography: {
      fontFamily: "Calibri",
      fontSize: 14,
      lineHeight: 1.5,
      marginSize: "normal",
      textColor: "#1a202c",
    },
    orientation: "portrait",
  },
  {
    id: "reading",
    title: "Comfortable reading",
    typography: {
      fontFamily: "Georgia",
      fontSize: 16,
      lineHeight: 1.8,
      marginSize: "wide",
      textColor: "#252525",
    },
    orientation: "portrait",
  },
  {
    id: "compact",
    title: "Compact report",
    typography: {
      fontFamily: "Arial",
      fontSize: 12,
      lineHeight: 1.2,
      marginSize: "narrow",
      textColor: "#1a202c",
    },
    orientation: "portrait",
  },
  {
    id: "landscape",
    title: "Landscape report",
    typography: {
      fontFamily: "Calibri",
      fontSize: 14,
      lineHeight: 1.5,
      marginSize: "normal",
    },
    orientation: "landscape",
  },
];

export const documentLayoutPlugin: StudioPlugin = {
  id: "studio.layout",
  name: "Document Layout",
  version: "1.0.0",
  description:
    "Apply editable typography, spacing, margins, and page orientation to Word documents.",
  permissions: ["document:read", "document:write"],
  commands: presets.map((preset) => ({
    id: `studio.layout.${preset.id}`,
    title: preset.title,
    documentTypes: ["word"],
    run(document) {
      if (document.type !== "word")
        throw new Error("Open a Word document to use a layout preset.");
      return {
        message: `Applied ${preset.title.toLowerCase()}. Save to keep these changes.`,
        document: {
          type: "word",
          data: {
            ...document.data,
            modifiedAt: new Date().toISOString(),
            typography: { ...document.data.typography, ...preset.typography },
            sections: document.data.sections.length
              ? document.data.sections.map((section) => ({
                  ...section,
                  orientation: preset.orientation,
                }))
              : [
                  {
                    id: `${document.data.id}-section-1`,
                    title: "Document",
                    pageNumber: 1,
                    orientation: preset.orientation,
                  },
                ],
          },
        },
      };
    },
  })),
};

export const documentContextPlugin: StudioPlugin = {
  id: "studio.context",
  name: "Document Context",
  version: "1.0.0",
  description:
    "Inspect current text, structure, and statistics; export a portable context snapshot without sending data away.",
  permissions: ["document:read"],
  commands: [
    {
      id: "studio.context.snapshot",
      title: "Inspect document context",
      documentTypes: ["word", "excel", "powerpoint", "code"],
      run: (document) => ({
        message: "Context snapshot generated from the current editor.",
        context: collectDocumentContext(document),
      }),
    },
    {
      id: "studio.context.check",
      title: "Run document checks",
      documentTypes: ["word", "excel", "powerpoint", "code"],
      run: (document) => {
        const diagnostics = getDocumentDiagnostics(document);
        return {
          message: diagnostics.length
            ? diagnostics
                .map(
                  (item) =>
                    `${item.severity.toUpperCase()} · ${item.location}: ${item.message}`,
                )
                .join("\n")
            : "No issues found by the available document checks. These checks do not provide grammar, spelling, or full language analysis.",
        };
      },
    },
  ],
};

export const builtinPlugins: readonly StudioPlugin[] = [
  documentLayoutPlugin,
  smartArtPlugin,
  documentContextPlugin,
];
export const defaultPluginHost = createPluginHost({ plugins: builtinPlugins });
export { smartArtPlugin } from "./smartArt";
