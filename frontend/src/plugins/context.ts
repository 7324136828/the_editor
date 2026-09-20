import type { OfficeDocument, WordParagraph } from "../types/office";
import type { ProblemItem } from "../types/vscode";
import type { DocumentContext } from "./types";

function paragraphText(paragraph: WordParagraph): string {
  if (paragraph.type === "table")
    return (paragraph.tableData ?? []).map((row) => row.join("\t")).join("\n");
  if (paragraph.type === "smartart")
    return paragraph.smartArt
      ? [
          paragraph.smartArt.title,
          ...paragraph.smartArt.items.map((item) => item.text),
        ].join("\n")
      : "";
  return (
    paragraph.text ?? paragraph.runs?.map((run) => run.text).join("") ?? ""
  );
}

export function collectDocumentContext(
  document: OfficeDocument,
  maxCharacters = 12000,
): DocumentContext {
  let text = "";
  let summary = "";
  let outline: string[] = [];
  let statistics: Record<string, number> = {};
  switch (document.type) {
    case "word": {
      text = document.data.paragraphs.map(paragraphText).join("\n");
      const words = text.trim() ? text.trim().split(/\s+/u).length : 0;
      outline = document.data.paragraphs
        .filter((paragraph) => paragraph.type.startsWith("heading-"))
        .map(paragraphText);
      statistics = {
        words,
        characters: text.length,
        paragraphs: document.data.paragraphs.length,
        headings: outline.length,
        unresolvedComments: document.data.comments.filter(
          (comment) => !comment.resolved,
        ).length,
      };
      summary = `${words.toLocaleString()} words in ${statistics.paragraphs} paragraphs; ${outline.length} headings and ${statistics.unresolvedComments} unresolved comments.`;
      break;
    }
    case "excel": {
      outline = document.data.sheets.map(
        (sheet) =>
          `${sheet.name} (${sheet.rowCount} rows × ${sheet.colCount} columns)`,
      );
      const cells = document.data.sheets.flatMap((sheet) =>
        Object.values(sheet.cells),
      );
      statistics = {
        sheets: document.data.sheets.length,
        populatedCells: cells.filter(
          (cell) => (cell.value !== null && cell.value !== "") || cell.formula,
        ).length,
        formulas: cells.filter((cell) => cell.formula).length,
      };
      text = document.data.sheets
        .map(
          (sheet) =>
            `[${sheet.name}]\n${Object.entries(sheet.cells)
              .filter(
                ([, cell]) =>
                  (cell.value !== null && cell.value !== "") || cell.formula,
              )
              .map(
                ([address, cell]) =>
                  `${address}: ${cell.formula ?? String(cell.value)}${cell.formula ? ` → ${String(cell.value ?? "")}` : ""}`,
              )
              .join("\n")}`,
        )
        .join("\n\n");
      summary = `${statistics.sheets} worksheets, ${statistics.populatedCells} populated cells and ${statistics.formulas} formulas. Formula values are the editor's current values.`;
      break;
    }
    case "powerpoint": {
      outline = document.data.slides.map(
        (slide, index) => `${index + 1}. ${slide.title || "Untitled slide"}`,
      );
      statistics = {
        slides: document.data.slides.length,
        objects: document.data.slides.reduce(
          (count, slide) => count + slide.objects.length,
          0,
        ),
        slidesWithNotes: document.data.slides.filter((slide) =>
          slide.notes.trim(),
        ).length,
      };
      text = document.data.slides
        .map(
          (slide, index) =>
            `Slide ${index + 1}: ${slide.title}\n${slide.objects
              .filter((object) => object.visible !== false)
              .map((object) =>
                [
                  object.text,
                  object.title,
                  object.subtitle,
                  object.metricValue,
                  object.metricLabel,
                ]
                  .filter(Boolean)
                  .join(" "),
              )
              .join("\n")}${slide.notes ? `\nNotes: ${slide.notes}` : ""}`,
        )
        .join("\n\n");
      summary = `${statistics.slides} slides with ${statistics.objects} objects; ${statistics.slidesWithNotes} slides have speaker notes.`;
      break;
    }
    case "code": {
      text = document.data.content;
      const lines = text.split("\n");
      outline = lines.flatMap((line, index) => {
        const match = line.match(
          /^\s*(?:(?:export\s+)?(?:async\s+)?(?:function|class|interface|def)\s+([\w$]+)|(?:#{1,6})\s+(.+))/,
        );
        return match ? [`Line ${index + 1}: ${match[1] || match[2]}`] : [];
      });
      statistics = {
        lines: lines.length,
        characters: text.length,
        symbols: outline.length,
      };
      summary = `${document.data.language} document with ${statistics.lines} lines and ${statistics.characters.toLocaleString()} characters.`;
      break;
    }
  }
  const limit = Number.isFinite(maxCharacters)
    ? Math.max(0, Math.floor(maxCharacters))
    : 12000;
  return {
    schemaVersion: 1,
    source: "current-editor",
    document: {
      id: document.data.id,
      title: document.data.title,
      type: document.type,
    },
    summary,
    statistics,
    outline,
    text: text.slice(0, limit),
    truncated: text.length > limit,
  };
}

/** Deterministic checks, not an LSP client, spell checker, or model-generated review. */
export function getDocumentDiagnostics(
  document: OfficeDocument,
): ProblemItem[] {
  const problems: ProblemItem[] = [];
  const add = (
    severity: ProblemItem["severity"],
    message: string,
    location: string,
  ) => {
    problems.push({
      id: `${document.data.id}-context-${problems.length}`,
      severity,
      message,
      location,
      fileId: document.data.id,
      fileName: document.data.title,
      source: "Document checks",
    });
  };
  if (document.type === "word") {
    let previousLevel = 0;
    document.data.paragraphs.forEach((paragraph, index) => {
      if (paragraph.type.startsWith("heading-")) {
        const level = Number(paragraph.type.slice(-1));
        if (level > previousLevel + 1)
          add(
            "warning",
            `Heading jumps from level ${previousLevel} to level ${level}.`,
            `Paragraph ${index + 1}`,
          );
        if (!paragraphText(paragraph).trim())
          add("warning", "Empty heading.", `Paragraph ${index + 1}`);
        previousLevel = level;
      }
    });
    const unresolved = document.data.comments.filter(
      (comment) => !comment.resolved,
    ).length;
    if (unresolved)
      add(
        "info",
        `${unresolved} unresolved comment${unresolved === 1 ? "" : "s"}.`,
        "Comments",
      );
  } else if (document.type === "excel") {
    document.data.sheets.forEach((sheet) => {
      Object.entries(sheet.cells).forEach(([address, cell]) => {
        if (
          typeof cell.value === "string" &&
          /^#(?:REF!|VALUE!|DIV\/0!|NAME\?|N\/A|NUM!|SPILL!|CYCLE!|ERROR!)/.test(
            cell.value,
          )
        )
          add(
            "error",
            `Cell contains ${cell.value}.`,
            `${sheet.name}!${address}`,
          );
        else if (cell.formula?.includes("#REF!"))
          add(
            "error",
            "Formula contains a broken reference.",
            `${sheet.name}!${address}`,
          );
      });
    });
  } else if (document.type === "powerpoint") {
    document.data.slides.forEach((slide, index) => {
      if (!slide.title.trim())
        add("warning", "Slide title is empty.", `Slide ${index + 1}`);
      if (!slide.objects.some((object) => object.visible !== false))
        add("info", "Slide has no visible objects.", `Slide ${index + 1}`);
    });
  } else {
    if (document.data.language === "json") {
      try {
        JSON.parse(document.data.content);
      } catch (error) {
        add(
          "error",
          error instanceof Error ? error.message : "Invalid JSON.",
          "JSON",
        );
      }
    }
    document.data.content.split("\n").forEach((line, index) => {
      if (/\b(?:TODO|FIXME)\b/.test(line))
        add("info", line.trim().slice(0, 160), `Line ${index + 1}`);
    });
  }
  return problems;
}
