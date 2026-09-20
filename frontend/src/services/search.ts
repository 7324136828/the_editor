import { OfficeDocument } from "../types/office";
import { SearchMatch } from "../types/vscode";
import { clone, normalizeDocument } from "./documents";
import { recalculateWorkbook } from "./formulas";
export interface SearchOptions {
  matchCase?: boolean;
  wholeWord?: boolean;
  useRegex?: boolean;
}
export function createMatcher(query: string, options: SearchOptions = {}) {
  if (!query || query.length > 200)
    throw new Error("Enter a search of 1–200 characters.");
  if (options.useRegex) {
    // A deliberately bounded regex subset keeps user input off a backtracking trap.
    let inClass = false,
      quantifiers = 0;
    for (let i = 0; i < query.length; i++) {
      const char = query[i];
      if (char === "\\") {
        if (/[1-9k]/.test(query[++i] ?? ""))
          throw new Error("Regex backreferences are unsupported.");
        continue;
      }
      if (char === "[") {
        inClass = true;
        continue;
      }
      if (char === "]") {
        inClass = false;
        continue;
      }
      if (inClass) continue;
      if ("()|{}".includes(char))
        throw new Error(
          "Regex groups, alternatives and repetition braces are unsupported. Use literal text or a simple character pattern.",
        );
      if ("*+?".includes(char) && ++quantifiers > 1)
        throw new Error(
          "Use at most one regex quantifier to keep searches responsive.",
        );
      if ("*+".includes(char) && !query.startsWith("^"))
        throw new Error(
          "Repeated regex patterns must start with ^ to keep searches responsive.",
        );
    }
  }
  let pattern = options.useRegex
    ? query
    : query.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  if (options.wholeWord) pattern = `\\b(?:${pattern})\\b`;
  return new RegExp(pattern, options.matchCase ? "g" : "gi");
}
function replaceRuns(
  p: import("../types/office").WordParagraph,
  text: string,
  regex: RegExp,
  replacement: string,
) {
  if (!p.runs?.length) return;
  const runs = p.runs,
    segments: {
      text: string;
      bold?: boolean;
      italic?: boolean;
      highlight?: string;
    }[] = [];
  const copyRange = (start: number, end: number) => {
    let offset = 0;
    for (const run of runs) {
      const from = Math.max(start - offset, 0),
        to = Math.min(end - offset, run.text.length);
      if (to > from) segments.push({ ...run, text: run.text.slice(from, to) });
      offset += run.text.length;
    }
  };
  let cursor = 0;
  for (const match of text.matchAll(regex)) {
    const start = match.index!;
    copyRange(cursor, start);
    let offset = 0;
    const style =
      runs.find((run) => {
        offset += run.text.length;
        return offset > start;
      }) ?? runs[runs.length - 1];
    if (replacement) segments.push({ ...style, text: replacement });
    cursor = start + match[0].length;
  }
  copyRange(cursor, text.length);
  p.runs = segments.length ? segments : [{ text: "" }];
}
function textFields(
  doc: OfficeDocument,
  visit: (
    text: string,
    set: (s: string) => void,
    location: string,
    target: SearchMatch["targetRef"],
  ) => void,
) {
  if (doc.type === "word")
    doc.data.paragraphs.forEach((p, i) => {
      if (p.tableData)
        p.tableData.forEach((row, r) =>
          row.forEach((v, c) =>
            visit(
              v,
              (s) => {
                row[c] = s;
              },
              `Table ${i + 1}, ${r + 1}:${c + 1}`,
              { headingId: p.id },
            ),
          ),
        );
      else if (p.smartArt) {
        visit(
          p.smartArt.title,
          (s) => {
            p.smartArt!.title = s;
          },
          `SmartArt ${i + 1}: title`,
          { headingId: p.id },
        );
        p.smartArt.items.forEach((item, itemIndex) =>
          visit(
            item.text,
            (s) => {
              item.text = s;
            },
            `SmartArt ${i + 1}, item ${itemIndex + 1}`,
            { headingId: p.id },
          ),
        );
      } else
        visit(
          p.runs?.map((r) => r.text).join("") ?? p.text ?? "",
          (s) => {
            p.text = s;
            if (p.runs) p.runs = [{ ...p.runs[0], text: s }];
          },
          `Paragraph ${i + 1}`,
          { headingId: p.id },
        );
    });
  else if (doc.type === "excel")
    doc.data.sheets.forEach((sheet) =>
      Object.entries(sheet.cells).forEach(([coord, cell]) =>
        visit(
          cell.formula ?? String(cell.value ?? ""),
          (s) => {
            if (cell.formula && s.startsWith("=")) cell.formula = s;
            else {
              delete cell.formula;
              cell.value =
                typeof cell.value === "number" &&
                s.trim() !== "" &&
                Number.isFinite(Number(s))
                  ? Number(s)
                  : typeof cell.value === "boolean" && /^(true|false)$/i.test(s)
                    ? s.toLowerCase() === "true"
                    : s;
            }
            delete cell.formatted;
          },
          `${sheet.name}!${coord}`,
          { sheetId: sheet.id, cellCoord: coord },
        ),
      ),
    );
  else if (doc.type === "powerpoint")
    doc.data.slides.forEach((slide, i) => {
      visit(
        slide.title,
        (s) => {
          slide.title = s;
        },
        `Slide ${i + 1}: title`,
        { slideIndex: i + 1 },
      );
      visit(
        slide.notes,
        (s) => {
          slide.notes = s;
        },
        `Slide ${i + 1}: notes`,
        { slideIndex: i + 1 },
      );
      slide.objects.forEach((obj) => {
        for (const key of [
          "text",
          "title",
          "subtitle",
          "metricValue",
          "metricLabel",
        ] as const)
          if (obj[key])
            visit(
              obj[key]!,
              (s) => {
                obj[key] = s;
              },
              `Slide ${i + 1}: ${key}`,
              { slideIndex: i + 1, elementId: obj.id },
            );
      });
    });
  else
    doc.data.content
      .split("\n")
      .forEach((line, i) =>
        visit(line, () => {}, `Line ${i + 1}`, { lineNumber: i + 1 }),
      );
}
export function searchDocuments(
  documents: { id: string; name: string; doc: OfficeDocument }[],
  query: string,
  options: SearchOptions = {},
): SearchMatch[] {
  if (!query) return [];
  const regex = createMatcher(query, options),
    results: SearchMatch[] = [];
  for (const { id, name, doc } of documents)
    textFields(doc, (text, _set, location, targetRef) => {
      regex.lastIndex = 0;
      for (const match of text.matchAll(regex)) {
        if (results.length >= 10000) break;
        results.push({
          id: `${id}-${results.length}`,
          fileId: id,
          fileName: name,
          fileType: doc.type,
          location,
          previewText: text.slice(
            Math.max(0, match.index! - 30),
            match.index! + 150,
          ),
          matchIndex: match.index!,
          matchLength: match[0].length,
          targetRef,
        });
      }
    });
  return results;
}
export function replaceInDocument(
  document: OfficeDocument,
  query: string,
  replacement: string,
  options: SearchOptions = {},
) {
  const doc = clone(document),
    regex = createMatcher(query, options);
  let count = 0;
  const replace = (text: string) =>
    text.replace(regex, () => {
      count++;
      return replacement;
    });
  if (doc.type === "code")
    doc.data.content = doc.data.content.split("\n").map(replace).join("\n");
  else if (doc.type === "word")
    for (const p of doc.data.paragraphs) {
      if (p.tableData) p.tableData = p.tableData.map((row) => row.map(replace));
      else if (p.smartArt) {
        p.smartArt.title = replace(p.smartArt.title);
        p.smartArt.items = p.smartArt.items.map((item) => ({
          ...item,
          text: replace(item.text),
        }));
      } else {
        const text = p.runs?.map((r) => r.text).join("") ?? p.text ?? "",
          next = replace(text);
        if (next !== text) {
          replaceRuns(p, text, regex, replacement);
          p.text = next;
        }
      }
    }
  else
    textFields(doc, (text, set) => {
      const next = replace(text);
      if (next !== text) set(next);
    });
  return {
    document: normalizeDocument(
      doc.type === "excel"
        ? { type: "excel", data: recalculateWorkbook(doc.data) }
        : doc,
    ),
    count,
  };
}
