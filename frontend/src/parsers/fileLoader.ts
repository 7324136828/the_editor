import { OfficeDocument, CodeDocumentModel } from "../types/office";
import { parseDocxFile } from "./docxParser";
import { parseXlsxFile } from "./xlsxParser";
import { parsePptxFile } from "./pptxParser";

export async function loadOfficeFile(file: File): Promise<OfficeDocument> {
  if (file.size > 20 * 1024 * 1024)
    throw new Error("This editor supports files up to 20 MB.");
  const ext = file.name.split(".").pop()?.toLowerCase();
  if (ext === "doc" || ext === "ppt")
    throw new Error(
      "Legacy DOC/PPT files are unsupported. Convert a copy to DOCX/PPTX first.",
    );
  if (
    !ext ||
    ![
      "docx",
      "xlsx",
      "xls",
      "csv",
      "pptx",
      "txt",
      "md",
      "json",
      "ts",
      "tsx",
      "js",
      "jsx",
      "py",
      "html",
      "css",
      "yaml",
      "yml",
      "xml",
    ].includes(ext)
  )
    throw new Error(
      "Unsupported file type. Open DOCX, XLSX, CSV, PPTX or a text/code file.",
    );

  if (ext === "docx" || ext === "doc") {
    const doc = await parseDocxFile(file);
    return { type: "word", data: doc };
  }

  if (ext === "xlsx" || ext === "xls" || ext === "csv") {
    const doc = await parseXlsxFile(file);
    return { type: "excel", data: doc };
  }

  if (ext === "pptx" || ext === "ppt") {
    const doc = await parsePptxFile(file);
    return { type: "powerpoint", data: doc };
  }

  // Fallback to text / code
  const text = await file.text();
  const codeDoc: CodeDocumentModel = {
    id: `doc-code-${Date.now()}`,
    title: file.name,
    language:
      ext === "ts" || ext === "tsx"
        ? "typescript"
        : ext === "py"
          ? "python"
          : ext === "json"
            ? "json"
            : ext === "js" || ext === "jsx"
              ? "javascript"
              : "markdown",
    content: text,
    symbols: [],
  };

  return { type: "code", data: codeDoc };
}
