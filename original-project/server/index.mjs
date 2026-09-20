import http from "node:http";
import path from "node:path";
import { fileURLToPath } from "node:url";
import {
  mkdir,
  readFile,
  readdir,
  rename,
  unlink,
  open,
} from "node:fs/promises";
import { randomUUID } from "node:crypto";
import { createMcpHandler } from "./mcp.mjs";

const MAX_BODY = 12 * 1024 * 1024;
const MAX_VERSIONS = 50;
const idPattern = /^[a-zA-Z0-9][a-zA-Z0-9_-]{0,127}$/;
const isObject = (value) =>
  value !== null && typeof value === "object" && !Array.isArray(value);
const fail = (status, message, extra = {}) =>
  Object.assign(new Error(message), { status, ...extra });

/** Reject incomplete models before accepting durable data that the editors cannot open. */
function validateDocument(document, id) {
  if (!isObject(document) || !isObject(document.data))
    throw fail(400, "A document model is required.");
  const data = document.data;
  if (data.id !== id || typeof data.title !== "string")
    throw fail(400, "Document ID must match the URL, and title must be text.");
  const textFields = (value, fields) =>
    isObject(value) &&
    fields.every((field) => typeof value[field] === "string");
  const finiteFields = (value, fields) =>
    isObject(value) && fields.every((field) => Number.isFinite(value[field]));
  const distinctIds = (items) =>
    items.every((item) => isObject(item) && typeof item.id === "string") &&
    new Set(items.map((item) => item.id)).size === items.length;
  const array = (key, max = 10000) => {
    if (!Array.isArray(data[key]) || data[key].length > max)
      throw fail(400, `Invalid ${key}.`);
    return data[key];
  };
  switch (document.type) {
    case "word":
      if (!textFields(data, ["author", "createdAt", "modifiedAt"]))
        throw fail(400, "Invalid Word document metadata.");
      if (
        array("headings").some(
          (heading) =>
            !textFields(heading, ["id", "text"]) ||
            ![1, 2, 3, 4].includes(heading.level) ||
            !Number.isFinite(heading.pageIndex),
        )
      )
        throw fail(400, "Invalid Word outline.");
      if (
        array("sections").some(
          (section) =>
            !textFields(section, ["id", "title"]) ||
            !["portrait", "landscape"].includes(section.orientation) ||
            !Number.isFinite(section.pageNumber),
        )
      )
        throw fail(400, "Invalid Word section.");
      if (
        array("comments").some(
          (comment) =>
            !textFields(comment, [
              "id",
              "author",
              "avatarColor",
              "timestamp",
              "selectedText",
              "comment",
            ]) || typeof comment.resolved !== "boolean",
        )
      )
        throw fail(400, "Invalid Word comment.");
      if (
        !textFields(data.typography, ["fontFamily", "textColor"]) ||
        !finiteFields(data.typography, ["fontSize", "lineHeight"]) ||
        data.typography.fontSize <= 0 ||
        data.typography.lineHeight <= 0 ||
        !["normal", "narrow", "wide"].includes(data.typography.marginSize) ||
        !finiteFields(data.stats, [
          "words",
          "characters",
          "charactersNoSpaces",
          "paragraphs",
          "pages",
          "readingTimeMinutes",
        ]) ||
        typeof data.stats.readabilityScore !== "string"
      )
        throw fail(400, "Word typography and statistics are required.");
      for (const paragraph of array("paragraphs")) {
        if (
          !isObject(paragraph) ||
          typeof paragraph.id !== "string" ||
          ![
            "heading-1",
            "heading-2",
            "heading-3",
            "body",
            "quote",
            "callout",
            "table",
            "smartart",
          ].includes(paragraph.type)
        )
          throw fail(400, "Invalid paragraph.");
        if (paragraph.text !== undefined && typeof paragraph.text !== "string")
          throw fail(400, "Paragraph text must be text.");
        if (
          paragraph.runs !== undefined &&
          (!Array.isArray(paragraph.runs) ||
            paragraph.runs.some(
              (run) => !isObject(run) || typeof run.text !== "string",
            ))
        )
          throw fail(400, "Invalid text runs.");
        if (
          paragraph.tableData !== undefined &&
          (!Array.isArray(paragraph.tableData) ||
            paragraph.tableData.some(
              (row) =>
                !Array.isArray(row) ||
                row.some((cell) => typeof cell !== "string"),
            ))
        )
          throw fail(400, "Invalid table.");
        if (
          (paragraph.type === "smartart") !==
            (paragraph.smartArt !== undefined) ||
          (paragraph.type === "smartart" &&
            (!isObject(paragraph.smartArt) ||
              typeof paragraph.smartArt.title !== "string" ||
              !["process", "cycle", "hierarchy", "pyramid"].includes(
                paragraph.smartArt.layout,
              ) ||
              !/^#[a-f\d]{6}$/i.test(paragraph.smartArt.accentColor) ||
              !Array.isArray(paragraph.smartArt.items) ||
              paragraph.smartArt.items.length < 1 ||
              paragraph.smartArt.items.length > 12 ||
              paragraph.smartArt.items.some(
                (item) =>
                  !isObject(item) ||
                  typeof item.id !== "string" ||
                  typeof item.text !== "string" ||
                  item.text.length > 500,
              ) ||
              !distinctIds(paragraph.smartArt.items)))
        )
          throw fail(400, "Invalid SmartArt diagram.");
      }
      if (!distinctIds(data.paragraphs))
        throw fail(400, "Paragraph IDs must be unique.");
      break;
    case "excel":
      array("formulasAudit");
      if (typeof data.activeSheetId !== "string")
        throw fail(400, "An active worksheet is required.");
      for (const sheet of array("sheets", 50)) {
        if (
          !isObject(sheet) ||
          typeof sheet.id !== "string" ||
          typeof sheet.name !== "string" ||
          !isObject(sheet.cells) ||
          !Array.isArray(sheet.columns) ||
          !Number.isInteger(sheet.rowCount) ||
          !Number.isInteger(sheet.colCount) ||
          sheet.rowCount < 1 ||
          sheet.rowCount > 1000 ||
          sheet.colCount < 1 ||
          sheet.colCount > 100 ||
          sheet.rowCount * sheet.colCount > 10000 ||
          Object.keys(sheet.cells).length > 10000
        )
          throw fail(
            400,
            "Invalid worksheet or grid exceeds 1,000 rows, 100 columns, or 10,000 cells.",
          );
        if (
          sheet.columns.length !== sheet.colCount ||
          sheet.columns.some(
            (column) =>
              !textFields(column, ["key", "label", "type"]) ||
              !/^[A-Z]{1,3}$/.test(column.key) ||
              !Number.isFinite(column.width),
          )
        )
          throw fail(400, "Invalid worksheet columns.");
        for (const [coord, cell] of Object.entries(sheet.cells)) {
          if (
            !/^[A-Z]{1,3}[1-9][0-9]{0,6}$/.test(coord) ||
            !isObject(cell) ||
            (!["string", "number", "boolean"].includes(typeof cell.value) &&
              cell.value !== null) ||
            (cell.formula !== undefined && typeof cell.formula !== "string")
          )
            throw fail(400, "Invalid worksheet cell.");
          const [, letters, row] = /^([A-Z]+)([0-9]+)$/.exec(coord);
          const column = [...letters].reduce(
            (sum, letter) => sum * 26 + letter.charCodeAt(0) - 64,
            0,
          );
          if (column > sheet.colCount || Number(row) > sheet.rowCount)
            throw fail(400, "A worksheet cell is outside the visible grid.");
        }
      }
      if (
        !data.sheets.length ||
        !distinctIds(data.sheets) ||
        !data.sheets.some((sheet) => sheet.id === data.activeSheetId)
      )
        throw fail(
          400,
          "An active worksheet and unique worksheet IDs are required.",
        );
      break;
    case "powerpoint":
      if (typeof data.activeSlideId !== "string")
        throw fail(400, "An active slide is required.");
      if (!textFields(data, ["themeName", "accentColor"]))
        throw fail(400, "Invalid presentation theme.");
      for (const slide of array("slides", 200)) {
        if (
          !textFields(slide, ["id", "title", "background", "notes"]) ||
          !Array.isArray(slide.objects) ||
          slide.objects.length > 1000 ||
          !Number.isFinite(slide.slideNumber) ||
          !["none", "fade", "slide", "zoom"].includes(slide.transition) ||
          !["title", "content", "two-column", "dashboard", "blank"].includes(
            slide.layout,
          )
        )
          throw fail(400, "Invalid slide.");
        for (const object of slide.objects) {
          if (
            !textFields(object, ["id", "fill", "color"]) ||
            !finiteFields(object, [
              "x",
              "y",
              "width",
              "height",
              "fontSize",
              "zIndex",
            ]) ||
            !["text", "shape", "metric", "image", "chart"].includes(
              object.kind,
            ) ||
            ["text", "title", "subtitle", "metricValue", "metricLabel"].some(
              (key) =>
                object[key] !== undefined && typeof object[key] !== "string",
            )
          )
            throw fail(400, "Invalid slide object.");
        }
        if (!distinctIds(slide.objects))
          throw fail(400, "Slide object IDs must be unique.");
      }
      if (
        !data.slides.length ||
        !distinctIds(data.slides) ||
        !data.slides.some((slide) => slide.id === data.activeSlideId)
      )
        throw fail(400, "An active slide and unique slide IDs are required.");
      break;
    case "code":
      array("symbols");
      if (
        typeof data.content !== "string" ||
        !["typescript", "javascript", "python", "markdown", "json"].includes(
          data.language,
        )
      )
        throw fail(400, "Invalid text document.");
      break;
    default:
      throw fail(400, "Unsupported document type.");
  }
}

async function readJson(request) {
  if (
    !/^application\/json(?:\s*;|$)/i.test(request.headers["content-type"] || "")
  )
    throw fail(415, "Use application/json.");
  const chunks = [];
  let size = 0;
  for await (const chunk of request) {
    size += chunk.length;
    if (size > MAX_BODY)
      throw fail(413, "Document exceeds the 12 MB save limit.");
    chunks.push(chunk);
  }
  try {
    return JSON.parse(Buffer.concat(chunks).toString("utf8"));
  } catch {
    throw fail(400, "Invalid JSON.");
  }
}

/** One envelope contains head and history, so an atomic rename commits both together. */
export function createOfficeServer({
  dataDir = path.resolve(".office-data"),
  allowedOrigins = [],
  staticDir = null,
} = {}) {
  const root = path.resolve(dataDir);
  const locks = new Map();
  const ready = mkdir(root, { recursive: true });
  const fileFor = (id) => path.join(root, `${id}.json`);
  const readEnvelope = async (id) => {
    try {
      return JSON.parse(await readFile(fileFor(id), "utf8"));
    } catch (error) {
      if (error.code === "ENOENT") return null;
      throw error;
    }
  };
  const listRecords = async () => {
    await ready;
    const names = (await readdir(root)).filter((name) =>
      /^[a-zA-Z0-9][a-zA-Z0-9_-]{0,127}\.json$/.test(name),
    );
    const records = await Promise.all(
      names.map(async (name) => (await readEnvelope(name.slice(0, -5)))?.head),
    );
    return records
      .filter(Boolean)
      .sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));
  };
  const mcpHandler = createMcpHandler({
    listDocuments: async () =>
      (await listRecords()).map((record) => record.document),
    allowedOrigins,
  });
  const commit = async (id, envelope) => {
    const temp = path.join(root, `${id}.${randomUUID()}.tmp`);
    let handle;
    try {
      handle = await open(temp, "wx");
      await handle.writeFile(JSON.stringify(envelope), "utf8");
      await handle.sync();
      await handle.close();
      handle = null;
      await rename(temp, fileFor(id));
    } finally {
      if (handle) await handle.close();
      await unlink(temp).catch((error) => {
        if (error.code !== "ENOENT") throw error;
      });
    }
  };
  const locked = (id, operation) => {
    const pending = (locks.get(id) || Promise.resolve())
      .catch(() => {})
      .then(operation);
    locks.set(id, pending);
    return pending.finally(() => {
      if (locks.get(id) === pending) locks.delete(id);
    });
  };

  return http.createServer(async (request, response) => {
    const send = (status, value) => {
      response.writeHead(status, {
        "Content-Type": "application/json; charset=utf-8",
        "Cache-Control": "no-store",
        "X-Content-Type-Options": "nosniff",
      });
      response.end(JSON.stringify(value));
    };
    try {
      await ready;
      const host = request.headers.host || "";
      if (!/^(localhost|127\.0\.0\.1|\[::1\])(?::\d+)?$/i.test(host))
        throw fail(403, "This service accepts localhost requests only.");
      const origin = request.headers.origin;
      const allowed = new Set([
        `http://${host}`,
        "http://localhost:5173",
        "http://127.0.0.1:5173",
        "http://localhost:5174",
        "http://127.0.0.1:5174",
        "http://localhost:4173",
        "http://127.0.0.1:4173",
        ...allowedOrigins,
      ]);
      if (origin && !allowed.has(origin))
        throw fail(403, "Origin is not allowed.");
      if (origin) {
        response.setHeader("Access-Control-Allow-Origin", origin);
        response.setHeader("Vary", "Origin");
      }
      const url = new URL(request.url, `http://${host}`);
      if (url.pathname === "/mcp") {
        await mcpHandler(request, response);
        return;
      }
      if (request.method === "OPTIONS") {
        response.writeHead(204, {
          "Access-Control-Allow-Methods": "GET, PUT, OPTIONS",
          "Access-Control-Allow-Headers": "Content-Type",
        });
        response.end();
        return;
      }
      if (url.pathname === "/api/health" && request.method === "GET") {
        send(200, {
          status: "ok",
          storage: "local-disk",
          maxVersions: MAX_VERSIONS,
        });
        return;
      }
      if (url.pathname === "/api/documents" && request.method === "GET") {
        send(200, { documents: await listRecords() });
        return;
      }
      if (
        staticDir &&
        !url.pathname.startsWith("/api/") &&
        ["GET", "HEAD"].includes(request.method)
      ) {
        let pathname;
        try {
          pathname = decodeURIComponent(url.pathname);
        } catch {
          throw fail(400, "Invalid asset path.");
        }
        if (
          pathname.includes("\0") ||
          pathname.includes("\\") ||
          pathname.split("/").includes("..")
        )
          throw fail(400, "Invalid asset path.");
        const assetsRoot = path.resolve(staticDir);
        let target = path.resolve(
          assetsRoot,
          `.${pathname === "/" ? "/index.html" : pathname}`,
        );
        const relative = path.relative(assetsRoot, target);
        if (relative.startsWith("..") || path.isAbsolute(relative))
          throw fail(403, "Asset path is outside the application.");
        let bytes;
        try {
          bytes = await readFile(target);
        } catch (error) {
          if (error.code !== "ENOENT" && error.code !== "EISDIR") throw error;
          if (path.extname(pathname)) throw fail(404, "Asset not found.");
          target = path.join(assetsRoot, "index.html");
          try {
            bytes = await readFile(target);
          } catch {
            throw fail(
              503,
              "Build the application with npm run build before using npm start.",
            );
          }
        }
        const mime =
          {
            ".html": "text/html; charset=utf-8",
            ".js": "application/javascript; charset=utf-8",
            ".css": "text/css; charset=utf-8",
            ".json": "application/json",
            ".svg": "image/svg+xml",
            ".png": "image/png",
            ".jpg": "image/jpeg",
            ".ico": "image/x-icon",
            ".woff2": "font/woff2",
            ".map": "application/json",
          }[path.extname(target)] || "application/octet-stream";
        response.writeHead(200, {
          "Content-Type": mime,
          "Content-Length": bytes.length,
          "X-Content-Type-Options": "nosniff",
          "Cache-Control":
            path.extname(target) === ".html"
              ? "no-cache"
              : "public, max-age=86400",
        });
        response.end(request.method === "HEAD" ? undefined : bytes);
        return;
      }
      const match =
        /^\/api\/documents\/([^/]+)(?:\/history(?:\/(\d+))?)?$/.exec(
          url.pathname,
        );
      if (!match) throw fail(404, "Endpoint not found.");
      let id;
      try {
        id = decodeURIComponent(match[1]);
      } catch {
        throw fail(400, "Invalid document ID.");
      }
      if (!idPattern.test(id))
        throw fail(
          400,
          "Invalid document ID. Use letters, numbers, underscores, and hyphens.",
        );
      const isHistory = url.pathname.includes("/history");
      if (request.method === "GET") {
        const envelope = await readEnvelope(id);
        if (!envelope) throw fail(404, "Document not found.");
        if (isHistory && !match[2]) {
          send(200, {
            versions: envelope.versions
              .map(({ revision, name, updatedAt }) => ({
                revision,
                name,
                updatedAt,
              }))
              .reverse(),
          });
          return;
        }
        const record = match[2]
          ? envelope.versions.find(
              (version) => version.revision === Number(match[2]),
            )
          : envelope.head;
        if (!record) throw fail(404, "Version is no longer available.");
        send(200, record);
        return;
      }
      if (request.method !== "PUT" || isHistory)
        throw fail(405, "Method not allowed.");
      const body = await readJson(request);
      if (
        !isObject(body) ||
        typeof body.name !== "string" ||
        !body.name.trim() ||
        body.name.length > 240 ||
        /[\x00-\x1f\\/]/.test(body.name)
      )
        throw fail(
          400,
          "A filename of 1–240 characters, without path separators, is required.",
        );
      if (
        body.revision !== null &&
        (!Number.isInteger(body.revision) || body.revision < 1)
      )
        throw fail(
          400,
          "Send revision:null for a new document or its current revision number.",
        );
      validateDocument(body.document, id);
      await locked(id, async () => {
        const existing = await readEnvelope(id);
        if ((existing?.head.revision ?? null) !== body.revision)
          throw fail(
            409,
            "This document was changed elsewhere. Reload the saved version or save a copy.",
            { current: existing?.head ?? null },
          );
        const record = {
          id,
          name: body.name.trim(),
          document: body.document,
          revision: (existing?.head.revision || 0) + 1,
          updatedAt: new Date().toISOString(),
        };
        const versions = [...(existing?.versions || []), record].slice(
          -MAX_VERSIONS,
        );
        await commit(id, { schemaVersion: 1, head: record, versions });
        send(existing ? 200 : 201, record);
      });
    } catch (error) {
      if (response.headersSent) {
        response.destroy();
        return;
      }
      if (!error.status)
        console.error("Document storage error:", error.message);
      send(error.status || 500, {
        error: error.status
          ? error.message
          : "The document could not be read or saved. Check the local server storage.",
        ...(error.current !== undefined ? { current: error.current } : {}),
      });
    }
  });
}

if (
  process.argv[1] &&
  path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)
) {
  const port = Number(process.env.OFFICE_PORT || 3001);
  const server = createOfficeServer({
    dataDir: process.env.OFFICE_DATA_DIR || path.resolve(".office-data"),
    allowedOrigins: (process.env.OFFICE_ALLOWED_ORIGINS || "")
      .split(",")
      .filter(Boolean),
    staticDir: process.argv.includes("--static") ? path.resolve("dist") : null,
  });
  server.listen(port, "127.0.0.1", () =>
    console.log(`Office Studio storage listening on http://127.0.0.1:${port}`),
  );
  server.on("error", (error) => {
    console.error(error.message);
    process.exitCode = 1;
  });
}
