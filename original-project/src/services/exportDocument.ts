import JSZip from "jszip";
import * as XLSX from "xlsx";
import type {
  OfficeDocument,
  WordDocumentModel,
  WordParagraph,
  PptDocumentModel,
  PptShapeObject,
} from "../types/office";

export interface SerializedDocument {
  data: Uint8Array;
  mimeType: string;
  extension: string;
}
const header = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const relNS = "http://schemas.openxmlformats.org/package/2006/relationships";
const officeRel =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const wordNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const pptNS = "http://schemas.openxmlformats.org/presentationml/2006/main";
const drawNS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const xml = (value: unknown) =>
  String(value ?? "")
    .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
const color = (value: string | undefined, fallback = "24292F"): string => {
  const hex = value?.match(/#([a-f\d]{6}|[a-f\d]{3})\b/i)?.[1];
  if (hex)
    return hex.length === 3
      ? hex
          .split("")
          .map((ch) => ch + ch)
          .join("")
          .toUpperCase()
      : hex.toUpperCase();
  const rgb = value?.match(/rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/i);
  return rgb
    ? rgb
        .slice(1)
        .map((n) => Math.min(255, Number(n)).toString(16).padStart(2, "0"))
        .join("")
        .toUpperCase()
    : fallback;
};
const relationships = (items: { id: string; type: string; target: string }[]) =>
  header +
  `<Relationships xmlns="${relNS}">${items.map((item) => `<Relationship Id="${item.id}" Type="${officeRel}/${item.type}" Target="${xml(item.target)}"/>`).join("")}</Relationships>`;
const types = (parts: [string, string][]) =>
  header +
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
  parts
    .map(
      ([part, contentType]) =>
        `<Override PartName="/${part}" ContentType="application/vnd.openxmlformats-officedocument.${contentType}+xml"/>`,
    )
    .join("") +
  "</Types>";

function wordRun(
  text: string,
  options: {
    bold?: boolean;
    italic?: boolean;
    highlight?: string;
    textColor?: string;
  } = {},
) {
  const formatting = `${options.bold ? "<w:b/>" : ""}${options.italic ? "<w:i/>" : ""}${options.highlight ? `<w:shd w:fill="${color(options.highlight, "FFFF00")}"/>` : ""}${options.textColor ? `<w:color w:val="${color(options.textColor)}"/>` : ""}`;
  const content = text
    .split(/(\n|\t)/)
    .map((part) =>
      part === "\n"
        ? "<w:br/>"
        : part === "\t"
          ? "<w:tab/>"
          : `<w:t xml:space="preserve">${xml(part)}</w:t>`,
    )
    .join("");
  return `<w:r>${formatting ? `<w:rPr>${formatting}</w:rPr>` : ""}${content}</w:r>`;
}

function wordParagraph(paragraph: WordParagraph): string {
  if (paragraph.type === "table") {
    const rows = paragraph.tableData || [];
    const columns = Math.max(1, ...rows.map((row) => row.length));
    const width = Math.floor(9360 / columns);
    return `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>${["top", "left", "bottom", "right", "insideH", "insideV"].map((edge) => `<w:${edge} w:val="single" w:sz="4" w:color="D0D7DE"/>`).join("")}</w:tblBorders></w:tblPr><w:tblGrid>${Array.from({ length: columns }, () => `<w:gridCol w:w="${width}"/>`).join("")}</w:tblGrid>${rows.map((row) => `<w:tr>${Array.from({ length: columns }, (_, index) => `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr><w:p>${wordRun(row[index] || "")}</w:p></w:tc>`).join("")}</w:tr>`).join("")}</w:tbl>`;
  }
  if (paragraph.type === "smartart") {
    const diagram = paragraph.smartArt;
    if (!diagram) return `<w:p>${wordRun("SmartArt diagram")}</w:p>`;
    const accent = color(diagram.accentColor, "2B6CB0");
    const columns = Math.max(1, diagram.items.length);
    const width = Math.floor(9360 / columns);
    const border = ["top", "left", "bottom", "right", "insideH", "insideV"]
      .map((edge) => `<w:${edge} w:val="single" w:sz="8" w:color="${accent}"/>`)
      .join("");
    return `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>${border}</w:tblBorders></w:tblPr><w:tblGrid>${diagram.items.map(() => `<w:gridCol w:w="${width}"/>`).join("")}</w:tblGrid><w:tr><w:tc><w:tcPr><w:gridSpan w:val="${columns}"/><w:shd w:fill="${accent}"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/></w:pPr>${wordRun(diagram.title, { bold: true, textColor: "#ffffff" })}</w:p></w:tc></w:tr><w:tr>${diagram.items.map((item, index) => `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/><w:shd w:fill="${index % 2 ? "F8FAFC" : "EEF4FB"}"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/></w:pPr>${wordRun(`${index + 1}. ${item.text}`, { bold: true })}</w:p></w:tc>`).join("")}</w:tr></w:tbl>`;
  }
  const level = paragraph.type.match(/^heading-(\d)$/)?.[1];
  const props = `${level ? `<w:pStyle w:val="Heading${level}"/>` : ""}${paragraph.align ? `<w:jc w:val="${paragraph.align === "justify" ? "both" : paragraph.align}"/>` : ""}${paragraph.type === "quote" ? '<w:ind w:left="480"/><w:pBdr><w:left w:val="single" w:sz="12" w:space="8" w:color="64748B"/></w:pBdr>' : ""}${paragraph.type === "callout" ? '<w:shd w:fill="EAF2F8"/>' : ""}`;
  const runs = paragraph.runs?.length
    ? paragraph.runs.map((run) => wordRun(run.text, run)).join("")
    : wordRun(paragraph.text || "");
  return `<w:p>${props ? `<w:pPr>${props}</w:pPr>` : ""}${runs}</w:p>`;
}

async function serializeWord(document: WordDocumentModel): Promise<Uint8Array> {
  const zip = new JSZip();
  const typography = document.typography;
  const margin = { normal: 1440, narrow: 720, wide: 2160 }[
    typography.marginSize
  ];
  const landscape = document.sections[0]?.orientation === "landscape";
  const comments = document.comments.filter((comment) => !comment.resolved);
  const commentParagraphs = comments.map((comment) => {
    const found = document.paragraphs.findIndex(
      (paragraph) =>
        paragraph.type !== "table" &&
        paragraph.type !== "smartart" &&
        (
          paragraph.runs?.map((run) => run.text).join("") ||
          paragraph.text ||
          ""
        ).includes(comment.selectedText),
    );
    return found >= 0
      ? found
      : Math.max(
          0,
          document.paragraphs.findIndex(
            (paragraph) =>
              paragraph.type !== "table" && paragraph.type !== "smartart",
          ),
        );
  });
  const body = document.paragraphs
    .map((paragraph, paragraphIndex) => {
      const anchored = comments
        .map((_, index) => index)
        .filter((index) => commentParagraphs[index] === paragraphIndex);
      const content = wordParagraph(paragraph);
      if (
        paragraph.type === "table" ||
        paragraph.type === "smartart" ||
        !anchored.length
      )
        return content;
      // The editor stores selected text, not range offsets; anchor at the matching paragraph.
      return content.replace(
        "</w:p>",
        `${anchored.map((index) => `<w:r><w:commentReference w:id="${index}"/></w:r>`).join("")}</w:p>`,
      );
    })
    .join("");
  zip.file(
    "[Content_Types].xml",
    types([
      ["word/document.xml", "wordprocessingml.document.main"],
      ["word/styles.xml", "wordprocessingml.styles"],
      ...(comments.length
        ? [
            ["word/comments.xml", "wordprocessingml.comments"] as [
              string,
              string,
            ],
          ]
        : []),
    ]),
  );
  zip.file(
    "_rels/.rels",
    relationships([
      { id: "rId1", type: "officeDocument", target: "word/document.xml" },
    ]),
  );
  zip.file(
    "word/_rels/document.xml.rels",
    relationships([
      { id: "rId1", type: "styles", target: "styles.xml" },
      ...(comments.length
        ? [{ id: "rId2", type: "comments", target: "comments.xml" }]
        : []),
      ...(document.paragraphs.some((paragraph) => paragraph.type === "smartart")
        ? [
            {
              id: comments.length ? "rId3" : "rId2",
              type: "customXml",
              target: "../customXml/officeStudioSmartArt.xml",
            },
          ]
        : []),
    ]),
  );
  zip.file(
    "word/document.xml",
    header +
      `<w:document xmlns:w="${wordNS}" xmlns:r="${officeRel}"><w:body>${body}<w:sectPr><w:pgSz w:w="${landscape ? 15840 : 12240}" w:h="${landscape ? 12240 : 15840}"${landscape ? ' w:orient="landscape"' : ""}/><w:pgMar w:top="${margin}" w:right="${margin}" w:bottom="${margin}" w:left="${margin}" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`,
  );
  zip.file(
    "word/styles.xml",
    header +
      `<w:styles xmlns:w="${wordNS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="${xml(typography.fontFamily)}" w:hAnsi="${xml(typography.fontFamily)}"/><w:color w:val="${color(typography.textColor)}"/><w:sz w:val="${Math.round(typography.fontSize * 1.5)}"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="${Math.round(typography.lineHeight * 240)}" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>${[1, 2, 3].map((level) => `<w:style w:type="paragraph" w:styleId="Heading${level}"><w:name w:val="heading ${level}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="${level - 1}"/></w:pPr><w:rPr><w:b/><w:sz w:val="${[40, 32, 28][level - 1]}"/></w:rPr></w:style>`).join("")}</w:styles>`,
  );
  if (comments.length)
    zip.file(
      "word/comments.xml",
      header +
        `<w:comments xmlns:w="${wordNS}">${comments.map((comment, index) => `<w:comment w:id="${index}" w:author="${xml(comment.author)}"><w:p>${wordRun(comment.comment)}</w:p></w:comment>`).join("")}</w:comments>`,
    );
  const smartArt = document.paragraphs.flatMap((paragraph, index) =>
    paragraph.type === "smartart" && paragraph.smartArt
      ? [{ index, id: paragraph.id, smartArt: paragraph.smartArt }]
      : [],
  );
  if (smartArt.length)
    zip.file(
      "customXml/officeStudioSmartArt.xml",
      `${header}<officeStudioSmartArt xmlns="urn:office-studio:smartart:1">${xml(JSON.stringify({ version: 1, blocks: smartArt }))}</officeStudioSmartArt>`,
    );
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}

function serializeExcel(
  document: Extract<OfficeDocument, { type: "excel" }>["data"],
): Uint8Array {
  const workbook = XLSX.utils.book_new();
  const usedNames = new Set<string>();
  for (const source of document.sheets) {
    const sheet: XLSX.WorkSheet = {};
    let maxRow = 0;
    let maxCol = 0;
    for (const [coord, cell] of Object.entries(source.cells)) {
      if (!/^[A-Z]+[1-9][0-9]*$/.test(coord)) continue;
      const position = XLSX.utils.decode_cell(coord);
      maxRow = Math.max(maxRow, position.r);
      maxCol = Math.max(maxCol, position.c);
      const value = cell.value;
      const entry: XLSX.CellObject = {
        t:
          typeof value === "number"
            ? "n"
            : typeof value === "boolean"
              ? "b"
              : "s",
        v: value ?? "",
      };
      if (cell.formula) entry.f = cell.formula.replace(/^=/, "");
      const digits = Math.max(0, Math.min(10, cell.format?.decimalPlaces ?? 2));
      const decimal = digits ? "." + "0".repeat(digits) : "";
      if (cell.format?.numberFormat === "currency")
        entry.z = "$#,##0" + decimal;
      if (cell.format?.numberFormat === "percent")
        entry.z = "0" + decimal + "%";
      if (cell.format?.numberFormat === "number") entry.z = "#,##0" + decimal;
      sheet[coord] = entry;
    }
    sheet["!ref"] = XLSX.utils.encode_range({
      s: { r: 0, c: 0 },
      e: { r: maxRow, c: maxCol },
    });
    sheet["!cols"] = source.columns.map((column) => ({ wpx: column.width }));
    const stem =
      source.name
        .replace(/[\\/?:*\[\]]/g, "_")
        .replace(/^'|'$/g, "")
        .slice(0, 31) || "Sheet";
    let name = stem;
    let suffix = 2;
    while (usedNames.has(name.toLowerCase())) {
      const tail = ` (${suffix++})`;
      name = stem.slice(0, 31 - tail.length) + tail;
    }
    usedNames.add(name.toLowerCase());
    XLSX.utils.book_append_sheet(workbook, sheet, name);
  }
  if (!workbook.SheetNames.length)
    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.aoa_to_sheet([[""]]),
      "Sheet1",
    );
  return new Uint8Array(
    XLSX.write(workbook, {
      type: "array",
      bookType: "xlsx",
      compression: true,
    }),
  );
}

const group =
  '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';
const clrMap =
  '<p:clrMap accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" bg1="lt1" bg2="lt2" folHlink="folHlink" hlink="hlink" tx1="dk1" tx2="dk2"/>';
const pptRoot = `xmlns:a="${drawNS}" xmlns:r="${officeRel}" xmlns:p="${pptNS}"`;
const emu = (pixels: number) => Math.round(pixels * 9525);
const solid = (value: string | undefined, fallback = "FFFFFF") =>
  !value || value === "transparent"
    ? "<a:noFill/>"
    : `<a:solidFill><a:srgbClr val="${color(value, fallback)}"/></a:solidFill>`;

function shapeXml(object: PptShapeObject, index: number): string {
  const text =
    object.kind === "metric"
      ? [object.title, object.metricValue, object.metricLabel, object.subtitle]
          .filter(Boolean)
          .join("\n")
      : object.text || object.title || "";
  const kind =
    object.shapeKind === "circle"
      ? "ellipse"
      : object.shapeKind === "rounded-rect" || object.shapeKind === "card"
        ? "roundRect"
        : "rect";
  const align = { left: "l", center: "ctr", right: "r" }[
    object.align || "left"
  ];
  const paragraphs = text
    .split("\n")
    .map(
      (line) =>
        `<a:p><a:pPr algn="${align}"/><a:r><a:rPr lang="en-US" sz="${Math.round(object.fontSize * 75)}" b="${object.fontWeight && object.fontWeight !== "normal" ? "1" : "0"}">${solid(object.color, "24292F")}<a:latin typeface="Aptos"/></a:rPr><a:t>${xml(line)}</a:t></a:r><a:endParaRPr lang="en-US"/></a:p>`,
    )
    .join("");
  return `<p:sp><p:nvSpPr><p:cNvPr id="${index + 2}" name="${xml(object.title || `Object ${index + 1}`)}"/><p:cNvSpPr${object.kind === "text" ? ' txBox="1"' : ""}/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm${object.rotation ? ` rot="${Math.round(object.rotation * 60000)}"` : ""}><a:off x="${emu(object.x)}" y="${emu(object.y)}"/><a:ext cx="${emu(Math.max(1, object.width))}" cy="${emu(Math.max(1, object.height))}"/></a:xfrm><a:prstGeom prst="${kind}"><a:avLst/></a:prstGeom>${solid(object.fill)}<a:ln w="${emu(object.strokeWidth || 1)}">${solid(object.stroke)}</a:ln></p:spPr><p:txBody><a:bodyPr wrap="square"/><a:lstStyle/>${paragraphs}</p:txBody></p:sp>`;
}

function themeXml(accent: string): string {
  const colors = [
    ["dk1", "000000"],
    ["lt1", "FFFFFF"],
    ["dk2", "1F2937"],
    ["lt2", "F3F4F6"],
    ["accent1", color(accent, "007ACC")],
    ["accent2", "10B981"],
    ["accent3", "F59E0B"],
    ["accent4", "8B5CF6"],
    ["accent5", "EF4444"],
    ["accent6", "06B6D4"],
    ["hlink", "0563C1"],
    ["folHlink", "954F72"],
  ];
  const fill = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
  return (
    header +
    `<a:theme xmlns:a="${drawNS}" name="Office Studio"><a:themeElements><a:clrScheme name="Office Studio">${colors.map(([name, value]) => `<a:${name}><a:srgbClr val="${value}"/></a:${name}>`).join("")}</a:clrScheme><a:fontScheme name="Office Studio"><a:majorFont><a:latin typeface="Aptos Display"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Aptos"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Office Studio"><a:fillStyleLst>${fill.repeat(3)}</a:fillStyleLst><a:lnStyleLst>${[9525, 19050, 28575].map((width) => `<a:ln w="${width}" cap="flat" cmpd="sng" algn="ctr">${fill}<a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>`).join("")}</a:lnStyleLst><a:effectStyleLst>${"<a:effectStyle><a:effectLst/></a:effectStyle>".repeat(3)}</a:effectStyleLst><a:bgFillStyleLst>${fill.repeat(3)}</a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`
  );
}

async function serializePowerPoint(
  document: PptDocumentModel,
): Promise<Uint8Array> {
  const zip = new JSZip();
  const parts: [string, string][] = [
    ["ppt/presentation.xml", "presentationml.presentation.main"],
    ["ppt/presProps.xml", "presentationml.presProps"],
    ["ppt/slideMasters/slideMaster1.xml", "presentationml.slideMaster"],
    ["ppt/slideLayouts/slideLayout1.xml", "presentationml.slideLayout"],
    ["ppt/theme/theme1.xml", "theme"],
    ["ppt/notesMasters/notesMaster1.xml", "presentationml.notesMaster"],
  ];
  zip.file(
    "_rels/.rels",
    relationships([
      { id: "rId1", type: "officeDocument", target: "ppt/presentation.xml" },
    ]),
  );
  const presentationRels = [
    {
      id: "rId1",
      type: "slideMaster",
      target: "slideMasters/slideMaster1.xml",
    },
    {
      id: "rId2",
      type: "notesMaster",
      target: "notesMasters/notesMaster1.xml",
    },
    { id: "rId3", type: "presProps", target: "presProps.xml" },
  ];
  document.slides.forEach((slide, index) => {
    const number = index + 1;
    parts.push(
      [`ppt/slides/slide${number}.xml`, "presentationml.slide"],
      [`ppt/notesSlides/notesSlide${number}.xml`, "presentationml.notesSlide"],
    );
    presentationRels.push({
      id: `rId${index + 4}`,
      type: "slide",
      target: `slides/slide${number}.xml`,
    });
    zip.file(
      `ppt/slides/slide${number}.xml`,
      header +
        `<p:sld ${pptRoot}><p:cSld name="${xml(slide.title)}"><p:bg><p:bgPr>${solid(slide.background, "FFFFFF")}<a:effectLst/></p:bgPr></p:bg><p:spTree>${group}${[
          ...slide.objects,
        ]
          .filter((object) => object.visible !== false)
          .sort((a, b) => a.zIndex - b.zIndex)
          .map(shapeXml)
          .join(
            "",
          )}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>${slide.transition === "fade" ? "<p:transition><p:fade/></p:transition>" : ""}</p:sld>`,
    );
    zip.file(
      `ppt/slides/_rels/slide${number}.xml.rels`,
      relationships([
        {
          id: "rId1",
          type: "slideLayout",
          target: "../slideLayouts/slideLayout1.xml",
        },
        {
          id: "rId2",
          type: "notesSlide",
          target: `../notesSlides/notesSlide${number}.xml`,
        },
      ]),
    );
    zip.file(
      `ppt/notesSlides/notesSlide${number}.xml`,
      header +
        `<p:notes ${pptRoot}><p:cSld><p:spTree>${group}<p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>${slide.notes
          .split("\n")
          .map((line) => `<a:p><a:r><a:t>${xml(line)}</a:t></a:r></a:p>`)
          .join(
            "",
          )}</p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`,
    );
    zip.file(
      `ppt/notesSlides/_rels/notesSlide${number}.xml.rels`,
      relationships([
        {
          id: "rId1",
          type: "notesMaster",
          target: "../notesMasters/notesMaster1.xml",
        },
        { id: "rId2", type: "slide", target: `../slides/slide${number}.xml` },
      ]),
    );
  });
  zip.file("[Content_Types].xml", types(parts));
  zip.file(
    "ppt/presentation.xml",
    header +
      `<p:presentation ${pptRoot}><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:notesMasterIdLst><p:notesMasterId r:id="rId2"/></p:notesMasterIdLst><p:sldIdLst>${document.slides.map((_, index) => `<p:sldId id="${index + 256}" r:id="rId${index + 4}"/>`).join("")}</p:sldIdLst><p:sldSz cx="${emu(880)}" cy="${emu(495)}" type="screen16x9"/><p:notesSz cx="6858000" cy="9144000"/><p:defaultTextStyle/></p:presentation>`,
  );
  zip.file("ppt/_rels/presentation.xml.rels", relationships(presentationRels));
  zip.file("ppt/presProps.xml", header + `<p:presentationPr ${pptRoot}/>`);
  zip.file(
    "ppt/slideMasters/slideMaster1.xml",
    header +
      `<p:sldMaster ${pptRoot}><p:cSld><p:spTree>${group}</p:spTree></p:cSld>${clrMap}<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles></p:sldMaster>`,
  );
  zip.file(
    "ppt/slideMasters/_rels/slideMaster1.xml.rels",
    relationships([
      {
        id: "rId1",
        type: "slideLayout",
        target: "../slideLayouts/slideLayout1.xml",
      },
      { id: "rId2", type: "theme", target: "../theme/theme1.xml" },
    ]),
  );
  zip.file(
    "ppt/slideLayouts/slideLayout1.xml",
    header +
      `<p:sldLayout ${pptRoot} type="blank" preserve="1"><p:cSld name="Blank"><p:spTree>${group}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`,
  );
  zip.file(
    "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
    relationships([
      {
        id: "rId1",
        type: "slideMaster",
        target: "../slideMasters/slideMaster1.xml",
      },
    ]),
  );
  zip.file(
    "ppt/notesMasters/notesMaster1.xml",
    header +
      `<p:notesMaster ${pptRoot}><p:cSld><p:spTree>${group}</p:spTree></p:cSld>${clrMap}<p:notesStyle/></p:notesMaster>`,
  );
  zip.file(
    "ppt/notesMasters/_rels/notesMaster1.xml.rels",
    relationships([
      { id: "rId1", type: "theme", target: "../theme/theme1.xml" },
    ]),
  );
  zip.file("ppt/theme/theme1.xml", themeXml(document.accentColor));
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}

/** Exports the supported editor model, not the original Office package or unsupported objects. */
export async function serializeDocument(
  document: OfficeDocument,
): Promise<SerializedDocument> {
  if (document.type === "word")
    return {
      data: await serializeWord(document.data),
      mimeType:
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      extension: "docx",
    };
  if (document.type === "excel")
    return {
      data: serializeExcel(document.data),
      mimeType:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      extension: "xlsx",
    };
  if (document.type === "powerpoint")
    return {
      data: await serializePowerPoint(document.data),
      mimeType:
        "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      extension: "pptx",
    };
  return {
    data: new TextEncoder().encode(document.data.content),
    mimeType: "text/plain;charset=utf-8",
    extension: {
      typescript: "ts",
      javascript: "js",
      python: "py",
      markdown: "md",
      json: "json",
    }[document.data.language],
  };
}

export async function exportDocument(
  document: OfficeDocument,
  filename: string,
): Promise<void> {
  const output = await serializeDocument(document);
  const basename =
    filename.replace(/[\\/]/g, "_").replace(/\.[^.]+$/, "") || "Untitled";
  const blob = new Blob([new Uint8Array(output.data)], {
    type: output.mimeType,
  });
  const url = URL.createObjectURL(blob);
  const anchor = window.document.createElement("a");
  anchor.href = url;
  anchor.download =
    document.type === "code"
      ? filename.replace(/[\\/]/g, "_")
      : `${basename}.${output.extension}`;
  window.document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
