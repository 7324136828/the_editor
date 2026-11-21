//! .docx read/write.
//!
//! READ (`parse_docx`): the main document part (located through
//! `_rels/.rels`) is parsed into a small DOM ([`xml`]) and walked into the
//! shared WordDoc model. Also consulted: styles.xml (style names, `basedOn`
//! chains, outline levels and style-level numbering, so localized headings
//! and style-based lists work), numbering.xml, the part relationships,
//! word/media (images become `img` blocks with data: URLs), footnotes /
//! endnotes (moved to numbered notes at the end), docProps/core.xml (title,
//! creator -> author, subject), the body `sectPr` (page size and margins,
//! kept in `meta.page`) and its default header / footer parts (`meta.header`
//! / `meta.footer` = `{runs, align}`; PAGE / NUMPAGES fields become the literal
//! tokens "{PAGE}" / "{NUMPAGES}").
//! Fields show their RESULT text only; hyperlink fields become links.
//! Everything that cannot be represented is reported in `meta.warnings`
//! (main.rs lifts these into the /api/parse response).
//!
//! WRITE (`write_docx`): emits a complete package Word opens without repair:
//! [Content_Types].xml, _rels/.rels, docProps/{core,app}.xml,
//! word/{document,styles,numbering,settings}.xml, word/_rels/document.xml.rels
//! and word/media/* for images, plus word/header1.xml / word/footer1.xml
//! (with real PAGE / NUMPAGES fields) when `meta.header` / `meta.footer` have
//! text. Lists use real numbering definitions
//! (ordered lists restart per list), links use hyperlink relationships (or
//! `w:anchor` for "#bookmark" links), rPr children follow the CT_RPr schema
//! order, XML-invalid characters are stripped, every table cell ends with a
//! paragraph and the page setup comes from `meta.page` (default US Letter).
//!
//! `meta.page` (read and written): `{ size?: "A4"|"Letter"|"Legal"|"A3"|"A5"|
//! "Executive"|"custom", orientation?: "portrait"|"landscape", width?, height?,
//! margins?: {top,right,bottom,left,header?,footer?} }` — all lengths in twips
//! (1/1440 inch). A plain `meta.pageSize: "A4"` string is also accepted.

use crate::model::{Block, Run, TableCell, WordDoc};
use serde_json::{json, Map, Value};
use std::borrow::Cow;
use std::collections::{BTreeSet, HashMap};
use std::io::{Cursor, Read, Write};
use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

use self::xml::Node;

const W_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG_REL_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const WP_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const PIC_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/picture";

/// Largest single XML part we are willing to inflate (zip-bomb guard; main.rs
/// additionally runs limits::check_zip on the whole container).
const MAX_XML_PART: u64 = 256 * 1024 * 1024;
/// Largest single embedded image we turn into a data: URL.
const MAX_IMAGE: u64 = 32 * 1024 * 1024;
/// Total data-URL characters of all images in one parse.
const MAX_MEDIA_TOTAL: usize = 160 * 1024 * 1024;

// User-facing warning texts (deduplicated per document).
const W_HF_VARIANTS: &str = "Only the main header and footer are kept (first-page and even-page versions were dropped).";
const W_HF_SECTIONS: &str = "Headers and footers of earlier sections were dropped; only the last section's are kept.";
const W_HF_SIMPLIFIED: &str = "Pictures and tables in headers and footers were simplified to text.";
const W_COMMENTS: &str = "Comments are not shown and are not kept when saving.";
const W_TRACKED: &str = "Tracked changes were accepted: insertions are shown as normal text and deletions are hidden.";
const W_NOTES: &str = "Footnotes/endnotes were converted to numbered notes at the end of the document.";
const W_FIELDS: &str = "Fields (page numbers, dates, tables of contents, cross-references) show their last calculated text and are saved as plain text.";
const W_TEXTBOX: &str = "Text boxes were converted to normal paragraphs placed after the paragraph that contained them.";
const W_SHAPES: &str = "Some shapes or drawings could not be displayed and are not kept when saving.";
const W_CHARTS: &str = "Charts and SmartArt graphics could not be displayed and are not kept when saving.";
const W_LINKED_IMG: &str = "Linked (not embedded) pictures could not be loaded.";
const W_VMERGE: &str = "Vertically merged table cells were split into separate cells.";
const W_SECTIONS: &str = "The document has several sections; only the last section's page size and margins are kept.";
const W_EQUATIONS: &str = "Equations are shown as plain text.";
const W_OBJECTS: &str = "Embedded objects (OLE) are shown as their preview picture only.";
const W_ALTCHUNK: &str = "Embedded HTML/RTF fragments (altChunk) could not be displayed.";
const W_IMG_FORMAT: &str = "Some pictures use formats browsers cannot display (EMF/WMF/TIFF); they are kept when saving as .docx.";
const W_IMG_LIMIT: &str = "Some pictures were too large to load and were left out.";
const W_HIDDEN: &str = "Hidden text is shown as normal text.";

// ============================ shared helpers ============================

/// True for characters allowed in XML 1.0 documents.
pub fn xml_char_ok(c: char) -> bool {
    matches!(c, '\t' | '\n' | '\r')
        || ('\u{20}'..='\u{D7FF}').contains(&c)
        || ('\u{E000}'..='\u{FFFD}').contains(&c)
        || (c as u32) >= 0x10000
}

/// Escape text content for XML, dropping characters XML 1.0 forbids.
pub fn xml_text(s: &str) -> String {
    let mut o = String::with_capacity(s.len() + 8);
    for c in s.chars() {
        match c {
            '&' => o.push_str("&amp;"),
            '<' => o.push_str("&lt;"),
            '>' => o.push_str("&gt;"),
            c if xml_char_ok(c) => o.push(c),
            _ => {}
        }
    }
    o
}

/// Escape a double-quoted attribute value (whitespace controls become
/// character references so they survive attribute normalization).
pub fn xml_attr(s: &str) -> String {
    let mut o = String::with_capacity(s.len() + 8);
    for c in s.chars() {
        match c {
            '&' => o.push_str("&amp;"),
            '<' => o.push_str("&lt;"),
            '>' => o.push_str("&gt;"),
            '"' => o.push_str("&quot;"),
            '\n' => o.push_str("&#10;"),
            '\r' => o.push_str("&#13;"),
            '\t' => o.push_str("&#9;"),
            c if xml_char_ok(c) => o.push(c),
            _ => {}
        }
    }
    o
}

/// Add a warning to `meta.warnings` (deduplicated).
pub fn push_warning(meta: &mut Map<String, Value>, msg: &str) {
    let entry = meta.entry("warnings").or_insert_with(|| Value::Array(Vec::new()));
    if !entry.is_array() {
        *entry = Value::Array(Vec::new());
    }
    if let Value::Array(a) = entry {
        if !a.iter().any(|v| v.as_str() == Some(msg)) {
            a.push(Value::String(msg.to_string()));
        }
    }
}

const B64: &[u8; 64] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";

pub fn base64_encode(data: &[u8]) -> String {
    let mut out = String::with_capacity(data.len().div_ceil(3) * 4);
    for chunk in data.chunks(3) {
        let b1 = *chunk.get(1).unwrap_or(&0);
        let b2 = *chunk.get(2).unwrap_or(&0);
        let n = ((chunk[0] as u32) << 16) | ((b1 as u32) << 8) | b2 as u32;
        out.push(B64[(n >> 18) as usize & 63] as char);
        out.push(B64[(n >> 12) as usize & 63] as char);
        out.push(if chunk.len() > 1 { B64[(n >> 6) as usize & 63] as char } else { '=' });
        out.push(if chunk.len() > 2 { B64[n as usize & 63] as char } else { '=' });
    }
    out
}

/// Lenient base64 decoder (standard or URL-safe alphabet, whitespace ignored,
/// padding optional). Returns None on any other character.
pub fn base64_decode(s: &str) -> Option<Vec<u8>> {
    let mut out = Vec::with_capacity(s.len() / 4 * 3 + 3);
    let mut acc: u32 = 0;
    let mut bits = 0u32;
    for c in s.bytes() {
        let v = match c {
            b'A'..=b'Z' => c - b'A',
            b'a'..=b'z' => c - b'a' + 26,
            b'0'..=b'9' => c - b'0' + 52,
            b'+' | b'-' => 62,
            b'/' | b'_' => 63,
            b'=' => break,
            b' ' | b'\n' | b'\r' | b'\t' => continue,
            _ => return None,
        };
        acc = (acc << 6) | v as u32;
        bits += 6;
        if bits >= 8 {
            bits -= 8;
            out.push((acc >> bits) as u8);
            acc &= (1 << bits) - 1;
        }
    }
    Some(out)
}

/// Percent-decode (invalid escapes are kept literally).
pub fn percent_decode(s: &str) -> Vec<u8> {
    let b = s.as_bytes();
    let mut out = Vec::with_capacity(b.len());
    let mut i = 0;
    while i < b.len() {
        if b[i] == b'%' && i + 2 < b.len() {
            let h = std::str::from_utf8(&b[i + 1..i + 3]).ok().and_then(|h| u8::from_str_radix(h, 16).ok());
            if let Some(v) = h {
                out.push(v);
                i += 3;
                continue;
            }
        }
        out.push(b[i]);
        i += 1;
    }
    out
}

/// Decode a `data:` URL into (mime type, bytes).
pub fn parse_data_url(src: &str) -> Option<(String, Vec<u8>)> {
    let s = src.trim();
    if !s.get(..5).map(|p| p.eq_ignore_ascii_case("data:")).unwrap_or(false) {
        return None;
    }
    let (meta, data) = s[5..].split_once(',')?;
    let mime = meta.split(';').next().unwrap_or("").trim().to_ascii_lowercase();
    let is_b64 = meta.split(';').any(|p| p.trim().eq_ignore_ascii_case("base64"));
    let bytes = if is_b64 {
        if data.contains('%') {
            base64_decode(&String::from_utf8_lossy(&percent_decode(data)))?
        } else {
            base64_decode(data)?
        }
    } else {
        percent_decode(data)
    };
    Some((if mime.is_empty() { "text/plain".into() } else { mime }, bytes))
}

pub fn data_url(mime: &str, bytes: &[u8]) -> String {
    format!("data:{mime};base64,{}", base64_encode(bytes))
}

/// Identify an image by its magic bytes.
pub fn sniff_image(b: &[u8]) -> Option<&'static str> {
    if b.starts_with(b"\x89PNG\r\n\x1a\n") { return Some("image/png"); }
    if b.starts_with(&[0xFF, 0xD8, 0xFF]) { return Some("image/jpeg"); }
    if b.starts_with(b"GIF87a") || b.starts_with(b"GIF89a") { return Some("image/gif"); }
    if b.starts_with(b"BM") && b.len() > 26 { return Some("image/bmp"); }
    if b.len() > 12 && &b[..4] == b"RIFF" && &b[8..12] == b"WEBP" { return Some("image/webp"); }
    if b.starts_with(b"II*\0") || b.starts_with(b"MM\0*") { return Some("image/tiff"); }
    if b.len() > 44 && b[..4] == [1, 0, 0, 0] && &b[40..44] == b" EMF" { return Some("image/x-emf"); }
    if b.starts_with(&[0xD7, 0xCD, 0xC6, 0x9A]) || b.starts_with(&[1, 0, 9, 0]) || b.starts_with(&[2, 0, 9, 0]) {
        return Some("image/x-wmf");
    }
    if b.starts_with(&[0, 0, 1, 0]) && b.len() > 6 { return Some("image/x-icon"); }
    let head = String::from_utf8_lossy(&b[..b.len().min(1024)]).to_ascii_lowercase();
    if head.trim_start().starts_with("<svg") || (head.trim_start().starts_with("<?xml") && head.contains("<svg")) {
        return Some("image/svg+xml");
    }
    None
}

/// File extension for an image mime type (None = not an image we can embed).
pub fn image_ext(mime: &str) -> Option<&'static str> {
    Some(match mime.trim().to_ascii_lowercase().as_str() {
        "image/png" | "image/apng" => "png",
        "image/jpeg" | "image/jpg" | "image/pjpeg" => "jpeg",
        "image/gif" => "gif",
        "image/bmp" | "image/x-bmp" | "image/x-ms-bmp" => "bmp",
        "image/tiff" | "image/tif" => "tiff",
        "image/x-emf" | "image/emf" => "emf",
        "image/x-wmf" | "image/wmf" => "wmf",
        "image/svg+xml" => "svg",
        "image/webp" => "webp",
        "image/x-icon" | "image/vnd.microsoft.icon" => "ico",
        _ => return None,
    })
}

pub fn mime_for_ext(ext: &str) -> &'static str {
    match ext.to_ascii_lowercase().as_str() {
        "png" => "image/png",
        "jpg" | "jpeg" | "jpe" | "jfif" => "image/jpeg",
        "gif" => "image/gif",
        "bmp" | "dib" => "image/bmp",
        "tif" | "tiff" => "image/tiff",
        "emf" => "image/x-emf",
        "wmf" => "image/x-wmf",
        "svg" => "image/svg+xml",
        "webp" => "image/webp",
        "ico" => "image/x-icon",
        _ => "application/octet-stream",
    }
}

/// Pixel dimensions from an image header (PNG, JPEG, GIF, BMP).
pub fn image_dims(b: &[u8]) -> Option<(u32, u32)> {
    let be32 = |o: usize| -> Option<u32> { b.get(o..o + 4).map(|s| u32::from_be_bytes([s[0], s[1], s[2], s[3]])) };
    let le16 = |o: usize| -> Option<u32> { b.get(o..o + 2).map(|s| u16::from_le_bytes([s[0], s[1]]) as u32) };
    let be16 = |o: usize| -> Option<u32> { b.get(o..o + 2).map(|s| u16::from_be_bytes([s[0], s[1]]) as u32) };
    let dims = match sniff_image(b)? {
        "image/png" => (be32(16)?, be32(20)?),
        "image/gif" => (le16(6)?, le16(8)?),
        "image/bmp" => {
            let w = i32::from_le_bytes(b.get(18..22)?.try_into().ok()?);
            let h = i32::from_le_bytes(b.get(22..26)?.try_into().ok()?);
            (w.unsigned_abs(), h.unsigned_abs())
        }
        "image/jpeg" => {
            let mut i = 2usize;
            loop {
                while *b.get(i)? != 0xFF { i += 1; }
                while *b.get(i)? == 0xFF { i += 1; }
                let m = *b.get(i)?;
                i += 1;
                if matches!(m, 0xD8 | 0x01 | 0xD0..=0xD7) { continue; }
                let len = be16(i)? as usize;
                if (0xC0..=0xCF).contains(&m) && !matches!(m, 0xC4 | 0xC8 | 0xCC) {
                    break (be16(i + 5)?, be16(i + 3)?);
                }
                if len < 2 { return None; }
                i += len;
            }
        }
        _ => return None,
    };
    if dims.0 == 0 || dims.1 == 0 { None } else { Some(dims) }
}

/// Fonts treated as "code" fonts.
pub fn is_mono_font(f: &str) -> bool {
    let l = f.trim().trim_matches(|c| c == '"' || c == '\'').to_ascii_lowercase();
    matches!(
        l.as_str(),
        "consolas" | "courier" | "courier new" | "menlo" | "monaco" | "lucida console" | "source code pro"
            | "dejavu sans mono" | "liberation mono" | "cascadia code" | "cascadia mono" | "fira code"
            | "monospace" | "jetbrains mono" | "roboto mono" | "ubuntu mono" | "sf mono"
    )
}

/// Reject links that must never become clickable (script / data schemes).
pub fn safe_link(s: &str) -> Option<String> {
    let t: String = s.trim().chars().filter(|c| xml_char_ok(*c) && *c != '\n' && *c != '\r' && *c != '\t').collect();
    if t.is_empty() {
        return None;
    }
    let scheme: String = t.chars().take_while(|c| *c != ':' && *c != '/' && *c != '#' && *c != '?').collect::<String>().to_ascii_lowercase();
    if t.len() > scheme.len() && t[scheme.len()..].starts_with(':') && matches!(scheme.trim(), "javascript" | "vbscript" | "data") {
        return None;
    }
    Some(t)
}

/// Normalize "#RRGGBB" / "#RGB" / "RRGGBB" / "rgb(r,g,b)" to "RRGGBB".
pub fn norm_hex(s: &str) -> Option<String> {
    let t = s.trim();
    if let Some(inner) = t.strip_prefix("rgb(").or_else(|| t.strip_prefix("rgba(")).and_then(|r| r.strip_suffix(')')) {
        let parts: Vec<u8> = inner.split(',').take(3).filter_map(|p| p.trim().parse::<f64>().ok()).map(|v| v.clamp(0.0, 255.0) as u8).collect();
        if parts.len() == 3 {
            return Some(format!("{:02X}{:02X}{:02X}", parts[0], parts[1], parts[2]));
        }
        return None;
    }
    let h = t.trim_start_matches('#');
    if h.len() == 6 && h.chars().all(|c| c.is_ascii_hexdigit()) {
        return Some(h.to_ascii_uppercase());
    }
    if h.len() == 3 && h.chars().all(|c| c.is_ascii_hexdigit()) {
        return Some(h.chars().flat_map(|c| [c, c]).collect::<String>().to_ascii_uppercase());
    }
    None
}

const HIGHLIGHTS: &[(&str, &str)] = &[
    ("yellow", "FFFF00"), ("green", "00FF00"), ("cyan", "00FFFF"), ("magenta", "FF00FF"),
    ("blue", "0000FF"), ("red", "FF0000"), ("darkBlue", "000080"), ("darkCyan", "008080"),
    ("darkGreen", "008000"), ("darkMagenta", "800080"), ("darkRed", "800000"), ("darkYellow", "808000"),
    ("darkGray", "808080"), ("lightGray", "C0C0C0"), ("black", "000000"), ("white", "FFFFFF"),
];

pub fn highlight_hex(name: &str) -> Option<String> {
    HIGHLIGHTS.iter().find(|(n, _)| n.eq_ignore_ascii_case(name)).map(|(_, h)| format!("#{h}"))
}

pub fn highlight_name(hex: &str) -> Option<&'static str> {
    HIGHLIGHTS.iter().find(|(_, h)| h.eq_ignore_ascii_case(hex)).map(|(n, _)| *n)
}

/// Same formatting (everything but the text)?
pub fn same_format(a: &Run, b: &Run) -> bool {
    a.b == b.b && a.i == b.i && a.u == b.u && a.s == b.s && a.sup == b.sup && a.sub == b.sub
        && a.color == b.color && a.size == b.size && a.font == b.font && a.link == b.link && a.hl == b.hl
}

/// Merge adjacent runs with identical formatting; drop empty runs.
pub fn merge_runs(runs: Vec<Run>) -> Vec<Run> {
    let mut out: Vec<Run> = Vec::with_capacity(runs.len());
    for r in runs {
        if r.t.is_empty() {
            continue;
        }
        if let Some(last) = out.last_mut() {
            if same_format(last, &r) {
                last.t.push_str(&r.t);
                continue;
            }
        }
        out.push(r);
    }
    out
}

/// Inline content of a source paragraph: text runs, or block-level items
/// (images, page breaks) that the model only allows between paragraphs.
pub enum Inline {
    Run(Run),
    Block(Block),
}

/// Turn the inline content of one source paragraph into model blocks. The
/// paragraph is split around images / page breaks; `make(runs, first)` builds
/// each text block (`first` = false for continuation pieces). A paragraph
/// without any block items always yields exactly one (possibly empty) block.
pub fn split_paragraph(items: Vec<Inline>, out: &mut Vec<Block>, make: &mut dyn FnMut(Vec<Run>, bool) -> Block) {
    if !items.iter().any(|i| matches!(i, Inline::Block(_))) {
        let runs = items.into_iter().filter_map(|i| if let Inline::Run(r) = i { Some(r) } else { None }).collect();
        out.push(make(merge_runs(runs), true));
        return;
    }
    fn emit(cur: &mut Vec<Run>, first: &mut bool, out: &mut Vec<Block>, make: &mut dyn FnMut(Vec<Run>, bool) -> Block) {
        let runs = merge_runs(std::mem::take(cur));
        if runs.iter().any(|r| !r.t.trim().is_empty()) {
            out.push(make(runs, *first));
            *first = false;
        }
    }
    let mut cur = Vec::new();
    let mut first = true;
    for it in items {
        match it {
            Inline::Run(r) => cur.push(r),
            Inline::Block(b) => {
                emit(&mut cur, &mut first, out, make);
                out.push(b);
            }
        }
    }
    emit(&mut cur, &mut first, out, make);
}

/// Resolve a block for writing: typed blocks are borrowed unchanged;
/// `Block::Other` (JSON that did not match the contract exactly, e.g. a float
/// image width, a negative list level or an unknown block type) is coerced
/// into the closest typed block so writers never silently drop content.
/// Returns None only for blocks that carry no content at all.
pub fn resolve_block(b: &Block) -> Option<Cow<'_, Block>> {
    match b {
        Block::Other(v) => normalize_value_block(v).map(Cow::Owned),
        _ => Some(Cow::Borrowed(b)),
    }
}

fn num_of(v: Option<&Value>) -> Option<f64> {
    match v? {
        Value::Number(n) => n.as_f64(),
        Value::String(s) => s.trim().trim_end_matches("px").trim_end_matches("pt").trim().parse().ok(),
        _ => None,
    }
    .filter(|f: &f64| f.is_finite())
}

fn dim_of(v: Option<&Value>) -> Option<u32> {
    num_of(v).filter(|f| *f >= 1.0).map(|f| f.round().min(100_000.0) as u32)
}

/// Best-effort runs from loosely-typed JSON.
pub fn lenient_runs(v: Option<&Value>) -> Vec<Run> {
    let Some(Value::Array(a)) = v else {
        return Vec::new();
    };
    a.iter()
        .filter_map(|r| {
            if let Value::String(s) = r {
                return Some(Run::text(s.clone()));
            }
            if let Ok(run) = serde_json::from_value::<Run>(r.clone()) {
                return Some(run);
            }
            let o = r.as_object()?;
            let t = match o.get("t").or_else(|| o.get("text")) {
                Some(Value::String(s)) => s.clone(),
                Some(Value::Number(n)) => n.to_string(),
                Some(Value::Bool(b)) => b.to_string(),
                _ => String::new(),
            };
            let truthy = |k: &str| {
                o.get(k)
                    .map(|x| match x {
                        Value::Bool(b) => *b,
                        Value::Number(n) => n.as_f64().unwrap_or(0.0) != 0.0,
                        Value::String(s) => !s.is_empty() && s != "0" && s != "false",
                        _ => false,
                    })
                    .unwrap_or(false)
            };
            let st = |k: &str| o.get(k).and_then(Value::as_str).filter(|x| !x.is_empty()).map(String::from);
            Some(Run {
                t,
                b: truthy("b"),
                i: truthy("i"),
                u: truthy("u"),
                s: truthy("s"),
                sup: truthy("sup"),
                sub: truthy("sub"),
                color: st("color"),
                size: num_of(o.get("size")).filter(|f| *f > 0.0 && *f < 2000.0),
                font: st("font"),
                link: st("link"),
                hl: st("hl"),
            })
        })
        .collect()
}

fn collect_strings(v: &Value, out: &mut String) {
    match v {
        Value::Object(o) => {
            for (k, x) in o {
                if (k == "t" || k == "text") && x.is_string() {
                    out.push_str(x.as_str().unwrap_or(""));
                } else if x.is_object() || x.is_array() {
                    collect_strings(x, out);
                }
            }
        }
        Value::Array(a) => {
            for x in a {
                collect_strings(x, out);
            }
        }
        _ => {}
    }
}

/// Coerce a loosely-typed block JSON value into a typed block.
pub fn normalize_value_block(v: &Value) -> Option<Block> {
    if let Value::String(s) = v {
        return Some(Block::P { runs: vec![Run::text(s.clone())], align: None });
    }
    let o = v.as_object()?;
    let ty = o.get("type").and_then(Value::as_str).unwrap_or("").to_ascii_lowercase();
    let runs = lenient_runs(o.get("runs"));
    let align = o
        .get("align")
        .and_then(Value::as_str)
        .filter(|a| matches!(*a, "left" | "center" | "right" | "justify"))
        .map(str::to_string);
    Some(match ty.as_str() {
        "p" | "paragraph" | "div" => Block::P { runs, align },
        "h1" => Block::H1 { runs, align },
        "h2" => Block::H2 { runs, align },
        "h3" => Block::H3 { runs, align },
        "h4" => Block::H4 { runs, align },
        "h5" => Block::H5 { runs, align },
        "h6" => Block::H6 { runs, align },
        "quote" | "blockquote" => Block::Quote { runs, align },
        "code" | "pre" => Block::Code { runs, align },
        "list" | "li" => Block::List {
            list: if o.get("list").and_then(Value::as_str) == Some("ol") { "ol".into() } else { "ul".into() },
            level: num_of(o.get("level")).map(|f| f.clamp(0.0, 8.0) as u32).unwrap_or(0),
            runs,
        },
        "img" | "image" => Block::Img {
            src: o.get("src").and_then(Value::as_str)?.to_string(),
            alt: o.get("alt").and_then(Value::as_str).map(String::from),
            w: dim_of(o.get("w")),
            h: dim_of(o.get("h")),
        },
        "table" => {
            let rows = o
                .get("rows")
                .and_then(Value::as_array)
                .map(|rows| {
                    rows.iter()
                        .filter_map(Value::as_array)
                        .map(|cells| {
                            cells
                                .iter()
                                .map(|c| match c {
                                    Value::Object(co) => TableCell {
                                        blocks: co
                                            .get("blocks")
                                            .and_then(Value::as_array)
                                            .map(|bs| bs.iter().filter_map(|b| serde_json::from_value::<Block>(b.clone()).ok()).collect())
                                            .unwrap_or_default(),
                                        colspan: dim_of(co.get("colspan")).filter(|c| *c > 1),
                                    },
                                    other => {
                                        let mut t = String::new();
                                        if let Value::String(s) = other { t = s.clone() } else { collect_strings(other, &mut t) }
                                        TableCell { blocks: vec![Block::P { runs: vec![Run::text(t)], align: None }], colspan: None }
                                    }
                                })
                                .collect()
                        })
                        .collect()
                })
                .unwrap_or_default();
            Block::Table { rows }
        }
        "pagebreak" | "page-break" | "pb" => Block::PageBreak,
        "hr" => Block::P { runs: Vec::new(), align: None },
        _ => {
            if !runs.is_empty() {
                Block::P { runs, align }
            } else {
                let mut t = String::new();
                collect_strings(v, &mut t);
                if t.trim().is_empty() {
                    return None;
                }
                Block::P { runs: vec![Run::text(t)], align: None }
            }
        }
    })
}

/// Read one ZIP entry with a hard cap on the inflated size.
pub fn read_zip_entry<R: Read + std::io::Seek>(zip: &mut ZipArchive<R>, name: &str, cap: u64) -> Result<Vec<u8>, String> {
    let f = zip.by_name(name).map_err(|e| format!("{name}: {e}"))?;
    if f.size() > cap {
        return Err(format!("{name} is too large to open ({} MB uncompressed)", f.size() / (1024 * 1024)));
    }
    let mut v = Vec::with_capacity(f.size().min(cap) as usize);
    f.take(cap + 1).read_to_end(&mut v).map_err(|e| format!("{name}: {e}"))?;
    if v.len() as u64 > cap {
        return Err(format!("{name} is too large to open"));
    }
    Ok(v)
}

/// XML bytes -> String (BOM-aware; UTF-16 parts are rare but legal).
pub fn xml_bytes_to_string(b: &[u8]) -> String {
    if let Some(rest) = b.strip_prefix(&[0xEF, 0xBB, 0xBF]) {
        return String::from_utf8_lossy(rest).into_owned();
    }
    if b.starts_with(&[0xFF, 0xFE]) || b.starts_with(&[0xFE, 0xFF]) {
        let le = b[0] == 0xFF;
        let units: Vec<u16> = b[2..]
            .chunks_exact(2)
            .map(|c| if le { u16::from_le_bytes([c[0], c[1]]) } else { u16::from_be_bytes([c[0], c[1]]) })
            .collect();
        return String::from_utf16_lossy(&units);
    }
    String::from_utf8_lossy(b).into_owned()
}

// ============================ mini DOM ============================

/// A tiny owned XML DOM on top of quick-xml. Element/attribute names keep
/// their prefixes; lookups use local names. Depth is capped (deeper subtrees
/// are dropped) so recursive walkers cannot overflow the stack.
pub mod xml {
    use quick_xml::events::{BytesStart, Event};

    pub const MAX_DEPTH: usize = 160;

    #[derive(Debug, Clone, Default)]
    pub struct Node {
        pub name: String,
        pub attrs: Vec<(String, String)>,
        pub children: Vec<Child>,
    }

    #[derive(Debug, Clone)]
    pub enum Child {
        El(Node),
        Text(String),
    }

    pub fn local_name(q: &str) -> &str {
        q.rsplit(':').next().unwrap_or(q)
    }

    impl Node {
        pub fn local(&self) -> &str {
            local_name(&self.name)
        }
        pub fn is(&self, l: &str) -> bool {
            self.local() == l
        }
        /// Attribute by local name (prefix ignored).
        pub fn attr(&self, l: &str) -> Option<&str> {
            self.attrs.iter().find(|(k, _)| local_name(k) == l).map(|(_, v)| v.as_str())
        }
        /// Attribute by exact qualified name.
        pub fn attr_q(&self, q: &str) -> Option<&str> {
            self.attrs.iter().find(|(k, _)| k == q).map(|(_, v)| v.as_str())
        }
        pub fn elems(&self) -> impl Iterator<Item = &Node> {
            self.children.iter().filter_map(|c| match c {
                Child::El(n) => Some(n),
                _ => None,
            })
        }
        pub fn child(&self, l: &str) -> Option<&Node> {
            self.elems().find(|n| n.is(l))
        }
        /// First descendant with this local name (depth-first).
        pub fn find(&self, l: &str) -> Option<&Node> {
            for c in self.elems() {
                if c.is(l) {
                    return Some(c);
                }
                if let Some(f) = c.find(l) {
                    return Some(f);
                }
            }
            None
        }
        /// All descendants with this local name (not descending into matches).
        pub fn find_all<'a>(&'a self, l: &str, out: &mut Vec<&'a Node>) {
            for c in self.elems() {
                if c.is(l) {
                    out.push(c);
                } else {
                    c.find_all(l, out);
                }
            }
        }
        /// Concatenated descendant text.
        pub fn text(&self) -> String {
            let mut s = String::new();
            self.collect_text(&mut s);
            s
        }
        fn collect_text(&self, s: &mut String) {
            for c in &self.children {
                match c {
                    Child::Text(t) => s.push_str(t),
                    Child::El(n) => n.collect_text(s),
                }
            }
        }
    }

    fn make_node(e: &BytesStart) -> Node {
        let name = String::from_utf8_lossy(e.name().as_ref()).into_owned();
        let mut attrs = Vec::new();
        for a in e.attributes().with_checks(false).flatten() {
            let k = String::from_utf8_lossy(a.key.as_ref()).into_owned();
            let v = match a.unescape_value() {
                Ok(v) => v.into_owned(),
                Err(_) => String::from_utf8_lossy(&a.value).into_owned(),
            };
            attrs.push((k, v));
        }
        Node { name, attrs, children: Vec::new() }
    }

    /// Parse an XML document. Whitespace-only text nodes are dropped unless
    /// `keep_ws(parent_local_name)` says they are significant.
    pub fn parse(src: &str, keep_ws: fn(&str) -> bool) -> Result<Node, String> {
        let mut rdr = quick_xml::Reader::from_str(src);
        rdr.config_mut().check_end_names = false;
        let mut stack: Vec<Node> = vec![Node { name: "#document".into(), ..Default::default() }];
        let mut skip = 0usize;
        loop {
            match rdr.read_event() {
                Ok(Event::Start(e)) => {
                    if skip > 0 || stack.len() > MAX_DEPTH {
                        skip += 1;
                        continue;
                    }
                    stack.push(make_node(&e));
                }
                Ok(Event::Empty(e)) => {
                    if skip > 0 || stack.len() > MAX_DEPTH {
                        continue;
                    }
                    let n = make_node(&e);
                    if let Some(top) = stack.last_mut() {
                        top.children.push(Child::El(n));
                    }
                }
                Ok(Event::End(_)) => {
                    if skip > 0 {
                        skip -= 1;
                        continue;
                    }
                    if stack.len() > 1 {
                        let n = stack.pop().unwrap();
                        stack.last_mut().unwrap().children.push(Child::El(n));
                    }
                }
                Ok(Event::Text(t)) => {
                    if skip > 0 || stack.len() < 2 {
                        continue;
                    }
                    let s = match t.unescape() {
                        Ok(s) => s.into_owned(),
                        Err(_) => String::from_utf8_lossy(&t).into_owned(),
                    };
                    push_text(&mut stack, s, keep_ws);
                }
                Ok(Event::CData(t)) => {
                    if skip > 0 || stack.len() < 2 {
                        continue;
                    }
                    let s = String::from_utf8_lossy(&t).into_owned();
                    push_text(&mut stack, s, keep_ws);
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(format!("XML error at byte {}: {e}", rdr.buffer_position())),
                _ => {}
            }
        }
        while stack.len() > 1 {
            let n = stack.pop().unwrap();
            stack.last_mut().unwrap().children.push(Child::El(n));
        }
        let doc = stack.pop().unwrap();
        doc.children
            .into_iter()
            .find_map(|c| match c {
                Child::El(n) => Some(n),
                _ => None,
            })
            .ok_or_else(|| "empty XML document".to_string())
    }

    fn push_text(stack: &mut [Node], s: String, keep_ws: fn(&str) -> bool) {
        let top = stack.last_mut().unwrap();
        if s.trim().is_empty() && !keep_ws(top.local()) {
            return;
        }
        if let Some(Child::Text(prev)) = top.children.last_mut() {
            prev.push_str(&s);
        } else {
            top.children.push(Child::Text(s));
        }
    }
}

// ============================ READ ============================

fn keep_ws_docx(parent: &str) -> bool {
    matches!(parent, "t" | "instrText" | "delText" | "delInstrText")
}

fn on_off(n: &Node) -> bool {
    !matches!(n.attr("val"), Some("0" | "false" | "off" | "none"))
}

fn hex_color(v: &str) -> Option<String> {
    if v.eq_ignore_ascii_case("auto") {
        return None;
    }
    norm_hex(v).map(|h| format!("#{h}"))
}

/// Apply one rPr child element to a run.
fn apply_rpr_child(c: &Node, run: &mut Run) {
    match c.local() {
        "b" => run.b = on_off(c),
        "i" => run.i = on_off(c),
        "u" => run.u = !matches!(c.attr("val"), Some("none")),
        "strike" | "dstrike" => run.s = on_off(c),
        "color" => run.color = c.attr("val").and_then(hex_color),
        "sz" => {
            if let Some(v) = c.attr("val").and_then(|v| v.trim().parse::<f64>().ok()).filter(|v| *v > 0.0) {
                run.size = Some(v / 2.0);
            }
        }
        "vertAlign" => match c.attr("val") {
            Some("superscript") => {
                run.sup = true;
                run.sub = false;
            }
            Some("subscript") => {
                run.sub = true;
                run.sup = false;
            }
            Some("baseline") => {
                run.sup = false;
                run.sub = false;
            }
            _ => {}
        },
        "rFonts" => {
            let f = c.attr("ascii").or_else(|| c.attr("hAnsi")).or_else(|| c.attr("eastAsia")).or_else(|| c.attr("cs"));
            if let Some(f) = f.map(str::trim).filter(|f| !f.is_empty()) {
                run.font = Some(f.to_string());
            }
        }
        "highlight" => run.hl = c.attr("val").and_then(highlight_hex),
        "shd" => {
            if run.hl.is_none() {
                if let Some(f) = c.attr("fill").and_then(hex_color) {
                    if f != "#FFFFFF" {
                        run.hl = Some(f);
                    }
                }
            }
        }
        _ => {}
    }
}

/// HYPERLINK field instruction -> link target.
pub fn field_link(instr: &str) -> Option<String> {
    let mut toks: Vec<String> = Vec::new();
    let mut cur = String::new();
    let mut in_q = false;
    for c in instr.chars() {
        match c {
            '"' => {
                if in_q {
                    toks.push(std::mem::take(&mut cur));
                }
                in_q = !in_q;
            }
            c if c.is_whitespace() && !in_q => {
                if !cur.is_empty() {
                    toks.push(std::mem::take(&mut cur));
                }
            }
            c => cur.push(c),
        }
    }
    if !cur.is_empty() {
        toks.push(cur);
    }
    let mut it = toks.into_iter();
    if !it.next()?.eq_ignore_ascii_case("HYPERLINK") {
        return None;
    }
    let (mut url, mut anchor) = (None, None);
    while let Some(t) = it.next() {
        if t.eq_ignore_ascii_case("\\l") {
            anchor = it.next();
        } else if t.eq_ignore_ascii_case("\\o") || t.eq_ignore_ascii_case("\\t") {
            it.next();
        } else if t.starts_with('\\') {
        } else if url.is_none() {
            url = Some(t);
        }
    }
    let link = match (url, anchor) {
        (Some(u), Some(a)) => format!("{u}#{a}"),
        (Some(u), None) => u,
        (None, Some(a)) => format!("#{a}"),
        _ => return None,
    };
    safe_link(&link)
}

pub fn heading_from_name(lname: &str) -> Option<u32> {
    let rest = lname.strip_prefix("heading")?.trim();
    let n: u32 = rest.parse().ok()?;
    if (1..=9).contains(&n) { Some(n.min(6)) } else { None }
}

fn norm_style_id(id: &str) -> String {
    id.chars().filter(|c| !matches!(c, ' ' | '_' | '-')).collect::<String>().to_lowercase()
}

/// Heading level from a style ID alone (used when styles.xml lacks the
/// style): English and common localized built-in heading IDs.
fn heading_from_id(id: &str) -> Option<u32> {
    let l = norm_style_id(id);
    if l == "title" {
        return Some(1);
    }
    let digits: String = l.chars().rev().take_while(|c| c.is_ascii_digit()).collect::<Vec<_>>().into_iter().rev().collect();
    if digits.is_empty() {
        return None;
    }
    let prefix = &l[..l.len() - digits.len()];
    const PREFIXES: &[&str] = &[
        "heading", "berschrift", "überschrift", "titre", "ttulo", "título", "titolo", "kop", "rubrik",
        "nagwek", "nagłówek", "otsikko", "overskrift", "nadpis", "címsor", "cmsor", "başlık", "balk",
        "заголовок", "kopjes", "otsikko", "rubrica", "cabeçalho",
    ];
    if !PREFIXES.contains(&prefix) {
        return None;
    }
    let n: u32 = digits.parse().ok()?;
    if (1..=9).contains(&n) { Some(n.min(6)) } else { None }
}

pub fn is_code_style_name(n: &str) -> bool {
    let l = norm_style_id(n);
    matches!(l.as_str(), "code" | "codeblock" | "sourcecode" | "htmlpreformatted" | "preformattedtext" | "macrotext")
}

/// Paragraph kind while building blocks (shared with textfmt readers).
#[derive(Clone, Debug, PartialEq)]
pub enum PKind {
    P,
    H(u32),
    Quote,
    Code,
    List(String, u32),
}

pub fn kind_block(k: &PKind, runs: Vec<Run>, align: Option<String>) -> Block {
    match k {
        PKind::P => Block::P { runs, align },
        PKind::H(n) => heading_block(*n, runs, align),
        PKind::Quote => Block::Quote { runs, align },
        PKind::Code => Block::Code { runs, align },
        PKind::List(l, lvl) => Block::List { list: l.clone(), level: *lvl, runs },
    }
}

pub fn heading_block(n: u32, runs: Vec<Run>, align: Option<String>) -> Block {
    match n {
        0 | 1 => Block::H1 { runs, align },
        2 => Block::H2 { runs, align },
        3 => Block::H3 { runs, align },
        4 => Block::H4 { runs, align },
        5 => Block::H5 { runs, align },
        _ => Block::H6 { runs, align },
    }
}

/// Merge consecutive code paragraphs into one code block (lines joined with
/// newlines) — Word stores one paragraph per code line.
pub fn merge_code_blocks(blocks: Vec<Block>) -> Vec<Block> {
    let mut out: Vec<Block> = Vec::with_capacity(blocks.len());
    for b in blocks {
        if let Block::Code { runs, align } = b {
            if let Some(Block::Code { runs: prev, align: pa }) = out.last_mut() {
                if *pa == align {
                    prev.push(Run::text("\n"));
                    prev.extend(runs);
                    let merged = merge_runs(std::mem::take(prev));
                    *prev = merged;
                    continue;
                }
            }
            out.push(Block::Code { runs, align });
        } else {
            out.push(b);
        }
    }
    out
}

/// Flatten blocks into one run list (used for notes).
fn flatten_runs(blocks: &[Block], out: &mut Vec<Run>) {
    for b in blocks {
        match b {
            Block::Table { rows } => {
                for row in rows {
                    for c in row {
                        flatten_runs(&c.blocks, out);
                    }
                }
            }
            _ => {
                let runs = b.para_runs();
                if runs.is_empty() {
                    continue;
                }
                if !out.is_empty() {
                    out.push(Run::text(" "));
                }
                out.extend(runs.iter().cloned());
            }
        }
    }
}

fn choose_alt(ac: &Node) -> Option<&Node> {
    ac.child("Choice").or_else(|| ac.child("Fallback"))
}

/// Resolve a relationship target against the source part's directory.
fn resolve_target(dir: &str, target: &str) -> String {
    let t = target.trim();
    let joined = if let Some(abs) = t.strip_prefix('/') { abs.to_string() } else { format!("{dir}{t}") };
    let mut parts: Vec<&str> = Vec::new();
    for seg in joined.split('/') {
        match seg {
            "" | "." => {}
            ".." => {
                parts.pop();
            }
            s => parts.push(s),
        }
    }
    parts.join("/")
}

fn split_part(path: &str) -> (String, String) {
    match path.rsplit_once('/') {
        Some((d, f)) => (format!("{d}/"), f.to_string()),
        None => (String::new(), path.to_string()),
    }
}

/// Parse a VML/CSS length ("12pt", "1.5in", "2cm", "40px", "30mm") into px.
pub fn css_len_px(v: &str) -> Option<u32> {
    let v = v.trim();
    let num_end = v.find(|c: char| !(c.is_ascii_digit() || c == '.' || c == '-')).unwrap_or(v.len());
    let n: f64 = v[..num_end].parse().ok()?;
    let px = match v[num_end..].trim().to_ascii_lowercase().as_str() {
        "pt" => n * 96.0 / 72.0,
        "in" => n * 96.0,
        "cm" => n * 96.0 / 2.54,
        "mm" => n * 96.0 / 25.4,
        "pc" => n * 16.0,
        "px" | "" => n,
        _ => return None,
    };
    if px >= 1.0 && px < 100_000.0 { Some(px.round() as u32) } else { None }
}

struct Rel {
    kind: String,
    target: String,
    external: bool,
}

struct StyleDef {
    name: String,
    based_on: Option<String>,
    outline: Option<u32>,
    num_id: Option<u32>,
    ilvl: Option<u32>,
    rpr: Option<Node>,
}

#[derive(Default)]
struct StyleInfo {
    heading: Option<u32>,
    quote: bool,
    code: bool,
    num_id: Option<u32>,
    ilvl: Option<u32>,
    /// Built-in list style recognised by ID only (styles.xml missing it).
    id_list: Option<(bool, u32)>,
}

/// Word's built-in list styles by ID ("ListBullet", "ListNumber3", ...),
/// used only when styles.xml does not define them: (ordered, level).
fn list_from_style_id(id: &str) -> Option<(bool, u32)> {
    let l = norm_style_id(id);
    let (ordered, rest) = if let Some(r) = l.strip_prefix("listbullet") {
        (false, r)
    } else if let Some(r) = l.strip_prefix("listnumber") {
        (true, r)
    } else {
        return None;
    };
    let lvl = if rest.is_empty() { 1 } else { rest.parse::<u32>().ok()? };
    if (1..=9).contains(&lvl) { Some((ordered, lvl - 1)) } else { None }
}

#[derive(Default)]
struct PInfo {
    style: Option<String>,
    align: Option<String>,
    num_id: Option<u32>,
    ilvl: Option<u32>,
    outline: Option<u32>,
    ind_left: Option<i64>,
    ind_hanging: Option<i64>,
    page_break_before: bool,
    mark_font: Option<String>,
}

struct Field {
    in_code: bool,
    instr: String,
    link: Option<String>,
    /// Page-number field in a header/footer: its result text is replaced by a token.
    suppress: bool,
}

/// "{PAGE}" / "{NUMPAGES}" for a page-number field instruction.
fn field_token(instr: &str) -> Option<&'static str> {
    let w = instr.trim_start();
    let name: String = w.chars().take_while(|c| c.is_ascii_alphabetic()).collect::<String>().to_ascii_uppercase();
    match name.as_str() {
        "PAGE" => Some("{PAGE}"),
        "NUMPAGES" | "SECTIONPAGES" => Some("{NUMPAGES}"),
        _ => None,
    }
}

struct DocxReader<'z> {
    raw: &'z [u8],
    zip: ZipArchive<Cursor<&'z [u8]>>,
    names: Vec<String>,
    rels: HashMap<String, Rel>,
    part_dir: String,
    styles: HashMap<String, StyleDef>,
    default_pstyle: Option<String>,
    numbering: HashMap<u32, HashMap<u32, String>>,
    warnings: Vec<String>,
    fields: Vec<Field>,
    media: HashMap<String, Option<String>>,
    media_total: usize,
    footnotes: HashMap<String, Vec<Block>>,
    endnotes: HashMap<String, Vec<Block>>,
    notes: Vec<(u32, Vec<Run>)>,
    note_no: u32,
    extra_sections: bool,
    page: Option<Value>,
    /// (kind "header"/"footer", type, relationship id) of the last body sectPr.
    hf_refs: Vec<(String, String, String)>,
    /// Parsing a header/footer part: page-number fields become tokens.
    hf_mode: bool,
}

/// Does this byte string look like an RTF document?
pub fn looks_rtf(b: &[u8]) -> bool {
    let b = b.strip_prefix(&[0xEF, 0xBB, 0xBF]).unwrap_or(b);
    let start = b.iter().position(|c| !c.is_ascii_whitespace()).unwrap_or(b.len());
    b[start..].starts_with(b"{\\rtf")
}

/// Parse .docx bytes into the word model. Misnamed inputs are routed to the
/// right parser (legacy .doc, RTF, ODT); encrypted files get a clear error.
pub fn parse_docx(bytes: &[u8]) -> Result<WordDoc, String> {
    if bytes.starts_with(&[0xD0, 0xCF, 0x11, 0xE0]) {
        if crate::textfmt::ole_is_encrypted(bytes) {
            return Err("This document is password-protected (encrypted). Remove the password in Word and try again.".into());
        }
        return Ok(crate::textfmt::parse_doc(bytes));
    }
    if looks_rtf(bytes) {
        return crate::textfmt::parse_rtf(bytes);
    }
    if !bytes.starts_with(b"PK") {
        return Err("This file is not a valid .docx document (it is not a ZIP package).".into());
    }
    let zip = ZipArchive::new(Cursor::new(bytes)).map_err(|e| format!("This .docx file is damaged (zip error: {e})."))?;
    let names = zip.file_names().filter_map(|n| n.ok().map(|n| n.to_string())).collect();
    DocxReader {
        raw: bytes,
        zip,
        names,
        rels: HashMap::new(),
        part_dir: "word/".into(),
        styles: HashMap::new(),
        default_pstyle: None,
        numbering: HashMap::new(),
        warnings: Vec::new(),
        fields: Vec::new(),
        media: HashMap::new(),
        media_total: 0,
        footnotes: HashMap::new(),
        endnotes: HashMap::new(),
        notes: Vec::new(),
        note_no: 0,
        extra_sections: false,
        page: None,
        hf_refs: Vec::new(),
        hf_mode: false,
    }
    .read()
}

impl<'z> DocxReader<'z> {
    fn warn(&mut self, msg: &str) {
        if !self.warnings.iter().any(|w| w == msg) {
            self.warnings.push(msg.to_string());
        }
    }

    fn entry_name(&self, name: &str) -> Option<String> {
        let n = name.trim_start_matches('/');
        if self.names.iter().any(|x| x == n) {
            return Some(n.to_string());
        }
        let dec = String::from_utf8_lossy(&percent_decode(n)).into_owned();
        self.names.iter().find(|x| x.eq_ignore_ascii_case(n) || x.eq_ignore_ascii_case(&dec)).cloned()
    }

    fn read_bytes(&mut self, name: &str, cap: u64) -> Result<Option<Vec<u8>>, String> {
        let Some(real) = self.entry_name(name) else { return Ok(None) };
        read_zip_entry(&mut self.zip, &real, cap).map(Some)
    }

    /// Read and parse an XML part. A missing part is Ok(None); parse errors
    /// in optional parts are reported as warnings by the callers.
    fn read_xml(&mut self, name: &str) -> Result<Option<Node>, String> {
        let Some(b) = self.read_bytes(name, MAX_XML_PART)? else { return Ok(None) };
        xml::parse(&xml_bytes_to_string(&b), keep_ws_docx).map(Some)
    }

    fn read_optional_xml(&mut self, name: &str) -> Option<Node> {
        match self.read_xml(name) {
            Ok(n) => n,
            Err(e) => {
                self.warn(&format!("Part {name} could not be read ({e})."));
                None
            }
        }
    }

    fn load_rels(&mut self, dir: &str, file: &str) -> HashMap<String, Rel> {
        let mut map = HashMap::new();
        let path = format!("{dir}_rels/{file}.rels");
        if let Some(n) = self.read_optional_xml(&path) {
            for r in n.elems().filter(|e| e.is("Relationship")) {
                let id = r.attr("Id").unwrap_or("").to_string();
                map.insert(
                    id,
                    Rel {
                        kind: r.attr("Type").unwrap_or("").to_string(),
                        target: r.attr("Target").unwrap_or("").to_string(),
                        external: r.attr("TargetMode").map(|m| m.eq_ignore_ascii_case("External")).unwrap_or(false),
                    },
                );
            }
        }
        map
    }

    fn read(mut self) -> Result<WordDoc, String> {
        // Main document part from the package relationships.
        let root = self.load_rels("", "");
        let mut main = root
            .values()
            .find(|r| r.kind.ends_with("/officeDocument"))
            .map(|r| resolve_target("", &r.target))
            .unwrap_or_else(|| "word/document.xml".into());
        if self.entry_name(&main).is_none() {
            if self.entry_name("word/document.xml").is_some() {
                main = "word/document.xml".into();
            } else if self.entry_name("content.xml").is_some() {
                return crate::textfmt::parse_odt(self.raw);
            } else if self.entry_name("xl/workbook.xml").is_some() {
                return Err("This is an Excel workbook (.xlsx), not a Word document.".into());
            } else {
                return Err("This file is not a Word document (word/document.xml is missing).".into());
            }
        }
        let core_path = root
            .values()
            .find(|r| r.kind.ends_with("/core-properties"))
            .map(|r| resolve_target("", &r.target))
            .unwrap_or_else(|| "docProps/core.xml".into());

        let (dir, file) = split_part(&main);
        let doc_rels = self.load_rels(&dir, &file);
        let rel_path = |kind: &str, default: &str| -> String {
            doc_rels
                .values()
                .find(|r| r.kind.ends_with(kind))
                .map(|r| resolve_target(&dir, &r.target))
                .unwrap_or_else(|| format!("{dir}{default}"))
        };
        let styles_path = rel_path("/styles", "styles.xml");
        let numbering_path = rel_path("/numbering", "numbering.xml");
        let foot_path = doc_rels.values().find(|r| r.kind.ends_with("/footnotes")).map(|r| resolve_target(&dir, &r.target));
        let end_path = doc_rels.values().find(|r| r.kind.ends_with("/endnotes")).map(|r| resolve_target(&dir, &r.target));
        let has_comments = doc_rels.values().any(|r| r.kind.ends_with("/comments"));
        if has_comments {
            self.warn(W_COMMENTS);
        }

        if let Some(n) = self.read_optional_xml(&styles_path) {
            self.load_styles(&n);
        }
        if let Some(n) = self.read_optional_xml(&numbering_path) {
            self.load_numbering(&n);
        }
        for (path, foot) in [(foot_path, true), (end_path, false)] {
            let Some(path) = path else { continue };
            let (ndir, nfile) = split_part(&path);
            self.rels = self.load_rels(&ndir, &nfile);
            self.part_dir = ndir;
            if let Some(n) = self.read_optional_xml(&path) {
                self.load_notes(&n, foot);
            }
            self.fields.clear();
        }

        self.rels = doc_rels;
        self.part_dir = dir;
        let doc = self.read_xml(&main)?.ok_or("The main document part is missing.")?;
        let mut blocks = Vec::new();
        if let Some(body) = doc.child("body") {
            self.walk_blocks(body, &mut blocks, true);
        }
        let mut blocks = merge_code_blocks(blocks);
        if !self.notes.is_empty() {
            self.warn(W_NOTES);
            blocks.push(Block::P { runs: vec![Run::text("Notes")], align: None });
            for (n, runs) in std::mem::take(&mut self.notes) {
                let mut rs = vec![Run { t: n.to_string(), sup: true, ..Default::default() }, Run::text(" ")];
                let mut runs = runs;
                if let Some(first) = runs.first_mut() {
                    first.t = first.t.trim_start().to_string();
                }
                rs.extend(runs);
                blocks.push(Block::P { runs: merge_runs(rs), align: None });
            }
        }
        if blocks.is_empty() {
            blocks.push(Block::P { runs: Vec::new(), align: None });
        }
        if self.extra_sections {
            self.warn(W_SECTIONS);
        }

        let mut meta = Map::new();
        if let Some(core) = self.read_optional_xml(&core_path) {
            for (key, local) in [("title", "title"), ("author", "creator"), ("subject", "subject")] {
                if let Some(t) = core.find(local).map(|t| t.text().trim().to_string()).filter(|t| !t.is_empty()) {
                    meta.insert(key.into(), Value::String(t));
                }
            }
        }
        for kind in ["header", "footer"] {
            if let Some(hf) = self.read_header_footer(kind) {
                meta.insert(kind.into(), hf);
            }
        }
        if let Some(p) = self.page.take() {
            meta.insert("page".into(), p);
        }
        if !self.warnings.is_empty() {
            meta.insert("warnings".into(), json!(self.warnings));
        }
        Ok(WordDoc { kind: "word".into(), meta, blocks })
    }

    fn load_styles(&mut self, root: &Node) {
        for s in root.elems().filter(|e| e.is("style")) {
            let Some(id) = s.attr("styleId") else { continue };
            let kind = s.attr("type").unwrap_or("paragraph");
            let ppr = s.child("pPr");
            let num = ppr.and_then(|p| p.child("numPr"));
            let def = StyleDef {
                name: s.child("name").and_then(|n| n.attr("val")).unwrap_or("").to_string(),
                based_on: s.child("basedOn").and_then(|n| n.attr("val")).map(String::from),
                outline: ppr.and_then(|p| p.child("outlineLvl")).and_then(|o| o.attr("val")).and_then(|v| v.parse().ok()),
                num_id: num.and_then(|n| n.child("numId")).and_then(|n| n.attr("val")).and_then(|v| v.parse().ok()),
                ilvl: num.and_then(|n| n.child("ilvl")).and_then(|n| n.attr("val")).and_then(|v| v.parse().ok()),
                rpr: s.child("rPr").cloned(),
            };
            if kind == "paragraph" && matches!(s.attr("default"), Some("1" | "true" | "on")) {
                self.default_pstyle = Some(id.to_string());
            }
            self.styles.insert(id.to_string(), def);
        }
    }

    fn load_numbering(&mut self, root: &Node) {
        let mut abs: HashMap<String, HashMap<u32, String>> = HashMap::new();
        let lvl_fmt = |l: &Node| l.child("numFmt").and_then(|f| f.attr("val")).unwrap_or("decimal").to_string();
        for a in root.elems().filter(|e| e.is("abstractNum")) {
            let id = a.attr("abstractNumId").unwrap_or("").to_string();
            let mut levels = HashMap::new();
            for l in a.elems().filter(|e| e.is("lvl")) {
                if let Some(ilvl) = l.attr("ilvl").and_then(|v| v.parse::<u32>().ok()) {
                    levels.insert(ilvl, lvl_fmt(l));
                }
            }
            abs.insert(id, levels);
        }
        for n in root.elems().filter(|e| e.is("num")) {
            let Some(num_id) = n.attr("numId").and_then(|v| v.parse::<u32>().ok()) else { continue };
            let abs_id = n.child("abstractNumId").and_then(|a| a.attr("val")).unwrap_or("");
            let mut levels = abs.get(abs_id).cloned().unwrap_or_default();
            for o in n.elems().filter(|e| e.is("lvlOverride")) {
                if let (Some(ilvl), Some(l)) = (o.attr("ilvl").and_then(|v| v.parse::<u32>().ok()), o.child("lvl")) {
                    if l.child("numFmt").is_some() {
                        levels.insert(ilvl, lvl_fmt(l));
                    }
                }
            }
            self.numbering.insert(num_id, levels);
        }
    }

    fn load_notes(&mut self, root: &Node, foot: bool) {
        for n in root.elems().filter(|e| e.is("footnote") || e.is("endnote")) {
            if n.attr("type").map(|t| t != "normal").unwrap_or(false) {
                continue; // separators / continuation notices
            }
            let Some(id) = n.attr("id").map(String::from) else { continue };
            let blocks = self.blocks_of(n);
            if foot {
                self.footnotes.insert(id, blocks);
            } else {
                self.endnotes.insert(id, blocks);
            }
        }
    }

    fn blocks_of(&mut self, n: &Node) -> Vec<Block> {
        let mut v = Vec::new();
        self.walk_blocks(n, &mut v, false);
        merge_code_blocks(v)
    }

    fn walk_blocks(&mut self, n: &Node, out: &mut Vec<Block>, body: bool) {
        for c in n.elems() {
            match c.local() {
                "p" => self.paragraph(c, out),
                "tbl" => out.push(self.table(c)),
                "sdt" => {
                    if let Some(sc) = c.child("sdtContent") {
                        self.walk_blocks(sc, out, false);
                    }
                }
                "customXml" | "sdtContent" | "smartTag" => self.walk_blocks(c, out, false),
                "ins" | "moveTo" => {
                    self.warn(W_TRACKED);
                    self.walk_blocks(c, out, false);
                }
                "del" | "moveFrom" => self.warn(W_TRACKED),
                "AlternateContent" => {
                    if let Some(ch) = choose_alt(c) {
                        self.walk_blocks(ch, out, false);
                    }
                }
                "altChunk" => self.warn(W_ALTCHUNK),
                "sectPr" if body => self.read_sect(c),
                _ => {}
            }
        }
    }

    fn read_sect(&mut self, s: &Node) {
        self.hf_refs.clear();
        for c in s.elems() {
            let kind = match c.local() {
                "headerReference" => "header",
                "footerReference" => "footer",
                _ => continue,
            };
            let rid = c.attr_q("r:id").or_else(|| c.attrs.iter().find(|(k, _)| k.ends_with(":id")).map(|(_, v)| v.as_str()));
            if let Some(rid) = rid {
                self.hf_refs.push((kind.into(), c.attr("type").unwrap_or("default").to_string(), rid.to_string()));
            }
        }
        let mut page = Map::new();
        if let Some(sz) = s.child("pgSz") {
            let w = sz.attr("w").and_then(|v| v.parse::<u32>().ok());
            let h = sz.attr("h").and_then(|v| v.parse::<u32>().ok());
            if let (Some(w), Some(h)) = (w, h) {
                page.insert("width".into(), json!(w));
                page.insert("height".into(), json!(h));
                let landscape = sz.attr("orient") == Some("landscape") || w > h;
                page.insert("orientation".into(), json!(if landscape { "landscape" } else { "portrait" }));
                page.insert("size".into(), json!(page_size_name(w, h)));
            }
        }
        if let Some(m) = s.child("pgMar") {
            let mut mm = Map::new();
            for k in ["top", "right", "bottom", "left", "header", "footer"] {
                if let Some(v) = m.attr(k).and_then(|v| v.parse::<i64>().ok()) {
                    mm.insert(k.into(), json!(v.unsigned_abs()));
                }
            }
            if !mm.is_empty() {
                page.insert("margins".into(), Value::Object(mm));
            }
        }
        if !page.is_empty() {
            self.page = Some(Value::Object(page));
        }
    }

    fn read_ppr(&mut self, ppr: Option<&Node>) -> PInfo {
        let mut pi = PInfo::default();
        let Some(ppr) = ppr else { return pi };
        for c in ppr.elems() {
            match c.local() {
                "pStyle" => pi.style = c.attr("val").map(String::from),
                "jc" => {
                    pi.align = match c.attr("val").unwrap_or("") {
                        "center" => Some("center".into()),
                        "right" | "end" => Some("right".into()),
                        "both" | "distribute" | "lowKashida" | "mediumKashida" | "highKashida" | "thaiDistribute" => Some("justify".into()),
                        _ => None,
                    }
                }
                "numPr" => {
                    pi.num_id = c.child("numId").and_then(|n| n.attr("val")).and_then(|v| v.parse().ok());
                    pi.ilvl = c.child("ilvl").and_then(|n| n.attr("val")).and_then(|v| v.parse().ok());
                }
                "outlineLvl" => pi.outline = c.attr("val").and_then(|v| v.parse().ok()),
                "ind" => {
                    pi.ind_left = c.attr("left").or_else(|| c.attr("start")).and_then(|v| v.parse().ok());
                    pi.ind_hanging = c.attr("hanging").and_then(|v| v.parse().ok());
                }
                "pageBreakBefore" => pi.page_break_before = on_off(c),
                "rPr" => pi.mark_font = c.child("rFonts").and_then(|f| f.attr("ascii").or_else(|| f.attr("hAnsi"))).map(String::from),
                "sectPr" => {
                    self.extra_sections = true;
                    if c.child("headerReference").is_some() || c.child("footerReference").is_some() {
                        self.warn(W_HF_SECTIONS);
                    }
                }
                _ => {}
            }
        }
        pi
    }

    fn style_info(&self, id: &str) -> StyleInfo {
        let mut info = StyleInfo::default();
        let mut heading_decided = false;
        let mut cur = Some(id.to_string());
        let mut guard = 0;
        while let Some(sid) = cur.take() {
            guard += 1;
            if guard > 24 {
                break;
            }
            let Some(sd) = self.styles.get(&sid) else {
                // Style not defined in styles.xml: fall back to the ID itself.
                if !heading_decided {
                    info.heading = heading_from_id(&sid);
                }
                let l = norm_style_id(&sid);
                if l.contains("quote") || l == "quotations" {
                    info.quote = true;
                }
                if is_code_style_name(&sid) {
                    info.code = true;
                }
                info.id_list = list_from_style_id(&sid);
                break;
            };
            let lname = sd.name.trim().to_lowercase();
            if !heading_decided {
                if lname == "subtitle" {
                    heading_decided = true;
                } else if lname == "title" {
                    info.heading = Some(1);
                    heading_decided = true;
                } else if let Some(n) = heading_from_name(&lname) {
                    info.heading = Some(n);
                    heading_decided = true;
                } else if let Some(o) = sd.outline {
                    if o < 6 {
                        info.heading = Some(o + 1);
                    }
                    heading_decided = true;
                } else if lname.is_empty() {
                    if let Some(h) = heading_from_id(&sid) {
                        info.heading = Some(h);
                        heading_decided = true;
                    }
                }
            }
            if lname.contains("quote") || lname == "quotations" {
                info.quote = true;
            }
            if is_code_style_name(&lname) || (lname.is_empty() && is_code_style_name(&sid)) {
                info.code = true;
            }
            if info.num_id.is_none() {
                info.num_id = sd.num_id;
            }
            if info.ilvl.is_none() {
                info.ilvl = sd.ilvl;
            }
            cur = sd.based_on.clone();
        }
        info
    }

    fn list_kind(&self, num_id: u32, ilvl: u32) -> Option<String> {
        let levels = self.numbering.get(&num_id)?;
        let fmt = levels.get(&ilvl).or_else(|| levels.get(&0)).map(String::as_str).unwrap_or("bullet");
        match fmt {
            "none" => None,
            "bullet" => Some("ul".into()),
            _ => Some("ol".into()),
        }
    }

    fn para_kind(&self, pi: &PInfo, items: &mut Vec<Inline>) -> PKind {
        let sid = pi.style.clone().or_else(|| self.default_pstyle.clone());
        let si = sid.as_deref().map(|s| self.style_info(s)).unwrap_or_default();
        if let Some(o) = pi.outline {
            if o < 6 {
                return PKind::H(o + 1);
            }
        }
        if let Some(h) = si.heading {
            return PKind::H(h);
        }
        // Direct numPr overrides the style's; numId 0 explicitly removes numbering.
        let num_id = pi.num_id.or(si.num_id);
        let ilvl = pi.ilvl.or(si.ilvl).unwrap_or(0).min(8);
        if let Some(id) = num_id.filter(|id| *id != 0) {
            if let Some(kind) = self.list_kind(id, ilvl) {
                return PKind::List(kind, ilvl);
            }
        }
        if let (None, Some((ordered, lvl))) = (num_id, si.id_list) {
            return PKind::List(if ordered { "ol".into() } else { "ul".into() }, pi.ilvl.unwrap_or(lvl).min(8));
        }
        if si.quote {
            return PKind::Quote;
        }
        if si.code {
            return PKind::Code;
        }
        if num_id.is_none() {
            if let Some(k) = legacy_marker(pi, items) {
                return k;
            }
        }
        // Code blocks written by older versions of this backend.
        if sid.as_deref().map(|s| s.eq_ignore_ascii_case("NoSpacing")).unwrap_or(false)
            && pi.mark_font.as_deref().map(is_mono_font).unwrap_or(false)
        {
            return PKind::Code;
        }
        PKind::P
    }

    fn paragraph(&mut self, p: &Node, out: &mut Vec<Block>) {
        let pi = self.read_ppr(p.child("pPr"));
        let mut items = Vec::new();
        let mut deferred = Vec::new();
        self.walk_inline(p, None, &mut items, &mut deferred);
        let kind = self.para_kind(&pi, &mut items);
        if pi.page_break_before {
            out.push(Block::PageBreak);
        }
        let cont = match &kind {
            PKind::List(..) => PKind::P,
            k => k.clone(),
        };
        let align = pi.align.clone();
        split_paragraph(items, out, &mut |runs, first| kind_block(if first { &kind } else { &cont }, runs, align.clone()));
        out.extend(deferred);
    }

    fn visible(&self) -> bool {
        self.fields.iter().all(|f| !f.in_code && !f.suppress)
    }

    /// A page-number token run (written even while the field result is hidden).
    fn push_token(&self, items: &mut Vec<Inline>, props: &Run, tok: &str) {
        let mut r = props.clone();
        r.t = tok.to_string();
        r.link = None;
        items.push(Inline::Run(r));
    }

    /// The default header or footer of the last section as `{runs, align}`.
    fn read_header_footer(&mut self, kind: &str) -> Option<Value> {
        let refs: Vec<(String, String)> =
            self.hf_refs.iter().filter(|(k, _, _)| k == kind).map(|(_, t, r)| (t.clone(), r.clone())).collect();
        if refs.iter().any(|(t, _)| t != "default") {
            self.warn(W_HF_VARIANTS);
        }
        let rid = refs.into_iter().find(|(t, _)| t == "default").map(|(_, r)| r)?;
        let rel = self.rels.get(&rid).filter(|r| !r.external)?;
        let path = resolve_target(&self.part_dir, &rel.target);
        let (hdir, hfile) = split_part(&path);
        let saved_rels = std::mem::replace(&mut self.rels, HashMap::new());
        let saved_dir = std::mem::replace(&mut self.part_dir, hdir.clone());
        self.rels = self.load_rels(&hdir, &hfile);
        self.hf_mode = true;
        self.fields.clear();
        let mut blocks = Vec::new();
        if let Some(root) = self.read_optional_xml(&path) {
            self.walk_blocks(&root, &mut blocks, false);
        }
        self.hf_mode = false;
        self.fields.clear();
        self.rels = saved_rels;
        self.part_dir = saved_dir;
        let (runs, align, simplified) = header_footer_runs(&blocks);
        if simplified {
            self.warn(W_HF_SIMPLIFIED);
        }
        if !runs.iter().any(|r| !r.t.trim().is_empty()) {
            return None;
        }
        let mut o = Map::new();
        o.insert("runs".into(), serde_json::to_value(&runs).unwrap_or(Value::Array(Vec::new())));
        if let Some(a) = align.filter(|a| a != "left") {
            o.insert("align".into(), Value::String(a));
        }
        Some(Value::Object(o))
    }

    fn push_text(&self, items: &mut Vec<Inline>, props: &Run, link: Option<&str>, t: &str) {
        if t.is_empty() || !self.visible() {
            return;
        }
        let mut r = props.clone();
        r.t = t.to_string();
        r.link = self.fields.iter().rev().find_map(|f| f.link.clone()).or_else(|| link.map(String::from));
        items.push(Inline::Run(r));
    }

    fn note_field(&mut self, instr: &str) {
        if !instr.trim().is_empty() {
            self.warn(W_FIELDS);
        }
    }

    fn hyperlink_target(&self, c: &Node) -> Option<String> {
        let target = c
            .attr_q("r:id")
            .or_else(|| c.attrs.iter().find(|(k, _)| k.ends_with(":id") && k != "w:id").map(|(_, v)| v.as_str()))
            .and_then(|id| self.rels.get(id))
            .map(|r| r.target.clone());
        let anchor = c.attr("anchor");
        let link = match (target, anchor) {
            (Some(t), Some(a)) => format!("{t}#{a}"),
            (Some(t), None) => t,
            (None, Some(a)) => format!("#{a}"),
            _ => return None,
        };
        safe_link(&link)
    }

    fn walk_inline(&mut self, n: &Node, link: Option<&str>, items: &mut Vec<Inline>, deferred: &mut Vec<Block>) {
        for c in n.elems() {
            match c.local() {
                "r" => self.run(c, link, items, deferred),
                "hyperlink" => {
                    let l = self.hyperlink_target(c);
                    self.walk_inline(c, l.as_deref().or(link), items, deferred);
                }
                "fldSimple" => {
                    let instr = c.attr("instr").unwrap_or("").to_string();
                    if self.hf_mode {
                        if let Some(tok) = field_token(&instr) {
                            let mut props = Run::default();
                            if let Some(rpr) = c.find("rPr") {
                                for rc in rpr.elems() {
                                    apply_rpr_child(rc, &mut props);
                                }
                            }
                            if self.visible() {
                                self.push_token(items, &props, tok);
                            }
                            continue;
                        }
                    }
                    let fl = field_link(&instr);
                    if fl.is_none() {
                        self.note_field(&instr);
                    }
                    self.walk_inline(c, fl.as_deref().or(link), items, deferred);
                }
                "sdt" => {
                    if let Some(sc) = c.child("sdtContent") {
                        self.walk_inline(sc, link, items, deferred);
                    }
                }
                "sdtContent" | "smartTag" | "customXml" | "dir" | "bdo" => self.walk_inline(c, link, items, deferred),
                "ins" | "moveTo" => {
                    self.warn(W_TRACKED);
                    self.walk_inline(c, link, items, deferred);
                }
                "del" | "moveFrom" => self.warn(W_TRACKED),
                "AlternateContent" => {
                    if let Some(b) = choose_alt(c) {
                        self.walk_inline(b, link, items, deferred);
                    }
                }
                "oMath" | "oMathPara" => {
                    self.warn(W_EQUATIONS);
                    let t = c.text();
                    self.push_text(items, &Run::default(), link, &t);
                }
                _ => {}
            }
        }
    }

    fn apply_char_style(&self, sid: &str, run: &mut Run, depth: u32) {
        if depth > 16 || sid.eq_ignore_ascii_case("Hyperlink") || sid.eq_ignore_ascii_case("FollowedHyperlink") {
            return;
        }
        let Some(sd) = self.styles.get(sid) else {
            // Built-in character styles referenced without a definition.
            match norm_style_id(sid).as_str() {
                "strong" => run.b = true,
                "emphasis" => run.i = true,
                _ => {}
            }
            return;
        };
        let ln = sd.name.to_ascii_lowercase();
        if ln == "hyperlink" || ln == "followedhyperlink" {
            return;
        }
        if let Some(b) = &sd.based_on {
            self.apply_char_style(b, run, depth + 1);
        }
        if let Some(rpr) = &sd.rpr {
            for c in rpr.elems() {
                apply_rpr_child(c, run);
            }
        }
    }

    fn run(&mut self, r: &Node, link: Option<&str>, items: &mut Vec<Inline>, deferred: &mut Vec<Block>) {
        let mut props = Run::default();
        if let Some(rpr) = r.child("rPr") {
            if let Some(sid) = rpr.child("rStyle").and_then(|s| s.attr("val")) {
                self.apply_char_style(sid, &mut props, 0);
            }
            for c in rpr.elems() {
                apply_rpr_child(c, &mut props);
            }
            if rpr.child("vanish").map(on_off).unwrap_or(false) {
                self.warn(W_HIDDEN);
            }
        }
        self.run_children(r, &props, link, items, deferred);
    }

    fn run_children(&mut self, r: &Node, props: &Run, link: Option<&str>, items: &mut Vec<Inline>, deferred: &mut Vec<Block>) {
        for c in r.elems() {
            match c.local() {
                "t" => {
                    let t = c.text();
                    self.push_text(items, props, link, &t);
                }
                "tab" | "ptab" => self.push_text(items, props, link, "\t"),
                "br" => {
                    if c.attr("type") == Some("page") {
                        if self.visible() {
                            items.push(Inline::Block(Block::PageBreak));
                        }
                    } else {
                        self.push_text(items, props, link, "\n");
                    }
                }
                "cr" => self.push_text(items, props, link, "\n"),
                "noBreakHyphen" => self.push_text(items, props, link, "\u{2011}"),
                "sym" => {
                    let ch = c
                        .attr("char")
                        .and_then(|h| u32::from_str_radix(h.trim(), 16).ok())
                        .map(|v| if (0xF000..=0xF0FF).contains(&v) { v - 0xF000 } else { v })
                        .and_then(char::from_u32);
                    if let Some(ch) = ch {
                        let ch = symbol_char(c.attr("font").unwrap_or(""), ch);
                        self.push_text(items, props, link, &ch.to_string());
                    }
                }
                "instrText" => {
                    let t = c.text();
                    if let Some(f) = self.fields.last_mut() {
                        if f.in_code {
                            f.instr.push_str(&t);
                        }
                    }
                }
                "fldChar" => match c.attr("fldCharType") {
                    Some("begin") => self.fields.push(Field { in_code: true, instr: String::new(), link: None, suppress: false }),
                    Some("separate") => {
                        let mut instr = String::new();
                        let mut token = None;
                        let hf = self.hf_mode;
                        if let Some(f) = self.fields.last_mut() {
                            f.in_code = false;
                            f.link = field_link(&f.instr);
                            token = if hf { field_token(&f.instr) } else { None };
                            if f.link.is_none() && token.is_none() {
                                instr = f.instr.clone();
                            }
                        }
                        if let Some(tok) = token {
                            // Outer fields still hidden? Then the token is too.
                            if self.visible() {
                                self.push_token(items, props, tok);
                            }
                            if let Some(f) = self.fields.last_mut() {
                                f.suppress = true;
                            }
                        }
                        self.note_field(&instr);
                    }
                    Some("end") => {
                        if let Some(f) = self.fields.pop() {
                            if f.in_code {
                                match field_token(&f.instr).filter(|_| self.hf_mode) {
                                    // Page-number field without a result part.
                                    Some(tok) if self.visible() => self.push_token(items, props, tok),
                                    Some(_) => {}
                                    // Field without a result part: nothing shown.
                                    None => self.note_field(&f.instr),
                                }
                            }
                        }
                    }
                    _ => {}
                },
                "drawing" => {
                    if self.visible() {
                        self.drawing(c, items, deferred);
                    }
                }
                "pict" | "object" => {
                    if self.visible() {
                        self.vml(c, items, deferred);
                    }
                }
                "AlternateContent" => {
                    if let Some(b) = choose_alt(c) {
                        self.run_children(b, props, link, items, deferred);
                    }
                }
                "footnoteReference" | "endnoteReference" => {
                    if self.visible() {
                        self.note_ref(c, props, items);
                    }
                }
                "commentReference" => self.warn(W_COMMENTS),
                "delText" | "delInstrText" => self.warn(W_TRACKED),
                "ruby" => {
                    if let Some(base) = c.child("rubyBase") {
                        self.walk_inline(base, link, items, deferred);
                    }
                }
                _ => {}
            }
        }
    }

    fn note_ref(&mut self, c: &Node, props: &Run, items: &mut Vec<Inline>) {
        let id = c.attr("id").unwrap_or("");
        let blocks = if c.is("footnoteReference") { self.footnotes.get(id) } else { self.endnotes.get(id) }.cloned();
        let Some(blocks) = blocks else { return };
        self.note_no += 1;
        let n = self.note_no;
        let mut r = props.clone();
        r.t = n.to_string();
        r.sup = true;
        r.sub = false;
        r.link = None;
        items.push(Inline::Run(r));
        let mut runs = Vec::new();
        flatten_runs(&blocks, &mut runs);
        self.notes.push((n, runs));
    }

    fn image_from_rel(&mut self, rid: Option<&str>, alt: Option<String>, w: Option<u32>, h: Option<u32>) -> Option<Block> {
        let rel = self.rels.get(rid?)?;
        if rel.external {
            self.warn(W_LINKED_IMG);
            return None;
        }
        let path = resolve_target(&self.part_dir, &rel.target);
        let url = if let Some(cached) = self.media.get(&path) {
            cached.clone()
        } else {
            let loaded = match self.read_bytes(&path, MAX_IMAGE) {
                Ok(Some(bytes)) if !bytes.is_empty() => {
                    let ext = path.rsplit('.').next().unwrap_or("");
                    let mime = sniff_image(&bytes).unwrap_or_else(|| mime_for_ext(ext));
                    if matches!(mime, "image/x-emf" | "image/x-wmf" | "image/tiff") {
                        self.warn(W_IMG_FORMAT);
                    }
                    let url = data_url(mime, &bytes);
                    if self.media_total + url.len() > MAX_MEDIA_TOTAL {
                        self.warn(W_IMG_LIMIT);
                        None
                    } else {
                        self.media_total += url.len();
                        Some(url)
                    }
                }
                Ok(_) => None,
                Err(_) => {
                    self.warn(W_IMG_LIMIT);
                    None
                }
            };
            self.media.insert(path.clone(), loaded.clone());
            loaded
        }?;
        let (w, h) = match (w, h) {
            (Some(w), Some(h)) => (Some(w), Some(h)),
            _ => match parse_data_url(&url).and_then(|(_, b)| image_dims(&b)) {
                Some((iw, ih)) => (Some(iw), Some(ih)),
                None => (w, h),
            },
        };
        Some(Block::Img { src: url, alt, w, h })
    }

    fn drawing(&mut self, d: &Node, items: &mut Vec<Inline>, deferred: &mut Vec<Block>) {
        for holder in d.elems() {
            let ext = holder.child("extent");
            let emu = |k: &str| ext.and_then(|e| e.attr(k)).and_then(|v| v.parse::<u64>().ok()).map(|v| ((v + 4762) / 9525) as u32).filter(|v| *v > 0);
            let (w, h) = (emu("cx"), emu("cy"));
            let alt = holder
                .child("docPr")
                .and_then(|p| p.attr("descr").filter(|s| !s.trim().is_empty()).or_else(|| p.attr("title").filter(|s| !s.trim().is_empty())))
                .map(String::from);
            let mut handled = false;
            let mut blips = Vec::new();
            holder.find_all("blip", &mut blips);
            let single = blips.len() == 1;
            for blip in blips {
                let rid = blip.attr("embed");
                if rid.is_none() && blip.attr("link").is_some() {
                    self.warn(W_LINKED_IMG);
                    continue;
                }
                let (bw, bh) = if single { (w, h) } else { (None, None) };
                if let Some(img) = self.image_from_rel(rid, alt.clone(), bw, bh) {
                    items.push(Inline::Block(img));
                    handled = true;
                }
            }
            let mut boxes = Vec::new();
            holder.find_all("txbxContent", &mut boxes);
            for b in boxes {
                let blocks = self.blocks_of(b);
                if !blocks.is_empty() {
                    deferred.extend(blocks);
                    handled = true;
                    self.warn(W_TEXTBOX);
                }
            }
            if !handled {
                if holder.find("chart").is_some() || holder.find("relIds").is_some() {
                    self.warn(W_CHARTS);
                } else if holder.local() == "inline" || holder.local() == "anchor" {
                    self.warn(W_SHAPES);
                }
            }
        }
    }

    fn vml(&mut self, n: &Node, items: &mut Vec<Inline>, deferred: &mut Vec<Block>) {
        let mut handled = false;
        let (mut w, mut h) = (None, None);
        if let Some(style) = n.find("shape").and_then(|s| s.attr("style")) {
            for decl in style.split(';') {
                if let Some((k, v)) = decl.split_once(':') {
                    match k.trim().to_ascii_lowercase().as_str() {
                        "width" => w = css_len_px(v),
                        "height" => h = css_len_px(v),
                        _ => {}
                    }
                }
            }
        }
        let mut imgs = Vec::new();
        n.find_all("imagedata", &mut imgs);
        let single = imgs.len() == 1;
        for im in imgs {
            let rid = im.attr_q("r:id").or_else(|| im.attr("relid")).or_else(|| im.attr("id"));
            let alt = im.attr("title").filter(|s| !s.trim().is_empty()).map(String::from);
            let (bw, bh) = if single { (w, h) } else { (None, None) };
            if let Some(img) = self.image_from_rel(rid, alt, bw, bh) {
                items.push(Inline::Block(img));
                handled = true;
            }
        }
        let mut boxes = Vec::new();
        n.find_all("txbxContent", &mut boxes);
        for b in boxes {
            let blocks = self.blocks_of(b);
            if !blocks.is_empty() {
                deferred.extend(blocks);
                handled = true;
                self.warn(W_TEXTBOX);
            }
        }
        if n.is("object") {
            self.warn(W_OBJECTS);
        } else if !handled && ["shape", "rect", "oval", "roundrect", "group", "polyline"].iter().any(|k| n.find(k).is_some()) {
            self.warn(W_SHAPES);
        }
    }

    fn table(&mut self, tbl: &Node) -> Block {
        let mut rows = Vec::new();
        self.collect_rows(tbl, &mut rows);
        Block::Table { rows }
    }

    fn collect_rows(&mut self, n: &Node, rows: &mut Vec<Vec<TableCell>>) {
        for c in n.elems() {
            match c.local() {
                "tr" => {
                    let mut cells = Vec::new();
                    self.collect_cells(c, &mut cells);
                    rows.push(cells);
                }
                "sdt" => {
                    if let Some(sc) = c.child("sdtContent") {
                        self.collect_rows(sc, rows);
                    }
                }
                "customXml" | "sdtContent" | "ins" | "moveTo" => self.collect_rows(c, rows),
                _ => {}
            }
        }
    }

    fn collect_cells(&mut self, n: &Node, cells: &mut Vec<TableCell>) {
        for c in n.elems() {
            match c.local() {
                "tc" => {
                    let cell = self.cell(c);
                    cells.push(cell);
                }
                "sdt" => {
                    if let Some(sc) = c.child("sdtContent") {
                        self.collect_cells(sc, cells);
                    }
                }
                "customXml" | "sdtContent" => self.collect_cells(c, cells),
                _ => {}
            }
        }
    }

    fn cell(&mut self, tc: &Node) -> TableCell {
        let mut colspan = None;
        if let Some(pr) = tc.child("tcPr") {
            if let Some(g) = pr.child("gridSpan").and_then(|g| g.attr("val")).and_then(|v| v.parse::<u32>().ok()) {
                if g > 1 {
                    colspan = Some(g.min(63));
                }
            }
            if pr.child("vMerge").is_some() {
                self.warn(W_VMERGE);
            }
        }
        TableCell { blocks: self.blocks_of(tc), colspan }
    }
}

/// Map Symbol/Wingdings private glyphs to Unicode where it matters (bullets,
/// arrows, check marks); other characters pass through.
fn symbol_char(font: &str, ch: char) -> char {
    let f = font.to_ascii_lowercase();
    let c = ch as u32;
    if f.contains("symbol") {
        return match c {
            0xB7 => '•',
            0xAE => '→',
            0xAC => '←',
            0xB4 => '×',
            0xB8 => '÷',
            0xA3 => '≤',
            0xB3 => '≥',
            0xB1 => '±',
            0xA5 => '∞',
            _ => ch,
        };
    }
    if f.contains("wingdings") {
        return match c {
            0x9F | 0xA1 | 0x6C => '●',
            0xA7 | 0x6E => '■',
            0xD8 => '➢',
            0xE0 | 0xE8 => '➔',
            0xF0 => '⇨',
            0xFC => '✓',
            0xFB => '✗',
            0xFE => '☑',
            0xA8 => '◻',
            0x71 => '❑',
            0x76 => '❖',
            _ => ch,
        };
    }
    ch
}

/// Recover lists written by older versions of this backend as literal text
/// ("    • " / "1. " in a separate first run, w:ind hanging=360).
fn legacy_marker(pi: &PInfo, items: &mut Vec<Inline>) -> Option<PKind> {
    if pi.ind_hanging? != 360 {
        return None;
    }
    let left = pi.ind_left?;
    if left < 360 || left % 360 != 0 {
        return None;
    }
    let Some(Inline::Run(first)) = items.first() else { return None };
    let t = first.t.as_str();
    let spaces = t.len() - t.trim_start_matches(' ').len();
    let rest = &t[spaces..];
    if spaces % 4 != 0 {
        return None;
    }
    let kind = if rest == "• " {
        "ul"
    } else if rest.len() > 2 && rest.ends_with(". ") && rest[..rest.len() - 2].chars().all(|c| c.is_ascii_digit()) {
        "ol"
    } else {
        return None;
    };
    let level = ((left / 360) - 1).clamp(0, 8) as u32;
    items.remove(0);
    Some(PKind::List(kind.into(), level))
}

/// Header/footer blocks -> one paragraph's runs: paragraphs joined with line
/// breaks, table cells with tabs; pictures dropped. Returns (runs, align of
/// the first aligned paragraph, whether anything was simplified).
fn header_footer_runs(blocks: &[Block]) -> (Vec<Run>, Option<String>, bool) {
    fn visit(blocks: &[Block], lines: &mut Vec<Vec<Run>>, align: &mut Option<String>, simplified: &mut bool, depth: u32) {
        if depth > 8 {
            return;
        }
        for b in blocks {
            match b {
                Block::Table { rows } => {
                    *simplified = true;
                    for row in rows {
                        let mut line: Vec<Run> = Vec::new();
                        for (i, c) in row.iter().enumerate() {
                            let mut sub = Vec::new();
                            visit(&c.blocks, &mut sub, align, simplified, depth + 1);
                            if i > 0 {
                                line.push(Run::text("\t"));
                            }
                            for (j, l) in sub.into_iter().enumerate() {
                                if j > 0 {
                                    line.push(Run::text(" "));
                                }
                                line.extend(l);
                            }
                        }
                        lines.push(line);
                    }
                }
                Block::Img { .. } => *simplified = true,
                Block::PageBreak | Block::Other(_) => {}
                _ => {
                    if align.is_none() {
                        if let Block::P { align: Some(a), .. }
                        | Block::H1 { align: Some(a), .. }
                        | Block::H2 { align: Some(a), .. }
                        | Block::H3 { align: Some(a), .. }
                        | Block::Quote { align: Some(a), .. } = b
                        {
                            *align = Some(a.clone());
                        }
                    }
                    lines.push(b.para_runs().to_vec());
                }
            }
        }
    }
    let mut lines = Vec::new();
    let mut align = None;
    let mut simplified = false;
    visit(blocks, &mut lines, &mut align, &mut simplified, 0);
    let blank = |l: &Vec<Run>| !l.iter().any(|r| !r.t.trim().is_empty());
    while lines.last().map(blank).unwrap_or(false) {
        lines.pop();
    }
    while lines.first().map(blank).unwrap_or(false) {
        lines.remove(0);
    }
    let mut runs = Vec::new();
    for (i, l) in lines.into_iter().enumerate() {
        if i > 0 {
            runs.push(Run::text("\n"));
        }
        runs.extend(l.into_iter().filter(|r| !r.t.is_empty()));
    }
    (merge_runs(runs), align, simplified)
}

fn page_size_name(w: u32, h: u32) -> &'static str {
    let (a, b) = if w <= h { (w, h) } else { (h, w) };
    let near = |x: u32, y: u32| a.abs_diff(x) <= 60 && b.abs_diff(y) <= 60;
    if near(11906, 16838) {
        "A4"
    } else if near(12240, 15840) {
        "Letter"
    } else if near(12240, 20160) {
        "Legal"
    } else if near(16838, 23811) {
        "A3"
    } else if near(8391, 11906) {
        "A5"
    } else if near(10440, 15120) {
        "Executive"
    } else {
        "custom"
    }
}

// ============================ WRITE ============================

/// Page setup in twips.
struct PageSetup {
    w: u32,
    h: u32,
    top: u32,
    right: u32,
    bottom: u32,
    left: u32,
    header: u32,
    footer: u32,
}

pub fn named_size(name: &str) -> Option<(u32, u32)> {
    Some(match name.trim().to_ascii_lowercase().as_str() {
        "a4" => (11906, 16838),
        "letter" | "us letter" => (12240, 15840),
        "legal" | "us legal" => (12240, 20160),
        "a3" => (16838, 23811),
        "a5" => (8391, 11906),
        "executive" => (10440, 15120),
        _ => return None,
    })
}

impl PageSetup {
    fn from_meta(meta: &Map<String, Value>) -> PageSetup {
        let mut ps = PageSetup { w: 12240, h: 15840, top: 1440, right: 1440, bottom: 1440, left: 1440, header: 720, footer: 720 };
        let twips = |v: Option<&Value>, lo: f64, hi: f64| num_of(v).filter(|x| *x >= lo && *x <= hi).map(|x| x.round() as u32);
        if let Some(p) = meta.get("page").and_then(Value::as_object) {
            if let Some((w, h)) = p.get("size").and_then(Value::as_str).and_then(named_size) {
                ps.w = w;
                ps.h = h;
            }
            if let (Some(w), Some(h)) = (twips(p.get("width"), 1440.0, 31680.0), twips(p.get("height"), 1440.0, 31680.0)) {
                ps.w = w;
                ps.h = h;
            }
            match p.get("orientation").and_then(Value::as_str) {
                Some("landscape") if ps.w < ps.h => std::mem::swap(&mut ps.w, &mut ps.h),
                Some("portrait") if ps.w > ps.h => std::mem::swap(&mut ps.w, &mut ps.h),
                _ => {}
            }
            if let Some(m) = p.get("margins").and_then(Value::as_object) {
                for (k, slot) in [("top", &mut ps.top), ("right", &mut ps.right), ("bottom", &mut ps.bottom), ("left", &mut ps.left), ("header", &mut ps.header), ("footer", &mut ps.footer)] {
                    if let Some(v) = twips(m.get(k), 0.0, 14400.0) {
                        *slot = v;
                    }
                }
            }
        } else if let Some((w, h)) = meta.get("pageSize").and_then(Value::as_str).and_then(named_size) {
            ps.w = w;
            ps.h = h;
        }
        // Keep at least one inch of usable width/height.
        if ps.left + ps.right + 1440 > ps.w {
            ps.left = 1440.min(ps.w / 4);
            ps.right = ps.left;
        }
        if ps.top + ps.bottom + 1440 > ps.h {
            ps.top = 1440.min(ps.h / 4);
            ps.bottom = ps.top;
        }
        ps
    }

    fn text_width(&self) -> u32 {
        (self.w - self.left - self.right).max(1440)
    }

    fn sect_xml(&self, refs: &str) -> String {
        let orient = if self.w > self.h { " w:orient=\"landscape\"" } else { "" };
        format!(
            "<w:sectPr>{refs}<w:pgSz w:w=\"{}\" w:h=\"{}\"{orient}/><w:pgMar w:top=\"{}\" w:right=\"{}\" w:bottom=\"{}\" w:left=\"{}\" w:header=\"{}\" w:footer=\"{}\" w:gutter=\"0\"/><w:cols w:space=\"720\"/><w:docGrid w:linePitch=\"360\"/></w:sectPr>",
            self.w, self.h, self.top, self.right, self.bottom, self.left, self.header, self.footer
        )
    }
}

/// Page setup from `meta` in twips: (width, height, top, right, bottom, left).
pub fn page_twips(meta: &Map<String, Value>) -> (u32, u32, u32, u32, u32, u32) {
    let p = PageSetup::from_meta(meta);
    (p.w, p.h, p.top, p.right, p.bottom, p.left)
}

/// Convert an IRI to a URI acceptable in a relationship Target.
fn uri_for_target(url: &str) -> String {
    let mut o = String::with_capacity(url.len());
    for c in url.chars() {
        if c.is_ascii() && !c.is_ascii_control() && !matches!(c, ' ' | '"' | '<' | '>' | '\\' | '^' | '`' | '{' | '|' | '}') {
            o.push(c);
        } else {
            let mut buf = [0u8; 4];
            for b in c.encode_utf8(&mut buf).bytes() {
                o.push_str(&format!("%{b:02X}"));
            }
        }
    }
    o
}

fn hash_str(s: &str) -> u64 {
    use std::hash::{Hash, Hasher};
    let mut h = std::collections::hash_map::DefaultHasher::new();
    s.hash(&mut h);
    h.finish()
}

/// Text of a run as w:t / w:br / w:tab elements.
fn run_content(t: &str) -> String {
    let mut out = String::new();
    let mut cur = String::new();
    let flush = |cur: &mut String, out: &mut String| {
        if !cur.is_empty() {
            out.push_str("<w:t xml:space=\"preserve\">");
            out.push_str(&xml_text(cur));
            out.push_str("</w:t>");
            cur.clear();
        }
    };
    let mut chars = t.chars().peekable();
    while let Some(ch) = chars.next() {
        match ch {
            '\r' => {
                if chars.peek() != Some(&'\n') {
                    flush(&mut cur, &mut out);
                    out.push_str("<w:br/>");
                }
            }
            '\n' | '\u{b}' | '\u{2028}' => {
                flush(&mut cur, &mut out);
                out.push_str("<w:br/>");
            }
            '\u{c}' => {
                flush(&mut cur, &mut out);
                out.push_str("<w:br w:type=\"page\"/>");
            }
            '\t' => {
                flush(&mut cur, &mut out);
                out.push_str("<w:tab/>");
            }
            '\u{ad}' => {
                flush(&mut cur, &mut out);
                out.push_str("<w:softHyphen/>");
            }
            c if xml_char_ok(c) => cur.push(c),
            _ => {}
        }
    }
    flush(&mut cur, &mut out);
    out
}

/// w:rPr with children in CT_RPr schema order:
/// rStyle, rFonts, b, bCs, i, iCs, strike, color, sz, szCs, highlight, u, shd, vertAlign.
fn rpr_xml(r: &Run, linked: bool) -> String {
    let mut s = String::new();
    if linked {
        s.push_str("<w:rStyle w:val=\"Hyperlink\"/>");
    }
    if let Some(f) = r.font.as_deref().map(str::trim).filter(|f| !f.is_empty()) {
        let f = xml_attr(f);
        s.push_str(&format!("<w:rFonts w:ascii=\"{f}\" w:hAnsi=\"{f}\" w:eastAsia=\"{f}\" w:cs=\"{f}\"/>"));
    }
    if r.b {
        s.push_str("<w:b/><w:bCs/>");
    }
    if r.i {
        s.push_str("<w:i/><w:iCs/>");
    }
    if r.s {
        s.push_str("<w:strike/>");
    }
    if let Some(c) = r.color.as_deref().and_then(norm_hex) {
        s.push_str(&format!("<w:color w:val=\"{c}\"/>"));
    }
    if let Some(sz) = r.size.filter(|v| v.is_finite() && *v > 0.0) {
        let hp = (sz * 2.0).round().clamp(2.0, 3276.0) as u32;
        s.push_str(&format!("<w:sz w:val=\"{hp}\"/><w:szCs w:val=\"{hp}\"/>"));
    }
    let hl = r.hl.as_deref().and_then(norm_hex);
    let named = hl.as_deref().and_then(highlight_name);
    if let Some(n) = named {
        s.push_str(&format!("<w:highlight w:val=\"{n}\"/>"));
    }
    if r.u {
        s.push_str("<w:u w:val=\"single\"/>");
    }
    if let (Some(h), None) = (&hl, named) {
        s.push_str(&format!("<w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"{h}\"/>"));
    }
    if r.sup {
        s.push_str("<w:vertAlign w:val=\"superscript\"/>");
    } else if r.sub {
        s.push_str("<w:vertAlign w:val=\"subscript\"/>");
    }
    if s.is_empty() { s } else { format!("<w:rPr>{s}</w:rPr>") }
}

struct DocxWriter {
    rels: Vec<String>,
    next_rel: u32,
    link_rels: HashMap<String, String>,
    media: Vec<(String, Vec<u8>)>,
    media_by_src: HashMap<(u64, usize), String>,
    exts: BTreeSet<&'static str>,
    nums: Vec<(u32, u32, Option<u32>)>,
    bullet_num: Option<u32>,
    ol_level: [Option<u32>; 9],
    in_list: bool,
    pic_id: u32,
}

impl DocxWriter {
    fn new() -> Self {
        DocxWriter {
            rels: Vec::new(),
            next_rel: 10,
            link_rels: HashMap::new(),
            media: Vec::new(),
            media_by_src: HashMap::new(),
            exts: BTreeSet::new(),
            nums: Vec::new(),
            bullet_num: None,
            ol_level: [None; 9],
            in_list: false,
            pic_id: 0,
        }
    }

    fn add_rel(&mut self, kind: &str, target: &str, external: bool) -> String {
        let id = format!("rId{}", self.next_rel);
        self.next_rel += 1;
        self.rels.push(format!(
            "<Relationship Id=\"{id}\" Type=\"{REL_NS}/{kind}\" Target=\"{}\"{}/>",
            xml_attr(target),
            if external { " TargetMode=\"External\"" } else { "" }
        ));
        id
    }

    fn link_rel(&mut self, url: &str) -> String {
        if let Some(id) = self.link_rels.get(url) {
            return id.clone();
        }
        let id = self.add_rel("hyperlink", &uri_for_target(url), true);
        self.link_rels.insert(url.to_string(), id.clone());
        id
    }

    fn new_num(&mut self, abs: u32, restart: Option<u32>) -> u32 {
        let id = self.nums.len() as u32 + 1;
        self.nums.push((id, abs, restart));
        id
    }

    /// numId for a list item. Bullets share one numbering instance; every
    /// ordered list (and every nested ordered sub-list) gets its own so that
    /// numbering restarts at 1.
    fn list_num(&mut self, ordered: bool, lvl: u32) -> u32 {
        let lvl = lvl.min(8) as usize;
        if !self.in_list {
            self.ol_level = [None; 9];
            self.in_list = true;
        }
        for l in self.ol_level.iter_mut().skip(lvl + 1) {
            *l = None;
        }
        if ordered {
            if let Some(n) = self.ol_level[lvl] {
                return n;
            }
            let n = self.new_num(1, Some(lvl as u32));
            self.ol_level[lvl] = Some(n);
            n
        } else {
            self.ol_level[lvl] = None;
            match self.bullet_num {
                Some(n) => n,
                None => {
                    let n = self.new_num(0, None);
                    self.bullet_num = Some(n);
                    n
                }
            }
        }
    }

    /// Write blocks; returns true when the last element written is a paragraph.
    fn blocks(&mut self, blocks: &[Block], out: &mut String, width: u32) -> bool {
        let mut last_para = false;
        let mut prev_table = false;
        for b in blocks {
            let Some(b) = resolve_block(b) else { continue };
            let b: &Block = &b;
            if !matches!(b, Block::List { .. }) {
                self.in_list = false;
            }
            match b {
                Block::Table { rows } => {
                    self.in_list = false;
                    if rows.iter().all(|r| r.is_empty()) {
                        continue; // nothing to write (Word rejects empty tables)
                    }
                    // Two adjacent tables would merge into one in Word.
                    if prev_table {
                        out.push_str("<w:p/>");
                    }
                    self.table(rows, out, width);
                    self.in_list = false;
                    prev_table = true;
                    last_para = false;
                    continue;
                }
                Block::Img { src, alt, w, h } => self.image(src, alt.as_deref(), *w, *h, width, out),
                Block::PageBreak => out.push_str("<w:p><w:r><w:br w:type=\"page\"/></w:r></w:p>"),
                Block::List { list, level, runs } => {
                    let lvl = (*level).min(8);
                    let num = self.list_num(list == "ol", lvl);
                    self.para(Some("ListParagraph"), Some((num, lvl)), None, runs, out);
                }
                Block::Other(_) => continue,
                _ => {
                    let (style, align) = match b {
                        Block::P { align, .. } => (None, align),
                        Block::H1 { align, .. } => (Some("Heading1"), align),
                        Block::H2 { align, .. } => (Some("Heading2"), align),
                        Block::H3 { align, .. } => (Some("Heading3"), align),
                        Block::H4 { align, .. } => (Some("Heading4"), align),
                        Block::H5 { align, .. } => (Some("Heading5"), align),
                        Block::H6 { align, .. } => (Some("Heading6"), align),
                        Block::Quote { align, .. } => (Some("Quote"), align),
                        Block::Code { align, .. } => (Some("Code"), align),
                        _ => (None, &None),
                    };
                    self.para(style, None, align.as_deref(), b.para_runs(), out);
                }
            }
            prev_table = false;
            last_para = true;
        }
        last_para
    }

    fn para(&mut self, style: Option<&str>, num: Option<(u32, u32)>, align: Option<&str>, runs: &[Run], out: &mut String) {
        let mut ppr = String::new();
        if let Some(s) = style {
            ppr.push_str(&format!("<w:pStyle w:val=\"{s}\"/>"));
        }
        if let Some((n, l)) = num {
            ppr.push_str(&format!("<w:numPr><w:ilvl w:val=\"{l}\"/><w:numId w:val=\"{n}\"/></w:numPr>"));
        }
        let jc = match align {
            Some("center") => Some("center"),
            Some("right") => Some("right"),
            Some("justify") => Some("both"),
            Some("left") => Some("left"),
            _ => None,
        };
        if let Some(j) = jc {
            ppr.push_str(&format!("<w:jc w:val=\"{j}\"/>"));
        }
        out.push_str("<w:p>");
        if !ppr.is_empty() {
            out.push_str("<w:pPr>");
            out.push_str(&ppr);
            out.push_str("</w:pPr>");
        }
        self.runs(runs, out);
        out.push_str("</w:p>");
    }

    fn runs(&mut self, runs: &[Run], out: &mut String) {
        let mut i = 0;
        while i < runs.len() {
            let mut j = i + 1;
            while j < runs.len() && runs[j].link == runs[i].link {
                j += 1;
            }
            let link = runs[i].link.as_deref().and_then(safe_link);
            match link {
                Some(l) => {
                    if let Some(anchor) = l.strip_prefix('#') {
                        out.push_str(&format!("<w:hyperlink w:anchor=\"{}\" w:history=\"1\">", xml_attr(anchor)));
                    } else {
                        let rid = self.link_rel(&l);
                        out.push_str(&format!("<w:hyperlink r:id=\"{rid}\" w:history=\"1\">"));
                    }
                    for r in &runs[i..j] {
                        push_run(r, true, out);
                    }
                    out.push_str("</w:hyperlink>");
                }
                None => {
                    for r in &runs[i..j] {
                        push_run(r, false, out);
                    }
                }
            }
            i = j;
        }
    }

    fn image(&mut self, src: &str, alt: Option<&str>, w: Option<u32>, h: Option<u32>, width_twips: u32, out: &mut String) {
        let placeholder = |out: &mut String| {
            let label = alt.filter(|a| !a.trim().is_empty()).unwrap_or("image");
            out.push_str(&format!("<w:p><w:r><w:t xml:space=\"preserve\">[image: {}]</w:t></w:r></w:p>", xml_text(label)));
        };
        let Some((mime, bytes)) = parse_data_url(src).filter(|(_, b)| !b.is_empty()) else {
            placeholder(out);
            return;
        };
        let mime = sniff_image(&bytes).map(str::to_string).unwrap_or(mime);
        let Some(ext) = image_ext(&mime) else {
            placeholder(out);
            return;
        };
        let key = (hash_str(src), src.len());
        let rid = match self.media_by_src.get(&key) {
            Some(r) => r.clone(),
            None => {
                let n = self.media.len() + 1;
                let target = format!("media/image{n}.{ext}");
                self.media.push((format!("word/{target}"), bytes.clone()));
                self.exts.insert(ext);
                let rid = self.add_rel("image", &target, false);
                self.media_by_src.insert(key, rid.clone());
                rid
            }
        };
        let intrinsic = image_dims(&bytes);
        let (mut pw, mut ph): (u64, u64) = match (w, h, intrinsic) {
            (Some(w), Some(h), _) => (w as u64, h as u64),
            (Some(w), None, Some((iw, ih))) => (w as u64, w as u64 * ih as u64 / iw as u64),
            (None, Some(h), Some((iw, ih))) => (h as u64 * iw as u64 / ih as u64, h as u64),
            (None, None, Some((iw, ih))) => (iw as u64, ih as u64),
            (Some(w), None, None) => (w as u64, w as u64 * 3 / 4),
            (None, Some(h), None) => (h as u64 * 4 / 3, h as u64),
            (None, None, None) => (320, 240),
        };
        pw = pw.max(1);
        ph = ph.max(1);
        let max_px = (width_twips / 15).max(32) as u64;
        if pw > max_px {
            ph = (ph * max_px / pw).max(1);
            pw = max_px;
        }
        let (cx, cy) = (pw * 9525, ph * 9525);
        self.pic_id += 1;
        let id = self.pic_id;
        let descr = xml_attr(alt.unwrap_or(""));
        out.push_str(&format!(
            "<w:p><w:r><w:drawing><wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\"><wp:extent cx=\"{cx}\" cy=\"{cy}\"/><wp:effectExtent l=\"0\" t=\"0\" r=\"0\" b=\"0\"/><wp:docPr id=\"{id}\" name=\"Picture {id}\" descr=\"{descr}\"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect=\"1\"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"><pic:pic><pic:nvPicPr><pic:cNvPr id=\"{id}\" name=\"image{id}.{ext}\" descr=\"{descr}\"/><pic:cNvPicPr><a:picLocks noChangeAspect=\"1\" noChangeArrowheads=\"1\"/></pic:cNvPicPr></pic:nvPicPr><pic:blipFill><a:blip r:embed=\"{rid}\"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr bwMode=\"auto\"><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{cx}\" cy=\"{cy}\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"
        ));
    }

    /// Returns false when nothing was written (no non-empty rows).
    fn table(&mut self, rows: &[Vec<TableCell>], out: &mut String, width: u32) -> bool {
        let rows: Vec<&Vec<TableCell>> = rows.iter().filter(|r| !r.is_empty()).collect();
        if rows.is_empty() {
            return false;
        }
        let span = |c: &TableCell| c.colspan.unwrap_or(1).clamp(1, 63);
        let cols = rows.iter().map(|r| r.iter().map(span).sum::<u32>()).max().unwrap_or(1).clamp(1, 63);
        let colw = (width / cols).max(360);
        out.push_str("<w:tbl><w:tblPr><w:tblStyle w:val=\"TableGrid\"/>");
        out.push_str(&format!("<w:tblW w:w=\"{}\" w:type=\"dxa\"/>", colw * cols));
        out.push_str("<w:tblBorders>");
        for side in ["top", "left", "bottom", "right", "insideH", "insideV"] {
            out.push_str(&format!("<w:{side} w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>"));
        }
        out.push_str("</w:tblBorders><w:tblLayout w:type=\"fixed\"/><w:tblLook w:val=\"04A0\" w:firstRow=\"1\" w:lastRow=\"0\" w:firstColumn=\"1\" w:lastColumn=\"0\" w:noHBand=\"0\" w:noVBand=\"1\"/></w:tblPr><w:tblGrid>");
        for _ in 0..cols {
            out.push_str(&format!("<w:gridCol w:w=\"{colw}\"/>"));
        }
        out.push_str("</w:tblGrid>");
        for row in rows {
            out.push_str("<w:tr>");
            let total: u32 = row.iter().map(span).sum::<u32>().min(cols);
            if total < cols {
                out.push_str(&format!("<w:trPr><w:gridAfter w:val=\"{}\"/></w:trPr>", cols - total));
            }
            let mut used = 0u32;
            for cell in row {
                if used >= cols {
                    break;
                }
                let s = span(cell).min(cols - used);
                used += s;
                out.push_str(&format!("<w:tc><w:tcPr><w:tcW w:w=\"{}\" w:type=\"dxa\"/>", colw * s));
                if s > 1 {
                    out.push_str(&format!("<w:gridSpan w:val=\"{s}\"/>"));
                }
                out.push_str("</w:tcPr>");
                self.in_list = false;
                let inner = (colw * s).saturating_sub(216).max(360);
                if !self.blocks(&cell.blocks, out, inner) {
                    out.push_str("<w:p/>");
                }
                self.in_list = false;
                out.push_str("</w:tc>");
            }
            out.push_str("</w:tr>");
        }
        out.push_str("</w:tbl>");
        true
    }

    fn numbering_xml(&self) -> String {
        let mut s = format!("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<w:numbering xmlns:w=\"{W_NS}\">");
        // abstractNum 0: bullets (Word's default glyph set).
        s.push_str("<w:abstractNum w:abstractNumId=\"0\"><w:multiLevelType w:val=\"hybridMultilevel\"/>");
        for l in 0..9u32 {
            let (glyph, font) = match l % 3 {
                0 => ("\u{F0B7}", "Symbol"),
                1 => ("o", "Courier New"),
                _ => ("\u{F0A7}", "Wingdings"),
            };
            s.push_str(&format!(
                "<w:lvl w:ilvl=\"{l}\"><w:start w:val=\"1\"/><w:numFmt w:val=\"bullet\"/><w:lvlText w:val=\"{glyph}\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"{}\" w:hanging=\"360\"/></w:pPr><w:rPr><w:rFonts w:ascii=\"{font}\" w:hAnsi=\"{font}\" w:hint=\"default\"/></w:rPr></w:lvl>",
                720 * (l + 1)
            ));
        }
        s.push_str("</w:abstractNum>");
        // abstractNum 1: 1. / a. / i. cycling.
        s.push_str("<w:abstractNum w:abstractNumId=\"1\"><w:multiLevelType w:val=\"hybridMultilevel\"/>");
        for l in 0..9u32 {
            let fmt = ["decimal", "lowerLetter", "lowerRoman"][(l % 3) as usize];
            s.push_str(&format!(
                "<w:lvl w:ilvl=\"{l}\"><w:start w:val=\"1\"/><w:numFmt w:val=\"{fmt}\"/><w:lvlText w:val=\"%{}.\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"{}\" w:hanging=\"360\"/></w:pPr></w:lvl>",
                l + 1,
                720 * (l + 1)
            ));
        }
        s.push_str("</w:abstractNum>");
        for (id, abs, restart) in &self.nums {
            s.push_str(&format!("<w:num w:numId=\"{id}\"><w:abstractNumId w:val=\"{abs}\"/>"));
            if let Some(l) = restart {
                s.push_str(&format!("<w:lvlOverride w:ilvl=\"{l}\"><w:startOverride w:val=\"1\"/></w:lvlOverride>"));
            }
            s.push_str("</w:num>");
        }
        s.push_str("</w:numbering>");
        s
    }
}

fn push_run(r: &Run, linked: bool, out: &mut String) {
    let content = run_content(&r.t);
    if content.is_empty() {
        return;
    }
    out.push_str("<w:r>");
    out.push_str(&rpr_xml(r, linked));
    out.push_str(&content);
    out.push_str("</w:r>");
}

const STYLES_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>
<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
<w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="360" w:after="120"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:bCs/><w:color w:val="1F3864"/><w:sz w:val="36"/><w:szCs w:val="36"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:unhideWhenUsed/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="240" w:after="80"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:bCs/><w:color w:val="1F3864"/><w:sz w:val="30"/><w:szCs w:val="30"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:unhideWhenUsed/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="200" w:after="80"/><w:outlineLvl w:val="2"/></w:pPr><w:rPr><w:b/><w:bCs/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading4"><w:name w:val="heading 4"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:unhideWhenUsed/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="160" w:after="40"/><w:outlineLvl w:val="3"/></w:pPr><w:rPr><w:b/><w:bCs/><w:i/><w:iCs/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading5"><w:name w:val="heading 5"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:unhideWhenUsed/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="120" w:after="40"/><w:outlineLvl w:val="4"/></w:pPr><w:rPr><w:b/><w:bCs/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading6"><w:name w:val="heading 6"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:unhideWhenUsed/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="120" w:after="40"/><w:outlineLvl w:val="5"/></w:pPr><w:rPr><w:b/><w:bCs/><w:i/><w:iCs/><w:color w:val="404040"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="10"/><w:qFormat/><w:pPr><w:spacing w:after="80" w:line="240" w:lineRule="auto"/><w:contextualSpacing/></w:pPr><w:rPr><w:sz w:val="56"/><w:szCs w:val="56"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="29"/><w:qFormat/><w:pPr><w:spacing w:before="160"/><w:ind w:left="864" w:right="864"/></w:pPr><w:rPr><w:i/><w:iCs/><w:color w:val="404040"/></w:rPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="Code"><w:name w:val="Code"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:eastAsia="Consolas" w:cs="Consolas"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="34"/><w:qFormat/><w:pPr><w:spacing w:after="40"/><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Header"><w:name w:val="header"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:pPr><w:tabs><w:tab w:val="center" w:pos="4680"/><w:tab w:val="right" w:pos="9360"/></w:tabs><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Footer"><w:name w:val="footer"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:pPr><w:tabs><w:tab w:val="center" w:pos="4680"/><w:tab w:val="right" w:pos="9360"/></w:tabs><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:style>
<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>
<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:uiPriority w:val="39"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr></w:style>
</w:styles>"#;

const SETTINGS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:zoom w:percent="100"/><w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="doNotCompress"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>"#;

fn civil_from_days(z: i64) -> (i64, u32, u32) {
    let z = z + 719_468;
    let era = z.div_euclid(146_097);
    let doe = z - era * 146_097;
    let yoe = (doe - doe / 1460 + doe / 36_524 - doe / 146_096) / 365;
    let y = yoe + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = doy - (153 * mp + 2) / 5 + 1;
    let m = if mp < 10 { mp + 3 } else { mp - 9 };
    (if m <= 2 { y + 1 } else { y }, m as u32, d as u32)
}

fn now_w3cdtf() -> String {
    let secs = std::time::SystemTime::now().duration_since(std::time::UNIX_EPOCH).map(|d| d.as_secs() as i64).unwrap_or(0);
    let (y, m, d) = civil_from_days(secs.div_euclid(86_400));
    let r = secs.rem_euclid(86_400);
    format!("{y:04}-{m:02}-{d:02}T{:02}:{:02}:{:02}Z", r / 3600, r % 3600 / 60, r % 60)
}

/// `meta.header` / `meta.footer` -> (runs, align) when there is any text.
pub fn header_footer_meta(v: Option<&Value>) -> Option<(Vec<Run>, Option<String>)> {
    let o = v?.as_object()?;
    let runs: Vec<Run> = lenient_runs(o.get("runs")).into_iter().filter(|r| !r.t.is_empty()).collect();
    if !runs.iter().any(|r| !r.t.trim().is_empty()) {
        return None;
    }
    let align = o.get("align").and_then(Value::as_str).filter(|a| matches!(*a, "left" | "center" | "right" | "justify")).map(String::from);
    Some((runs, align))
}

/// Split runs at the "{PAGE}" / "{NUMPAGES}" tokens: (run, Some(field)) for tokens.
pub fn split_field_runs(runs: &[Run]) -> Vec<(Run, Option<&'static str>)> {
    let mut out = Vec::new();
    for r in runs {
        let mut rest = r.t.as_str();
        while !rest.is_empty() {
            let next = [("{PAGE}", "PAGE"), ("{NUMPAGES}", "NUMPAGES")]
                .iter()
                .filter_map(|(tok, f)| rest.find(tok).map(|i| (i, *tok, *f)))
                .min_by_key(|(i, _, _)| *i);
            match next {
                Some((i, tok, f)) => {
                    if i > 0 {
                        out.push((Run { t: rest[..i].to_string(), ..r.clone() }, None));
                    }
                    out.push((Run { t: String::new(), ..r.clone() }, Some(f)));
                    rest = &rest[i + tok.len()..];
                }
                None => {
                    out.push((Run { t: rest.to_string(), ..r.clone() }, None));
                    rest = "";
                }
            }
        }
    }
    out
}

/// word/header1.xml or footer1.xml (one paragraph, real PAGE / NUMPAGES
/// fields) and its relationships part (for hyperlinks), if any.
fn header_footer_xml(kind: &str, runs: &[Run], align: Option<&str>) -> (String, Option<String>) {
    let mut hw = DocxWriter::new();
    let mut content = String::new();
    let mut pending: Vec<Run> = Vec::new();
    for (run, field) in split_field_runs(runs) {
        match field {
            None => pending.push(run),
            Some(f) => {
                hw.runs(&std::mem::take(&mut pending), &mut content);
                let rpr = rpr_xml(&run, false);
                for inner in [
                    "<w:fldChar w:fldCharType=\"begin\"/>".to_string(),
                    format!("<w:instrText xml:space=\"preserve\"> {f} </w:instrText>"),
                    "<w:fldChar w:fldCharType=\"separate\"/>".to_string(),
                    "<w:t>1</w:t>".to_string(),
                    "<w:fldChar w:fldCharType=\"end\"/>".to_string(),
                ] {
                    content.push_str(&format!("<w:r>{rpr}{inner}</w:r>"));
                }
            }
        }
    }
    hw.runs(&pending, &mut content);
    let style = if kind == "header" { "Header" } else { "Footer" };
    let jc = match align {
        Some("center") => "<w:jc w:val=\"center\"/>",
        Some("right") => "<w:jc w:val=\"right\"/>",
        Some("justify") => "<w:jc w:val=\"both\"/>",
        _ => "",
    };
    let tag = if kind == "header" { "hdr" } else { "ftr" };
    let xml = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<w:{tag} xmlns:w=\"{W_NS}\" xmlns:r=\"{REL_NS}\"><w:p><w:pPr><w:pStyle w:val=\"{style}\"/>{jc}</w:pPr>{content}</w:p></w:{tag}>"
    );
    let rels = if hw.rels.is_empty() {
        None
    } else {
        Some(format!(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"{PKG_REL_NS}\">{}</Relationships>",
            hw.rels.concat()
        ))
    };
    (xml, rels)
}

/// Write a WordDoc model to .docx bytes.
pub fn write_docx(doc: &WordDoc) -> Result<Vec<u8>, String> {
    let page = PageSetup::from_meta(&doc.meta);
    let mut w = DocxWriter::new();
    let mut body = String::new();
    if !w.blocks(&doc.blocks, &mut body, page.text_width()) {
        body.push_str("<w:p/>");
    }
    // Header / footer parts: (part name, kind, xml, rels xml).
    let mut hf_parts: Vec<(String, &str, String, Option<String>)> = Vec::new();
    let mut refs = String::new();
    for kind in ["header", "footer"] {
        let Some((runs, align)) = header_footer_meta(doc.meta.get(kind)) else { continue };
        let (xml, rels) = header_footer_xml(kind, &runs, align.as_deref());
        let name = format!("{kind}1.xml");
        let rid = w.add_rel(kind, &name, false);
        refs.push_str(&format!("<w:{kind}Reference w:type=\"default\" r:id=\"{rid}\"/>"));
        hf_parts.push((name, kind, xml, rels));
    }
    body.push_str(&page.sect_xml(&refs));

    let document_xml = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<w:document xmlns:w=\"{W_NS}\" xmlns:r=\"{REL_NS}\" xmlns:wp=\"{WP_NS}\" xmlns:a=\"{A_NS}\" xmlns:pic=\"{PIC_NS}\"><w:body>{body}</w:body></w:document>"
    );
    let numbering_xml = w.numbering_xml();
    let doc_rels = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"{PKG_REL_NS}\"><Relationship Id=\"rId1\" Type=\"{REL_NS}/styles\" Target=\"styles.xml\"/><Relationship Id=\"rId2\" Type=\"{REL_NS}/numbering\" Target=\"numbering.xml\"/><Relationship Id=\"rId3\" Type=\"{REL_NS}/settings\" Target=\"settings.xml\"/>{}</Relationships>",
        w.rels.concat()
    );
    let mut defaults = String::new();
    for ext in &w.exts {
        defaults.push_str(&format!("<Default Extension=\"{ext}\" ContentType=\"{}\"/>", mime_for_ext(ext)));
    }
    for (name, kind, _, _) in &hf_parts {
        defaults.push_str(&format!(
            "<Override PartName=\"/word/{name}\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.{kind}+xml\"/>"
        ));
    }
    let content_types = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/>{defaults}<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/><Override PartName=\"/word/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\"/><Override PartName=\"/word/numbering.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml\"/><Override PartName=\"/word/settings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml\"/><Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/><Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/></Types>"
    );
    let root_rels = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"{PKG_REL_NS}\"><Relationship Id=\"rId1\" Type=\"{REL_NS}/officeDocument\" Target=\"word/document.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/><Relationship Id=\"rId3\" Type=\"{REL_NS}/extended-properties\" Target=\"docProps/app.xml\"/></Relationships>"
    );
    let prop = |key: &str, tag: &str| -> String {
        doc.meta
            .get(key)
            .and_then(Value::as_str)
            .map(str::trim)
            .filter(|t| !t.is_empty())
            .map(|t| format!("<{tag}>{}</{tag}>", xml_text(t)))
            .unwrap_or_default()
    };
    let props = format!("{}{}{}{}", prop("title", "dc:title"), prop("subject", "dc:subject"), prop("author", "dc:creator"), prop("author", "cp:lastModifiedBy"));
    let now = now_w3cdtf();
    let core = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">{props}<dcterms:created xsi:type=\"dcterms:W3CDTF\">{now}</dcterms:created><dcterms:modified xsi:type=\"dcterms:W3CDTF\">{now}</dcterms:modified></cp:coreProperties>"
    );
    let app = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\"><Application>office-backend</Application></Properties>";

    let mut buf = Cursor::new(Vec::new());
    {
        let mut zip = ZipWriter::new(&mut buf);
        let opt = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
        let mut put = |name: &str, data: &[u8]| -> Result<(), String> {
            zip.start_file(name, opt).map_err(|e| e.to_string())?;
            zip.write_all(data).map_err(|e| e.to_string())
        };
        put("[Content_Types].xml", content_types.as_bytes())?;
        put("_rels/.rels", root_rels.as_bytes())?;
        put("docProps/core.xml", core.as_bytes())?;
        put("docProps/app.xml", app.as_bytes())?;
        put("word/document.xml", document_xml.as_bytes())?;
        put("word/styles.xml", STYLES_XML.as_bytes())?;
        put("word/numbering.xml", numbering_xml.as_bytes())?;
        put("word/settings.xml", SETTINGS_XML.as_bytes())?;
        put("word/_rels/document.xml.rels", doc_rels.as_bytes())?;
        for (name, _, xml, rels) in &hf_parts {
            put(&format!("word/{name}"), xml.as_bytes())?;
            if let Some(r) = rels {
                put(&format!("word/_rels/{name}.rels"), r.as_bytes())?;
            }
        }
        for (name, data) in &w.media {
            put(name, data)?;
        }
        zip.finish().map_err(|e| e.to_string())?;
    }
    Ok(buf.into_inner())
}

// ============================ tests ============================

#[cfg(test)]
pub(crate) mod tests {
    use super::*;

    /// 1x1 transparent PNG.
    pub const PNG_1X1: &[u8] = &[
        0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52, 0x00, 0x00, 0x00, 0x01,
        0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4, 0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41,
        0x54, 0x78, 0x9c, 0x63, 0x60, 0x00, 0x00, 0x00, 0x02, 0x00, 0x01, 0x54, 0xa2, 0x4f, 0x5d, 0x00, 0x00, 0x00, 0x00, 0x49,
        0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
    ];

    pub fn zip_parts(parts: &[(&str, &[u8])]) -> Vec<u8> {
        let mut buf = Cursor::new(Vec::new());
        {
            let mut z = ZipWriter::new(&mut buf);
            let opt = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
            for (n, d) in parts {
                z.start_file(*n, opt).unwrap();
                z.write_all(d).unwrap();
            }
            z.finish().unwrap();
        }
        buf.into_inner()
    }

    fn unzip(bytes: &[u8]) -> HashMap<String, Vec<u8>> {
        let mut z = ZipArchive::new(Cursor::new(bytes)).unwrap();
        let mut m = HashMap::new();
        for i in 0..z.len() {
            let mut f = z.by_index(i).unwrap();
            let mut v = Vec::new();
            f.read_to_end(&mut v).unwrap();
            m.insert(f.name().map(|n| n.to_string()).unwrap_or_default(), v);
        }
        m
    }

    /// Strict well-formedness check: balanced tags, valid attributes, no
    /// XML-forbidden characters, all entities known.
    pub fn assert_well_formed(name: &str, xml_src: &str) {
        for c in xml_src.chars() {
            assert!(xml_char_ok(c), "{name}: forbidden char U+{:04X}", c as u32);
        }
        let mut rdr = quick_xml::Reader::from_str(xml_src);
        let mut depth = 0i64;
        loop {
            match rdr.read_event() {
                Ok(quick_xml::events::Event::Start(e)) => {
                    depth += 1;
                    for a in e.attributes() {
                        let a = a.unwrap_or_else(|err| panic!("{name}: bad attribute: {err}"));
                        a.unescape_value().unwrap_or_else(|err| panic!("{name}: bad attr value: {err}"));
                    }
                }
                Ok(quick_xml::events::Event::Empty(e)) => {
                    for a in e.attributes() {
                        let a = a.unwrap_or_else(|err| panic!("{name}: bad attribute: {err}"));
                        a.unescape_value().unwrap_or_else(|err| panic!("{name}: bad attr value: {err}"));
                    }
                }
                Ok(quick_xml::events::Event::End(_)) => depth -= 1,
                Ok(quick_xml::events::Event::Text(t)) => {
                    t.unescape().unwrap_or_else(|err| panic!("{name}: bad text: {err}"));
                }
                Ok(quick_xml::events::Event::Eof) => break,
                Err(e) => panic!("{name}: not well-formed: {e}"),
                _ => {}
            }
        }
        assert_eq!(depth, 0, "{name}: unbalanced tags");
    }

    fn check_package(bytes: &[u8]) -> HashMap<String, Vec<u8>> {
        let parts = unzip(bytes);
        for (n, d) in &parts {
            if n.ends_with(".xml") || n.ends_with(".rels") {
                assert_well_formed(n, std::str::from_utf8(d).unwrap());
            }
        }
        parts
    }

    fn part(parts: &HashMap<String, Vec<u8>>, name: &str) -> String {
        String::from_utf8(parts.get(name).unwrap_or_else(|| panic!("missing part {name}")).clone()).unwrap()
    }

    fn word(blocks: Vec<Block>) -> WordDoc {
        WordDoc { kind: "word".into(), meta: Map::new(), blocks }
    }

    fn texts(b: &Block) -> String {
        b.para_runs().iter().map(|r| r.t.as_str()).collect()
    }

    fn li(kind: &str, level: u32, t: &str) -> Block {
        Block::List { list: kind.into(), level, runs: vec![Run::text(t)] }
    }

    #[test]
    fn writes_real_numbering_and_round_trips_list_levels() {
        let doc = word(vec![
            li("ol", 0, "one"),
            li("ol", 0, "two"),
            li("ul", 1, "nested bullet"),
            li("ol", 0, "three"),
            Block::P { runs: vec![Run::text("between")], align: None },
            li("ol", 0, "restart"),
        ]);
        let bytes = write_docx(&doc).unwrap();
        let parts = check_package(&bytes);
        let d = part(&parts, "word/document.xml");
        assert!(!d.contains("• ") && !d.contains(">1. <"), "no literal markers: {d}");
        assert!(d.contains("<w:numPr><w:ilvl w:val=\"1\"/>"));
        let num = part(&parts, "word/numbering.xml");
        assert!(num.contains("w:numFmt w:val=\"bullet\"") && num.contains("w:numFmt w:val=\"decimal\""));
        // Two separate ordered lists -> two numbering instances restarting at 1.
        assert!(num.matches("<w:startOverride w:val=\"1\"/>").count() >= 2, "{num}");
        let rels = part(&parts, "word/_rels/document.xml.rels");
        assert!(rels.contains("relationships/styles") && rels.contains("relationships/numbering"));
        let ct = part(&parts, "[Content_Types].xml");
        assert!(ct.contains("/word/numbering.xml"));

        let back = parse_docx(&bytes).unwrap();
        let kinds: Vec<(String, u32, String)> = back
            .blocks
            .iter()
            .filter_map(|b| match b {
                Block::List { list, level, runs } => Some((list.clone(), *level, runs.iter().map(|r| r.t.clone()).collect())),
                _ => None,
            })
            .collect();
        assert_eq!(
            kinds,
            vec![
                ("ol".into(), 0, "one".into()),
                ("ol".into(), 0, "two".into()),
                ("ul".into(), 1, "nested bullet".into()),
                ("ol".into(), 0, "three".into()),
                ("ol".into(), 0, "restart".into()),
            ]
        );
    }

    #[test]
    fn images_round_trip_as_media_parts() {
        let src = data_url("image/png", PNG_1X1);
        let doc = word(vec![Block::Img { src: src.clone(), alt: Some("Logo & \"mark\"".into()), w: Some(40), h: Some(20) }]);
        let bytes = write_docx(&doc).unwrap();
        let parts = check_package(&bytes);
        assert_eq!(parts.get("word/media/image1.png").map(|v| v.as_slice()), Some(PNG_1X1));
        let d = part(&parts, "word/document.xml");
        assert!(d.contains("<wp:extent cx=\"381000\" cy=\"190500\"/>"), "{d}");
        assert!(d.contains("descr=\"Logo &amp; &quot;mark&quot;\""));
        assert!(part(&parts, "[Content_Types].xml").contains("Extension=\"png\""));
        let back = parse_docx(&bytes).unwrap();
        match &back.blocks[0] {
            Block::Img { src: s2, alt, w, h } => {
                assert_eq!(s2, &src);
                assert_eq!(alt.as_deref(), Some("Logo & \"mark\""));
                assert_eq!((*w, *h), (Some(40), Some(20)));
            }
            other => panic!("expected img, got {other:?}"),
        }
    }

    #[test]
    fn links_with_ampersands_and_anchors() {
        let url = "https://example.com/a?x=1&y=2";
        let doc = word(vec![Block::P {
            runs: vec![
                Run { t: "ext".into(), link: Some(url.into()), ..Default::default() },
                Run::text(" and "),
                Run { t: "internal".into(), link: Some("#sec2".into()), ..Default::default() },
                Run { t: "bad".into(), link: Some("javascript:alert(1)".into()), ..Default::default() },
            ],
            align: None,
        }]);
        let bytes = write_docx(&doc).unwrap();
        let parts = check_package(&bytes);
        let rels = part(&parts, "word/_rels/document.xml.rels");
        assert!(rels.contains("Target=\"https://example.com/a?x=1&amp;y=2\""), "{rels}");
        assert!(!rels.contains("&amp;amp;"));
        assert!(!rels.contains("javascript"));
        let d = part(&parts, "word/document.xml");
        assert!(d.contains("w:anchor=\"sec2\""));
        let back = parse_docx(&bytes).unwrap();
        let runs = back.blocks[0].para_runs();
        assert_eq!(runs[0].link.as_deref(), Some(url));
        assert!(runs.iter().any(|r| r.t == "internal" && r.link.as_deref() == Some("#sec2")));
        assert!(runs.iter().all(|r| r.link.as_deref() != Some("javascript:alert(1)")));
        // Re-saving does not double-escape.
        let again = check_package(&write_docx(&back).unwrap());
        assert!(!part(&again, "word/_rels/document.xml.rels").contains("&amp;amp;"));
    }

    #[test]
    fn tables_colspan_nested_and_empty_rows_are_valid() {
        let cell = |t: &str| TableCell { blocks: vec![Block::P { runs: vec![Run::text(t)], align: None }], colspan: None };
        let nested = Block::Table { rows: vec![vec![cell("inner")]] };
        let doc = word(vec![
            Block::Table {
                rows: vec![
                    vec![TableCell { blocks: vec![Block::P { runs: vec![Run::text("wide")], align: None }], colspan: Some(2) }],
                    vec![cell("a"), TableCell { blocks: vec![nested], colspan: None }],
                    vec![],
                ],
            },
            Block::Table { rows: vec![] },
            Block::Table { rows: vec![vec![cell("second table")]] },
        ]);
        let bytes = write_docx(&doc).unwrap();
        let parts = check_package(&bytes);
        let d = part(&parts, "word/document.xml");
        assert!(!d.contains("<w:tr></w:tr>"));
        assert!(d.contains("<w:gridSpan w:val=\"2\"/>"));
        // A cell whose last block is a table must end with a paragraph.
        assert!(d.contains("</w:tbl><w:p/></w:tc>"), "{d}");
        // Adjacent tables are separated by a paragraph; body ends with one.
        assert!(d.contains("</w:tbl><w:p/><w:tbl>"));
        assert!(d.contains("</w:tbl><w:p/><w:sectPr>"));
        let back = parse_docx(&bytes).unwrap();
        match &back.blocks[0] {
            Block::Table { rows } => {
                assert_eq!(rows[0][0].colspan, Some(2));
                assert!(matches!(rows[1][1].blocks[0], Block::Table { .. }));
            }
            other => panic!("{other:?}"),
        }
    }

    #[test]
    fn control_characters_never_break_the_xml() {
        let doc = word(vec![Block::P {
            runs: vec![Run { t: "ctrl\u{1}char, formfeed\u{c}page, \u{b}vt, \u{fffe}x".into(), font: Some("A\u{2}rial".into()), ..Default::default() }],
            align: None,
        }]);
        let bytes = write_docx(&doc).unwrap();
        let parts = check_package(&bytes);
        let d = part(&parts, "word/document.xml");
        assert!(d.contains("<w:br w:type=\"page\"/>"));
        assert!(d.contains("<w:br/>"));
        assert!(d.contains("w:ascii=\"Arial\""));
    }

    #[test]
    fn rpr_children_follow_schema_order() {
        let r = Run {
            t: "x".into(),
            b: true,
            i: true,
            u: true,
            s: true,
            sup: true,
            color: Some("#ff0000".into()),
            size: Some(14.0),
            font: Some("Arial".into()),
            hl: Some("#FFFF00".into()),
            ..Default::default()
        };
        let x = rpr_xml(&r, true);
        let order = ["w:rStyle", "w:rFonts", "<w:b/>", "<w:i/>", "w:strike", "w:color", "<w:sz ", "w:highlight", "<w:u ", "w:vertAlign"];
        let mut last = 0;
        for tag in order {
            let p = x.find(tag).unwrap_or_else(|| panic!("{tag} missing in {x}"));
            assert!(p >= last, "{tag} out of order in {x}");
            last = p;
        }
        // Non-palette highlight uses shading, which comes after u.
        let x2 = rpr_xml(&Run { t: "x".into(), u: true, hl: Some("#123456".into()), ..Default::default() }, false);
        assert!(x2.find("<w:u ").unwrap() < x2.find("<w:shd ").unwrap());
    }

    #[test]
    fn styles_code_quote_highlight_and_page_setup_round_trip() {
        let mut meta = Map::new();
        meta.insert("page".into(), json!({"size": "A4", "orientation": "landscape", "margins": {"top": 720, "right": 720, "bottom": 720, "left": 720}}));
        meta.insert("title".into(), json!("Quarterly <report>"));
        let doc = WordDoc {
            kind: "word".into(),
            meta,
            blocks: vec![
                Block::H2 { runs: vec![Run::text("Heading")], align: Some("center".into()) },
                Block::Code { runs: vec![Run::text("fn main() {\n    x();\n}")], align: None },
                Block::Quote { runs: vec![Run::text("quoted")], align: None },
                Block::P { runs: vec![Run { t: "marked".into(), hl: Some("#ffff00".into()), ..Default::default() }, Run { t: "shaded".into(), hl: Some("#ABCDEF".into()), ..Default::default() }], align: None },
                Block::PageBreak,
                Block::P { runs: vec![Run::text("after")], align: None },
            ],
        };
        let bytes = write_docx(&doc).unwrap();
        let parts = check_package(&bytes);
        let d = part(&parts, "word/document.xml");
        assert!(d.contains("<w:pgSz w:w=\"16838\" w:h=\"11906\" w:orient=\"landscape\"/>"), "{d}");
        assert!(part(&parts, "word/styles.xml").contains("w:styleId=\"Code\""));
        assert!(part(&parts, "docProps/core.xml").contains("<dc:title>Quarterly &lt;report&gt;</dc:title>"));
        let back = parse_docx(&bytes).unwrap();
        assert!(matches!(&back.blocks[0], Block::H2 { align: Some(a), .. } if a == "center"));
        assert!(matches!(&back.blocks[1], Block::Code { .. }));
        assert_eq!(texts(&back.blocks[1]), "fn main() {\n    x();\n}");
        assert!(matches!(&back.blocks[2], Block::Quote { .. }));
        let runs = back.blocks[3].para_runs();
        assert_eq!(runs[0].hl.as_deref(), Some("#FFFF00"));
        assert_eq!(runs[1].hl.as_deref(), Some("#ABCDEF"));
        assert!(matches!(back.blocks[4], Block::PageBreak));
        assert_eq!(back.meta["title"], json!("Quarterly <report>"));
        assert_eq!(back.meta["page"]["size"], json!("A4"));
        assert_eq!(back.meta["page"]["orientation"], json!("landscape"));
        assert_eq!(back.meta["page"]["margins"]["left"], json!(720));
    }

    #[test]
    fn loosely_typed_blocks_are_not_dropped() {
        let blocks: Vec<Block> = serde_json::from_value(json!([
            {"type": "list", "list": "ul", "level": -1, "runs": [{"t": "negative level item"}]},
            {"type": "img", "src": data_url("image/png", PNG_1X1), "w": 100.5, "h": 50.2},
            {"type": "hr"},
            {"type": "callout", "runs": [{"t": "unknown block", "size": "12"}]},
        ]))
        .unwrap();
        assert!(blocks.iter().all(|b| matches!(b, Block::Other(_))));
        let bytes = write_docx(&word(blocks)).unwrap();
        let back = parse_docx(&bytes).unwrap();
        assert!(back.blocks.iter().any(|b| matches!(b, Block::List { level: 0, .. }) && texts(b) == "negative level item"));
        assert!(back.blocks.iter().any(|b| matches!(b, Block::Img { w: Some(101), h: Some(50), .. })));
        assert!(back.blocks.iter().any(|b| texts(b) == "unknown block"));
    }

    const W: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape""#;

    const NUMBERING: &str = r#"<?xml version="1.0" encoding="UTF-8"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/></w:lvl><w:lvl w:ilvl="1"><w:numFmt w:val="bullet"/></w:lvl></w:abstractNum>
<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl></w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>"#;

    const STYLES: &str = r#"<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="ListBullet"><w:name w:val="List Bullet"/><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="berschrift1"><w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Subtitle"><w:name w:val="Subtitle"/></w:style>
<w:style w:type="paragraph" w:styleId="MyChapter"><w:name w:val="My Chapter"/><w:basedOn w:val="berschrift1"/></w:style>
<w:style w:type="paragraph" w:styleId="Outline3"><w:name w:val="Custom Outline"/><w:pPr><w:outlineLvl w:val="2"/></w:pPr></w:style>
<w:style w:type="character" w:styleId="Strong"><w:name w:val="Strong"/><w:rPr><w:b/></w:rPr></w:style>
</w:styles>"#;

    fn rich_docx(body: &str) -> Vec<u8> {
        let rels = r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/a?x=1&amp;y=2" TargetMode="External"/>
<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
</Relationships>"#;
        let root = r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;
        let foot = format!(r#"<?xml version="1.0"?><w:footnotes {W}><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> The footnote text.</w:t></w:r></w:p></w:footnote></w:footnotes>"#);
        let doc = format!(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}</w:body></w:document>"#);
        zip_parts(&[
            ("[Content_Types].xml", b"<Types/>"),
            ("_rels/.rels", root.as_bytes()),
            ("word/_rels/document.xml.rels", rels.as_bytes()),
            ("word/numbering.xml", NUMBERING.as_bytes()),
            ("word/styles.xml", STYLES.as_bytes()),
            ("word/media/image1.png", PNG_1X1),
            ("word/footnotes.xml", foot.as_bytes()),
            ("word/header1.xml", b"<w:hdr/>"),
            ("word/document.xml", doc.as_bytes()),
        ])
    }

    fn find_text<'a>(blocks: &'a [Block], needle: &str) -> &'a Block {
        blocks.iter().find(|b| texts(b).contains(needle)).unwrap_or_else(|| panic!("no block containing {needle:?}: {blocks:#?}"))
    }

    #[test]
    fn parses_styles_lists_fields_and_content_controls() {
        let body = r#"
<w:p><w:pPr><w:pStyle w:val="berschrift1"/></w:pPr><w:r><w:t>German heading</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="MyChapter"/></w:pPr><w:r><w:t>Inherited heading</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="Outline3"/></w:pPr><w:r><w:t>Outline heading</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="Subtitle"/></w:pPr><w:r><w:t>A subtitle</w:t></w:r></w:p>
<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>bullet numPr</w:t></w:r></w:p>
<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr><w:r><w:t>decimal numPr</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="ListBullet"/></w:pPr><w:r><w:t>bullet via style</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="ListBullet"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr></w:pPr><w:r><w:t>numId zero</w:t></w:r></w:p>
<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="77"/></w:numPr></w:pPr><w:r><w:t>unknown numId</w:t></w:r></w:p>
<w:p><w:r><w:t xml:space="preserve">Field: </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> HYPERLINK "https://field.example" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>field link text</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
<w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE \* MERGEFORMAT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>7</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t xml:space="preserve"> of doc</w:t></w:r></w:p>
<w:p><w:r><w:t xml:space="preserve">Simple </w:t></w:r><w:fldSimple w:instr=" DATE "><w:r><w:t>2024-01-02</w:t></w:r></w:fldSimple><w:r><w:t xml:space="preserve"> end</w:t></w:r></w:p>
<w:p><w:r><w:t xml:space="preserve">Before sdt </w:t></w:r><w:sdt><w:sdtPr><w:alias w:val="x"/></w:sdtPr><w:sdtContent><w:r><w:t>INSIDE_CONTENT_CONTROL</w:t></w:r></w:sdtContent></w:sdt><w:r><w:t xml:space="preserve"> after sdt</w:t></w:r></w:p>
<w:sdt><w:sdtContent><w:p><w:r><w:t>Block level control</w:t></w:r></w:p></w:sdtContent></w:sdt>
<w:p><w:hyperlink r:id="rId4"><w:r><w:t>amp link</w:t></w:r></w:hyperlink><w:r><w:t xml:space="preserve"> and </w:t></w:r><w:hyperlink w:anchor="sec2"><w:r><w:t>internal anchor</w:t></w:r></w:hyperlink></w:p>
<w:p><w:r><w:t xml:space="preserve">kept </w:t></w:r><w:ins w:id="1"><w:r><w:t>inserted</w:t></w:r></w:ins><w:del w:id="2"><w:r><w:delText>deleted</w:delText></w:r></w:del></w:p>
<w:p><w:r><w:rPr><w:rStyle w:val="Strong"/><w:rFonts w:ascii="Arial"/><w:color w:val="FF0000"/><w:sz w:val="28"/></w:rPr><w:t>styled</w:t></w:r><w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t>hl</w:t></w:r><w:r><w:rPr><w:shd w:val="clear" w:fill="00FF00"/></w:rPr><w:t>shd</w:t></w:r><w:r><w:rPr><w:color w:val="auto"/></w:rPr><w:t>auto</w:t></w:r></w:p>
<w:p><w:r><w:t xml:space="preserve">Note here</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r></w:p>
<w:sectPr><w:headerReference w:type="default" r:id="rId5"/><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/></w:sectPr>"#;
        let doc = parse_docx(&rich_docx(body)).unwrap();
        let b = &doc.blocks;
        assert!(matches!(find_text(b, "German heading"), Block::H1 { .. }));
        assert!(matches!(find_text(b, "Inherited heading"), Block::H1 { .. }));
        assert!(matches!(find_text(b, "Outline heading"), Block::H3 { .. }));
        assert!(matches!(find_text(b, "A subtitle"), Block::P { .. }));
        assert!(matches!(find_text(b, "bullet numPr"), Block::List { list, .. } if list == "ul"));
        assert!(matches!(find_text(b, "decimal numPr"), Block::List { list, .. } if list == "ol"));
        assert!(matches!(find_text(b, "bullet via style"), Block::List { list, .. } if list == "ul"));
        assert!(matches!(find_text(b, "numId zero"), Block::P { .. }));
        assert!(matches!(find_text(b, "unknown numId"), Block::P { .. }));
        let f = find_text(b, "field link text");
        assert_eq!(texts(f), "Field: field link text");
        assert!(f.para_runs().iter().any(|r| r.t == "field link text" && r.link.as_deref() == Some("https://field.example")));
        assert_eq!(texts(find_text(b, "Page ")), "Page 7 of doc");
        assert_eq!(texts(find_text(b, "Simple")), "Simple 2024-01-02 end");
        assert_eq!(texts(find_text(b, "Before sdt")), "Before sdt INSIDE_CONTENT_CONTROL after sdt");
        find_text(b, "Block level control");
        let l = find_text(b, "amp link");
        assert_eq!(l.para_runs()[0].link.as_deref(), Some("https://example.com/a?x=1&y=2"));
        assert!(l.para_runs().iter().any(|r| r.t == "internal anchor" && r.link.as_deref() == Some("#sec2")));
        assert_eq!(texts(find_text(b, "kept")), "kept inserted");
        let s = find_text(b, "styled").para_runs();
        assert!(s[0].b && s[0].font.as_deref() == Some("Arial") && s[0].color.as_deref() == Some("#FF0000") && s[0].size == Some(14.0));
        assert_eq!(s.iter().find(|r| r.t == "hl").unwrap().hl.as_deref(), Some("#FFFF00"));
        assert_eq!(s.iter().find(|r| r.t == "shd").unwrap().hl.as_deref(), Some("#00FF00"));
        assert_eq!(s.iter().find(|r| r.t == "auto").unwrap().color, None);
        let note = find_text(b, "Note here").para_runs();
        assert!(note.iter().any(|r| r.t == "1" && r.sup));
        assert!(texts(b.last().unwrap()).contains("The footnote text."));
        let warnings = doc.meta["warnings"].as_array().unwrap();
        let has = |s: &str| warnings.iter().any(|w| w.as_str().unwrap().contains(s));
        // The (empty) default header holds nothing to lose: no warning, no meta.header.
        assert!(!has("header"), "{warnings:?}");
        assert!(doc.meta.get("header").is_none());
        assert!(has("Tracked changes"));
        assert!(has("Fields"));
        assert!(has("Footnotes"));
        assert_eq!(doc.meta["page"]["size"], json!("A4"));
    }

    #[test]
    fn parses_images_text_boxes_and_vml_without_losing_text() {
        let body = r#"
<w:p><w:r><w:t xml:space="preserve">Image here: </w:t></w:r><w:r><w:drawing><wp:inline><wp:extent cx="952500" cy="476250"/><wp:docPr id="1" name="Picture 1" descr="Company logo"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rId3"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r><w:r><w:t xml:space="preserve"> after.</w:t></w:r></w:p>
<w:p><w:r><w:t xml:space="preserve">Before shape </w:t></w:r><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor><wp:extent cx="100" cy="100"/><wp:docPr id="2" name="Text Box 2"/><a:graphic><a:graphicData><wps:wsp><wps:txbx><w:txbxContent><w:p><w:r><w:t>BOX</w:t></w:r></w:p></w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict><v:shape><v:textbox><w:txbxContent><w:p><w:r><w:t>BOX</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r><w:r><w:t>after shape.</w:t></w:r></w:p>
<w:p><w:r><w:t xml:space="preserve">VML: </w:t></w:r><w:r><w:pict><v:shape style="width:75pt;height:30pt"><v:imagedata r:id="rId3" o:title="vml pic" xmlns:o="urn:schemas-microsoft-com:office:office"/></v:shape></w:pict></w:r><w:r><w:t xml:space="preserve"> after pict</w:t></w:r></w:p>
<w:p><w:r><w:t>Para after pict</w:t></w:r></w:p>
<w:p><w:r><w:t>text</w:t></w:r><w:r><w:br w:type="page"/></w:r></w:p>
<w:tbl><w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>span</w:t></w:r></w:p></w:tc></w:tr>
<w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>tall</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>r1c2</w:t></w:r></w:p></w:tc></w:tr>
<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc><w:tc><w:p><w:r><w:t>r2c2</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#;
        let doc = parse_docx(&rich_docx(body)).unwrap();
        let b = &doc.blocks;
        let i = b.iter().position(|x| texts(x) == "Image here: ").expect("text before image");
        match &b[i + 1] {
            Block::Img { src, alt, w, h } => {
                assert!(src.starts_with("data:image/png;base64,"));
                assert_eq!(alt.as_deref(), Some("Company logo"));
                assert_eq!((*w, *h), (Some(100), Some(50)));
            }
            other => panic!("{other:?}"),
        }
        assert_eq!(texts(&b[i + 2]), " after.");
        let s = b.iter().position(|x| texts(x).starts_with("Before shape")).unwrap();
        assert_eq!(texts(&b[s]), "Before shape after shape.");
        assert_eq!(texts(&b[s + 1]), "BOX");
        assert_eq!(b.iter().filter(|x| texts(x) == "BOX").count(), 1, "text box must not be duplicated");
        let v = b.iter().position(|x| texts(x) == "VML: ").unwrap();
        assert!(matches!(&b[v + 1], Block::Img { alt: Some(a), w: Some(100), h: Some(40), .. } if a == "vml pic"));
        assert_eq!(texts(&b[v + 2]), " after pict");
        find_text(b, "Para after pict");
        let t = b.iter().position(|x| texts(x) == "text").unwrap();
        assert!(matches!(b[t + 1], Block::PageBreak));
        let table = b.iter().find_map(|x| if let Block::Table { rows } = x { Some(rows) } else { None }).unwrap();
        assert_eq!(table[0][0].colspan, Some(2));
        assert_eq!(table[2].len(), 2);
        let warnings = doc.meta["warnings"].as_array().unwrap();
        assert!(warnings.iter().any(|w| w.as_str().unwrap().contains("merged")));
        assert!(warnings.iter().any(|w| w.as_str().unwrap().contains("Text boxes")));
    }

    #[test]
    fn recovers_lists_damaged_by_older_saves() {
        let body = r#"
<w:p><w:pPr><w:ind w:left="360" w:hanging="360"/></w:pPr><w:r><w:t xml:space="preserve">• </w:t></w:r><w:r><w:t>Revenue up</w:t></w:r></w:p>
<w:p><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:r><w:t xml:space="preserve">    1. </w:t></w:r><w:r><w:t>Hire engineers</w:t></w:r></w:p>
<w:p><w:r><w:t xml:space="preserve">1. </w:t></w:r><w:r><w:t>plain numbered text stays</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="NoSpacing"/><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/></w:rPr></w:pPr><w:r><w:t>let x = 1;</w:t></w:r></w:p>"#;
        let doc = parse_docx(&rich_docx(body)).unwrap();
        let b = &doc.blocks;
        assert!(matches!(&b[0], Block::List { list, level: 0, .. } if list == "ul"));
        assert_eq!(texts(&b[0]), "Revenue up");
        assert!(matches!(&b[1], Block::List { list, level: 1, .. } if list == "ol"));
        assert_eq!(texts(&b[1]), "Hire engineers");
        assert!(matches!(&b[2], Block::P { .. }));
        assert!(matches!(&b[3], Block::Code { .. }));
    }

    #[test]
    fn misnamed_and_broken_inputs_are_routed_or_rejected() {
        let rtf = parse_docx(br"{\rtf1\ansi Hello\par}").unwrap();
        assert_eq!(texts(&rtf.blocks[0]), "Hello");
        assert!(parse_docx(b"just some text").is_err());
        let xlsx = zip_parts(&[("xl/workbook.xml", b"<workbook/>")]);
        assert!(parse_docx(&xlsx).unwrap_err().contains("Excel"));
    }

    #[test]
    fn deeply_nested_xml_does_not_overflow() {
        let mut body = String::new();
        for _ in 0..5000 {
            body.push_str("<w:sdt><w:sdtContent>");
        }
        body.push_str("<w:p><w:r><w:t>deep</w:t></w:r></w:p>");
        for _ in 0..5000 {
            body.push_str("</w:sdtContent></w:sdt>");
        }
        body.push_str("<w:p><w:r><w:t>after</w:t></w:r></w:p>");
        let doc = parse_docx(&rich_docx(&body)).unwrap();
        find_text(&doc.blocks, "after");
    }

    fn hf_doc() -> WordDoc {
        let mut meta = Map::new();
        meta.insert("title".into(), json!("Annual <plan>"));
        meta.insert("author".into(), json!("Ann & Bob"));
        meta.insert("subject".into(), json!("Budget"));
        meta.insert("page".into(), json!({"size": "Legal", "orientation": "portrait", "margins": {"top": 1080, "right": 1080, "bottom": 1080, "left": 1080, "header": 500, "footer": 400}}));
        meta.insert("header".into(), json!({"runs": [{"t": "Acme ", "b": true}, {"t": "site", "link": "https://acme.example/?a=1&b=2"}], "align": "center"}));
        meta.insert("footer".into(), json!({"runs": [{"t": "Page {PAGE} of {NUMPAGES}", "i": true}], "align": "right"}));
        WordDoc { kind: "word".into(), meta, blocks: vec![Block::P { runs: vec![Run::text("Body")], align: None }] }
    }

    #[test]
    fn header_footer_and_properties_round_trip() {
        let bytes = write_docx(&hf_doc()).unwrap();
        let parts = check_package(&bytes);
        let d = part(&parts, "word/document.xml");
        assert!(d.contains("<w:sectPr><w:headerReference w:type=\"default\" r:id=\""), "{d}");
        assert!(d.contains("<w:footerReference w:type=\"default\""), "{d}");
        assert!(d.contains("<w:pgSz w:w=\"12240\" w:h=\"20160\"/>"), "{d}");
        let h = part(&parts, "word/header1.xml");
        assert!(h.starts_with("<?xml") && h.contains("<w:hdr ") && h.contains("<w:pStyle w:val=\"Header\"/><w:jc w:val=\"center\"/>"), "{h}");
        assert!(h.contains("<w:hyperlink r:id="), "{h}");
        assert!(part(&parts, "word/_rels/header1.xml.rels").contains("https://acme.example/?a=1&amp;b=2"));
        let f = part(&parts, "word/footer1.xml");
        assert!(f.contains("<w:instrText xml:space=\"preserve\"> PAGE </w:instrText>"), "{f}");
        assert!(f.contains("<w:instrText xml:space=\"preserve\"> NUMPAGES </w:instrText>"), "{f}");
        assert_eq!(f.matches("w:fldCharType=\"begin\"").count(), 2);
        assert!(!f.contains("{PAGE}"));
        let ct = part(&parts, "[Content_Types].xml");
        assert!(ct.contains("/word/header1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml\""));
        assert!(ct.contains("/word/footer1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml\""));
        let rels = part(&parts, "word/_rels/document.xml.rels");
        assert!(rels.contains("/relationships/header\" Target=\"header1.xml\""));
        assert!(rels.contains("/relationships/footer\" Target=\"footer1.xml\""));
        let core = part(&parts, "docProps/core.xml");
        assert!(core.contains("<dc:title>Annual &lt;plan&gt;</dc:title>"), "{core}");
        assert!(core.contains("<dc:creator>Ann &amp; Bob</dc:creator>"), "{core}");
        assert!(core.contains("<dc:subject>Budget</dc:subject>"), "{core}");
        assert!(part(&parts, "word/styles.xml").contains("w:styleId=\"Footer\""));

        let back = parse_docx(&bytes).unwrap();
        assert!(back.meta.get("warnings").is_none(), "{:?}", back.meta.get("warnings"));
        assert_eq!(back.meta["title"], json!("Annual <plan>"));
        assert_eq!(back.meta["author"], json!("Ann & Bob"));
        assert_eq!(back.meta["subject"], json!("Budget"));
        assert_eq!(back.meta["page"]["size"], json!("Legal"));
        assert_eq!(back.meta["page"]["margins"]["header"], json!(500));
        assert_eq!(back.meta["header"], json!({"runs": [{"t": "Acme ", "b": true}, {"t": "site", "link": "https://acme.example/?a=1&b=2"}], "align": "center"}));
        assert_eq!(back.meta["footer"], json!({"runs": [{"t": "Page {PAGE} of {NUMPAGES}", "i": true}], "align": "right"}));
    }

    #[test]
    fn no_header_parts_without_text() {
        let mut doc = hf_doc();
        doc.meta.insert("header".into(), json!({"runs": [{"t": "   "}]}));
        doc.meta.insert("footer".into(), json!("not an object"));
        let parts = check_package(&write_docx(&doc).unwrap());
        assert!(!parts.contains_key("word/header1.xml") && !parts.contains_key("word/footer1.xml"));
        assert!(!part(&parts, "word/document.xml").contains("Reference"));
    }

    #[test]
    fn reads_word_header_fields_as_tokens() {
        let bytes = include_bytes!(concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/word/hf-fields.docx"));
        let doc = parse_docx(bytes).unwrap();
        assert_eq!(doc.meta["title"], json!("Fixture title"));
        assert_eq!(doc.meta["author"], json!("Fixture Author"));
        assert_eq!(doc.meta["subject"], json!("Header and footer fields"));
        assert_eq!(doc.meta["header"]["align"], json!("center"));
        let hr: String = doc.meta["header"]["runs"].as_array().unwrap().iter().map(|r| r["t"].as_str().unwrap()).collect();
        assert_eq!(hr, "Acme Corp — Quarterly report");
        assert_eq!(doc.meta["header"]["runs"][0]["b"], json!(true));
        let fr: String = doc.meta["footer"]["runs"].as_array().unwrap().iter().map(|r| r["t"].as_str().unwrap()).collect();
        assert_eq!(fr, "Page {PAGE} of {NUMPAGES}");
        assert_eq!(doc.meta["footer"]["align"], json!("right"));
        // Body fields keep their result text.
        assert_eq!(texts(find_text(&doc.blocks, "Body text")), "Body text with a PAGE field: 3");
        let warnings: Vec<String> = doc.meta["warnings"].as_array().unwrap().iter().map(|w| w.as_str().unwrap().to_string()).collect();
        assert!(warnings.iter().any(|w| w.contains("first-page")), "{warnings:?}");
        // Re-written: real fields again.
        let out = check_package(&write_docx(&doc).unwrap());
        assert!(part(&out, "word/footer1.xml").contains("> NUMPAGES <"));
    }

    #[test]
    fn header_with_table_and_picture_is_simplified() {
        let hdr = format!(
            r#"<w:hdr {W}><w:tbl><w:tr><w:tc><w:p><w:r><w:t>Left</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Right </w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>4</w:t></w:r></w:fldSimple></w:p></w:tc></w:tr></w:tbl><w:p/></w:hdr>"#
        );
        let rels = r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>"#;
        let root = r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;
        let doc = format!(r#"<w:document {W}><w:body><w:p><w:r><w:t>x</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rId5"/></w:sectPr></w:body></w:document>"#);
        let bytes = zip_parts(&[
            ("_rels/.rels", root.as_bytes()),
            ("word/_rels/document.xml.rels", rels.as_bytes()),
            ("word/header1.xml", hdr.as_bytes()),
            ("word/document.xml", doc.as_bytes()),
        ]);
        let d = parse_docx(&bytes).unwrap();
        assert_eq!(d.meta["header"]["runs"], json!([{"t": "Left\tRight {PAGE}"}]));
        assert!(d.meta["warnings"].as_array().unwrap().iter().any(|w| w.as_str().unwrap().contains("simplified")));
    }

    #[test]
    fn base64_and_data_urls() {
        for n in 0..10 {
            let data: Vec<u8> = (0..n as u8).collect();
            assert_eq!(base64_decode(&base64_encode(&data)).unwrap(), data);
        }
        let (m, b) = parse_data_url("data:image/svg+xml,%3Csvg%3E").unwrap();
        assert_eq!(m, "image/svg+xml");
        assert_eq!(b, b"<svg>");
        assert_eq!(image_dims(PNG_1X1), Some((1, 1)));
    }
}
