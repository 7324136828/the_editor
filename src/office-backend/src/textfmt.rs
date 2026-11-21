//! Text-family and legacy word formats.
//!
//! Readers (all produce the shared WordDoc model; fidelity notes go to
//! `meta.warnings`, which main.rs lifts into the /api/parse response):
//! * `parse_text`  — .txt (one paragraph per line) and Markdown (headings,
//!   emphasis, links, nested lists, quotes, fenced/indented code, GFM tables,
//!   images). Encoding: BOM, UTF-8, UTF-16 or Windows-1252 fallback.
//! * `parse_rtf`   — a real RTF tokenizer: groups, control words/symbols,
//!   ignorable destinations, \'hh through \ansicpg / \fcharset code pages,
//!   \uN with \ucN, fonts, colors, sizes, highlight, alignment, styles
//!   (headings), lists, tables, hyperlink fields and PNG/JPEG pictures.
//! * `parse_odt`   — headings, lists (with levels), tables (colspan),
//!   bold/italic/etc. via automatic styles, links, text:s / tabs / line
//!   breaks, images, notes.
//! * `parse_doc`   — legacy Word 97-2003 binary: reads the OLE compound
//!   file, the piece table and paragraph properties to recover paragraphs,
//!   headings, simple lists, tables and hyperlinks (formatting is lost; a
//!   warning says so). Misnamed RTF/DOCX/text files are routed correctly.
//!
//! Writers (called by main.rs /api/render): `write_rtf`, `write_markdown`,
//! `write_html`.

use crate::docx::{
    self, data_url, heading_block, is_mono_font, kind_block, merge_runs, push_warning, resolve_block, safe_link,
    split_paragraph, xml, xml::Child, xml::Node, Inline, PKind,
};
use crate::model::{Block, Run, TableCell, WordDoc};
use serde_json::{Map, Value};
use std::collections::HashMap;
use std::io::Cursor;
use zip::ZipArchive;

const MAX_XML_PART: u64 = 256 * 1024 * 1024;
const MAX_IMAGE: u64 = 32 * 1024 * 1024;
const MAX_MEDIA_TOTAL: usize = 160 * 1024 * 1024;

const W_LEGACY_DOC: &str = "Legacy .doc file: text, headings, simple lists and tables were recovered, but character formatting, images, headers/footers and footnotes are not available. Saving creates a new .docx.";
const W_DOC_FALLBACK: &str = "This .doc file could not be fully decoded; only plain text fragments were recovered and some text may be missing or out of order.";
const W_ENCRYPTED: &str = "This document is password-protected (encrypted) and cannot be opened. Remove the password in Word and try again.";
const W_NOT_RTF: &str = "The file did not contain RTF markup and was opened as plain text.";
const W_RTF_HEADERS: &str = "Headers and footers are not shown and are not kept when saving.";
const W_RTF_NOTES: &str = "Footnotes are not shown and are not kept when saving.";
const W_RTF_COMMENTS: &str = "Comments (annotations) are not shown and are not kept when saving.";
const W_RTF_PICT: &str = "Some pictures use formats that cannot be displayed (old Windows bitmaps) and were left out.";
const W_IMG_FORMAT: &str = "Some pictures use formats browsers cannot display (EMF/WMF/TIFF); they are kept when saving as .docx.";
const W_IMG_LIMIT: &str = "Some pictures were too large to load and were left out.";
const W_FIELDS: &str = "Fields (page numbers, dates, tables of contents, cross-references) show their last calculated text and are saved as plain text.";
const W_ODT_NOTES: &str = "Footnotes/endnotes were converted to numbered notes at the end of the document.";
const W_ODT_COMMENTS: &str = "Comments are not shown and are not kept when saving.";
const W_ODT_TRACKED: &str = "Tracked changes are not shown; the current text is displayed.";
const W_ODT_SHAPES: &str = "Some drawings or embedded objects could not be displayed and are not kept when saving.";
const W_ODT_TEXTBOX: &str = "Text boxes were converted to normal paragraphs placed after the paragraph that contained them.";
const W_ODT_VMERGE: &str = "Vertically merged table cells were split into separate cells.";
const W_ODT_HEADERS: &str = "Headers and footers are not shown and are not kept when saving.";

fn doc_from(blocks: Vec<Block>, warnings: &[String], title: Option<String>) -> WordDoc {
    let mut meta = Map::new();
    if let Some(t) = title.map(|t| t.trim().to_string()).filter(|t| !t.is_empty()) {
        meta.insert("title".into(), Value::String(t));
    }
    for w in warnings {
        push_warning(&mut meta, w);
    }
    let blocks = if blocks.is_empty() { vec![Block::P { runs: Vec::new(), align: None }] } else { blocks };
    WordDoc { kind: "word".into(), meta, blocks }
}

fn warn(list: &mut Vec<String>, msg: &str) {
    if !list.iter().any(|w| w == msg) {
        list.push(msg.to_string());
    }
}

// ============================ code pages ============================

/// Windows-1252 0x80..0x9F (0 = undefined -> C1 control passthrough).
const CP1252_80: [u16; 32] = [
    0x20AC, 0x0081, 0x201A, 0x0192, 0x201E, 0x2026, 0x2020, 0x2021, 0x02C6, 0x2030, 0x0160, 0x2039, 0x0152, 0x008D, 0x017D, 0x008F,
    0x0090, 0x2018, 0x2019, 0x201C, 0x201D, 0x2022, 0x2013, 0x2014, 0x02DC, 0x2122, 0x0161, 0x203A, 0x0153, 0x009D, 0x017E, 0x0178,
];

/// Windows-1250 (Central European) 0x80..0xFF.
const CP1250: [u16; 128] = [
    0x20AC, 0x0081, 0x201A, 0x0083, 0x201E, 0x2026, 0x2020, 0x2021, 0x0088, 0x2030, 0x0160, 0x2039, 0x015A, 0x0164, 0x017D, 0x0179,
    0x0090, 0x2018, 0x2019, 0x201C, 0x201D, 0x2022, 0x2013, 0x2014, 0x0098, 0x2122, 0x0161, 0x203A, 0x015B, 0x0165, 0x017E, 0x017A,
    0x00A0, 0x02C7, 0x02D8, 0x0141, 0x00A4, 0x0104, 0x00A6, 0x00A7, 0x00A8, 0x00A9, 0x015E, 0x00AB, 0x00AC, 0x00AD, 0x00AE, 0x017B,
    0x00B0, 0x00B1, 0x02DB, 0x0142, 0x00B4, 0x00B5, 0x00B6, 0x00B7, 0x00B8, 0x0105, 0x015F, 0x00BB, 0x013D, 0x02DD, 0x013E, 0x017C,
    0x0154, 0x00C1, 0x00C2, 0x0102, 0x00C4, 0x0139, 0x0106, 0x00C7, 0x010C, 0x00C9, 0x0118, 0x00CB, 0x011A, 0x00CD, 0x00CE, 0x010E,
    0x0110, 0x0143, 0x0147, 0x00D3, 0x00D4, 0x0150, 0x00D6, 0x00D7, 0x0158, 0x016E, 0x00DA, 0x0170, 0x00DC, 0x00DD, 0x0162, 0x00DF,
    0x0155, 0x00E1, 0x00E2, 0x0103, 0x00E4, 0x013A, 0x0107, 0x00E7, 0x010D, 0x00E9, 0x0119, 0x00EB, 0x011B, 0x00ED, 0x00EE, 0x010F,
    0x0111, 0x0144, 0x0148, 0x00F3, 0x00F4, 0x0151, 0x00F6, 0x00F7, 0x0159, 0x016F, 0x00FA, 0x0171, 0x00FC, 0x00FD, 0x0163, 0x02D9,
];

/// Windows-1251 (Cyrillic) 0x80..0xBF; 0xC0..0xFF map to U+0410..U+044F.
const CP1251_80: [u16; 64] = [
    0x0402, 0x0403, 0x201A, 0x0453, 0x201E, 0x2026, 0x2020, 0x2021, 0x20AC, 0x2030, 0x0409, 0x2039, 0x040A, 0x040C, 0x040B, 0x040F,
    0x0452, 0x2018, 0x2019, 0x201C, 0x201D, 0x2022, 0x2013, 0x2014, 0x0098, 0x2122, 0x0459, 0x203A, 0x045A, 0x045C, 0x045B, 0x045F,
    0x00A0, 0x040E, 0x045E, 0x0408, 0x00A4, 0x0490, 0x00A6, 0x00A7, 0x0401, 0x00A9, 0x0404, 0x00AB, 0x00AC, 0x00AD, 0x00AE, 0x0407,
    0x00B0, 0x00B1, 0x0406, 0x0456, 0x0491, 0x00B5, 0x00B6, 0x00B7, 0x0451, 0x2116, 0x0454, 0x00BB, 0x0458, 0x0405, 0x0455, 0x0457,
];

/// Mac OS Roman 0x80..0xFF.
const MACROMAN: [u16; 128] = [
    0x00C4, 0x00C5, 0x00C7, 0x00C9, 0x00D1, 0x00D6, 0x00DC, 0x00E1, 0x00E0, 0x00E2, 0x00E4, 0x00E3, 0x00E5, 0x00E7, 0x00E9, 0x00E8,
    0x00EA, 0x00EB, 0x00ED, 0x00EC, 0x00EE, 0x00EF, 0x00F1, 0x00F3, 0x00F2, 0x00F4, 0x00F6, 0x00F5, 0x00FA, 0x00F9, 0x00FB, 0x00FC,
    0x2020, 0x00B0, 0x00A2, 0x00A3, 0x00A7, 0x2022, 0x00B6, 0x00DF, 0x00AE, 0x00A9, 0x2122, 0x00B4, 0x00A8, 0x2260, 0x00C6, 0x00D8,
    0x221E, 0x00B1, 0x2264, 0x2265, 0x00A5, 0x00B5, 0x2202, 0x2211, 0x220F, 0x03C0, 0x222B, 0x00AA, 0x00BA, 0x03A9, 0x00E6, 0x00F8,
    0x00BF, 0x00A1, 0x00AC, 0x221A, 0x0192, 0x2248, 0x2206, 0x00AB, 0x00BB, 0x2026, 0x00A0, 0x00C0, 0x00C3, 0x00D5, 0x0152, 0x0153,
    0x2013, 0x2014, 0x201C, 0x201D, 0x2018, 0x2019, 0x00F7, 0x25CA, 0x00FF, 0x0178, 0x2044, 0x20AC, 0x2039, 0x203A, 0xFB01, 0xFB02,
    0x2021, 0x00B7, 0x201A, 0x201E, 0x2030, 0x00C2, 0x00CA, 0x00C1, 0x00CB, 0x00C8, 0x00CD, 0x00CE, 0x00CF, 0x00CC, 0x00D3, 0x00D4,
    0xF8FF, 0x00D2, 0x00DA, 0x00DB, 0x00D9, 0x0131, 0x02C6, 0x02DC, 0x00AF, 0x02D8, 0x02D9, 0x02DA, 0x00B8, 0x02DD, 0x02DB, 0x02C7,
];

/// Code page used for the Symbol font (bullets etc.).
const CP_SYMBOL: u32 = 42;

/// Decode one byte of a single-byte code page (unknown pages fall back to 1252).
pub fn decode_cp_byte(cp: u32, b: u8) -> char {
    if b < 0x80 {
        return b as char;
    }
    let hi = (b - 0x80) as usize;
    let cp1252 = |b: u8| -> u32 { if b < 0xA0 { CP1252_80[(b - 0x80) as usize] as u32 } else { b as u32 } };
    let v: u32 = match cp {
        1250 => CP1250[hi] as u32,
        1251 => {
            if b >= 0xC0 { 0x0410 + (b - 0xC0) as u32 } else { CP1251_80[hi] as u32 }
        }
        1253 => match b {
            0xA1 => 0x0385,
            0xA2 => 0x0386,
            0xAF => 0x2015,
            0xB4 => 0x0384,
            0xBB => 0x00BB,
            0xBD => 0x00BD,
            0xB8..=0xFE => 0x0388 + (b - 0xB8) as u32,
            _ => cp1252(b),
        },
        1254 => match b {
            0xD0 => 0x011E,
            0xDD => 0x0130,
            0xDE => 0x015E,
            0xF0 => 0x011F,
            0xFD => 0x0131,
            0xFE => 0x015F,
            _ => cp1252(b),
        },
        10000 => MACROMAN[hi] as u32,
        CP_SYMBOL => match b {
            0xB7 => 0x2022,
            0xA7 => 0x25AA,
            0xD8 => 0x27A2,
            0xFC => 0x2713,
            0xAE => 0x2192,
            _ => cp1252(b),
        },
        _ => cp1252(b),
    };
    char::from_u32(v).unwrap_or('\u{FFFD}')
}

/// Reverse Windows-1252 lookup (for RTF \'hh fallbacks).
fn cp1252_byte(c: char) -> Option<u8> {
    let v = c as u32;
    if v < 0x80 || (0xA0..=0xFF).contains(&v) {
        return Some(v as u8);
    }
    CP1252_80.iter().position(|x| *x as u32 == v && v > 0xFF).map(|i| 0x80 + i as u8)
}

/// Decode a text file: BOM (UTF-8 / UTF-16), valid UTF-8, BOM-less UTF-16
/// (detected by NUL patterns), else Windows-1252. The second value is a
/// warning when a legacy encoding had to be guessed.
pub fn decode_text(b: &[u8]) -> (String, Option<&'static str>) {
    if let Some(rest) = b.strip_prefix(&[0xEF, 0xBB, 0xBF]) {
        return (String::from_utf8_lossy(rest).into_owned(), None);
    }
    let utf16 = |data: &[u8], le: bool| -> String {
        let units: Vec<u16> = data
            .chunks_exact(2)
            .map(|c| if le { u16::from_le_bytes([c[0], c[1]]) } else { u16::from_be_bytes([c[0], c[1]]) })
            .collect();
        String::from_utf16_lossy(&units)
    };
    if b.starts_with(&[0xFF, 0xFE]) {
        return (utf16(&b[2..], true), None);
    }
    if b.starts_with(&[0xFE, 0xFF]) {
        return (utf16(&b[2..], false), None);
    }
    if let Ok(s) = std::str::from_utf8(b) {
        return (s.to_string(), None);
    }
    if b.len() >= 4 {
        let sample = &b[..b.len().min(4096) & !1];
        let odd_nuls = sample.iter().skip(1).step_by(2).filter(|x| **x == 0).count();
        let even_nuls = sample.iter().step_by(2).filter(|x| **x == 0).count();
        let half = sample.len() / 2;
        if half > 0 && odd_nuls * 10 > half * 6 && even_nuls * 10 < half {
            return (utf16(b, true), None);
        }
        if half > 0 && even_nuls * 10 > half * 6 && odd_nuls * 10 < half {
            return (utf16(b, false), None);
        }
    }
    (b.iter().map(|x| decode_cp_byte(1252, *x)).collect(), Some("The file is not UTF-8; it was read as Windows-1252 (Western European)."))
}

// ============================ plain text / markdown ============================

/// Parse .txt or .md bytes into a word model.
pub fn parse_text(bytes: &[u8], markdown: bool) -> WordDoc {
    let (text, note) = decode_text(bytes);
    let text = text.replace("\r\n", "\n").replace('\r', "\n");
    let mut warnings = Vec::new();
    if let Some(n) = note {
        warnings.push(n.to_string());
    }
    if markdown {
        return doc_from(parse_markdown(&text, 0), &warnings, None);
    }
    let mut lines: Vec<&str> = text.split('\n').collect();
    if lines.len() > 1 && lines.last() == Some(&"") {
        lines.pop();
    }
    let mut blocks = Vec::with_capacity(lines.len());
    for l in lines {
        for (k, seg) in l.split('\u{c}').enumerate() {
            if k > 0 {
                blocks.push(Block::PageBreak);
            }
            if k == 0 || !seg.is_empty() {
                let runs = if seg.is_empty() { Vec::new() } else { vec![Run::text(seg)] };
                blocks.push(Block::P { runs, align: None });
            }
        }
    }
    doc_from(blocks, &warnings, None)
}

fn expand_leading_tabs(l: &str) -> String {
    let mut out = String::new();
    let mut col = 0usize;
    let mut rest_start = l.len();
    for (i, c) in l.char_indices() {
        match c {
            ' ' => {
                out.push(' ');
                col += 1;
            }
            '\t' => {
                let n = 4 - col % 4;
                out.push_str(&" ".repeat(n));
                col += n;
            }
            _ => {
                rest_start = i;
                break;
            }
        }
    }
    out.push_str(&l[rest_start.min(l.len())..]);
    out
}

fn md_fence(t: &str) -> Option<(char, usize)> {
    let c = t.chars().next()?;
    if c != '`' && c != '~' {
        return None;
    }
    let n = t.chars().take_while(|x| *x == c).count();
    if n < 3 || (c == '`' && t[n..].contains('`')) {
        return None;
    }
    Some((c, n))
}

fn md_atx(t: &str) -> Option<(u32, &str)> {
    let n = t.chars().take_while(|c| *c == '#').count();
    if !(1..=6).contains(&n) {
        return None;
    }
    let rest = &t[n..];
    if !rest.is_empty() && !rest.starts_with(' ') && !rest.starts_with('\t') {
        return None;
    }
    let mut content = rest.trim();
    let stripped = content.trim_end_matches('#');
    if stripped.is_empty() || stripped.ends_with(' ') {
        content = stripped.trim_end();
    }
    Some((n as u32, content))
}

fn md_is_hr(t: &str) -> bool {
    let c = match t.chars().next() {
        Some(c @ ('-' | '*' | '_')) => c,
        _ => return false,
    };
    t.chars().all(|x| x == c || x == ' ' || x == '\t') && t.chars().filter(|x| *x == c).count() >= 3
}

/// Returns (ordered, content) for a list item line (already left-trimmed).
fn md_list_item(t: &str) -> Option<(bool, &str)> {
    let mut chars = t.chars();
    let c = chars.next()?;
    if matches!(c, '-' | '*' | '+') {
        let rest = &t[1..];
        if rest.is_empty() {
            return Some((false, ""));
        }
        if rest.starts_with(' ') || rest.starts_with('\t') {
            return Some((false, rest.trim_start()));
        }
        return None;
    }
    let digits = t.chars().take_while(|c| c.is_ascii_digit()).count();
    if (1..=9).contains(&digits) {
        let rest = &t[digits..];
        if let Some(r) = rest.strip_prefix('.').or_else(|| rest.strip_prefix(')')) {
            if r.is_empty() {
                return Some((true, ""));
            }
            if r.starts_with(' ') || r.starts_with('\t') {
                return Some((true, r.trim_start()));
            }
        }
    }
    None
}

fn md_split_row(line: &str) -> Vec<String> {
    let t = line.trim();
    let t = t.strip_prefix('|').unwrap_or(t);
    let t = if t.ends_with('|') && !t.ends_with("\\|") { &t[..t.len() - 1] } else { t };
    let mut cells = Vec::new();
    let mut cur = String::new();
    let mut prev_bs = false;
    for c in t.chars() {
        if c == '|' && !prev_bs {
            cells.push(cur.trim().to_string());
            cur.clear();
        } else {
            if c == '|' && prev_bs {
                cur.pop();
            }
            cur.push(c);
        }
        prev_bs = c == '\\' && !prev_bs;
    }
    cells.push(cur.trim().to_string());
    cells
}

fn md_is_table_sep(line: &str) -> bool {
    if !line.contains('-') {
        return false;
    }
    let t = line.trim();
    if !t.contains('|') && !t.starts_with(':') {
        return false;
    }
    let cells = md_split_row(line);
    !cells.is_empty()
        && cells.iter().all(|c| {
            let c = c.trim();
            let inner = c.trim_start_matches(':').trim_end_matches(':');
            !inner.is_empty() && inner.chars().all(|x| x == '-')
        })
}

/// `![alt](src)` occupying a whole line.
fn md_image_line(t: &str) -> Option<Block> {
    let chars: Vec<char> = t.trim().chars().collect();
    if chars.first() != Some(&'!') {
        return None;
    }
    let (alt, url, end) = md_parse_link(&chars, 1)?;
    if end != chars.len() {
        return None;
    }
    if !url.trim().to_ascii_lowercase().starts_with("data:image/") {
        return None;
    }
    let dims = docx::parse_data_url(&url).and_then(|(_, b)| docx::image_dims(&b));
    let alt: String = alt.into_iter().collect();
    Some(Block::Img { src: url.trim().to_string(), alt: if alt.is_empty() { None } else { Some(alt) }, w: dims.map(|d| d.0), h: dims.map(|d| d.1) })
}

fn md_starts_block(t: &str) -> bool {
    md_fence(t).is_some() || md_atx(t).is_some() || t.starts_with('>') || md_is_hr(t) || md_list_item(t).is_some()
}

fn md_join_lines(lines: &[String]) -> String {
    let mut s = String::new();
    for (k, l) in lines.iter().enumerate() {
        let last = k + 1 == lines.len();
        if last {
            s.push_str(l.trim_end());
        } else if l.ends_with("  ") {
            s.push_str(l.trim_end());
            s.push('\n');
        } else if let Some(x) = l.trim_end().strip_suffix('\\') {
            s.push_str(x);
            s.push('\n');
        } else {
            s.push_str(l.trim_end());
            s.push(' ');
        }
    }
    s
}

/// Block-level Markdown parser (CommonMark/GFM subset).
fn parse_markdown(text: &str, depth: u32) -> Vec<Block> {
    let lines: Vec<String> = text.split('\n').map(expand_leading_tabs).collect();
    let n = lines.len();
    let mut blocks: Vec<Block> = Vec::new();
    let mut para: Vec<String> = Vec::new();
    let mut stack: Vec<usize> = Vec::new(); // list indents
    let flush = |para: &mut Vec<String>, blocks: &mut Vec<Block>| {
        if !para.is_empty() {
            blocks.push(Block::P { runs: md_inline(&md_join_lines(para)), align: None });
            para.clear();
        }
    };
    let mut i = 0;
    while i < n {
        let line = &lines[i];
        let t = line.trim_start();
        let indent = line.len() - t.len();
        if t.is_empty() {
            flush(&mut para, &mut blocks);
            i += 1;
            continue;
        }
        if indent < 4 || !stack.is_empty() {
            if let Some((fc, fl)) = md_fence(t) {
                flush(&mut para, &mut blocks);
                let mut code: Vec<String> = Vec::new();
                i += 1;
                while i < n {
                    let l = &lines[i];
                    let lt = l.trim_start();
                    if lt.chars().take_while(|c| *c == fc).count() >= fl && lt.trim_start_matches(fc).trim().is_empty() {
                        i += 1;
                        break;
                    }
                    let strip = (l.len() - lt.len()).min(indent);
                    code.push(l[strip..].to_string());
                    i += 1;
                }
                blocks.push(Block::Code { runs: if code.is_empty() { Vec::new() } else { vec![Run::text(code.join("\n"))] }, align: None });
                continue;
            }
        }
        if indent < 4 {
            if let Some((lvl, content)) = md_atx(t) {
                flush(&mut para, &mut blocks);
                stack.clear();
                blocks.push(heading_block(lvl, md_inline(content), None));
                i += 1;
                continue;
            }
            if !para.is_empty() && (t.chars().all(|c| c == '=') || (t.chars().all(|c| c == '-') && t.len() >= 2)) {
                let lvl = if t.starts_with('=') { 1 } else { 2 };
                let runs = md_inline(&md_join_lines(&para));
                para.clear();
                blocks.push(heading_block(lvl, runs, None));
                i += 1;
                continue;
            }
            if md_is_hr(t) && md_list_item(t).map(|(_, c)| c.is_empty() || md_is_hr(c)).unwrap_or(true) {
                flush(&mut para, &mut blocks);
                stack.clear();
                blocks.push(Block::PageBreak);
                i += 1;
                continue;
            }
            if t.starts_with('>') {
                flush(&mut para, &mut blocks);
                stack.clear();
                let mut q: Vec<String> = Vec::new();
                while i < n {
                    let lt = lines[i].trim_start();
                    let Some(rest) = lt.strip_prefix('>') else { break };
                    q.push(rest.strip_prefix(' ').unwrap_or(rest).to_string());
                    i += 1;
                }
                let inner = if depth < 8 { parse_markdown(&q.join("\n"), depth + 1) } else { vec![Block::P { runs: vec![Run::text(q.join("\n"))], align: None }] };
                for b in inner {
                    blocks.push(match b {
                        Block::P { runs, align }
                        | Block::H1 { runs, align }
                        | Block::H2 { runs, align }
                        | Block::H3 { runs, align }
                        | Block::H4 { runs, align }
                        | Block::H5 { runs, align }
                        | Block::H6 { runs, align }
                        | Block::Quote { runs, align } => Block::Quote { runs, align },
                        other => other,
                    });
                }
                continue;
            }
            if t.contains('|') && i + 1 < n && md_is_table_sep(&lines[i + 1]) {
                flush(&mut para, &mut blocks);
                stack.clear();
                let header = md_split_row(t);
                let aligns: Vec<Option<String>> = md_split_row(&lines[i + 1])
                    .iter()
                    .map(|c| {
                        let c = c.trim();
                        match (c.starts_with(':'), c.ends_with(':')) {
                            (true, true) => Some("center".to_string()),
                            (false, true) => Some("right".to_string()),
                            _ => None,
                        }
                    })
                    .collect();
                let cols = header.len();
                let mk = |cells: Vec<String>, bold: bool| -> Vec<TableCell> {
                    (0..cols)
                        .map(|k| {
                            let mut runs = md_inline(cells.get(k).map(String::as_str).unwrap_or(""));
                            if bold {
                                for r in runs.iter_mut() {
                                    r.b = true;
                                }
                            }
                            TableCell { blocks: vec![Block::P { runs, align: aligns.get(k).cloned().flatten() }], colspan: None }
                        })
                        .collect()
                };
                let mut rows = vec![mk(header, true)];
                i += 2;
                while i < n && !lines[i].trim().is_empty() && lines[i].contains('|') {
                    rows.push(mk(md_split_row(&lines[i]), false));
                    i += 1;
                }
                blocks.push(Block::Table { rows });
                continue;
            }
        }
        if let Some((ordered, content)) = md_list_item(t) {
            if para.is_empty() || indent > 0 || !content.is_empty() {
                flush(&mut para, &mut blocks);
                while let Some(&top) = stack.last() {
                    if indent < top {
                        stack.pop();
                    } else {
                        break;
                    }
                }
                let level = match stack.last() {
                    None => {
                        stack.push(indent);
                        0
                    }
                    Some(&top) if indent >= top + 2 => {
                        stack.push(indent);
                        stack.len() - 1
                    }
                    Some(_) => stack.len() - 1,
                };
                let mut text = content.to_string();
                if let Some(r) = text.strip_prefix("[ ] ") {
                    text = format!("☐ {r}");
                } else if let Some(r) = text.strip_prefix("[x] ").or_else(|| text.strip_prefix("[X] ")) {
                    text = format!("☑ {r}");
                }
                let mut parts = vec![text];
                i += 1;
                while i < n {
                    let lt = lines[i].trim_start();
                    if lt.is_empty() || md_starts_block(lt) || (lt.contains('|') && i + 1 < n && md_is_table_sep(&lines[i + 1])) {
                        break;
                    }
                    parts.push(lt.to_string());
                    i += 1;
                }
                blocks.push(Block::List { list: if ordered { "ol".into() } else { "ul".into() }, level: (level as u32).min(8), runs: md_inline(&md_join_lines(&parts)) });
                continue;
            }
        }
        if indent < 4 {
            if let Some(img) = md_image_line(t) {
                flush(&mut para, &mut blocks);
                blocks.push(img);
                i += 1;
                continue;
            }
        }
        if indent >= 4 && para.is_empty() && stack.is_empty() {
            let mut code: Vec<String> = Vec::new();
            while i < n {
                let l = &lines[i];
                if l.trim().is_empty() {
                    let next_indented = lines[i + 1..].iter().find(|x| !x.trim().is_empty()).map(|x| x.starts_with("    ")).unwrap_or(false);
                    if !next_indented {
                        break;
                    }
                    code.push(String::new());
                } else if l.starts_with("    ") {
                    code.push(l[4..].to_string());
                } else {
                    break;
                }
                i += 1;
            }
            blocks.push(Block::Code { runs: vec![Run::text(code.join("\n"))], align: None });
            continue;
        }
        if indent == 0 && para.is_empty() {
            stack.clear();
        }
        para.push(t.to_string() + if line.ends_with("  ") { "  " } else { "" });
        i += 1;
    }
    flush(&mut para, &mut blocks);
    blocks
}

fn run_len(s: &[char], i: usize, c: char) -> usize {
    s[i..].iter().take_while(|x| **x == c).count()
}

/// `[text](url)` starting at s[i] == '['. Returns (text, url, index after ')').
fn md_parse_link(s: &[char], i: usize) -> Option<(Vec<char>, String, usize)> {
    if s.get(i) != Some(&'[') {
        return None;
    }
    let mut depth = 0;
    let mut j = i;
    let close = loop {
        match s.get(j)? {
            '\\' => j += 1,
            '[' => depth += 1,
            ']' => {
                depth -= 1;
                if depth == 0 {
                    break j;
                }
            }
            _ => {}
        }
        j += 1;
    };
    if s.get(close + 1) != Some(&'(') {
        return None;
    }
    let mut k = close + 2;
    let mut pdepth = 1;
    let start = k;
    let end = loop {
        match s.get(k)? {
            '\\' => k += 1,
            '(' => pdepth += 1,
            ')' => {
                pdepth -= 1;
                if pdepth == 0 {
                    break k;
                }
            }
            _ => {}
        }
        k += 1;
    };
    let inner: String = s[start..end].iter().collect();
    let inner = inner.trim();
    let url = if let Some(rest) = inner.strip_prefix('<') {
        rest.split('>').next().unwrap_or("").to_string()
    } else {
        inner.split_whitespace().next().unwrap_or("").to_string()
    };
    Some((s[i + 1..close].to_vec(), url, end + 1))
}

fn md_find_emph_close(s: &[char], from: usize, c: char, n: usize) -> Option<usize> {
    let mut k = from;
    while k < s.len() {
        let ch = s[k];
        if ch == '\\' {
            k += 2;
            continue;
        }
        if ch == '`' {
            let m = run_len(s, k, '`');
            let mut q = k + m;
            let mut found = None;
            while q < s.len() {
                if s[q] == '`' {
                    let mm = run_len(s, q, '`');
                    if mm == m {
                        found = Some(q + mm);
                        break;
                    }
                    q += mm;
                } else {
                    q += 1;
                }
            }
            k = found.unwrap_or(k + m);
            continue;
        }
        if ch == c {
            let m = run_len(s, k, c);
            let ok_prev = k > from && !s[k - 1].is_whitespace();
            let ok_next = !(c == '_' && k + m < s.len() && s[k + m].is_alphanumeric());
            if ok_prev && ok_next && (m == n || (m == 3 && n < 3)) {
                return Some(k + m - n);
            }
            k += m;
            continue;
        }
        k += 1;
    }
    None
}

fn md_inline(s: &str) -> Vec<Run> {
    let chars: Vec<char> = s.chars().collect();
    let mut out = Vec::new();
    md_inline_into(&chars, &Run::default(), &mut out, 0);
    merge_runs(out)
}

fn md_inline_into(s: &[char], fmt: &Run, out: &mut Vec<Run>, depth: u32) {
    let mut buf = String::new();
    let flush = |buf: &mut String, out: &mut Vec<Run>| {
        if !buf.is_empty() {
            let mut r = fmt.clone();
            r.t = std::mem::take(buf);
            out.push(r);
        }
    };
    let mut i = 0;
    while i < s.len() {
        let c = s[i];
        if c == '\\' && i + 1 < s.len() && (s[i + 1].is_ascii_punctuation() || s[i + 1] == '\n') {
            buf.push(s[i + 1]);
            i += 2;
            continue;
        }
        if depth < 12 {
            if c == '`' {
                let n = run_len(s, i, '`');
                let mut q = i + n;
                let mut close = None;
                while q < s.len() {
                    if s[q] == '`' {
                        let m = run_len(s, q, '`');
                        if m == n {
                            close = Some(q);
                            break;
                        }
                        q += m;
                    } else {
                        q += 1;
                    }
                }
                if let Some(j) = close {
                    flush(&mut buf, out);
                    let mut code: String = s[i + n..j].iter().collect();
                    if code.len() >= 2 && code.starts_with(' ') && code.ends_with(' ') && !code.trim().is_empty() {
                        code = code[1..code.len() - 1].to_string();
                    }
                    let mut r = fmt.clone();
                    r.t = code;
                    r.font = Some("Consolas".into());
                    out.push(r);
                    i = j + n;
                } else {
                    buf.extend(std::iter::repeat('`').take(n));
                    i += n;
                }
                continue;
            }
            if c == '*' || c == '_' || c == '~' {
                let n = run_len(s, i, c);
                let can_open = i + n < s.len() && !s[i + n].is_whitespace() && !(c == '_' && i > 0 && s[i - 1].is_alphanumeric());
                if can_open && n <= 3 && (c != '~' || n == 2) {
                    if let Some(j) = md_find_emph_close(s, i + n, c, n) {
                        flush(&mut buf, out);
                        let mut f = fmt.clone();
                        match (c, n) {
                            ('~', _) => f.s = true,
                            (_, 1) => f.i = true,
                            (_, 2) => f.b = true,
                            _ => {
                                f.b = true;
                                f.i = true;
                            }
                        }
                        md_inline_into(&s[i + n..j], &f, out, depth + 1);
                        i = j + n;
                        continue;
                    }
                }
                buf.extend(std::iter::repeat(c).take(n));
                i += n;
                continue;
            }
            if c == '!' && s.get(i + 1) == Some(&'[') {
                if let Some((alt, url, end)) = md_parse_link(s, i + 1) {
                    flush(&mut buf, out);
                    let mut r = fmt.clone();
                    let alt: String = alt.into_iter().collect();
                    r.t = if alt.is_empty() { "[image]".into() } else { alt };
                    if !url.to_ascii_lowercase().starts_with("data:") {
                        r.link = safe_link(&url);
                    }
                    out.push(r);
                    i = end;
                    continue;
                }
            }
            if c == '[' {
                if let Some((text, url, end)) = md_parse_link(s, i) {
                    flush(&mut buf, out);
                    let mut f = fmt.clone();
                    f.link = safe_link(&url);
                    md_inline_into(&text, &f, out, depth + 1);
                    i = end;
                    continue;
                }
            }
            if c == '<' {
                if let Some(rel) = s[i..].iter().position(|x| *x == '>') {
                    let inner: String = s[i + 1..i + rel].iter().collect();
                    let lower = inner.to_ascii_lowercase();
                    let end = i + rel + 1;
                    if !inner.contains(' ') && (lower.starts_with("http://") || lower.starts_with("https://") || lower.starts_with("mailto:") || (inner.contains('@') && !inner.contains('/'))) {
                        flush(&mut buf, out);
                        let mut r = fmt.clone();
                        r.t = inner.trim_start_matches("mailto:").to_string();
                        r.link = safe_link(&if inner.contains('@') && !lower.starts_with("mailto:") { format!("mailto:{inner}") } else { inner.clone() });
                        out.push(r);
                        i = end;
                        continue;
                    }
                    if matches!(lower.trim_end_matches('/').trim(), "br") {
                        buf.push('\n');
                        i = end;
                        continue;
                    }
                    let tag = lower.trim();
                    if matches!(tag, "u" | "ins" | "sup" | "sub" | "s" | "del" | "strike" | "b" | "strong" | "i" | "em" | "code" | "mark") {
                        let closing: Vec<char> = format!("</{tag}>").chars().collect();
                        let rest = &s[end..];
                        if let Some(p) = rest.windows(closing.len()).position(|w| w.iter().zip(&closing).all(|(a, b)| a.to_ascii_lowercase() == *b)) {
                            flush(&mut buf, out);
                            let mut f = fmt.clone();
                            match tag {
                                "u" | "ins" => f.u = true,
                                "sup" => f.sup = true,
                                "sub" => f.sub = true,
                                "s" | "del" | "strike" => f.s = true,
                                "b" | "strong" => f.b = true,
                                "i" | "em" => f.i = true,
                                "code" => f.font = Some("Consolas".into()),
                                _ => f.hl = Some("#FFFF00".into()),
                            }
                            md_inline_into(&rest[..p], &f, out, depth + 1);
                            i = end + p + closing.len();
                            continue;
                        }
                    }
                }
            }
        }
        buf.push(c);
        i += 1;
    }
    flush(&mut buf, out);
}

// ============================ RTF ============================

#[derive(Clone, Copy, PartialEq, Debug)]
enum Dest {
    Normal,
    Skip,
    FontTbl,
    ColorTbl,
    StyleSheet,
    Info,
    Title,
    Pict,
    FldInst,
    ListText,
    Pn,
}

#[derive(Clone)]
struct RState {
    dest: Dest,
    b: bool,
    i: bool,
    u: bool,
    s: bool,
    sup: bool,
    sub: bool,
    hidden: bool,
    cf: i32,
    hl: i32,
    fs: i32,
    font: i32,
    uc: usize,
    align: Option<&'static str>,
    style: i32,
    intbl: bool,
    ls: Option<i32>,
    ilvl: u32,
    outline: Option<u32>,
    pn: Option<(bool, u32)>,
    link: Option<String>,
}

impl RState {
    fn new() -> Self {
        RState {
            dest: Dest::Normal,
            b: false,
            i: false,
            u: false,
            s: false,
            sup: false,
            sub: false,
            hidden: false,
            cf: 0,
            hl: 0,
            fs: 24,
            font: -1,
            uc: 1,
            align: None,
            style: 0,
            intbl: false,
            ls: None,
            ilvl: 0,
            outline: None,
            pn: None,
            link: None,
        }
    }
    fn reset_char(&mut self) {
        self.b = false;
        self.i = false;
        self.u = false;
        self.s = false;
        self.sup = false;
        self.sub = false;
        self.hidden = false;
        self.cf = 0;
        self.hl = 0;
        self.fs = 24;
        self.font = -1;
    }
    fn reset_para(&mut self) {
        self.align = None;
        self.style = 0;
        self.intbl = false;
        self.ls = None;
        self.ilvl = 0;
        self.outline = None;
        self.pn = None;
    }
}

#[derive(Default)]
struct Pict {
    mime: Option<&'static str>,
    data: Vec<u8>,
    half: Option<u8>,
    picw: i64,
    pich: i64,
    goalw: i64,
    goalh: i64,
    scalex: i64,
    scaley: i64,
    too_big: bool,
}

fn charset_cp(n: i32) -> u32 {
    match n {
        2 => CP_SYMBOL,
        77 => 10000,
        128 => 932,
        129 => 949,
        134 => 936,
        136 => 950,
        161 => 1253,
        162 => 1254,
        163 => 1258,
        177 => 1255,
        178 => 1256,
        186 => 1257,
        204 => 1251,
        222 => 874,
        238 => 1250,
        255 => 437,
        _ => 0,
    }
}

const RTF_SKIP_DESTS: &[&str] = &[
    "xe", "tc", "tcn", "txe", "rxe", "listtable", "listoverridetable", "revtbl", "rsidtbl", "themedata",
    "colorschememapping", "latentstyles", "datastore", "xmlnstbl", "generator", "pgdsctbl", "filetbl", "mmathPr",
    "objdata", "objclass", "objname", "bkmkstart", "bkmkend", "pntxta", "pntxtb", "userprops", "docvar", "template",
    "ftnsep", "ftnsepc", "ftncn", "aftnsep", "aftnsepc", "aftncn", "do", "sp", "sn", "sv", "shpinst", "fchars",
    "lchars", "nonshppict", "nonesttables", "nesttableprops", "blipuid", "picprop", "wgrffmtfilter", "defchp",
    "defpap", "passwordhash", "protusertbl", "ud", "pgptbl", "fldtype", "formfield", "datafield", "falt", "panose",
    "fname", "keycode", "pxe", "atnid", "atnauthor", "atrfstart", "atrfend", "atnref", "atndate", "atnparent",
    "atnicn", "atntime", "listpicture", "background", "mhtmltag", "htmltag", "oldcprops", "oldpprops", "oldsprops",
    "oldtprops", "stylesheet_", "ebcstart", "ebcend",
];

struct RtfParser<'a> {
    d: &'a [u8],
    pos: usize,
    st: RState,
    stack: Vec<RState>,
    overflow: usize,
    cp: u32,
    deff: i32,
    fonts: HashMap<i32, (String, u32)>,
    font_buf: String,
    font_cur: i32,
    font_cp: u32,
    colors: Vec<Option<String>>,
    rgb: [u8; 3],
    rgb_set: bool,
    styles: HashMap<i32, (String, Option<u32>)>,
    style_buf: String,
    style_cur: i32,
    style_outline: Option<u32>,
    title: String,
    pict: Option<Pict>,
    fields: Vec<(usize, String)>,
    listtext: String,
    pn_tmp: Option<(bool, u32)>,
    ls_kind: HashMap<(i32, u32), bool>,
    text: String,
    items: Vec<Inline>,
    blocks: Vec<Block>,
    cell_blocks: Vec<Block>,
    cur_row: Vec<TableCell>,
    row_cellx: Vec<(i32, bool)>,
    pending_merge: bool,
    table_rows: Vec<(Vec<TableCell>, Vec<(i32, bool)>)>,
    uc_skip: usize,
    high_surrogate: Option<u32>,
    mb_buf: Vec<u8>,
    star: bool,
    /// Unescaped bytes >= 0x80 are UTF-8 (some generators write raw UTF-8
    /// instead of \'hh / \uN escapes); decided once from the whole input.
    raw_utf8: bool,
    media_total: usize,
    warnings: Vec<String>,
}

impl<'a> RtfParser<'a> {
    fn new(d: &'a [u8]) -> Self {
        RtfParser {
            d,
            pos: 0,
            st: RState::new(),
            stack: Vec::new(),
            overflow: 0,
            cp: 1252,
            deff: 0,
            fonts: HashMap::new(),
            font_buf: String::new(),
            font_cur: 0,
            font_cp: 0,
            colors: Vec::new(),
            rgb: [0; 3],
            rgb_set: false,
            styles: HashMap::new(),
            style_buf: String::new(),
            style_cur: 0,
            style_outline: None,
            title: String::new(),
            pict: None,
            fields: Vec::new(),
            listtext: String::new(),
            pn_tmp: None,
            ls_kind: HashMap::new(),
            text: String::new(),
            items: Vec::new(),
            blocks: Vec::new(),
            cell_blocks: Vec::new(),
            cur_row: Vec::new(),
            row_cellx: Vec::new(),
            pending_merge: false,
            table_rows: Vec::new(),
            uc_skip: 0,
            high_surrogate: None,
            mb_buf: Vec::new(),
            star: false,
            raw_utf8: d.iter().any(|b| *b >= 0x80) && std::str::from_utf8(d).is_ok(),
            media_total: 0,
            warnings: Vec::new(),
        }
    }

    fn run(mut self) -> WordDoc {
        while self.pos < self.d.len() {
            let c = self.d[self.pos];
            match c {
                b'{' => {
                    self.pos += 1;
                    self.flush_text();
                    self.uc_skip = 0;
                    self.star = false;
                    if self.stack.len() >= 4096 {
                        self.overflow += 1;
                    } else {
                        self.stack.push(self.st.clone());
                    }
                }
                b'}' => {
                    self.pos += 1;
                    self.star = false;
                    self.end_group();
                }
                b'\\' => self.control(),
                b'\r' | b'\n' | 0 => self.pos += 1,
                b'\t' => {
                    self.pos += 1;
                    if self.uc_skip > 0 {
                        self.uc_skip -= 1;
                    } else {
                        self.emit_char('\t');
                    }
                }
                _ => {
                    self.pos += 1;
                    if self.uc_skip > 0 {
                        // A multi-byte UTF-8 fallback character counts once.
                        if !(self.raw_utf8 && (0x80..0xC0).contains(&c)) {
                            self.uc_skip -= 1;
                        }
                        continue;
                    }
                    if self.raw_utf8 && c >= 0x80 {
                        self.raw_utf8_byte(c);
                    } else {
                        self.byte(c);
                    }
                }
            }
        }
        self.flush_text();
        if !self.items.is_empty() {
            let intbl = self.st.intbl;
            self.finish_para(intbl);
        }
        if !self.cur_row.is_empty() || !self.cell_blocks.is_empty() {
            if !self.cell_blocks.is_empty() {
                let blocks = std::mem::take(&mut self.cell_blocks);
                self.cur_row.push(TableCell { blocks, colspan: None });
            }
            let row = std::mem::take(&mut self.cur_row);
            self.table_rows.push((row, self.row_cellx.clone()));
        }
        self.flush_table();
        let title = std::mem::take(&mut self.title);
        let blocks = docx::merge_code_blocks(std::mem::take(&mut self.blocks));
        doc_from(blocks, &self.warnings, Some(title))
    }

    fn cur_cp(&self) -> u32 {
        if self.st.font >= 0 {
            if let Some((_, cp)) = self.fonts.get(&self.st.font) {
                if *cp != 0 {
                    return *cp;
                }
            }
        }
        self.cp
    }

    /// A raw (or \'hh) byte in the current code page.
    fn byte(&mut self, b: u8) {
        if b < 0x80 && self.mb_buf.is_empty() {
            self.emit_char(b as char);
            return;
        }
        let cp = self.cur_cp();
        match cp {
            65001 => {
                self.mb_buf.push(b);
                match std::str::from_utf8(&self.mb_buf) {
                    Ok(s) => {
                        let s = s.to_string();
                        self.mb_buf.clear();
                        for ch in s.chars() {
                            self.emit_char(ch);
                        }
                    }
                    Err(e) if e.error_len().is_some() || self.mb_buf.len() >= 4 => {
                        self.mb_buf.clear();
                        self.emit_char('\u{FFFD}');
                    }
                    Err(_) => {}
                }
            }
            932 | 936 | 949 | 950 => {
                // Double-byte code pages: Word always adds \uN for these, so raw
                // pairs only appear in old files; show a replacement character.
                if self.mb_buf.is_empty() && b >= 0x81 {
                    self.mb_buf.push(b);
                } else {
                    self.mb_buf.clear();
                    self.emit_char('\u{FFFD}');
                }
            }
            _ => {
                self.mb_buf.clear();
                let ch = decode_cp_byte(cp, b);
                self.emit_char(ch);
            }
        }
    }

    fn raw_utf8_byte(&mut self, b: u8) {
        self.mb_buf.push(b);
        match std::str::from_utf8(&self.mb_buf) {
            Ok(s) => {
                let s = s.to_string();
                self.mb_buf.clear();
                for ch in s.chars() {
                    self.emit_char(ch);
                }
            }
            Err(e) if e.error_len().is_some() || self.mb_buf.len() >= 4 => {
                self.mb_buf.clear();
                self.emit_char('\u{FFFD}');
            }
            Err(_) => {}
        }
    }

    fn emit_char(&mut self, ch: char) {
        match self.st.dest {
            Dest::Normal => {
                if !self.st.hidden {
                    self.text.push(ch);
                }
            }
            Dest::FontTbl => {
                if ch == ';' {
                    self.commit_font();
                } else {
                    self.font_buf.push(ch);
                }
            }
            Dest::ColorTbl => {
                if ch == ';' {
                    self.colors.push(if self.rgb_set { Some(format!("#{:02X}{:02X}{:02X}", self.rgb[0], self.rgb[1], self.rgb[2])) } else { None });
                    self.rgb = [0; 3];
                    self.rgb_set = false;
                }
            }
            Dest::StyleSheet => {
                if ch == ';' {
                    self.commit_style();
                } else {
                    self.style_buf.push(ch);
                }
            }
            Dest::Title => self.title.push(ch),
            Dest::FldInst => {
                if let Some(f) = self.fields.last_mut() {
                    f.1.push(ch);
                }
            }
            Dest::ListText => self.listtext.push(ch),
            Dest::Pict => {
                if let Some(v) = ch.to_digit(16) {
                    if let Some(p) = self.pict.as_mut() {
                        if p.data.len() as u64 >= MAX_IMAGE {
                            p.too_big = true;
                            return;
                        }
                        match p.half.take() {
                            Some(h) => p.data.push((h << 4) | v as u8),
                            None => p.half = Some(v as u8),
                        }
                    }
                }
            }
            Dest::Skip | Dest::Info | Dest::Pn => {}
        }
    }

    fn commit_font(&mut self) {
        let name = self.font_buf.trim().trim_end_matches(';').trim().to_string();
        if !name.is_empty() || !self.fonts.contains_key(&self.font_cur) {
            let cp = if self.font_cp == 0 && (name.eq_ignore_ascii_case("symbol") || name.to_ascii_lowercase().starts_with("wingdings")) { CP_SYMBOL } else { self.font_cp };
            self.fonts.insert(self.font_cur, (name, cp));
        }
        self.font_buf.clear();
    }

    fn commit_style(&mut self) {
        let name = self.style_buf.trim().to_string();
        if self.style_cur >= 0 && !name.is_empty() {
            self.styles.insert(self.style_cur, (name, self.style_outline));
        }
        self.style_buf.clear();
        self.style_cur = 0;
        self.style_outline = None;
    }

    fn flush_text(&mut self) {
        if self.text.is_empty() {
            return;
        }
        let t = std::mem::take(&mut self.text);
        let st = &self.st;
        let color = |idx: i32| -> Option<String> { if idx <= 0 { None } else { self.colors.get(idx as usize).cloned().flatten() } };
        let font = if st.font >= 0 && st.font != self.deff {
            self.fonts.get(&st.font).filter(|(n, cp)| !n.is_empty() && *cp != CP_SYMBOL).map(|(n, _)| n.clone())
        } else {
            None
        };
        // Hyperlink text is written with the standard "link look" (underline +
        // link blue); that is implied by the link itself, so don't keep it.
        let mut cf_color = color(st.cf);
        let mut u = st.u;
        if st.link.is_some() && u && matches!(cf_color.as_deref(), Some("#0563C1" | "#0000FF")) {
            cf_color = None;
            u = false;
        }
        let run = Run {
            t,
            b: st.b,
            i: st.i,
            u,
            s: st.s,
            sup: st.sup,
            sub: st.sub && !st.sup,
            color: cf_color,
            size: if st.fs > 0 && st.fs != 24 { Some(st.fs as f64 / 2.0) } else { None },
            font,
            link: st.link.clone(),
            hl: color(st.hl),
        };
        self.items.push(Inline::Run(run));
    }

    fn end_group(&mut self) {
        self.flush_text();
        self.uc_skip = 0;
        if self.overflow > 0 {
            self.overflow -= 1;
            return;
        }
        let Some(parent) = self.stack.pop() else { return };
        let old = std::mem::replace(&mut self.st, parent);
        match old.dest {
            Dest::FontTbl => {
                if !self.font_buf.trim().is_empty() {
                    self.commit_font();
                }
            }
            Dest::StyleSheet => {
                if !self.style_buf.trim().is_empty() {
                    self.commit_style();
                }
            }
            Dest::Pict if self.st.dest != Dest::Pict => self.finish_pict(),
            Dest::Pn if self.st.dest != Dest::Pn => {
                if let Some(pn) = self.pn_tmp.take() {
                    self.st.pn = Some(pn);
                }
            }
            _ => {}
        }
        while let Some((depth, _)) = self.fields.last() {
            if *depth > self.stack.len() {
                self.fields.pop();
            } else {
                break;
            }
        }
    }

    fn control(&mut self) {
        self.pos += 1;
        let Some(&c) = self.d.get(self.pos) else { return };
        if c.is_ascii_alphabetic() {
            let start = self.pos;
            while self.pos < self.d.len() && self.d[self.pos].is_ascii_alphabetic() && self.pos - start < 32 {
                self.pos += 1;
            }
            let word = String::from_utf8_lossy(&self.d[start..self.pos]).into_owned();
            let mut num: Option<i32> = None;
            if self.pos < self.d.len() && (self.d[self.pos] == b'-' || self.d[self.pos].is_ascii_digit()) {
                let ns = self.pos;
                if self.d[self.pos] == b'-' {
                    self.pos += 1;
                }
                while self.pos < self.d.len() && self.d[self.pos].is_ascii_digit() && self.pos - ns < 11 {
                    self.pos += 1;
                }
                num = std::str::from_utf8(&self.d[ns..self.pos]).ok().and_then(|s| s.parse::<i64>().ok()).map(|v| v.clamp(i32::MIN as i64, i32::MAX as i64) as i32);
            }
            if self.pos < self.d.len() && self.d[self.pos] == b' ' {
                self.pos += 1;
            }
            self.flush_text();
            if self.uc_skip > 0 && word != "bin" {
                self.uc_skip -= 1;
                return;
            }
            self.word(&word, num);
        } else {
            self.pos += 1;
            match c {
                b'\'' => {
                    let v = self.d.get(self.pos..self.pos + 2).and_then(|h| std::str::from_utf8(h).ok()).and_then(|h| u8::from_str_radix(h, 16).ok());
                    if let Some(v) = v {
                        self.pos += 2;
                        if self.uc_skip > 0 {
                            self.uc_skip -= 1;
                            return;
                        }
                        self.byte(v);
                    }
                }
                b'*' => self.star = true,
                b'\\' | b'{' | b'}' => {
                    if self.uc_skip > 0 {
                        self.uc_skip -= 1;
                    } else {
                        self.emit_char(c as char);
                    }
                }
                b'~' => self.emit_char('\u{a0}'),
                b'_' => self.emit_char('\u{2011}'),
                b'\r' | b'\n' => {
                    self.flush_text();
                    if self.st.dest == Dest::Normal {
                        self.par();
                    }
                }
                b'\t' => self.emit_char('\t'),
                _ => {}
            }
        }
    }

    fn unicode(&mut self, n: i32) {
        let v = if n < 0 { (n + 65536) as u32 } else { n as u32 };
        let ch = if (0xD800..0xDC00).contains(&v) {
            self.high_surrogate = Some(v);
            None
        } else if (0xDC00..0xE000).contains(&v) {
            self.high_surrogate.take().and_then(|h| char::from_u32(0x10000 + ((h - 0xD800) << 10) + (v - 0xDC00)))
        } else {
            self.high_surrogate = None;
            char::from_u32(v)
        };
        if let Some(ch) = ch {
            if self.st.dest == Dest::Pict {
                return;
            }
            let ch = if self.cur_cp() == CP_SYMBOL && (0xF000..=0xF0FF).contains(&(ch as u32)) { decode_cp_byte(CP_SYMBOL, (ch as u32 - 0xF000) as u8) } else { ch };
            self.emit_char(ch);
        }
        self.uc_skip = self.st.uc;
    }

    fn special_char(&mut self, w: &str) -> bool {
        let ch = match w {
            "tab" => '\t',
            "emdash" => '\u{2014}',
            "endash" => '\u{2013}',
            "emspace" => '\u{2003}',
            "enspace" => '\u{2002}',
            "qmspace" => '\u{2005}',
            "bullet" => '\u{2022}',
            "lquote" => '\u{2018}',
            "rquote" => '\u{2019}',
            "ldblquote" => '\u{201C}',
            "rdblquote" => '\u{201D}',
            "zwj" => '\u{200D}',
            "zwnj" => '\u{200C}',
            "line" | "nestrow" => '\n',
            "nestcell" => '\t',
            _ => return false,
        };
        self.emit_char(ch);
        true
    }

    fn word(&mut self, w: &str, n: Option<i32>) {
        let star = std::mem::take(&mut self.star);
        if w == "bin" {
            let len = n.unwrap_or(0).max(0) as usize;
            let end = (self.pos + len).min(self.d.len());
            if self.st.dest == Dest::Pict {
                if let Some(p) = self.pict.as_mut() {
                    if (p.data.len() + (end - self.pos)) as u64 <= MAX_IMAGE {
                        p.data.extend_from_slice(&self.d[self.pos..end]);
                    } else {
                        p.too_big = true;
                    }
                }
            }
            self.pos = end;
            return;
        }
        if self.st.dest == Dest::Skip {
            return;
        }
        // Destinations.
        match w {
            "fonttbl" => {
                self.st.dest = Dest::FontTbl;
                return;
            }
            "colortbl" => {
                self.st.dest = Dest::ColorTbl;
                self.colors.clear();
                self.rgb_set = false;
                return;
            }
            "stylesheet" => {
                self.st.dest = Dest::StyleSheet;
                return;
            }
            "info" => {
                self.st.dest = Dest::Info;
                return;
            }
            "title" if self.st.dest == Dest::Info => {
                self.st.dest = Dest::Title;
                return;
            }
            "pict" => {
                self.st.dest = Dest::Pict;
                self.pict = Some(Pict { scalex: 100, scaley: 100, ..Default::default() });
                return;
            }
            "fldinst" => {
                self.st.dest = Dest::FldInst;
                return;
            }
            "fldrslt" => {
                self.st.dest = Dest::Normal;
                let instr = self.fields.last().map(|f| f.1.clone()).unwrap_or_default();
                let link = docx::field_link(&instr);
                if link.is_none() && !instr.trim().is_empty() {
                    warn(&mut self.warnings, W_FIELDS);
                }
                if link.is_some() {
                    self.st.link = link;
                }
                return;
            }
            "listtext" | "pntext" => {
                self.st.dest = Dest::ListText;
                self.listtext.clear();
                return;
            }
            "pn" => {
                self.st.dest = Dest::Pn;
                self.pn_tmp = Some((false, 0));
                return;
            }
            "shppict" | "shprslt" | "result" | "upr" => return,
            "header" | "headerl" | "headerr" | "headerf" | "footer" | "footerl" | "footerr" | "footerf" => {
                warn(&mut self.warnings, W_RTF_HEADERS);
                self.st.dest = Dest::Skip;
                return;
            }
            "footnote" => {
                warn(&mut self.warnings, W_RTF_NOTES);
                self.st.dest = Dest::Skip;
                return;
            }
            "annotation" => {
                warn(&mut self.warnings, W_RTF_COMMENTS);
                self.st.dest = Dest::Skip;
                return;
            }
            _ => {}
        }
        if RTF_SKIP_DESTS.contains(&w) || (star && w != "field") {
            if star && w == "shp" {
                return;
            }
            self.st.dest = Dest::Skip;
            return;
        }
        let v = n.unwrap_or(1);
        match self.st.dest {
            Dest::FontTbl => {
                match w {
                    "f" => {
                        if !self.font_buf.trim().is_empty() {
                            self.commit_font();
                        }
                        self.font_cur = v;
                        self.font_cp = 0;
                    }
                    "fcharset" => self.font_cp = charset_cp(v),
                    "cpg" => self.font_cp = v.max(0) as u32,
                    "u" => self.unicode(v),
                    _ => {}
                }
                return;
            }
            Dest::ColorTbl => {
                let c = v.clamp(0, 255) as u8;
                match w {
                    "red" => self.rgb[0] = c,
                    "green" => self.rgb[1] = c,
                    "blue" => self.rgb[2] = c,
                    _ => return,
                }
                self.rgb_set = true;
                return;
            }
            Dest::StyleSheet => {
                match w {
                    "s" => self.style_cur = v,
                    "outlinelevel" => self.style_outline = Some(v.max(0) as u32),
                    "cs" | "ds" | "ts" | "tsrowd" => self.style_cur = -1,
                    "u" => self.unicode(v),
                    _ => {}
                }
                return;
            }
            Dest::Pict => {
                if let Some(p) = self.pict.as_mut() {
                    match w {
                        "pngblip" => p.mime = Some("image/png"),
                        "jpegblip" => p.mime = Some("image/jpeg"),
                        "emfblip" => p.mime = Some("image/x-emf"),
                        "wmetafile" => p.mime = Some("image/x-wmf"),
                        "dibitmap" | "wbitmap" | "macpict" | "pmmetafile" => p.mime = None,
                        "picw" => p.picw = v as i64,
                        "pich" => p.pich = v as i64,
                        "picwgoal" => p.goalw = v as i64,
                        "pichgoal" => p.goalh = v as i64,
                        "picscalex" => p.scalex = v as i64,
                        "picscaley" => p.scaley = v as i64,
                        _ => {}
                    }
                }
                return;
            }
            Dest::Pn => {
                if let Some(pn) = self.pn_tmp.as_mut() {
                    match w {
                        "pnlvlblt" => pn.0 = false,
                        "pnlvlbody" | "pndec" | "pnucrm" | "pnlcrm" | "pnucltr" | "pnlcltr" => pn.0 = true,
                        "pnlvl" => {
                            pn.0 = true;
                            pn.1 = (v - 1).clamp(0, 8) as u32;
                        }
                        _ => {}
                    }
                }
                return;
            }
            Dest::FldInst | Dest::ListText | Dest::Title => {
                if w == "u" {
                    self.unicode(v);
                } else {
                    self.special_char(w);
                }
                return;
            }
            Dest::Info | Dest::Skip => return,
            Dest::Normal => {}
        }
        if self.special_char(w) {
            return;
        }
        let on = n != Some(0);
        match w {
            "ansi" => self.cp = 1252,
            "mac" => self.cp = 10000,
            "pc" => self.cp = 437,
            "pca" => self.cp = 850,
            "ansicpg" => self.cp = v.max(0) as u32,
            "deff" => self.deff = v,
            "uc" => self.st.uc = v.clamp(0, 10) as usize,
            "u" => self.unicode(v),
            "par" => self.par(),
            "page" => {
                self.flush_text();
                self.items.push(Inline::Block(Block::PageBreak));
            }
            "sect" => {
                if !self.items.is_empty() {
                    self.par();
                }
            }
            "cell" => self.cell(),
            "row" => self.row(),
            "field" => self.fields.push((self.stack.len(), String::new())),
            "b" => self.st.b = on,
            "i" => self.st.i = on,
            "ul" | "uld" | "uldash" | "uldashd" | "uldashdd" | "uldb" | "ulhwave" | "ulldash" | "ulth" | "ulthd" | "ulthdash"
            | "ulthdashd" | "ulthdashdd" | "ulthldash" | "ululdbwave" | "ulw" | "ulwave" => self.st.u = on,
            "ulnone" => self.st.u = false,
            "strike" | "striked" => self.st.s = on,
            "super" => {
                self.st.sup = on;
                if on {
                    self.st.sub = false;
                }
            }
            "sub" => {
                self.st.sub = on;
                if on {
                    self.st.sup = false;
                }
            }
            "nosupersub" => {
                self.st.sup = false;
                self.st.sub = false;
            }
            "cf" => self.st.cf = v,
            "cb" | "highlight" | "chcbpat" => self.st.hl = v,
            "fs" => self.st.fs = v,
            "f" => self.st.font = v,
            "plain" => self.st.reset_char(),
            "v" => self.st.hidden = on,
            "pard" => self.st.reset_para(),
            "ql" => self.st.align = None,
            "qc" => self.st.align = Some("center"),
            "qr" => self.st.align = Some("right"),
            "qj" | "qd" => self.st.align = Some("justify"),
            "s" => self.st.style = v,
            "outlinelevel" => self.st.outline = Some(v.max(0) as u32),
            "intbl" => self.st.intbl = true,
            "itap" => self.st.intbl = v > 0,
            "ls" => self.st.ls = Some(v),
            "ilvl" => self.st.ilvl = v.clamp(0, 8) as u32,
            "trowd" => {
                self.row_cellx.clear();
                self.pending_merge = false;
            }
            "clmrg" => self.pending_merge = true,
            "cellx" => {
                let m = std::mem::take(&mut self.pending_merge);
                self.row_cellx.push((v, m));
            }
            _ => {}
        }
    }

    fn finish_pict(&mut self) {
        let Some(p) = self.pict.take() else { return };
        if p.too_big {
            warn(&mut self.warnings, W_IMG_LIMIT);
            return;
        }
        let Some(mime) = p.mime else {
            if !p.data.is_empty() {
                warn(&mut self.warnings, W_RTF_PICT);
            }
            return;
        };
        if p.data.is_empty() {
            return;
        }
        let mime = docx::sniff_image(&p.data).unwrap_or(mime);
        if matches!(mime, "image/x-emf" | "image/x-wmf") {
            warn(&mut self.warnings, W_IMG_FORMAT);
        }
        let metafile = matches!(mime, "image/x-emf" | "image/x-wmf");
        let scale = |v: i64, s: i64| -> i64 { v * if s > 0 { s } else { 100 } / 100 };
        let mut w = if p.goalw > 0 { Some(scale(p.goalw, p.scalex) / 15) } else if p.picw > 0 { Some(scale(if metafile { p.picw * 96 / 2540 } else { p.picw }, p.scalex)) } else { None };
        let mut h = if p.goalh > 0 { Some(scale(p.goalh, p.scaley) / 15) } else if p.pich > 0 { Some(scale(if metafile { p.pich * 96 / 2540 } else { p.pich }, p.scaley)) } else { None };
        if w.is_none() || h.is_none() {
            if let Some((iw, ih)) = docx::image_dims(&p.data) {
                w = Some(iw as i64);
                h = Some(ih as i64);
            }
        }
        let src = data_url(mime, &p.data);
        if self.media_total + src.len() > MAX_MEDIA_TOTAL {
            warn(&mut self.warnings, W_IMG_LIMIT);
            return;
        }
        self.media_total += src.len();
        let dim = |v: Option<i64>| v.filter(|x| *x >= 1).map(|x| x.min(100_000) as u32);
        let img = Block::Img { src, alt: None, w: dim(w), h: dim(h) };
        if self.st.dest == Dest::Normal || self.st.dest == Dest::Pict {
            self.items.push(Inline::Block(img));
        }
    }

    fn para_kind(&mut self) -> PKind {
        let st = &self.st;
        let style = self.styles.get(&st.style);
        let lname = style.map(|s| s.0.to_lowercase()).unwrap_or_default();
        if let Some(o) = st.outline.filter(|o| *o < 6) {
            return PKind::H(o + 1);
        }
        if let Some(o) = style.and_then(|s| s.1).filter(|o| *o < 6) {
            return PKind::H(o + 1);
        }
        if let Some(h) = docx::heading_from_name(&lname) {
            return PKind::H(h);
        }
        if lname == "title" {
            return PKind::H(1);
        }
        let lt = self.listtext.trim().to_string();
        let ordered_from_text = |t: &str| -> bool {
            let t = t.trim_end_matches(['.', ')']);
            !t.is_empty() && (t.chars().any(|c| c.is_ascii_digit()) || (t.len() <= 4 && t.chars().all(|c| c.is_ascii_alphabetic())))
        };
        if let Some(ls) = st.ls.filter(|l| *l > 0) {
            let key = (ls, st.ilvl);
            let ordered = if !lt.is_empty() {
                let o = ordered_from_text(&lt);
                self.ls_kind.insert(key, o);
                o
            } else {
                *self.ls_kind.get(&key).unwrap_or(&false)
            };
            return PKind::List(if ordered { "ol".into() } else { "ul".into() }, st.ilvl.min(8));
        }
        if let Some((ordered, lvl)) = st.pn {
            return PKind::List(if ordered { "ol".into() } else { "ul".into() }, lvl.min(8));
        }
        if !lt.is_empty() {
            return PKind::List(if ordered_from_text(&lt) { "ol".into() } else { "ul".into() }, 0);
        }
        if lname.contains("quote") || lname == "quotations" {
            return PKind::Quote;
        }
        if docx::is_code_style_name(&lname) {
            return PKind::Code;
        }
        PKind::P
    }

    fn finish_para(&mut self, to_cell: bool) {
        self.flush_text();
        let items = std::mem::take(&mut self.items);
        let kind = self.para_kind();
        self.listtext.clear();
        let align = self.st.align.map(String::from);
        let cont = if matches!(kind, PKind::List(..)) { PKind::P } else { kind.clone() };
        let mut items = items;
        strip_style_formatting(&kind, &mut items);
        let mut out = Vec::new();
        split_paragraph(items, &mut out, &mut |runs, first| kind_block(if first { &kind } else { &cont }, runs, align.clone()));
        if to_cell {
            self.cell_blocks.extend(out);
        } else {
            self.flush_table();
            self.blocks.extend(out);
        }
    }

    fn par(&mut self) {
        let intbl = self.st.intbl;
        self.finish_para(intbl);
    }

    fn cell(&mut self) {
        self.flush_text();
        if !self.items.is_empty() || self.cell_blocks.is_empty() {
            self.finish_para(true);
        }
        let blocks = std::mem::take(&mut self.cell_blocks);
        self.cur_row.push(TableCell { blocks, colspan: None });
    }

    fn row(&mut self) {
        self.flush_text();
        if !self.items.is_empty() || !self.cell_blocks.is_empty() {
            self.cell();
        }
        let row = std::mem::take(&mut self.cur_row);
        if !row.is_empty() {
            self.table_rows.push((row, self.row_cellx.clone()));
        }
    }

    fn flush_table(&mut self) {
        if self.table_rows.is_empty() {
            return;
        }
        let rows = std::mem::take(&mut self.table_rows);
        let mut grid: Vec<i32> = rows.iter().flat_map(|(_, d)| d.iter().map(|x| x.0)).collect();
        grid.sort_unstable();
        grid.dedup();
        let out: Vec<Vec<TableCell>> = rows
            .into_iter()
            .map(|(cells, defs)| {
                let mut res: Vec<TableCell> = Vec::new();
                let mut prev = i32::MIN;
                for (k, mut cell) in cells.into_iter().enumerate() {
                    let (span, merged) = match defs.get(k) {
                        Some(&(x, m)) => {
                            let s = grid.iter().filter(|g| **g > prev && **g <= x).count().max(1) as u32;
                            prev = x;
                            (s, m)
                        }
                        None => (1, false),
                    };
                    if merged && !res.is_empty() {
                        let last = res.last_mut().unwrap();
                        last.colspan = Some(last.colspan.unwrap_or(1) + span);
                        if cell.blocks.iter().any(|b| b.para_runs().iter().any(|r| !r.t.trim().is_empty())) {
                            last.blocks.append(&mut cell.blocks);
                        }
                        continue;
                    }
                    cell.colspan = if span > 1 { Some(span) } else { None };
                    res.push(cell);
                }
                res
            })
            .collect();
        self.blocks.push(Block::Table { rows: out });
    }
}

/// RTF repeats a style's character formatting on every run (Word expands
/// heading/quote/code styles into direct formatting). When every run of a
/// heading/quote/code paragraph carries the same style-implied formatting,
/// drop it so the model stays clean (the block type already implies it).
fn strip_style_formatting(kind: &PKind, items: &mut [Inline]) {
    let mut runs: Vec<&mut Run> = items.iter_mut().filter_map(|i| if let Inline::Run(r) = i { Some(r) } else { None }).collect();
    if runs.is_empty() {
        return;
    }
    match kind {
        PKind::H(_) => {
            let size = runs[0].size;
            let uniform = runs.iter().all(|r| r.size == size);
            let all_b = runs.iter().all(|r| r.b);
            for r in runs.iter_mut() {
                if uniform {
                    r.size = None;
                }
                if all_b {
                    r.b = false;
                }
            }
        }
        PKind::Quote => {
            if runs.iter().all(|r| r.i) {
                for r in runs.iter_mut() {
                    r.i = false;
                }
            }
        }
        PKind::Code => {
            for r in runs.iter_mut() {
                if r.font.as_deref().map(is_mono_font).unwrap_or(false) {
                    r.font = None;
                }
                r.size = None;
            }
        }
        _ => {}
    }
}

/// RTF -> word model.
pub fn parse_rtf(bytes: &[u8]) -> Result<WordDoc, String> {
    if !docx::looks_rtf(bytes) {
        if bytes.starts_with(b"PK") {
            return docx::parse_docx(bytes);
        }
        if bytes.starts_with(&[0xD0, 0xCF, 0x11, 0xE0]) {
            return Ok(parse_doc(bytes));
        }
        let mut d = parse_text(bytes, false);
        push_warning(&mut d.meta, W_NOT_RTF);
        return Ok(d);
    }
    Ok(RtfParser::new(bytes).run())
}

// ============================ ODT ============================

fn keep_ws_odt(parent: &str) -> bool {
    matches!(parent, "p" | "h" | "span" | "a" | "ruby-base" | "meta" | "date" | "time" | "page-number" | "title" | "sequence")
}

#[derive(Default, Clone)]
struct OStyle {
    parent: Option<String>,
    display: String,
    props: Vec<(String, String)>,
    outline: Option<u32>,
}

fn odt_len_pt(v: &str) -> Option<f64> {
    docx::css_len_px(v).map(|px| (px as f64 * 72.0 / 96.0 * 2.0).round() / 2.0).or_else(|| {
        let n: f64 = v.trim().trim_end_matches("pt").trim().parse().ok()?;
        Some(n)
    })
}

struct OdtReader<'z> {
    zip: ZipArchive<Cursor<&'z [u8]>>,
    names: Vec<String>,
    styles: HashMap<String, OStyle>,
    list_styles: HashMap<String, Vec<bool>>,
    warnings: Vec<String>,
    notes: Vec<(String, Vec<Run>)>,
    note_no: u32,
    media_total: usize,
}

impl<'z> OdtReader<'z> {
    fn warn(&mut self, m: &str) {
        warn(&mut self.warnings, m);
    }

    fn load_styles(&mut self, root: &Node) {
        for section in root.elems().filter(|e| matches!(e.local(), "styles" | "automatic-styles")) {
            for s in section.elems() {
                match s.local() {
                    "style" | "default-style" => {
                        let Some(name) = s.attr("name") else { continue };
                        let mut props = Vec::new();
                        for p in s.elems().filter(|e| matches!(e.local(), "text-properties" | "paragraph-properties")) {
                            for (k, v) in &p.attrs {
                                props.push((xml::local_name(k).to_string(), v.clone()));
                            }
                        }
                        self.styles.insert(
                            name.to_string(),
                            OStyle {
                                parent: s.attr("parent-style-name").map(String::from),
                                display: s.attr("display-name").unwrap_or(name).to_string(),
                                props,
                                outline: s.attr("default-outline-level").and_then(|v| v.parse().ok()),
                            },
                        );
                    }
                    "list-style" => {
                        let Some(name) = s.attr("name") else { continue };
                        let mut levels = vec![false; 10];
                        for l in s.elems() {
                            let lvl = l.attr("level").and_then(|v| v.parse::<usize>().ok()).unwrap_or(1).clamp(1, 10) - 1;
                            levels[lvl] = l.local() == "list-level-style-number" && !l.attr("num-format").unwrap_or("1").is_empty();
                        }
                        self.list_styles.insert(name.to_string(), levels);
                    }
                    _ => {}
                }
            }
        }
        if let Some(master) = root.child("master-styles") {
            let mut hf = Vec::new();
            master.find_all("header", &mut hf);
            master.find_all("footer", &mut hf);
            if hf.iter().any(|h| !h.text().trim().is_empty()) {
                self.warn(W_ODT_HEADERS);
            }
        }
    }

    /// Style chain from root ancestor to `name`.
    fn chain(&self, name: Option<&str>) -> Vec<&OStyle> {
        let mut v = Vec::new();
        let mut cur = name.map(String::from);
        while let Some(n) = cur {
            if v.len() > 20 {
                break;
            }
            match self.styles.get(&n) {
                Some(s) => {
                    v.push(s);
                    cur = s.parent.clone();
                }
                None => break,
            }
        }
        v.reverse();
        v
    }

    fn apply_text_props(&self, name: Option<&str>, run: &mut Run) {
        for s in self.chain(name) {
            for (k, v) in &s.props {
                let v = v.trim();
                match k.as_str() {
                    "font-weight" => run.b = v == "bold" || v == "bolder" || v.parse::<u32>().map(|w| w >= 600).unwrap_or(false),
                    "font-style" => run.i = v == "italic" || v == "oblique",
                    "text-underline-style" => run.u = v != "none",
                    "text-line-through-style" => run.s = v != "none",
                    "text-position" => {
                        let first = v.split_whitespace().next().unwrap_or("");
                        let pct = first.trim_end_matches('%').parse::<f64>().ok();
                        run.sup = first == "super" || pct.map(|p| p > 0.0).unwrap_or(false);
                        run.sub = first == "sub" || pct.map(|p| p < 0.0).unwrap_or(false);
                    }
                    "color" => run.color = docx::norm_hex(v).map(|h| format!("#{h}")),
                    "font-size" => {
                        if !v.ends_with('%') {
                            run.size = odt_len_pt(v).filter(|x| *x > 0.0);
                        }
                    }
                    "font-name" | "font-family" => {
                        let f = v.trim_matches(|c| c == '\'' || c == '"').to_string();
                        if !f.is_empty() {
                            run.font = Some(f);
                        }
                    }
                    "background-color" => run.hl = if v == "transparent" { None } else { docx::norm_hex(v).map(|h| format!("#{h}")) },
                    _ => {}
                }
            }
        }
    }

    fn para_info(&self, name: Option<&str>) -> (Option<String>, PKind, bool, bool) {
        let mut align = None;
        let mut kind = PKind::P;
        let mut brk_before = false;
        let mut brk_after = false;
        for s in self.chain(name) {
            let d = s.display.to_lowercase().replace("_20_", " ");
            if let Some(h) = docx::heading_from_name(&d) {
                kind = PKind::H(h);
            } else if d == "title" {
                kind = PKind::H(1);
            } else if d == "subtitle" {
                kind = PKind::P;
            } else if d.contains("quotation") || d.contains("quote") {
                kind = PKind::Quote;
            } else if docx::is_code_style_name(&d) {
                kind = PKind::Code;
            }
            if let Some(o) = s.outline.filter(|o| *o >= 1 && *o <= 6) {
                kind = PKind::H(o);
            }
            for (k, v) in &s.props {
                match k.as_str() {
                    "text-align" => {
                        align = match v.as_str() {
                            "center" => Some("center".to_string()),
                            "end" | "right" => Some("right".to_string()),
                            "justify" => Some("justify".to_string()),
                            _ => None,
                        }
                    }
                    "break-before" => brk_before = v == "page",
                    "break-after" => brk_after = v == "page",
                    _ => {}
                }
            }
        }
        (align, kind, brk_before, brk_after)
    }

    fn walk(&mut self, n: &Node, out: &mut Vec<Block>) {
        for c in n.elems() {
            match c.local() {
                "p" | "h" => self.paragraph(c, out, None),
                "list" => self.list(c, out, None, 0),
                "numbered-paragraph" => {
                    let lvl = c.attr("level").and_then(|v| v.parse::<u32>().ok()).unwrap_or(1).saturating_sub(1);
                    let style = c.attr("style-name").map(String::from);
                    for p in c.elems().filter(|e| e.is("p") || e.is("h")) {
                        self.paragraph(p, out, Some((style.clone(), lvl)));
                    }
                }
                "table" => {
                    let t = self.table(c);
                    out.push(t);
                }
                "section" | "index-body" | "table-of-content" | "alphabetical-index" | "illustration-index" | "user-index"
                | "bibliography" | "object-index" | "table-index" | "index-title" | "deletion" => {
                    if c.is("deletion") {
                        continue;
                    }
                    self.walk(c, out)
                }
                "tracked-changes" => {
                    if c.elems().next().is_some() {
                        self.warn(W_ODT_TRACKED);
                    }
                }
                _ => {}
            }
        }
    }

    fn list(&mut self, l: &Node, out: &mut Vec<Block>, style: Option<String>, level: u32) {
        let style = l.attr("style-name").map(String::from).or(style);
        for item in l.elems() {
            let header = item.is("list-header");
            if !item.is("list-item") && !header {
                continue;
            }
            for x in item.elems() {
                match x.local() {
                    "p" | "h" => self.paragraph(x, out, if header { None } else { Some((style.clone(), level)) }),
                    "list" => self.list(x, out, style.clone(), level + 1),
                    "table" => {
                        let t = self.table(x);
                        out.push(t);
                    }
                    _ => {}
                }
            }
        }
    }

    fn paragraph(&mut self, p: &Node, out: &mut Vec<Block>, list: Option<(Option<String>, u32)>) {
        let sname = p.attr("style-name").map(String::from);
        let (align, mut kind, before, after) = self.para_info(sname.as_deref());
        if p.is("h") {
            let lvl = p.attr("outline-level").and_then(|v| v.parse::<u32>().ok()).unwrap_or(1).clamp(1, 6);
            kind = PKind::H(lvl);
        } else if let Some((ls, lvl)) = &list {
            let ordered = ls.as_ref().and_then(|s| self.list_styles.get(s)).and_then(|v| v.get(*lvl as usize).copied()).unwrap_or(false);
            kind = PKind::List(if ordered { "ol".into() } else { "ul".into() }, (*lvl).min(8));
        }
        let mut base = Run::default();
        self.apply_text_props(sname.as_deref(), &mut base);
        let mut items = Vec::new();
        let mut deferred = Vec::new();
        let mut prev_space = true;
        self.inline(p, &base, None, &mut items, &mut deferred, &mut prev_space);
        // Trailing collapsible whitespace is not significant in ODF.
        if let Some(Inline::Run(r)) = items.last_mut() {
            if prev_space && r.t.ends_with(' ') {
                r.t.pop();
            }
        }
        if before {
            out.push(Block::PageBreak);
        }
        let cont = if matches!(kind, PKind::List(..)) { PKind::P } else { kind.clone() };
        split_paragraph(items, out, &mut |runs, first| kind_block(if first { &kind } else { &cont }, runs, align.clone()));
        if after {
            out.push(Block::PageBreak);
        }
        out.extend(deferred);
    }

    fn push_text(items: &mut Vec<Inline>, props: &Run, link: Option<&str>, t: String) {
        if t.is_empty() {
            return;
        }
        let mut r = props.clone();
        r.t = t;
        r.link = link.map(String::from);
        items.push(Inline::Run(r));
    }

    fn inline(&mut self, n: &Node, props: &Run, link: Option<&str>, items: &mut Vec<Inline>, deferred: &mut Vec<Block>, prev_space: &mut bool) {
        for child in &n.children {
            match child {
                Child::Text(t) => {
                    let mut s = String::with_capacity(t.len());
                    for ch in t.chars() {
                        if matches!(ch, ' ' | '\t' | '\n' | '\r') {
                            if !*prev_space {
                                s.push(' ');
                                *prev_space = true;
                            }
                        } else {
                            s.push(ch);
                            *prev_space = false;
                        }
                    }
                    Self::push_text(items, props, link, s);
                }
                Child::El(e) => match e.local() {
                    "span" => {
                        let mut p = props.clone();
                        self.apply_text_props(e.attr("style-name"), &mut p);
                        self.inline(e, &p, link, items, deferred, prev_space);
                    }
                    "a" => {
                        let href = e.attr("href").and_then(safe_link);
                        self.inline(e, props, href.as_deref().or(link), items, deferred, prev_space);
                    }
                    "s" => {
                        let c = e.attr("c").and_then(|v| v.parse::<usize>().ok()).unwrap_or(1).clamp(1, 1000);
                        Self::push_text(items, props, link, " ".repeat(c));
                        *prev_space = false;
                    }
                    "tab" => {
                        Self::push_text(items, props, link, "\t".into());
                        *prev_space = false;
                    }
                    "line-break" => {
                        Self::push_text(items, props, link, "\n".into());
                        *prev_space = true;
                    }
                    "note" => {
                        let cit = e.child("note-citation").map(|c| c.text().trim().to_string()).filter(|c| !c.is_empty());
                        self.note_no += 1;
                        let label = cit.unwrap_or_else(|| self.note_no.to_string());
                        let mut r = props.clone();
                        r.t = label.clone();
                        r.sup = true;
                        r.sub = false;
                        r.link = None;
                        items.push(Inline::Run(r));
                        if let Some(body) = e.child("note-body") {
                            let mut blocks = Vec::new();
                            self.walk(body, &mut blocks);
                            let mut runs = Vec::new();
                            for b in &blocks {
                                if !runs.is_empty() {
                                    runs.push(Run::text(" "));
                                }
                                runs.extend(b.para_runs().iter().cloned());
                            }
                            self.notes.push((label, runs));
                        }
                        self.warn(W_ODT_NOTES);
                    }
                    "annotation" => self.warn(W_ODT_COMMENTS),
                    "frame" => {
                        self.frame(e, items, deferred);
                        // An anchored frame occupies a character position.
                        *prev_space = false;
                    }
                    "custom-shape" | "rect" | "ellipse" | "line" | "polygon" | "polyline" | "path" | "connector" | "g" | "control" | "circle" => {
                        let mut paras = Vec::new();
                        e.find_all("p", &mut paras);
                        if paras.iter().any(|p| !p.text().trim().is_empty()) {
                            for p in paras {
                                self.paragraph(p, deferred, None);
                            }
                            self.warn(W_ODT_TEXTBOX);
                        } else {
                            self.warn(W_ODT_SHAPES);
                        }
                    }
                    "ruby" => {
                        if let Some(b) = e.child("ruby-base") {
                            self.inline(b, props, link, items, deferred, prev_space);
                        }
                    }
                    "soft-page-break" | "bookmark" | "bookmark-start" | "bookmark-end" | "reference-mark" | "reference-mark-start"
                    | "reference-mark-end" | "toc-mark" | "toc-mark-start" | "toc-mark-end" | "alphabetical-index-mark"
                    | "alphabetical-index-mark-start" | "alphabetical-index-mark-end" | "user-index-mark" | "change"
                    | "change-start" | "change-end" | "annotation-end" | "hidden-text" | "hidden-paragraph" | "tracked-changes" => {}
                    _ => self.inline(e, props, link, items, deferred, prev_space),
                },
            }
        }
    }

    fn entry_name(&self, href: &str) -> Option<String> {
        let h = href.trim().trim_start_matches("./");
        let dec = String::from_utf8_lossy(&docx::percent_decode(h)).into_owned();
        self.names.iter().find(|n| n.as_str() == h || n.as_str() == dec).cloned()
    }

    fn frame(&mut self, f: &Node, items: &mut Vec<Inline>, deferred: &mut Vec<Block>) {
        let w = f.attr("width").and_then(docx::css_len_px);
        let h = f.attr("height").and_then(docx::css_len_px);
        let alt = f
            .child("desc")
            .or_else(|| f.child("title"))
            .map(|d| d.text().trim().to_string())
            .filter(|t| !t.is_empty());
        let mut done = false;
        for img in f.elems().filter(|e| e.is("image")) {
            let bytes = if let Some(bin) = img.child("binary-data") {
                docx::base64_decode(&bin.text())
            } else if let Some(href) = img.attr("href") {
                if href.contains("://") {
                    None
                } else {
                    self.entry_name(href).and_then(|n| docx::read_zip_entry(&mut self.zip, &n, MAX_IMAGE).ok())
                }
            } else {
                None
            };
            let Some(bytes) = bytes.filter(|b| !b.is_empty()) else { continue };
            let Some(mime) = docx::sniff_image(&bytes) else { continue };
            if matches!(mime, "image/x-emf" | "image/x-wmf" | "image/tiff") {
                // Prefer a browser-friendly sibling if there is one.
                if f.elems().filter(|e| e.is("image")).count() > 1 && !done {
                    warn(&mut self.warnings, W_IMG_FORMAT);
                }
            }
            let src = data_url(mime, &bytes);
            if self.media_total + src.len() > MAX_MEDIA_TOTAL {
                self.warn(W_IMG_LIMIT);
                break;
            }
            self.media_total += src.len();
            let (w, h) = match (w, h) {
                (Some(w), Some(h)) => (Some(w), Some(h)),
                _ => docx::image_dims(&bytes).map(|(a, b)| (Some(a), Some(b))).unwrap_or((w, h)),
            };
            items.push(Inline::Block(Block::Img { src, alt: alt.clone(), w, h }));
            done = true;
            break;
        }
        if let Some(tb) = f.child("text-box") {
            let mut blocks = Vec::new();
            self.walk(tb, &mut blocks);
            if !blocks.is_empty() {
                deferred.extend(blocks);
                self.warn(W_ODT_TEXTBOX);
                done = true;
            }
        }
        if !done {
            self.warn(W_ODT_SHAPES);
        }
    }

    fn table(&mut self, t: &Node) -> Block {
        let mut rows_nodes = Vec::new();
        fn collect<'n>(n: &'n Node, out: &mut Vec<&'n Node>) {
            for c in n.elems() {
                match c.local() {
                    "table-row" => out.push(c),
                    "table-header-rows" | "table-rows" | "table-row-group" => collect(c, out),
                    _ => {}
                }
            }
        }
        collect(t, &mut rows_nodes);
        let mut rows = Vec::new();
        for r in rows_nodes {
            let rep = r.attr("number-rows-repeated").and_then(|v| v.parse::<usize>().ok()).unwrap_or(1).clamp(1, 100);
            let mut cells = Vec::new();
            let mut skip = 0u32;
            for c in r.elems() {
                match c.local() {
                    "table-cell" => {
                        let span = c.attr("number-columns-spanned").and_then(|v| v.parse::<u32>().ok()).unwrap_or(1).clamp(1, 63);
                        if c.attr("number-rows-spanned").and_then(|v| v.parse::<u32>().ok()).unwrap_or(1) > 1 {
                            self.warn(W_ODT_VMERGE);
                        }
                        let crep = c.attr("number-columns-repeated").and_then(|v| v.parse::<usize>().ok()).unwrap_or(1).clamp(1, 64);
                        let mut blocks = Vec::new();
                        self.walk(c, &mut blocks);
                        for _ in 0..crep {
                            cells.push(TableCell { blocks: blocks.clone(), colspan: if span > 1 { Some(span) } else { None } });
                        }
                        skip = span - 1;
                    }
                    "covered-table-cell" => {
                        if skip > 0 {
                            skip -= 1;
                        } else {
                            cells.push(TableCell { blocks: Vec::new(), colspan: None });
                        }
                    }
                    _ => {}
                }
            }
            for _ in 0..rep {
                rows.push(cells.clone());
            }
        }
        // Drop trailing empty rows that spreadsheet-like tables often repeat.
        while rows.len() > 1 && rows.last().map(|r| r.iter().all(|c| c.blocks.iter().all(|b| b.para_runs().is_empty()))).unwrap_or(false) {
            rows.pop();
        }
        Block::Table { rows }
    }
}

/// .odt -> word model.
pub fn parse_odt(bytes: &[u8]) -> Result<WordDoc, String> {
    if bytes.starts_with(&[0xD0, 0xCF, 0x11, 0xE0]) {
        return Ok(parse_doc(bytes));
    }
    if docx::looks_rtf(bytes) {
        return parse_rtf(bytes);
    }
    let zip = ZipArchive::new(Cursor::new(bytes)).map_err(|e| format!("This .odt file is damaged or not an OpenDocument file (zip error: {e})."))?;
    let names: Vec<String> = zip.file_names().filter_map(|n| n.ok().map(|n| n.to_string())).collect();
    if !names.iter().any(|n| n == "content.xml") {
        if names.iter().any(|n| n == "word/document.xml") {
            return docx::parse_docx(bytes);
        }
        return Err("This file is not an OpenDocument text file (content.xml is missing).".into());
    }
    let mut r = OdtReader { zip, names, styles: HashMap::new(), list_styles: HashMap::new(), warnings: Vec::new(), notes: Vec::new(), note_no: 0, media_total: 0 };
    let read = |r: &mut OdtReader, name: &str| -> Result<Option<Node>, String> {
        if !r.names.iter().any(|n| n == name) {
            return Ok(None);
        }
        let b = docx::read_zip_entry(&mut r.zip, name, MAX_XML_PART)?;
        xml::parse(&docx::xml_bytes_to_string(&b), keep_ws_odt).map(Some)
    };
    let content = read(&mut r, "content.xml")?.ok_or("content.xml is missing")?;
    if let Ok(Some(styles)) = read(&mut r, "styles.xml") {
        r.load_styles(&styles);
    }
    r.load_styles(&content);
    let title = read(&mut r, "meta.xml").ok().flatten().and_then(|m| m.find("title").map(|t| t.text()));
    let mut blocks = Vec::new();
    if let Some(text) = content.child("body").and_then(|b| b.child("text")) {
        r.walk(text, &mut blocks);
    }
    let mut blocks = docx::merge_code_blocks(blocks);
    if !r.notes.is_empty() {
        blocks.push(Block::P { runs: vec![Run::text("Notes")], align: None });
        for (label, runs) in std::mem::take(&mut r.notes) {
            let mut rs = vec![Run { t: label, sup: true, ..Default::default() }, Run::text(" ")];
            rs.extend(runs);
            blocks.push(Block::P { runs: merge_runs(rs), align: None });
        }
    }
    Ok(doc_from(blocks, &r.warnings, title))
}

// ============================ legacy .doc (OLE2) ============================

/// Minimal read-only OLE2 / Compound File Binary reader.
mod cfb {
    pub struct Cfb<'a> {
        data: &'a [u8],
        ss: usize,
        mss: usize,
        cutoff: u64,
        fat: Vec<u32>,
        minifat: Vec<u32>,
        mini_stream: Vec<u8>,
        pub entries: Vec<Entry>,
    }

    pub struct Entry {
        pub name: String,
        pub kind: u8,
        pub start: u32,
        pub size: u64,
    }

    fn le16(b: &[u8], o: usize) -> Option<u16> {
        b.get(o..o + 2).map(|s| u16::from_le_bytes([s[0], s[1]]))
    }
    fn le32(b: &[u8], o: usize) -> Option<u32> {
        b.get(o..o + 4).map(|s| u32::from_le_bytes([s[0], s[1], s[2], s[3]]))
    }

    const MAXREG: u32 = 0xFFFF_FFFA;

    impl<'a> Cfb<'a> {
        pub fn open(data: &'a [u8]) -> Option<Cfb<'a>> {
            if data.len() < 512 || data[..8] != [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1] {
                return None;
            }
            let shift = le16(data, 0x1E)?;
            if shift != 9 && shift != 12 {
                return None;
            }
            let ss = 1usize << shift;
            let mss = 1usize << le16(data, 0x20)?.min(12);
            let num_fat = le32(data, 0x2C)? as usize;
            let dir_start = le32(data, 0x30)?;
            let cutoff = le32(data, 0x38)? as u64;
            let minifat_start = le32(data, 0x3C)?;
            let difat_start = le32(data, 0x44)?;
            let max_sectors = data.len() / ss + 1;
            let mut cfb = Cfb { data, ss, mss, cutoff, fat: Vec::new(), minifat: Vec::new(), mini_stream: Vec::new(), entries: Vec::new() };
            let mut fat_secs: Vec<u32> = (0..109).filter_map(|i| le32(data, 0x4C + i * 4)).filter(|s| *s < MAXREG).collect();
            let mut next = difat_start;
            let mut guard = 0;
            while next < MAXREG && guard < max_sectors {
                let sec = cfb.sector(next)?;
                for i in 0..(ss / 4 - 1) {
                    if let Some(s) = le32(sec, i * 4).filter(|s| *s < MAXREG) {
                        fat_secs.push(s);
                    }
                }
                next = le32(sec, ss - 4)?;
                guard += 1;
            }
            fat_secs.truncate(num_fat.max(1).min(max_sectors));
            for s in fat_secs {
                let sec = cfb.sector(s)?;
                for i in 0..ss / 4 {
                    cfb.fat.push(le32(sec, i * 4)?);
                }
            }
            let dir = cfb.chain(dir_start, None);
            for e in dir.chunks_exact(128) {
                let name_len = (le16(e, 64)? as usize).min(64);
                let units: Vec<u16> = (0..name_len.saturating_sub(2) / 2).filter_map(|i| le16(e, i * 2)).collect();
                let size = if ss == 512 { le32(e, 120)? as u64 } else { le32(e, 120)? as u64 | ((le32(e, 124)? as u64) << 32) };
                cfb.entries.push(Entry { name: String::from_utf16_lossy(&units), kind: e[66], start: le32(e, 116)?, size });
            }
            let mf = cfb.chain(minifat_start, None);
            cfb.minifat = mf.chunks_exact(4).map(|c| u32::from_le_bytes([c[0], c[1], c[2], c[3]])).collect();
            if let Some(root) = cfb.entries.iter().find(|e| e.kind == 5) {
                let (start, size) = (root.start, root.size);
                cfb.mini_stream = cfb.chain(start, Some(size));
            }
            Some(cfb)
        }

        fn sector(&self, n: u32) -> Option<&'a [u8]> {
            let off = (n as usize).checked_add(1)?.checked_mul(self.ss)?;
            self.data.get(off..off + self.ss)
        }

        fn chain(&self, start: u32, size: Option<u64>) -> Vec<u8> {
            let mut out = Vec::new();
            let mut s = start;
            let mut guard = 0usize;
            while s < MAXREG && guard <= self.fat.len() {
                let Some(sec) = self.sector(s) else { break };
                out.extend_from_slice(sec);
                if let Some(sz) = size {
                    if out.len() as u64 >= sz {
                        break;
                    }
                }
                s = *self.fat.get(s as usize).unwrap_or(&0xFFFF_FFFE);
                guard += 1;
            }
            if let Some(sz) = size {
                out.truncate(sz as usize);
            }
            out
        }

        fn mini_chain(&self, start: u32, size: u64) -> Vec<u8> {
            let mut out = Vec::new();
            let mut s = start;
            let mut guard = 0usize;
            while s < MAXREG && guard <= self.minifat.len() && (out.len() as u64) < size {
                let off = s as usize * self.mss;
                let Some(chunk) = self.mini_stream.get(off..off + self.mss) else { break };
                out.extend_from_slice(chunk);
                s = *self.minifat.get(s as usize).unwrap_or(&0xFFFF_FFFE);
                guard += 1;
            }
            out.truncate(size as usize);
            out
        }

        pub fn has(&self, name: &str) -> bool {
            self.entries.iter().any(|e| e.kind == 2 && e.name.eq_ignore_ascii_case(name))
        }

        pub fn stream(&self, name: &str) -> Option<Vec<u8>> {
            let e = self.entries.iter().find(|e| e.kind == 2 && e.name.eq_ignore_ascii_case(name))?;
            Some(if e.size < self.cutoff { self.mini_chain(e.start, e.size) } else { self.chain(e.start, Some(e.size)) })
        }
    }
}

/// True for OLE2 files that are encrypted Office documents (OOXML packages
/// wrapped in EncryptedPackage, or .doc files with the fEncrypted flag).
pub fn ole_is_encrypted(bytes: &[u8]) -> bool {
    let Some(c) = cfb::Cfb::open(bytes) else { return false };
    if c.has("EncryptedPackage") || c.has("EncryptionInfo") {
        return true;
    }
    c.stream("WordDocument").and_then(|wd| le16(&wd, 0x0A)).map(|f| f & 0x0100 != 0).unwrap_or(false)
}

fn le16(b: &[u8], o: usize) -> Option<u16> {
    b.get(o..o + 2).map(|s| u16::from_le_bytes([s[0], s[1]]))
}
fn le32(b: &[u8], o: usize) -> Option<u32> {
    b.get(o..o + 4).map(|s| u32::from_le_bytes([s[0], s[1], s[2], s[3]]))
}

#[derive(Clone, Copy, Default, Debug)]
struct PapInfo {
    in_table: bool,
    ttp: bool,
    istd: u16,
    ilfo: u16,
    ilvl: u8,
}

/// Paragraph properties we care about from a PAPX grpprl.
fn parse_grpprl(g: &[u8], info: &mut PapInfo) {
    let mut i = 0usize;
    while i + 2 <= g.len() {
        let op = u16::from_le_bytes([g[i], g[i + 1]]);
        i += 2;
        let size = match op >> 13 {
            0 | 1 => 1,
            2 | 4 | 5 => 2,
            3 => 4,
            7 => 3,
            _ => {
                if op == 0xD608 || op == 0xD606 {
                    let Some(cb) = le16(g, i) else { return };
                    i += 2;
                    (cb as usize).saturating_sub(1)
                } else {
                    let Some(&cb) = g.get(i) else { return };
                    i += 1;
                    if op == 0xC615 && cb == 255 {
                        return;
                    }
                    cb as usize
                }
            }
        };
        let Some(arg) = g.get(i..i + size) else { return };
        match op {
            0x2416 | 0x244B => info.in_table |= arg[0] != 0,
            0x2417 | 0x244C => info.ttp |= arg[0] != 0,
            0x6649 => info.in_table |= le32(arg, 0).unwrap_or(0) > 0,
            0x460B => info.ilfo = le16(arg, 0).unwrap_or(0),
            0x260A => info.ilvl = arg[0].min(8),
            _ => {}
        }
        i += size;
    }
}

/// (fc_start, fc_end, props) from the PAPX FKPs, sorted by fc.
fn doc_papx(wd: &[u8], table: &[u8]) -> Vec<(u32, u32, PapInfo)> {
    let mut out = Vec::new();
    let (Some(fc), Some(lcb)) = (le32(wd, 0x102), le32(wd, 0x106)) else { return out };
    let Some(plc) = table.get(fc as usize..(fc as usize).saturating_add(lcb as usize)) else { return out };
    if plc.len() < 12 {
        return out;
    }
    let n = (plc.len() - 4) / 8;
    for k in 0..n {
        let Some(pn) = le32(plc, (n + 1) * 4 + k * 4) else { break };
        let off = (pn & 0x3F_FFFF) as usize * 512;
        let Some(page) = wd.get(off..off + 512) else { continue };
        let crun = page[511] as usize;
        for i in 0..crun {
            let (Some(fc0), Some(fc1)) = (le32(page, i * 4), le32(page, (i + 1) * 4)) else { break };
            let bx = (crun + 1) * 4 + i * 13;
            let mut info = PapInfo::default();
            if let Some(&b) = page.get(bx) {
                let b_off = b as usize * 2;
                if b_off > 0 && b_off < 511 {
                    let cb = page[b_off];
                    let (start, len) = if cb != 0 { (b_off + 1, 2 * cb as usize - 1) } else { (b_off + 2, 2 * *page.get(b_off + 1).unwrap_or(&0) as usize) };
                    let end = (start + len).min(511);
                    if start + 2 <= end {
                        info.istd = le16(page, start).unwrap_or(0);
                        parse_grpprl(&page[start + 2..end], &mut info);
                    }
                }
            }
            out.push((fc0, fc1, info));
        }
    }
    out.sort_by_key(|x| x.0);
    out
}

/// ilfo (1-based) -> per-level "ordered" flags, from PlfLst/PlfLfo.
fn doc_lists(wd: &[u8], table: &[u8]) -> HashMap<u16, Vec<bool>> {
    let mut out = HashMap::new();
    let get = |fco: usize| -> Option<&[u8]> {
        let fc = le32(wd, fco)? as usize;
        let lcb = le32(wd, fco + 4)? as usize;
        if lcb == 0 {
            return None;
        }
        table.get(fc..fc.checked_add(lcb)?)
    };
    let (Some(lst), Some(lfo)) = (get(0x2E2), get(0x2EA)) else { return out };
    let Some(clst) = le16(lst, 0) else { return out };
    let clst = clst as usize;
    let mut lists: Vec<(i32, bool)> = Vec::new();
    for k in 0..clst {
        let o = 2 + k * 28;
        let (Some(lsid), Some(&flags)) = (le32(lst, o), lst.get(o + 26)) else { return out };
        lists.push((lsid as i32, flags & 1 != 0));
    }
    let mut pos = 2 + clst * 28;
    let mut levels: HashMap<i32, Vec<bool>> = HashMap::new();
    for (lsid, simple) in &lists {
        let n = if *simple { 1 } else { 9 };
        let mut v = Vec::new();
        for _ in 0..n {
            let Some(lvlf) = lst.get(pos..pos + 28) else { return out };
            let nfc = lvlf[4];
            let cb_chpx = lvlf[24] as usize;
            let cb_papx = lvlf[25] as usize;
            pos += 28 + cb_papx + cb_chpx;
            let Some(cch) = le16(lst, pos) else { return out };
            pos += 2 + cch as usize * 2;
            v.push(nfc != 23 && nfc != 255);
        }
        levels.insert(*lsid, v);
    }
    let Some(count) = le32(lfo, 0) else { return out };
    for k in 0..(count as usize).min(10_000) {
        let Some(lsid) = le32(lfo, 4 + k * 16) else { break };
        if let Some(v) = levels.get(&(lsid as i32)) {
            out.insert((k + 1) as u16, v.clone());
        }
    }
    out
}

/// Real .doc text: piece table + paragraph properties. None if the file is
/// not a Word 97+ binary we understand.
fn parse_word97(bytes: &[u8]) -> Option<(Vec<Block>, bool)> {
    let c = cfb::Cfb::open(bytes)?;
    let wd = c.stream("WordDocument")?;
    if le16(&wd, 0)? != 0xA5EC {
        return None;
    }
    let nfib = le16(&wd, 2)?;
    let flags = le16(&wd, 0x0A)?;
    if flags & 0x0100 != 0 {
        return Some((Vec::new(), true));
    }
    if nfib < 0x00C0 {
        return None; // Word 6/95
    }
    let table = c.stream(if flags & 0x0200 != 0 { "1Table" } else { "0Table" })?;
    let ccp_text = le32(&wd, 0x4C)?;
    let fc_clx = le32(&wd, 0x1A2)? as usize;
    let lcb_clx = le32(&wd, 0x1A6)? as usize;
    let clx = table.get(fc_clx..fc_clx.checked_add(lcb_clx)?)?;
    let mut i = 0;
    while i < clx.len() && clx[i] == 1 {
        i += 3 + le16(clx, i + 1)? as usize;
    }
    if clx.get(i) != Some(&2) {
        return None;
    }
    let lcb = le32(clx, i + 1)? as usize;
    let plc = clx.get(i + 5..i + 5 + lcb)?;
    if lcb < 16 {
        return None;
    }
    let n = (lcb - 4) / 12;
    // (char, fc) stream of the main document text.
    let mut chars: Vec<(char, u32)> = Vec::new();
    for k in 0..n {
        let cp0 = le32(plc, k * 4)?;
        let cp1 = le32(plc, (k + 1) * 4)?.min(ccp_text);
        if cp0 >= ccp_text || cp1 <= cp0 {
            continue;
        }
        let pcd = plc.get((n + 1) * 4 + k * 8..(n + 1) * 4 + k * 8 + 8)?;
        let raw = le32(pcd, 2)?;
        let compressed = raw & 0x4000_0000 != 0;
        let count = (cp1 - cp0) as usize;
        if compressed {
            let fc = ((raw & 0x3FFF_FFFF) / 2) as usize;
            let seg = wd.get(fc..fc + count)?;
            for (j, b) in seg.iter().enumerate() {
                chars.push((decode_cp_byte(1252, *b), (fc + j) as u32));
            }
        } else {
            let fc = (raw & 0x3FFF_FFFF) as usize;
            let seg = wd.get(fc..fc + count * 2)?;
            let units: Vec<u16> = seg.chunks_exact(2).map(|c| u16::from_le_bytes([c[0], c[1]])).collect();
            let mut j = 0;
            for ch in char::decode_utf16(units.iter().copied()) {
                let ch = ch.unwrap_or('\u{FFFD}');
                chars.push((ch, (fc + j * 2) as u32));
                j += ch.len_utf16();
            }
        }
    }
    let papx = doc_papx(&wd, &table);
    let lists = doc_lists(&wd, &table);
    let lookup = |fc: u32| -> Option<PapInfo> {
        let idx = papx.partition_point(|x| x.0 <= fc);
        if idx == 0 {
            return None;
        }
        let (a, b, info) = papx[idx - 1];
        if fc >= a && fc < b { Some(info) } else { None }
    };

    let mut blocks: Vec<Block> = Vec::new();
    let mut runs: Vec<Run> = Vec::new();
    let mut cur = String::new();
    let mut cur_link: Option<String> = None;
    let mut fields: Vec<(bool, String, Option<String>)> = Vec::new();
    let mut cell_blocks: Vec<Block> = Vec::new();
    let mut row: Vec<TableCell> = Vec::new();
    let mut rows: Vec<Vec<TableCell>> = Vec::new();
    let mut prev_cell_mark = false;
    let mut warned_fields = false;

    fn flush_cur(cur: &mut String, link: &Option<String>, runs: &mut Vec<Run>) {
        if !cur.is_empty() {
            runs.push(Run { t: std::mem::take(cur), link: link.clone(), ..Default::default() });
        }
    }
    fn flush_table(rows: &mut Vec<Vec<TableCell>>, row: &mut Vec<TableCell>, blocks: &mut Vec<Block>) {
        if !row.is_empty() {
            rows.push(std::mem::take(row));
        }
        if !rows.is_empty() {
            blocks.push(Block::Table { rows: std::mem::take(rows) });
        }
    }
    let para_block = |runs: Vec<Run>, info: Option<PapInfo>| -> Block {
        let runs = merge_runs(runs);
        match info {
            Some(p) if p.ilfo > 0 && p.ilfo < 0xF801 => {
                let ordered = lists.get(&p.ilfo).and_then(|v| v.get(p.ilvl as usize).or(v.first())).copied().unwrap_or(false);
                Block::List { list: if ordered { "ol".into() } else { "ul".into() }, level: p.ilvl as u32, runs }
            }
            Some(p) if (1..=9).contains(&p.istd) && !p.in_table => heading_block(p.istd.min(6) as u32, runs, None),
            _ => Block::P { runs, align: None },
        }
    };

    for (ch, fc) in chars {
        let in_code = fields.iter().any(|f| f.0);
        match ch {
            '\u{13}' => {
                flush_cur(&mut cur, &cur_link, &mut runs);
                fields.push((true, String::new(), None));
                continue;
            }
            '\u{14}' => {
                if let Some(f) = fields.last_mut() {
                    f.0 = false;
                    f.2 = docx::field_link(&f.1);
                    if f.2.is_none() && !f.1.trim().is_empty() {
                        warned_fields = true;
                    }
                }
                flush_cur(&mut cur, &cur_link, &mut runs);
                cur_link = fields.iter().rev().find_map(|f| f.2.clone());
                continue;
            }
            '\u{15}' => {
                flush_cur(&mut cur, &cur_link, &mut runs);
                fields.pop();
                cur_link = fields.iter().rev().find_map(|f| f.2.clone());
                continue;
            }
            _ if in_code => {
                if let Some(f) = fields.iter_mut().rev().find(|f| f.0) {
                    f.1.push(ch);
                }
                continue;
            }
            _ => {}
        }
        match ch {
            '\r' => {
                flush_cur(&mut cur, &cur_link, &mut runs);
                let info = lookup(fc);
                let b = para_block(std::mem::take(&mut runs), info);
                if info.map(|i| i.in_table).unwrap_or(false) {
                    cell_blocks.push(b);
                } else {
                    flush_table(&mut rows, &mut row, &mut blocks);
                    blocks.push(b);
                }
                prev_cell_mark = false;
            }
            '\u{7}' => {
                flush_cur(&mut cur, &cur_link, &mut runs);
                let info = lookup(fc);
                let row_end = match info {
                    Some(i) => i.ttp,
                    None => prev_cell_mark && runs.is_empty() && cell_blocks.is_empty(),
                };
                if row_end {
                    runs.clear();
                    cell_blocks.clear();
                    if !row.is_empty() {
                        rows.push(std::mem::take(&mut row));
                    }
                } else {
                    if !runs.is_empty() || cell_blocks.is_empty() {
                        cell_blocks.push(Block::P { runs: merge_runs(std::mem::take(&mut runs)), align: None });
                    }
                    row.push(TableCell { blocks: std::mem::take(&mut cell_blocks), colspan: None });
                }
                prev_cell_mark = true;
            }
            '\u{b}' => cur.push('\n'),
            '\u{c}' => {
                flush_cur(&mut cur, &cur_link, &mut runs);
                if !runs.is_empty() {
                    flush_table(&mut rows, &mut row, &mut blocks);
                    blocks.push(Block::P { runs: merge_runs(std::mem::take(&mut runs)), align: None });
                }
                blocks.push(Block::PageBreak);
            }
            '\u{1e}' => cur.push('\u{2011}'),
            '\t' => cur.push('\t'),
            '\u{a0}' => cur.push('\u{a0}'),
            c if (c as u32) < 0x20 => {}
            c => {
                let link = fields.iter().rev().find_map(|f| f.2.clone());
                if link != cur_link {
                    flush_cur(&mut cur, &cur_link, &mut runs);
                    cur_link = link;
                }
                cur.push(c);
            }
        }
    }
    flush_cur(&mut cur, &cur_link, &mut runs);
    if !runs.is_empty() {
        flush_table(&mut rows, &mut row, &mut blocks);
        blocks.push(Block::P { runs: merge_runs(runs), align: None });
    }
    flush_table(&mut rows, &mut row, &mut blocks);
    Some((blocks, warned_fields))
}

/// Last-resort extraction: printable UTF-16LE runs from the raw container.
fn scrape_doc(bytes: &[u8]) -> Vec<Block> {
    let mut texts: Vec<String> = Vec::new();
    let mut cur = String::new();
    let mut i = 0usize;
    while i + 1 < bytes.len() {
        let u = u16::from_le_bytes([bytes[i], bytes[i + 1]]);
        let ok = matches!(u, 0x0D | 0x09) || (0x20..0x7F).contains(&u) || (0xA0..0x2000).contains(&u) || (0x2010..0x2070).contains(&u);
        if ok {
            cur.push(if u == 0x0D { '\n' } else { char::from_u32(u as u32).unwrap_or(' ') });
            i += 2;
        } else {
            if cur.trim().chars().count() >= 8 {
                texts.push(std::mem::take(&mut cur));
            } else {
                cur.clear();
            }
            i += if i % 2 == 0 { 2 } else { 1 };
        }
    }
    if cur.trim().chars().count() >= 8 {
        texts.push(cur);
    }
    texts
        .join("\n")
        .split('\n')
        .map(str::trim_end)
        .filter(|l| l.chars().filter(|c| c.is_alphanumeric()).count() >= 2)
        .map(|l| Block::P { runs: vec![Run::text(l)], align: None })
        .collect()
}

/// Legacy .doc -> word model (lossy; always carries a warning). Misnamed
/// RTF / DOCX / plain-text files are routed to the right parser.
pub fn parse_doc(bytes: &[u8]) -> WordDoc {
    if docx::looks_rtf(bytes) {
        if let Ok(d) = parse_rtf(bytes) {
            return d;
        }
    }
    if bytes.starts_with(b"PK") {
        if let Ok(d) = docx::parse_docx(bytes) {
            return d;
        }
    }
    if !bytes.starts_with(&[0xD0, 0xCF, 0x11, 0xE0]) {
        // Not an OLE file: many ".doc" downloads are really text or HTML.
        let mut d = parse_text(bytes, false);
        push_warning(&mut d.meta, "This .doc file is not a Word 97-2003 document; it was opened as plain text.");
        return d;
    }
    if ole_is_encrypted(bytes) {
        return doc_from(Vec::new(), &[W_ENCRYPTED.to_string()], None);
    }
    let mut warnings = Vec::new();
    let blocks = match parse_word97(bytes) {
        Some((blocks, fields)) if !blocks.is_empty() => {
            warnings.push(W_LEGACY_DOC.to_string());
            if fields {
                warnings.push(W_FIELDS.to_string());
            }
            blocks
        }
        _ => {
            warnings.push(W_DOC_FALLBACK.to_string());
            scrape_doc(bytes)
        }
    };
    let mut d = doc_from(blocks, &warnings, None);
    d.meta.insert("lossy".into(), Value::Bool(true));
    d
}

// ============================ CSV ============================

/// CSV/TSV bytes -> excel model. Encoding detection, quoting and type
/// inference live in xlsx::parse_csv_bytes (owned by the excel backend).
#[allow(dead_code)] // frozen API; main.rs now calls xlsx::parse_csv_bytes directly
pub fn csv_book(bytes: &[u8], delim: char, name: &str) -> crate::model::ExcelBook {
    crate::xlsx::parse_csv_bytes(bytes, Some(delim), name)
}

// ============================ writers ============================

/// Simple list-numbering tracker shared by the RTF writer: returns
/// (list instance id, item number) per list item so ordered lists restart.
struct ListTrack {
    in_list: bool,
    counters: [u32; 9],
    kinds: [Option<bool>; 9],
}

impl ListTrack {
    fn new() -> Self {
        ListTrack { in_list: false, counters: [0; 9], kinds: [None; 9] }
    }
    fn reset(&mut self) {
        self.in_list = false;
    }
    /// Returns (item number, restarted) for a list item.
    fn item(&mut self, ordered: bool, level: u32) -> (u32, bool) {
        let l = level.min(8) as usize;
        if !self.in_list {
            self.counters = [0; 9];
            self.kinds = [None; 9];
            self.in_list = true;
        }
        for k in l + 1..9 {
            self.counters[k] = 0;
            self.kinds[k] = None;
        }
        let restarted = self.kinds[l] != Some(ordered) || self.counters[l] == 0;
        if self.kinds[l] != Some(ordered) {
            self.counters[l] = 0;
            self.kinds[l] = Some(ordered);
        }
        self.counters[l] += 1;
        (self.counters[l], restarted)
    }
}

fn to_roman(mut n: u32) -> String {
    let t = [(1000, "m"), (900, "cm"), (500, "d"), (400, "cd"), (100, "c"), (90, "xc"), (50, "l"), (40, "xl"), (10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i")];
    let mut s = String::new();
    for (v, r) in t {
        while n >= v {
            s.push_str(r);
            n -= v;
        }
    }
    s
}

fn to_alpha(n: u32) -> String {
    let mut n = n.max(1);
    let mut s = Vec::new();
    while n > 0 {
        n -= 1;
        s.push((b'a' + (n % 26) as u8) as char);
        n /= 26;
    }
    s.iter().rev().collect()
}

fn list_marker(ordered: bool, level: u32, n: u32) -> String {
    if !ordered {
        return ["\u{2022}", "\u{25E6}", "\u{25AA}"][(level % 3) as usize].to_string();
    }
    match level % 3 {
        0 => format!("{n}."),
        1 => format!("{}.", to_alpha(n)),
        _ => format!("{}.", to_roman(n)),
    }
}

// ---------- RTF ----------

struct RtfWriter {
    fonts: Vec<String>,
    colors: Vec<String>,
    /// Per \ls N (index N-1): (ordered, level whose numbering restarts at 1).
    /// Two list definitions are shared (listid 1 bullets, 2 numbers); each
    /// ordered list instance only adds a small \listoverride entry.
    lists: Vec<(bool, Option<u32>)>,
    track: ListTrack,
    cur_ls: [u32; 9],
    bullet_ls: Option<u32>,
    text_width: u32,
}

fn rtf_escape(s: &str, out: &mut String) {
    for c in s.chars() {
        match c {
            '\\' => out.push_str("\\\\"),
            '{' => out.push_str("\\{"),
            '}' => out.push_str("\\}"),
            '\n' | '\u{b}' | '\u{2028}' => out.push_str("\\line "),
            '\t' => out.push_str("\\tab "),
            '\u{c}' => out.push_str("\\page "),
            '\u{a0}' => out.push_str("\\~"),
            '\u{ad}' => out.push_str("\\-"),
            '\u{2011}' => out.push_str("\\_"),
            '\r' => {}
            c if (c as u32) < 0x20 => {}
            c if c.is_ascii() => out.push(c),
            c => {
                let mut buf = [0u16; 2];
                let units = c.encode_utf16(&mut buf);
                let single = units.len() == 1;
                for u in units.iter() {
                    out.push_str(&format!("\\u{}", *u as i16));
                    match cp1252_byte(c).filter(|_| single) {
                        Some(b) => out.push_str(&format!("\\'{b:02x}")),
                        None => out.push('?'),
                    }
                }
            }
        }
    }
}

impl RtfWriter {
    fn font(&mut self, name: &str) -> usize {
        let clean: String = name.chars().filter(|c| !matches!(c, '{' | '}' | '\\' | ';') && !c.is_control()).collect();
        if let Some(i) = self.fonts.iter().position(|f| f.eq_ignore_ascii_case(&clean)) {
            return i;
        }
        self.fonts.push(clean);
        self.fonts.len() - 1
    }

    fn color(&mut self, c: &str) -> Option<usize> {
        let hex = docx::norm_hex(c)?;
        if let Some(i) = self.colors.iter().position(|x| *x == hex) {
            return Some(i + 1);
        }
        self.colors.push(hex);
        Some(self.colors.len())
    }

    fn runs(&mut self, runs: &[Run], out: &mut String) {
        let mut i = 0;
        while i < runs.len() {
            let mut j = i + 1;
            while j < runs.len() && runs[j].link == runs[i].link {
                j += 1;
            }
            match runs[i].link.as_deref().and_then(safe_link) {
                Some(l) => {
                    let instr = match l.strip_prefix('#') {
                        Some(a) => format!("HYPERLINK \\l \"{}\"", a.replace('"', "")),
                        None => format!("HYPERLINK \"{}\"", l.replace('"', "%22")),
                    };
                    out.push_str("{\\field{\\*\\fldinst{");
                    rtf_escape(&instr, out);
                    out.push_str("}}{\\fldrslt{");
                    for r in &runs[i..j] {
                        self.run(r, true, out);
                    }
                    out.push_str("}}}");
                }
                None => {
                    for r in &runs[i..j] {
                        self.run(r, false, out);
                    }
                }
            }
            i = j;
        }
    }

    fn run(&mut self, r: &Run, link: bool, out: &mut String) {
        if r.t.is_empty() {
            return;
        }
        out.push('{');
        let props_start = out.len();
        if r.b {
            out.push_str("\\b");
        }
        if r.i {
            out.push_str("\\i");
        }
        if r.u || link {
            out.push_str("\\ul");
        }
        if r.s {
            out.push_str("\\strike");
        }
        if r.sup {
            out.push_str("\\super");
        } else if r.sub {
            out.push_str("\\sub");
        }
        let color = r.color.as_deref().or(if link { Some("#0563C1") } else { None });
        if let Some(c) = color.and_then(|c| self.color(c)) {
            out.push_str(&format!("\\cf{c}"));
        }
        if let Some(h) = r.hl.as_deref() {
            let named = docx::norm_hex(h).map(|x| docx::highlight_name(&x).is_some()).unwrap_or(false);
            if let Some(c) = self.color(h) {
                out.push_str(&format!("\\chcbpat{c}"));
                if named {
                    out.push_str(&format!("\\highlight{c}"));
                }
            }
        }
        if let Some(sz) = r.size.filter(|s| s.is_finite() && *s > 0.0) {
            out.push_str(&format!("\\fs{}", (sz * 2.0).round().clamp(2.0, 3276.0) as u32));
        }
        if let Some(f) = r.font.as_deref().filter(|f| !f.trim().is_empty()) {
            let idx = self.font(f.trim());
            out.push_str(&format!("\\f{idx}"));
        }
        // A space only terminates a control word; without one it would be text.
        if out.len() > props_start {
            out.push(' ');
        }
        rtf_escape(&r.t, out);
        out.push('}');
    }

    fn list_ls(&mut self, ordered: bool, level: u32, restarted: bool) -> u32 {
        let l = level.min(8) as usize;
        if !ordered {
            if let Some(b) = self.bullet_ls {
                return b;
            }
            self.lists.push((false, None));
            let id = self.lists.len() as u32;
            self.bullet_ls = Some(id);
            return id;
        }
        if restarted || self.cur_ls[l] == 0 {
            self.lists.push((true, Some(l as u32)));
            self.cur_ls[l] = self.lists.len() as u32;
        }
        self.cur_ls[l]
    }

    fn image(&mut self, src: &str, alt: Option<&str>, w: Option<u32>, h: Option<u32>, width: u32, out: &mut String) {
        let blip = docx::parse_data_url(src).and_then(|(mime, bytes)| {
            let mime = docx::sniff_image(&bytes).map(str::to_string).unwrap_or(mime);
            let kind = match mime.as_str() {
                "image/png" => "\\pngblip",
                "image/jpeg" => "\\jpegblip",
                "image/x-emf" => "\\emfblip",
                "image/x-wmf" => "\\wmetafile8",
                _ => return None,
            };
            Some((kind, bytes))
        });
        let Some((kind, bytes)) = blip else {
            out.push_str("{\\pard\\plain ");
            rtf_escape(&format!("[image: {}]", alt.filter(|a| !a.trim().is_empty()).unwrap_or("image")), out);
            out.push_str("\\par}\n");
            return;
        };
        let (iw, ih) = docx::image_dims(&bytes).unwrap_or((w.unwrap_or(320), h.unwrap_or(240)));
        let (mut pw, mut ph) = match (w, h) {
            (Some(w), Some(h)) => (w as u64, h as u64),
            (Some(w), None) => (w as u64, w as u64 * ih as u64 / iw.max(1) as u64),
            (None, Some(h)) => (h as u64 * iw as u64 / ih.max(1) as u64, h as u64),
            _ => (iw as u64, ih as u64),
        };
        let max = (width / 15) as u64;
        if pw > max && max > 0 {
            ph = ph * max / pw;
            pw = max;
        }
        out.push_str(&format!("{{\\pard\\plain{{\\pict{kind}\\picw{iw}\\pich{ih}\\picwgoal{}\\pichgoal{}\n", pw.max(1) * 15, ph.max(1) * 15));
        let hex = b"0123456789abcdef";
        for chunk in bytes.chunks(64) {
            for b in chunk {
                out.push(hex[(b >> 4) as usize] as char);
                out.push(hex[(b & 15) as usize] as char);
            }
            out.push('\n');
        }
        out.push_str("}\\par}\n");
    }

    fn cell_text(&mut self, blocks: &[Block], out: &mut String, width: u32) {
        let mut first = true;
        for b in blocks {
            let Some(b) = resolve_block(b) else { continue };
            if !first {
                out.push_str("\\par ");
            }
            first = false;
            match &*b {
                Block::Table { rows } => {
                    for row in rows {
                        for (k, c) in row.iter().enumerate() {
                            if k > 0 {
                                out.push_str("\\tab ");
                            }
                            let t: String = c.blocks.iter().map(|x| x.para_runs().iter().map(|r| r.t.as_str()).collect::<String>()).collect::<Vec<_>>().join(" ");
                            rtf_escape(&t, out);
                        }
                        out.push_str("\\line ");
                    }
                }
                Block::Img { src, alt, w, h } => {
                    let mut tmp = String::new();
                    self.image(src, alt.as_deref(), *w, *h, width, &mut tmp);
                    // Strip the paragraph wrapper; keep just the picture group.
                    let pict = tmp.find("{\\pict").map(|p| &tmp[p..tmp.rfind("}\\par}").unwrap_or(tmp.len())]).unwrap_or("").to_string();
                    out.push_str(&pict);
                    if !pict.ends_with('}') {
                        out.push('}');
                    }
                }
                Block::List { list, level, runs } => {
                    let (n, _) = self.track.item(list == "ol", *level);
                    rtf_escape(&format!("{}{} ", "    ".repeat(*level as usize), list_marker(list == "ol", *level, n)), out);
                    self.runs(runs, out);
                }
                Block::PageBreak => {}
                other => {
                    let runs = other.para_runs();
                    if matches!(other, Block::H1 { .. } | Block::H2 { .. } | Block::H3 { .. } | Block::H4 { .. } | Block::H5 { .. } | Block::H6 { .. }) {
                        out.push_str("{\\b ");
                        self.runs(runs, out);
                        out.push('}');
                    } else {
                        self.runs(runs, out);
                    }
                }
            }
        }
        self.track.reset();
    }

    fn blocks(&mut self, blocks: &[Block], out: &mut String) {
        for b in blocks {
            let Some(b) = resolve_block(b) else { continue };
            let b: &Block = &b;
            if !matches!(b, Block::List { .. }) {
                self.track.reset();
            }
            let align = |a: &Option<String>| match a.as_deref() {
                Some("center") => "\\qc",
                Some("right") => "\\qr",
                Some("justify") => "\\qj",
                _ => "\\ql",
            };
            match b {
                Block::P { runs, align: a } => {
                    out.push_str(&format!("\\pard\\plain\\s0{}\\f0 ", align(a)));
                    self.runs(runs, out);
                    out.push_str("\\par\n");
                }
                Block::H1 { runs, align: a } | Block::H2 { runs, align: a } | Block::H3 { runs, align: a } | Block::H4 { runs, align: a } | Block::H5 { runs, align: a } | Block::H6 { runs, align: a } => {
                    let n = match b {
                        Block::H1 { .. } => 1,
                        Block::H2 { .. } => 2,
                        Block::H3 { .. } => 3,
                        Block::H4 { .. } => 4,
                        Block::H5 { .. } => 5,
                        _ => 6,
                    };
                    let fs = [36, 30, 26, 24, 22, 22][n - 1];
                    out.push_str(&format!("\\pard\\plain\\s{n}\\outlinelevel{}\\keepn\\sb240\\sa80{}\\f0\\b\\fs{fs} ", n - 1, align(a)));
                    self.runs(runs, out);
                    out.push_str("\\par\n");
                }
                Block::Quote { runs, align: a } => {
                    out.push_str(&format!("\\pard\\plain\\s7\\li864\\ri864{}\\f0\\i ", align(a)));
                    self.runs(runs, out);
                    out.push_str("\\par\n");
                }
                Block::Code { runs, align: a } => {
                    out.push_str(&format!("\\pard\\plain\\s8\\sa0{}\\f1\\fs20 ", align(a)));
                    self.runs(runs, out);
                    out.push_str("\\par\n");
                }
                Block::List { list, level, runs } => {
                    let ordered = list == "ol";
                    let lvl = (*level).min(8);
                    let (n, restarted) = self.track.item(ordered, lvl);
                    let ls = self.list_ls(ordered, lvl, restarted && n == 1);
                    out.push_str("{\\listtext\\pard\\plain\\f0 ");
                    rtf_escape(&list_marker(ordered, lvl, n), out);
                    out.push_str("\\tab}");
                    out.push_str(&format!("\\pard\\plain\\s9\\fi-360\\li{}\\ls{ls}\\ilvl{lvl}\\f0 ", 720 * (lvl + 1)));
                    self.runs(runs, out);
                    out.push_str("\\par\n");
                }
                Block::Table { rows } => {
                    let rows: Vec<&Vec<TableCell>> = rows.iter().filter(|r| !r.is_empty()).collect();
                    if rows.is_empty() {
                        continue;
                    }
                    let span = |c: &TableCell| c.colspan.unwrap_or(1).clamp(1, 63);
                    let cols = rows.iter().map(|r| r.iter().map(span).sum::<u32>()).max().unwrap_or(1).max(1);
                    let colw = (self.text_width / cols).max(300);
                    for row in rows {
                        out.push_str("\\trowd\\trgaph108\\trleft0");
                        let mut x = 0;
                        for c in row.iter() {
                            x += colw * span(c);
                            out.push_str(&format!("\\clbrdrt\\brdrs\\brdrw10\\clbrdrl\\brdrs\\brdrw10\\clbrdrb\\brdrs\\brdrw10\\clbrdrr\\brdrs\\brdrw10\\cellx{x}"));
                        }
                        out.push('\n');
                        for c in row.iter() {
                            out.push_str("\\pard\\plain\\intbl\\f0 ");
                            let w = colw * span(c);
                            self.cell_text(&c.blocks, out, w);
                            out.push_str("\\cell\n");
                        }
                        out.push_str("\\row\n");
                    }
                    out.push_str("\\pard\n");
                }
                Block::Img { src, alt, w, h } => {
                    let tw = self.text_width;
                    self.image(src, alt.as_deref(), *w, *h, tw, out);
                }
                Block::PageBreak => out.push_str("\\pard\\plain\\page\n"),
                Block::Other(_) => {}
            }
        }
    }
}

/// Word model -> RTF document bytes (7-bit ASCII; non-ASCII as \uN with a
/// Windows-1252 fallback). Formatting, colors, highlight, fonts, alignment,
/// headings (with outline levels), real lists, tables, hyperlinks and
/// PNG/JPEG/EMF/WMF pictures are written.
pub fn write_rtf(doc: &WordDoc) -> Vec<u8> {
    let (pw, ph, mt, mr, mb, ml) = docx::page_twips(&doc.meta);
    let mut w = RtfWriter {
        fonts: vec!["Calibri".into(), "Consolas".into()],
        colors: Vec::new(),
        lists: Vec::new(),
        track: ListTrack::new(),
        cur_ls: [0; 9],
        bullet_ls: None,
        text_width: pw.saturating_sub(ml + mr).max(1440),
    };
    let mut body = String::new();
    w.blocks(&doc.blocks, &mut body);
    // Header / footer (before the font/color tables are written: they may add entries).
    let mut hf = String::new();
    for kind in ["header", "footer"] {
        let Some((runs, align)) = docx::header_footer_meta(doc.meta.get(kind)) else { continue };
        let q = match align.as_deref() {
            Some("center") => "\\qc",
            Some("right") => "\\qr",
            Some("justify") => "\\qj",
            _ => "\\ql",
        };
        hf.push_str(&format!("{{\\{kind}\\pard\\plain{q}\\f0\\fs20 "));
        let mut pending: Vec<Run> = Vec::new();
        for (run, field) in docx::split_field_runs(&runs) {
            match field {
                None => pending.push(run),
                Some(f) => {
                    w.runs(&std::mem::take(&mut pending), &mut hf);
                    if f == "PAGE" {
                        hf.push_str("{\\chpgn}");
                    } else {
                        hf.push_str("{\\field{\\*\\fldinst NUMPAGES }{\\fldrslt 1}}");
                    }
                }
            }
        }
        w.runs(&pending, &mut hf);
        hf.push_str("\\par}\n");
    }

    let mut out = String::from("{\\rtf1\\ansi\\ansicpg1252\\deff0\\uc1\\deflang1033\n{\\fonttbl");
    for (i, f) in w.fonts.iter().enumerate() {
        let family = if i == 1 || is_mono_font(f) { "\\fmodern" } else { "\\fswiss" };
        out.push_str(&format!("{{\\f{i}{family}\\fcharset0 "));
        rtf_escape(f, &mut out);
        out.push_str(";}");
    }
    out.push_str("}\n{\\colortbl;");
    for c in &w.colors {
        let v = u32::from_str_radix(c, 16).unwrap_or(0);
        out.push_str(&format!("\\red{}\\green{}\\blue{};", v >> 16, (v >> 8) & 255, v & 255));
    }
    out.push_str("}\n{\\stylesheet{\\s0\\f0 Normal;}");
    for n in 1..=6 {
        let fs = [36, 30, 26, 24, 22, 22][n - 1];
        out.push_str(&format!("{{\\s{n}\\outlinelevel{}\\keepn\\sb240\\sa80\\sbasedon0\\snext0\\f0\\b\\fs{fs} heading {n};}}", n - 1));
    }
    out.push_str("{\\s7\\li864\\ri864\\sbasedon0\\snext0\\f0\\i Quote;}{\\s8\\sa0\\sbasedon0\\snext8\\f1\\fs20 Code;}{\\s9\\li720\\sbasedon0\\snext9\\f0 List Paragraph;}}\n");
    if !w.lists.is_empty() {
        out.push_str("{\\*\\listtable");
        for (id, ordered) in [(1, false), (2, true)] {
            out.push_str(&format!("{{\\list\\listtemplateid{id}\\listhybrid"));
            for l in 0..9u32 {
                let indent = 720 * (l + 1);
                if ordered {
                    let nfc = [0, 4, 2][(l % 3) as usize];
                    out.push_str(&format!(
                        "{{\\listlevel\\levelnfc{nfc}\\levelnfcn{nfc}\\leveljc0\\leveljcn0\\levelfollow0\\levelstartat1\\levelspace0\\levelindent0{{\\leveltext\\'02\\'{l:02x}.;}}{{\\levelnumbers\\'01;}}\\fi-360\\li{indent}\\lin{indent} }}"
                    ));
                } else {
                    let glyph = [8226, 9702, 9642][(l % 3) as usize];
                    out.push_str(&format!(
                        "{{\\listlevel\\levelnfc23\\levelnfcn23\\leveljc0\\leveljcn0\\levelfollow0\\levelstartat1\\levelspace0\\levelindent0{{\\leveltext\\'01\\u{glyph} ?;}}{{\\levelnumbers;}}\\fi-360\\li{indent}\\lin{indent} }}"
                    ));
                }
            }
            out.push_str(&format!("{{\\listname ;}}\\listid{id}}}"));
        }
        out.push_str("}\n{\\*\\listoverridetable");
        for (k, (ordered, restart)) in w.lists.iter().enumerate() {
            let ls = k + 1;
            match (ordered, restart) {
                (true, Some(lvl)) => {
                    // One \lfolevel per level up to the restarted one.
                    out.push_str(&format!("{{\\listoverride\\listid2\\listoverridecount{}", lvl + 1));
                    for _ in 0..*lvl {
                        out.push_str("{\\lfolevel}");
                    }
                    out.push_str(&format!("{{\\lfolevel\\listoverridestartat\\levelstartat1}}\\ls{ls}}}"));
                }
                _ => out.push_str(&format!("{{\\listoverride\\listid{}\\listoverridecount0\\ls{ls}}}", if *ordered { 2 } else { 1 })),
            }
        }
        out.push_str("}\n");
    }
    if let Some(t) = doc.meta.get("title").and_then(Value::as_str).filter(|t| !t.trim().is_empty()) {
        out.push_str("{\\info{\\title ");
        rtf_escape(t, &mut out);
        out.push_str("}}\n");
    }
    out.push_str(&format!("\\paperw{pw}\\paperh{ph}\\margl{ml}\\margr{mr}\\margt{mt}\\margb{mb}{}\\viewkind4\n", if pw > ph { "\\landscape" } else { "" }));
    out.push_str(&hf);
    out.push_str(&body);
    out.push('}');
    out.into_bytes()
}

// ---------- Markdown ----------

fn md_escape(t: &str) -> String {
    let mut o = String::with_capacity(t.len());
    for c in t.chars() {
        match c {
            '\\' | '`' | '*' | '_' | '[' | ']' | '<' | '>' | '~' => {
                o.push('\\');
                o.push(c);
            }
            '\n' | '\u{b}' => o.push_str("\\\n"),
            '\r' => {}
            '\u{a0}' => o.push_str("&nbsp;"),
            c if (c as u32) < 0x20 && c != '\t' => {}
            c => o.push(c),
        }
    }
    o
}

fn md_url(u: &str) -> String {
    if u.contains(' ') || u.contains('(') || u.contains(')') || u.contains('<') {
        format!("<{}>", u.replace('<', "%3C").replace('>', "%3E"))
    } else {
        u.to_string()
    }
}

fn md_run(r: &Run) -> String {
    if r.t.is_empty() {
        return String::new();
    }
    if r.font.as_deref().map(is_mono_font).unwrap_or(false) && !r.t.contains('\n') && !r.t.trim().is_empty() {
        let max = r.t.split(|c| c != '`').map(str::len).max().unwrap_or(0);
        let fence = "`".repeat(max + 1);
        let pad = if r.t.starts_with('`') || r.t.ends_with('`') { " " } else { "" };
        return format!("{fence}{pad}{}{pad}{fence}", r.t);
    }
    let core = r.t.trim();
    if core.is_empty() {
        return md_escape(&r.t);
    }
    let lead = &r.t[..r.t.len() - r.t.trim_start().len()];
    let trail = &r.t[r.t.trim_end().len()..];
    let mut c = md_escape(core);
    if r.sup {
        c = format!("<sup>{c}</sup>");
    } else if r.sub {
        c = format!("<sub>{c}</sub>");
    }
    if r.u {
        c = format!("<u>{c}</u>");
    }
    if r.s {
        c = format!("~~{c}~~");
    }
    c = match (r.b, r.i) {
        (true, true) => format!("***{c}***"),
        (true, false) => format!("**{c}**"),
        (false, true) => format!("*{c}*"),
        _ => c,
    };
    format!("{}{c}{}", md_escape(lead), md_escape(trail))
}

fn md_runs(runs: &[Run]) -> String {
    let mut s = String::new();
    let mut i = 0;
    while i < runs.len() {
        let mut j = i + 1;
        while j < runs.len() && runs[j].link == runs[i].link {
            j += 1;
        }
        let inner: String = runs[i..j].iter().map(md_run).collect();
        match runs[i].link.as_deref().and_then(safe_link) {
            Some(l) if !inner.trim().is_empty() => s.push_str(&format!("[{inner}]({})", md_url(&l))),
            _ => s.push_str(&inner),
        }
        i = j;
    }
    s
}

/// Escape characters that would start a block construct at line start.
fn md_line_start(s: &str) -> String {
    s.split('\n')
        .map(|line| {
            let t = line;
            if t.starts_with('#') || t.starts_with('>') || t.starts_with('-') || t.starts_with('+') || t.starts_with('=') || t.starts_with('|') {
                return format!("\\{t}");
            }
            let digits = t.chars().take_while(|c| c.is_ascii_digit()).count();
            if digits > 0 && (t[digits..].starts_with(". ") || t[digits..].starts_with(") ") || &t[digits..] == "." || &t[digits..] == ")") {
                return format!("{}\\{}", &t[..digits], &t[digits..]);
            }
            t.to_string()
        })
        .collect::<Vec<_>>()
        .join("\n")
}

fn md_img(src: &str, alt: Option<&str>) -> String {
    let alt = alt.unwrap_or("").replace('[', "\\[").replace(']', "\\]").replace('\n', " ");
    format!("![{alt}]({})", md_url(src))
}

fn md_cell(blocks: &[Block]) -> String {
    let mut parts = Vec::new();
    for b in blocks {
        let Some(b) = resolve_block(b) else { continue };
        let s = match &*b {
            Block::Img { src, alt, .. } => md_img(src, alt.as_deref()),
            Block::Table { rows } => rows
                .iter()
                .map(|r| r.iter().map(|c| md_cell(&c.blocks)).collect::<Vec<_>>().join(" / "))
                .collect::<Vec<_>>()
                .join("<br>"),
            Block::List { list, runs, .. } => format!("{} {}", if list == "ol" { "1." } else { "•" }, md_runs(runs)),
            Block::PageBreak => String::new(),
            other => md_runs(other.para_runs()),
        };
        if !s.trim().is_empty() {
            parts.push(s);
        }
    }
    parts.join("<br>").replace("\\\n", "<br>").replace('\n', "<br>").replace('|', "\\|")
}

/// Word model -> Markdown (CommonMark + GFM tables).
pub fn write_markdown(doc: &WordDoc) -> String {
    let mut out = String::new();
    let mut track = ListTrack::new();
    let mut prev_list = false;
    // Markdown cannot skip list levels (a jump of more than one level, or a
    // first item that is indented, would read back as code or continuation
    // text), so each item is at most one level deeper than the previous one.
    let mut md_level: Option<u32> = None;
    for b in &doc.blocks {
        let Some(b) = resolve_block(b) else { continue };
        let b: &Block = &b;
        let is_list = matches!(b, Block::List { .. });
        if !is_list {
            track.reset();
            md_level = None;
        }
        let text = match b {
            Block::P { runs, .. } => md_line_start(&md_runs(runs)),
            Block::H1 { runs, .. } | Block::H2 { runs, .. } | Block::H3 { runs, .. } | Block::H4 { runs, .. } | Block::H5 { runs, .. } | Block::H6 { runs, .. } => {
                let n = match b {
                    Block::H1 { .. } => 1,
                    Block::H2 { .. } => 2,
                    Block::H3 { .. } => 3,
                    Block::H4 { .. } => 4,
                    Block::H5 { .. } => 5,
                    _ => 6,
                };
                let t = md_runs(runs).replace("\\\n", " ");
                if t.trim().is_empty() { String::new() } else { format!("{} {}", "#".repeat(n), t) }
            }
            Block::Quote { runs, .. } => {
                let t = md_line_start(&md_runs(runs));
                if t.trim().is_empty() { String::new() } else { t.split('\n').map(|l| format!("> {l}")).collect::<Vec<_>>().join("\n") }
            }
            Block::Code { runs, .. } => {
                let t: String = runs.iter().map(|r| r.t.as_str()).collect();
                let max = t.split(|c| c != '`').map(str::len).max().unwrap_or(0);
                let fence = "`".repeat(max.max(2) + 1);
                format!("{fence}\n{t}\n{fence}")
            }
            Block::List { list, level, runs } => {
                let ordered = list == "ol";
                let lvl = (*level).min(8).min(md_level.map_or(0, |p| p + 1));
                md_level = Some(lvl);
                let (n, _) = track.item(ordered, lvl);
                let indent = "    ".repeat(lvl as usize);
                let marker = if ordered { format!("{n}.") } else { "-".to_string() };
                let cont = format!("\n{indent}{}", " ".repeat(marker.len() + 1));
                let content = md_runs(runs).replace('\n', &cont);
                format!("{indent}{marker} {content}")
            }
            Block::Table { rows } => {
                let rows: Vec<&Vec<TableCell>> = rows.iter().filter(|r| !r.is_empty()).collect();
                if rows.is_empty() {
                    String::new()
                } else {
                    let span = |c: &TableCell| c.colspan.unwrap_or(1).clamp(1, 63) as usize;
                    let cols = rows.iter().map(|r| r.iter().map(span).sum::<usize>()).max().unwrap_or(1);
                    let mut lines = Vec::new();
                    let mut header_align: Vec<String> = Vec::new();
                    for (ri, row) in rows.iter().enumerate() {
                        let mut cells: Vec<String> = Vec::new();
                        for c in row.iter() {
                            if ri == 0 {
                                let a = c.blocks.first().and_then(|b| match b {
                                    Block::P { align, .. } => align.clone(),
                                    _ => None,
                                });
                                let sep = match a.as_deref() {
                                    Some("center") => ":---:",
                                    Some("right") => "---:",
                                    _ => "---",
                                };
                                for _ in 0..span(c) {
                                    header_align.push(sep.to_string());
                                }
                            }
                            cells.push(md_cell(&c.blocks));
                            for _ in 1..span(c) {
                                cells.push(String::new());
                            }
                        }
                        while cells.len() < cols {
                            cells.push(String::new());
                        }
                        lines.push(format!("| {} |", cells.join(" | ")));
                        if ri == 0 {
                            while header_align.len() < cols {
                                header_align.push("---".into());
                            }
                            lines.push(format!("| {} |", header_align.join(" | ")));
                        }
                    }
                    lines.join("\n")
                }
            }
            Block::Img { src, alt, .. } => md_img(src, alt.as_deref()),
            Block::PageBreak => "---".to_string(),
            Block::Other(_) => String::new(),
        };
        if text.trim().is_empty() && !matches!(b, Block::Code { .. }) {
            prev_list = false;
            continue;
        }
        if !out.is_empty() {
            out.push_str(if is_list && prev_list { "\n" } else { "\n\n" });
        }
        out.push_str(&text);
        prev_list = is_list;
    }
    out.push('\n');
    out
}

// ---------- HTML ----------

fn html_esc(s: &str) -> String {
    let mut o = String::with_capacity(s.len());
    for c in s.chars() {
        match c {
            '&' => o.push_str("&amp;"),
            '<' => o.push_str("&lt;"),
            '>' => o.push_str("&gt;"),
            '"' => o.push_str("&quot;"),
            '\'' => o.push_str("&#39;"),
            c if (c as u32) < 0x20 && !matches!(c, '\t' | '\n') => {}
            c => o.push(c),
        }
    }
    o
}

fn html_run(r: &Run) -> String {
    if r.t.is_empty() {
        return String::new();
    }
    let mut c = html_esc(&r.t).replace('\n', "<br>").replace('\t', "<span class=\"tab\">\t</span>").replace('\u{c}', "<br>");
    let mut style = Vec::new();
    if let Some(col) = r.color.as_deref().and_then(docx::norm_hex) {
        style.push(format!("color:#{col}"));
    }
    if let Some(h) = r.hl.as_deref().and_then(docx::norm_hex) {
        style.push(format!("background-color:#{h}"));
    }
    if let Some(sz) = r.size.filter(|s| s.is_finite() && *s > 0.0 && *s < 1000.0) {
        style.push(format!("font-size:{sz}pt"));
    }
    if let Some(f) = r.font.as_deref() {
        let f: String = f.chars().filter(|c| !matches!(c, '"' | '\'' | ';' | '<' | '>' | '{' | '}' | '\\')).collect();
        if !f.trim().is_empty() {
            style.push(format!("font-family:'{}'", f.trim()));
        }
    }
    if !style.is_empty() {
        c = format!("<span style=\"{}\">{c}</span>", html_esc(&style.join(";")));
    }
    if r.sup {
        c = format!("<sup>{c}</sup>");
    } else if r.sub {
        c = format!("<sub>{c}</sub>");
    }
    if r.s {
        c = format!("<s>{c}</s>");
    }
    if r.u {
        c = format!("<u>{c}</u>");
    }
    if r.i {
        c = format!("<em>{c}</em>");
    }
    if r.b {
        c = format!("<strong>{c}</strong>");
    }
    c
}

fn html_runs(runs: &[Run]) -> String {
    let mut s = String::new();
    let mut i = 0;
    while i < runs.len() {
        let mut j = i + 1;
        while j < runs.len() && runs[j].link == runs[i].link {
            j += 1;
        }
        let inner: String = runs[i..j].iter().map(html_run).collect();
        match runs[i].link.as_deref().and_then(safe_link) {
            Some(l) => s.push_str(&format!("<a href=\"{}\">{inner}</a>", html_esc(&l))),
            None => s.push_str(&inner),
        }
        i = j;
    }
    s
}

fn html_align(a: &Option<String>) -> String {
    match a.as_deref() {
        Some(v @ ("center" | "right" | "justify" | "left")) => format!(" style=\"text-align:{v}\""),
        _ => String::new(),
    }
}

fn html_blocks(blocks: &[Block], out: &mut String) {
    // Open list stack: (level, ordered).
    let mut stack: Vec<(u32, bool)> = Vec::new();
    let close_all = |stack: &mut Vec<(u32, bool)>, out: &mut String| {
        while let Some((_, o)) = stack.pop() {
            out.push_str(if o { "</li></ol>\n" } else { "</li></ul>\n" });
        }
    };
    for b in blocks {
        let Some(b) = resolve_block(b) else { continue };
        let b: &Block = &b;
        if let Block::List { list, level, runs } = b {
            let ordered = list == "ol";
            let lvl = (*level).min(8);
            while let Some(&(l, o)) = stack.last() {
                if l > lvl || (l == lvl && o != ordered) {
                    stack.pop();
                    out.push_str(if o { "</li></ol>\n" } else { "</li></ul>\n" });
                } else {
                    break;
                }
            }
            match stack.last() {
                Some(&(l, _)) if l == lvl => out.push_str("</li>\n<li>"),
                _ => {
                    let start = stack.last().map(|(l, _)| l + 1).unwrap_or(0);
                    for k in start..=lvl {
                        let tag = if ordered { "ol" } else { "ul" };
                        if k < lvl {
                            out.push_str(&format!("<{tag}><li class=\"nested\">"));
                        } else {
                            out.push_str(&format!("<{tag}>\n<li>"));
                        }
                        stack.push((k, ordered));
                    }
                }
            }
            out.push_str(&html_runs(runs));
            continue;
        }
        close_all(&mut stack, out);
        match b {
            Block::P { runs, align } => {
                let inner = html_runs(runs);
                out.push_str(&format!("<p{}>{}</p>\n", html_align(align), if inner.is_empty() { "<br>".to_string() } else { inner }));
            }
            Block::H1 { runs, align } => out.push_str(&format!("<h1{}>{}</h1>\n", html_align(align), html_runs(runs))),
            Block::H2 { runs, align } => out.push_str(&format!("<h2{}>{}</h2>\n", html_align(align), html_runs(runs))),
            Block::H3 { runs, align } => out.push_str(&format!("<h3{}>{}</h3>\n", html_align(align), html_runs(runs))),
            Block::H4 { runs, align } => out.push_str(&format!("<h4{}>{}</h4>\n", html_align(align), html_runs(runs))),
            Block::H5 { runs, align } => out.push_str(&format!("<h5{}>{}</h5>\n", html_align(align), html_runs(runs))),
            Block::H6 { runs, align } => out.push_str(&format!("<h6{}>{}</h6>\n", html_align(align), html_runs(runs))),
            Block::Quote { runs, align } => out.push_str(&format!("<blockquote><p{}>{}</p></blockquote>\n", html_align(align), html_runs(runs))),
            Block::Code { runs, .. } => {
                let t: String = runs.iter().map(|r| r.t.as_str()).collect();
                out.push_str(&format!("<pre><code>{}</code></pre>\n", html_esc(&t)));
            }
            Block::Table { rows } => {
                out.push_str("<table>\n");
                for row in rows.iter().filter(|r| !r.is_empty()) {
                    out.push_str("<tr>");
                    for c in row {
                        let span = c.colspan.filter(|s| *s > 1).map(|s| format!(" colspan=\"{}\"", s.min(63))).unwrap_or_default();
                        out.push_str(&format!("<td{span}>"));
                        let mut inner = String::new();
                        html_blocks(&c.blocks, &mut inner);
                        out.push_str(inner.trim_end());
                        out.push_str("</td>");
                    }
                    out.push_str("</tr>\n");
                }
                out.push_str("</table>\n");
            }
            Block::Img { src, alt, w, h } => {
                let src_ok = src.trim().to_ascii_lowercase().starts_with("data:image/") || safe_link(src).is_some();
                if src_ok {
                    let mut attrs = String::new();
                    if let Some(w) = w {
                        attrs.push_str(&format!(" width=\"{w}\""));
                    }
                    if let Some(h) = h {
                        attrs.push_str(&format!(" height=\"{h}\""));
                    }
                    out.push_str(&format!("<p class=\"figure\"><img src=\"{}\" alt=\"{}\"{attrs}></p>\n", html_esc(src.trim()), html_esc(alt.as_deref().unwrap_or(""))));
                }
            }
            Block::PageBreak => out.push_str("<div class=\"page-break\"></div>\n"),
            Block::List { .. } | Block::Other(_) => {}
        }
    }
    close_all(&mut stack, out);
}

const HTML_CSS: &str = "body{margin:0;background:#f3f3f3;color:#1f1f1f;font:11pt/1.5 Calibri,'Segoe UI',Arial,sans-serif}\
article{max-width:48rem;margin:2rem auto;padding:3rem 3.5rem;background:#fff;box-shadow:0 1px 4px rgba(0,0,0,.15)}\
h1,h2,h3,h4,h5,h6{line-height:1.25;margin:1.2em 0 .4em}h1{font-size:1.8em}h2{font-size:1.45em}h3{font-size:1.2em}\
p{margin:0 0 .6em}blockquote{margin:1em 2em;padding:.2em 1em;border-left:3px solid #c8c8c8;color:#404040;font-style:italic}\
pre{background:#f2f2f2;padding:.75em 1em;overflow:auto;border-radius:4px}code,pre{font-family:Consolas,'Courier New',monospace;font-size:10pt}\
table{border-collapse:collapse;margin:.6em 0 1em;max-width:100%}td,th{border:1px solid #bbb;padding:.3em .6em;vertical-align:top}td>p:last-child{margin-bottom:0}\
img{max-width:100%;height:auto}.figure{margin:.6em 0}li.nested{list-style:none}ul,ol{margin:0 0 .6em;padding-left:2em}.tab{white-space:pre}\
a{color:#0563c1}.doc-header,.doc-footer{color:#595959;font-size:9.5pt}.doc-header{border-bottom:1px solid #e3e3e3;padding-bottom:.4em;margin-bottom:1.2em}.doc-footer{border-top:1px solid #e3e3e3;padding-top:.4em;margin-top:1.6em}.page-break{break-after:page;page-break-after:always;border-top:1px dashed #bbb;margin:1.5em 0}\
@media print{body{background:#fff}article{box-shadow:none;margin:0;max-width:none;padding:0}.page-break{border:0;margin:0}}";

/// Word model -> standalone, styled HTML document.
pub fn write_html(doc: &WordDoc) -> String {
    let title = doc
        .meta
        .get("title")
        .and_then(Value::as_str)
        .map(str::to_string)
        .filter(|t| !t.trim().is_empty())
        .or_else(|| {
            doc.blocks.iter().find_map(|b| match b {
                Block::H1 { runs, .. } | Block::H2 { runs, .. } => Some(runs.iter().map(|r| r.t.as_str()).collect::<String>()).filter(|t| !t.trim().is_empty()),
                _ => None,
            })
        })
        .unwrap_or_else(|| "Document".to_string());
    let mut body = String::new();
    // A web page is one page: page-number fields read "1".
    let band = |kind: &str, tag: &str| -> String {
        match docx::header_footer_meta(doc.meta.get(kind)) {
            Some((runs, align)) => {
                let runs: Vec<Run> = runs.into_iter().map(|r| Run { t: r.t.replace("{PAGE}", "1").replace("{NUMPAGES}", "1"), ..r }).collect();
                format!("<{tag} class=\"doc-{kind}\"{}>{}</{tag}>\n", html_align(&align), html_runs(&runs))
            }
            None => String::new(),
        }
    };
    body.push_str(&band("header", "header"));
    html_blocks(&doc.blocks, &mut body);
    body.push_str(&band("footer", "footer"));
    format!(
        "<!DOCTYPE html>\n<html lang=\"en\">\n<head>\n<meta charset=\"utf-8\">\n<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\n<meta name=\"generator\" content=\"office-backend\">\n<title>{}</title>\n<style>{HTML_CSS}</style>\n</head>\n<body>\n<article>\n{body}</article>\n</body>\n</html>\n",
        html_esc(title.trim())
    )
}

// ============================ tests ============================

#[cfg(test)]
mod tests {
    use super::*;
    use crate::docx::tests::{zip_parts, PNG_1X1};
    use serde_json::json;

    fn texts(b: &Block) -> String {
        b.para_runs().iter().map(|r| r.t.as_str()).collect()
    }

    fn warnings(d: &WordDoc) -> Vec<String> {
        d.meta.get("warnings").and_then(Value::as_array).map(|a| a.iter().filter_map(|v| v.as_str().map(String::from)).collect()).unwrap_or_default()
    }

    // ---------- RTF ----------

    #[test]
    fn rtf_control_words_paragraphs_and_formatting() {
        let rtf = br"{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\froman Times New Roman;}{\f1\fswiss Arial;}}{\colortbl;\red255\green0\blue0;}{\stylesheet{\s0 Normal;}{\s1 heading 1;}}\f0\fs24 Hello {\b bold} world\par Second \cf1 red\cf0  line\par}";
        let d = parse_rtf(rtf).unwrap();
        assert_eq!(d.blocks.len(), 2, "{:?}", d.blocks);
        assert_eq!(texts(&d.blocks[0]), "Hello bold world");
        let runs = d.blocks[0].para_runs();
        assert!(runs.iter().any(|r| r.t == "bold" && r.b));
        assert!(runs.iter().all(|r| r.t == "bold" || !r.b));
        assert_eq!(texts(&d.blocks[1]), "Second red line");
        assert_eq!(d.blocks[1].para_runs().iter().find(|r| r.t == "red").unwrap().color.as_deref(), Some("#FF0000"));
        assert!(!texts(&d.blocks[0]).contains("Times"));
    }

    #[test]
    fn rtf_hex_escapes_unicode_and_apostrophes() {
        let d = parse_rtf(br"{\rtf1\ansi Don't caf\'e9 \u8364? \uc2\u8364\'80\'80 euro\par Next}").unwrap();
        assert_eq!(texts(&d.blocks[0]), "Don't café € € euro");
        assert_eq!(texts(&d.blocks[1]), "Next");
        // Cyrillic via \fcharset204 and \ansicpg1251
        let d = parse_rtf(br"{\rtf1\ansi\ansicpg1251{\fonttbl{\f0\fcharset204 Arial;}}\f0 \'cf\'f0\'e8\'e2\'e5\'f2\par}").unwrap();
        assert_eq!(texts(&d.blocks[0]), "Привет");
        // Raw UTF-8 written by sloppy generators (no escapes) is detected.
        let d = parse_rtf("{\\rtf1\\ansi Grüße 荤 \\'e9\\par}".as_bytes()).unwrap();
        assert_eq!(texts(&d.blocks[0]), "Grüße 荤 é");
        // Surrogate pair
        let d = parse_rtf(br"{\rtf1 smile \u-10179?\u-8704?\par}").unwrap();
        assert_eq!(texts(&d.blocks[0]), "smile 😀");
    }

    #[test]
    fn rtf_wordpad_style_document() {
        let rtf = b"{\\rtf1\\ansi\\ansicpg1252\\deff0{\\fonttbl{\\f0\\fswiss\\fcharset0 Arial;}{\\f1\\fnil Calibri;}}{\\colortbl ;\\red255\\green0\\blue0;}{\\stylesheet{\\s0 Normal;}{\\s1 heading 1;}}{\\*\\generator Riched20 10.0.19041}\\viewkind4\\uc1 \\pard\\f0\\fs22 Caf\\'e9 cr\\'e8me \\b bold\\b0  text\\par\r\n\\pard\\s1\\qc Title line\\par\r\n\\pard\\i 'quoted' Don't stop\\i0\\par\r\n{\\pntext\\f0\\'B7\\tab}{\\*\\pn\\pnlvlblt\\pnf0\\pnindent0{\\pntxtb\\'B7}}\\fi-360\\li720 Bullet one\\par\r\n{\\pntext\\f0\\'B7\\tab}Bullet two\\par\r\n\\pard{\\pntext\\f0 1.\\tab}{\\*\\pn\\pnlvlbody\\pnf0\\pnindent0\\pnstart1\\pndec{\\pntxta.}}\\fi-360\\li720 Numbered\\par\r\n\\pard Plain again\\par\r\n}";
        let d = parse_rtf(rtf).unwrap();
        let b = &d.blocks;
        assert_eq!(texts(&b[0]), "Café crème bold text");
        assert!(b[0].para_runs().iter().any(|r| r.t == "bold" && r.b && r.size == Some(11.0)));
        assert!(matches!(&b[1], Block::H1 { align: Some(a), .. } if a == "center"));
        assert_eq!(texts(&b[2]), "'quoted' Don't stop");
        assert!(b[2].para_runs()[0].i);
        assert!(matches!(&b[3], Block::List { list, .. } if list == "ul"));
        assert_eq!(texts(&b[3]), "Bullet one");
        assert!(matches!(&b[4], Block::List { list, .. } if list == "ul"));
        assert!(matches!(&b[5], Block::List { list, .. } if list == "ol"));
        assert!(matches!(&b[6], Block::P { .. }));
        assert_eq!(texts(&b[6]), "Plain again");
    }

    #[test]
    fn rtf_word_lists_tables_fields_and_pictures() {
        let png_hex: String = PNG_1X1.iter().map(|b| format!("{b:02x}")).collect();
        let rtf = format!(
            r#"{{\rtf1\ansi\deff0{{\fonttbl{{\f0 Calibri;}}{{\f3\fcharset2 Symbol;}}}}{{\*\listtable{{\list\listid1{{\listlevel\levelnfc23}}}}}}{{\*\listoverridetable{{\listoverride\listid1\ls1}}}}
{{\header\pard Header text\par}}
{{\listtext\pard\plain\f3 \'b7\tab}}\pard\ls1\ilvl0 Item A\par
{{\listtext\pard\plain\f3 \'b7\tab}}\pard\ls1\ilvl1 Item B\par
{{\listtext\pard\plain 1.\tab}}\pard\ls2\ilvl0 Num 1\par
\pard See {{\field{{\*\fldinst{{HYPERLINK "https://example.com/?a=1&b=2"}}}}{{\fldrslt{{\ul\cf1 the site}}}}}} now.\par
\pard Page {{\field{{\*\fldinst PAGE}}{{\fldrslt 3}}}}\par
\trowd\cellx2000\cellx4000\cellx6000
\pard\intbl A1\cell B1\cell C1\cell\row
\trowd\cellx4000\cellx6000
\pard\intbl wide\cell C2\cell\row
\pard After table\par
{{\*\shppict{{\pict\pngblip\picw1\pich1\picwgoal300\pichgoal150 {png_hex}}}}}{{\nonshppict{{\pict\wmetafile8 0102}}}}\par
{{\footnote\pard footnote text\par}}End\par
}}"#
        );
        let d = parse_rtf(rtf.as_bytes()).unwrap();
        let b = &d.blocks;
        let all: String = b.iter().map(texts).collect::<Vec<_>>().join("|");
        assert!(!all.contains("Header text") && !all.contains("footnote text"), "{all}");
        assert!(matches!(&b[0], Block::List { list, level: 0, .. } if list == "ul"));
        assert!(matches!(&b[1], Block::List { list, level: 1, .. } if list == "ul"));
        assert!(matches!(&b[2], Block::List { list, level: 0, .. } if list == "ol"));
        assert_eq!(texts(&b[3]), "See the site now.");
        let link = b[3].para_runs().iter().find(|r| r.t == "the site").unwrap();
        assert_eq!(link.link.as_deref(), Some("https://example.com/?a=1&b=2"));
        assert_eq!(texts(&b[4]), "Page 3");
        match &b[5] {
            Block::Table { rows } => {
                assert_eq!(rows.len(), 2);
                assert_eq!(rows[0].len(), 3);
                assert_eq!(rows[1][0].colspan, Some(2));
                assert_eq!(texts(&rows[1][0].blocks[0]), "wide");
            }
            other => panic!("{other:?}"),
        }
        assert_eq!(texts(&b[6]), "After table");
        assert!(matches!(&b[7], Block::Img { w: Some(20), h: Some(10), .. }), "{:?}", b[7]);
        assert_eq!(b.iter().filter(|x| matches!(x, Block::Img { .. })).count(), 1);
        assert_eq!(texts(b.last().unwrap()), "End");
        let w = warnings(&d);
        assert!(w.iter().any(|x| x.contains("Headers")));
        assert!(w.iter().any(|x| x.contains("Footnotes")));
    }

    #[test]
    fn rtf_round_trip_through_writer() {
        let doc = WordDoc {
            kind: "word".into(),
            meta: Map::new(),
            blocks: vec![
                Block::H2 { runs: vec![Run::text("Überschrift")], align: None },
                Block::P {
                    runs: vec![
                        Run { t: "bold".into(), b: true, ..Default::default() },
                        Run::text(" {braces} \\ "),
                        Run { t: "red".into(), color: Some("#FF0000".into()), hl: Some("#FFFF00".into()), size: Some(14.0), font: Some("Georgia".into()), ..Default::default() },
                        Run { t: "link & more".into(), link: Some("https://e.com/?a=1&b=2".into()), ..Default::default() },
                        Run { t: " see".into(), link: Some("#sec2".into()), ..Default::default() },
                        Run::text(" € 😀"),
                    ],
                    align: Some("right".into()),
                },
                Block::List { list: "ol".into(), level: 0, runs: vec![Run::text("one")] },
                Block::List { list: "ul".into(), level: 1, runs: vec![Run::text("sub")] },
                Block::List { list: "ol".into(), level: 0, runs: vec![Run::text("two")] },
                Block::Table { rows: vec![vec![TableCell { blocks: vec![Block::P { runs: vec![Run::text("span")], align: None }], colspan: Some(2) }], vec![TableCell { blocks: vec![Block::P { runs: vec![Run::text("a")], align: None }], colspan: None }, TableCell { blocks: vec![Block::P { runs: vec![Run::text("b")], align: None }], colspan: None }]] },
                Block::Img { src: data_url("image/png", PNG_1X1), alt: None, w: Some(10), h: Some(10) },
                Block::Quote { runs: vec![Run::text("q")], align: None },
                Block::Code { runs: vec![Run::text("a\nb")], align: None },
            ],
        };
        let bytes = write_rtf(&doc);
        assert!(bytes.is_ascii());
        let back = parse_rtf(&bytes).unwrap();
        let b = &back.blocks;
        assert!(matches!(&b[0], Block::H2 { .. }));
        assert_eq!(texts(&b[0]), "Überschrift");
        assert_eq!(texts(&b[1]), "bold {braces} \\ redlink & more see € 😀");
        assert_eq!(b[1].para_runs().iter().find(|r| r.t == " see").unwrap().link.as_deref(), Some("#sec2"));
        let runs = b[1].para_runs();
        assert!(runs[0].b);
        let red = runs.iter().find(|r| r.t == "red").unwrap();
        assert_eq!(red.color.as_deref(), Some("#FF0000"));
        assert_eq!(red.hl.as_deref(), Some("#FFFF00"));
        assert_eq!(red.size, Some(14.0));
        assert_eq!(red.font.as_deref(), Some("Georgia"));
        assert_eq!(runs.iter().find(|r| r.t == "link & more").unwrap().link.as_deref(), Some("https://e.com/?a=1&b=2"));
        assert!(matches!(&b[1], Block::P { align: Some(a), .. } if a == "right"));
        assert!(matches!(&b[2], Block::List { list, level: 0, .. } if list == "ol"));
        assert!(matches!(&b[3], Block::List { list, level: 1, .. } if list == "ul"));
        assert!(matches!(&b[4], Block::List { list, level: 0, .. } if list == "ol"));
        match &b[5] {
            Block::Table { rows } => {
                assert_eq!(rows[0][0].colspan, Some(2));
                assert_eq!(rows[1].len(), 2);
            }
            other => panic!("{other:?}"),
        }
        assert!(matches!(&b[6], Block::Img { .. }));
        assert!(matches!(&b[7], Block::Quote { .. }));
        assert!(matches!(&b[8], Block::Code { .. }));
        assert_eq!(texts(&b[8]), "a\nb");
    }

    #[test]
    fn header_footer_in_rtf_and_html_writers() {
        let mut meta = Map::new();
        meta.insert("title".into(), json!("T & T"));
        meta.insert("header".into(), json!({"runs": [{"t": "Head {x}", "b": true}], "align": "center"}));
        meta.insert("footer".into(), json!({"runs": [{"t": "Page {PAGE} of {NUMPAGES}"}], "align": "right"}));
        let doc = WordDoc { kind: "word".into(), meta, blocks: vec![Block::P { runs: vec![Run::text("body")], align: None }] };
        let rtf = String::from_utf8(write_rtf(&doc)).unwrap();
        assert!(rtf.contains("{\\header\\pard\\plain\\qc"), "{rtf}");
        assert!(rtf.contains("Head \\{x\\}"), "{rtf}");
        assert!(rtf.contains("{\\footer\\pard\\plain\\qr"), "{rtf}");
        assert!(rtf.contains("{\\chpgn}") && rtf.contains("\\fldinst NUMPAGES"), "{rtf}");
        let back = parse_rtf(rtf.as_bytes()).unwrap();
        assert_eq!(back.blocks.len(), 1, "{:?}", back.blocks);
        assert_eq!(texts(&back.blocks[0]), "body");
        let html = write_html(&doc);
        assert!(html.contains("<title>T &amp; T</title>"));
        assert!(html.contains("<header class=\"doc-header\" style=\"text-align:center\"><strong>Head {x}</strong></header>"), "{html}");
        assert!(html.contains("<footer class=\"doc-footer\" style=\"text-align:right\">Page 1 of 1</footer>"), "{html}");
        assert!(!write_markdown(&doc).contains("Page"));
    }

    #[test]
    fn rtf_misnamed_inputs() {
        let d = parse_rtf(b"plain words\nline two").unwrap();
        assert_eq!(texts(&d.blocks[0]), "plain words");
        assert!(warnings(&d).iter().any(|w| w.contains("plain text")));
        // Deeply nested / unbalanced groups do not crash.
        let mut s = String::from("{\\rtf1 ");
        for _ in 0..20000 {
            s.push('{');
        }
        s.push_str("deep");
        let d = parse_rtf(s.as_bytes()).unwrap();
        assert_eq!(texts(&d.blocks[0]), "deep");
    }

    // ---------- Markdown / text ----------

    #[test]
    fn markdown_blocks_and_inline() {
        let md = "# Title\n\nSome **bold** and *it* and `code` with [a link](https://e.com/?a=1&b=2).\nsoft line\n\n- one\n- two\n  - nested\n    - deeper\n1. first\n2) second\n\n> quoted **text**\n\n```rust\nfn x() {}\n```\n\n| A | B |\n|:-:|--:|\n| 1 | 2 |\n\n---\n\nSetext\n======\n\n- [ ] todo\n- [x] done\n";
        let d = parse_text(md.as_bytes(), true);
        let b = &d.blocks;
        assert!(matches!(&b[0], Block::H1 { .. }));
        let runs = b[1].para_runs();
        assert!(runs.iter().any(|r| r.t == "bold" && r.b));
        assert!(runs.iter().any(|r| r.t == "it" && r.i));
        assert!(runs.iter().any(|r| r.t == "code" && r.font.as_deref() == Some("Consolas")));
        assert!(runs.iter().any(|r| r.t == "a link" && r.link.as_deref() == Some("https://e.com/?a=1&b=2")));
        assert!(texts(&b[1]).ends_with(". soft line"));
        assert!(matches!(&b[2], Block::List { level: 0, .. }));
        assert!(matches!(&b[4], Block::List { level: 1, .. }));
        assert!(matches!(&b[5], Block::List { level: 2, .. }));
        assert!(matches!(&b[6], Block::List { list, level: 0, .. } if list == "ol"));
        assert!(matches!(&b[7], Block::List { list, .. } if list == "ol"));
        assert!(matches!(&b[8], Block::Quote { .. }));
        assert!(b[8].para_runs().iter().any(|r| r.b));
        assert!(matches!(&b[9], Block::Code { .. }));
        assert_eq!(texts(&b[9]), "fn x() {}");
        match &b[10] {
            Block::Table { rows } => {
                assert_eq!(rows.len(), 2);
                assert!(rows[0][0].blocks[0].para_runs()[0].b);
                assert!(matches!(&rows[1][1].blocks[0], Block::P { align: Some(a), .. } if a == "right"));
            }
            other => panic!("{other:?}"),
        }
        assert!(matches!(&b[11], Block::PageBreak));
        assert!(matches!(&b[12], Block::H1 { .. }));
        assert!(texts(&b[13]).starts_with('☐'));
        assert!(texts(&b[14]).starts_with('☑'));
    }

    #[test]
    fn markdown_never_skips_list_levels() {
        let item = |level: u32, t: &str| Block::List { list: "ul".into(), level, runs: vec![Run::text(t)] };
        let doc = WordDoc {
            kind: "word".into(),
            meta: Map::new(),
            blocks: vec![item(2, "a"), item(0, "b"), item(2, "c"), item(1, "d")],
        };
        let md = write_markdown(&doc);
        assert!(md.contains("- a\n- b\n    - c\n    - d"), "{md}");
        let back = parse_text(md.as_bytes(), true);
        let levels: Vec<u32> = back.blocks.iter().filter_map(|b| match b { Block::List { level, .. } => Some(*level), _ => None }).collect();
        assert_eq!(levels, vec![0, 0, 1, 1]);
    }

    #[test]
    fn markdown_round_trip_through_writer() {
        let doc = WordDoc {
            kind: "word".into(),
            meta: Map::new(),
            blocks: vec![
                Block::H1 { runs: vec![Run::text("Report")], align: None },
                Block::P {
                    runs: vec![
                        Run::text("Plain *star* "),
                        Run { t: "bold".into(), b: true, ..Default::default() },
                        Run::text(" "),
                        Run { t: "both ".into(), b: true, i: true, ..Default::default() },
                        Run { t: "link".into(), link: Some("https://e.com/a b".into()), ..Default::default() },
                        Run { t: "x^2".into(), font: Some("Consolas".into()), ..Default::default() },
                    ],
                    align: None,
                },
                Block::P { runs: vec![Run::text("1. not a list")], align: None },
                Block::List { list: "ol".into(), level: 0, runs: vec![Run::text("a")] },
                Block::List { list: "ol".into(), level: 0, runs: vec![Run::text("b")] },
                Block::List { list: "ul".into(), level: 1, runs: vec![Run::text("c")] },
                Block::Quote { runs: vec![Run::text("q1\nq2")], align: None },
                Block::Code { runs: vec![Run::text("let a = `x`;\n```")], align: None },
                Block::Table { rows: vec![vec![TableCell { blocks: vec![Block::P { runs: vec![Run::text("H|1")], align: None }], colspan: Some(2) }], vec![TableCell { blocks: vec![Block::P { runs: vec![Run::text("x")], align: None }], colspan: None }, TableCell { blocks: vec![Block::P { runs: vec![Run::text("y")], align: None }], colspan: None }]] },
                Block::Img { src: data_url("image/png", PNG_1X1), alt: Some("logo".into()), w: None, h: None },
                Block::PageBreak,
            ],
        };
        let md = write_markdown(&doc);
        assert!(md.contains("# Report"));
        assert!(md.contains("Plain \\*star\\* **bold** ***both*** [link](<https://e.com/a b>)`x^2`"), "{md}");
        assert!(md.contains("1\\. not a list"));
        assert!(md.contains("1. a\n2. b\n    - c"), "{md}");
        assert!(md.contains("> q1\\\n> q2"));
        assert!(md.contains("````\nlet a = `x`;\n```\n````"), "{md}");
        assert!(md.contains("| H\\|1 |  |\n| --- | --- |\n| x | y |"), "{md}");
        assert!(md.contains("![logo](data:image/png;base64,"));
        let back = parse_text(md.as_bytes(), true);
        let b = &back.blocks;
        assert!(matches!(&b[0], Block::H1 { .. }));
        assert_eq!(texts(&b[1]), "Plain *star* bold both linkx^2");
        assert!(b[1].para_runs().iter().any(|r| r.t == "link" && r.link.as_deref() == Some("https://e.com/a b")));
        assert_eq!(texts(&b[2]), "1. not a list");
        assert!(matches!(&b[2], Block::P { .. }));
        assert!(matches!(&b[3], Block::List { list, .. } if list == "ol"));
        assert!(matches!(&b[5], Block::List { list, level: 1, .. } if list == "ul"));
        assert!(matches!(&b[6], Block::Quote { .. }));
        assert_eq!(texts(&b[6]), "q1\nq2");
        assert_eq!(texts(&b[7]), "let a = `x`;\n```");
        assert!(matches!(&b[8], Block::Table { .. }));
        assert!(matches!(&b[9], Block::Img { w: Some(1), h: Some(1), .. }));
        assert!(matches!(&b[10], Block::PageBreak));
    }

    #[test]
    fn text_encodings() {
        let d = parse_text(b"hello\n\nworld", false);
        assert_eq!(d.blocks.len(), 3);
        let d = parse_text(b"caf\xe9 cr\xe8me\r\n", false);
        assert_eq!(texts(&d.blocks[0]), "café crème");
        assert_eq!(d.blocks.len(), 1);
        assert!(!warnings(&d).is_empty());
        let mut u16le = vec![0xFF, 0xFE];
        for u in "Grüße\nok".encode_utf16() {
            u16le.extend_from_slice(&u.to_le_bytes());
        }
        let d = parse_text(&u16le, false);
        assert_eq!(texts(&d.blocks[0]), "Grüße");
        let d = parse_text("\u{feff}bom text".as_bytes(), false);
        assert_eq!(texts(&d.blocks[0]), "bom text");
    }

    // ---------- writers ----------

    #[test]
    fn html_writer_is_complete_and_safe() {
        let doc = WordDoc {
            kind: "word".into(),
            meta: [("title".to_string(), json!("My <Doc>"))].into_iter().collect(),
            blocks: vec![
                Block::P { runs: vec![Run { t: "x<y>".into(), b: true, color: Some("#ff0000".into()), ..Default::default() }, Run { t: "bad".into(), link: Some("javascript:alert(1)".into()), ..Default::default() }, Run { t: "good".into(), link: Some("https://e.com/?a=1&b=2".into()), ..Default::default() }], align: Some("center".into()) },
                Block::List { list: "ul".into(), level: 0, runs: vec![Run::text("a")] },
                Block::List { list: "ul".into(), level: 2, runs: vec![Run::text("deep")] },
                Block::List { list: "ol".into(), level: 0, runs: vec![Run::text("n")] },
                Block::Table { rows: vec![vec![TableCell { blocks: vec![Block::P { runs: vec![Run::text("c")], align: None }], colspan: Some(2) }]] },
                Block::Code { runs: vec![Run::text("<b>")], align: None },
                Block::Img { src: "javascript:alert(1)".into(), alt: None, w: None, h: None },
            ],
        };
        let h = write_html(&doc);
        assert!(h.starts_with("<!DOCTYPE html>"));
        assert!(h.contains("<title>My &lt;Doc&gt;</title>"));
        assert!(h.contains("<strong><span style=\"color:#FF0000\">x&lt;y&gt;</span></strong>"));
        assert!(!h.contains("javascript:"));
        assert!(h.contains("href=\"https://e.com/?a=1&amp;b=2\""));
        assert!(h.contains("<td colspan=\"2\">"));
        assert!(h.contains("<pre><code>&lt;b&gt;</code></pre>"));
        assert_eq!(h.matches("<ul>").count(), h.matches("</ul>").count());
        assert_eq!(h.matches("<ol>").count(), h.matches("</ol>").count());
        assert_eq!(h.matches("<li").count(), h.matches("</li>").count(), "{h}");
    }

    // ---------- ODT ----------

    fn odt(content: &str) -> Vec<u8> {
        let c = format!(
            r##"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0">
<office:automatic-styles>
<style:style style:name="T1" style:family="text"><style:text-properties fo:font-weight="bold"/></style:style>
<style:style style:name="T2" style:family="text"><style:text-properties fo:font-style="italic" fo:color="#ff0000"/></style:style>
<style:style style:name="P1" style:family="paragraph" style:parent-style-name="Standard"><style:paragraph-properties fo:text-align="center"/></style:style>
<text:list-style style:name="L1"><text:list-level-style-bullet text:level="1"/><text:list-level-style-number text:level="2" style:num-format="1"/></text:list-style>
<text:list-style style:name="L2"><text:list-level-style-number text:level="1" style:num-format="1"/></text:list-style>
</office:automatic-styles><office:body><office:text>{content}</office:text></office:body></office:document-content>"##
        );
        let styles = r#"<?xml version="1.0"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"><office:styles><style:style style:name="Quotations" style:family="paragraph"/><style:style style:name="Preformatted_20_Text" style:display-name="Preformatted Text" style:family="paragraph"/><style:style style:name="Title" style:family="paragraph"/></office:styles></office:document-styles>"#;
        zip_parts(&[
            ("mimetype", b"application/vnd.oasis.opendocument.text"),
            ("content.xml", c.as_bytes()),
            ("styles.xml", styles.as_bytes()),
            ("Pictures/a.png", PNG_1X1),
        ])
    }

    #[test]
    fn odt_structure_and_formatting() {
        let content = r#"<text:h text:outline-level="2">Section <text:span text:style-name="T1">two</text:span></text:h>
<text:p text:style-name="Title">Doc title</text:p>
<text:p text:style-name="P1">Hello<text:s text:c="4"/>world <text:span text:style-name="T1">bold</text:span> <text:span text:style-name="T2">red it</text:span><text:tab/>tab<text:line-break/>next <text:a xlink:href="https://e.com/?a=1&amp;b=2">link</text:a></text:p>
<text:list text:style-name="L1"><text:list-item><text:p>item one</text:p><text:list><text:list-item><text:p>nested num</text:p></text:list-item></text:list></text:list-item><text:list-item><text:p>item two</text:p></text:list-item></text:list>
<text:list text:style-name="L2"><text:list-item><text:p>numbered</text:p></text:list-item></text:list>
<table:table><table:table-row><table:table-cell table:number-columns-spanned="2"><text:p>c1</text:p></table:table-cell><table:covered-table-cell/></table:table-row><table:table-row><table:table-cell><text:p>a</text:p></table:table-cell><table:table-cell><text:p>b</text:p></table:table-cell></table:table-row></table:table>
<text:p text:style-name="Quotations">A quote</text:p>
<text:p text:style-name="Preformatted_20_Text">code 1</text:p><text:p text:style-name="Preformatted_20_Text">code 2</text:p>
<text:p>Pic <draw:frame svg:width="1in" svg:height="0.5in"><draw:image xlink:href="Pictures/a.png"/><svg:desc>alt text</svg:desc></draw:frame> after</text:p>
<text:p>Note<text:note text:note-class="footnote"><text:note-citation>1</text:note-citation><text:note-body><text:p>The note.</text:p></text:note-body></text:note> end</text:p>"#;
        let d = parse_odt(&odt(content)).unwrap();
        let b = &d.blocks;
        assert!(matches!(&b[0], Block::H2 { .. }));
        assert!(b[0].para_runs().iter().any(|r| r.t == "two" && r.b));
        assert!(matches!(&b[1], Block::H1 { .. }));
        assert!(matches!(&b[2], Block::P { align: Some(a), .. } if a == "center"));
        assert_eq!(texts(&b[2]), "Hello    world bold red it\ttab\nnext link");
        let r = b[2].para_runs();
        assert!(r.iter().any(|x| x.t == "bold" && x.b));
        assert!(r.iter().any(|x| x.t == "red it" && x.i && x.color.as_deref() == Some("#FF0000")));
        assert!(r.iter().any(|x| x.t == "link" && x.link.as_deref() == Some("https://e.com/?a=1&b=2")));
        assert!(matches!(&b[3], Block::List { list, level: 0, .. } if list == "ul"));
        assert!(matches!(&b[4], Block::List { list, level: 1, .. } if list == "ol"));
        assert!(matches!(&b[5], Block::List { list, level: 0, .. } if list == "ul"));
        assert!(matches!(&b[6], Block::List { list, .. } if list == "ol"));
        match &b[7] {
            Block::Table { rows } => {
                assert_eq!(rows[0].len(), 1);
                assert_eq!(rows[0][0].colspan, Some(2));
                assert_eq!(rows[1].len(), 2);
            }
            other => panic!("{other:?}"),
        }
        assert!(matches!(&b[8], Block::Quote { .. }));
        assert!(matches!(&b[9], Block::Code { .. }));
        assert_eq!(texts(&b[9]), "code 1\ncode 2");
        assert_eq!(texts(&b[10]), "Pic ");
        assert!(matches!(&b[11], Block::Img { w: Some(96), h: Some(48), alt: Some(a), .. } if a == "alt text"));
        assert_eq!(texts(&b[12]), " after");
        assert!(b[13].para_runs().iter().any(|x| x.t == "1" && x.sup));
        assert!(texts(b.last().unwrap()).contains("The note."));
    }

    // ---------- .doc ----------

    /// Build a minimal Word 97 .doc (CFB v3) with the given UTF-16 text and
    /// per-paragraph properties (end-cp -> grpprl bytes).
    fn build_doc(text: &str, paras: &[(usize, Vec<u8>, u16)]) -> Vec<u8> {
        let units: Vec<u16> = text.encode_utf16().collect();
        let mut wd = vec![0u8; 0x1000];
        wd[0..2].copy_from_slice(&0xA5ECu16.to_le_bytes());
        wd[2..4].copy_from_slice(&0x00C1u16.to_le_bytes());
        wd[0x0A..0x0C].copy_from_slice(&0x0200u16.to_le_bytes()); // 1Table
        wd[0x4C..0x50].copy_from_slice(&(units.len() as u32).to_le_bytes());
        let text_fc = 0x800u32;
        for (k, u) in units.iter().enumerate() {
            let o = text_fc as usize + k * 2;
            wd[o..o + 2].copy_from_slice(&u.to_le_bytes());
        }
        // FKP page at 0xC00 (page number 6).
        let page_off = 0xC00usize;
        let crun = paras.len();
        let mut page = vec![0u8; 512];
        let mut fc_prev = text_fc;
        let mut papx_pos = 511usize;
        for (i, (end_cp, grpprl, istd)) in paras.iter().enumerate() {
            let fc_end = text_fc + (*end_cp as u32 + 1) * 2;
            page[i * 4..i * 4 + 4].copy_from_slice(&fc_prev.to_le_bytes());
            page[(i + 1) * 4..(i + 1) * 4 + 4].copy_from_slice(&fc_end.to_le_bytes());
            fc_prev = fc_end;
            let mut body = istd.to_le_bytes().to_vec();
            body.extend_from_slice(grpprl);
            if body.len() % 2 == 1 {
                body.push(0);
            }
            let total = 2 + body.len();
            papx_pos = (papx_pos - total) & !1;
            page[papx_pos] = 0;
            page[papx_pos + 1] = (body.len() / 2) as u8;
            page[papx_pos + 2..papx_pos + 2 + body.len()].copy_from_slice(&body);
            page[(crun + 1) * 4 + i * 13] = (papx_pos / 2) as u8;
        }
        page[511] = crun as u8;
        wd[page_off..page_off + 512].copy_from_slice(&page);
        // Table stream: Clx at 0, PlcBtePapx after.
        let mut tbl = Vec::new();
        tbl.push(2u8);
        let plc_len = 4 * 2 + 8;
        tbl.extend_from_slice(&(plc_len as u32).to_le_bytes());
        tbl.extend_from_slice(&0u32.to_le_bytes());
        tbl.extend_from_slice(&(units.len() as u32).to_le_bytes());
        tbl.extend_from_slice(&0u16.to_le_bytes());
        tbl.extend_from_slice(&text_fc.to_le_bytes());
        tbl.extend_from_slice(&0u16.to_le_bytes());
        let clx_len = tbl.len();
        let bte_off = tbl.len();
        tbl.extend_from_slice(&text_fc.to_le_bytes());
        tbl.extend_from_slice(&fc_prev.to_le_bytes());
        tbl.extend_from_slice(&((page_off / 512) as u32).to_le_bytes());
        let bte_len = tbl.len() - bte_off;
        wd[0x1A2..0x1A6].copy_from_slice(&0u32.to_le_bytes());
        wd[0x1A6..0x1AA].copy_from_slice(&(clx_len as u32).to_le_bytes());
        wd[0x102..0x106].copy_from_slice(&(bte_off as u32).to_le_bytes());
        wd[0x106..0x10A].copy_from_slice(&(bte_len as u32).to_le_bytes());
        tbl.resize(0x1000, 0);
        // CFB: header, FAT sector 0, directory sector 1, WordDocument sectors 2..9, 1Table 10..17.
        let mut f = vec![0u8; 512];
        f[..8].copy_from_slice(&[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);
        f[0x18..0x1A].copy_from_slice(&0x3Eu16.to_le_bytes());
        f[0x1A..0x1C].copy_from_slice(&3u16.to_le_bytes());
        f[0x1C..0x1E].copy_from_slice(&0xFFFEu16.to_le_bytes());
        f[0x1E..0x20].copy_from_slice(&9u16.to_le_bytes());
        f[0x20..0x22].copy_from_slice(&6u16.to_le_bytes());
        f[0x2C..0x30].copy_from_slice(&1u32.to_le_bytes());
        f[0x30..0x34].copy_from_slice(&1u32.to_le_bytes());
        f[0x38..0x3C].copy_from_slice(&4096u32.to_le_bytes());
        f[0x3C..0x40].copy_from_slice(&0xFFFF_FFFEu32.to_le_bytes());
        f[0x44..0x48].copy_from_slice(&0xFFFF_FFFEu32.to_le_bytes());
        for i in 0..109 {
            f[0x4C + i * 4..0x50 + i * 4].copy_from_slice(&0xFFFF_FFFFu32.to_le_bytes());
        }
        f[0x4C..0x50].copy_from_slice(&0u32.to_le_bytes());
        let mut fat = vec![0xFFFF_FFFFu32; 128];
        fat[0] = 0xFFFF_FFFD;
        fat[1] = 0xFFFF_FFFE;
        for s in 2..10 {
            fat[s] = if s == 9 { 0xFFFF_FFFE } else { s as u32 + 1 };
        }
        for s in 10..18 {
            fat[s] = if s == 17 { 0xFFFF_FFFE } else { s as u32 + 1 };
        }
        for v in fat {
            f.extend_from_slice(&v.to_le_bytes());
        }
        let mut dir = vec![0u8; 512];
        let entry = |dir: &mut Vec<u8>, idx: usize, name: &str, kind: u8, start: u32, size: u32, child: u32, right: u32| {
            let o = idx * 128;
            let units: Vec<u16> = name.encode_utf16().collect();
            for (k, u) in units.iter().enumerate() {
                dir[o + k * 2..o + k * 2 + 2].copy_from_slice(&u.to_le_bytes());
            }
            dir[o + 64..o + 66].copy_from_slice(&(((units.len() + 1) * 2) as u16).to_le_bytes());
            dir[o + 66] = kind;
            dir[o + 68..o + 72].copy_from_slice(&0xFFFF_FFFFu32.to_le_bytes());
            dir[o + 72..o + 76].copy_from_slice(&right.to_le_bytes());
            dir[o + 76..o + 80].copy_from_slice(&child.to_le_bytes());
            dir[o + 116..o + 120].copy_from_slice(&start.to_le_bytes());
            dir[o + 120..o + 124].copy_from_slice(&size.to_le_bytes());
        };
        entry(&mut dir, 0, "Root Entry", 5, 0xFFFF_FFFE, 0, 1, 0xFFFF_FFFF);
        entry(&mut dir, 1, "WordDocument", 2, 2, 0x1000, 0xFFFF_FFFF, 2);
        entry(&mut dir, 2, "1Table", 2, 10, 0x1000, 0xFFFF_FFFF, 0xFFFF_FFFF);
        f.extend_from_slice(&dir);
        f.extend_from_slice(&wd);
        f.extend_from_slice(&tbl);
        f
    }

    #[test]
    fn legacy_doc_paragraphs_headings_tables_and_fields() {
        // Paragraphs: heading, body with HYPERLINK field, a 2x1 table row, after.
        let text = "Heading\rSee \u{13} HYPERLINK \"https://e.com/?a=1&b=2\" \u{14}site\u{15} and \u{13}PAGE\u{14}4\u{15}.\rA\u{7}B\u{7}\u{7}After\r";
        let pos = |needle: &str, nth: usize| -> usize {
            let u: Vec<char> = text.chars().collect();
            let mut found = 0;
            for (i, c) in u.iter().enumerate() {
                if c.to_string() == needle {
                    if found == nth {
                        return i;
                    }
                    found += 1;
                }
            }
            panic!("not found")
        };
        let in_table = vec![0x16, 0x24, 1];
        let ttp = vec![0x16, 0x24, 1, 0x17, 0x24, 1];
        let paras = vec![
            (pos("\r", 0), vec![], 1u16),
            (pos("\r", 1), vec![], 0),
            (pos("\u{7}", 0), in_table.clone(), 0),
            (pos("\u{7}", 1), in_table.clone(), 0),
            (pos("\u{7}", 2), ttp, 0),
            (pos("\r", 2), vec![], 0),
        ];
        let bytes = build_doc(text, &paras);
        assert!(!ole_is_encrypted(&bytes));
        let d = parse_doc(&bytes);
        let b = &d.blocks;
        assert!(matches!(&b[0], Block::H1 { .. }), "{b:?}");
        assert_eq!(texts(&b[0]), "Heading");
        assert_eq!(texts(&b[1]), "See site and 4.");
        assert!(b[1].para_runs().iter().any(|r| r.t == "site" && r.link.as_deref() == Some("https://e.com/?a=1&b=2")));
        match &b[2] {
            Block::Table { rows } => {
                assert_eq!(rows.len(), 1);
                assert_eq!(rows[0].len(), 2);
                assert_eq!(texts(&rows[0][1].blocks[0]), "B");
            }
            other => panic!("{other:?}"),
        }
        assert_eq!(texts(&b[3]), "After");
        assert!(warnings(&d).iter().any(|w| w.contains("Legacy .doc")));
        // Misnamed files are routed.
        let rtf = parse_doc(br"{\rtf1 Hi\par}");
        assert_eq!(texts(&rtf.blocks[0]), "Hi");
        let txt = parse_doc(b"just text");
        assert_eq!(texts(&txt.blocks[0]), "just text");
        // Garbage OLE falls back to scraping with a warning.
        let mut junk = bytes.clone();
        junk[0x600] = 0; // break wIdent
        junk[0x601] = 0;
        let d = parse_doc(&junk);
        assert!(warnings(&d).iter().any(|w| w.contains("could not be fully decoded")));
    }
}
