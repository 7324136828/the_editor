//! Format detection: filename extension + content sniffing.
//!
//! The extension is only a hint. The bytes decide whenever they are
//! unambiguous, so commonly misnamed files are routed to the right parser:
//! * ZIP containers are told apart by `[Content_Types].xml` (OOXML) or the
//!   `mimetype` entry (OpenDocument): a .docx renamed .doc opens as docx.
//! * OLE2 compound files are told apart by their stream names
//!   (`WordDocument` = .doc, `Workbook`/`Book` = .xls; `EncryptedPackage` =
//!   password-protected OOXML).
//! * `{\rtf` is RTF whatever the name (Word happily saves RTF as .doc).
//! * HTML / SpreadsheetML-2003 "Excel exports" saved as .xls are reported
//!   with a clear error; delimited text saved as .xls opens as CSV.
//!
//! Mirrors (and extends) lib/formats.js in the extension.

use std::io::{Cursor, Read};

pub fn ext_of(name: &str) -> String {
    // Only the last path component counts ("a.b/c" has no extension).
    let base = name.rsplit(['/', '\\']).next().unwrap_or(name);
    match base.rsplit_once('.') {
        Some((stem, ext)) if !stem.is_empty() || !ext.is_empty() => ext.trim().to_ascii_lowercase(),
        _ => String::new(),
    }
}

pub fn format_of_name(name: &str) -> Option<&'static str> {
    Some(match ext_of(name).as_str() {
        "docx" => "docx", "doc" => "doc", "rtf" => "rtf", "odt" => "odt",
        "txt" => "txt", "md" => "md",
        "xlsx" => "xlsx", "xls" => "xls", "xlsm" => "xlsm",
        "csv" => "csv", "tsv" => "tsv", "ods" => "ods",
        "pptx" => "pptx",
        "tex" | "ltx" | "latex" => "tex",
        _ => return None,
    })
}

/// "word" | "excel" | None
pub fn kind_of_format(fmt: &str) -> Option<&'static str> {
    match fmt {
        "docx" | "doc" | "rtf" | "odt" | "txt" | "md" => Some("word"),
        "xlsx" | "xls" | "xlsm" | "csv" | "tsv" | "ods" => Some("excel"),
        "pptx" => Some("slides"),
        "tex" => Some("latex"),
        _ => None,
    }
}

/// Every format /api/parse accepts (advertised by /health).
pub const PARSE_FORMATS: &[&str] = &[
    "docx", "doc", "rtf", "odt", "txt", "md", "xlsx", "xlsm", "xls", "ods", "csv", "tsv",
    "pptx", "tex",
];

/// Human-readable description used in messages.
pub fn describe(fmt: &str) -> &'static str {
    match fmt {
        "docx" => "a Word document (.docx)",
        "doc" => "a legacy Word 97-2003 document (.doc)",
        "rtf" => "a Rich Text document (.rtf)",
        "odt" => "an OpenDocument text file (.odt)",
        "txt" => "a plain-text file",
        "md" => "a Markdown file",
        "xlsx" => "an Excel workbook (.xlsx)",
        "xlsm" => "a macro-enabled Excel workbook (.xlsm)",
        "xls" => "a legacy Excel 97-2003 workbook (.xls)",
        "ods" => "an OpenDocument spreadsheet (.ods)",
        "csv" => "delimited text (CSV)",
        "tsv" => "tab-separated text",
        "pptx" => "a PowerPoint presentation (.pptx)",
        "tex" => "a LaTeX document (.tex)",
        _ => "an unknown file type",
    }
}

/// Cheap magic-byte classification: "ole2" | "zip" | "rtf" | None.
pub fn sniff(bytes: &[u8]) -> Option<&'static str> {
    if bytes.len() < 4 { return None; }
    if bytes.starts_with(&[0xd0, 0xcf, 0x11, 0xe0]) { return Some("ole2"); }
    if bytes.starts_with(b"PK\x03\x04") || bytes.starts_with(b"PK\x05\x06") || bytes.starts_with(b"PK\x07\x08") {
        return Some("zip");
    }
    if bytes.starts_with(b"{\\rt") { return Some("rtf"); }
    None
}

/// Result of a successful detection.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Detected {
    /// One of PARSE_FORMATS.
    pub format: &'static str,
    /// Set when the content overrode the file extension (shown to the user).
    pub note: Option<String>,
}

/// Why detection failed. `damaged` separates "this is a broken file of a
/// supported type" (HTTP 422) from "this type is not supported" (HTTP 415).
/// Derefs to the user-facing message.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DetectError {
    pub message: String,
    pub damaged: bool,
}

impl DetectError {
    fn unsupported(m: impl Into<String>) -> Self { DetectError { message: m.into(), damaged: false } }
    fn damaged(m: impl Into<String>) -> Self { DetectError { message: m.into(), damaged: true } }
}

impl std::ops::Deref for DetectError {
    type Target = str;
    fn deref(&self) -> &str { &self.message }
}

impl std::fmt::Display for DetectError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result { f.write_str(&self.message) }
}

/// What a ZIP / OLE2 container really holds.
#[derive(Debug, Clone, PartialEq, Eq)]
enum Content {
    Format(&'static str),
    /// Recognised but unsupported: the message explains why.
    Unsupported(String),
    /// A container of a supported kind that cannot be read.
    Damaged(String),
}

/// Resolve the effective format of an upload. `Err` carries a user-facing
/// explanation (unsupported / unrecognisable file).
pub fn detect(name: &str, bytes: &[u8]) -> Result<Detected, DetectError> {
    let by_name = format_of_name(name);
    let ext = ext_of(name);
    let content: Option<Content> = match sniff(bytes) {
        Some("zip") => Some(zip_content(bytes)),
        Some("ole2") => Some(ole_content(bytes)),
        Some("rtf") => Some(Content::Format("rtf")),
        _ => text_content(bytes, by_name),
    };
    match content {
        Some(Content::Unsupported(msg)) => Err(DetectError::unsupported(msg)),
        Some(Content::Damaged(msg)) => Err(DetectError::damaged(msg)),
        Some(Content::Format(actual)) => {
            let note = match by_name {
                // .xlsm vs .xlsx are the same container; don't nag about it.
                Some(n) if n == actual || matches!((n, actual), ("xlsm", "xlsx") | ("xlsx", "xlsm")) => None,
                Some(_) => Some(format!(
                    "The file is named .{ext} but is actually {}; it was opened as .{actual}.",
                    describe(actual))),
                None => None,
            };
            // Keep the user's extension for the same-container xlsx/xlsm case.
            let format = match (by_name, actual) {
                (Some("xlsm"), "xlsx") => "xlsm",
                (Some("xlsx"), "xlsm") => "xlsx",
                _ => actual,
            };
            Ok(Detected { format, note })
        }
        None => match by_name {
            // Binary formats must have recognisable content.
            Some(f @ ("docx" | "xlsx" | "xlsm" | "odt" | "ods" | "pptx")) => Err(DetectError::damaged(format!(
                "This is not a valid {}: it may be damaged or incomplete, or be a different kind of file renamed to .{ext}.",
                describe(f).trim_start_matches("an ").trim_start_matches("a ")))),
            Some("doc") => Err(DetectError::damaged("This .doc file is not a Word document (unrecognised content). It may be damaged.")),
            Some("xls") => Err(DetectError::damaged("This .xls file is not an Excel workbook (unrecognised content). It may be damaged.")),
            Some(f) => Ok(Detected { format: f, note: None }),
            None if ext.is_empty() => Err(DetectError::unsupported("Unsupported file type: the file has no extension and its content was not recognised.")),
            None => Err(DetectError::unsupported(format!("Unsupported file type: .{ext}"))),
        },
    }
}

/// Backwards-compatible wrapper: effective format or None.
#[allow(dead_code)] // public API (tests, other callers)
pub fn resolve_format(name: &str, bytes: &[u8]) -> Option<String> {
    detect(name, bytes).ok().map(|d| d.format.to_string())
}

// ---------------- ZIP (OOXML / OpenDocument) ----------------

fn read_entry_prefix(ar: &mut zip::ZipArchive<Cursor<&[u8]>>, name: &str, max: u64) -> Option<String> {
    let f = ar.by_name(name).ok()?;
    let mut s = Vec::new();
    f.take(max).read_to_end(&mut s).ok()?;
    Some(String::from_utf8_lossy(&s).into_owned())
}

fn zip_content(bytes: &[u8]) -> Content {
    let Ok(mut ar) = zip::ZipArchive::new(Cursor::new(bytes)) else {
        return Content::Damaged("The file looks like a ZIP package but it is damaged or incomplete and cannot be opened.".into());
    };
    // OpenDocument: "mimetype" entry (stored first, uncompressed).
    if let Some(m) = read_entry_prefix(&mut ar, "mimetype", 200) {
        let m = m.trim();
        if m == "application/vnd.oasis.opendocument.text" || m == "application/vnd.oasis.opendocument.text-template" {
            return Content::Format("odt");
        }
        if m == "application/vnd.oasis.opendocument.spreadsheet" || m == "application/vnd.oasis.opendocument.spreadsheet-template" {
            return Content::Format("ods");
        }
        if m.starts_with("application/vnd.oasis.opendocument.") {
            return Content::Unsupported(format!(
                "This OpenDocument file ({}) is not a text document or spreadsheet and cannot be opened.",
                m.trim_start_matches("application/vnd.oasis.opendocument.")));
        }
    }
    // OOXML: main part content type in [Content_Types].xml.
    if let Some(ct) = read_entry_prefix(&mut ar, "[Content_Types].xml", 1 << 20) {
        let ct = ct.to_ascii_lowercase();
        if ct.contains("wordprocessingml.document.main+xml") || ct.contains("wordprocessingml.template.main+xml")
            || ct.contains("ms-word.document.macroenabled.main+xml") || ct.contains("ms-word.template.macroenabledtemplate.main+xml") {
            return Content::Format("docx");
        }
        if ct.contains("ms-excel.sheet.macroenabled.main+xml") || ct.contains("ms-excel.template.macroenabled.main+xml") {
            return Content::Format("xlsm");
        }
        if ct.contains("spreadsheetml.sheet.main+xml") || ct.contains("spreadsheetml.template.main+xml") {
            return Content::Format("xlsx");
        }
        if ct.contains("ms-excel.sheet.binary.macroenabled.main") {
            return Content::Unsupported("Excel binary workbooks (.xlsb) are not supported. Save it as .xlsx in Excel and try again.".into());
        }
        if ct.contains("presentationml.presentation.main+xml") || ct.contains("presentationml.template.main+xml")
            || ct.contains("ms-powerpoint.presentation.macroenabled.main+xml") || ct.contains("presentationml.slideshow.main+xml") {
            return Content::Format("pptx");
        }
    }
    // Fallback on well-known part names (packages with a sloppy content-types part).
    let names: Vec<String> = ar.file_names().filter_map(|r| r.ok()).map(|c| c.into_owned()).collect();
    if names.iter().any(|n| n.eq_ignore_ascii_case("word/document.xml")) { return Content::Format("docx"); }
    if names.iter().any(|n| n.eq_ignore_ascii_case("xl/workbook.xml")) { return Content::Format("xlsx"); }
    if names.iter().any(|n| n.eq_ignore_ascii_case("xl/workbook.bin")) {
        return Content::Unsupported("Excel binary workbooks (.xlsb) are not supported. Save it as .xlsx in Excel and try again.".into());
    }
    if names.iter().any(|n| n.eq_ignore_ascii_case("ppt/presentation.xml")) { return Content::Format("pptx"); }
    if names.iter().any(|n| n.eq_ignore_ascii_case("content.xml")) && names.iter().any(|n| n.eq_ignore_ascii_case("meta.xml") || n.eq_ignore_ascii_case("styles.xml")) {
        // OpenDocument without a mimetype entry: peek at the body element.
        if let Some(c) = read_entry_prefix(&mut ar, "content.xml", 1 << 20) {
            if c.contains("<office:text") { return Content::Format("odt"); }
            if c.contains("<office:spreadsheet") { return Content::Format("ods"); }
        }
    }
    Content::Unsupported("This is a ZIP archive, not a Word, Excel or PowerPoint document.".into())
}

// ---------------- OLE2 compound files (.doc / .xls) ----------------

/// Does the compound file contain a directory entry (stream or storage)
/// called `name`? Directory entries are 128-byte records that always start
/// at offsets that are multiples of 128 (header = 512 bytes, sectors are 512
/// or 4096 bytes), with the UTF-16LE name at +0, its byte length incl. NUL
/// at +64 (u16) and the object type at +66 (1 storage, 2 stream, 5 root).
/// Scanning aligned records is robust and needs no FAT walk.
fn ole_has_entry(bytes: &[u8], name: &str) -> bool {
    let want: Vec<u16> = name.encode_utf16().collect();
    let want_len = ((want.len() + 1) * 2) as u16;
    let mut off = 512;
    while off + 128 <= bytes.len() {
        let rec = &bytes[off..off + 128];
        let len = u16::from_le_bytes([rec[64], rec[65]]);
        if len == want_len && matches!(rec[66], 1 | 2) {
            let matches = want.iter().enumerate().all(|(i, &u)| {
                let got = u16::from_le_bytes([rec[i * 2], rec[i * 2 + 1]]);
                // Names compare case-insensitively in CFB.
                char::from_u32(got as u32).map(|c| c.to_ascii_lowercase())
                    == char::from_u32(u as u32).map(|c| c.to_ascii_lowercase())
            });
            if matches { return true; }
        }
        off += 128;
    }
    false
}

fn ole_content(bytes: &[u8]) -> Content {
    if ole_has_entry(bytes, "EncryptedPackage") || ole_has_entry(bytes, "EncryptionInfo") {
        return Content::Unsupported("This document is password-protected (encrypted). Remove the password in Word/Excel and try again.".into());
    }
    if ole_has_entry(bytes, "WordDocument") { return Content::Format("doc"); }
    if ole_has_entry(bytes, "Workbook") || ole_has_entry(bytes, "Book") { return Content::Format("xls"); }
    if ole_has_entry(bytes, "PowerPoint Document") {
        return Content::Unsupported("PowerPoint presentations are not supported (only Word and Excel documents).".into());
    }
    Content::Unsupported("This is a Microsoft compound file that is neither a Word (.doc) nor an Excel (.xls) document.".into())
}

// ---------------- text-ish content ----------------

/// Classify non-container content. Only overrides the extension where the
/// extension is a binary format (.doc/.xls/...) and the bytes are clearly
/// something else; text extensions (.txt/.md/.csv/.tsv) are trusted.
fn text_content(bytes: &[u8], by_name: Option<&str>) -> Option<Content> {
    let head = &bytes[..bytes.len().min(8192)];
    let lower = String::from_utf8_lossy(head).trim_start_matches('\u{feff}').trim_start().to_ascii_lowercase();
    let is_html = lower.starts_with("<!doctype html") || lower.starts_with("<html") || lower.starts_with("<table")
        || (lower.starts_with("<?xml") || lower.starts_with("<")) && (lower.contains("<html") || lower.contains("urn:schemas-microsoft-com:office:excel"));
    let is_spreadsheetml = lower.starts_with("<?xml") && lower.contains("urn:schemas-microsoft-com:office:spreadsheet");
    match by_name {
        Some("xls" | "xlsx" | "xlsm" | "ods") => {
            if is_spreadsheetml {
                return Some(Content::Unsupported(
                    "This file is an \"XML Spreadsheet 2003\" saved with an Excel extension. Open it in Excel or LibreOffice and save it as .xlsx, then try again.".into()));
            }
            if is_html {
                return Some(Content::Unsupported(
                    "This file is a web page (HTML table export) saved with an Excel extension, not a real workbook. Open it in Excel or LibreOffice and save it as .xlsx, then try again.".into()));
            }
            // "Export to Excel" text files are very commonly named .xls (and
            // Excel's "Unicode Text" is UTF-16 TSV). The newer containers
            // are never legitimately plain text: report them as damaged.
            if by_name == Some("xls") && looks_like_text(bytes) { return Some(Content::Format("csv")); }
            None
        }
        Some("doc" | "docx" | "odt") => {
            if is_html {
                return Some(Content::Unsupported(
                    "This file is a web page (HTML) saved with a Word extension, not a real Word document. Open it in Word or LibreOffice and save it as .docx, then try again.".into()));
            }
            // Plain text renamed .doc is common; .docx/.odt never are.
            if by_name == Some("doc") && looks_like_text(bytes) { return Some(Content::Format("txt")); }
            None
        }
        _ => None,
    }
}

/// Heuristic: decodable text with no binary control bytes. UTF-16 with a
/// BOM counts as text.
pub fn looks_like_text(bytes: &[u8]) -> bool {
    // An empty or truncated "PK" file is a broken ZIP, not text.
    if bytes.is_empty() || bytes.starts_with(b"PK") { return false; }
    if bytes.starts_with(&[0xFF, 0xFE]) || bytes.starts_with(&[0xFE, 0xFF]) { return true; }
    let head = &bytes[..bytes.len().min(8192)];
    let bad = head.iter().filter(|&&b| b < 0x20 && !matches!(b, b'\t' | b'\n' | b'\r' | 0x0c)).count();
    bad == 0
}

#[cfg(test)]
/// Minimal fake compound file: 512-byte header + one directory sector
/// holding root + one named stream.
pub(crate) fn fake_ole(stream: &str) -> Vec<u8> {
    let mut b = vec![0u8; 512 + 512];
    b[..8].copy_from_slice(&[0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
    let put = |b: &mut Vec<u8>, off: usize, name: &str, ty: u8| {
        let u: Vec<u16> = name.encode_utf16().collect();
        for (i, c) in u.iter().enumerate() { b[off + i * 2..off + i * 2 + 2].copy_from_slice(&c.to_le_bytes()); }
        b[off + 64..off + 66].copy_from_slice(&(((u.len() + 1) * 2) as u16).to_le_bytes());
        b[off + 66] = ty;
    };
    put(&mut b, 512, "Root Entry", 5);
    put(&mut b, 640, stream, 2);
    b
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Write;

    fn zip_with(entries: &[(&str, &str)]) -> Vec<u8> {
        let mut w = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let opts = zip::write::SimpleFileOptions::default();
        for (n, d) in entries {
            w.start_file(*n, opts).unwrap();
            w.write_all(d.as_bytes()).unwrap();
        }
        w.finish().unwrap().into_inner()
    }

    #[test]
    fn extension_helpers() {
        assert_eq!(ext_of("a.DOCX"), "docx");
        assert_eq!(ext_of("dir.v2/file"), "");
        assert_eq!(ext_of("noext"), "");
        assert_eq!(format_of_name("C:\\x\\Report.Final.XLSX"), Some("xlsx"));
    }

    #[test]
    fn zip_containers_by_content() {
        let docx = zip_with(&[("[Content_Types].xml", "<Override ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>"), ("word/document.xml", "<w:document/>")]);
        let xlsx = zip_with(&[("[Content_Types].xml", "<Override ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"), ("xl/workbook.xml", "<workbook/>")]);
        let odt = zip_with(&[("mimetype", "application/vnd.oasis.opendocument.text"), ("content.xml", "<x/>")]);
        let ods = zip_with(&[("mimetype", "application/vnd.oasis.opendocument.spreadsheet"), ("content.xml", "<x/>")]);
        let bare = zip_with(&[("word/document.xml", "<w:document/>")]);
        assert_eq!(detect("a.docx", &docx).unwrap(), Detected { format: "docx", note: None });
        let d = detect("renamed.doc", &docx).unwrap();
        assert_eq!(d.format, "docx");
        assert!(d.note.unwrap().contains(".doc"));
        assert_eq!(detect("x.xls", &xlsx).unwrap().format, "xlsx");
        assert_eq!(detect("x.xlsm", &xlsx).unwrap(), Detected { format: "xlsm", note: None });
        assert_eq!(detect("x.zip", &odt).unwrap().format, "odt");
        assert_eq!(detect("document", &ods).unwrap().format, "ods");
        assert_eq!(detect("x.docx", &bare).unwrap().format, "docx");
        assert!(detect("x.docx", &zip_with(&[("readme.txt", "hi")])).unwrap_err().contains("ZIP archive"));
        // Word file sent as Excel: kind changes.
        assert_eq!(kind_of_format(detect("book.xlsx", &docx).unwrap().format), Some("word"));
    }

    #[test]
    fn ole_by_stream_names() {
        assert_eq!(detect("a.doc", &fake_ole("WordDocument")).unwrap().format, "doc");
        assert_eq!(detect("a.xls", &fake_ole("Workbook")).unwrap().format, "xls");
        assert_eq!(detect("a.xls", &fake_ole("Book")).unwrap().format, "xls");
        let d = detect("a.xls", &fake_ole("WordDocument")).unwrap();
        assert_eq!(d.format, "doc");
        assert!(d.note.is_some());
        assert!(detect("a.docx", &fake_ole("EncryptedPackage")).unwrap_err().contains("password"));
        assert!(detect("a.doc", &fake_ole("Something")).is_err());
    }

    #[test]
    fn text_and_html_disguises() {
        assert_eq!(detect("x.doc", b"{\\rtf1 hi}").unwrap().format, "rtf");
        assert_eq!(detect("x.bin", b"{\\rtf1 hi}").unwrap().format, "rtf");
        assert!(detect("export.xls", b"<html><body><table><tr><td>1</td></tr></table></body></html>").unwrap_err().contains("HTML"));
        assert!(detect("export.xls", b"<?xml version=\"1.0\"?><Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\">").unwrap_err().contains("2003"));
        let d = detect("export.xls", b"a,b\n1,2\n").unwrap();
        assert_eq!(d.format, "csv");
        assert!(d.note.is_some());
        assert_eq!(detect("x.csv", b"a,b").unwrap().format, "csv");
        assert_eq!(detect("notes.txt", b"<html>not parsed as html</html>").unwrap().format, "txt");
        assert!(detect("x.bin", &[0u8, 1, 2]).is_err());
        assert!(detect("x.doc", &[0u8, 1, 2, 3, 4]).is_err());
        assert!(detect("noext", b"hello").is_err());
        assert_eq!(detect("old.doc", b"just text").unwrap().format, "txt");
        assert!(detect("x.docx", b"PK").is_err());
        // damaged files of a supported type vs. unsupported types
        assert!(detect("x.docx", b" garbage-bytes").unwrap_err().damaged);
        // The message is one plain sentence a person can act on (shown as-is by the viewer).
        let msg = detect("x.docx", b" garbage").unwrap_err().message;
        assert!(msg.starts_with("This is not a valid Word document"), "{msg}");
        assert!(msg.contains("damaged") && msg.contains(".docx") && !msg.contains(") ("), "{msg}");
        assert!(!detect("x.bin", &[0u8, 1, 2]).unwrap_err().damaged);
        assert!(!detect("export.xls", b"<html><body></body></html>").unwrap_err().damaged);
        assert!(detect("x.docx", b"plain text").is_err());
        assert!(detect("x.xlsx", b"a,b").is_err());
        assert_eq!(detect("u.xls", &[0xff, 0xfe, b'a', 0, 9, 0, b'b', 0]).unwrap().format, "csv");
    }
}
