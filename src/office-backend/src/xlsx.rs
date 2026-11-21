//! Spreadsheet formats: .xlsx/.xls/.xlsm/.ods via calamine (read) and
//! rust_xlsxwriter (write), plus a CSV/TSV codec. Produces / consumes the
//! shared ExcelBook model (see the Excel CONTRACT in model.rs).
//!
//! Reading (`parse_book`)
//! - Cell positions are ABSOLUTE: calamine ranges start at the first used
//!   cell, so every relative (row, col) from `used_cells()` is offset by
//!   `range.start()`. The formula range has its own, independent start.
//! - Formulas come from `worksheet_formula` and are merged into the value
//!   cells as `f: "=..."`, keeping the cached result in `v`.
//! - Dates use calamine's own serial conversion (handles the 1900 leap-year
//!   bug and the 1904 date system); out-of-range serials stay numbers.
//! - Styles / number formats come from calamine's `worksheet_style` (xlsx);
//!   column widths, the active tab, frozen panes (`<pane state="frozen">`),
//!   the sheet auto filter, embedded charts and fidelity warnings come from a
//!   light read of the OOXML package (zip + quick-xml).
//! - Charts (xlsx/xlsm): worksheet -> drawing part -> chart parts. Anchors
//!   (two-cell / one-cell / absolute) are converted to grid pixels with the
//!   CANONICAL geometry shared with lib/model-excel.js (see `ColMetrics`).
//!   The chart type is the first chart group of the plot area (column / bar /
//!   line / pie / scatter / area; 3-D and doughnut variants map to their 2-D
//!   kind), the title comes from rich text or a cached string reference, and
//!   the data range is the bounding box of every series reference (names,
//!   categories / x values, values), which must all be on the host sheet.
//!   Charts that cannot be represented are counted in ONE warning; pictures
//!   and shapes are reported as unsupported features. Everything is bounded
//!   (part sizes, MAX_CHARTS, XML nesting depth) and never panics.
//! - Hidden and empty sheets are kept; anything that cannot be represented in
//!   the model is reported in `book.warnings`.
//!
//! Writing (`write_xlsx`)
//! - Every sheet is written (hidden ones with `set_hidden`), the active sheet
//!   is preserved, sheet names are sanitized to Excel's rules, formulas are
//!   written with their cached results, dates as real Excel dates, and
//!   z / s / merges / column widths / frozen panes / the auto filter are
//!   mapped to rust_xlsxwriter.
//! - Model charts become native Excel charts (header row = series names,
//!   first column = categories / x values, one series per further column),
//!   placed with the same canonical geometry. Invalid chart entries, frozen
//!   panes or filter ranges are skipped WITH a warning; they never fail a save.
//! - Per-cell write errors are reported, never swallowed.
//!
//! CSV (`parse_csv_bytes`, `parse_delimited`, `sheet_to_delimited`)
//! - RFC 4180 parser (quoted fields with embedded quotes, delimiters and
//!   newlines; a quote only opens a quoted field at the start of a field).
//! - Encoding detection (UTF-8 ± BOM, UTF-16 LE/BE, Windows-1252 fallback)
//!   and delimiter sniffing (`,` `;` tab `|`, plus Excel's `sep=` line).
//! - Values stay text unless they are plain decimal numbers without leading
//!   zeros (so ZIP codes, IDs, "3%", dates and long digit strings round-trip
//!   byte-for-byte).

use crate::model::{Cell, CellStyle, ExcelBook, Freeze, Sheet};
use calamine::{
    open_workbook_auto_from_rs, Data, ExcelDateTime, ExcelDateTimeType, FillPattern,
    HorizontalAlignment, Range, Reader, SheetType, SheetVisible, Sheets, Style,
    TextRotation, UnderlineStyle, VerticalAlignment,
};
use rust_xlsxwriter::{
    Chart, ChartType, Color, Format, FormatAlign, FormatUnderline, Formula, Workbook, Worksheet,
};
use serde_json::{json, Value};
use std::collections::{BTreeMap, BTreeSet, HashMap};
use std::io::{BufReader, Cursor, Read};

/// Excel's hard row limit (rows are 1..=1_048_576).
pub const MAX_ROWS: u32 = 1_048_576;
/// Excel's hard column limit (A..=XFD).
pub const MAX_COLS: u32 = 16_384;
/// Maximum number of cells loaded into one model. Larger workbooks are
/// truncated WITH a warning (the JSON model would otherwise be gigabytes).
pub const MAX_CELLS: usize = 2_000_000;
/// Excel's maximum string length in a cell.
const MAX_STRING_CHARS: usize = 32_767;
/// Largest valid Excel date serial (9999-12-31) + 1.
const MAX_DATE_SERIAL: f64 = 2_958_466.0;

/// Excel error literals, recognised in cached formula results.
const ERROR_LITERALS: &[&str] = &[
    "#DIV/0!", "#N/A", "#NAME?", "#NULL!", "#NUM!", "#REF!", "#VALUE!", "#DATA!",
    "#GETTING_DATA", "#SPILL!", "#CALC!", "#FIELD!", "#BLOCKED!", "#CONNECT!", "#BUSY!",
    "#UNKNOWN!", "#EXTERNAL!", "#PYTHON!",
];

// =====================================================================
// Addresses
// =====================================================================

fn col_to_index(col: &str) -> Option<u32> {
    if col.is_empty() || col.len() > 3 { return None; }
    let mut n: u32 = 0;
    for ch in col.chars() {
        if !ch.is_ascii_alphabetic() { return None; }
        n = n.checked_mul(26)?.checked_add(ch.to_ascii_uppercase() as u32 - 64)?;
    }
    if n == 0 || n > MAX_COLS { return None; }
    Some(n - 1)
}

/// 0-based column index -> letters ("A", "Z", "AA", ...).
pub fn index_to_col(i: u32) -> String {
    let mut s = Vec::new();
    let mut n = i as u64 + 1;
    while n > 0 {
        let r = ((n - 1) % 26) as u8;
        s.push(b'A' + r);
        n = (n - 1) / 26;
    }
    s.reverse();
    String::from_utf8(s).unwrap_or_default()
}

/// "B12" / "$B$12" -> 0-based (row, col). Rejects row 0, rows/cols beyond
/// Excel's limits and anything that is not exactly letters followed by digits.
#[allow(dead_code)] // public API; wired by main.rs / used by tests
pub fn parse_addr(addr: &str) -> Option<(u32, u32)> {
    let a: String = addr.trim().chars().filter(|&c| c != '$').collect();
    let split = a.find(|c: char| c.is_ascii_digit())?;
    let (col, row) = a.split_at(split);
    if col.is_empty() || row.is_empty() || !row.bytes().all(|b| b.is_ascii_digit()) {
        return None;
    }
    let c = col_to_index(col)?;
    let r: u32 = row.parse().ok()?;
    let r = r.checked_sub(1)?;
    if r >= MAX_ROWS { return None; }
    Some((r, c))
}

/// 0-based (row, col) -> "B12".
#[allow(dead_code)] // public API; wired by main.rs / used by tests
pub fn cell_addr(row: u32, col: u32) -> String {
    format!("{}{}", index_to_col(col), row as u64 + 1)
}

/// "A1:B2" -> ((r1, c1), (r2, c2)) normalised so r1<=r2, c1<=c2.
#[allow(dead_code)] // public API; wired by main.rs / used by tests
pub fn parse_range_ref(s: &str) -> Option<((u32, u32), (u32, u32))> {
    let (a, b) = s.split_once(':').unwrap_or((s, s));
    let (r1, c1) = parse_addr(a)?;
    let (r2, c2) = parse_addr(b)?;
    Some(((r1.min(r2), c1.min(c2)), (r1.max(r2), c1.max(c2))))
}

// =====================================================================
// Dates
// =====================================================================

/// Days since 1970-01-01 for a proleptic Gregorian date (Howard Hinnant).
fn days_from_civil(y: i64, m: u32, d: u32) -> i64 {
    let y = if m <= 2 { y - 1 } else { y };
    let era = if y >= 0 { y } else { y - 399 } / 400;
    let yoe = y - era * 400;
    let mp = (m as i64 + 9) % 12;
    let doy = (153 * mp + 2) / 5 + d as i64 - 1;
    let doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    era * 146097 + doe - 719468
}

fn is_leap(y: i64) -> bool { (y % 4 == 0 && y % 100 != 0) || y % 400 == 0 }

fn days_in_month(y: i64, m: u32) -> u32 {
    match m {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        2 => if is_leap(y) { 29 } else { 28 },
        _ => 0,
    }
}

/// What an ISO-ish date/time string in the model means.
#[derive(Debug, Clone, Copy, PartialEq)]
enum IsoKind {
    Date,
    DateTime,
    Time,
    Duration,
}

/// Parse the model's date strings into an Excel serial (1900 system).
/// Accepts "YYYY-MM-DD", "YYYY-MM-DD[T ]HH:MM[:SS[.fff]]" (a trailing "Z" or
/// "+hh:mm" offset is IGNORED — values are naive local times and must never
/// be shifted), time-only "HH:MM[:SS]" and ISO durations "PT1H30M"/"P1DT2H".
fn iso_to_serial(s: &str) -> Option<(f64, IsoKind)> {
    let s = s.trim();
    if s.starts_with('P') || s.starts_with("-P") {
        return parse_iso_duration(s).map(|secs| (secs / 86400.0, IsoKind::Duration));
    }
    if let Some(secs) = parse_hms(s) {
        return Some((secs / 86400.0, IsoKind::Time));
    }
    let b = s.as_bytes();
    if b.len() < 10 || b[4] != b'-' || b[7] != b'-' { return None; }
    if ![0, 1, 2, 3, 5, 6, 8, 9].iter().all(|&i| b[i].is_ascii_digit()) { return None; }
    let y: i64 = s[0..4].parse().ok()?;
    let m: u32 = s[5..7].parse().ok()?;
    let d: u32 = s[8..10].parse().ok()?;
    if !(1..=12).contains(&m) || d == 0 || d > days_in_month(y, m) { return None; }
    if !(1900..=9999).contains(&y) { return None; }
    let mut serial = (days_from_civil(y, m, d) - days_from_civil(1899, 12, 30)) as f64;
    // Excel's phantom 1900-02-29: real dates before 1900-03-01 are one lower.
    if serial < 61.0 { serial -= 1.0; }
    let rest = &s[10..];
    if rest.is_empty() { return Some((serial, IsoKind::Date)); }
    let rest = rest.strip_prefix('T').or_else(|| rest.strip_prefix(' '))?;
    // Drop a timezone designator; it is intentionally ignored (naive time).
    let time = rest.trim_end_matches('Z');
    let time = match time.char_indices().skip(5).find(|(_, c)| *c == '+' || *c == '-') {
        Some((i, _)) => &time[..i],
        None => time,
    };
    let secs = parse_hms(time)?;
    if secs >= 86400.0 { return None; }
    Some((serial + secs / 86400.0, IsoKind::DateTime))
}

/// "HH:MM", "HH:MM:SS", "HH:MM:SS.fff" -> seconds since midnight.
fn parse_hms(s: &str) -> Option<f64> {
    let parts: Vec<&str> = s.split(':').collect();
    if parts.len() < 2 || parts.len() > 3 { return None; }
    if parts[0].is_empty() || parts[0].len() > 2 || parts[1].len() != 2 { return None; }
    if !parts[0].bytes().chain(parts[1].bytes()).all(|b| b.is_ascii_digit()) { return None; }
    let h: f64 = parts[0].parse().ok()?;
    let m: f64 = parts[1].parse().ok()?;
    let sec: f64 = match parts.get(2) {
        Some(p) => {
            if p.len() < 2 || !p.bytes().all(|b| b.is_ascii_digit() || b == b'.') { return None; }
            p.parse().ok()?
        }
        None => 0.0,
    };
    if h > 23.0 || m > 59.0 || sec >= 60.0 { return None; }
    Some(h * 3600.0 + m * 60.0 + sec)
}

/// ISO 8601 duration ("PT10H30M00S", "P1DT2H", "-PT5M") -> seconds.
fn parse_iso_duration(s: &str) -> Option<f64> {
    let (neg, s) = match s.strip_prefix('-') { Some(r) => (true, r), None => (false, s) };
    let s = s.strip_prefix('P')?;
    let mut total = 0.0;
    let mut num = String::new();
    let mut in_time = false;
    let mut any = false;
    for ch in s.chars() {
        match ch {
            'T' => in_time = true,
            '0'..='9' | '.' | ',' => num.push(if ch == ',' { '.' } else { ch }),
            _ => {
                let n: f64 = num.parse().ok()?;
                num.clear();
                any = true;
                total += n * match (ch, in_time) {
                    ('W', false) => 7.0 * 86400.0,
                    ('D', false) => 86400.0,
                    ('H', true) => 3600.0,
                    ('M', true) => 60.0,
                    ('S', true) => 1.0,
                    _ => return None, // years/months have no fixed length
                };
            }
        }
    }
    if !num.is_empty() || !any { return None; }
    Some(if neg { -total } else { total })
}

/// Convert a calamine date/time cell into a model cell.
/// - real dates -> t "d", v "YYYY-MM-DD" or "YYYY-MM-DDTHH:MM:SS" (naive)
/// - durations, time-only values (1900 system, serial < 1), out-of-range
///   serials and Excel's phantom 1900-02-29 -> t "n" with a number format,
///   so the exact value round-trips.
fn datetime_cell(dt: &ExcelDateTime) -> Cell {
    let v = dt.as_f64();
    let number = |z: &str| {
        let mut c = Cell::new(num_json(v), "n");
        c.z = Some(z.to_string());
        c
    };
    if !v.is_finite() { return Cell::new(Value::Null, "n"); }
    if dt.is_duration() { return number("[h]:mm:ss"); }
    // ExcelDateTime keeps its 1904 flag private; recover it by comparison.
    let is_1904 = *dt == ExcelDateTime::new(v, ExcelDateTimeType::DateTime, true);
    if v < 0.0 || v >= MAX_DATE_SERIAL { return number("General"); }
    if !is_1904 && v < 1.0 { return number("hh:mm:ss"); }
    // Round to whole seconds first so 23:59:59.9 does not print as 24:00:00.
    let total = (v * 86400.0).round() as u64;
    let (days, secs) = (total / 86400, total % 86400);
    if !is_1904 && days as f64 >= MAX_DATE_SERIAL { return number("General"); }
    let (y, m, d, ..) =
        ExcelDateTime::new(days as f64, ExcelDateTimeType::DateTime, is_1904).to_ymd_hms_milli();
    if (y, m, d) == (1900, 2, 29) { return number("yyyy-mm-dd"); }
    let mut iso = format!("{y:04}-{m:02}-{d:02}");
    if secs > 0 {
        iso.push_str(&format!("T{:02}:{:02}:{:02}", secs / 3600, (secs / 60) % 60, secs % 60));
    }
    Cell::new(json!(iso), "d")
}

/// Normalise an ODS ISO date ("2024-01-15T10:30:00.000") to the model form.
fn normalize_iso_datetime(s: &str) -> Option<String> {
    let s = s.trim();
    let (serial, kind) = iso_to_serial(s)?;
    match kind {
        IsoKind::Date => Some(s[..10].to_string()),
        IsoKind::DateTime => {
            let secs = ((serial.fract()) * 86400.0).round() as u64;
            if secs == 0 { return Some(s[..10].to_string()); }
            if secs >= 86400 { return None; }
            Some(format!("{}T{:02}:{:02}:{:02}", &s[..10], secs / 3600, (secs / 60) % 60, secs % 60))
        }
        _ => None,
    }
}

// =====================================================================
// Reading
// =====================================================================

/// JSON number without a spurious ".0" for integral values.
fn num_json(n: f64) -> Value {
    if n.is_finite() && n.fract() == 0.0 && n.abs() < 9_007_199_254_740_992.0 {
        json!(n as i64)
    } else {
        json!(n)
    }
}

fn cell_from_data(data: &Data) -> Option<Cell> {
    Some(match data {
        Data::Empty => return None,
        Data::String(s) => Cell::new(json!(s), "s"),
        Data::Float(f) => Cell::new(num_json(*f), "n"),
        Data::Int(i) => Cell::new(json!(i), "n"),
        Data::Bool(b) => Cell::new(json!(b), "b"),
        Data::Error(e) => Cell::new(json!(e.to_string()), "e"),
        Data::DateTime(dt) => datetime_cell(dt),
        Data::DateTimeIso(s) => match normalize_iso_datetime(s) {
            Some(iso) => Cell::new(json!(iso), "d"),
            None => Cell::new(json!(s), "s"),
        },
        Data::DurationIso(s) => match parse_iso_duration(s) {
            Some(secs) => {
                let mut c = Cell::new(num_json(secs / 86400.0), "n");
                c.z = Some("[h]:mm:ss".into());
                c
            }
            None => Cell::new(json!(s), "s"),
        },
    })
}

/// Remove the storage-only prefixes Excel puts in front of newer functions
/// ("_xlfn.XLOOKUP" -> "XLOOKUP"). rust_xlsxwriter re-adds them on write.
/// Lambda parameter prefixes ("_xlpm.") are kept because they are not re-added.
fn clean_formula(f: &str) -> String {
    let mut out = String::with_capacity(f.len());
    let mut in_str = false;
    let mut rest = f;
    while let Some(ch) = rest.chars().next() {
        if ch == '"' { in_str = !in_str; }
        if !in_str {
            if let Some(r) = rest.strip_prefix("_xlfn.") { rest = r; continue; }
            if let Some(r) = rest.strip_prefix("_xlws.") { rest = r; continue; }
        }
        out.push(ch);
        rest = &rest[ch.len_utf8()..];
    }
    out
}

/// Ordered, de-duplicated warning list.
#[derive(Default)]
struct Warnings {
    list: Vec<String>,
    seen: BTreeSet<String>,
}

impl Warnings {
    fn push(&mut self, w: impl Into<String>) {
        let w = w.into();
        if self.seen.insert(w.clone()) { self.list.push(w); }
    }
}

/// Parse state shared across sheets.
struct ReadCtx {
    warn: Warnings,
    /// Remaining cell budget (MAX_CELLS for the whole book).
    budget: usize,
    truncated: bool,
    /// Formatting features that the model cannot carry (for one warning).
    lost_formatting: BTreeSet<&'static str>,
    /// (font name, size*100) -> count, to detect non-uniform fonts.
    fonts: HashMap<(String, i64), usize>,
}

impl ReadCtx {
    fn new(budget: usize) -> Self {
        ReadCtx { warn: Warnings::default(), budget, truncated: false, lost_formatting: BTreeSet::new(), fonts: HashMap::new() }
    }
    /// Reserve room for one new cell; false (and mark truncation) if full.
    fn take(&mut self) -> bool {
        if self.budget == 0 { self.truncated = true; return false; }
        self.budget -= 1;
        true
    }
}

fn hex_color(c: &calamine::Color) -> String {
    format!("#{:02X}{:02X}{:02X}", c.red, c.green, c.blue)
}

/// Map a calamine style to the model's (z, s). Unsupported properties are
/// recorded in `ctx.lost_formatting` so a single warning can be emitted.
fn style_to_model(st: &Style, ctx: &mut ReadCtx) -> (Option<String>, Option<CellStyle>) {
    let mut s = CellStyle::default();
    if let Some(font) = &st.font {
        s.b = font.is_bold();
        s.i = font.is_italic();
        s.u = font.underline != UnderlineStyle::None;
        if font.strikethrough { ctx.lost_formatting.insert("strikethrough"); }
        if let Some(c) = &font.color {
            // Black is the automatic/default text colour; don't pin it.
            if (c.red, c.green, c.blue) != (0, 0, 0) { s.color = Some(hex_color(c)); }
        }
        let key = (font.name.clone().unwrap_or_default(), (font.size.unwrap_or(11.0) * 100.0).round() as i64);
        *ctx.fonts.entry(key).or_insert(0) += 1;
    }
    if let Some(fill) = &st.fill {
        match fill.pattern {
            FillPattern::None => {}
            FillPattern::Solid => {
                if let Some(c) = fill.get_color() { s.bg = Some(hex_color(&c)); }
            }
            _ => { ctx.lost_formatting.insert("pattern fills"); }
        }
    }
    if let Some(al) = &st.alignment {
        match al.horizontal {
            HorizontalAlignment::Left => s.align = Some("left".into()),
            HorizontalAlignment::Center | HorizontalAlignment::CenterContinuous => s.align = Some("center".into()),
            HorizontalAlignment::Right => s.align = Some("right".into()),
            HorizontalAlignment::General => {}
            _ => { ctx.lost_formatting.insert("justified/filled alignment"); }
        }
        s.wrap = al.wrap_text;
        if al.vertical != VerticalAlignment::Bottom { ctx.lost_formatting.insert("vertical alignment"); }
        if al.text_rotation != TextRotation::None { ctx.lost_formatting.insert("text rotation"); }
        if al.indent.is_some_and(|i| i > 0) { ctx.lost_formatting.insert("indentation"); }
    }
    if st.borders.as_ref().is_some_and(|b| b.has_visible_borders()) {
        ctx.lost_formatting.insert("borders");
    }
    let z = st.number_format.as_ref()
        .map(|n| n.format_code.trim().to_string())
        .filter(|z| !z.is_empty() && !z.eq_ignore_ascii_case("General"));
    (z, if s.is_empty() { None } else { Some(s) })
}

/// Attach formula text (calamine gives it without "=") to a cell.
fn set_formula(cell: &mut Cell, f: &str) {
    let body = f.strip_prefix('=').unwrap_or(f);
    cell.f = Some(format!("={}", clean_formula(body)));
    // A cached error result stored as a string (t="str") is an error value.
    if cell.t == "s" {
        if let Some(s) = cell.v.as_str() {
            if ERROR_LITERALS.contains(&s) { cell.t = "e".into(); }
        }
    }
}

/// Generic path (xls / xlsb / ods): dense calamine ranges. Positions from
/// used_cells() are RELATIVE to range.start(); the formula range has its own
/// start. (calamine reads no cell styles for these formats.)
fn fill_sheet_from_ranges(
    sheet: &mut Sheet,
    values: &Range<Data>,
    formulas: Option<&Range<String>>,
    ctx: &mut ReadCtx,
) {
    let (r0, c0) = values.start().unwrap_or((0, 0));
    for (r, c, d) in values.used_cells() {
        let Some(cell) = cell_from_data(d) else { continue };
        if !ctx.take() { break; }
        sheet.cells.insert(cell_addr(r0 + r as u32, c0 + c as u32), cell);
    }
    if let Some(fr) = formulas {
        let (fr0, fc0) = fr.start().unwrap_or((0, 0));
        for (r, c, f) in fr.used_cells() {
            let f = f.trim();
            if f.is_empty() { continue; }
            let addr = cell_addr(fr0 + r as u32, fc0 + c as u32);
            if !sheet.cells.contains_key(&addr) {
                if !ctx.take() { break; }
                sheet.cells.insert(addr.clone(), Cell::new(Value::Null, "f"));
            }
            if let Some(cell) = sheet.cells.get_mut(&addr) { set_formula(cell, f); }
        }
    }
}

/// .xlsx / .xlsm path: stream the sheet XML with calamine's cell reader.
/// Unlike `worksheet_range`, this never allocates the dense bounding box, so
/// a sheet with cells at A1 and XFD1048576 costs two cells, not 17 billion.
/// Pass 1: values + (shared-expanded) formulas. Pass 2: style ids, mapped
/// through the workbook palette (also yields formatted blank cells).
fn fill_sheet_xlsx<RS: Read + std::io::Seek>(
    x: &mut calamine::Xlsx<RS>,
    sheet: &mut Sheet,
    ctx: &mut ReadCtx,
) -> Result<(), String> {
    let name = sheet.name.clone();
    let err = |e: calamine::XlsxError| format!("sheet {name:?}: {e}");
    {
        let mut rd = x.worksheet_cells_reader(&name).map_err(err)?;
        while let Some(rec) = rd.next_cell_with_formula().map_err(err)? {
            let data: Data = rec.value.into();
            let formula = rec.formula.as_deref().map(str::trim).filter(|f| !f.is_empty());
            let cell = match (cell_from_data(&data), formula) {
                (None, None) => continue,
                (c, f) => {
                    let mut cell = c.unwrap_or_else(|| Cell::new(Value::Null, "f"));
                    if let Some(f) = f { set_formula(&mut cell, f); }
                    cell
                }
            };
            if !ctx.take() { break; }
            sheet.cells.insert(cell_addr(rec.pos.0, rec.pos.1), cell);
        }
    }
    // Formatting is best effort: a failure here keeps the values loaded.
    let mut rd = match x.worksheet_cells_reader(&name) {
        Ok(rd) => rd,
        Err(e) => {
            ctx.warn.push(format!("Formatting on sheet \"{name}\" could not be read ({e})."));
            return Ok(());
        }
    };
    let palette: Vec<Style> = rd.styles().to_vec();
    let mut mapped: HashMap<usize, (Option<String>, Option<CellStyle>)> = HashMap::new();
    loop {
        let (r, c, id) = match rd.next_style_id() {
            Ok(Some(t)) => t,
            Ok(None) => break,
            Err(e) => {
                ctx.warn.push(format!("Formatting on sheet \"{name}\" could not be read completely ({e})."));
                break;
            }
        };
        let Some(st) = palette.get(id) else { continue };
        let (z, s) = mapped.entry(id).or_insert_with(|| style_to_model(st, ctx)).clone();
        let addr = cell_addr(r, c);
        if let Some(cell) = sheet.cells.get_mut(&addr) {
            if z.is_some() { cell.z = z; }
            cell.s = s;
        } else if s.is_some() || z.is_some() {
            // Formatted empty cell (fill colour, bold header slot, a column
            // pre-formatted as currency, ...): keep both style and format so
            // a value typed there later is shown the way the author intended.
            if ctx.budget == 0 { continue; }
            ctx.take();
            let mut cell = Cell::new(Value::Null, "s");
            cell.s = s;
            cell.z = z;
            sheet.cells.insert(addr, cell);
        }
    }
    Ok(())
}

/// Things read straight from the OOXML package (calamine does not expose them).
#[derive(Default)]
struct XlsxPackage {
    /// workbookView@activeTab (index into ALL <sheet> entries).
    active_tab: usize,
    /// Sheet name -> <col> spans (min0, max0, width-in-chars, hidden).
    cols: HashMap<String, Vec<(u32, u32, f64, bool)>>,
    /// Human-readable names of unsupported features found in the package.
    features: BTreeSet<&'static str>,
    /// Sheet name -> frozen panes, auto filter and charts of that sheet.
    extras: HashMap<String, SheetExtras>,
}

/// Per-sheet things read from the worksheet part and its drawing.
#[derive(Default)]
struct SheetExtras {
    freeze: Option<Freeze>,
    auto_filter: Option<String>,
    charts: Vec<ChartRead>,
    /// Embedded charts that could not be represented in the model.
    unreadable_charts: usize,
}

/// A chart read from the package (geometry still in file units).
#[derive(Debug, Clone, PartialEq)]
struct ChartRead {
    kind: &'static str,
    title: Option<String>,
    /// Bounding box of all series references: ((r1, c1), (r2, c2)).
    range: ((u32, u32), (u32, u32)),
    anchor: RawAnchor,
}

/// Read one zip part (bounded by MAX_UNZIPPED_TOTAL).
fn read_zip_part(zip: &mut zip::ZipArchive<Cursor<&[u8]>>, name: &str) -> Option<Vec<u8>> {
    let f = zip.by_name(name).ok()?;
    let mut out = Vec::new();
    f.take(crate::limits::MAX_UNZIPPED_TOTAL).read_to_end(&mut out).ok()?;
    Some(out)
}

/// Relationships of a part: Id -> resolved internal target (external
/// targets are skipped). Missing / malformed rels give an empty map.
fn read_rels(zip: &mut zip::ZipArchive<Cursor<&[u8]>>, part: &str) -> HashMap<String, String> {
    use quick_xml::events::Event;
    let (dir, file) = part.rsplit_once('/').unwrap_or(("", part));
    let rels_path = if dir.is_empty() { format!("_rels/{file}.rels") } else { format!("{dir}/_rels/{file}.rels") };
    let mut out = HashMap::new();
    let Some(xml) = read_zip_part(zip, &rels_path) else { return out };
    let mut r = quick_xml::Reader::from_reader(xml.as_slice());
    let mut buf = Vec::new();
    loop {
        match r.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) if e.local_name().as_ref() == b"Relationship" => {
                let external = attr_value(&e, b"TargetMode").is_some_and(|m| m.eq_ignore_ascii_case("External"));
                if let (false, Some(id), Some(t)) = (external, attr_value(&e, b"Id"), attr_value(&e, b"Target")) {
                    out.insert(id, resolve_target(dir, &t));
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    out
}

fn attr_value(e: &quick_xml::events::BytesStart, local: &[u8]) -> Option<String> {
    e.attributes().flatten()
        .find(|a| a.key.local_name().as_ref() == local)
        .and_then(|a| a.unescape_value().ok().map(|v| v.into_owned()))
}

/// Resolve a relationship target relative to the part's directory.
fn resolve_target(base_dir: &str, target: &str) -> String {
    if let Some(abs) = target.strip_prefix('/') { return abs.to_string(); }
    let mut parts: Vec<&str> = base_dir.split('/').filter(|p| !p.is_empty()).collect();
    for seg in target.split('/') {
        match seg {
            "" | "." => {}
            ".." => { parts.pop(); }
            s => parts.push(s),
        }
    }
    parts.join("/")
}

/// Excel stores <col width> in "character + padding" units; convert back to
/// the plain character count the model (and rust_xlsxwriter) use.
fn stored_width_to_chars(w: f64) -> f64 {
    if !(w.is_finite()) || w <= 0.0 { return 0.0; }
    let px = (w * 7.0).round();
    let chars = if px >= 12.0 { (px - 5.0) / 7.0 } else { px / 12.0 };
    (chars * 100.0).round() / 100.0
}

//// Read workbook view, column widths, frozen panes, auto filters, charts and
/// feature flags from an .xlsx/.xlsm.
/// Best effort: returns None for non-OOXML input or malformed packages.
fn read_xlsx_package(bytes: &[u8]) -> Option<XlsxPackage> {
    use quick_xml::events::Event;
    let mut zip = zip::ZipArchive::new(Cursor::new(bytes)).ok()?;
    let limit = crate::limits::MAX_UNZIPPED_TOTAL;
    let mut pkg = XlsxPackage::default();

    // Chart parts present in the package (lower-case); those not handled by a
    // worksheet drawing below (chart sheets, orphans) are flagged as lost.
    let mut chart_parts: BTreeSet<String> = BTreeSet::new();
    let entry_names: Vec<String> = zip.file_names().flatten().map(|n| n.to_ascii_lowercase()).collect();
    for n in entry_names {
        if n.contains("/chartsheets/") { pkg.features.insert("charts"); }
        if n.contains("/charts/chart") && n.ends_with(".xml") && !n.contains("/_rels/") { chart_parts.insert(n.clone()); }
        if n.contains("/media/") { pkg.features.insert("images"); }
        if n.contains("/pivottables/") { pkg.features.insert("pivot tables"); }
        if n.contains("comments") { pkg.features.insert("comments"); }
        if n.ends_with("vbaproject.bin") { pkg.features.insert("macros (VBA)"); }
        if n.contains("/tables/table") { pkg.features.insert("table formatting"); }
        if n.contains("/externallinks/") { pkg.features.insert("links to other workbooks"); }
    }

    // Locate the workbook part via the package relationships.
    let mut wb_path = "xl/workbook.xml".to_string();
    if let Some(rels) = read_zip_part(&mut zip, "_rels/.rels") {
        let mut r = quick_xml::Reader::from_reader(rels.as_slice());
        let mut buf = Vec::new();
        loop {
            match r.read_event_into(&mut buf) {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) if e.local_name().as_ref() == b"Relationship" => {
                    let ty = attr_value(&e, b"Type").unwrap_or_default();
                    if ty.ends_with("/officeDocument") {
                        if let Some(t) = attr_value(&e, b"Target") { wb_path = resolve_target("", &t); }
                    }
                }
                Ok(Event::Eof) | Err(_) => break,
                _ => {}
            }
            buf.clear();
        }
    }
    let wb_xml = read_zip_part(&mut zip, &wb_path)?;

    // workbook.xml: sheet list (name, r:id) and activeTab.
    let mut sheets: Vec<(String, String)> = Vec::new();
    {
        let mut r = quick_xml::Reader::from_reader(wb_xml.as_slice());
        let mut buf = Vec::new();
        loop {
            match r.read_event_into(&mut buf) {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match e.local_name().as_ref() {
                    b"workbookView" => {
                        if let Some(t) = attr_value(&e, b"activeTab").and_then(|v| v.parse().ok()) {
                            pkg.active_tab = t;
                        }
                    }
                    b"sheet" => {
                        let name = attr_value(&e, b"name").unwrap_or_default();
                        let rid = attr_value(&e, b"id").unwrap_or_default();
                        sheets.push((name, rid));
                    }
                    b"definedName" => {
                        let n = attr_value(&e, b"name").unwrap_or_default();
                        if !n.starts_with("_xlnm.") { pkg.features.insert("named ranges"); }
                    }
                    _ => {}
                },
                Ok(Event::Eof) | Err(_) => break,
                _ => {}
            }
            buf.clear();
        }
    }
    let targets = read_rels(&mut zip, &wb_path);

    // Stream every worksheet part once: <col> spans, frozen panes, the auto
    // filter, the drawing reference and feature flags.
    let mut handled_charts: BTreeSet<String> = BTreeSet::new();
    let mut charts_seen = 0usize;
    for (name, rid) in sheets {
        let Some(path) = targets.get(&rid).cloned() else { continue };
        let mut spans = Vec::new();
        let mut extras = SheetExtras::default();
        let mut pane_seen = false;
        let mut filter_seen = false;
        let mut drawing_rid: Option<String> = None;
        {
            let Ok(file) = zip.by_name(&path) else { continue };
            let mut r = quick_xml::Reader::from_reader(BufReader::new(file.take(limit)));
            let mut buf = Vec::new();
            loop {
                match r.read_event_into(&mut buf) {
                    Ok(Event::Start(e)) | Ok(Event::Empty(e)) => match e.local_name().as_ref() {
                        b"col" => {
                            let min: u32 = attr_value(&e, b"min").and_then(|v| v.parse().ok()).unwrap_or(0);
                            let max: u32 = attr_value(&e, b"max").and_then(|v| v.parse().ok()).unwrap_or(min);
                            let hidden = matches!(attr_value(&e, b"hidden").as_deref(), Some("1") | Some("true"));
                            let width = attr_value(&e, b"width").and_then(|v| v.parse::<f64>().ok());
                            if min >= 1 && max >= min && (width.is_some() || hidden) {
                                let w = if hidden { 0.0 } else { stored_width_to_chars(width.unwrap_or(0.0)) };
                                spans.push((min - 1, (max - 1).min(MAX_COLS - 1), w, hidden));
                            }
                        }
                        b"row" => {
                            if !pkg.features.contains("custom row heights")
                                && matches!(attr_value(&e, b"customHeight").as_deref(), Some("1") | Some("true"))
                            {
                                pkg.features.insert("custom row heights");
                            }
                        }
                        b"conditionalFormatting" => { pkg.features.insert("conditional formatting"); }
                        b"dataValidation" => { pkg.features.insert("data validation"); }
                        b"hyperlink" => { pkg.features.insert("hyperlinks"); }
                        b"pane" if !pane_seen => {
                            // Only the first sheet view counts; plain (non-frozen)
                            // splits are a view setting and are simply ignored.
                            pane_seen = true;
                            let state = attr_value(&e, b"state").unwrap_or_default();
                            if state == "frozen" || state == "frozenSplit" {
                                let split = |k: &[u8], max: u32| {
                                    attr_value(&e, k).and_then(|v| v.trim().parse::<f64>().ok())
                                        .filter(|f| f.is_finite() && *f > 0.0)
                                        .map(|f| (f.round() as u64).min(max as u64 - 1) as u32)
                                        .unwrap_or(0)
                                };
                                let freeze = Freeze { rows: split(b"ySplit", MAX_ROWS), cols: split(b"xSplit", MAX_COLS) };
                                if freeze.rows > 0 || freeze.cols > 0 { extras.freeze = Some(freeze); }
                            }
                        }
                        b"autoFilter" if !filter_seen => {
                            // The worksheet-level filter precedes any custom-view copy.
                            filter_seen = true;
                            if let Some(((r1, c1), (r2, c2))) = attr_value(&e, b"ref").and_then(|v| parse_range_ref(v.trim())) {
                                extras.auto_filter = Some(format!("{}:{}", cell_addr(r1, c1), cell_addr(r2, c2)));
                            }
                        }
                        b"filterColumn" => { pkg.features.insert("filter conditions (hidden rows)"); }
                        b"drawing" => {
                            if drawing_rid.is_none() { drawing_rid = attr_value(&e, b"id"); }
                        }
                        b"sheetProtection" => { pkg.features.insert("sheet protection"); }
                        _ => {}
                    },
                    Ok(Event::Eof) | Err(_) => break,
                    _ => {}
                }
                buf.clear();
            }
        }
        if let Some(drid) = drawing_rid {
            let mut ctx = DrawingCtx {
                zip: &mut zip, features: &mut pkg.features,
                handled_charts: &mut handled_charts, charts_seen: &mut charts_seen,
            };
            ctx.read_sheet_drawing(&path, &drid, &name, &mut extras);
        }
        if !spans.is_empty() { pkg.cols.insert(name.clone(), spans); }
        pkg.extras.insert(name, extras);
    }
    if chart_parts.iter().any(|p| !handled_charts.contains(p)) { pkg.features.insert("charts"); }
    Some(pkg)
}

/// Shared state while following worksheet drawings to their charts.
struct DrawingCtx<'a, 'b> {
    zip: &'a mut zip::ZipArchive<Cursor<&'b [u8]>>,
    features: &'a mut BTreeSet<&'static str>,
    /// Chart parts (lower-case paths) referenced from worksheet drawings.
    handled_charts: &'a mut BTreeSet<String>,
    /// Charts looked at so far in the whole workbook (MAX_CHARTS cap).
    charts_seen: &'a mut usize,
}

impl DrawingCtx<'_, '_> {
    /// Follow a worksheet's <drawing r:id> to its drawing part and read every
    /// anchored chart into `extras`. Pictures / shapes / unsupported charts
    /// become feature flags or unreadable-chart counts.
    fn read_sheet_drawing(&mut self, sheet_path: &str, drawing_rid: &str, sheet_name: &str, extras: &mut SheetExtras) {
        let sheet_rels = read_rels(self.zip, sheet_path);
        let Some(drawing_path) = sheet_rels.get(drawing_rid) else { return };
        let Some(xml) = read_zip_part(self.zip, drawing_path) else { return };
        let anchors = parse_drawing_xml(&xml);
        if anchors.is_empty() { return; }
        let rels = read_rels(self.zip, drawing_path);
        for a in anchors {
            if a.picture { self.features.insert("images"); }
            if a.shape { self.features.insert("drawings and shapes"); }
            for rid in &a.other_chart_rids {
                if let Some(p) = rels.get(rid) { self.handled_charts.insert(p.to_ascii_lowercase()); }
            }
            extras.unreadable_charts += a.other_charts;
            // A chart is only representable when it is the anchor's sole content
            // (charts inside groups would lose their group).
            let single = a.chart_rids.len() == 1 && !a.shape && !a.picture;
            for rid in &a.chart_rids {
                *self.charts_seen += 1;
                let path = rels.get(rid);
                if let Some(p) = path { self.handled_charts.insert(p.to_ascii_lowercase()); }
                let read = match (single, a.anchor, path) {
                    (true, Some(anchor), Some(p)) if *self.charts_seen <= MAX_CHARTS => read_zip_part(self.zip, p)
                        .and_then(|x| parse_chart_xml(&x))
                        .and_then(|info| {
                            let range = chart_bbox(&info.refs, sheet_name)?;
                            Some(ChartRead { kind: info.kind?, title: info.title, range, anchor })
                        }),
                    _ => None,
                };
                match read {
                    Some(c) => extras.charts.push(c),
                    None => extras.unreadable_charts += 1,
                }
            }
        }
    }
}

// Parse .xlsx/.xls/.xlsm/.xlsb/.ods bytes into an ExcelBook.
/// `sheet_name_hint`: when Some, only that sheet is returned.
pub fn parse_book(bytes: &[u8], sheet_name_hint: Option<&str>) -> Result<ExcelBook, String> {
    parse_book_limited(bytes, sheet_name_hint, MAX_CELLS)
}

fn parse_book_limited(bytes: &[u8], sheet_name_hint: Option<&str>, max_cells: usize) -> Result<ExcelBook, String> {
    let pkg = if bytes.starts_with(b"PK") { read_xlsx_package(bytes) } else { None };
    let mut wb: Sheets<_> = open_workbook_auto_from_rs(Cursor::new(bytes))
        .map_err(|e| format!("spreadsheet parse failed: {e}"))?;
    let metas: Vec<calamine::Sheet> = wb.sheets_metadata().to_vec();
    let mut ctx = ReadCtx::new(max_cells);
    let mut book = ExcelBook { kind: "excel".into(), active: 0, sheets: Vec::new(), warnings: Vec::new() };
    let active_tab = pkg.as_ref().map(|p| p.active_tab).unwrap_or(0);
    let mut active_model: Option<usize> = None;
    let mut unreadable_charts = 0usize;

    for (idx, meta) in metas.iter().enumerate() {
        if let Some(h) = sheet_name_hint { if meta.name != h { continue; } }
        if meta.typ != SheetType::WorkSheet {
            let what = match meta.typ {
                SheetType::ChartSheet => "chart sheet",
                SheetType::MacroSheet => "macro sheet",
                SheetType::DialogSheet => "dialog sheet",
                _ => "non-worksheet sheet",
            };
            ctx.warn.push(format!("\"{}\" is a {what}; it cannot be shown and will be dropped if you save.", meta.name));
            continue;
        }
        let name = meta.name.clone();
        let mut sheet = Sheet {
            name: name.clone(),
            hidden: !matches!(meta.visible, SheetVisible::Visible),
            ..Default::default()
        };
        if let Sheets::Xlsx(x) = &mut wb {
            fill_sheet_xlsx(x, &mut sheet, &mut ctx)?;
        } else {
            let values = wb.worksheet_range(&name).map_err(|e| format!("sheet {name:?}: {e}"))?;
            let formulas = match wb.worksheet_formula(&name) {
                Ok(f) => Some(f),
                Err(e) => {
                    ctx.warn.push(format!("Formulas on sheet \"{name}\" could not be read ({e}); only their values were loaded."));
                    None
                }
            };
            fill_sheet_from_ranges(&mut sheet, &values, formulas.as_ref(), &mut ctx);
        }

        // Merged ranges (xlsx/xls only; calamine has no merge API for ods/xlsb).
        let merges = match &mut wb {
            Sheets::Xlsx(x) => x.merge_cells_by_sheet_name(&name).ok(),
            Sheets::Xls(x) => x.merge_cells_by_sheet_name(&name).ok(),
            _ => None,
        };
        for d in merges.unwrap_or_default() {
            if d.start == d.end { continue; }
            sheet.merges.push(format!("{}:{}", cell_addr(d.start.0, d.start.1), cell_addr(d.end.0, d.end.1)));
        }

        // Frozen panes, auto filter and charts (xlsx/xlsm package read).
        let extras = pkg.as_ref().and_then(|p| p.extras.get(&name));

        // Column widths: expand <col> spans, but don't blow up whole-sheet
        // spans (min=1 max=16384) into 16k entries. Columns under charts are
        // kept so chart geometry stays exact.
        if let Some(spans) = pkg.as_ref().and_then(|p| p.cols.get(&name)) {
            let used_max = sheet.cells.keys().filter_map(|a| parse_addr(a).map(|(_, c)| c))
                .chain(sheet.merges.iter().filter_map(|m| parse_range_ref(m).map(|(_, (_, c))| c)))
                .chain(extras.into_iter().flat_map(|e| e.charts.iter().map(|c| c.anchor.max_col())))
                .max().unwrap_or(0);
            let cap = (used_max + 1).max(25);
            let mut widths = BTreeMap::new();
            for &(min, max, w, _) in spans {
                let hi = if max > cap { cap.max(min) } else { max };
                for c in min..=hi { widths.insert(index_to_col(c), w); }
            }
            if !widths.is_empty() { sheet.col_widths = Some(widths); }
        }

        if let Some(ex) = extras {
            sheet.freeze = ex.freeze;
            sheet.auto_filter = ex.auto_filter.clone();
            let metrics = ColMetrics::new(sheet.col_widths.as_ref());
            sheet.charts = ex.charts.iter().map(|c| chart_to_json(c, &metrics)).collect();
            unreadable_charts += ex.unreadable_charts;
        }

        if idx == active_tab { active_model = Some(book.sheets.len()); }
        book.sheets.push(sheet);
        if ctx.truncated { break; }
    }

    if book.sheets.is_empty() {
        return Err(match sheet_name_hint {
            Some(h) => format!("sheet {h:?} not found"),
            None => "workbook contains no worksheets".into(),
        });
    }

    // Active sheet: the file's active tab when it is a visible worksheet,
    // else the first visible sheet.
    let mut active = active_model.unwrap_or(0);
    if book.sheets.get(active).is_none_or(|s| s.hidden) {
        active = book.sheets.iter().position(|s| !s.hidden).unwrap_or(0);
    }
    book.active = active as u32;

    // ---- fidelity warnings ----
    if ctx.truncated {
        ctx.warn.push(format!(
            "This workbook is too large to load completely: only the first {} cells were loaded. \
             Saving would permanently drop the rest — do not overwrite the original file.",
            max_cells
        ));
    }
    if unreadable_charts > 0 {
        ctx.warn.push(format!("{unreadable_charts} chart(s) could not be read and will be lost if you save."));
    }
    if ctx.fonts.len() > 1 { ctx.lost_formatting.insert("font faces and sizes"); }
    let mut features: Vec<&str> = pkg.as_ref().map(|p| p.features.iter().copied().collect()).unwrap_or_default();
    if pkg.is_none() && !wb.defined_names().is_empty() { features.push("named ranges"); }
    if !features.is_empty() {
        ctx.warn.push(format!(
            "This workbook contains {} that this editor does not keep; they will be lost if you save over the original.",
            join_list(&features)
        ));
    }
    if !ctx.lost_formatting.is_empty() {
        let lost: Vec<&str> = ctx.lost_formatting.iter().copied().collect();
        ctx.warn.push(format!("Some formatting is not supported and will be lost on save: {}.", join_list(&lost)));
    }
    if matches!(wb, Sheets::Xls(_) | Sheets::Ods(_) | Sheets::Xlsb(_)) {
        ctx.warn.push("Cell formatting (fonts, fills, alignment) is not read from this file type; only values, formulas and number formats were loaded.".to_string());
    }
    book.warnings = ctx.warn.list;
    Ok(book)
}

fn join_list(items: &[&str]) -> String {
    match items.len() {
        0 => String::new(),
        1 => items[0].to_string(),
        n => format!("{} and {}", items[..n - 1].join(", "), items[n - 1]),
    }
}

// =====================================================================
// Charts, frozen panes, auto filter
// =====================================================================
//
// CANONICAL chart geometry (shared with lib/model-excel.js; keep in sync):
// model x/y/w/h are px relative to the grid's A1 at 100% zoom.
//   grid column px  = colWidths[col] defined ? max(4, round(w*7+5)) : 88; grid rows 24 px
//   Excel column px = colWidths[col] defined ? max(4, round(w*7+5)) : 64; Excel rows 20 px
// WRITE: col = last c with gridX(c) <= x, x offset scaled by excelPx/gridPx;
//        row = floor(y/24), y offset scaled by 20/24; size = w x h px.
// READ:  anchor cell + offset scaled back to grid px; size measured in Excel px.

const EMU_PER_PX: f64 = 9525.0;
const GRID_ROW_PX: f64 = 24.0;
const EXCEL_ROW_PX: f64 = 20.0;
const GRID_DEFAULT_COL_PX: f64 = 88.0;
const EXCEL_DEFAULT_COL_PX: f64 = 64.0;
/// Most charts read per workbook / written per sheet.
const MAX_CHARTS: usize = 200;
/// Excel's limit of series per chart.
const MAX_CHART_SERIES: u32 = 255;
const DEFAULT_CHART_W: f64 = 480.0;
const DEFAULT_CHART_H: f64 = 288.0;
const MIN_CHART_PX: f64 = 50.0;
const MAX_CHART_PX: f64 = 10_000.0;
/// Chart positions beyond these are clamped (keeps anchors inside the sheet).
const MAX_CHART_X: f64 = 1_000_000.0;
const MAX_CHART_Y: f64 = 20_000_000.0;
/// Nesting / size guards for the drawing and chart XML readers.
const MAX_XML_DEPTH: usize = 256;
const MAX_ANCHORS: usize = 10_000;
const MAX_CHART_REFS: usize = 4_096;

/// ((r1, c1), (r2, c2)), 0-based and normalised.
type Rect = ((u32, u32), (u32, u32));

/// Column pixel metrics of one sheet, derived from the model's colWidths.
struct ColMetrics {
    /// Column -> px for columns with a defined width (same in grid and Excel).
    custom: BTreeMap<u32, f64>,
}

impl ColMetrics {
    fn new(widths: Option<&BTreeMap<String, f64>>) -> Self {
        let mut custom = BTreeMap::new();
        for (k, w) in widths.into_iter().flatten() {
            if let (Some(c), true) = (col_to_index(k.trim()), w.is_finite()) {
                custom.insert(c, (w * 7.0 + 5.0).round().max(4.0));
            }
        }
        ColMetrics { custom }
    }
    fn px(&self, c: u32, grid: bool) -> f64 {
        self.custom.get(&c).copied().unwrap_or(if grid { GRID_DEFAULT_COL_PX } else { EXCEL_DEFAULT_COL_PX })
    }
    fn grid_px(&self, c: u32) -> f64 { self.px(c, true) }
    fn excel_px(&self, c: u32) -> f64 { self.px(c, false) }
    /// Left edge of column `c` in grid (true) or Excel (false) px.
    fn left(&self, c: u32, grid: bool) -> f64 {
        let def = if grid { GRID_DEFAULT_COL_PX } else { EXCEL_DEFAULT_COL_PX };
        let mut x = c as f64 * def;
        for (_, px) in self.custom.range(..c) { x += px - def; }
        x
    }
    fn grid_x(&self, c: u32) -> f64 { self.left(c, true) }
    fn excel_x(&self, c: u32) -> f64 { self.left(c, false) }
    /// Last column whose left edge is <= x (x >= 0) and the remaining offset.
    fn locate(&self, x: f64, grid: bool) -> (u32, f64) {
        let mut left = 0.0;
        let mut c = 0u32;
        loop {
            let w = self.px(c, grid);
            if left + w > x || c + 1 >= MAX_COLS { return (c, x - left); }
            left += w;
            c += 1;
        }
    }
}

/// Anchor cell of a two-/one-cell anchor (offsets in EMU).
#[derive(Debug, Clone, Copy, Default, PartialEq)]
struct AnchorCell {
    col: u32,
    col_off: f64,
    row: u32,
    row_off: f64,
}

/// Chart position as stored in the drawing part (EMU units).
#[derive(Debug, Clone, Copy, PartialEq)]
enum RawAnchor {
    Two { from: AnchorCell, to: AnchorCell },
    One { from: AnchorCell, cx: f64, cy: f64 },
    Abs { x: f64, y: f64, cx: f64, cy: f64 },
}

impl RawAnchor {
    /// Right-most column the anchor (roughly) touches, so column widths up
    /// to there are kept in the model and the geometry stays exact.
    fn max_col(&self) -> u32 {
        let est = |px: f64| (px / 4.0).ceil().clamp(0.0, MAX_COLS as f64) as u32;
        let c = match *self {
            RawAnchor::Two { to, .. } => to.col,
            RawAnchor::One { from, cx, .. } => from.col.saturating_add(est(cx / EMU_PER_PX)),
            RawAnchor::Abs { x, cx, .. } => est((x + cx) / EMU_PER_PX),
        };
        c.min(MAX_COLS - 1)
    }

    /// Model geometry (x, y, w, h) in grid px.
    fn to_grid(&self, m: &ColMetrics) -> (i64, i64, i64, i64) {
        let px = |emu: f64| emu / EMU_PER_PX;
        let cell_to_grid = |cell: &AnchorCell| -> (f64, f64) {
            let x = m.grid_x(cell.col) + (px(cell.col_off) * m.grid_px(cell.col) / m.excel_px(cell.col)).round();
            let y = cell.row as f64 * GRID_ROW_PX + (px(cell.row_off) * GRID_ROW_PX / EXCEL_ROW_PX).round();
            (x, y)
        };
        let (x, y, w, h) = match *self {
            RawAnchor::Two { from, to } => {
                let (x, y) = cell_to_grid(&from);
                let w = m.excel_x(to.col) + px(to.col_off) - (m.excel_x(from.col) + px(from.col_off));
                let h = (to.row as f64 * EXCEL_ROW_PX + px(to.row_off)) - (from.row as f64 * EXCEL_ROW_PX + px(from.row_off));
                (x, y, w, h)
            }
            RawAnchor::One { from, cx, cy } => {
                let (x, y) = cell_to_grid(&from);
                (x, y, px(cx), px(cy))
            }
            RawAnchor::Abs { x, y, cx, cy } => {
                // Absolute position in Excel px -> the cell it falls in -> grid.
                let (col, dx) = m.locate(px(x), false);
                let row = ((px(y) / EXCEL_ROW_PX).floor().max(0.0) as u64).min(MAX_ROWS as u64 - 1) as u32;
                let dy = px(y) - row as f64 * EXCEL_ROW_PX;
                let cell = AnchorCell { col, col_off: dx * EMU_PER_PX, row, row_off: dy * EMU_PER_PX };
                let (gx, gy) = cell_to_grid(&cell);
                (gx, gy, px(cx), px(cy))
            }
        };
        (x.round() as i64, y.round() as i64, w.round().max(1.0) as i64, h.round().max(1.0) as i64)
    }
}

/// Grid px position -> (row, col, x offset, y offset) in Excel px.
fn grid_to_excel_anchor(m: &ColMetrics, x: f64, y: f64) -> (u32, u16, u32, u32) {
    let (c, dx) = m.locate(x.max(0.0), true);
    let xoff = (dx * m.excel_px(c) / m.grid_px(c)).round().max(0.0) as u32;
    let y = y.max(0.0);
    let r = ((y / GRID_ROW_PX).floor() as u64).min(MAX_ROWS as u64 - 1) as u32;
    let yoff = ((y - r as f64 * GRID_ROW_PX) * EXCEL_ROW_PX / GRID_ROW_PX).round().max(0.0) as u32;
    (r, c as u16, xoff, yoff)
}

/// One anchor of a drawing part and what it contains.
#[derive(Debug, Default)]
struct DrawingAnchor {
    anchor: Option<RawAnchor>,
    /// r:ids of regular charts (c:chart).
    chart_rids: Vec<String>,
    /// r:ids of charts we cannot read (chartex: waterfall, treemap, ...).
    other_chart_rids: Vec<String>,
    other_charts: usize,
    picture: bool,
    shape: bool,
}

fn xml_f64_attr(e: &quick_xml::events::BytesStart, k: &[u8]) -> f64 {
    attr_value(e, k).and_then(|v| v.trim().parse::<f64>().ok()).filter(|f| f.is_finite()).unwrap_or(0.0).max(0.0)
}

/// Parse xdr:wsDr: every two-cell / one-cell / absolute anchor with its
/// position and content. Malformed XML ends the scan (anchors read so far
/// are kept).
fn parse_drawing_xml(xml: &[u8]) -> Vec<DrawingAnchor> {
    use quick_xml::events::Event;
    #[derive(Clone, Copy, PartialEq)]
    enum Gd { None, Chart, ChartEx, Other }
    let mut out = Vec::new();
    let mut r = quick_xml::Reader::from_reader(xml);
    let mut buf = Vec::new();
    let mut stack: Vec<Vec<u8>> = Vec::new();
    // (stack length inside the anchor, kind: 2 two-cell, 1 one-cell, 0 absolute)
    let mut cur: Option<(usize, u8)> = None;
    let mut a = DrawingAnchor::default();
    let (mut from, mut to) = (AnchorCell::default(), AnchorCell::default());
    let (mut ext, mut pos) = ((0.0, 0.0), (0.0, 0.0));
    let (mut has_from, mut has_to, mut has_ext, mut has_pos) = (false, false, false, false);
    let mut fallback: Option<usize> = None;
    let mut gd = Gd::None;
    loop {
        let ev = match r.read_event_into(&mut buf) {
            Ok(ev) => ev,
            Err(_) => break,
        };
        match ev {
            Event::Eof => break,
            Event::Start(ref e) | Event::Empty(ref e) => {
                let is_start = matches!(ev, Event::Start(_));
                let name = e.local_name().as_ref().to_vec();
                match cur {
                    None => {
                        let kind = match name.as_slice() {
                            b"twoCellAnchor" => Some(2u8),
                            b"oneCellAnchor" => Some(1),
                            b"absoluteAnchor" => Some(0),
                            _ => None,
                        };
                        if let (Some(k), true) = (kind, is_start) {
                            cur = Some((stack.len() + 1, k));
                            a = DrawingAnchor::default();
                            from = AnchorCell::default();
                            to = AnchorCell::default();
                            (has_from, has_to, has_ext, has_pos) = (false, false, false, false);
                            fallback = None;
                            gd = Gd::None;
                        }
                    }
                    // Fallback content (e.g. the "chart not supported" shape
                    // behind a chartex) is a stand-in, not real content.
                    Some((depth, _)) if fallback.is_none() => {
                        if stack.len() == depth {
                            match name.as_slice() {
                                b"from" => has_from = true,
                                b"to" => has_to = true,
                                b"ext" => { ext = (xml_f64_attr(e, b"cx"), xml_f64_attr(e, b"cy")); has_ext = true; }
                                b"pos" => { pos = (xml_f64_attr(e, b"x"), xml_f64_attr(e, b"y")); has_pos = true; }
                                _ => {}
                            }
                        }
                        match name.as_slice() {
                            b"Fallback" if is_start => fallback = Some(stack.len() + 1),
                            b"pic" => a.picture = true,
                            b"sp" | b"cxnSp" | b"grpSp" | b"contentPart" => a.shape = true,
                            b"graphicData" => {
                                let uri = attr_value(e, b"uri").unwrap_or_default();
                                gd = if uri.ends_with("/drawingml/2006/chart") {
                                    Gd::Chart
                                } else if uri.contains("chartex") {
                                    a.other_charts += 1;
                                    Gd::ChartEx
                                } else {
                                    a.shape = true; // SmartArt, ink, tables, ...
                                    Gd::Other
                                };
                            }
                            b"chart" => {
                                if let Some(id) = attr_value(e, b"id") {
                                    match gd {
                                        Gd::Chart => a.chart_rids.push(id),
                                        Gd::ChartEx => a.other_chart_rids.push(id),
                                        _ => {}
                                    }
                                }
                            }
                            _ => {}
                        }
                    }
                    _ => {}
                }
                if is_start {
                    if stack.len() >= MAX_XML_DEPTH { break; }
                    stack.push(name);
                }
            }
            Event::End(_) => {
                stack.pop();
                if fallback.is_some_and(|fd| stack.len() < fd) { fallback = None; }
                if let Some((depth, kind)) = cur {
                    if stack.len() < depth {
                        a.anchor = match kind {
                            2 if has_from && has_to => Some(RawAnchor::Two { from, to }),
                            1 if has_from && has_ext => Some(RawAnchor::One { from, cx: ext.0, cy: ext.1 }),
                            0 if has_pos && has_ext => Some(RawAnchor::Abs { x: pos.0, y: pos.1, cx: ext.0, cy: ext.1 }),
                            _ => None,
                        };
                        out.push(std::mem::take(&mut a));
                        cur = None;
                        if out.len() >= MAX_ANCHORS { break; }
                    }
                }
            }
            Event::Text(ref t) => {
                if let Some((depth, _)) = cur {
                    if stack.len() == depth + 2 {
                        let txt = t.unescape().map(|s| s.trim().to_string()).unwrap_or_default();
                        let cell = match stack[depth].as_slice() {
                            b"from" => Some(&mut from),
                            b"to" => Some(&mut to),
                            _ => None,
                        };
                        if let Some(cell) = cell {
                            let num = || txt.parse::<f64>().ok().filter(|f| f.is_finite()).unwrap_or(0.0).max(0.0);
                            match stack[depth + 1].as_slice() {
                                b"col" => cell.col = (num() as u64).min(MAX_COLS as u64 - 1) as u32,
                                b"row" => cell.row = (num() as u64).min(MAX_ROWS as u64 - 1) as u32,
                                b"colOff" => cell.col_off = num(),
                                b"rowOff" => cell.row_off = num(),
                                _ => {}
                            }
                        }
                    }
                }
            }
            _ => {}
        }
        buf.clear();
    }
    out
}

/// What a chart part says (before checking it fits the model).
#[derive(Debug, Default)]
struct ChartInfo {
    /// Model type of the first chart group; None when unsupported/missing.
    kind: Option<&'static str>,
    title: Option<String>,
    /// (role, formula) for every series reference: tx / cat / val / xVal / yVal.
    refs: Vec<(String, String)>,
}

/// Chart group element -> Some(model type) / Some(None) for unsupported groups.
fn chart_group_kind(name: &[u8]) -> Option<Option<&'static str>> {
    Some(match name {
        b"barChart" | b"bar3DChart" => Some("column"),
        b"lineChart" | b"line3DChart" => Some("line"),
        b"pieChart" | b"pie3DChart" | b"doughnutChart" => Some("pie"),
        b"scatterChart" => Some("scatter"),
        b"areaChart" | b"area3DChart" => Some("area"),
        n if n.ends_with(b"Chart") => None, // radar, stock, bubble, surface, ofPie
        _ => return None,
    })
}

/// Parse a c:chartSpace part. None when the XML is malformed.
fn parse_chart_xml(xml: &[u8]) -> Option<ChartInfo> {
    use quick_xml::events::Event;
    let mut info = ChartInfo::default();
    let mut r = quick_xml::Reader::from_reader(xml);
    let mut buf = Vec::new();
    let mut stack: Vec<Vec<u8>> = Vec::new();
    let mut group_seen = false;
    let mut bar_group_depth: Option<usize> = None;
    let (mut rich, mut cached) = (String::new(), String::new());
    loop {
        let ev = r.read_event_into(&mut buf).ok()?;
        match ev {
            Event::Eof => break,
            Event::Start(ref e) | Event::Empty(ref e) => {
                let is_start = matches!(ev, Event::Start(_));
                let name = e.local_name().as_ref().to_vec();
                if !group_seen && stack.last().is_some_and(|p| p == b"plotArea") {
                    if let Some(kind) = chart_group_kind(&name) {
                        group_seen = true;
                        info.kind = kind;
                        if name == b"barChart" || name == b"bar3DChart" { bar_group_depth = Some(stack.len() + 1); }
                    }
                }
                if name == b"barDir" && bar_group_depth == Some(stack.len()) {
                    info.kind = Some(if attr_value(e, b"val").as_deref() == Some("bar") { "bar" } else { "column" });
                }
                if is_start {
                    if stack.len() >= MAX_XML_DEPTH { return None; }
                    stack.push(name);
                }
            }
            Event::End(_) => { stack.pop(); }
            Event::Text(ref t) if !stack.is_empty() => {
                let last = &stack[stack.len() - 1];
                let in_title = stack.len() >= 3 && stack[1] == b"chart" && stack[2] == b"title";
                let is_f = last == b"f";
                if (in_title && (last == b"t" || last == b"v")) || is_f {
                    let text = t.unescape().ok()?.into_owned();
                    if in_title && last == b"t" {
                        rich.push_str(&text);
                    } else if in_title && last == b"v" && stack.iter().any(|n| n == b"strCache") {
                        cached.push_str(&text);
                    } else if is_f {
                        let role = stack.iter().position(|n| n == b"ser").and_then(|i| stack.get(i + 1));
                        if let Some(role) = role {
                            if matches!(role.as_slice(), b"tx" | b"cat" | b"val" | b"xVal" | b"yVal") {
                                if info.refs.len() >= MAX_CHART_REFS { return None; }
                                info.refs.push((String::from_utf8_lossy(role).into_owned(), text));
                            }
                        }
                    }
                }
            }
            _ => {}
        }
        buf.clear();
    }
    info.title = [rich, cached].into_iter().find(|t| !t.trim().is_empty());
    Some(info)
}

/// One area of a chart formula: "Sheet1!$B$1:$B$5", "'My Sheet'!$A$2".
fn parse_chart_ref(p: &str) -> Option<(Option<String>, Rect)> {
    let p = p.trim();
    let mut quoted = false;
    let mut bang = None;
    for (i, ch) in p.char_indices() {
        match ch {
            '\'' => quoted = !quoted,
            '!' if !quoted => bang = Some(i),
            _ => {}
        }
    }
    let (sheet, range) = match bang {
        Some(i) => (Some(p[..i].trim()), &p[i + 1..]),
        None => (None, p),
    };
    let sheet = match sheet {
        None => None,
        Some(s) => {
            let name = if s.len() >= 2 && s.starts_with('\'') && s.ends_with('\'') {
                s[1..s.len() - 1].replace("''", "'")
            } else {
                s.to_string()
            };
            // External workbooks ("[1]Sheet1") are not ours.
            if name.is_empty() || name.contains('[') || name.contains(']') { return None; }
            Some(name)
        }
    };
    Some((sheet, parse_range_ref(range.trim())?))
}

/// A chart series formula, possibly a union "(Sheet1!$A$1,Sheet1!$A$3)".
fn parse_chart_formula(f: &str) -> Option<Vec<(Option<String>, Rect)>> {
    let f = f.trim();
    let f = f.strip_prefix('=').unwrap_or(f).trim();
    let inner = f.strip_prefix('(').and_then(|s| s.strip_suffix(')')).unwrap_or(f);
    let mut parts = Vec::new();
    let mut cur = String::new();
    let mut quoted = false;
    for ch in inner.chars() {
        match ch {
            '\'' => { quoted = !quoted; cur.push(ch); }
            ',' if !quoted => {
                parts.push(std::mem::take(&mut cur));
                if parts.len() > 256 { return None; }
            }
            _ => cur.push(ch),
        }
    }
    parts.push(cur);
    parts.iter().map(|p| parse_chart_ref(p)).collect()
}

/// Bounding box of all series references, or None when the chart cannot be
/// represented (no values, refs on another sheet, unparsable refs, no data
/// row below the header).
fn chart_bbox(refs: &[(String, String)], host: &str) -> Option<Rect> {
    if !refs.iter().any(|(role, _)| role == "val" || role == "yVal") { return None; }
    let host = host.to_lowercase();
    let mut bb: Option<Rect> = None;
    for (_, f) in refs {
        for (sheet, ((r1, c1), (r2, c2))) in parse_chart_formula(f)? {
            if sheet.is_some_and(|s| s.to_lowercase() != host) { return None; }
            bb = Some(match bb {
                None => ((r1, c1), (r2, c2)),
                Some(((a1, b1), (a2, b2))) => ((a1.min(r1), b1.min(c1)), (a2.max(r2), b2.max(c2))),
            });
        }
    }
    bb.filter(|((r1, _), (r2, _))| r2 > r1)
}

/// Model JSON for a chart read from the file.
fn chart_to_json(c: &ChartRead, m: &ColMetrics) -> Value {
    let (x, y, w, h) = c.anchor.to_grid(m);
    let ((r1, c1), (r2, c2)) = c.range;
    let mut o = serde_json::Map::new();
    o.insert("type".into(), json!(c.kind));
    if let Some(t) = &c.title { o.insert("title".into(), json!(t)); }
    o.insert("range".into(), json!(format!("{}:{}", cell_addr(r1, c1), cell_addr(r2, c2))));
    o.insert("x".into(), json!(x));
    o.insert("y".into(), json!(y));
    o.insert("w".into(), json!(w));
    o.insert("h".into(), json!(h));
    Value::Object(o)
}

/// A model chart validated for writing.
#[derive(Debug)]
struct ChartSpec {
    kind: &'static str,
    title: Option<String>,
    range: Rect,
    x: f64,
    y: f64,
    w: f64,
    h: f64,
}

fn json_f64(v: Option<&Value>) -> Option<f64> {
    match v? {
        Value::Number(n) => n.as_f64(),
        Value::String(s) => s.trim().parse().ok(),
        _ => None,
    }
    .filter(|f: &f64| f.is_finite())
}

/// Validate one model chart. Err = plain-language reason it is skipped.
fn chart_spec_from_json(v: &Value) -> Result<ChartSpec, String> {
    let o = v.as_object().ok_or("it is not a chart description")?;
    let ty = o.get("type").and_then(Value::as_str).unwrap_or("").trim().to_ascii_lowercase();
    let kind = ["column", "bar", "line", "pie", "scatter", "area"].into_iter().find(|k| *k == ty)
        .ok_or_else(|| format!("unknown chart type {ty:?}"))?;
    let range = o.get("range").and_then(Value::as_str).map(str::trim).filter(|r| !r.is_empty())
        .ok_or("it has no data range")?;
    let rect = parse_range_ref(range).ok_or_else(|| format!("invalid data range {range:?}"))?;
    let ((r1, c1), (r2, c2)) = rect;
    if r2 <= r1 { return Err(format!("the data range {range} needs a header row and at least one row of data")); }
    let series = if c2 > c1 { c2 - c1 } else { 1 };
    if series > MAX_CHART_SERIES {
        return Err(format!("it has {series} data series (Excel allows at most {MAX_CHART_SERIES})"));
    }
    if kind == "scatter" && c2 == c1 {
        return Err("a scatter chart needs two columns (x values and y values)".into());
    }
    let title = o.get("title").and_then(Value::as_str).filter(|t| !t.trim().is_empty()).map(str::to_string);
    let size = |k: &str, def: f64| match json_f64(o.get(k)) {
        Some(v) => v.clamp(MIN_CHART_PX, MAX_CHART_PX),
        None => def,
    };
    Ok(ChartSpec {
        kind,
        title,
        range: rect,
        x: json_f64(o.get("x")).unwrap_or(0.0).clamp(0.0, MAX_CHART_X),
        y: json_f64(o.get("y")).unwrap_or(0.0).clamp(0.0, MAX_CHART_Y),
        w: size("w", DEFAULT_CHART_W),
        h: size("h", DEFAULT_CHART_H),
    })
}

/// rust_xlsxwriter treats "Name!A1"-like titles as cell references; a
/// trailing space keeps such a title literal.
fn literal_chart_title(t: &str) -> String {
    if let Some(i) = t.find('!') {
        let rest = t[i + 1..].replace('$', "");
        if rest.starts_with(|c: char| c.is_ascii_uppercase())
            && rest.ends_with(|c: char| c.is_ascii_digit())
            && rest.chars().all(|c| c.is_ascii_uppercase() || c.is_ascii_digit() || c == ':')
        {
            return format!("{t} ");
        }
    }
    t.to_string()
}

/// Native Excel chart for a model chart: header row = series names, first
/// column = categories (x values for scatter), one series per further column.
fn build_chart(spec: &ChartSpec, sheet: &str) -> Chart {
    let ty = match spec.kind {
        "column" => ChartType::Column,
        "bar" => ChartType::Bar,
        "line" => ChartType::Line,
        "pie" => ChartType::Pie,
        "scatter" => ChartType::Scatter,
        _ => ChartType::Area,
    };
    let mut chart = Chart::new(ty);
    let ((r1, c1), (r2, c2)) = spec.range;
    let (c1, c2) = (c1 as u16, c2 as u16);
    if c2 > c1 {
        for c in c1 + 1..=c2 {
            chart.add_series()
                .set_categories((sheet, r1 + 1, c1, r2, c1))
                .set_values((sheet, r1 + 1, c, r2, c))
                .set_name((sheet, r1, c));
        }
    } else {
        chart.add_series().set_values((sheet, r1 + 1, c1, r2, c1)).set_name((sheet, r1, c1));
    }
    match &spec.title {
        Some(t) => { chart.title().set_name(literal_chart_title(t).as_str()); }
        // The model shows no title; stop Excel from adding an automatic one.
        None => { chart.title().set_hidden(); }
    }
    chart.set_width(spec.w.round() as u32).set_height(spec.h.round() as u32);
    chart
}

/// Frozen panes, auto filter and charts of one sheet. Invalid entries are
/// skipped with a warning; they never fail the save.
fn write_sheet_features(ws: &mut Worksheet, sheet: &Sheet, sheet_name: &str, warn: &mut Warnings) {
    if let Some(f) = sheet.freeze.filter(|f| f.rows > 0 || f.cols > 0) {
        let res = if f.rows >= MAX_ROWS || f.cols >= MAX_COLS {
            Err("outside the sheet".to_string())
        } else {
            ws.set_freeze_panes(f.rows, f.cols as u16).map(|_| ()).map_err(|e| e.to_string())
        };
        if let Err(e) = res {
            warn.push(format!("Frozen panes on sheet \"{sheet_name}\" ({} rows, {} columns) were not saved: {e}.", f.rows, f.cols));
        }
    }
    if let Some(af) = sheet.auto_filter.as_deref().map(str::trim).filter(|a| !a.is_empty()) {
        let res = match parse_range_ref(af) {
            Some(((r1, c1), (r2, c2))) => ws.autofilter(r1, c1 as u16, r2, c2 as u16).map(|_| ()).map_err(|e| e.to_string()),
            None => Err("it is not a valid range".to_string()),
        };
        if let Err(e) = res {
            warn.push(format!("The filter range {af:?} on sheet \"{sheet_name}\" was not saved: {e}."));
        }
    }
    if sheet.charts.is_empty() { return; }
    let m = ColMetrics::new(sheet.col_widths.as_ref());
    for (i, v) in sheet.charts.iter().enumerate() {
        if i >= MAX_CHARTS {
            warn.push(format!("Only the first {MAX_CHARTS} charts on sheet \"{sheet_name}\" were saved."));
            break;
        }
        let spec = match chart_spec_from_json(v) {
            Ok(s) => s,
            Err(why) => {
                warn.push(format!("Chart {} on sheet \"{sheet_name}\" was not saved: {why}.", i + 1));
                continue;
            }
        };
        let chart = build_chart(&spec, sheet_name);
        let (row, col, xoff, yoff) = grid_to_excel_anchor(&m, spec.x, spec.y);
        if let Err(e) = ws.insert_chart_with_offset(row, col, &chart, xoff, yoff) {
            warn.push(format!("Chart {} on sheet \"{sheet_name}\" was not saved: {e}.", i + 1));
        }
    }
}

// =====================================================================
// Writing
// =====================================================================

/// Make a sheet name legal for Excel: no []:*?/\ (replaced by "_"), no
/// leading/trailing apostrophe, at most 31 characters, not blank.
/// Returns "" for names that end up empty (caller substitutes "SheetN").
#[allow(dead_code)] // public API; wired by main.rs / used by tests
pub fn sanitize_sheet_name(name: &str) -> String {
    let cleaned: String = name.chars()
        .map(|c| if matches!(c, '[' | ']' | ':' | '*' | '?' | '/' | '\\') || c.is_control() { '_' } else { c })
        .collect();
    let trimmed = cleaned.trim().trim_matches('\'').trim();
    trimmed.chars().take(31).collect::<String>().trim_end().to_string()
}

/// Sanitize and de-duplicate (case-insensitively) all sheet names.
fn sanitize_sheet_names(sheets: &[Sheet], warn: &mut Warnings) -> Vec<String> {
    let mut used: BTreeSet<String> = BTreeSet::new();
    let mut out = Vec::with_capacity(sheets.len());
    for (i, s) in sheets.iter().enumerate() {
        let mut base = sanitize_sheet_name(&s.name);
        if base.is_empty() { base = format!("Sheet{}", i + 1); }
        let mut candidate = base.clone();
        let mut n = 2;
        while used.contains(&candidate.to_lowercase()) {
            let suffix = format!(" ({n})");
            let keep = 31 - suffix.chars().count();
            candidate = format!("{}{}", base.chars().take(keep).collect::<String>().trim_end(), suffix);
            n += 1;
        }
        used.insert(candidate.to_lowercase());
        if candidate != s.name {
            warn.push(format!(
                "Sheet name \"{}\" is not allowed in Excel; it was saved as \"{}\" (formulas that refer to the old name may need updating).",
                s.name, candidate
            ));
        }
        out.push(candidate);
    }
    out
}

/// "#RRGGBB" / "RRGGBB" / "#RGB" -> 0xRRGGBB.
fn parse_hex_color(s: &str) -> Option<u32> {
    let h = s.trim().trim_start_matches('#');
    let h: String = match h.len() {
        3 => h.chars().flat_map(|c| [c, c]).collect(),
        6 => h.to_string(),
        8 => h[2..].to_string(), // ARGB
        _ => return None,
    };
    u32::from_str_radix(&h, 16).ok()
}

/// Builds and caches rust_xlsxwriter Formats keyed by (z, style).
#[derive(Default)]
struct FormatCache {
    map: HashMap<String, Format>,
}

impl FormatCache {
    /// Format for a cell, or None when the cell needs no formatting.
    fn get(&mut self, z: Option<&str>, s: Option<&CellStyle>, warn: &mut Warnings) -> Option<Format> {
        let z = z.map(str::trim).filter(|z| !z.is_empty() && !z.eq_ignore_ascii_case("General"));
        let s = s.filter(|s| !s.is_empty());
        if z.is_none() && s.is_none() { return None; }
        let key = format!("{:?}|{:?}", z, s);
        if let Some(f) = self.map.get(&key) { return Some(f.clone()); }
        let mut f = Format::new();
        if let Some(z) = z { f = f.set_num_format(z); }
        if let Some(s) = s {
            if s.b { f = f.set_bold(); }
            if s.i { f = f.set_italic(); }
            if s.u { f = f.set_underline(FormatUnderline::Single); }
            if s.wrap { f = f.set_text_wrap(); }
            if let Some(c) = &s.color {
                match parse_hex_color(c) {
                    Some(rgb) => f = f.set_font_color(Color::RGB(rgb)),
                    None => warn.push(format!("Ignored invalid text colour {c:?}.")),
                }
            }
            if let Some(c) = &s.bg {
                match parse_hex_color(c) {
                    Some(rgb) => f = f.set_background_color(Color::RGB(rgb)),
                    None => warn.push(format!("Ignored invalid fill colour {c:?}.")),
                }
            }
            match s.align.as_deref().map(str::to_ascii_lowercase).as_deref() {
                Some("left") => f = f.set_align(FormatAlign::Left),
                Some("center") | Some("centre") => f = f.set_align(FormatAlign::Center),
                Some("right") => f = f.set_align(FormatAlign::Right),
                Some("justify") => f = f.set_align(FormatAlign::Justify),
                _ => {}
            }
        }
        self.map.insert(key, f.clone());
        Some(f)
    }
}

/// Number text for formula cached results / CSV export.
fn fmt_number(n: f64) -> String {
    if n.fract() == 0.0 && n.abs() < 1e15 { format!("{}", n as i64) } else { format!("{n}") }
}

fn value_as_f64(v: &Value) -> Option<f64> {
    match v {
        Value::Number(n) => n.as_f64(),
        Value::String(s) => {
            let t = s.trim();
            if is_number_str(t) { t.parse().ok().filter(|f: &f64| f.is_finite()) } else { None }
        }
        Value::Bool(b) => Some(if *b { 1.0 } else { 0.0 }),
        _ => None,
    }
}

fn value_as_bool(v: &Value) -> Option<bool> {
    match v {
        Value::Bool(b) => Some(*b),
        Value::Number(n) => n.as_f64().map(|f| f != 0.0),
        Value::String(s) if s.trim().eq_ignore_ascii_case("true") => Some(true),
        Value::String(s) if s.trim().eq_ignore_ascii_case("false") => Some(false),
        _ => None,
    }
}

/// Default number format for a date-like value without an explicit z.
fn default_date_format(kind: IsoKind) -> &'static str {
    match kind {
        IsoKind::Date => "yyyy-mm-dd",
        IsoKind::DateTime => "yyyy-mm-dd hh:mm:ss",
        IsoKind::Time => "hh:mm:ss",
        IsoKind::Duration => "[h]:mm:ss",
    }
}

/// Truncate a string to Excel's 32,767-character cell limit (with warning).
fn excel_string<'a>(s: &'a str, addr: &str, sheet: &str, warn: &mut Warnings) -> std::borrow::Cow<'a, str> {
    if s.chars().count() <= MAX_STRING_CHARS { return std::borrow::Cow::Borrowed(s); }
    warn.push(format!(
        "Cell {sheet}!{addr} has more than {MAX_STRING_CHARS} characters (Excel's limit); the text was cut off."
    ));
    std::borrow::Cow::Owned(s.chars().take(MAX_STRING_CHARS).collect())
}

/// The cached result of a formula cell as the string rust_xlsxwriter stores
/// in <v>, plus the date kind when the result is a date (needs a date format).
fn formula_result(cell: &Cell) -> (Option<String>, Option<IsoKind>) {
    match (&cell.v, cell.t.as_str()) {
        (Value::Null, _) => (None, None),
        (Value::String(s), "d") => match iso_to_serial(s) {
            Some((serial, kind)) => (Some(fmt_number(serial)), Some(kind)),
            None => (Some(s.clone()), None),
        },
        (Value::Bool(b), _) => (Some(if *b { "TRUE" } else { "FALSE" }.into()), None),
        (Value::Number(n), _) => (n.as_f64().map(fmt_number), None),
        (Value::String(s), _) => (Some(s.clone()), None),
        (other, _) => (Some(other.to_string()), None),
    }
}

#[allow(clippy::too_many_arguments)]
fn write_cell(
    ws: &mut Worksheet,
    r: u32,
    c: u16,
    addr: &str,
    sheet_name: &str,
    cell: &Cell,
    formats: &mut FormatCache,
    warn: &mut Warnings,
) -> Result<(), String> {
    let e = |err: rust_xlsxwriter::XlsxError| format!("cell {sheet_name}!{addr}: {err}");
    let style = cell.s.as_ref();

    // Formula (with cached result so non-recalculating readers see values).
    if let Some(f) = cell.f.as_deref().map(str::trim).filter(|f| !f.is_empty() && *f != "=") {
        let text = if f.starts_with('=') { f.to_string() } else { format!("={f}") };
        let (result, date_kind) = formula_result(cell);
        let z = cell.z.as_deref().or(date_kind.map(default_date_format));
        let mut formula = Formula::new(text);
        if let Some(res) = result { formula = formula.set_result(res); }
        match formats.get(z, style, warn) {
            Some(fmt) => ws.write_formula_with_format(r, c, formula, &fmt).map_err(e)?,
            None => ws.write_formula(r, c, formula).map_err(e)?,
        };
        return Ok(());
    }

    enum Out { Num(f64), Str(String), Bool(bool), Blank }
    let mut z: Option<String> = cell.z.clone();
    let out = match (cell.t.as_str(), &cell.v) {
        (_, Value::Null) => Out::Blank,
        ("n", v) => match value_as_f64(v) {
            Some(n) => Out::Num(n),
            None => Out::Str(cell_display(v)),
        },
        ("b", v) => match value_as_bool(v) {
            Some(b) => Out::Bool(b),
            None => Out::Str(cell_display(v)),
        },
        ("d", Value::String(s)) => match iso_to_serial(s) {
            Some((serial, kind)) => {
                if z.is_none() { z = Some(default_date_format(kind).into()); }
                Out::Num(serial)
            }
            None => Out::Str(s.clone()),
        },
        ("d", Value::Number(n)) => {
            if z.is_none() { z = Some("yyyy-mm-dd".into()); }
            Out::Num(n.as_f64().unwrap_or(0.0))
        }
        // Error literals without a formula cannot be stored as errors by
        // rust_xlsxwriter; keep the text so nothing is lost.
        (_, Value::String(s)) => Out::Str(s.clone()),
        (_, Value::Number(n)) => Out::Num(n.as_f64().unwrap_or(0.0)),
        (_, Value::Bool(b)) => Out::Bool(*b),
        (_, other) => Out::Str(other.to_string()),
    };
    let fmt = formats.get(z.as_deref(), style, warn);
    match (out, fmt) {
        (Out::Num(n), Some(f)) => { ws.write_number_with_format(r, c, n, &f).map_err(e)?; }
        (Out::Num(n), None) => { ws.write_number(r, c, n).map_err(e)?; }
        (Out::Bool(b), Some(f)) => { ws.write_boolean_with_format(r, c, b, &f).map_err(e)?; }
        (Out::Bool(b), None) => { ws.write_boolean(r, c, b).map_err(e)?; }
        (Out::Str(s), fmt) => {
            let s = excel_string(&s, addr, sheet_name, warn);
            match fmt {
                Some(f) => ws.write_string_with_format(r, c, s.as_ref(), &f).map_err(e)?,
                None => ws.write_string(r, c, s.as_ref()).map_err(e)?,
            };
        }
        (Out::Blank, Some(f)) => { ws.write_blank(r, c, &f).map_err(e)?; }
        (Out::Blank, None) => {}
    }
    Ok(())
}

fn write_sheet(ws: &mut Worksheet, sheet: &Sheet, sheet_name: &str, formats: &mut FormatCache, warn: &mut Warnings) -> Result<(), String> {
    if let Some(widths) = &sheet.col_widths {
        for (col, w) in widths {
            let Some(ci) = col_to_index(col.trim()) else {
                warn.push(format!("Ignored column width for invalid column {col:?} on sheet \"{sheet_name}\"."));
                continue;
            };
            if !w.is_finite() || *w < 0.0 { continue; }
            // 0 hides the column (that is how hidden columns are read back).
            ws.set_column_width(ci as u16, w.min(255.0))
                .map_err(|e| format!("column {col} width on sheet \"{sheet_name}\": {e}"))?;
        }
    }

    // Merges first; the top-left cell's real value is written afterwards
    // (rust_xlsxwriter explicitly supports overwriting it).
    let mut merged: Vec<((u32, u32), (u32, u32))> = Vec::new();
    for m in &sheet.merges {
        let Some(((r1, c1), (r2, c2))) = parse_range_ref(m) else {
            warn.push(format!("Ignored invalid merged range {m:?} on sheet \"{sheet_name}\"."));
            continue;
        };
        if r1 == r2 && c1 == c2 { continue; }
        if merged.iter().any(|&((a1, b1), (a2, b2))| r1 <= a2 && a1 <= r2 && c1 <= b2 && b1 <= c2) {
            warn.push(format!("Merged range {m} on sheet \"{sheet_name}\" overlaps another merged range and was not merged."));
            continue;
        }
        let top_left = sheet.cells.get(&cell_addr(r1, c1));
        let fmt = formats
            .get(top_left.and_then(|c| c.z.as_deref()), top_left.and_then(|c| c.s.as_ref()), warn)
            .unwrap_or_default();
        match ws.merge_range(r1, c1 as u16, r2, c2 as u16, "", &fmt) {
            Ok(_) => merged.push(((r1, c1), (r2, c2))),
            Err(err) => warn.push(format!("Could not merge {m} on sheet \"{sheet_name}\": {err}.")),
        }
    }

    for (addr, cell) in &sheet.cells {
        let Some((r, c)) = parse_addr(addr) else {
            warn.push(format!("Skipped cell with invalid address {addr:?} on sheet \"{sheet_name}\"."));
            continue;
        };
        write_cell(ws, r, c as u16, addr, sheet_name, cell, formats, warn)?;
    }
    write_sheet_features(ws, sheet, sheet_name, warn);
    Ok(())
}

/// Write the whole book to .xlsx bytes. See `write_xlsx_with_warnings`.
/// (Frozen signature; the HTTP handler uses the `_with_warnings` variant.)
#[allow(dead_code)]
pub fn write_xlsx(book: &ExcelBook) -> Result<Vec<u8>, String> {
    write_xlsx_with_warnings(book).map(|(bytes, _)| bytes)
}

/// Write every sheet (hidden ones stay hidden, the active sheet stays
/// active) and return the bytes plus notes about anything that had to be
/// changed (renamed sheets, truncated text, skipped invalid input).
/// Hard failures (a cell that cannot be written) are errors, never dropped.
#[allow(dead_code)] // public API; wired by main.rs / used by tests
pub fn write_xlsx_with_warnings(book: &ExcelBook) -> Result<(Vec<u8>, Vec<String>), String> {
    let mut wb = Workbook::new();
    let mut warn = Warnings::default();
    let mut formats = FormatCache::default();
    let n = book.sheets.len();
    if n == 0 {
        wb.add_worksheet().set_name("Sheet1").map_err(|e| e.to_string())?;
        return wb.save_to_buffer().map(|b| (b, warn.list)).map_err(|e| format!("xlsx write failed: {e}"));
    }
    let names = sanitize_sheet_names(&book.sheets, &mut warn);
    let mut hidden: Vec<bool> = book.sheets.iter().map(|s| s.hidden).collect();
    let mut active = (book.active as usize).min(n - 1);
    if hidden.iter().all(|h| *h) {
        hidden[active] = false;
        warn.push(format!("Excel needs at least one visible sheet; \"{}\" was made visible.", names[active]));
    }
    if hidden[active] {
        active = hidden.iter().position(|h| !*h).unwrap_or(0);
    }
    for (i, sheet) in book.sheets.iter().enumerate() {
        let ws = wb.add_worksheet();
        ws.set_name(&names[i]).map_err(|e| format!("sheet name {:?}: {e}", names[i]))?;
        if hidden[i] { ws.set_hidden(true); }
        if i == active { ws.set_active(true); }
        write_sheet(ws, sheet, &names[i], &mut formats, &mut warn)?;
    }
    let bytes = wb.save_to_buffer().map_err(|e| format!("xlsx write failed: {e}"))?;
    Ok((bytes, warn.list))
}


fn cell_display(v: &Value) -> String {
    match v {
        Value::String(s) => s.clone(),
        Value::Null => String::new(),
        Value::Bool(b) => if *b { "TRUE".into() } else { "FALSE".into() },
        Value::Number(n) => n.as_f64().map(fmt_number).unwrap_or_else(|| n.to_string()),
        other => other.to_string(),
    }
}

// =====================================================================
// CSV / TSV
// =====================================================================

/// Streaming RFC 4180 record reader. A '"' opens a quoted field only at the
/// start of a field; elsewhere it is literal. Inside quotes, "" is a quote
/// and delimiters / CR / LF are data. Records end at CRLF, LF or CR.
struct CsvRecords<'a> {
    chars: std::iter::Peekable<std::str::Chars<'a>>,
    delim: char,
    done: bool,
    /// Set when the input ended inside a quoted field.
    unterminated: bool,
}

impl<'a> CsvRecords<'a> {
    fn new(text: &'a str, delim: char) -> Self {
        CsvRecords { chars: text.chars().peekable(), delim, done: false, unterminated: false }
    }
}

impl Iterator for CsvRecords<'_> {
    type Item = Vec<String>;
    fn next(&mut self) -> Option<Vec<String>> {
        if self.done { return None; }
        let mut rec = Vec::new();
        let mut field = String::new();
        let mut at_start = true;
        let mut in_q = false;
        let mut any = false;
        loop {
            let Some(ch) = self.chars.next() else {
                self.done = true;
                if in_q { self.unterminated = true; }
                if !any { return None; } // EOF right after a line break
                rec.push(field);
                return Some(rec);
            };
            any = true;
            if in_q {
                if ch == '"' {
                    if self.chars.peek() == Some(&'"') { field.push('"'); self.chars.next(); } else { in_q = false; }
                } else {
                    field.push(ch);
                }
                continue;
            }
            if ch == '"' && at_start { in_q = true; at_start = false; continue; }
            if ch == self.delim { rec.push(std::mem::take(&mut field)); at_start = true; continue; }
            if ch == '\r' || ch == '\n' {
                if ch == '\r' && self.chars.peek() == Some(&'\n') { self.chars.next(); }
                rec.push(field);
                return Some(rec);
            }
            field.push(ch);
            at_start = false;
        }
    }
}

/// CSV field -> cell. Only plain decimal numbers become numbers
/// ("-12", "3.5", "0.25"); anything with leading zeros ("007"), a percent
/// sign, thousands separators, exponent, leading "+"/".", more than 15
/// significant digits, or date/boolean/formula look stays TEXT, so a CSV
/// opened and saved again is byte-identical.
fn csv_scalar(s: &str) -> Cell {
    if is_plain_csv_number(s) {
        if let Ok(n) = s.parse::<f64>() {
            if n.is_finite() { return Cell::new(num_json(n), "n"); }
        }
    }
    Cell::new(json!(s), "s")
}

fn is_plain_csv_number(s: &str) -> bool {
    let b = s.as_bytes();
    let mut i = 0;
    if b.first() == Some(&b'-') { i += 1; }
    let int_start = i;
    while i < b.len() && b[i].is_ascii_digit() { i += 1; }
    let int_len = i - int_start;
    if int_len == 0 { return false; }
    if int_len > 1 && b[int_start] == b'0' { return false; } // leading zero
    let mut sig = if b[int_start] == b'0' { 0 } else { int_len };
    if i < b.len() && b[i] == b'.' {
        i += 1;
        let frac_start = i;
        while i < b.len() && b[i].is_ascii_digit() { i += 1; }
        if i == frac_start { return false; }
        sig += i - frac_start;
    }
    i == b.len() && sig <= 15
}

/// Parse delimited text into a one-sheet book (no encoding / sniffing; see
/// `parse_csv_bytes`). Empty fields produce no cell; blank lines keep their
/// row so positions are preserved.
pub fn parse_delimited(text: &str, delim: char, sheet_name: &str) -> ExcelBook {
    parse_delimited_limited(text, delim, sheet_name, MAX_CELLS)
}

fn parse_delimited_limited(text: &str, delim: char, sheet_name: &str, max_cells: usize) -> ExcelBook {
    let text = text.strip_prefix('\u{feff}').unwrap_or(text);
    let mut sheet = Sheet { name: sheet_name.into(), ..Default::default() };
    let mut warn = Warnings::default();
    let mut records = CsvRecords::new(text, delim);
    let mut budget = max_cells;
    let mut r: u32 = 0;
    'rows: for rec in records.by_ref() {
        if r >= MAX_ROWS {
            warn.push(format!("The file has more than {MAX_ROWS} rows (Excel's limit); the extra rows were not loaded. Saving would drop them."));
            break;
        }
        if rec.len() > MAX_COLS as usize {
            warn.push(format!("Some rows have more than {MAX_COLS} columns (Excel's limit); the extra columns were not loaded."));
        }
        for (c, field) in rec.iter().enumerate().take(MAX_COLS as usize) {
            if field.is_empty() { continue; }
            if budget == 0 {
                warn.push(format!(
                    "This file is too large to load completely: only the first {max_cells} cells were loaded. \
                     Saving would permanently drop the rest — do not overwrite the original file."
                ));
                break 'rows;
            }
            budget -= 1;
            sheet.cells.insert(cell_addr(r, c as u32), csv_scalar(field));
        }
        r += 1;
    }
    if records.unterminated {
        warn.push("The file ends inside a quoted field (a quote is missing); the rest of the file was read into one cell.");
    }
    ExcelBook { kind: "excel".into(), active: 0, sheets: vec![sheet], warnings: warn.list }
}

/// Windows-1252 code points for bytes 0x80..=0x9F (others map to U+00xx).
const CP1252_HIGH: [u16; 32] = [
    0x20AC, 0x0081, 0x201A, 0x0192, 0x201E, 0x2026, 0x2020, 0x2021, 0x02C6, 0x2030, 0x0160, 0x2039,
    0x0152, 0x008D, 0x017D, 0x008F, 0x0090, 0x2018, 0x2019, 0x201C, 0x201D, 0x2022, 0x2013, 0x2014,
    0x02DC, 0x2122, 0x0161, 0x203A, 0x0153, 0x009D, 0x017E, 0x0178,
];

fn decode_cp1252(bytes: &[u8]) -> String {
    bytes.iter().map(|&b| match b {
        0x80..=0x9F => char::from_u32(CP1252_HIGH[(b - 0x80) as usize] as u32).unwrap_or('\u{fffd}'),
        _ => b as char,
    }).collect()
}

fn decode_utf16(bytes: &[u8], le: bool) -> String {
    let units = bytes.chunks_exact(2).map(|p| if le { u16::from_le_bytes([p[0], p[1]]) } else { u16::from_be_bytes([p[0], p[1]]) });
    char::decode_utf16(units).map(|r| r.unwrap_or('\u{fffd}')).collect()
}

/// Decode text bytes: UTF-8 (± BOM), UTF-16 LE/BE (BOM, or BOM-less when
/// the NUL pattern is unambiguous), else Windows-1252 (Excel's "ANSI" CSV).
/// Returns the text and an optional note about the encoding used.
#[allow(dead_code)] // public API; wired by main.rs / used by tests
pub fn decode_text(bytes: &[u8]) -> (String, Option<String>) {
    if let Some(rest) = bytes.strip_prefix(&[0xEF, 0xBB, 0xBF]) {
        return match std::str::from_utf8(rest) {
            Ok(s) => (s.to_string(), None),
            Err(_) => (String::from_utf8_lossy(rest).into_owned(),
                       Some("The file contains invalid UTF-8; some characters were replaced.".into())),
        };
    }
    if let Some(rest) = bytes.strip_prefix(&[0xFF, 0xFE]) { return (decode_utf16(rest, true), None); }
    if let Some(rest) = bytes.strip_prefix(&[0xFE, 0xFF]) { return (decode_utf16(rest, false), None); }
    // BOM-less UTF-16: mostly-ASCII text has a NUL in every other byte.
    let sample = &bytes[..bytes.len().min(4096) & !1];
    if sample.len() >= 4 {
        let pairs = sample.len() / 2;
        let odd_nul = sample.iter().skip(1).step_by(2).filter(|&&b| b == 0).count();
        let even_nul = sample.iter().step_by(2).filter(|&&b| b == 0).count();
        if odd_nul * 10 >= pairs * 4 && even_nul * 20 < pairs { return (decode_utf16(bytes, true), None); }
        if even_nul * 10 >= pairs * 4 && odd_nul * 20 < pairs { return (decode_utf16(bytes, false), None); }
    }
    match std::str::from_utf8(bytes) {
        Ok(s) => (s.to_string(), None),
        Err(_) => (decode_cp1252(bytes),
                   Some("The file is not UTF-8; it was read as Windows-1252 (Western European). It will be saved as UTF-8.".into())),
    }
}

/// Guess the delimiter from the first records: the candidate that yields
/// the most rows with the same (>1) field count wins; ties favour , ; tab |.
#[allow(dead_code)] // public API; wired by main.rs / used by tests
pub fn sniff_delimiter(text: &str) -> char {
    let mut end = text.len().min(64 * 1024);
    while !text.is_char_boundary(end) { end -= 1; }
    let sample = &text[..end];
    let truncated = end < text.len();
    let mut best = (',', 0usize, 0usize);
    for cand in [',', ';', '\t', '|'] {
        let mut recs: Vec<usize> = CsvRecords::new(sample, cand).take(51)
            .filter(|r| !(r.len() == 1 && r[0].is_empty()))
            .map(|r| r.len())
            .collect();
        if truncated && recs.len() > 1 { recs.pop(); } // last record may be cut
        let mut freq: HashMap<usize, usize> = HashMap::new();
        for n in &recs { *freq.entry(*n).or_insert(0) += 1; }
        let Some((&mode, &count)) = freq.iter().filter(|(n, _)| **n > 1).max_by_key(|(n, c)| (**c, **n)) else { continue };
        if (count, mode) > (best.1, best.2) { best = (cand, count, mode); }
    }
    best.0
}

/// CONTRACT (called by main.rs for .csv/.tsv uploads): decode bytes (UTF-8
/// with/without BOM, UTF-16 LE/BE, else Windows-1252), honour an Excel
/// "sep=X" first line, sniff the delimiter when `delim` is None
/// (comma / semicolon / tab / pipe), and parse into a one-sheet book named
/// `name`. Encoding / size notes go to `book.warnings`.
#[allow(dead_code)] // public API; wired by main.rs / used by tests
pub fn parse_csv_bytes(bytes: &[u8], delim: Option<char>, name: &str) -> ExcelBook {
    let (text, note) = decode_text(bytes);
    let mut body: &str = text.strip_prefix('\u{feff}').unwrap_or(&text);
    let mut sep_line: Option<char> = None;
    if body.len() >= 5 && body.get(..4).is_some_and(|p| p.eq_ignore_ascii_case("sep=")) {
        let mut it = body[4..].chars();
        if let Some(ch) = it.next() {
            let rest = it.as_str();
            if rest.is_empty() || rest.starts_with('\n') || rest.starts_with('\r') {
                sep_line = Some(ch);
                body = rest.strip_prefix("\r\n").or_else(|| rest.strip_prefix('\n')).or_else(|| rest.strip_prefix('\r')).unwrap_or(rest);
            }
        }
    }
    let d = delim.or(sep_line).unwrap_or_else(|| sniff_delimiter(body));
    let mut book = parse_delimited(body, d, name);
    if let Some(n) = note { book.warnings.insert(0, n); }
    book
}

/// Mirror of parseCellInput() in lib/model-excel.js: interpret what a user
/// TYPED into a cell ("=f" formula, TRUE/FALSE, numbers with thousands
/// separators, "12%" -> 0.12 with a percent format, ISO dates).
#[allow(dead_code)] // public API; wired by main.rs / used by tests
pub fn parse_scalar(s: &str) -> Cell {
    if s.starts_with('=') {
        return Cell { v: Value::Null, t: "f".into(), f: Some(s.to_string()), z: None, s: None };
    }
    let trimmed = s.trim();
    if trimmed.is_empty() { return Cell::new(Value::Null, "s"); }
    if trimmed.eq_ignore_ascii_case("TRUE") { return Cell::new(json!(true), "b"); }
    if trimmed.eq_ignore_ascii_case("FALSE") { return Cell::new(json!(false), "b"); }
    let (body, pct) = match trimmed.strip_suffix('%') {
        Some(b) => (b.trim(), true),
        None => (trimmed, false),
    };
    let numeric = body.replace(',', "");
    if is_number_str(&numeric) {
        if let Ok(n) = numeric.parse::<f64>() {
            if n.is_finite() {
                if pct {
                    let decimals = numeric.split_once('.').map(|(_, f)| f.len()).unwrap_or(0);
                    let mut c = Cell::new(json!(n / 100.0), "n");
                    c.z = Some(if decimals == 0 { "0%".into() } else { format!("0.{}%", "0".repeat(decimals.min(10))) });
                    return c;
                }
                return Cell::new(num_json(n), "n");
            }
        }
    }
    if let Some((_, kind)) = iso_to_serial(trimmed) {
        if matches!(kind, IsoKind::Date | IsoKind::DateTime) {
            if let Some(iso) = normalize_iso_datetime(trimmed) { return Cell::new(json!(iso), "d"); }
        }
    }
    Cell::new(json!(s), "s")
}

/// Strict numeric literal check: [+-]?digits[.digits][eE[+-]digits]
fn is_number_str(s: &str) -> bool {
    let b = s.as_bytes();
    let mut i = 0;
    if i < b.len() && (b[i] == b'+' || b[i] == b'-') { i += 1; }
    let mut digits = 0;
    while i < b.len() && b[i].is_ascii_digit() { i += 1; digits += 1; }
    if i < b.len() && b[i] == b'.' {
        i += 1;
        while i < b.len() && b[i].is_ascii_digit() { i += 1; digits += 1; }
    }
    if digits == 0 { return false; }
    if i < b.len() && (b[i] == b'e' || b[i] == b'E') {
        i += 1;
        if i < b.len() && (b[i] == b'+' || b[i] == b'-') { i += 1; }
        let mut ed = 0;
        while i < b.len() && b[i].is_ascii_digit() { i += 1; ed += 1; }
        if ed == 0 { return false; }
    }
    i == b.len()
}

/// Serialize one sheet to CSV/TSV text. The output always starts at A1 (a
/// sheet whose data starts at B2 keeps its blank first row/column), cells
/// are exported by VALUE (formulas -> cached result, booleans TRUE/FALSE,
/// dates "YYYY-MM-DD[ HH:MM:SS]"), fields are quoted when they contain the
/// delimiter, a quote, CR or LF. A UTF-8 BOM is prepended when the text is
/// not pure ASCII so Excel detects the encoding.
pub fn sheet_to_delimited(book: &ExcelBook, sheet_idx: usize, delim: char) -> String {
    let Some(sheet) = book.sheets.get(sheet_idx) else { return String::new() };
    let mut rows: BTreeMap<u32, BTreeMap<u32, String>> = BTreeMap::new();
    let mut max_c = 0u32;
    for (addr, cell) in &sheet.cells {
        let Some((r, c)) = parse_addr(addr) else { continue };
        let text = csv_cell_text(cell);
        if text.is_empty() { continue; }
        max_c = max_c.max(c);
        rows.entry(r).or_default().insert(c, text);
    }
    let Some(&max_r) = rows.keys().next_back() else { return String::new() };
    let mut out = String::new();
    let empty = BTreeMap::new();
    for r in 0..=max_r {
        let row = rows.get(&r).unwrap_or(&empty);
        for c in 0..=max_c {
            if c > 0 { out.push(delim); }
            if let Some(t) = row.get(&c) { out.push_str(&escape_field(t, delim)); }
        }
        out.push_str("\r\n");
    }
    if !out.is_ascii() { out.insert(0, '\u{feff}'); }
    out
}

/// The text a cell contributes to CSV: its value (never the formula text).
fn csv_cell_text(cell: &Cell) -> String {
    match (&cell.v, cell.t.as_str()) {
        (Value::String(s), "d") => s.replacen('T', " ", 1),
        (v, _) => cell_display(v),
    }
}

fn escape_field(s: &str, delim: char) -> String {
    if s.contains('"') || s.contains(delim) || s.contains('\n') || s.contains('\r') {
        format!("\"{}\"", s.replace('"', "\"\""))
    } else {
        s.to_string()
    }
}

// =====================================================================
// Tests
// =====================================================================

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Write;

    fn cell(v: Value, t: &str) -> Cell { Cell::new(v, t) }

    fn book_of(sheets: Vec<Sheet>) -> ExcelBook {
        ExcelBook { kind: "excel".into(), active: 0, sheets, warnings: vec![] }
    }

    fn sheet_of(name: &str, cells: &[(&str, Cell)]) -> Sheet {
        Sheet { name: name.into(), cells: cells.iter().cloned().map(|(a, c)| (a.to_string(), c)).collect(), ..Default::default() }
    }

    fn round_trip(book: &ExcelBook) -> ExcelBook {
        let bytes = write_xlsx(book).expect("write");
        parse_book(&bytes, None).expect("parse")
    }

    /// Unzip one part of an xlsx as text.
    fn part(bytes: &[u8], name: &str) -> String {
        let mut z = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
        let mut s = String::new();
        z.by_name(name).unwrap().read_to_string(&mut s).unwrap();
        s
    }

    /// Build a minimal xlsx by hand (for features rust_xlsxwriter can't write:
    /// error cells, the 1904 date system, offset dimension ranges).
    fn handmade_xlsx(sheet_xml: &str, workbook_pr: &str, styles_xml: &str) -> Vec<u8> {
        let mut w = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let o = zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
        let files: Vec<(&str, String)> = vec![
            ("[Content_Types].xml", r#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>"#.into()),
            ("_rels/.rels", r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#.into()),
            ("xl/workbook.xml", format!(r#"<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">{workbook_pr}<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>"#)),
            ("xl/_rels/workbook.xml.rels", r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#.into()),
            ("xl/styles.xml", styles_xml.to_string()),
            ("xl/worksheets/sheet1.xml", format!(r#"<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">{sheet_xml}</worksheet>"#)),
        ];
        for (name, body) in files {
            w.start_file(name, o).unwrap();
            w.write_all(body.as_bytes()).unwrap();
        }
        w.finish().unwrap().into_inner()
    }

    const DATE_STYLES: &str = r#"<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs></styleSheet>"#;

    // ---------------- addresses ----------------

    #[test]
    fn address_parsing_keeps_row_digit_order() {
        assert_eq!(parse_addr("A1"), Some((0, 0)));
        assert_eq!(parse_addr("A10"), Some((9, 0)));
        assert_eq!(parse_addr("A12"), Some((11, 0)));
        assert_eq!(parse_addr("B100"), Some((99, 1)));
        assert_eq!(parse_addr("$C$5"), Some((4, 2)));
        assert_eq!(parse_addr("XFD1048576"), Some((1_048_575, 16_383)));
        assert_eq!(parse_addr("A0"), None);
        assert_eq!(parse_addr("XFE1"), None);
        assert_eq!(parse_addr("A1048577"), None);
        assert_eq!(parse_addr("1A"), None);
        assert_eq!(parse_addr("A1B"), None);
        assert_eq!(index_to_col(0), "A");
        assert_eq!(index_to_col(25), "Z");
        assert_eq!(index_to_col(26), "AA");
        assert_eq!(index_to_col(16_383), "XFD");
        assert_eq!(parse_addr("A12").map(|(r, _)| r), Some(11));
    }

    #[test]
    fn writer_puts_cells_on_the_right_rows() {
        let book = book_of(vec![sheet_of("S", &[
            ("A1", cell(json!("row1"), "s")),
            ("A10", cell(json!("row10"), "s")),
            ("A12", cell(json!("row12"), "s")),
            ("A23", cell(json!("row23"), "s")),
            ("B100", cell(json!(100), "n")),
            ("XFD1048576", cell(json!("corner"), "s")),
        ])]);
        let back = round_trip(&book);
        let c = &back.sheets[0].cells;
        assert_eq!(c["A1"].v, json!("row1"));
        assert_eq!(c["A10"].v, json!("row10"));
        assert_eq!(c["A12"].v, json!("row12"));
        assert_eq!(c["A23"].v, json!("row23"));
        assert_eq!(c["B100"].v.as_f64(), Some(100.0));
        assert_eq!(c["XFD1048576"].v, json!("corner"));
        assert_eq!(c.len(), 6);
    }

    // ---------------- reading ----------------

    #[test]
    fn sheet_not_starting_at_a1_keeps_every_cell() {
        let mut cells = vec![];
        for r in 5..=8 {
            for col in ["C", "D", "E", "F"] {
                cells.push((format!("{col}{r}"), cell(json!(format!("{col}{r}")), "s")));
            }
        }
        let sheet = Sheet { name: "Offset".into(), cells: cells.into_iter().collect(), ..Default::default() };
        let back = round_trip(&book_of(vec![sheet]));
        let c = &back.sheets[0].cells;
        assert_eq!(c.len(), 16, "{c:?}");
        assert_eq!(c["C5"].v, json!("C5"));
        assert_eq!(c["F8"].v, json!("F8"));
    }

    #[test]
    fn sparse_offset_sheet_with_formula_at_its_own_offset() {
        let mut f = cell(json!(5), "n");
        f.f = Some("=C5+2".into());
        let book = book_of(vec![sheet_of("Data", &[
            ("E1", cell(json!("far"), "s")),
            ("C5", cell(json!(3), "n")),
            ("D5", cell(json!("hdr"), "s")),
            ("D7", f),
        ])]);
        let back = round_trip(&book);
        let c = &back.sheets[0].cells;
        assert_eq!(c["E1"].v, json!("far"));
        assert_eq!(c["C5"].v.as_f64(), Some(3.0));
        assert_eq!(c["D5"].v, json!("hdr"));
        assert_eq!(c["D7"].f.as_deref(), Some("=C5+2"));
        assert_eq!(c["D7"].v.as_f64(), Some(5.0));
    }

    #[test]
    fn formulas_round_trip_with_cached_values() {
        let mut sum = cell(json!(30), "n");
        sum.f = Some("=SUM(A1:A2)".into());
        let mut text = cell(json!("hi"), "s");
        text.f = Some("=\"h\"&\"i\"".into());
        let mut err = cell(json!("#DIV/0!"), "e");
        err.f = Some("=1/0".into());
        let mut no_cache = cell(Value::Null, "f");
        no_cache.f = Some("=A1*2".into());
        let mut no_eq = cell(json!(11), "n");
        no_eq.f = Some("A1+1".into());
        let mut future = cell(json!(10), "n");
        future.f = Some("=XLOOKUP(10,A1:A2,A1:A2)".into());
        let book = book_of(vec![sheet_of("S", &[
            ("A1", cell(json!(10), "n")), ("A2", cell(json!(20), "n")),
            ("B1", sum), ("B2", text), ("B3", err), ("B4", no_cache), ("B5", no_eq), ("B6", future),
        ])]);
        let bytes = write_xlsx(&book).unwrap();
        let xml = part(&bytes, "xl/worksheets/sheet1.xml");
        assert!(xml.contains("<f>SUM(A1:A2)</f><v>30</v>"), "{xml}");
        assert!(xml.contains("_xlfn.XLOOKUP"), "future function must be prefixed on disk");
        let back = parse_book(&bytes, None).unwrap();
        let c = &back.sheets[0].cells;
        assert_eq!(c["B1"].f.as_deref(), Some("=SUM(A1:A2)"));
        assert_eq!(c["B1"].v.as_f64(), Some(30.0));
        assert_eq!(c["B1"].t, "n");
        assert_eq!(c["B2"].v, json!("hi"));
        assert_eq!(c["B3"].t, "e");
        assert_eq!(c["B3"].v, json!("#DIV/0!"));
        assert_eq!(c["B4"].f.as_deref(), Some("=A1*2"));
        assert_eq!(c["B5"].f.as_deref(), Some("=A1+1"));
        assert_eq!(c["B6"].f.as_deref(), Some("=XLOOKUP(10,A1:A2,A1:A2)"), "prefix stripped on read");
    }

    #[test]
    fn clean_formula_strips_storage_prefixes_outside_strings() {
        assert_eq!(clean_formula("_xlfn._xlws.FILTER(A1:A3,B1:B3)"), "FILTER(A1:A3,B1:B3)");
        assert_eq!(clean_formula("_xlfn.CONCAT(\"_xlfn.\",A1)"), "CONCAT(\"_xlfn.\",A1)");
    }

    #[test]
    fn hidden_empty_and_active_sheets_survive() {
        let mut secret = sheet_of("Secret", &[("A1", cell(json!("s3cr3t"), "s"))]);
        secret.hidden = true;
        let mut book = book_of(vec![
            sheet_of("Data", &[("A1", cell(json!(1), "n"))]),
            sheet_of("Empty", &[]),
            secret,
        ]);
        book.active = 1;
        let bytes = write_xlsx(&book).unwrap();
        let back = parse_book(&bytes, None).unwrap();
        let names: Vec<&str> = back.sheets.iter().map(|s| s.name.as_str()).collect();
        assert_eq!(names, ["Data", "Empty", "Secret"]);
        assert!(!back.sheets[0].hidden && !back.sheets[1].hidden);
        assert!(back.sheets[2].hidden);
        assert_eq!(back.sheets[2].cells["A1"].v, json!("s3cr3t"));
        assert!(back.sheets[1].cells.is_empty());
        assert_eq!(back.active, 1);
    }

    #[test]
    fn blank_workbook_opens() {
        let bytes = write_xlsx(&book_of(vec![sheet_of("Sheet1", &[])])).unwrap();
        let back = parse_book(&bytes, None).expect("blank workbook must parse");
        assert_eq!(back.sheets.len(), 1);
        assert!(back.sheets[0].cells.is_empty());
    }

    #[test]
    fn hidden_active_sheet_falls_back_to_visible_one() {
        let mut h = sheet_of("H", &[]);
        h.hidden = true;
        let mut book = book_of(vec![h, sheet_of("V", &[])]);
        book.active = 0;
        let (bytes, _) = write_xlsx_with_warnings(&book).unwrap();
        // Excel rejects a hidden sheet that is selected; check the XML.
        let wbx = part(&bytes, "xl/workbook.xml");
        assert!(wbx.contains("activeTab=\"1\""), "{wbx}");
        assert!(!part(&bytes, "xl/worksheets/sheet1.xml").contains("tabSelected"));
        assert!(part(&bytes, "xl/worksheets/sheet2.xml").contains("tabSelected=\"1\""));
        let back = parse_book(&bytes, None).unwrap();
        assert!(back.sheets[0].hidden);
        assert_eq!(back.active, 1);
    }

    #[test]
    fn all_hidden_sheets_unhides_one_with_warning() {
        let mut a = sheet_of("A", &[]);
        a.hidden = true;
        let mut b = sheet_of("B", &[]);
        b.hidden = true;
        let (bytes, warns) = write_xlsx_with_warnings(&book_of(vec![a, b])).unwrap();
        assert!(warns.iter().any(|w| w.contains("visible")), "{warns:?}");
        let back = parse_book(&bytes, None).unwrap();
        assert_eq!(back.sheets.len(), 2);
        assert!(!back.sheets[0].hidden);
        assert!(back.sheets[1].hidden);
    }

    // ---------------- dates ----------------

    #[test]
    fn date_serials_convert_correctly() {
        let d = |v: f64| datetime_cell(&ExcelDateTime::new(v, ExcelDateTimeType::DateTime, false));
        assert_eq!(d(45306.0).v, json!("2024-01-15"));
        assert_eq!(d(45306.0).t, "d");
        assert_eq!(d(45306.4375).v, json!("2024-01-15T10:30:00"));
        assert_eq!(d(1.0).v, json!("1900-01-01"));
        assert_eq!(d(59.0).v, json!("1900-02-28"));
        assert_eq!(d(61.0).v, json!("1900-03-01"));
        assert_eq!(d(60.0).t, "n", "phantom 1900-02-29 stays a number");
        assert_eq!(d(45306.999999999).v, json!("2024-01-16"), "rounds up to midnight, never 24:00:00");
        // Time-only values are numbers with a time format.
        let t = d(0.5);
        assert_eq!((t.t.as_str(), t.v.as_f64(), t.z.as_deref()), ("n", Some(0.5), Some("hh:mm:ss")));
        // Huge / negative serials: no panic, stay numbers.
        assert_eq!(d(1e300).t, "n");
        assert_eq!(d(-5.0).t, "n");
        assert_eq!(d(f64::MAX).t, "n");
        // 1904 date system.
        let d4 = |v: f64| datetime_cell(&ExcelDateTime::new(v, ExcelDateTimeType::DateTime, true));
        assert_eq!(d4(0.0).v, json!("1904-01-01"));
        assert_eq!(d4(43844.0).v, json!("2024-01-15"));
        // Durations.
        let dur = datetime_cell(&ExcelDateTime::new(1.5, ExcelDateTimeType::TimeDelta, false));
        assert_eq!((dur.t.as_str(), dur.z.as_deref()), ("n", Some("[h]:mm:ss")));
    }

    #[test]
    fn iso_strings_convert_to_serials() {
        assert_eq!(iso_to_serial("2024-01-15"), Some((45306.0, IsoKind::Date)));
        assert_eq!(iso_to_serial("2024-01-15T10:30:00"), Some((45306.4375, IsoKind::DateTime)));
        assert_eq!(iso_to_serial("2024-01-15 10:30"), Some((45306.4375, IsoKind::DateTime)));
        // Timezone designators are ignored, never applied.
        assert_eq!(iso_to_serial("2024-01-15T10:30:00Z"), Some((45306.4375, IsoKind::DateTime)));
        assert_eq!(iso_to_serial("2024-01-15T10:30:00+05:00"), Some((45306.4375, IsoKind::DateTime)));
        assert_eq!(iso_to_serial("1900-01-01"), Some((1.0, IsoKind::Date)));
        assert_eq!(iso_to_serial("1900-03-01"), Some((61.0, IsoKind::Date)));
        assert_eq!(iso_to_serial("12:00"), Some((0.5, IsoKind::Time)));
        assert_eq!(iso_to_serial("PT36H"), Some((1.5, IsoKind::Duration)));
        assert_eq!(iso_to_serial("2024-02-30"), None);
        assert_eq!(iso_to_serial("2024-01-+1"), None);
        assert_eq!(iso_to_serial("1899-12-31"), None);
        assert_eq!(iso_to_serial("not a date"), None);
    }

    #[test]
    fn dates_are_written_as_real_excel_dates() {
        let mut custom = cell(json!("2024-03-05"), "d");
        custom.z = Some("dd/mm/yyyy".into());
        let book = book_of(vec![sheet_of("S", &[
            ("A1", cell(json!("2024-01-15"), "d")),
            ("A2", cell(json!("2024-01-15T10:30:00"), "d")),
            ("A3", custom),
            ("A4", cell(json!("not a date"), "d")),
            ("A5", cell(json!("12:30:00"), "d")),
        ])]);
        let bytes = write_xlsx(&book).unwrap();
        let xml = part(&bytes, "xl/worksheets/sheet1.xml");
        assert!(xml.contains("<v>45306</v>"), "{xml}");
        let back = parse_book(&bytes, None).unwrap();
        let c = &back.sheets[0].cells;
        assert_eq!((c["A1"].t.as_str(), &c["A1"].v), ("d", &json!("2024-01-15")));
        assert_eq!(c["A2"].v, json!("2024-01-15T10:30:00"));
        assert_eq!(c["A3"].v, json!("2024-03-05"));
        assert_eq!(c["A3"].z.as_deref(), Some("dd/mm/yyyy"));
        assert_eq!((c["A4"].t.as_str(), &c["A4"].v), ("s", &json!("not a date")));
        assert_eq!(c["A5"].v.as_f64(), Some(0.5208333333333334));
        assert_eq!(c["A5"].z.as_deref(), Some("hh:mm:ss"));
    }

    #[test]
    fn handmade_1904_workbook_and_error_cells() {
        let sheet = r#"<dimension ref="B2:C4"/><sheetData><row r="2"><c r="B2" s="1"><v>0</v></c><c r="C2" t="e"><v>#DIV/0!</v></c></row><row r="3"><c r="B3" s="1"><v>43844</v></c><c r="C3" t="e"><v>#N/A</v></c></row><row r="4"><c r="C4" s="1"><v>1e300</v></c></row></sheetData>"#;
        let bytes = handmade_xlsx(sheet, r#"<workbookPr date1904="1"/>"#, DATE_STYLES);
        let book = parse_book(&bytes, None).unwrap();
        let c = &book.sheets[0].cells;
        assert_eq!(c["B2"].v, json!("1904-01-01"));
        assert_eq!(c["B3"].v, json!("2024-01-15"));
        assert_eq!((c["C2"].t.as_str(), &c["C2"].v), ("e", &json!("#DIV/0!")));
        assert_eq!(c["C3"].v, json!("#N/A"));
        assert_eq!(c["C4"].t, "n", "huge date serial must not panic");
    }

    // ---------------- styles, merges, widths ----------------

    #[test]
    fn styles_and_number_formats_round_trip() {
        let mut head = cell(json!("Header"), "s");
        head.s = Some(CellStyle { b: true, i: true, u: true, color: Some("#FF0000".into()), bg: Some("#FFFF00".into()), align: Some("center".into()), wrap: true });
        let mut pct = cell(json!(0.25), "n");
        pct.z = Some("0.00%".into());
        let mut money = cell(json!(1234.5), "n");
        money.z = Some("\"$\"#,##0.00".into());
        let mut blank = cell(Value::Null, "s");
        blank.s = Some(CellStyle { bg: Some("#00FF00".into()), ..Default::default() });
        let mut blank_fmt = cell(Value::Null, "s");
        blank_fmt.z = Some("#,##0.00".into());
        let book = book_of(vec![sheet_of("S", &[("A1", head), ("A2", pct), ("A3", money), ("A4", blank), ("A5", blank_fmt)])]);
        let back = round_trip(&book);
        let c = &back.sheets[0].cells;
        let s = c["A1"].s.as_ref().expect("style read back");
        assert!(s.b && s.i && s.u && s.wrap);
        assert_eq!(s.color.as_deref(), Some("#FF0000"));
        assert_eq!(s.bg.as_deref(), Some("#FFFF00"));
        assert_eq!(s.align.as_deref(), Some("center"));
        assert_eq!(c["A2"].z.as_deref(), Some("0.00%"));
        assert_eq!(c["A3"].z.as_deref(), Some("\"$\"#,##0.00"));
        assert_eq!(c["A4"].s.as_ref().and_then(|s| s.bg.as_deref()), Some("#00FF00"), "styled blank cell kept");
        assert_eq!(c["A5"].z.as_deref(), Some("#,##0.00"), "number format of a blank cell kept");
        assert!(c["A4"].v.is_null());
    }

    #[test]
    fn merges_and_column_widths_round_trip() {
        let mut title = cell(json!(42), "n");
        title.s = Some(CellStyle { b: true, ..Default::default() });
        let mut s = sheet_of("S", &[("B2", title), ("B3", cell(json!("x"), "s"))]);
        s.merges = vec!["B2:D2".into(), "C2:C2".into(), "C2:E3".into(), "bogus".into()];
        s.col_widths = Some([("A".to_string(), 20.0), ("C".to_string(), 8.43), ("D".to_string(), 0.0)].into_iter().collect());
        let (bytes, warns) = write_xlsx_with_warnings(&book_of(vec![s])).unwrap();
        assert!(warns.iter().any(|w| w.contains("overlaps")), "{warns:?}");
        assert!(warns.iter().any(|w| w.contains("bogus")), "{warns:?}");
        let back = parse_book(&bytes, None).unwrap();
        let sh = &back.sheets[0];
        assert_eq!(sh.merges, vec!["B2:D2".to_string()]);
        assert_eq!(sh.cells["B2"].v.as_f64(), Some(42.0), "merge must not clobber the top-left value");
        assert!(sh.cells["B2"].s.as_ref().unwrap().b);
        let w = sh.col_widths.as_ref().expect("widths");
        assert_eq!(w["A"], 20.0);
        assert_eq!(w["C"], 8.43);
        assert_eq!(w["D"], 0.0, "hidden column");
    }

    #[test]
    fn unsupported_features_produce_warnings() {
        let mut wb = Workbook::new();
        let ws = wb.add_worksheet();
        ws.write_string_with_format(0, 0, "boxed", &Format::new().set_border(rust_xlsxwriter::FormatBorder::Thin)).unwrap();
        ws.write_string_with_format(1, 0, "big", &Format::new().set_font_size(20)).unwrap();
        ws.set_freeze_panes(1, 0).unwrap();
        ws.insert_note(0, 0, &rust_xlsxwriter::Note::new("hi")).unwrap();
        ws.set_row_height(3, 40).unwrap();
        let bytes = wb.save_to_buffer().unwrap();
        let book = parse_book(&bytes, None).unwrap();
        let all = book.warnings.join(" | ");
        assert!(!all.contains("frozen panes"), "frozen panes are supported now: {all}");
        assert_eq!(book.sheets[0].freeze, Some(Freeze { rows: 1, cols: 0 }));
        assert!(all.contains("comments"), "{all}");
        assert!(all.contains("custom row heights"), "{all}");
        assert!(all.contains("borders"), "{all}");
        assert!(all.contains("font faces and sizes"), "{all}");
        // A plain workbook has no warnings at all.
        let plain = write_xlsx(&book_of(vec![sheet_of("S", &[("A1", cell(json!(1), "n"))])])).unwrap();
        assert!(parse_book(&plain, None).unwrap().warnings.is_empty());
    }

    #[test]
    fn stored_width_conversion() {
        assert_eq!(stored_width_to_chars(9.140625), 8.43);
        assert_eq!(stored_width_to_chars(20.7109375), 20.0);
        assert_eq!(stored_width_to_chars(0.0), 0.0);
    }

    // ---------------- writer robustness ----------------

    #[test]
    fn sheet_names_are_sanitized_and_deduplicated() {
        let book = book_of(vec![
            sheet_of("Quarterly revenue report for 2026 final version", &[]),
            sheet_of("report [final]", &[]),
            sheet_of("a/b:c*d?e\\f", &[]),
            sheet_of("'quoted'", &[]),
            sheet_of("REPORT _FINAL_", &[]),
            sheet_of("", &[]),
            sheet_of("Report _final_", &[]),
        ]);
        let (bytes, warns) = write_xlsx_with_warnings(&book).unwrap();
        let back = parse_book(&bytes, None).unwrap();
        let names: Vec<&str> = back.sheets.iter().map(|s| s.name.as_str()).collect();
        assert_eq!(names, [
            "Quarterly revenue report for 20",
            "report _final_",
            "a_b_c_d_e_f",
            "quoted",
            "REPORT _FINAL_ (2)",
            "Sheet6",
            "Report _final_ (3)",
        ]);
        assert!(names.iter().all(|n| n.chars().count() <= 31));
        assert!(warns.len() >= 6, "{warns:?}");
    }

    #[test]
    fn long_strings_truncated_with_warning_and_bad_addresses_reported() {
        let long = "x".repeat(40_000);
        let book = book_of(vec![sheet_of("S", &[
            ("A1", cell(json!("ok"), "s")),
            ("A4", cell(json!(long), "s")),
            ("not-an-address", cell(json!(1), "n")),
            ("A5", cell(json!("after"), "s")),
        ])]);
        let (bytes, warns) = write_xlsx_with_warnings(&book).unwrap();
        assert!(warns.iter().any(|w| w.contains("S!A4")), "{warns:?}");
        assert!(warns.iter().any(|w| w.contains("not-an-address")), "{warns:?}");
        let back = parse_book(&bytes, None).unwrap();
        let c = &back.sheets[0].cells;
        assert_eq!(c["A4"].v.as_str().unwrap().len(), 32_767);
        assert_eq!(c["A5"].v, json!("after"));
    }

    #[test]
    fn loose_value_types_are_written_sensibly() {
        let book = book_of(vec![sheet_of("S", &[
            ("A1", cell(json!("12.5"), "n")),
            ("A2", cell(json!("abc"), "n")),
            ("A3", cell(json!("TRUE"), "b")),
            ("A4", cell(json!(7), "s")),
            ("A5", cell(json!("#N/A"), "e")),
        ])]);
        let back = round_trip(&book);
        let c = &back.sheets[0].cells;
        assert_eq!(c["A1"].v.as_f64(), Some(12.5));
        assert_eq!(c["A2"].v, json!("abc"));
        assert_eq!(c["A3"].v, json!(true));
        assert_eq!(c["A4"].v.as_f64(), Some(7.0));
        assert_eq!(c["A5"].v, json!("#N/A"));
    }

    #[test]
    fn cell_budget_truncation_warns() {
        let cells: Vec<(String, Cell)> = (0..50).map(|i| (format!("A{}", i + 1), cell(json!(i), "n"))).collect();
        let s = Sheet { name: "S".into(), cells: cells.into_iter().collect(), ..Default::default() };
        let bytes = write_xlsx(&book_of(vec![s])).unwrap();
        let back = parse_book_limited(&bytes, None, 10).unwrap();
        assert_eq!(back.sheets[0].cells.len(), 10);
        assert!(back.warnings.iter().any(|w| w.contains("too large")), "{:?}", back.warnings);
        let full = parse_book(&bytes, None).unwrap();
        assert_eq!(full.sheets[0].cells.len(), 50);
        assert!(full.warnings.is_empty(), "{:?}", full.warnings);
    }

    #[test]
    fn sheet_hint_selects_one_sheet() {
        let bytes = write_xlsx(&book_of(vec![sheet_of("A", &[]), sheet_of("B", &[("A1", cell(json!(1), "n"))])])).unwrap();
        let b = parse_book(&bytes, Some("B")).unwrap();
        assert_eq!(b.sheets.len(), 1);
        assert_eq!(b.sheets[0].name, "B");
        assert!(parse_book(&bytes, Some("Nope")).is_err());
    }

    // ---------------- CSV ----------------

    fn csv_cells(text: &str, delim: char) -> BTreeMap<String, Cell> {
        parse_delimited(text, delim, "T").sheets.remove(0).cells
    }

    #[test]
    fn csv_quoted_fields_with_newlines_quotes_and_delimiters() {
        let c = csv_cells("1,\"line one\r\nline two\",02134\r\n\"a,b\",\"say \"\"hi\"\"\",x\n", ',');
        assert_eq!(c["B1"].v, json!("line one\r\nline two"));
        assert_eq!(c["C1"].v, json!("02134"));
        assert_eq!(c["A2"].v, json!("a,b"));
        assert_eq!(c["B2"].v, json!("say \"hi\""));
        assert_eq!(c["C2"].v, json!("x"));
        assert!(!c.contains_key("A3"));
    }

    #[test]
    fn csv_mid_field_quote_is_literal() {
        let c = csv_cells("ruler,12\" long,3\r\npipe,2\" wide,5", ',');
        assert_eq!(c["B1"].v, json!("12\" long"));
        assert_eq!(c["C1"].v.as_f64(), Some(3.0));
        assert_eq!(c["A2"].v, json!("pipe"));
        assert_eq!(c["B2"].v, json!("2\" wide"));
        let c = csv_cells("=HYPERLINK(\"x\"),b", ',');
        assert_eq!(c["A1"].v, json!("=HYPERLINK(\"x\")"));
        assert_eq!(c["A1"].t, "s", "CSV formulas are kept as text");
    }

    #[test]
    fn csv_values_stay_text_unless_plain_numbers() {
        let c = csv_cells("007,3%,1234567890123456789,1.50,-2,0.25,1e5,+3,.5,TRUE,2024-01-15,1 234,\" 5\",0", ',');
        let t = |a: &str| (c[a].t.clone(), c[a].v.clone());
        assert_eq!(t("A1"), ("s".into(), json!("007")));
        assert_eq!(t("B1"), ("s".into(), json!("3%")));
        assert_eq!(t("C1"), ("s".into(), json!("1234567890123456789")));
        assert_eq!(t("D1"), ("n".into(), json!(1.5)));
        assert_eq!(t("E1"), ("n".into(), json!(-2)));
        assert_eq!(t("F1"), ("n".into(), json!(0.25)));
        assert_eq!(t("G1").0, "s");
        assert_eq!(t("H1").0, "s");
        assert_eq!(t("I1").0, "s");
        assert_eq!(t("J1"), ("s".into(), json!("TRUE")));
        assert_eq!(t("K1"), ("s".into(), json!("2024-01-15")));
        assert_eq!(t("L1").0, "s");
        assert_eq!(t("M1"), ("s".into(), json!(" 5")));
        assert_eq!(t("N1"), ("n".into(), json!(0)));
    }

    #[test]
    fn csv_blank_lines_keep_row_positions_and_unterminated_quote_warns() {
        let c = csv_cells("a\n\nc\n", ',');
        assert_eq!(c["A1"].v, json!("a"));
        assert_eq!(c["A3"].v, json!("c"));
        let b = parse_delimited("a,\"never closed\nb", ',', "T");
        assert_eq!(b.sheets[0].cells["B1"].v, json!("never closed\nb"));
        assert!(b.warnings.iter().any(|w| w.contains("quote")));
    }

    #[test]
    fn csv_encodings() {
        let (t, w) = decode_text(b"caf\xe9,\x80");
        assert_eq!(t, "café,€");
        assert!(w.is_some());
        let (t, w) = decode_text("\u{feff}café".as_bytes());
        assert_eq!(t, "café");
        assert!(w.is_none());
        let mut le = vec![0xFF, 0xFE];
        for u in "a,é\r\n".encode_utf16() { le.extend_from_slice(&u.to_le_bytes()); }
        assert_eq!(decode_text(&le).0, "a,é\r\n");
        let mut be = vec![0xFE, 0xFF];
        for u in "x;y".encode_utf16() { be.extend_from_slice(&u.to_be_bytes()); }
        assert_eq!(decode_text(&be).0, "x;y");
        let mut nobom = vec![];
        for u in "name,value\r\nfoo,1\r\n".encode_utf16() { nobom.extend_from_slice(&u.to_le_bytes()); }
        assert_eq!(decode_text(&nobom).0, "name,value\r\nfoo,1\r\n");
        let book = parse_csv_bytes(b"caf\xe9;1\r\n", None, "T");
        assert_eq!(book.sheets[0].cells["A1"].v, json!("café"));
        assert_eq!(book.sheets[0].cells["B1"].v, json!(1));
        assert!(!book.warnings.is_empty());
    }

    #[test]
    fn csv_delimiter_sniffing() {
        assert_eq!(sniff_delimiter("a;b;c\n1,5;2,5;3\n4;5;6\n"), ';');
        assert_eq!(sniff_delimiter("a\tb\n1\t2\n"), '\t');
        assert_eq!(sniff_delimiter("a,b,c\n1,2,3\n"), ',');
        assert_eq!(sniff_delimiter("a|b\n1|2\n"), '|');
        assert_eq!(sniff_delimiter("\"x;y\",z\n\"p;q\",r\n"), ',', "quoted delimiters don't count");
        assert_eq!(sniff_delimiter("single column\nvalues\n"), ',');
        let book = parse_csv_bytes(b"sep=;\r\na;b\r\n1;2\r\n", None, "T");
        let c = &book.sheets[0].cells;
        assert_eq!(c["A1"].v, json!("a"));
        assert_eq!(c["B2"].v, json!(2));
        let book = parse_csv_bytes(b"a\tb\n", Some('\t'), "Tab");
        assert_eq!(book.sheets[0].name, "Tab");
        assert_eq!(book.sheets[0].cells["B1"].v, json!("b"));
    }

    #[test]
    fn csv_export_values_from_a1_and_round_trip() {
        let mut f = cell(json!(30), "n");
        f.f = Some("=SUM(B2:B3)".into());
        let mut f_nocache = cell(Value::Null, "f");
        f_nocache.f = Some("=1/0".into());
        let book = book_of(vec![sheet_of("S", &[
            ("B2", cell(json!(10), "n")),
            ("B3", cell(json!(20.5), "n")),
            ("C2", cell(json!(true), "b")),
            ("C3", cell(json!("a,\"b\"\nc"), "s")),
            ("D2", f),
            ("D3", f_nocache),
            ("E2", cell(json!("2024-01-15T10:30:00"), "d")),
            ("E3", cell(json!("tab\there"), "s")),
        ])]);
        let out = sheet_to_delimited(&book, 0, ',');
        assert_eq!(out, ",,,,\r\n,10,TRUE,30,2024-01-15 10:30:00\r\n,20.5,\"a,\"\"b\"\"\nc\",,tab\there\r\n");
        let tsv = sheet_to_delimited(&book, 0, '\t');
        assert!(tsv.contains("\"tab\there\""), "{tsv}");
        let back = parse_delimited(&tsv, '\t', "S");
        assert_eq!(back.sheets[0].cells["E3"].v, json!("tab\there"));
        assert_eq!(back.sheets[0].cells["C3"].v, json!("a,\"b\"\nc"));
        assert_eq!(back.sheets[0].cells["B2"].v, json!(10));
        // Text that looks like numbers survives a CSV round trip unchanged.
        let src = "zip,pct,id\r\n00501,3%,12345678901234567890\r\n";
        let b = parse_csv_bytes(src.as_bytes(), None, "T");
        assert_eq!(sheet_to_delimited(&b, 0, ','), src);
        // Non-ASCII output carries a BOM for Excel.
        let b = parse_delimited("café", ',', "T");
        assert!(sheet_to_delimited(&b, 0, ',').starts_with('\u{feff}'));
        assert_eq!(sheet_to_delimited(&book_of(vec![sheet_of("E", &[])]), 0, ','), "");
        assert_eq!(sheet_to_delimited(&book, 9, ','), "");
    }

    #[test]
    fn csv_limits_warn() {
        let b = parse_delimited_limited("1,2,3\n4,5,6\n", ',', "T", 4);
        assert_eq!(b.sheets[0].cells.len(), 4);
        assert!(b.warnings.iter().any(|w| w.contains("too large")));
    }

    #[test]
    fn parse_scalar_mirrors_user_input_rules() {
        assert_eq!(parse_scalar("42").t, "n");
        assert_eq!(parse_scalar("1,234").v.as_f64(), Some(1234.0));
        let p = parse_scalar("3%");
        assert_eq!((p.v.as_f64(), p.z.as_deref()), (Some(0.03), Some("0%")));
        assert_eq!(parse_scalar("2.5%").z.as_deref(), Some("0.0%"));
        assert_eq!(parse_scalar("=SUM(A1:A2)").f.as_deref(), Some("=SUM(A1:A2)"));
        assert_eq!(parse_scalar("TRUE").t, "b");
        assert_eq!(parse_scalar("hello").t, "s");
        assert_eq!(parse_scalar("2024-01-02").t, "d");
        assert_eq!(parse_scalar("2024-01-02 10:00").v, json!("2024-01-02T10:00:00"));
        assert_eq!(parse_scalar("2024-13-02").t, "s");
        assert_eq!(parse_scalar("2024-01-02abc").t, "s");
    }

    // ---------------- frozen panes, auto filter, charts ----------------

    /// Header row + 5 data rows in A1:D6.
    fn chart_data_cells() -> Vec<(String, Cell)> {
        let mut cells = Vec::new();
        for (c, h) in ["Month", "Sales", "Costs", "Profit"].iter().enumerate() {
            cells.push((cell_addr(0, c as u32), cell(json!(h), "s")));
        }
        for r in 1..=5u32 {
            cells.push((cell_addr(r, 0), cell(json!(format!("M{r}")), "s")));
            for c in 1..=3u32 {
                cells.push((cell_addr(r, c), cell(json!(r * 10 + c), "n")));
            }
        }
        cells
    }

    fn chart_json(ty: &str, title: Option<&str>, range: &str, x: i64, y: i64, w: i64, h: i64) -> Value {
        let mut v = json!({ "type": ty, "range": range, "x": x, "y": y, "w": w, "h": h });
        if let Some(t) = title { v["title"] = json!(t); }
        v
    }

    fn assert_chart_close(got: &Value, want: &Value) {
        for k in ["type", "range", "title"] {
            assert_eq!(got.get(k), want.get(k), "{k}: got {got} want {want}");
        }
        for k in ["x", "y", "w", "h"] {
            let (g, w) = (got[k].as_i64().expect("integer px"), want[k].as_i64().unwrap());
            assert!((g - w).abs() <= 2, "{k}: got {got} want {want}");
        }
    }

    /// The sheet used for the write -> read round trip (also dumped for the
    /// openpyxl check, see `dump_written_workbook_for_external_check`).
    fn feature_book() -> ExcelBook {
        let mut s = Sheet {
            name: "My Sheet".into(),
            cells: chart_data_cells().into_iter().collect(),
            ..Default::default()
        };
        s.col_widths = Some([("A".to_string(), 20.0), ("C".to_string(), 5.0)].into_iter().collect());
        s.freeze = Some(Freeze { rows: 1, cols: 1 });
        s.auto_filter = Some("A1:D6".into());
        s.charts = vec![
            chart_json("column", Some("Sales vs costs"), "A1:C6", 500, 50, 400, 250),
            chart_json("bar", Some("Horizontal"), "A1:D6", 0, 200, 480, 288),
            chart_json("line", Some("Trend"), "A1:B6", 300, 520, 360, 240),
            chart_json("pie", None, "A1:B6", 700, 530, 300, 300),
            chart_json("scatter", Some("Costs by sales"), "B1:C6", 1234, 777, 451, 263),
            chart_json("area", Some("Area"), "A1:C6", 90, 900, 520, 310),
        ];
        let mut other = sheet_of("Other", &[("A1", cell(json!(1), "n"))]);
        other.freeze = Some(Freeze { rows: 0, cols: 2 });
        book_of(vec![s, other])
    }

    #[test]
    fn freeze_filter_and_charts_round_trip() {
        let book = feature_book();
        let (bytes, warns) = write_xlsx_with_warnings(&book).unwrap();
        assert!(warns.is_empty(), "{warns:?}");
        // Real chart parts referencing the (quoted) sheet name.
        let c1 = part(&bytes, "xl/charts/chart1.xml");
        assert!(c1.contains("'My Sheet'!$B$2:$B$6"), "{c1}");
        let back = parse_book(&bytes, None).unwrap();
        assert!(back.warnings.is_empty(), "{:?}", back.warnings);
        let s = &back.sheets[0];
        assert_eq!(s.freeze, Some(Freeze { rows: 1, cols: 1 }));
        assert_eq!(s.auto_filter.as_deref(), Some("A1:D6"));
        assert_eq!(s.charts.len(), 6, "{:?}", s.charts);
        // Charts come back in file order (rust_xlsxwriter sorts by position).
        for want in &book.sheets[0].charts {
            let got = s.charts.iter().find(|c| c["type"] == want["type"]).expect("chart type read back");
            assert_chart_close(got, want);
        }
        assert!(s.charts.iter().find(|c| c["type"] == "pie").unwrap().get("title").is_none());
        let o = &back.sheets[1];
        assert_eq!(o.freeze, Some(Freeze { rows: 0, cols: 2 }));
        assert!(o.auto_filter.is_none() && o.charts.is_empty());
        // JSON shape of the model.
        let js = serde_json::to_value(s).unwrap();
        assert_eq!(js["autoFilter"], json!("A1:D6"));
        assert_eq!(js["freeze"], json!({ "rows": 1, "cols": 1 }));
    }

    #[test]
    fn chart_geometry_follows_canonical_spec() {
        // Default widths: grid 88 px / Excel 64 px columns, 24 / 20 px rows.
        let m = ColMetrics::new(None);
        assert_eq!(grid_to_excel_anchor(&m, 100.0, 30.0), (1, 1, 9, 5));
        assert_eq!(grid_to_excel_anchor(&m, 0.0, 0.0), (0, 0, 0, 0));
        assert_eq!(grid_to_excel_anchor(&m, 88.0, 24.0), (1, 1, 0, 0));
        let one = RawAnchor::One {
            from: AnchorCell { col: 1, col_off: 9.0 * EMU_PER_PX, row: 1, row_off: 5.0 * EMU_PER_PX },
            cx: 480.0 * EMU_PER_PX,
            cy: 288.0 * EMU_PER_PX,
        };
        assert_eq!(one.to_grid(&m), (100, 30, 480, 288));
        // Two-cell anchor: size measured in Excel px (64 px columns, 20 px rows).
        let two = RawAnchor::Two {
            from: AnchorCell { col: 2, col_off: 0.0, row: 3, row_off: 0.0 },
            to: AnchorCell { col: 9, col_off: 32.0 * EMU_PER_PX, row: 18, row_off: 10.0 * EMU_PER_PX },
        };
        assert_eq!(two.to_grid(&m), (176, 72, 7 * 64 + 32, 15 * 20 + 10));
        // Absolute anchor: Excel px position -> cell + offset -> grid.
        let abs = RawAnchor::Abs { x: 96.0 * EMU_PER_PX, y: 30.0 * EMU_PER_PX, cx: 300.0 * EMU_PER_PX, cy: 200.0 * EMU_PER_PX };
        assert_eq!(abs.to_grid(&m), (88 + 44, 24 + 12, 300, 200));
        // Custom widths are the same px in grid and Excel: max(4, round(w*7+5)).
        let widths: BTreeMap<String, f64> = [("A".to_string(), 20.0), ("B".to_string(), 0.0)].into_iter().collect();
        let m = ColMetrics::new(Some(&widths));
        assert_eq!((m.grid_px(0), m.excel_px(0), m.grid_px(1), m.excel_px(2)), (145.0, 145.0, 5.0, 64.0));
        assert_eq!(m.grid_x(3), 145.0 + 5.0 + 88.0);
        assert_eq!(m.excel_x(3), 145.0 + 5.0 + 64.0);
        assert_eq!(grid_to_excel_anchor(&m, 160.0, 0.0), (0, 2, 7, 0));
        // Huge positions are clamped, never panic.
        let (r, c, _, _) = grid_to_excel_anchor(&m, 1e12, 1e12);
        assert!(r < MAX_ROWS && (c as u32) < MAX_COLS);
    }

    #[test]
    fn invalid_chart_freeze_and_filter_input_warns_but_saves() {
        let mut s = Sheet { name: "S".into(), cells: chart_data_cells().into_iter().collect(), ..Default::default() };
        s.freeze = Some(Freeze { rows: 2_000_000, cols: 0 });
        s.auto_filter = Some("bogus".into());
        s.charts = vec![
            json!("not a chart"),
            json!({ "type": "radar", "range": "A1:B6" }),
            json!({ "type": "line" }),
            json!({ "type": "line", "range": "A1:B" }),
            json!({ "type": "column", "range": "A1:C1" }),
            json!({ "type": "scatter", "range": "B1:B6" }),
            json!({ "type": "column", "range": "A1:XFD6" }),
            // Valid; missing geometry uses defaults, tiny / bogus sizes are clamped.
            json!({ "type": "Column", "range": "A1:C6", "title": "  " }),
            json!({ "type": "line", "range": "$A$1:$B$6", "x": "12", "y": -40, "w": 3, "h": "abc" }),
        ];
        let (bytes, warns) = write_xlsx_with_warnings(&book_of(vec![s])).unwrap();
        assert_eq!(warns.iter().filter(|w| w.starts_with("Chart ")).count(), 7, "{warns:?}");
        assert!(warns.iter().any(|w| w.contains("Chart 2") && w.contains("radar")), "{warns:?}");
        assert!(warns.iter().any(|w| w.contains("Frozen panes")), "{warns:?}");
        assert!(warns.iter().any(|w| w.contains("\"bogus\"")), "{warns:?}");
        let back = parse_book(&bytes, None).unwrap();
        let sh = &back.sheets[0];
        assert!(sh.freeze.is_none() && sh.auto_filter.is_none());
        assert_eq!(sh.charts.len(), 2, "{:?}", sh.charts);
        let col = sh.charts.iter().find(|c| c["type"] == "column").unwrap();
        assert_chart_close(col, &chart_json("column", None, "A1:C6", 0, 0, 480, 288));
        let line = sh.charts.iter().find(|c| c["type"] == "line").unwrap();
        assert_chart_close(line, &chart_json("line", None, "A1:B6", 12, 0, 50, 288));
    }

    #[test]
    fn chart_titles_that_look_like_references_stay_literal() {
        let mut s = Sheet { name: "S".into(), cells: chart_data_cells().into_iter().collect(), ..Default::default() };
        s.charts = vec![chart_json("column", Some("Q1!A1"), "A1:B6", 0, 0, 480, 288)];
        let back = round_trip(&book_of(vec![s]));
        let t = back.sheets[0].charts[0]["title"].as_str().unwrap().to_string();
        assert_eq!(t.trim_end(), "Q1!A1");
    }

    #[test]
    fn reads_openpyxl_fixture_with_charts_freeze_and_filter() {
        let bytes = include_bytes!("../tests/fixtures/excel/charts_freeze_filter.xlsx");
        let book = parse_book(bytes, None).unwrap();
        let all = book.warnings.join(" | ");
        for lost in ["chart", "frozen", "filter", "drawings"] {
            assert!(!all.contains(lost), "{lost}: {all}");
        }
        let s = &book.sheets[0];
        assert_eq!(s.name, "Data");
        assert_eq!(s.freeze, Some(Freeze { rows: 1, cols: 1 }));
        assert_eq!(s.auto_filter.as_deref(), Some("A1:D6"));
        // openpyxl: one-cell anchors, 15 x 7.5 cm (567 x 283 px), default columns.
        let want = [
            chart_json("column", Some("Sales vs costs"), "A1:C6", 5 * 88, 24, 567, 283),
            chart_json("bar", Some("Horizontal"), "A1:D6", 15 * 88, 24, 567, 283),
            chart_json("line", Some("Trend"), "A1:B6", 5 * 88, 19 * 24, 567, 283),
            chart_json("pie", None, "A1:B6", 15 * 88, 19 * 24, 567, 283),
            chart_json("scatter", Some("Scatter"), "B1:C6", 5 * 88, 37 * 24, 567, 283),
            chart_json("area", Some("Area"), "A1:C6", 15 * 88, 37 * 24, 567, 283),
        ];
        assert_eq!(s.charts.len(), want.len(), "{:?}", s.charts);
        for (got, want) in s.charts.iter().zip(want.iter()) {
            assert_chart_close(got, want);
        }
        // And it survives our own save.
        let (bytes2, warns) = write_xlsx_with_warnings(&book).unwrap();
        assert!(warns.is_empty(), "{warns:?}");
        let again = parse_book(&bytes2, None).unwrap();
        assert_eq!(again.sheets[0].charts.len(), 6);
        assert_eq!(again.sheets[0].freeze, s.freeze);
        for (got, want) in again.sheets[0].charts.iter().zip(s.charts.iter()) {
            assert_chart_close(got, want);
        }
    }

    #[test]
    fn unreadable_charts_pictures_and_shapes_are_reported() {
        const PNG: &[u8] = &[
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
            0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
            0x89, 0x00, 0x00, 0x00, 0x0A, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9C, 0x63, 0x00, 0x01, 0x00, 0x00,
            0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE,
            0x42, 0x60, 0x82,
        ];
        let mut wb = Workbook::new();
        let ws = wb.add_worksheet().set_name("Main").unwrap();
        ws.write(0, 0, "h").unwrap();
        ws.write(1, 0, 1).unwrap();
        ws.write(2, 0, 2).unwrap();
        // Readable.
        let mut ok = Chart::new(ChartType::Line);
        ok.add_series().set_values("Main!$A$2:$A$3").set_name("Main!$A$1");
        ws.insert_chart(5, 5, &ok).unwrap();
        // Unsupported type.
        let mut radar = Chart::new(ChartType::Radar);
        radar.add_series().set_values("Main!$A$2:$A$3");
        ws.insert_chart(20, 5, &radar).unwrap();
        // Data on another sheet.
        let mut foreign = Chart::new(ChartType::Column);
        foreign.add_series().set_values("Data!$A$2:$A$3");
        ws.insert_chart(40, 5, &foreign).unwrap();
        let ws2 = wb.add_worksheet().set_name("Data").unwrap();
        ws2.write(1, 0, 5).unwrap();
        ws2.write(2, 0, 6).unwrap();
        let bytes = wb.save_to_buffer().unwrap();
        let book = parse_book(&bytes, None).unwrap();
        assert_eq!(book.sheets[0].charts.len(), 1, "{:?}", book.sheets[0].charts);
        assert_eq!(book.sheets[0].charts[0]["range"], json!("A1:A3"));
        let all = book.warnings.join(" | ");
        assert!(all.contains("2 chart(s) could not be read and will be lost if you save."), "{all}");
        assert_eq!(all.matches("could not be read").count(), 1, "{all}");
        assert!(!all.contains("drawings and shapes") && !all.contains("images"), "{all}");

        // Pictures and shapes are flagged; charts alone are not.
        let mut wb = Workbook::new();
        let ws = wb.add_worksheet();
        ws.write(0, 0, "x").unwrap();
        ws.insert_image(3, 3, &rust_xlsxwriter::Image::new_from_buffer(PNG).unwrap()).unwrap();
        ws.insert_shape(10, 3, &rust_xlsxwriter::Shape::textbox().set_text("note")).unwrap();
        let book = parse_book(&wb.save_to_buffer().unwrap(), None).unwrap();
        let all = book.warnings.join(" | ");
        assert!(all.contains("images") && all.contains("drawings and shapes"), "{all}");
        assert!(!all.contains("chart"), "{all}");
    }

    #[test]
    fn chart_formula_and_xml_parsing_is_robust() {
        assert_eq!(
            parse_chart_formula("'My ''Q'' Sheet'!$A$2:$A$6"),
            Some(vec![(Some("My 'Q' Sheet".to_string()), ((1, 0), (5, 0)))])
        );
        assert_eq!(
            parse_chart_formula("(Sheet1!$A$1,Sheet1!$C$3:$C$4)"),
            Some(vec![(Some("Sheet1".to_string()), ((0, 0), (0, 0))), (Some("Sheet1".to_string()), ((2, 2), (3, 2)))])
        );
        assert_eq!(parse_chart_formula("[1]Sheet1!$A$1"), None);
        assert_eq!(parse_chart_formula("MyName"), None);
        let refs = |v: &[(&str, &str)]| v.iter().map(|(a, b)| (a.to_string(), b.to_string())).collect::<Vec<_>>();
        assert_eq!(chart_bbox(&refs(&[("tx", "s!B1"), ("cat", "S!$A$2:$A$6"), ("val", "S!$C$2:$C$6")]), "S"), Some(((0, 0), (5, 2))));
        assert_eq!(chart_bbox(&refs(&[("val", "Other!$C$2:$C$6")]), "S"), None, "foreign data");
        assert_eq!(chart_bbox(&refs(&[("cat", "S!$A$2:$A$6")]), "S"), None, "no values");
        assert_eq!(chart_bbox(&refs(&[("val", "S!$B$2:$F$2")]), "S"), None, "no data row");

        // Chart XML: rich title, 3-D bar group, axis titles ignored.
        let xml = br#"<c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Hello </a:t></a:r><a:r><a:t>&amp; bye</a:t></a:r></a:p></c:rich></c:tx></c:title><c:plotArea><c:bar3DChart><c:barDir val="bar"/><c:ser><c:tx><c:strRef><c:f>S!$B$1</c:f></c:strRef></c:tx><c:val><c:numRef><c:f>S!$B$2:$B$3</c:f></c:numRef></c:val></c:ser></c:bar3DChart><c:lineChart/><c:valAx><c:title><c:tx><c:rich><a:p><a:r><a:t>axis</a:t></a:r></a:p></c:rich></c:tx></c:title></c:valAx></c:plotArea></c:chart></c:chartSpace>"#;
        let info = parse_chart_xml(xml).unwrap();
        assert_eq!(info.kind, Some("bar"));
        assert_eq!(info.title.as_deref(), Some("Hello & bye"));
        assert_eq!(info.refs.len(), 2);
        // Cached string-reference title; doughnut -> pie; radar unsupported.
        let xml = br#"<chartSpace><chart><title><tx><strRef><f>S!$A$1</f><strCache><pt idx="0"><v>Cached</v></pt></strCache></strRef></tx></title><plotArea><doughnutChart/></plotArea></chart></chartSpace>"#;
        let info = parse_chart_xml(xml).unwrap();
        assert_eq!((info.kind, info.title.as_deref()), (Some("pie"), Some("Cached")));
        assert_eq!(parse_chart_xml(b"<chartSpace><chart><plotArea><radarChart/></plotArea></chart></chartSpace>").unwrap().kind, None);
        // Malformed / truncated / deeply nested input never panics.
        assert!(parse_chart_xml(b"<chartSpace><chart><plotArea><barChart></lineChart>").is_none());
        let deep = "<a>".repeat(10_000);
        assert!(parse_chart_xml(deep.as_bytes()).is_none());
        assert!(parse_drawing_xml(deep.as_bytes()).is_empty());
        let partial = parse_drawing_xml(b"<wsDr><twoCellAnchor><from><col>3</col><colOff>x</colOff><row>2</row></from><to><col>9");
        assert!(partial.is_empty());
        let d = parse_drawing_xml(br#"<wsDr><absoluteAnchor><pos x="952500" y="-5"/><ext cx="9525" cy="19050"/><graphicFrame><graphic><graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><chart id="rId7"/></graphicData></graphic></graphicFrame></absoluteAnchor><oneCellAnchor><from><col>99999999</col><colOff>0</colOff><row>1</row><rowOff>0</rowOff></from><ext cx="1" cy="1"/><AlternateContent><Choice><graphicFrame><graphic><graphicData uri="http://schemas.microsoft.com/office/drawing/2014/chartex"><chart id="rId8"/></graphicData></graphic></graphicFrame></Choice><Fallback><sp/></Fallback></AlternateContent></oneCellAnchor></wsDr>"#);
        assert_eq!(d.len(), 2);
        assert_eq!(d[0].anchor, Some(RawAnchor::Abs { x: 952500.0, y: 0.0, cx: 9525.0, cy: 19050.0 }));
        assert_eq!(d[0].chart_rids, vec!["rId7".to_string()]);
        assert!(!d[0].shape && !d[0].picture);
        assert_eq!((d[1].other_charts, d[1].other_chart_rids.len(), d[1].shape), (1, 1, false), "fallback shape ignored");
        assert!(matches!(d[1].anchor, Some(RawAnchor::One { from: AnchorCell { col, .. }, .. }) if col == MAX_COLS - 1));
    }

    /// Writes the round-trip workbook to $EXCEL_FEATURES_DUMP (when set) so
    /// it can be checked with openpyxl / opened in Excel.
    #[test]
    fn dump_written_workbook_for_external_check() {
        let Ok(path) = std::env::var("EXCEL_FEATURES_DUMP") else { return };
        let (bytes, warns) = write_xlsx_with_warnings(&feature_book()).unwrap();
        assert!(warns.is_empty(), "{warns:?}");
        std::fs::write(&path, bytes).unwrap();
    }

    /// Manual check against real files: OFFICE_XLSX_SAMPLES=dir cargo test -- --ignored
    #[test]
    #[ignore]
    fn parse_sample_files() {
        let Ok(dir) = std::env::var("OFFICE_XLSX_SAMPLES") else { return };
        for entry in std::fs::read_dir(dir).unwrap().flatten() {
            let p = entry.path();
            let ext = p.extension().and_then(|e| e.to_str()).unwrap_or("").to_ascii_lowercase();
            if !matches!(ext.as_str(), "xlsx" | "xls" | "xlsm" | "ods") { continue; }
            let bytes = std::fs::read(&p).unwrap();
            match parse_book(&bytes, None) {
                Ok(b) => {
                    println!("== {} : active={} warnings={:?}", p.display(), b.active, b.warnings);
                    for s in &b.sheets {
                        let json = serde_json::to_string(s).unwrap();
                        println!("  {} hidden={} cells={} {}", s.name, s.hidden, s.cells.len(), &json[..json.len().min(600)]);
                    }
                    write_xlsx(&b).expect("re-write");
                }
                Err(e) => println!("== {} : ERROR {e}", p.display()),
            }
        }
    }
}
