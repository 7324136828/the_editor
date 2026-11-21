//! Shared document model — JSON contract with the extension's lib/*.js.
//!
//! WORD CONTRACT (shared with office-chrome-extension/lib/model-word.js):
//! doc   = { kind:"word", meta:{ title?, warnings?:[string], ... }, blocks:[Block] }
//! Block = p|h1..h6|quote|code {runs, align?:"left"|"center"|"right"|"justify"}
//!       | list {list:"ul"|"ol", level:0..8, runs}       (one block per list item)
//!       | table {rows:[[{blocks:[Block], colspan?}]]}
//!       | img {src:"data:image/...;base64,..." , alt?, w?, h?}  (w/h in CSS px)
//!       | pagebreak
//! Run   = {t, b?, i?, u?, s?, sup?, sub?, color?:"#RRGGBB", size?:pt, font?, link?, hl?:"#RRGGBB"}
//! meta.warnings: fidelity-loss notes from the parser; main.rs lifts them to the
//! top-level `warnings` of the /api/parse response and removes them from meta.
//! Optional document-level meta keys (free-form JSON in `meta`):
//!   meta.title / meta.author / meta.subject : string   (docProps/core.xml)
//!   meta.page   = { size?:"A4"|"Letter"|"Legal"|"A3"|"A5"|"Executive"|"custom",
//!                   orientation?:"portrait"|"landscape", width?, height?,
//!                   margins?:{top,right,bottom,left,header?,footer?} }  (twips; see docx.rs)
//!   meta.header / meta.footer = { runs:[Run], align? }  (one paragraph each;
//!                   the literal text tokens "{PAGE}" and "{NUMPAGES}" inside a
//!                   run's `t` are page-number fields)

use serde::{Deserialize, Serialize};
use serde_json::{Map, Value};
use std::collections::BTreeMap;

// ---------------- word ----------------

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Run {
    pub t: String,
    #[serde(default, skip_serializing_if = "is_false")]
    pub b: bool,
    #[serde(default, skip_serializing_if = "is_false")]
    pub i: bool,
    #[serde(default, skip_serializing_if = "is_false")]
    pub u: bool,
    #[serde(default, skip_serializing_if = "is_false")]
    pub s: bool,
    #[serde(default, skip_serializing_if = "is_false")]
    pub sup: bool,
    #[serde(default, skip_serializing_if = "is_false")]
    pub sub: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub size: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub link: Option<String>,
    /// Highlight / background color "#RRGGBB".
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub hl: Option<String>,
}

fn is_false(b: &bool) -> bool { !*b }

impl Run {
    pub fn text(t: impl Into<String>) -> Self { Run { t: t.into(), ..Default::default() } }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TableCell {
    #[serde(default)]
    pub blocks: Vec<Block>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub colspan: Option<u32>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "type", rename_all = "lowercase")]
pub enum Block {
    P { #[serde(default)] runs: Vec<Run>, #[serde(default, skip_serializing_if = "Option::is_none")] align: Option<String> },
    H1 { #[serde(default)] runs: Vec<Run>, #[serde(default, skip_serializing_if = "Option::is_none")] align: Option<String> },
    H2 { #[serde(default)] runs: Vec<Run>, #[serde(default, skip_serializing_if = "Option::is_none")] align: Option<String> },
    H3 { #[serde(default)] runs: Vec<Run>, #[serde(default, skip_serializing_if = "Option::is_none")] align: Option<String> },
    H4 { #[serde(default)] runs: Vec<Run>, #[serde(default, skip_serializing_if = "Option::is_none")] align: Option<String> },
    H5 { #[serde(default)] runs: Vec<Run>, #[serde(default, skip_serializing_if = "Option::is_none")] align: Option<String> },
    H6 { #[serde(default)] runs: Vec<Run>, #[serde(default, skip_serializing_if = "Option::is_none")] align: Option<String> },
    Quote { #[serde(default)] runs: Vec<Run>, #[serde(default, skip_serializing_if = "Option::is_none")] align: Option<String> },
    Code { #[serde(default)] runs: Vec<Run>, #[serde(default, skip_serializing_if = "Option::is_none")] align: Option<String> },
    List { list: String, #[serde(default)] level: u32, #[serde(default)] runs: Vec<Run> },
    Table { #[serde(default)] rows: Vec<Vec<TableCell>> },
    Img { src: String, #[serde(default, skip_serializing_if = "Option::is_none")] alt: Option<String>, #[serde(default, skip_serializing_if = "Option::is_none")] w: Option<u32>, #[serde(default, skip_serializing_if = "Option::is_none")] h: Option<u32> },
    #[serde(rename = "pagebreak")]
    PageBreak,
    /// Unknown block types round-trip as raw JSON so nothing is silently lost.
    #[serde(untagged)]
    Other(Value),
}

impl Block {
    pub fn para_runs(&self) -> &[Run] {
        match self {
            Block::P { runs, .. } | Block::H1 { runs, .. } | Block::H2 { runs, .. }
            | Block::H3 { runs, .. } | Block::H4 { runs, .. } | Block::H5 { runs, .. }
            | Block::H6 { runs, .. } | Block::Quote { runs, .. } | Block::Code { runs, .. }
            | Block::List { runs, .. } => runs,
            _ => &[],
        }
    }
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct WordDoc {
    #[serde(default = "default_word_kind")]
    pub kind: String,
    #[serde(default)]
    pub meta: Map<String, Value>,
    #[serde(default)]
    pub blocks: Vec<Block>,
}

fn default_word_kind() -> String { "word".into() }

// ---------------- excel ----------------
//
// CONTRACT (shared with office-chrome-extension/lib/model-excel.js):
// book  = { kind:"excel", active:number, sheets:[Sheet], warnings?:[string] }
// sheet = { name, hidden?:bool, cells:{ "A1": Cell }, colWidths?:{ "A": chars },
//           merges?:["A1:B2", ...], freeze?:{rows, cols}, autoFilter?:"A1:D20",
//           charts?:[Chart] }
// Chart = { type:"column"|"bar"|"line"|"pie"|"scatter"|"area", title?, range:"A1:C10"
//           (data INCLUDING a header row; first column = categories / x values,
//           each further column = one series named by its header cell),
//           x, y, w, h (px, relative to the grid's top-left cell A1 at 100% zoom) }
// cell  = { v:any, t:"n"|"s"|"b"|"d"|"e"|"f", f?:"=FORMULA", z?:"number format code",
//           s?:CellStyle }
//   t "d": v is a NAIVE local ISO string "YYYY-MM-DD" or "YYYY-MM-DDTHH:MM:SS" (no TZ).
//   t "e": v is the Excel error text, e.g. "#DIV/0!", "#N/A".
//   f: formula text INCLUDING leading "="; v is the cached value (may be null).
// CellStyle = { b?, i?, u?, color?:"#RRGGBB", bg?:"#RRGGBB", align?:"left"|"center"|"right",
//               wrap? }  — all optional; unknown keys must be ignored/preserved.

#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
pub struct CellStyle {
    #[serde(default, skip_serializing_if = "is_false")]
    pub b: bool,
    #[serde(default, skip_serializing_if = "is_false")]
    pub i: bool,
    #[serde(default, skip_serializing_if = "is_false")]
    pub u: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bg: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub align: Option<String>,
    #[serde(default, skip_serializing_if = "is_false")]
    pub wrap: bool,
}

impl CellStyle {
    pub fn is_empty(&self) -> bool { *self == CellStyle::default() }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Cell {
    #[serde(default)]
    pub v: Value,
    #[serde(default = "default_cell_t")]
    pub t: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub f: Option<String>,
    /// Number format code (e.g. "0.00%", "yyyy-mm-dd"); None = General.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub z: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub s: Option<CellStyle>,
}

fn default_cell_t() -> String { "s".into() }

impl Cell {
    pub fn new(v: Value, t: &str) -> Self { Cell { v, t: t.into(), f: None, z: None, s: None } }
}

#[derive(Debug, Clone, Default)]
pub struct Sheet {
    pub name: String,
    pub hidden: bool,
    pub cells: BTreeMap<String, Cell>,
    pub col_widths: Option<BTreeMap<String, f64>>,
    /// Merged ranges as "A1:B2".
    pub merges: Vec<String>,
    /// Frozen panes: rows above / columns left of the split stay visible.
    pub freeze: Option<Freeze>,
    /// AutoFilter range, e.g. "A1:D20" (JSON key `autoFilter`).
    pub auto_filter: Option<String>,
    /// Charts (raw JSON, see CONTRACT above; interpreted by xlsx.rs).
    pub charts: Vec<Value>,
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Serialize, Deserialize)]
pub struct Freeze {
    #[serde(default)]
    pub rows: u32,
    #[serde(default)]
    pub cols: u32,
}

// colWidths in JS is camelCase; accept both.
impl<'de> Deserialize<'de> for Sheet {
    fn deserialize<D: serde::Deserializer<'de>>(d: D) -> Result<Self, D::Error> {
        #[derive(Deserialize)]
        struct Raw {
            name: String,
            #[serde(default)] hidden: bool,
            #[serde(default)] cells: BTreeMap<String, Cell>,
            #[serde(default)] col_widths: Option<BTreeMap<String, f64>>,
            #[serde(default, rename = "colWidths")] col_widths_js: Option<BTreeMap<String, f64>>,
            #[serde(default)] merges: Vec<String>,
            #[serde(default)] freeze: Option<Freeze>,
            #[serde(default, rename = "autoFilter")] auto_filter: Option<String>,
            #[serde(default)] charts: Vec<Value>,
        }
        let r = Raw::deserialize(d)?;
        Ok(Sheet {
            name: r.name, hidden: r.hidden, cells: r.cells,
            col_widths: r.col_widths.or(r.col_widths_js), merges: r.merges,
            freeze: r.freeze.filter(|f| f.rows > 0 || f.cols > 0),
            auto_filter: r.auto_filter.filter(|a| !a.trim().is_empty()),
            charts: r.charts,
        })
    }
}

impl Serialize for Sheet {
    fn serialize<S: serde::Serializer>(&self, s: S) -> Result<S::Ok, S::Error> {
        #[derive(Serialize)]
        struct Raw<'a> {
            name: &'a str,
            #[serde(skip_serializing_if = "is_false")] hidden: bool,
            cells: &'a BTreeMap<String, Cell>,
            #[serde(skip_serializing_if = "Option::is_none", rename = "colWidths")] col_widths: Option<&'a BTreeMap<String, f64>>,
            #[serde(skip_serializing_if = "<[String]>::is_empty")] merges: &'a [String],
            #[serde(skip_serializing_if = "Option::is_none")] freeze: Option<&'a Freeze>,
            #[serde(skip_serializing_if = "Option::is_none", rename = "autoFilter")] auto_filter: Option<&'a String>,
            #[serde(skip_serializing_if = "<[Value]>::is_empty")] charts: &'a [Value],
        }
        Raw {
            name: &self.name, hidden: self.hidden, cells: &self.cells,
            col_widths: self.col_widths.as_ref(), merges: &self.merges,
            freeze: self.freeze.as_ref(), auto_filter: self.auto_filter.as_ref(),
            charts: &self.charts,
        }.serialize(s)
    }
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct ExcelBook {
    #[serde(default = "default_excel_kind")]
    pub kind: String,
    #[serde(default)]
    pub active: u32,
    #[serde(default)]
    pub sheets: Vec<Sheet>,
    /// Fidelity-loss notes produced by the parser; main.rs lifts these into
    /// the top-level `warnings` of the /api/parse response and removes them.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub warnings: Vec<String>,
}

fn default_excel_kind() -> String { "excel".into() }
// ---------------- slides ----------------
//
// CONTRACT (shared with office-chrome-extension/lib/model-slides.js):
// deck  = { kind:"slides", meta:{ title?, author?, warnings?:[string] },
//           size:{ w, h }              points; 16:9 = 960x540 (default), 4:3 = 720x540
//           theme:{ font?, headFont?, color?:"#RRGGBB" (body text), accent?:"#RRGGBB",
//                   bg?:"#RRGGBB" (default slide background) },
//           slides:[Slide] }
// slide = { layout?:"title"|"content"|"section"|"two"|"blank", bg?:"#RRGGBB",
//           notes?:string (speaker notes, "\n" between paragraphs), hidden?:bool,
//           transition?:"none"|"fade"|"push"|"wipe"|"zoom", items:[Item] }  (z-order = array order)
// Item  — geometry in points from the slide's top-left; rot = degrees clockwise:
//   text  { type:"text",  x,y,w,h, rot?, ph?:"title"|"subtitle"|"body", paras:[Para],
//           fill?, line?, lineW?, valign?:"top"|"middle"|"bottom" }
//   shape { type:"shape", x,y,w,h, rot?, shape:"rect"|"roundRect"|"ellipse"|"triangle"|"diamond"
//           |"rightArrow"|"leftArrow"|"line"|"star5"|"hexagon"|"chevron", fill?, line?, lineW?,
//           paras?:[Para], valign? }
//   image { type:"image", x,y,w,h, rot?, src:"data:image/...;base64,...", alt? }
//   table { type:"table", x,y,w,h, rows:[[{ paras:[Para], fill? }]], colW?:[pt] }
//   Unknown item types round-trip as raw JSON.
// Para  = { runs:[Run], align?:"left"|"center"|"right"|"justify", level?:0..8,
//           bullet?:"ul"|"ol" }     (Run = the word Run; size in pt)
// Colors: "#RRGGBB"; an absent fill/line means none (transparent / no outline).
// lineW in points. meta.warnings is lifted into the /api/parse `warnings`.

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Para {
    #[serde(default)]
    pub runs: Vec<Run>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub align: Option<String>,
    #[serde(default, skip_serializing_if = "is_zero")]
    pub level: u32,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bullet: Option<String>,
}

fn is_zero(n: &u32) -> bool { *n == 0 }
fn is_zero_f(n: &f64) -> bool { *n == 0.0 }

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SlideSize {
    #[serde(default = "default_slide_w")]
    pub w: f64,
    #[serde(default = "default_slide_h")]
    pub h: f64,
}
impl Default for SlideSize {
    fn default() -> Self { SlideSize { w: 960.0, h: 540.0 } }
}
fn default_slide_w() -> f64 { 960.0 }
fn default_slide_h() -> f64 { 540.0 }

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Theme {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font: Option<String>,
    #[serde(default, rename = "headFont", skip_serializing_if = "Option::is_none")]
    pub head_font: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub accent: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bg: Option<String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct TextItem {
    #[serde(default)] pub x: f64,
    #[serde(default)] pub y: f64,
    #[serde(default)] pub w: f64,
    #[serde(default)] pub h: f64,
    #[serde(default, skip_serializing_if = "is_zero_f")]
    pub rot: f64,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub ph: Option<String>,
    #[serde(default)]
    pub paras: Vec<Para>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line: Option<String>,
    #[serde(default, rename = "lineW", skip_serializing_if = "Option::is_none")]
    pub line_w: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub valign: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ShapeItem {
    #[serde(default)] pub x: f64,
    #[serde(default)] pub y: f64,
    #[serde(default)] pub w: f64,
    #[serde(default)] pub h: f64,
    #[serde(default, skip_serializing_if = "is_zero_f")]
    pub rot: f64,
    #[serde(default = "default_shape")]
    pub shape: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line: Option<String>,
    #[serde(default, rename = "lineW", skip_serializing_if = "Option::is_none")]
    pub line_w: Option<f64>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub paras: Vec<Para>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub valign: Option<String>,
}
impl Default for ShapeItem {
    fn default() -> Self {
        ShapeItem { x: 0.0, y: 0.0, w: 0.0, h: 0.0, rot: 0.0, shape: default_shape(), fill: None, line: None, line_w: None, paras: Vec::new(), valign: None }
    }
}
fn default_shape() -> String { "rect".into() }

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct ImageItem {
    #[serde(default)] pub x: f64,
    #[serde(default)] pub y: f64,
    #[serde(default)] pub w: f64,
    #[serde(default)] pub h: f64,
    #[serde(default, skip_serializing_if = "is_zero_f")]
    pub rot: f64,
    pub src: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub alt: Option<String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct SlideTableCell {
    #[serde(default)]
    pub paras: Vec<Para>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill: Option<String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct TableItem {
    #[serde(default)] pub x: f64,
    #[serde(default)] pub y: f64,
    #[serde(default)] pub w: f64,
    #[serde(default)] pub h: f64,
    #[serde(default)]
    pub rows: Vec<Vec<SlideTableCell>>,
    #[serde(default, rename = "colW", skip_serializing_if = "Vec::is_empty")]
    pub col_w: Vec<f64>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "type", rename_all = "lowercase")]
pub enum Item {
    Text(TextItem),
    Shape(ShapeItem),
    Image(ImageItem),
    Table(TableItem),
    /// Unknown item types round-trip as raw JSON.
    #[serde(untagged)]
    Other(Value),
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Slide {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub layout: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bg: Option<String>,
    #[serde(default, skip_serializing_if = "String::is_empty")]
    pub notes: String,
    #[serde(default, skip_serializing_if = "is_false")]
    pub hidden: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub transition: Option<String>,
    #[serde(default)]
    pub items: Vec<Item>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct SlidesDoc {
    #[serde(default = "default_slides_kind")]
    pub kind: String,
    #[serde(default)]
    pub meta: Map<String, Value>,
    #[serde(default)]
    pub size: SlideSize,
    #[serde(default)]
    pub theme: Theme,
    #[serde(default)]
    pub slides: Vec<Slide>,
}
fn default_slides_kind() -> String { "slides".into() }

// ---------------- latex ----------------
//
// CONTRACT: a LaTeX document travels as its source text only:
//   { kind:"latex", source:"\documentclass{article} ..." }
// /api/render kind "latex": pdf via an installed TeX engine (only advertised
// when one is found); docx/html/md/txt via latex::parse_latex -> WordDoc ->
// the word writers; tex returns the source unchanged.
// Word documents can also be rendered as format "tex" (latex::write_latex).

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct LatexDoc {
    #[serde(default = "default_latex_kind")]
    pub kind: String,
    #[serde(default)]
    pub source: String,
}
fn default_latex_kind() -> String { "latex".into() }

// ---------------- helpers ----------------


