//! PowerPoint (.pptx) <-> slides model (see model.rs "slides" CONTRACT).
//!
//! READ (`parse_pptx`): the presentation part is located through
//! `_rels/.rels`; slide order and size come from `presentation.xml`. Every
//! slide is walked together with its layout and master: placeholder geometry
//! (`spPr/xfrm`), fills, body properties and text styles are inherited
//! slide -> layout -> master (placeholders matched by `idx`, then by type),
//! text defaults come from the master `txStyles` / `presentation.xml`
//! `defaultTextStyle` and the placeholders' `lstStyle`s, so runs carry their
//! EFFECTIVE size, color, font and bold/italic (colors equal to `theme.color`
//! and fonts equal to the context default — `theme.headFont` for titles,
//! `theme.font` otherwise — are left implicit). Theme colors (`schemeClr`
//! through the master/layout/slide color map, `lumMod`/`lumOff`/`tint`/
//! `shade`/`satMod`) and theme fonts (`+mj-lt`/`+mn-lt`) are resolved.
//! Non-placeholder decorations of the layout and master (logos, lines) are
//! copied onto each slide (behind the slide's own content) unless
//! `showMasterSp="0"`. Groups are flattened with their child transform.
//! Charts, SmartArt, OLE objects and media become light-gray placeholder
//! shapes (or their preview picture) plus one user-facing warning per kind.
//!
//! WRITE (`write_pptx`): a complete package PowerPoint opens without repair:
//! [Content_Types].xml, _rels/.rels, docProps/{core,app}.xml,
//! ppt/presentation.xml (+rels), presProps/viewProps/tableStyles, theme1.xml
//! built from `doc.theme`, one slide master with five layouts (Title Slide,
//! Title and Content, Section Header, Two Content, Blank), the slides, media
//! parts, and a notes master + notes slides when a slide has speaker notes.
//! Every paragraph states its alignment, bullet and indentation explicitly and
//! every shape its anchor, so the result does not depend on master defaults.
//!
//! Line breaks inside a paragraph (`a:br`) are "\n" inside a run's `t`.

use crate::docx::xml::{self, Node};
use crate::docx::{
    data_url, mime_for_ext, norm_hex, parse_data_url, read_zip_entry, safe_link, sniff_image, xml_attr,
    xml_bytes_to_string, xml_text,
};
use crate::model::{
    ImageItem, Item, Para, Run, ShapeItem, Slide, SlideSize, SlideTableCell, SlidesDoc, TableItem, TextItem, Theme,
};
use serde_json::Value;
use std::collections::{BTreeSet, HashMap};
use std::io::{Cursor, Write};
use std::rc::Rc;
use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

const EMU_PER_PT: f64 = 12700.0;
/// Hostile-input caps.
const MAX_SLIDES: usize = 2000;
const MAX_ITEMS: usize = 4000;
const MAX_GROUP_DEPTH: usize = 24;
const MAX_PARAS: usize = 20_000;
const MAX_RUNS: usize = 5_000;
const MAX_TABLE_CELLS: usize = 20_000;
const MAX_XML_PART: u64 = 128 * 1024 * 1024;
const MAX_IMAGE: u64 = 32 * 1024 * 1024;
const MAX_MEDIA_TOTAL: usize = 192 * 1024 * 1024;

// User-facing warnings (deduplicated per document).
const W_SHAPES: &str = "Some shapes (such as clouds, arcs or freeform drawings) are shown as rectangles.";
const W_CHARTS: &str = "Charts are shown as gray placeholders and are not kept when saving.";
const W_SMARTART: &str = "SmartArt graphics are shown as gray placeholders and are not kept when saving.";
const W_MEDIA: &str = "Video and audio clips cannot be played here; they are shown as their preview picture or a placeholder and are not kept when saving.";
const W_OLE: &str = "Embedded objects (such as spreadsheets) are shown as their preview picture or a placeholder and are not kept when saving.";
const W_OTHER_OBJ: &str = "Some objects (such as 3D models or ink drawings) could not be displayed and are not kept when saving.";
const W_GRADIENT: &str = "Gradient, pattern and picture fills are shown as a solid color.";
const W_IMG_FORMAT: &str = "Some pictures use formats browsers cannot display (EMF/WMF/TIFF) and were left out.";
const W_IMG_LIMIT: &str = "Some pictures were too large to load and were left out.";
const W_IMG_LINKED: &str = "Linked (not embedded) pictures could not be loaded.";
const W_MERGED: &str = "Merged table cells were split into separate cells.";
const W_ANIM: &str = "Animations are not shown and are not kept when saving.";
const W_COMMENTS: &str = "Comments are not shown and are not kept when saving.";
const W_TRANSITION: &str = "Some slide transitions were replaced by the closest available effect.";
const W_ITEMS: &str = "Some slides contain too many objects; only the first ones were opened.";
const W_DEPTH: &str = "Some deeply nested groups of shapes were left out.";
const W_SLIDE_BROKEN: &str = "One or more slides are damaged and were skipped.";
const W_SLIDE_LIST: &str = "The slide list of this file is damaged; the slides were opened in file order.";

fn keep_ws_pptx(parent: &str) -> bool {
    parent == "t"
}

fn bad(msg: &str) -> String {
    format!("This is not a valid PowerPoint file: {msg}.")
}

fn pt(emu: f64) -> f64 {
    if !emu.is_finite() {
        return 0.0;
    }
    (emu / EMU_PER_PT * 100.0).round() / 100.0
}

fn num(n: &Node, a: &str) -> Option<f64> {
    n.attr(a).and_then(|v| v.trim().parse::<f64>().ok()).filter(|v| v.is_finite())
}

fn hex(c: [u8; 3]) -> String {
    format!("#{:02X}{:02X}{:02X}", c[0], c[1], c[2])
}

fn parse_hex(s: &str) -> Option<[u8; 3]> {
    let h = norm_hex(s)?;
    let v = u32::from_str_radix(&h, 16).ok()?;
    Some([(v >> 16) as u8, (v >> 8) as u8, v as u8])
}

/// `r:id`-style attribute (a prefixed attribute with the given local name),
/// so `sldId id="256" r:id="rId2"` resolves to "rId2".
fn rel_attr<'a>(n: &'a Node, local: &str) -> Option<&'a str> {
    n.attrs
        .iter()
        .find(|(k, _)| k.contains(':') && xml::local_name(k) == local)
        .map(|(_, v)| v.as_str())
}

fn split_part(path: &str) -> (String, String) {
    match path.rsplit_once('/') {
        Some((d, f)) => (format!("{d}/"), f.to_string()),
        None => (String::new(), path.to_string()),
    }
}

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

// ============================ colors ============================

fn rgb_to_hsl(c: [u8; 3]) -> (f64, f64, f64) {
    let r = c[0] as f64 / 255.0;
    let g = c[1] as f64 / 255.0;
    let b = c[2] as f64 / 255.0;
    let max = r.max(g).max(b);
    let min = r.min(g).min(b);
    let l = (max + min) / 2.0;
    if (max - min).abs() < 1e-9 {
        return (0.0, 0.0, l);
    }
    let d = max - min;
    let s = if l > 0.5 { d / (2.0 - max - min) } else { d / (max + min) };
    let h = if max == r {
        (g - b) / d + if g < b { 6.0 } else { 0.0 }
    } else if max == g {
        (b - r) / d + 2.0
    } else {
        (r - g) / d + 4.0
    };
    (h / 6.0, s, l)
}

fn hsl_to_rgb(h: f64, s: f64, l: f64) -> [u8; 3] {
    let to = |v: f64| (v.clamp(0.0, 1.0) * 255.0).round() as u8;
    if s <= 0.0 {
        return [to(l), to(l), to(l)];
    }
    let q = if l < 0.5 { l * (1.0 + s) } else { l + s - l * s };
    let p = 2.0 * l - q;
    let hue = |mut t: f64| {
        if t < 0.0 {
            t += 1.0;
        }
        if t > 1.0 {
            t -= 1.0;
        }
        if t < 1.0 / 6.0 {
            p + (q - p) * 6.0 * t
        } else if t < 0.5 {
            q
        } else if t < 2.0 / 3.0 {
            p + (q - p) * (2.0 / 3.0 - t) * 6.0
        } else {
            p
        }
    };
    [to(hue(h + 1.0 / 3.0)), to(hue(h)), to(hue(h - 1.0 / 3.0))]
}

/// Apply DrawingML color transforms (approximations good enough for display).
fn apply_mods(mut c: [u8; 3], el: &Node) -> [u8; 3] {
    for m in el.elems() {
        let v = num(m, "val").unwrap_or(100_000.0) / 100_000.0;
        match m.local() {
            "lumMod" | "lumOff" | "satMod" | "satOff" | "hueOff" | "hueMod" | "lum" | "sat" | "hue" => {
                let (mut h, mut s, mut l) = rgb_to_hsl(c);
                match m.local() {
                    "lumMod" => l *= v,
                    "lumOff" => l += v,
                    "satMod" => s *= v,
                    "satOff" => s += v,
                    "lum" => l = v,
                    "sat" => s = v,
                    "hue" => h = num(m, "val").unwrap_or(0.0) / 21_600_000.0,
                    "hueOff" => h += num(m, "val").unwrap_or(0.0) / 21_600_000.0,
                    "hueMod" => h *= v,
                    _ => {}
                }
                c = hsl_to_rgb(h.rem_euclid(1.0), s.clamp(0.0, 1.0), l.clamp(0.0, 1.0));
            }
            "tint" => {
                let t = v.clamp(0.0, 1.0);
                c = c.map(|x| (x as f64 + (255.0 - x as f64) * (1.0 - t)).round() as u8);
            }
            "shade" => {
                let t = v.clamp(0.0, 1.0);
                c = c.map(|x| (x as f64 * t).round() as u8);
            }
            "comp" => {
                let (h, s, l) = rgb_to_hsl(c);
                c = hsl_to_rgb((h + 0.5).rem_euclid(1.0), s, l);
            }
            "inv" => c = c.map(|x| 255 - x),
            "gray" => {
                let g = (0.299 * c[0] as f64 + 0.587 * c[1] as f64 + 0.114 * c[2] as f64).round() as u8;
                c = [g, g, g];
            }
            _ => {}
        }
    }
    c
}

fn preset_color(name: &str) -> Option<[u8; 3]> {
    const P: &[(&str, &str)] = &[
        ("black", "000000"), ("white", "FFFFFF"), ("red", "FF0000"), ("green", "008000"), ("blue", "0000FF"),
        ("yellow", "FFFF00"), ("cyan", "00FFFF"), ("aqua", "00FFFF"), ("magenta", "FF00FF"), ("fuchsia", "FF00FF"),
        ("gray", "808080"), ("grey", "808080"), ("silver", "C0C0C0"), ("maroon", "800000"), ("navy", "000080"),
        ("olive", "808000"), ("purple", "800080"), ("teal", "008080"), ("lime", "00FF00"), ("orange", "FFA500"),
        ("ltGray", "D3D3D3"), ("lightGray", "D3D3D3"), ("dkGray", "A9A9A9"), ("darkGray", "A9A9A9"),
        ("dkBlue", "00008B"), ("darkBlue", "00008B"), ("dkRed", "8B0000"), ("darkRed", "8B0000"),
        ("dkGreen", "006400"), ("darkGreen", "006400"), ("gold", "FFD700"), ("pink", "FFC0CB"), ("brown", "A52A2A"),
    ];
    P.iter().find(|(n, _)| n.eq_ignore_ascii_case(name)).and_then(|(_, h)| parse_hex(h))
}

#[derive(Clone)]
struct ColorCtx {
    scheme: Rc<HashMap<String, [u8; 3]>>,
    map: HashMap<String, String>,
}

impl ColorCtx {
    fn scheme_color(&self, name: &str) -> Option<[u8; 3]> {
        let mapped = self.map.get(name).map(String::as_str).unwrap_or(match name {
            "bg1" => "lt1",
            "tx1" => "dk1",
            "bg2" => "lt2",
            "tx2" => "dk2",
            n => n,
        });
        self.scheme.get(mapped).or_else(|| self.scheme.get(name)).copied()
    }

    /// Color of the first color element inside `parent` (e.g. a solidFill).
    fn color_in(&self, parent: &Node, ph: Option<[u8; 3]>) -> Option<[u8; 3]> {
        for c in parent.elems() {
            let base = match c.local() {
                "srgbClr" => c.attr("val").and_then(parse_hex),
                "schemeClr" => match c.attr("val") {
                    Some("phClr") => ph.or_else(|| self.scheme_color("accent1")),
                    Some(v) => self.scheme_color(v),
                    None => None,
                },
                "sysClr" => c.attr("lastClr").and_then(parse_hex).or(match c.attr("val") {
                    Some("window") => Some([255, 255, 255]),
                    Some(_) => Some([0, 0, 0]),
                    None => None,
                }),
                "prstClr" => c.attr("val").and_then(preset_color),
                "scrgbClr" => {
                    let ch = |a| {
                        let lin = (num(c, a).unwrap_or(0.0) / 100_000.0).clamp(0.0, 1.0);
                        let s = if lin <= 0.003_130_8 { lin * 12.92 } else { 1.055 * lin.powf(1.0 / 2.4) - 0.055 };
                        (s * 255.0).round() as u8
                    };
                    Some([ch("r"), ch("g"), ch("b")])
                }
                "hslClr" => Some(hsl_to_rgb(
                    num(c, "hue").unwrap_or(0.0) / 21_600_000.0,
                    num(c, "sat").unwrap_or(0.0) / 100_000.0,
                    num(c, "lum").unwrap_or(0.0) / 100_000.0,
                )),
                _ => continue,
            };
            return base.map(|b| apply_mods(b, c));
        }
        None
    }
}

fn clr_map_of(n: &Node) -> HashMap<String, String> {
    n.attrs.iter().map(|(k, v)| (xml::local_name(k).to_string(), v.clone())).collect()
}

/// A fill as far as the model can show it.
#[derive(Clone, Debug, PartialEq)]
enum Fill {
    None,
    Solid([u8; 3]),
    /// Gradient / pattern: shown as one color (warning).
    Approx([u8; 3]),
    /// Picture fill: relationship id of the blip.
    Picture(String),
}

fn fill_from_el(el: &Node, cc: &ColorCtx, ph: Option<[u8; 3]>) -> Option<Fill> {
    match el.local() {
        "noFill" => Some(Fill::None),
        "solidFill" => Some(cc.color_in(el, ph).map(Fill::Solid).unwrap_or(Fill::None)),
        "gradFill" => {
            let mut stops = Vec::new();
            if let Some(lst) = el.child("gsLst") {
                for gs in lst.elems().filter(|g| g.is("gs")) {
                    if let Some(c) = cc.color_in(gs, ph) {
                        stops.push(c);
                    }
                }
            }
            if stops.is_empty() {
                return Some(Fill::None);
            }
            let n = stops.len() as u32;
            let avg = |i: usize| (stops.iter().map(|s| s[i] as u32).sum::<u32>() / n) as u8;
            Some(Fill::Approx([avg(0), avg(1), avg(2)]))
        }
        "pattFill" => Some(
            el.child("fgClr").and_then(|f| cc.color_in(f, ph)).map(Fill::Approx).unwrap_or(Fill::None),
        ),
        "blipFill" => el.child("blip").and_then(|b| rel_attr(b, "embed")).map(|r| Fill::Picture(r.to_string())),
        _ => None,
    }
}

/// Fill declared directly in an spPr / bgPr.
fn fill_of(sppr: &Node, cc: &ColorCtx, ph: Option<[u8; 3]>) -> Option<Fill> {
    sppr.elems().find_map(|c| fill_from_el(c, cc, ph))
}

/// Outline: Some((color or None for "no line", width pt)).
#[derive(Clone, Debug, Default)]
struct Line {
    color: Option<[u8; 3]>,
    width: Option<f64>,
    /// The fill of the line was stated (solid or none).
    stated: bool,
}

fn line_from_ln(ln: &Node, cc: &ColorCtx, ph: Option<[u8; 3]>) -> Line {
    let mut l = Line { width: num(ln, "w").map(pt), ..Default::default() };
    if let Some(f) = ln.elems().find_map(|c| fill_from_el(c, cc, ph)) {
        l.stated = true;
        l.color = match f {
            Fill::Solid(c) | Fill::Approx(c) => Some(c),
            _ => None,
        };
    }
    l
}

// ============================ theme ============================

struct ThemeInfo {
    scheme: Rc<HashMap<String, [u8; 3]>>,
    major: String,
    minor: String,
    fill_styles: Vec<Node>,
    ln_styles: Vec<Node>,
    bg_styles: Vec<Node>,
}

fn office_scheme() -> HashMap<String, [u8; 3]> {
    [
        ("dk1", "000000"), ("lt1", "FFFFFF"), ("dk2", "44546A"), ("lt2", "E7E6E6"), ("accent1", "4472C4"),
        ("accent2", "ED7D31"), ("accent3", "A5A5A5"), ("accent4", "FFC000"), ("accent5", "5B9BD5"),
        ("accent6", "70AD47"), ("hlink", "0563C1"), ("folHlink", "954F72"),
    ]
    .iter()
    .filter_map(|(k, v)| parse_hex(v).map(|c| (k.to_string(), c)))
    .collect()
}

fn theme_from_node(node: Option<&Node>) -> ThemeInfo {
    let mut scheme = office_scheme();
    let mut major = "Calibri Light".to_string();
    let mut minor = "Calibri".to_string();
    let (mut fill_styles, mut ln_styles, mut bg_styles) = (Vec::new(), Vec::new(), Vec::new());
    if let Some(el) = node.and_then(|n| n.find("themeElements")) {
        if let Some(cs) = el.child("clrScheme") {
            let empty = ColorCtx { scheme: Rc::new(HashMap::new()), map: HashMap::new() };
            for c in cs.elems() {
                if let Some(col) = empty.color_in(c, None) {
                    scheme.insert(c.local().to_string(), col);
                }
            }
        }
        if let Some(fs) = el.child("fontScheme") {
            let face = |n: Option<&Node>| {
                n.and_then(|f| f.child("latin")).and_then(|l| l.attr("typeface")).map(str::trim).filter(|s| !s.is_empty()).map(String::from)
            };
            if let Some(f) = face(fs.child("majorFont")) {
                major = f;
            }
            if let Some(f) = face(fs.child("minorFont")) {
                minor = f;
            }
        }
        if let Some(fm) = el.child("fmtScheme") {
            let list = |name: &str| fm.child(name).map(|l| l.elems().cloned().collect::<Vec<_>>()).unwrap_or_default();
            fill_styles = list("fillStyleLst");
            ln_styles = list("lnStyleLst");
            bg_styles = list("bgFillStyleLst");
        }
    }
    ThemeInfo { scheme: Rc::new(scheme), major, minor, fill_styles, ln_styles, bg_styles }
}

// ============================ package reader ============================

#[derive(Default)]
struct Rels(HashMap<String, Rel>);

#[derive(Clone)]
struct Rel {
    kind: String,
    target: String,
    external: bool,
}

impl Rels {
    fn get(&self, id: &str) -> Option<&Rel> {
        self.0.get(id)
    }
    fn first(&self, kind: &str) -> Option<&Rel> {
        let mut v: Vec<(&String, &Rel)> = self.0.iter().filter(|(_, r)| r.kind == kind && !r.external).collect();
        v.sort_by(|a, b| a.0.cmp(b.0));
        v.first().map(|(_, r)| *r)
    }
    fn has(&self, kind: &str) -> bool {
        self.0.values().any(|r| r.kind == kind)
    }
}

struct MasterInfo {
    node: Node,
    rels: Rels,
    clr_map: HashMap<String, String>,
    theme: Rc<ThemeInfo>,
}

struct LayoutInfo {
    node: Node,
    rels: Rels,
    master: Rc<MasterInfo>,
    kind: Option<String>,
}

struct Reader<'a> {
    zip: ZipArchive<Cursor<&'a [u8]>>,
    lower: HashMap<String, String>,
    warnings: Vec<String>,
    media_cache: HashMap<String, Option<String>>,
    media_total: usize,
    themes: HashMap<String, Rc<ThemeInfo>>,
    masters: HashMap<String, Rc<MasterInfo>>,
    layouts: HashMap<String, Rc<LayoutInfo>>,
    deftext: Rc<Node>,
    slide_w: f64,
    slide_h: f64,
    deck_color: Option<[u8; 3]>,
    deck_major: String,
    deck_minor: String,
}

impl<'a> Reader<'a> {
    fn warn(&mut self, msg: &str) {
        if !self.warnings.iter().any(|w| w == msg) {
            self.warnings.push(msg.to_string());
        }
    }

    fn real_name(&self, name: &str) -> Option<String> {
        if let Some(r) = self.lower.get(&name.to_lowercase()) {
            return Some(r.clone());
        }
        let dec = String::from_utf8_lossy(&crate::docx::percent_decode(name)).to_lowercase();
        self.lower.get(&dec).cloned()
    }

    fn exists(&self, name: &str) -> bool {
        self.real_name(name).is_some()
    }

    fn read(&mut self, name: &str, cap: u64) -> Result<Vec<u8>, String> {
        let real = self.real_name(name).ok_or_else(|| format!("{name} is missing"))?;
        read_zip_entry(&mut self.zip, &real, cap)
    }

    fn xml(&mut self, name: &str) -> Result<Node, String> {
        let b = self.read(name, MAX_XML_PART)?;
        xml::parse(&xml_bytes_to_string(&b), keep_ws_pptx).map_err(|e| format!("{name}: {e}"))
    }

    fn rels(&mut self, part: &str) -> Rels {
        let (dir, file) = split_part(part);
        let rp = format!("{dir}_rels/{file}.rels");
        let mut out = Rels::default();
        let Ok(node) = self.xml(&rp) else { return out };
        for r in node.elems().filter(|n| n.is("Relationship")) {
            let (Some(id), Some(t)) = (r.attr("Id"), r.attr("Target")) else { continue };
            let kind = r.attr("Type").unwrap_or("").rsplit('/').next().unwrap_or("").to_string();
            let external = r.attr("TargetMode").map(|m| m.eq_ignore_ascii_case("External")).unwrap_or(false);
            let target = if external { t.trim().to_string() } else { resolve_target(&dir, t) };
            out.0.insert(id.to_string(), Rel { kind, target, external });
        }
        out
    }

    fn theme(&mut self, path: Option<&str>) -> Rc<ThemeInfo> {
        let key = path.unwrap_or("").to_string();
        if let Some(t) = self.themes.get(&key) {
            return t.clone();
        }
        let node = path.and_then(|p| self.xml(p).ok());
        let t = Rc::new(theme_from_node(node.as_ref()));
        self.themes.insert(key, t.clone());
        t
    }

    fn master(&mut self, path: &str) -> Rc<MasterInfo> {
        if let Some(m) = self.masters.get(path) {
            return m.clone();
        }
        let node = self.xml(path).unwrap_or_default();
        let rels = self.rels(path);
        let theme_path = rels.first("theme").map(|r| r.target.clone());
        let theme = self.theme(theme_path.as_deref());
        let clr_map = node.child("clrMap").map(clr_map_of).unwrap_or_default();
        let m = Rc::new(MasterInfo { node, rels, clr_map, theme });
        self.masters.insert(path.to_string(), m.clone());
        m
    }

    fn layout(&mut self, path: &str) -> Rc<LayoutInfo> {
        if let Some(l) = self.layouts.get(path) {
            return l.clone();
        }
        let node = self.xml(path).unwrap_or_default();
        let rels = self.rels(path);
        let master_path = rels.first("slideMaster").map(|r| r.target.clone()).unwrap_or_default();
        let master = self.master(&master_path);
        let kind = layout_kind(&node);
        let l = Rc::new(LayoutInfo { node, rels, master, kind });
        self.layouts.insert(path.to_string(), l.clone());
        l
    }

    /// Load a picture part as a data: URL (cached; counted against the media
    /// budget on every use because each use is serialized separately).
    fn image(&mut self, path: &str) -> Option<String> {
        if !self.media_cache.contains_key(path) {
            let res = self.load_image(path);
            self.media_cache.insert(path.to_string(), res);
        }
        let url = self.media_cache.get(path).cloned().flatten()?;
        if self.media_total + url.len() > MAX_MEDIA_TOTAL {
            self.warn(W_IMG_LIMIT);
            return None;
        }
        self.media_total += url.len();
        Some(url)
    }

    fn load_image(&mut self, path: &str) -> Option<String> {
        let bytes = match self.read(path, MAX_IMAGE) {
            Ok(b) => b,
            Err(e) => {
                if e.contains("too large") {
                    self.warn(W_IMG_LIMIT);
                }
                return None;
            }
        };
        let ext = path.rsplit('.').next().unwrap_or("");
        let mime = sniff_image(&bytes).unwrap_or_else(|| mime_for_ext(ext));
        match mime {
            "image/png" | "image/jpeg" | "image/gif" | "image/bmp" | "image/webp" | "image/svg+xml" | "image/x-icon" => {
                Some(data_url(mime, &bytes))
            }
            _ => {
                self.warn(W_IMG_FORMAT);
                None
            }
        }
    }
}

fn layout_kind(node: &Node) -> Option<String> {
    let t = node.attr("type").unwrap_or("");
    let k = match t {
        "title" => "title",
        "obj" | "tx" | "objTx" | "txAndObj" | "objAndTx" | "txAndChart" | "chartAndTx" | "txAndClipArt" | "clipArtAndTx"
        | "txAndMedia" | "mediaAndTx" | "tbl" | "chart" | "dgm" | "objOnly" | "picTx" | "vertTx" | "vertTitleAndTx" => "content",
        "secHead" => "section",
        "twoObj" | "twoTxTwoObj" | "twoColTx" | "objAndTwoObj" | "twoObjAndObj" | "twoObjAndTx" | "txAndTwoObj"
        | "twoObjOverTx" | "fourObj" => "two",
        "blank" => "blank",
        _ => {
            let name = node.child("cSld").and_then(|c| c.attr("name")).unwrap_or("").to_ascii_lowercase();
            if name.contains("title slide") {
                "title"
            } else if name.contains("section") {
                "section"
            } else if name.contains("two") || name.contains("comparison") {
                "two"
            } else if name.contains("blank") {
                "blank"
            } else if name.contains("content") {
                "content"
            } else {
                return None;
            }
        }
    };
    Some(k.to_string())
}

// ============================ placeholders ============================

#[derive(Clone, Debug)]
struct PhKey {
    typ: String,
    idx: Option<u32>,
}

fn nv_pr(el: &Node) -> Option<&Node> {
    el.elems().find(|c| c.local().starts_with("nv")).and_then(|nv| nv.child("nvPr"))
}

fn ph_of(el: &Node) -> Option<PhKey> {
    let ph = nv_pr(el)?.child("ph")?;
    Some(PhKey {
        typ: ph.attr("type").unwrap_or("obj").to_string(),
        idx: ph.attr("idx").and_then(|v| v.trim().parse().ok()),
    })
}

/// Placeholder category used for master matching / text styles.
fn ph_cat(t: &str) -> &str {
    match t {
        "title" | "ctrTitle" => "title",
        "dt" | "ftr" | "sldNum" | "hdr" => t,
        _ => "body",
    }
}

/// Contract `ph` value.
fn ph_model(t: &str) -> Option<&'static str> {
    match t {
        "title" | "ctrTitle" => Some("title"),
        "subTitle" => Some("subtitle"),
        "dt" | "ftr" | "sldNum" | "hdr" | "sldImg" => None,
        _ => Some("body"),
    }
}

fn sppr_of(n: Option<&Node>) -> Option<&Node> {
    n.and_then(|n| n.child("spPr"))
}

fn body_pr_of(n: Option<&Node>) -> Option<&Node> {
    n.and_then(|n| n.child("txBody")).and_then(|t| t.child("bodyPr"))
}

fn tree_of(part: &Node) -> Option<&Node> {
    part.child("cSld").and_then(|c| c.child("spTree"))
}

fn find_layout_ph<'n>(layout: &'n Node, key: &PhKey) -> Option<&'n Node> {
    let tree = tree_of(layout)?;
    let cands: Vec<(&Node, PhKey)> = tree.elems().filter_map(|e| ph_of(e).map(|k| (e, k))).collect();
    if let Some(i) = key.idx {
        if let Some((n, _)) = cands.iter().find(|(_, k)| k.idx == Some(i) && (i != 0 || ph_cat(&k.typ) == ph_cat(&key.typ))) {
            return Some(n);
        }
    }
    cands
        .iter()
        .find(|(_, k)| k.typ == key.typ)
        .or_else(|| cands.iter().find(|(_, k)| ph_cat(&k.typ) == ph_cat(&key.typ) && (key.idx.is_none() || ph_cat(&key.typ) == "title")))
        .map(|(n, _)| *n)
}

fn find_master_ph<'n>(master: &'n Node, key: &PhKey) -> Option<&'n Node> {
    let tree = tree_of(master)?;
    let cat = ph_cat(&key.typ);
    tree.elems().find(|e| ph_of(e).map(|k| ph_cat(&k.typ) == cat).unwrap_or(false))
}

// ============================ text properties ============================

#[derive(Clone, Default, Debug)]
struct TxProps {
    algn: Option<String>,
    /// Some(None) = explicitly no bullet.
    bullet: Option<Option<&'static str>>,
    sz: Option<f64>,
    b: Option<bool>,
    i: Option<bool>,
    u: Option<bool>,
    s: Option<bool>,
    baseline: Option<f64>,
    color: Option<[u8; 3]>,
    font: Option<String>,
    hl: Option<[u8; 3]>,
}

fn flag(n: &Node, a: &str) -> Option<bool> {
    n.attr(a).map(|v| matches!(v.trim(), "1" | "true" | "on"))
}

impl TxProps {
    fn overlay(&mut self, o: &TxProps) {
        macro_rules! take {
            ($($f:ident),*) => { $( if o.$f.is_some() { self.$f = o.$f.clone(); } )* };
        }
        take!(algn, bullet, sz, b, i, u, s, baseline, color, font, hl);
    }

    fn apply_ppr(&mut self, ppr: &Node, cc: &ColorCtx) {
        if let Some(a) = ppr.attr("algn") {
            self.algn = Some(a.to_string());
        }
        for c in ppr.elems() {
            match c.local() {
                "buNone" => self.bullet = Some(None),
                "buChar" | "buBlip" => self.bullet = Some(Some("ul")),
                "buAutoNum" => self.bullet = Some(Some("ol")),
                "defRPr" => self.apply_rpr(c, cc),
                _ => {}
            }
        }
    }

    fn apply_rpr(&mut self, rpr: &Node, cc: &ColorCtx) {
        if let Some(v) = num(rpr, "sz") {
            if v > 0.0 {
                self.sz = Some(v / 100.0);
            }
        }
        if let Some(v) = flag(rpr, "b") {
            self.b = Some(v);
        }
        if let Some(v) = flag(rpr, "i") {
            self.i = Some(v);
        }
        if let Some(v) = rpr.attr("u") {
            self.u = Some(v != "none");
        }
        if let Some(v) = rpr.attr("strike") {
            self.s = Some(v != "noStrike");
        }
        if let Some(v) = num(rpr, "baseline") {
            self.baseline = Some(v);
        }
        for c in rpr.elems() {
            match c.local() {
                "solidFill" => {
                    if let Some(col) = cc.color_in(c, None) {
                        self.color = Some(col);
                    }
                }
                "gradFill" => {
                    if let Some(Fill::Approx(col)) = fill_from_el(c, cc, None) {
                        self.color = Some(col);
                    }
                }
                "highlight" => self.hl = cc.color_in(c, None),
                "latin" => {
                    if let Some(f) = c.attr("typeface").map(str::trim).filter(|f| !f.is_empty()) {
                        self.font = Some(f.to_string());
                    }
                }
                _ => {}
            }
        }
    }
}

struct TextCtx<'n> {
    /// List-style containers (lstStyle / titleStyle / ...) from low to high priority.
    chain: Vec<&'n Node>,
    /// Shape-style / table-style text defaults, applied just before chain[base_at].
    base: TxProps,
    base_at: usize,
    font_scale: f64,
    /// Font that stays implicit (theme.headFont for titles, theme.font otherwise).
    default_font: String,
    theme_color: Option<[u8; 3]>,
    major: String,
    minor: String,
}

impl TextCtx<'_> {
    fn level_props(&self, lvl: u32, cc: &ColorCtx) -> TxProps {
        let mut p = TxProps::default();
        let name = format!("lvl{}pPr", lvl + 1);
        for (k, c) in self.chain.iter().enumerate() {
            if k == self.base_at {
                p.overlay(&self.base);
            }
            if let Some(l) = c.child(&name) {
                p.apply_ppr(l, cc);
            }
        }
        if self.base_at >= self.chain.len() {
            p.overlay(&self.base);
        }
        p
    }

    fn make_run(&self, t: String, p: &TxProps) -> Run {
        let font = p.font.as_deref().map(|f| match f {
            "+mj-lt" | "+mj-ea" | "+mj-cs" => self.major.clone(),
            "+mn-lt" | "+mn-ea" | "+mn-cs" => self.minor.clone(),
            f => f.to_string(),
        });
        Run {
            t,
            b: p.b.unwrap_or(false),
            i: p.i.unwrap_or(false),
            u: p.u.unwrap_or(false),
            s: p.s.unwrap_or(false),
            sup: p.baseline.map(|b| b > 0.0).unwrap_or(false),
            sub: p.baseline.map(|b| b < 0.0).unwrap_or(false),
            color: p.color.filter(|c| Some(*c) != self.theme_color).map(hex),
            size: p.sz.map(|s| ((s * self.font_scale) * 10.0).round() / 10.0),
            font: font.filter(|f| !f.eq_ignore_ascii_case(&self.default_font)),
            link: None,
            hl: p.hl.map(hex),
        }
    }
}

fn map_align(a: Option<&str>) -> Option<String> {
    match a? {
        "ctr" => Some("center".into()),
        "r" => Some("right".into()),
        "just" | "dist" | "justLow" | "thaiDist" => Some("justify".into()),
        _ => None,
    }
}

// ============================ geometry ============================

#[derive(Clone, Copy, Debug)]
struct GroupXf {
    off: (f64, f64),
    ext: (f64, f64),
    ch_off: (f64, f64),
    ch_ext: (f64, f64),
    rot: f64,
}

#[derive(Clone, Copy, Debug, Default)]
struct Rect {
    x: f64,
    y: f64,
    w: f64,
    h: f64,
    rot: f64,
}

fn xfrm_rect(x: &Node) -> Option<Rect> {
    let off = x.child("off");
    let ext = x.child("ext");
    if off.is_none() && ext.is_none() {
        return None;
    }
    Some(Rect {
        x: off.and_then(|o| num(o, "x")).unwrap_or(0.0),
        y: off.and_then(|o| num(o, "y")).unwrap_or(0.0),
        w: ext.and_then(|e| num(e, "cx")).unwrap_or(0.0).max(0.0),
        h: ext.and_then(|e| num(e, "cy")).unwrap_or(0.0).max(0.0),
        rot: num(x, "rot").unwrap_or(0.0) / 60000.0,
    })
}

fn group_xf(grp: &Node) -> GroupXf {
    let x = grp.child("grpSpPr").and_then(|s| s.child("xfrm"));
    let pair = |n: &str, a: &str, b: &str| {
        x.and_then(|x| x.child(n)).map(|e| (num(e, a).unwrap_or(0.0), num(e, b).unwrap_or(0.0)))
    };
    let off = pair("off", "x", "y").unwrap_or((0.0, 0.0));
    let ext = pair("ext", "cx", "cy").unwrap_or((0.0, 0.0));
    GroupXf {
        off,
        ext,
        ch_off: pair("chOff", "x", "y").unwrap_or(off),
        ch_ext: pair("chExt", "cx", "cy").unwrap_or(ext),
        rot: x.and_then(|x| num(x, "rot")).unwrap_or(0.0) / 60000.0,
    }
}

fn map_rect(xfs: &[GroupXf], mut r: Rect) -> Rect {
    for g in xfs.iter().rev() {
        let sx = if g.ch_ext.0 > 0.0 { g.ext.0 / g.ch_ext.0 } else { 1.0 };
        let sy = if g.ch_ext.1 > 0.0 { g.ext.1 / g.ch_ext.1 } else { 1.0 };
        r.x = g.off.0 + (r.x - g.ch_off.0) * sx;
        r.y = g.off.1 + (r.y - g.ch_off.1) * sy;
        r.w *= sx;
        r.h *= sy;
        if g.rot != 0.0 {
            let (gcx, gcy) = (g.off.0 + g.ext.0 / 2.0, g.off.1 + g.ext.1 / 2.0);
            let (cx, cy) = (r.x + r.w / 2.0, r.y + r.h / 2.0);
            let a = g.rot.to_radians();
            let (dx, dy) = (cx - gcx, cy - gcy);
            let (nx, ny) = (gcx + dx * a.cos() - dy * a.sin(), gcy + dx * a.sin() + dy * a.cos());
            r.x = nx - r.w / 2.0;
            r.y = ny - r.h / 2.0;
            r.rot += g.rot;
        }
    }
    r
}

fn norm_rot(r: f64) -> f64 {
    let v = r.rem_euclid(360.0);
    let v = (v * 100.0).round() / 100.0;
    if v >= 360.0 || v.abs() < 0.005 { 0.0 } else { v }
}

fn map_prst(p: &str) -> Option<&'static str> {
    Some(match p {
        "rect" | "flowChartProcess" | "snip1Rect" => "rect",
        "roundRect" | "flowChartAlternateProcess" | "round2SameRect" | "round1Rect" | "snipRoundRect" => "roundRect",
        "ellipse" | "flowChartConnector" | "circle" => "ellipse",
        "triangle" | "flowChartExtract" => "triangle",
        "diamond" | "flowChartDecision" => "diamond",
        "rightArrow" => "rightArrow",
        "leftArrow" => "leftArrow",
        "line" | "straightConnector1" | "bentConnector2" | "bentConnector3" | "bentConnector4" | "bentConnector5"
        | "curvedConnector2" | "curvedConnector3" | "curvedConnector4" | "curvedConnector5" | "lineInv" => "line",
        "star5" => "star5",
        "hexagon" | "flowChartPreparation" => "hexagon",
        "chevron" | "homePlate" => "chevron",
        _ => return None,
    })
}

fn placeholder_item(r: Rect, label: &str) -> Item {
    Item::Shape(ShapeItem {
        x: pt(r.x),
        y: pt(r.y),
        w: pt(r.w),
        h: pt(r.h),
        rot: norm_rot(r.rot),
        shape: "rect".into(),
        fill: Some("#F2F2F2".into()),
        line: Some("#BFBFBF".into()),
        line_w: Some(0.75),
        paras: vec![Para {
            runs: vec![Run { t: label.into(), color: Some("#595959".into()), size: Some(14.0), ..Default::default() }],
            align: Some("center".into()),
            ..Default::default()
        }],
        valign: Some("middle".into()),
    })
}

// ============================ slide walker ============================

struct Ctx<'n> {
    cc: ColorCtx,
    theme: Rc<ThemeInfo>,
    /// Deck-level text color / fonts (theme of the first master): run
    /// properties equal to these stay implicit.
    theme_color: Option<[u8; 3]>,
    deck_major: &'n str,
    deck_minor: &'n str,
    rels: &'n Rels,
    /// Layout + master of the slide (placeholder inheritance); None when
    /// walking a layout/master's own decorations.
    layout: Option<&'n LayoutInfo>,
    master: &'n MasterInfo,
    deftext: &'n Node,
    slide_part: bool,
}

impl<'a> Reader<'a> {
    fn walk(&mut self, tree: &Node, ctx: &Ctx, xfs: &mut Vec<GroupXf>, depth: usize, out: &mut Vec<Item>) {
        for el in tree.elems() {
            if out.len() >= MAX_ITEMS {
                self.warn(W_ITEMS);
                return;
            }
            match el.local() {
                "sp" => self.sp_item(el, ctx, xfs, out),
                "cxnSp" => self.sp_item(el, ctx, xfs, out),
                "pic" => self.pic_item(el, ctx, xfs, out),
                "graphicFrame" => self.frame_item(el, ctx, xfs, out),
                "grpSp" => {
                    if depth >= MAX_GROUP_DEPTH {
                        self.warn(W_DEPTH);
                        continue;
                    }
                    xfs.push(group_xf(el));
                    self.walk(el, ctx, xfs, depth + 1, out);
                    xfs.pop();
                }
                "AlternateContent" => {
                    if depth >= MAX_GROUP_DEPTH {
                        continue;
                    }
                    let chosen = el.child("Fallback").filter(|f| f.elems().next().is_some()).or_else(|| el.child("Choice"));
                    if let Some(c) = chosen {
                        self.walk(c, ctx, xfs, depth + 1, out);
                    }
                }
                "contentPart" => self.warn(W_OTHER_OBJ),
                _ => {}
            }
        }
    }

    /// Placeholder inheritance sources for a slide element.
    fn ph_sources<'n>(&self, key: Option<&PhKey>, ctx: &Ctx<'n>) -> (Option<&'n Node>, Option<&'n Node>) {
        let (Some(key), Some(layout)) = (key, ctx.layout) else { return (None, None) };
        let lay = find_layout_ph(&layout.node, key);
        let mkey = lay.and_then(ph_of).unwrap_or_else(|| key.clone());
        let mas = find_master_ph(&ctx.master.node, &mkey);
        (lay, mas)
    }

    fn sp_item(&mut self, sp: &Node, ctx: &Ctx, xfs: &[GroupXf], out: &mut Vec<Item>) {
        let key = ph_of(sp);
        if key.is_some() && !ctx.slide_part {
            return; // layout/master placeholders are prompts, not content
        }
        let (lay, mas) = self.ph_sources(key.as_ref(), ctx);
        let sppr = sp.child("spPr");
        let inh_sppr = sppr_of;
        let xfrm = sppr
            .and_then(|s| s.child("xfrm"))
            .and_then(xfrm_rect)
            .or_else(|| inh_sppr(lay).and_then(|s| s.child("xfrm")).and_then(xfrm_rect))
            .or_else(|| inh_sppr(mas).and_then(|s| s.child("xfrm")).and_then(xfrm_rect));
        let Some(r0) = xfrm else { return };
        let r = map_rect(xfs, r0);
        let cc = &ctx.cc;
        let style = sp.child("style");
        let style_color = |name: &str| style.and_then(|s| s.child(name)).and_then(|f| cc.color_in(f, None));
        let style_idx = |name: &str| style.and_then(|s| s.child(name)).and_then(|f| num(f, "idx")).unwrap_or(0.0) as usize;

        // ---- fill
        let mut fill = [sppr, inh_sppr(lay), inh_sppr(mas)].into_iter().flatten().find_map(|s| fill_of(s, cc, None));
        if fill.is_none() && style.is_some() {
            let idx = style_idx("fillRef");
            let phc = style_color("fillRef");
            fill = Some(if idx == 0 {
                Fill::None
            } else {
                let st = if idx >= 1001 { ctx.theme.bg_styles.get(idx - 1001) } else { ctx.theme.fill_styles.get(idx - 1) };
                match st.and_then(|s| fill_from_el(s, cc, phc)) {
                    Some(Fill::Approx(c)) => Fill::Solid(c), // theme style gradients: quiet approximation
                    Some(Fill::Picture(_)) | None => phc.map(Fill::Solid).unwrap_or(Fill::None),
                    Some(f) => f,
                }
            });
        }
        let mut fill_hex = None;
        let mut pic_fill = None;
        match fill {
            Some(Fill::Solid(c)) => fill_hex = Some(hex(c)),
            Some(Fill::Approx(c)) => {
                self.warn(W_GRADIENT);
                fill_hex = Some(hex(c));
            }
            Some(Fill::Picture(rid)) => pic_fill = Some(rid),
            _ => {}
        }

        // ---- line
        let mut line = [sppr, inh_sppr(lay), inh_sppr(mas)]
            .into_iter()
            .flatten()
            .find_map(|s| s.child("ln"))
            .map(|ln| line_from_ln(ln, cc, None))
            .unwrap_or_default();
        if !line.stated && style.is_some() {
            let idx = style_idx("lnRef");
            let phc = style_color("lnRef");
            if idx > 0 {
                let st = ctx.theme.ln_styles.get(idx - 1).map(|s| line_from_ln(s, cc, phc));
                match st {
                    Some(s) => {
                        line.color = if s.stated { s.color } else { phc };
                        if line.width.is_none() {
                            line.width = s.width;
                        }
                    }
                    None => line.color = phc,
                }
            }
        }
        let (line_hex, line_w) = match line.color {
            Some(c) => (Some(hex(c)), line.width),
            None => (None, None),
        };

        // ---- geometry
        let is_cxn = sp.is("cxnSp");
        let geom = sppr.and_then(|s| s.child("prstGeom")).and_then(|g| g.attr("prst"));
        let has_cust = sppr.and_then(|s| s.child("custGeom")).is_some();
        let shape_name = match geom {
            Some(g) => match map_prst(g) {
                Some(s) => s,
                None => {
                    self.warn(W_SHAPES);
                    "rect"
                }
            },
            None if is_cxn => "line",
            None => {
                if has_cust {
                    self.warn(W_SHAPES);
                }
                "rect"
            }
        };

        // ---- text
        let txb = sp.child("txBody");
        let is_title = key.as_ref().map(|k| ph_cat(&k.typ) == "title").unwrap_or(false);
        let body_pr = body_pr_of;
        let anchor = [body_pr(Some(sp)), body_pr(lay), body_pr(mas)].into_iter().flatten().find_map(|b| b.attr("anchor"));
        let font_scale = body_pr(Some(sp))
            .and_then(|b| b.child("normAutofit"))
            .and_then(|a| num(a, "fontScale"))
            .map(|f| (f / 100_000.0).clamp(0.05, 1.0))
            .unwrap_or(1.0);
        let mut chain: Vec<&Node> = Vec::new();
        let mut base = TxProps::default();
        if let Some(k) = &key {
            let styles = ctx.master.node.child("txStyles");
            let st = match ph_cat(&k.typ) {
                "title" => styles.and_then(|s| s.child("titleStyle")),
                "body" => styles.and_then(|s| s.child("bodyStyle")),
                _ => styles.and_then(|s| s.child("otherStyle")),
            };
            chain.extend(st);
            for n in [mas, lay].into_iter().flatten() {
                if let Some(l) = n.child("txBody").and_then(|t| t.child("lstStyle")) {
                    chain.push(l);
                }
            }
        } else {
            chain.push(ctx.deftext);
            if let Some(fr) = style.and_then(|s| s.child("fontRef")) {
                if let Some(c) = cc.color_in(fr, None) {
                    base.color = Some(c);
                }
                match fr.attr("idx") {
                    Some("major") => base.font = Some("+mj-lt".into()),
                    Some("minor") => base.font = Some("+mn-lt".into()),
                    _ => {}
                }
            }
        }
        if let Some(l) = txb.and_then(|t| t.child("lstStyle")) {
            chain.push(l);
        }
        let tctx = TextCtx {
            base_at: if key.is_some() { 0 } else { 1 },
            chain,
            base,
            font_scale,
            default_font: if is_title { ctx.deck_major.to_string() } else { ctx.deck_minor.to_string() },
            theme_color: ctx.theme_color,
            major: ctx.theme.major.clone(),
            minor: ctx.theme.minor.clone(),
        };
        let mut paras = match txb {
            Some(t) => self.paras(t, &tctx, ctx),
            None => Vec::new(),
        };
        let has_text = paras.iter().any(|p| p.runs.iter().any(|r| !r.t.trim().is_empty()));
        let valign = match anchor {
            Some("ctr") => Some("middle".to_string()),
            Some("b") => Some("bottom".to_string()),
            _ => None,
        };

        let model_ph = key.as_ref().and_then(|k| ph_model(&k.typ));
        if key.is_some() && model_ph.is_none() && !has_text {
            return; // empty date/footer/slide-number placeholder
        }
        let txbox = sp.child("nvSpPr").and_then(|n| n.child("cNvSpPr")).and_then(|c| c.attr("txBox")).map(|v| v == "1" || v == "true").unwrap_or(false);

        if let Some(rid) = pic_fill.clone().filter(|_| !has_text) {
            if let Some(src) = self.rel_image(ctx.rels, &rid) {
                let alt = sp.find("cNvPr").and_then(|c| c.attr("descr")).filter(|d| !d.trim().is_empty()).map(String::from);
                out.push(Item::Image(ImageItem { x: pt(r.x), y: pt(r.y), w: pt(r.w), h: pt(r.h), rot: norm_rot(r.rot), src, alt }));
                return;
            }
        } else if pic_fill.is_some() {
            self.warn(W_GRADIENT);
        }

        if model_ph.is_some() || txbox || (key.is_some() && has_text) {
            if paras.is_empty() {
                let p = tctx.level_props(0, cc);
                paras.push(Para {
                    runs: vec![tctx.make_run(String::new(), &p)],
                    align: map_align(p.algn.as_deref()),
                    level: 0,
                    bullet: p.bullet.flatten().map(String::from),
                });
            }
            out.push(Item::Text(TextItem {
                x: pt(r.x),
                y: pt(r.y),
                w: pt(r.w),
                h: pt(r.h),
                rot: norm_rot(r.rot),
                ph: model_ph.map(String::from),
                paras,
                fill: fill_hex,
                line: line_hex,
                line_w,
                valign,
            }));
        } else {
            let shape_line = shape_name == "line";
            out.push(Item::Shape(ShapeItem {
                x: pt(r.x),
                y: pt(r.y),
                w: pt(r.w),
                h: pt(r.h),
                rot: norm_rot(r.rot),
                shape: shape_name.to_string(),
                fill: if shape_line { None } else { fill_hex },
                line: line_hex,
                line_w,
                paras: if has_text { paras } else { Vec::new() },
                valign: Some(valign.unwrap_or_else(|| "top".into())),
            }));
        }
    }

    fn paras(&mut self, txb: &Node, tctx: &TextCtx, ctx: &Ctx) -> Vec<Para> {
        let cc = &ctx.cc;
        let mut out = Vec::new();
        for p in txb.elems().filter(|n| n.is("p")).take(MAX_PARAS) {
            let ppr = p.child("pPr");
            let lvl = ppr.and_then(|x| num(x, "lvl")).unwrap_or(0.0).clamp(0.0, 8.0) as u32;
            let mut props = tctx.level_props(lvl, cc);
            if let Some(x) = ppr {
                props.apply_ppr(x, cc);
            }
            let mut runs: Vec<Run> = Vec::new();
            for c in p.elems() {
                if runs.len() >= MAX_RUNS {
                    break;
                }
                match c.local() {
                    "r" | "fld" => {
                        let mut rp = props.clone();
                        let rpr = c.child("rPr");
                        if let Some(x) = rpr {
                            rp.apply_rpr(x, cc);
                        }
                        let t = c.child("t").map(|t| t.text()).unwrap_or_default();
                        let mut run = tctx.make_run(t.replace('\u{b}', "\n"), &rp);
                        if let Some(h) = rpr.and_then(|x| x.child("hlinkClick")) {
                            if let Some(rel) = rel_attr(h, "id").and_then(|id| ctx.rels.get(id)) {
                                if rel.external || rel.kind == "hyperlink" {
                                    run.link = safe_link(&rel.target);
                                }
                            }
                        }
                        runs.push(run);
                    }
                    "br" => {
                        let mut rp = props.clone();
                        if let Some(x) = c.child("rPr") {
                            rp.apply_rpr(x, cc);
                        }
                        runs.push(tctx.make_run("\n".into(), &rp));
                    }
                    _ => {}
                }
            }
            // Merge adjacent runs with identical formatting (keeps models small).
            let mut merged: Vec<Run> = Vec::with_capacity(runs.len());
            for r in runs {
                if let Some(last) = merged.last_mut() {
                    if same_fmt(last, &r) {
                        last.t.push_str(&r.t);
                        continue;
                    }
                }
                merged.push(r);
            }
            if merged.is_empty() {
                let mut rp = props.clone();
                if let Some(x) = p.child("endParaRPr") {
                    rp.apply_rpr(x, cc);
                }
                merged.push(tctx.make_run(String::new(), &rp));
            }
            out.push(Para { runs: merged, align: map_align(props.algn.as_deref()), level: lvl, bullet: props.bullet.flatten().map(String::from) });
        }
        out
    }

    fn rel_image(&mut self, rels: &Rels, rid: &str) -> Option<String> {
        let rel = rels.get(rid)?.clone();
        if rel.external {
            self.warn(W_IMG_LINKED);
            return None;
        }
        self.image(&rel.target)
    }

    fn pic_item(&mut self, pic: &Node, ctx: &Ctx, xfs: &[GroupXf], out: &mut Vec<Item>) {
        let key = ph_of(pic);
        if key.is_some() && !ctx.slide_part {
            return;
        }
        let (lay, mas) = self.ph_sources(key.as_ref(), ctx);
        let inh = |n: Option<&Node>| n.and_then(|n| n.child("spPr")).and_then(|s| s.child("xfrm")).and_then(xfrm_rect);
        let Some(r0) = inh(Some(pic)).or_else(|| inh(lay)).or_else(|| inh(mas)) else { return };
        let r = map_rect(xfs, r0);
        let nvpr = nv_pr(pic);
        let is_media = nvpr
            .map(|n| n.child("videoFile").is_some() || n.child("audioFile").is_some() || n.child("quickTimeFile").is_some() || n.find("media").is_some())
            .unwrap_or(false);
        if is_media {
            self.warn(W_MEDIA);
        }
        let blip = pic.child("blipFill").and_then(|b| b.child("blip"));
        let src = match blip {
            Some(b) => {
                if let Some(rid) = rel_attr(b, "embed") {
                    self.rel_image(ctx.rels, rid)
                } else {
                    if rel_attr(b, "link").is_some() {
                        self.warn(W_IMG_LINKED);
                    }
                    None
                }
            }
            None => None,
        };
        let alt = pic.find("cNvPr").and_then(|c| c.attr("descr")).filter(|d| !d.trim().is_empty()).map(String::from);
        match src {
            Some(src) => out.push(Item::Image(ImageItem { x: pt(r.x), y: pt(r.y), w: pt(r.w), h: pt(r.h), rot: norm_rot(r.rot), src, alt })),
            None if is_media => out.push(placeholder_item(r, "[Video or audio]")),
            None if blip.is_some() => out.push(placeholder_item(r, "[Picture]")),
            None => {}
        }
    }

    fn frame_item(&mut self, gf: &Node, ctx: &Ctx, xfs: &[GroupXf], out: &mut Vec<Item>) {
        let key = ph_of(gf);
        if key.is_some() && !ctx.slide_part {
            return;
        }
        let (lay, mas) = self.ph_sources(key.as_ref(), ctx);
        let inh = |n: Option<&Node>| n.and_then(|n| n.child("spPr")).and_then(|s| s.child("xfrm")).and_then(xfrm_rect);
        let own = gf.child("xfrm").and_then(xfrm_rect);
        let Some(r0) = own.or_else(|| inh(lay)).or_else(|| inh(mas)) else { return };
        let r = map_rect(xfs, r0);
        let Some(gd) = gf.child("graphic").and_then(|g| g.child("graphicData")) else { return };
        let uri = gd.attr("uri").unwrap_or("");
        if let Some(tbl) = gd.child("tbl") {
            let item = self.table(tbl, r, ctx);
            out.push(item);
        } else if uri.ends_with("/chart") || uri.contains("chartex") || gd.child("chart").is_some() {
            self.warn(W_CHARTS);
            out.push(placeholder_item(r, "[Chart]"));
        } else if uri.ends_with("/diagram") || gd.child("relIds").is_some() {
            self.warn(W_SMARTART);
            out.push(placeholder_item(r, "[SmartArt]"));
        } else if uri.ends_with("/ole") || gd.find("oleObj").is_some() {
            self.warn(W_OLE);
            // Preview picture: oleObj > p:pic (2010+) or an AlternateContent fallback.
            let mut pics = Vec::new();
            gd.find_all("pic", &mut pics);
            let mut tmp = Vec::new();
            for p in pics {
                if let Some(rid) = p.child("blipFill").and_then(|b| b.child("blip")).and_then(|b| rel_attr(b, "embed")) {
                    if let Some(src) = self.rel_image(ctx.rels, rid) {
                        tmp.push(Item::Image(ImageItem { x: pt(r.x), y: pt(r.y), w: pt(r.w), h: pt(r.h), rot: norm_rot(r.rot), src, alt: None }));
                        break;
                    }
                }
            }
            if tmp.is_empty() {
                out.push(placeholder_item(r, "[Embedded object]"));
            } else {
                out.extend(tmp);
            }
        } else {
            self.warn(W_OTHER_OBJ);
            out.push(placeholder_item(r, "[Object]"));
        }
    }

    fn table(&mut self, tbl: &Node, r: Rect, ctx: &Ctx) -> Item {
        let cc = &ctx.cc;
        let col_w: Vec<f64> = tbl
            .child("tblGrid")
            .map(|g| g.elems().filter(|c| c.is("gridCol")).take(1000).map(|c| pt(num(c, "w").unwrap_or(0.0))).collect())
            .unwrap_or_default();
        let tblpr = tbl.child("tblPr");
        // Approximate the default PowerPoint table style ("Medium Style 2 -
        // Accent 1"), which python-pptx and PowerPoint use for new tables.
        let style_id = tblpr.and_then(|p| p.child("tableStyleId")).map(|s| s.text()).unwrap_or_default();
        let medium2 = style_id.trim().eq_ignore_ascii_case("{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}");
        let first_row = tblpr.and_then(|p| flag(p, "firstRow")).unwrap_or(false);
        let band_row = tblpr.and_then(|p| flag(p, "bandRow")).unwrap_or(false);
        let accent = ctx.cc.scheme_color("accent1").unwrap_or([0x44, 0x72, 0xC4]);
        let tint = |c: [u8; 3], t: f64| c.map(|x| (x as f64 + (255.0 - x as f64) * (1.0 - t)).round() as u8);
        let tctx = TextCtx {
            chain: vec![ctx.deftext],
            base: TxProps::default(),
            base_at: 1,
            font_scale: 1.0,
            default_font: ctx.deck_minor.to_string(),
            theme_color: ctx.theme_color,
            major: ctx.theme.major.clone(),
            minor: ctx.theme.minor.clone(),
        };
        let mut rows = Vec::new();
        let mut cells = 0usize;
        for (ri, tr) in tbl.elems().filter(|e| e.is("tr")).enumerate() {
            let mut row = Vec::new();
            for tc in tr.elems().filter(|e| e.is("tc")) {
                cells += 1;
                if cells > MAX_TABLE_CELLS {
                    break;
                }
                if num(tc, "gridSpan").unwrap_or(1.0) > 1.0 || num(tc, "rowSpan").unwrap_or(1.0) > 1.0 {
                    self.warn(W_MERGED);
                }
                let merged_away = flag(tc, "hMerge").unwrap_or(false) || flag(tc, "vMerge").unwrap_or(false);
                let header = medium2 && first_row && ri == 0;
                let mut ctx2 = TextCtx { base: TxProps::default(), chain: tctx.chain.clone(), ..tctx.clone_shallow() };
                if header {
                    ctx2.base.b = Some(true);
                    ctx2.base.color = ctx.cc.scheme_color("lt1");
                }
                let mut paras = if merged_away {
                    Vec::new()
                } else {
                    tc.child("txBody").map(|t| self.paras(t, &ctx2, ctx)).unwrap_or_default()
                };
                if paras.is_empty() {
                    paras.push(Para { runs: vec![Run::default()], ..Default::default() });
                }
                let mut fill = tc.child("tcPr").and_then(|p| fill_of(p, cc, None)).and_then(|f| match f {
                    Fill::Solid(c) | Fill::Approx(c) => Some(hex(c)),
                    _ => None,
                });
                let stated = tc.child("tcPr").and_then(|p| fill_of(p, cc, None)).is_some();
                if !stated && medium2 {
                    let body_ri = if first_row { ri as i64 - 1 } else { ri as i64 };
                    fill = Some(hex(if header {
                        accent
                    } else if band_row && body_ri % 2 == 0 {
                        tint(accent, 0.4)
                    } else {
                        tint(accent, 0.2)
                    }));
                }
                row.push(SlideTableCell { paras, fill });
            }
            rows.push(row);
            if cells > MAX_TABLE_CELLS {
                self.warn(W_ITEMS);
                break;
            }
        }
        Item::Table(TableItem { x: pt(r.x), y: pt(r.y), w: pt(r.w), h: pt(r.h), rows, col_w })
    }
}

impl TextCtx<'_> {
    fn clone_shallow(&self) -> Self {
        TextCtx {
            chain: self.chain.clone(),
            base: self.base.clone(),
            base_at: self.base_at,
            font_scale: self.font_scale,
            default_font: self.default_font.clone(),
            theme_color: self.theme_color,
            major: self.major.clone(),
            minor: self.minor.clone(),
        }
    }
}

fn same_fmt(a: &Run, b: &Run) -> bool {
    a.b == b.b && a.i == b.i && a.u == b.u && a.s == b.s && a.sup == b.sup && a.sub == b.sub && a.color == b.color
        && a.size == b.size && a.font == b.font && a.link == b.link && a.hl == b.hl
}

// ============================ transitions / backgrounds ============================

fn map_transition(name: &str) -> Option<(&'static str, bool)> {
    Some(match name {
        "fade" => ("fade", true),
        "push" => ("push", true),
        "wipe" => ("wipe", true),
        "zoom" => ("zoom", true),
        "cut" => ("none", false),
        "dissolve" | "random" | "morph" | "flash" | "glitter" | "shred" | "fracture" | "crush" | "prestige"
        | "vortex" | "ripple" | "honeycomb" => ("fade", false),
        "cover" | "pull" | "pan" | "conveyor" | "ferris" | "gallery" | "reveal" | "switch" | "flip" | "cube" | "box"
        | "rotate" | "orbit" | "flythrough" => ("push", false),
        "split" | "strips" | "wedge" | "wheel" | "wheelReverse" | "circle" | "diamond" | "plus" | "blinds" | "checker"
        | "comb" | "randomBar" | "shape" | "curtains" | "wind" | "doors" | "window" | "pageCurlDouble"
        | "pageCurlSingle" | "peelOff" | "airplane" | "origami" => ("wipe", false),
        "newsflash" | "warp" | "prism" => ("zoom", false),
        "sndAc" | "extLst" => return None,
        _ => ("fade", false),
    })
}

fn transition_effect(tr: &Node) -> Option<(&'static str, bool)> {
    tr.elems().find_map(|c| map_transition(c.local()))
}

fn transition_of(sld: &Node) -> Option<(&'static str, bool)> {
    if let Some(t) = sld.child("transition") {
        return transition_effect(t);
    }
    for ac in sld.elems().filter(|e| e.is("AlternateContent")) {
        let choice = ac.child("Choice").and_then(|c| c.child("transition")).and_then(transition_effect);
        let fallback = ac.child("Fallback").and_then(|c| c.child("transition")).and_then(transition_effect);
        match (choice, fallback) {
            (Some(c), _) if c.1 => return Some(c),
            (_, Some(f)) => return Some(f),
            (Some(c), None) => return Some(c),
            _ => {}
        }
    }
    None
}

enum Bg {
    Color([u8; 3]),
    Picture(String),
}

impl<'a> Reader<'a> {
    fn bg_of(&mut self, part: &Node, rels: &Rels, cc: &ColorCtx, theme: &ThemeInfo) -> Option<Bg> {
        let bg = part.child("cSld")?.child("bg")?;
        if let Some(pr) = bg.child("bgPr") {
            return match fill_of(pr, cc, None)? {
                Fill::Solid(c) => Some(Bg::Color(c)),
                Fill::Approx(c) => {
                    self.warn(W_GRADIENT);
                    Some(Bg::Color(c))
                }
                Fill::Picture(rid) => self.rel_image(rels, &rid).map(Bg::Picture),
                Fill::None => None,
            };
        }
        if let Some(r) = bg.child("bgRef") {
            let phc = cc.color_in(r, None);
            let idx = num(r, "idx").unwrap_or(0.0) as usize;
            let st = if idx >= 1001 { theme.bg_styles.get(idx - 1001) } else if idx >= 1 { theme.fill_styles.get(idx - 1) } else { None };
            return match st.and_then(|s| fill_from_el(s, cc, phc)) {
                Some(Fill::Solid(c)) => Some(Bg::Color(c)),
                Some(Fill::Approx(c)) => {
                    self.warn(W_GRADIENT);
                    Some(Bg::Color(c))
                }
                _ => phc.map(Bg::Color),
            };
        }
        None
    }

    fn notes_text(&mut self, path: &str) -> String {
        let Ok(node) = self.xml(path) else { return String::new() };
        let Some(tree) = tree_of(&node) else { return String::new() };
        let mut lines = Vec::new();
        for sp in tree.elems().filter(|e| e.is("sp")) {
            let Some(k) = ph_of(sp) else { continue };
            if k.typ != "body" {
                continue;
            }
            if let Some(tb) = sp.child("txBody") {
                for p in tb.elems().filter(|e| e.is("p")).take(MAX_PARAS) {
                    let mut s = String::new();
                    for c in p.elems() {
                        match c.local() {
                            "r" | "fld" => s.push_str(&c.child("t").map(|t| t.text()).unwrap_or_default()),
                            "br" => s.push('\n'),
                            _ => {}
                        }
                    }
                    lines.push(s.replace('\u{b}', "\n"));
                }
            }
        }
        while lines.last().map(|l| l.trim().is_empty()).unwrap_or(false) {
            lines.pop();
        }
        lines.join("\n")
    }

    fn read_slide(&mut self, path: &str, theme_bg: Option<[u8; 3]>) -> Result<Slide, String> {
        let node = self.xml(path)?;
        if !node.is("sld") {
            return Err(format!("{path} is not a slide"));
        }
        let rels = self.rels(path);
        let layout = rels.first("slideLayout").map(|r| r.target.clone()).map(|p| self.layout(&p));
        let master = match &layout {
            Some(l) => l.master.clone(),
            None => {
                let mp = self.masters.keys().next().cloned().unwrap_or_default();
                self.master(&mp)
            }
        };
        let theme = master.theme.clone();
        let mut map = master.clr_map.clone();
        for part in [layout.as_ref().map(|l| &l.node), Some(&node)].into_iter().flatten() {
            if let Some(o) = part.child("clrMapOvr").and_then(|c| c.child("overrideClrMapping")) {
                map = clr_map_of(o);
            }
        }
        let cc = ColorCtx { scheme: theme.scheme.clone(), map };
        let theme_color = self.deck_color;
        let (deck_major, deck_minor) = (self.deck_major.clone(), self.deck_minor.clone());
        let mut slide = Slide { hidden: matches!(node.attr("show"), Some("0" | "false")), ..Default::default() };
        slide.layout = layout.as_ref().and_then(|l| l.kind.clone());

        let mut items: Vec<Item> = Vec::new();
        // Background: slide -> layout -> master.
        let mut bg = self.bg_of(&node, &rels, &cc, &theme).map(|b| (b, true));
        if bg.is_none() {
            if let Some(l) = &layout {
                bg = self.bg_of(&l.node, &l.rels, &cc, &theme).map(|b| (b, false));
            }
        }
        if bg.is_none() {
            bg = self.bg_of(&master.node, &master.rels, &cc, &theme).map(|b| (b, false));
        }
        match bg {
            Some((Bg::Color(c), own)) => {
                if own || Some(c) != theme_bg {
                    slide.bg = Some(hex(c));
                }
            }
            Some((Bg::Picture(src), _)) => items.push(Item::Image(ImageItem {
                x: 0.0,
                y: 0.0,
                w: pt(self.slide_w),
                h: pt(self.slide_h),
                rot: 0.0,
                src,
                alt: Some("Background picture".into()),
            })),
            None => {}
        }

        let deftext = self.deftext.clone();
        let show_master = |n: &Node| !matches!(n.attr("showMasterSp"), Some("0" | "false"));
        if show_master(&node) {
            let lay_ok = layout.as_ref().map(|l| show_master(&l.node)).unwrap_or(true);
            if lay_ok {
                if let Some(tree) = tree_of(&master.node) {
                    let ctx = Ctx { cc: cc.clone(), theme: theme.clone(), theme_color, deck_major: &deck_major, deck_minor: &deck_minor, rels: &master.rels, layout: None, master: &master, deftext: &deftext, slide_part: false };
                    self.walk(tree, &ctx, &mut Vec::new(), 0, &mut items);
                }
            }
            if let Some(l) = &layout {
                if let Some(tree) = tree_of(&l.node) {
                    let ctx = Ctx { cc: cc.clone(), theme: theme.clone(), theme_color, deck_major: &deck_major, deck_minor: &deck_minor, rels: &l.rels, layout: None, master: &master, deftext: &deftext, slide_part: false };
                    self.walk(tree, &ctx, &mut Vec::new(), 0, &mut items);
                }
            }
        }
        if let Some(tree) = tree_of(&node) {
            let ctx = Ctx { cc: cc.clone(), theme: theme.clone(), theme_color, deck_major: &deck_major, deck_minor: &deck_minor, rels: &rels, layout: layout.as_deref(), master: &master, deftext: &deftext, slide_part: true };
            self.walk(tree, &ctx, &mut Vec::new(), 0, &mut items);
        }
        slide.items = items;

        if let Some((t, exact)) = transition_of(&node) {
            if !exact {
                self.warn(W_TRANSITION);
            }
            slide.transition = Some(t.to_string());
        }
        if let Some(n) = rels.first("notesSlide").map(|r| r.target.clone()) {
            slide.notes = self.notes_text(&n);
        }
        if node.child("timing").map(|t| t.find("par").is_some()).unwrap_or(false) {
            self.warn(W_ANIM);
        }
        if rels.has("comments") || rels.has("comment") {
            self.warn(W_COMMENTS);
        }
        Ok(slide)
    }
}

// ============================ parse_pptx ============================

/// Read a .pptx package into the slides model. Fidelity notes go into
/// `doc.meta["warnings"]` (array of strings).
pub fn parse_pptx(bytes: &[u8]) -> Result<SlidesDoc, String> {
    if bytes.starts_with(&[0xD0, 0xCF, 0x11, 0xE0]) {
        return Err("This is not a .pptx file: it is either password-protected (encrypted) or an old PowerPoint 97-2003 (.ppt) presentation. Open it in PowerPoint and save it as an unprotected .pptx file.".into());
    }
    if !bytes.starts_with(b"PK") {
        return Err(bad("it is not a ZIP package"));
    }
    let zip = ZipArchive::new(Cursor::new(bytes)).map_err(|e| bad(&format!("the ZIP package is damaged ({e})")))?;
    let lower: HashMap<String, String> = zip.file_names().filter_map(|n| n.ok()).map(|n| (n.to_lowercase(), n.to_string())).collect();
    let mut rd = Reader {
        zip,
        lower,
        warnings: Vec::new(),
        media_cache: HashMap::new(),
        media_total: 0,
        themes: HashMap::new(),
        masters: HashMap::new(),
        layouts: HashMap::new(),
        deftext: Rc::new(Node::default()),
        slide_w: 9_144_000.0,
        slide_h: 6_858_000.0,
        deck_color: None,
        deck_major: "Calibri Light".into(),
        deck_minor: "Calibri".into(),
    };

    let root = rd.rels("");
    let main = root
        .0
        .values()
        .find(|r| r.kind == "officeDocument" && !r.external)
        .map(|r| r.target.clone())
        .filter(|p| rd.exists(p))
        .unwrap_or_else(|| "ppt/presentation.xml".into());
    if !rd.exists(&main) {
        if rd.exists("word/document.xml") {
            return Err(bad("it is a Word document, not a presentation"));
        }
        if rd.exists("xl/workbook.xml") {
            return Err(bad("it is an Excel workbook, not a presentation"));
        }
        return Err(bad("the presentation part (ppt/presentation.xml) is missing"));
    }
    let pres = rd.xml(&main).map_err(|e| bad(&format!("the presentation part is damaged ({e})")))?;
    if !pres.is("presentation") {
        return Err(bad("the main part is not a presentation"));
    }
    let prels = rd.rels(&main);

    let mut doc = SlidesDoc { kind: "slides".into(), ..Default::default() };
    if let Some(sz) = pres.child("sldSz") {
        let w = num(sz, "cx").unwrap_or(9_144_000.0);
        let h = num(sz, "cy").unwrap_or(6_858_000.0);
        if w > 0.0 && h > 0.0 && w < 1e12 && h < 1e12 {
            rd.slide_w = w;
            rd.slide_h = h;
        }
    }
    doc.size = SlideSize { w: pt(rd.slide_w), h: pt(rd.slide_h) };
    if let Some(d) = pres.child("defaultTextStyle") {
        rd.deftext = Rc::new(d.clone());
    }

    // Theme from the first master.
    let first_master = pres
        .child("sldMasterIdLst")
        .and_then(|l| l.elems().next())
        .and_then(|m| rel_attr(m, "id"))
        .and_then(|id| prels.get(id))
        .map(|r| r.target.clone())
        .or_else(|| prels.first("slideMaster").map(|r| r.target.clone()));
    let mut theme_bg = None;
    if let Some(mp) = first_master {
        let m = rd.master(&mp);
        let cc = ColorCtx { scheme: m.theme.scheme.clone(), map: m.clr_map.clone() };
        let tc = cc.scheme_color("tx1");
        let mbg = match rd.bg_of(&m.node, &m.rels, &cc, &m.theme) {
            Some(Bg::Color(c)) => Some(c),
            _ => cc.scheme_color("bg1"),
        };
        theme_bg = mbg;
        rd.deck_color = tc;
        rd.deck_major = m.theme.major.clone();
        rd.deck_minor = m.theme.minor.clone();
        doc.theme = Theme {
            font: Some(m.theme.minor.clone()),
            head_font: Some(m.theme.major.clone()),
            color: tc.map(hex),
            accent: cc.scheme_color("accent1").map(hex),
            bg: mbg.map(hex),
        };
    }

    // Slides in sldIdLst order.
    let mut paths: Vec<String> = Vec::new();
    let mut listed = 0usize;
    if let Some(lst) = pres.child("sldIdLst") {
        for s in lst.elems().filter(|e| e.is("sldId")) {
            listed += 1;
            if let Some(r) = rel_attr(s, "id").and_then(|id| prels.get(id)) {
                if !r.external && rd.exists(&r.target) {
                    paths.push(r.target.clone());
                }
            }
        }
    }
    if paths.is_empty() {
        let mut found: Vec<(u64, String)> = rd
            .lower
            .values()
            .filter_map(|n| {
                let l = n.to_lowercase();
                let num = l.strip_prefix("ppt/slides/slide")?.strip_suffix(".xml")?.parse::<u64>().ok()?;
                Some((num, n.clone()))
            })
            .collect();
        found.sort();
        if !found.is_empty() {
            rd.warn(W_SLIDE_LIST);
        } else if listed > 0 {
            return Err(bad("its slides are missing"));
        }
        paths = found.into_iter().map(|(_, n)| n).collect();
    }
    if paths.len() > MAX_SLIDES {
        rd.warn(&format!("This presentation has {} slides; only the first {MAX_SLIDES} were opened.", paths.len()));
        paths.truncate(MAX_SLIDES);
    }
    let mut failed = 0usize;
    for p in &paths {
        match rd.read_slide(p, theme_bg) {
            Ok(s) => doc.slides.push(s),
            Err(_) => failed += 1,
        }
    }
    if failed > 0 {
        if doc.slides.is_empty() {
            return Err(bad("none of its slides could be read"));
        }
        rd.warn(W_SLIDE_BROKEN);
    }

    // docProps/core.xml
    let core_path = root.0.values().find(|r| r.kind == "core-properties").map(|r| r.target.clone()).unwrap_or_else(|| "docProps/core.xml".into());
    if let Ok(core) = rd.xml(&core_path) {
        for (tag, key) in [("title", "title"), ("creator", "author"), ("subject", "subject")] {
            if let Some(v) = core.child(tag).map(|n| n.text()).map(|s| s.trim().to_string()).filter(|s| !s.is_empty()) {
                doc.meta.insert(key.into(), Value::String(v));
            }
        }
    }
    if !rd.warnings.is_empty() {
        doc.meta.insert("warnings".into(), Value::Array(rd.warnings.iter().cloned().map(Value::String).collect()));
    }
    Ok(doc)
}

// ============================ WRITE ============================

const NS_DECL: &str = r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main""#;
const XML_HEAD: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
const RT: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const CT_P: &str = "application/vnd.openxmlformats-officedocument.presentationml";
const TBL_NO_STYLE: &str = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}";

fn emu(v: f64) -> i64 {
    if !v.is_finite() {
        return 0;
    }
    (v * EMU_PER_PT).round().clamp(-27_273_042_316_900.0, 27_273_042_316_900.0) as i64
}

fn color_hex(s: Option<&String>) -> Option<String> {
    s.and_then(|c| norm_hex(c))
}

fn srgb_fill(h: &str) -> String {
    format!("<a:solidFill><a:srgbClr val=\"{h}\"/></a:solidFill>")
}

/// Percent-encode a hyperlink target so it is a valid URI (PowerPoint
/// repairs files whose relationship targets are not).
fn uri_escape(s: &str) -> String {
    let b = s.as_bytes();
    let mut o = String::with_capacity(s.len());
    let mut i = 0;
    while i < b.len() {
        let c = b[i];
        let ok_pct = c == b'%' && i + 2 < b.len() && b[i + 1].is_ascii_hexdigit() && b[i + 2].is_ascii_hexdigit();
        if ok_pct || (c.is_ascii_graphic() && !matches!(c, b'"' | b'<' | b'>' | b'\\' | b'^' | b'`' | b'{' | b'|' | b'}' | b'%')) {
            o.push(c as char);
        } else {
            o.push_str(&format!("%{c:02X}"));
        }
        i += 1;
    }
    o
}

fn rels_xml(rels: &[(String, String, String, bool)]) -> String {
    let mut s = format!("{XML_HEAD}<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
    for (id, typ, target, ext) in rels {
        s.push_str(&format!(
            "<Relationship Id=\"{id}\" Type=\"{typ}\" Target=\"{}\"{}/>",
            xml_attr(target),
            if *ext { " TargetMode=\"External\"" } else { "" }
        ));
    }
    s.push_str("</Relationships>");
    s
}

struct PartRels {
    list: Vec<(String, String, String, bool)>,
}

impl PartRels {
    fn new() -> Self {
        PartRels { list: Vec::new() }
    }
    fn add(&mut self, typ: &str, target: &str, external: bool) -> String {
        let full = format!("{RT}/{typ}");
        if let Some((id, ..)) = self.list.iter().find(|(_, t, tg, e)| *t == full && tg == target && *e == external) {
            return id.clone();
        }
        let id = format!("rId{}", self.list.len() + 1);
        self.list.push((id.clone(), full, target.to_string(), external));
        id
    }
}

struct Writer {
    notes: Vec<String>,
    media: Vec<(String, Vec<u8>)>,
    media_index: HashMap<(usize, u64), String>,
    exts: BTreeSet<&'static str>,
}

fn hash_bytes(b: &[u8]) -> u64 {
    let mut h: u64 = 0xcbf2_9ce4_8422_2325;
    for x in b {
        h ^= *x as u64;
        h = h.wrapping_mul(0x0100_0000_01b3);
    }
    h
}

impl Writer {
    fn note(&mut self, n: &str) {
        if !self.notes.iter().any(|x| x == n) {
            self.notes.push(n.to_string());
        }
    }

    /// Store a picture; returns the part name (ppt/media/imageN.ext).
    fn add_media(&mut self, src: &str) -> Option<String> {
        let Some((mime, bytes)) = parse_data_url(src) else {
            self.note("Some pictures could not be read (invalid image data) and were left out.");
            return None;
        };
        let real = sniff_image(&bytes).unwrap_or(mime.as_str()).to_string();
        let ext: &'static str = match real.as_str() {
            "image/png" => "png",
            "image/jpeg" | "image/jpg" => "jpeg",
            "image/gif" => "gif",
            "image/bmp" => "bmp",
            "image/tiff" => "tiff",
            "image/x-emf" => "emf",
            "image/x-wmf" => "wmf",
            "image/svg+xml" | "image/webp" => {
                self.note("Some pictures (SVG/WebP) cannot be stored in PowerPoint files and were left out; insert them as PNG or JPEG instead.");
                return None;
            }
            _ => {
                self.note("Some pictures could not be read (invalid image data) and were left out.");
                return None;
            }
        };
        if bytes.is_empty() || sniff_image(&bytes).is_none() {
            self.note("Some pictures could not be read (invalid image data) and were left out.");
            return None;
        }
        let key = (bytes.len(), hash_bytes(&bytes));
        if let Some(n) = self.media_index.get(&key) {
            return Some(n.clone());
        }
        let name = format!("ppt/media/image{}.{ext}", self.media.len() + 1);
        self.exts.insert(ext);
        self.media.push((name.clone(), bytes));
        self.media_index.insert(key, name.clone());
        Some(name)
    }
}

fn xfrm_xml(tag: &str, x: f64, y: f64, w: f64, h: f64, rot: f64) -> String {
    let (mut x, mut y, mut w, mut h) = (x, y, w, h);
    if w < 0.0 {
        x += w;
        w = -w;
    }
    if h < 0.0 {
        y += h;
        h = -h;
    }
    let r = if rot.is_finite() { (rot.rem_euclid(360.0) * 60000.0).round() as i64 % 21_600_000 } else { 0 };
    let rot_attr = if r != 0 { format!(" rot=\"{r}\"") } else { String::new() };
    format!(
        "<{tag}{rot_attr}><a:off x=\"{}\" y=\"{}\"/><a:ext cx=\"{}\" cy=\"{}\"/></{tag}>",
        emu(x),
        emu(y),
        emu(w).max(0),
        emu(h).max(0)
    )
}

fn clean_text(s: &str) -> String {
    s.chars().filter(|c| *c != '\r').collect()
}

fn rpr_xml(tag: &str, r: &Run, link_rid: Option<&str>) -> String {
    let mut a = String::from(" lang=\"en-US\"");
    if let Some(sz) = r.size.filter(|s| s.is_finite() && *s > 0.0) {
        a.push_str(&format!(" sz=\"{}\"", ((sz * 100.0).round() as i64).clamp(100, 400_000)));
    }
    if r.b {
        a.push_str(" b=\"1\"");
    }
    if r.i {
        a.push_str(" i=\"1\"");
    }
    if r.u {
        a.push_str(" u=\"sng\"");
    }
    if r.s {
        a.push_str(" strike=\"sngStrike\"");
    }
    if r.sup {
        a.push_str(" baseline=\"30000\"");
    } else if r.sub {
        a.push_str(" baseline=\"-25000\"");
    }
    a.push_str(" dirty=\"0\"");
    let mut c = String::new();
    if let Some(h) = color_hex(r.color.as_ref()) {
        c.push_str(&srgb_fill(&h));
    }
    if let Some(h) = color_hex(r.hl.as_ref()) {
        c.push_str(&format!("<a:highlight><a:srgbClr val=\"{h}\"/></a:highlight>"));
    }
    if let Some(f) = r.font.as_deref().map(str::trim).filter(|f| !f.is_empty()) {
        let f = xml_attr(f);
        c.push_str(&format!("<a:latin typeface=\"{f}\"/><a:ea typeface=\"{f}\"/><a:cs typeface=\"{f}\"/>"));
    }
    if let Some(id) = link_rid {
        c.push_str(&format!("<a:hlinkClick r:id=\"{id}\"/>"));
    }
    if c.is_empty() {
        format!("<a:{tag}{a}/>")
    } else {
        format!("<a:{tag}{a}>{c}</a:{tag}>")
    }
}

fn paras_xml(paras: &[Para], rels: &mut PartRels) -> String {
    let mut s = String::new();
    for p in paras {
        s.push_str(&para_xml(p, rels));
    }
    if s.is_empty() {
        s.push_str("<a:p><a:endParaRPr lang=\"en-US\" dirty=\"0\"/></a:p>");
    }
    s
}

fn para_xml(p: &Para, rels: &mut PartRels) -> String {
    let lvl = p.level.min(8);
    let algn = match p.align.as_deref() {
        Some("center") => "ctr",
        Some("right") => "r",
        Some("justify") => "just",
        _ => "l",
    };
    let (mar, ind, bu) = match p.bullet.as_deref() {
        Some("ul") => (
            285_750 + lvl as i64 * 457_200,
            -285_750,
            format!("<a:buFont typeface=\"Arial\"/><a:buChar char=\"{}\"/>", if lvl % 2 == 1 { "\u{2013}" } else { "\u{2022}" }),
        ),
        Some("ol") => (
            342_900 + lvl as i64 * 457_200,
            -342_900,
            format!("<a:buFont typeface=\"+mj-lt\"/><a:buAutoNum type=\"{}\"/>", ["arabicPeriod", "alphaLcPeriod", "romanLcPeriod"][lvl as usize % 3]),
        ),
        _ => (lvl as i64 * 457_200, 0, "<a:buNone/>".to_string()),
    };
    let mut s = format!("<a:p><a:pPr marL=\"{mar}\" lvl=\"{lvl}\" indent=\"{ind}\" algn=\"{algn}\">{bu}</a:pPr>");
    let mut last: Option<&Run> = None;
    for r in &p.runs {
        last = Some(r);
        let text = clean_text(&r.t).replace('\u{b}', "\n");
        if text.is_empty() {
            continue;
        }
        let rid = r.link.as_deref().and_then(safe_link).filter(|l| !l.starts_with('#')).map(|l| rels.add("hyperlink", &uri_escape(&l), true));
        for (k, piece) in text.split('\n').enumerate() {
            if k > 0 {
                s.push_str(&format!("<a:br>{}</a:br>", rpr_xml("rPr", r, None)));
            }
            if !piece.is_empty() {
                s.push_str(&format!("<a:r>{}<a:t>{}</a:t></a:r>", rpr_xml("rPr", r, rid.as_deref()), xml_text(piece)));
            }
        }
    }
    if let Some(r) = last {
        let end = Run { t: String::new(), link: None, ..r.clone() };
        s.push_str(&rpr_xml("endParaRPr", &end, None));
    }
    s.push_str("</a:p>");
    s
}

fn anchor_attr(v: Option<&str>) -> &'static str {
    match v {
        Some("middle") | Some("center") | Some("ctr") => "ctr",
        Some("bottom") | Some("b") => "b",
        _ => "t",
    }
}

#[derive(Clone, Copy, PartialEq)]
enum Lay {
    Title,
    Content,
    Section,
    Two,
    Blank,
}

impl Lay {
    fn of(s: Option<&str>, items: &[Item]) -> Lay {
        match s {
            Some("title") => Lay::Title,
            Some("content") => Lay::Content,
            Some("section") => Lay::Section,
            Some("two") => Lay::Two,
            Some("blank") => Lay::Blank,
            _ => {
                let has_ph = items.iter().any(|i| matches!(i, Item::Text(t) if t.ph.is_some()));
                if has_ph { Lay::Content } else { Lay::Blank }
            }
        }
    }
    fn file_no(self) -> usize {
        match self {
            Lay::Title => 1,
            Lay::Content => 2,
            Lay::Section => 3,
            Lay::Two => 4,
            Lay::Blank => 5,
        }
    }
}

struct SlideOut {
    xml: String,
    rels: PartRels,
}

fn write_slide(w: &mut Writer, slide: &Slide, lay: Lay) -> SlideOut {
    let mut rels = PartRels::new();
    rels.add("slideLayout", &format!("../slideLayouts/slideLayout{}.xml", lay.file_no()), false);
    let mut tree = String::new();
    let mut title_used = false;
    let mut ph_idx = 1u32;
    let mut unknown = false;
    for (cid, item) in (2u32..).zip(slide.items.iter().take(MAX_ITEMS)) {
        match item {
            Item::Text(t) => {
                let ph = match t.ph.as_deref() {
                    Some("title") if !title_used => {
                        title_used = true;
                        Some(if lay == Lay::Title { "<p:ph type=\"ctrTitle\"/>".to_string() } else { "<p:ph type=\"title\"/>".to_string() })
                    }
                    Some("subtitle") => {
                        let i = ph_idx;
                        ph_idx += 1;
                        Some(format!("<p:ph type=\"subTitle\" idx=\"{i}\"/>"))
                    }
                    Some("body") => {
                        let i = ph_idx;
                        ph_idx += 1;
                        Some(if matches!(lay, Lay::Content | Lay::Two) { format!("<p:ph idx=\"{i}\"/>") } else { format!("<p:ph type=\"body\" idx=\"{i}\"/>") })
                    }
                    _ => None,
                };
                let (name, nvsp, nvpr) = match &ph {
                    Some(p) => (format!("Placeholder {cid}"), "<p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr>".to_string(), format!("<p:nvPr>{p}</p:nvPr>")),
                    None => (format!("TextBox {cid}"), "<p:cNvSpPr txBox=\"1\"/>".to_string(), "<p:nvPr/>".to_string()),
                };
                let fill = match color_hex(t.fill.as_ref()) {
                    Some(h) => srgb_fill(&h),
                    None => "<a:noFill/>".into(),
                };
                let ln = ln_xml(t.line.as_ref(), t.line_w, ph.is_some());
                tree.push_str(&format!(
                    "<p:sp><p:nvSpPr><p:cNvPr id=\"{cid}\" name=\"{name}\"/>{nvsp}{nvpr}</p:nvSpPr><p:spPr>{}<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>{fill}{ln}</p:spPr><p:txBody><a:bodyPr vert=\"horz\" wrap=\"square\" lIns=\"91440\" tIns=\"45720\" rIns=\"91440\" bIns=\"45720\" rtlCol=\"0\" anchor=\"{}\"><a:noAutofit/></a:bodyPr><a:lstStyle/>{}</p:txBody></p:sp>",
                    xfrm_xml("a:xfrm", t.x, t.y, t.w, t.h, t.rot),
                    anchor_attr(t.valign.as_deref()),
                    paras_xml(&t.paras, &mut rels)
                ));
            }
            Item::Shape(sh) => {
                let prst = match sh.shape.as_str() {
                    s @ ("rect" | "roundRect" | "ellipse" | "triangle" | "diamond" | "rightArrow" | "leftArrow" | "line" | "star5" | "hexagon" | "chevron") => s,
                    _ => {
                        w.note("Some shapes of an unknown kind were saved as rectangles.");
                        "rect"
                    }
                };
                let is_line = prst == "line";
                let fill = match color_hex(sh.fill.as_ref()) {
                    Some(h) if !is_line => srgb_fill(&h),
                    _ => "<a:noFill/>".into(),
                };
                let ln = ln_xml(sh.line.as_ref(), sh.line_w, true);
                let tx = if sh.paras.is_empty() {
                    format!("<p:txBody><a:bodyPr rtlCol=\"0\" anchor=\"{}\"/><a:lstStyle/><a:p><a:pPr algn=\"ctr\"/><a:endParaRPr lang=\"en-US\" dirty=\"0\"/></a:p></p:txBody>", anchor_attr(sh.valign.as_deref()))
                } else {
                    format!(
                        "<p:txBody><a:bodyPr wrap=\"square\" rtlCol=\"0\" anchor=\"{}\"><a:noAutofit/></a:bodyPr><a:lstStyle/>{}</p:txBody>",
                        anchor_attr(sh.valign.as_deref()),
                        paras_xml(&sh.paras, &mut rels)
                    )
                };
                let (open, nvsp, close) = if is_line {
                    ("<p:cxnSp><p:nvCxnSpPr>", "<p:cNvCxnSpPr/>", "</p:nvCxnSpPr>")
                } else {
                    ("<p:sp><p:nvSpPr>", "<p:cNvSpPr/>", "</p:nvSpPr>")
                };
                tree.push_str(&format!(
                    "{open}<p:cNvPr id=\"{cid}\" name=\"Shape {cid}\"/>{nvsp}<p:nvPr/>{close}<p:spPr>{}<a:prstGeom prst=\"{prst}\"><a:avLst/></a:prstGeom>{fill}{ln}</p:spPr>{}{}",
                    xfrm_xml("a:xfrm", sh.x, sh.y, sh.w, sh.h, sh.rot),
                    if is_line { "" } else { &tx },
                    if is_line { "</p:cxnSp>" } else { "</p:sp>" }
                ));
            }
            Item::Image(im) => {
                let Some(part) = w.add_media(&im.src) else { continue };
                let rid = rels.add("image", &format!("../media/{}", part.rsplit('/').next().unwrap_or("")), false);
                let descr = im.alt.as_deref().map(|a| format!(" descr=\"{}\"", xml_attr(a))).unwrap_or_default();
                tree.push_str(&format!(
                    "<p:pic><p:nvPicPr><p:cNvPr id=\"{cid}\" name=\"Picture {cid}\"{descr}/><p:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed=\"{rid}\"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>{}<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr></p:pic>",
                    xfrm_xml("a:xfrm", im.x, im.y, im.w, im.h, im.rot)
                ));
            }
            Item::Table(tb) => {
                if let Some(x) = table_xml(tb, cid, &mut rels) {
                    tree.push_str(&x);
                } else {
                    w.note("Empty tables were left out.");
                }
            }
            Item::Other(_) => unknown = true,
        }
    }
    if slide.items.len() > MAX_ITEMS {
        w.note("Some slides had too many objects; the extra ones were left out.");
    }
    if unknown {
        w.note("Some slide objects of an unknown type were not saved.");
    }
    let bg = color_hex(slide.bg.as_ref())
        .map(|h| format!("<p:bg><p:bgPr>{}<a:effectLst/></p:bgPr></p:bg>", srgb_fill(&h)))
        .unwrap_or_default();
    let transition = match slide.transition.as_deref() {
        Some("fade") => "<p:transition spd=\"med\"><p:fade/></p:transition>",
        Some("push") => "<p:transition spd=\"med\"><p:push dir=\"u\"/></p:transition>",
        Some("wipe") => "<p:transition spd=\"med\"><p:wipe dir=\"r\"/></p:transition>",
        Some("zoom") => "<p:transition spd=\"med\"><p:zoom/></p:transition>",
        _ => "",
    };
    let show = if slide.hidden { " show=\"0\"" } else { "" };
    let xml = format!(
        "{XML_HEAD}<p:sld {NS_DECL}{show}><p:cSld>{bg}<p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr>{tree}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>{transition}</p:sld>"
    );
    SlideOut { xml, rels }
}

fn ln_xml(line: Option<&String>, w: Option<f64>, explicit_none: bool) -> String {
    match color_hex(line) {
        Some(h) => {
            let wa = w.filter(|w| w.is_finite() && *w > 0.0).map(|w| format!(" w=\"{}\"", emu(w).clamp(0, 20_116_800))).unwrap_or_default();
            format!("<a:ln{wa}>{}</a:ln>", srgb_fill(&h))
        }
        None if explicit_none => "<a:ln><a:noFill/></a:ln>".into(),
        None => String::new(),
    }
}

fn table_xml(tb: &TableItem, cid: u32, rels: &mut PartRels) -> Option<String> {
    let rows: Vec<&Vec<SlideTableCell>> = tb.rows.iter().take(1000).collect();
    let ncols = rows.iter().map(|r| r.len()).max().unwrap_or(0).min(1000);
    if rows.is_empty() || ncols == 0 {
        return None;
    }
    let total_w = if tb.w > 0.0 { tb.w } else { 72.0 * ncols as f64 };
    let given: f64 = tb.col_w.iter().take(ncols).filter(|v| v.is_finite() && **v > 0.0).sum();
    let n_missing = (0..ncols).filter(|i| !tb.col_w.get(*i).map(|v| v.is_finite() && *v > 0.0).unwrap_or(false)).count();
    let rest = if n_missing > 0 { ((total_w - given).max(0.0) / n_missing as f64).max(18.0) } else { 0.0 };
    let widths: Vec<f64> = (0..ncols).map(|i| tb.col_w.get(i).copied().filter(|v| v.is_finite() && *v > 0.0).unwrap_or(rest)).collect();
    let sum_w: f64 = widths.iter().sum();
    let row_h = if tb.h > 0.0 { tb.h / rows.len() as f64 } else { 28.0 };
    let border = "<a:solidFill><a:srgbClr val=\"BFBFBF\"/></a:solidFill>";
    let mut s = format!(
        "<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id=\"{cid}\" name=\"Table {cid}\"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp=\"1\"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>{}<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/table\"><a:tbl><a:tblPr firstRow=\"1\" bandRow=\"1\"><a:tableStyleId>{TBL_NO_STYLE}</a:tableStyleId></a:tblPr><a:tblGrid>",
        xfrm_xml("p:xfrm", tb.x, tb.y, sum_w, row_h * rows.len() as f64, 0.0)
    );
    for w in &widths {
        s.push_str(&format!("<a:gridCol w=\"{}\"/>", emu(*w).max(0)));
    }
    s.push_str("</a:tblGrid>");
    for row in rows {
        s.push_str(&format!("<a:tr h=\"{}\">", emu(row_h).max(0)));
        for c in 0..ncols {
            let cell = row.get(c);
            let paras = cell.map(|c| paras_xml(&c.paras, rels)).unwrap_or_else(|| paras_xml(&[], rels));
            let fill = match cell.and_then(|c| color_hex(c.fill.as_ref())) {
                Some(h) => srgb_fill(&h),
                None => "<a:noFill/>".into(),
            };
            s.push_str(&format!(
                "<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>{paras}</a:txBody><a:tcPr><a:lnL w=\"12700\">{border}</a:lnL><a:lnR w=\"12700\">{border}</a:lnR><a:lnT w=\"12700\">{border}</a:lnT><a:lnB w=\"12700\">{border}</a:lnB>{fill}</a:tcPr></a:tc>"
            ));
        }
        s.push_str("</a:tr>");
    }
    s.push_str("</a:tbl></a:graphicData></a:graphic></p:graphicFrame>");
    Some(s)
}

fn notes_slide_xml(notes: &str) -> String {
    let paras: Vec<Para> = notes
        .split('\n')
        .map(|l| Para { runs: vec![Run::text(l)], ..Default::default() })
        .collect();
    let body: String = paras
        .iter()
        .map(|p| {
            let t = clean_text(&p.runs[0].t);
            if t.is_empty() {
                "<a:p><a:endParaRPr lang=\"en-US\" dirty=\"0\"/></a:p>".to_string()
            } else {
                format!("<a:p><a:r><a:rPr lang=\"en-US\" dirty=\"0\"/><a:t>{}</a:t></a:r></a:p>", xml_text(&t))
            }
        })
        .collect();
    format!(
        "{XML_HEAD}<p:notes {NS_DECL}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id=\"2\" name=\"Slide Image Placeholder 1\"/><p:cNvSpPr><a:spLocks noGrp=\"1\" noRot=\"1\" noChangeAspect=\"1\"/></p:cNvSpPr><p:nvPr><p:ph type=\"sldImg\"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp><p:sp><p:nvSpPr><p:cNvPr id=\"3\" name=\"Notes Placeholder 2\"/><p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr><p:ph type=\"body\" idx=\"1\"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>{body}</p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>"
    )
}

fn lvl_styles(sizes: &[u32], font: &str, bullets: bool, algn: &str) -> String {
    let mut s = String::new();
    for i in 0..9usize {
        let sz = sizes.get(i).or(sizes.last()).copied().unwrap_or(1800);
        let (mar, ind, bu) = if bullets {
            (285_750 + i as i64 * 457_200, -285_750, format!("<a:buFont typeface=\"Arial\"/><a:buChar char=\"{}\"/>", if i % 2 == 1 { "\u{2013}" } else { "\u{2022}" }))
        } else {
            (i as i64 * 457_200, 0, "<a:buNone/>".to_string())
        };
        s.push_str(&format!(
            "<a:lvl{n}pPr marL=\"{mar}\" indent=\"{ind}\" algn=\"{algn}\" defTabSz=\"914400\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\"><a:spcBef><a:spcPts val=\"{}\"/></a:spcBef>{bu}<a:defRPr sz=\"{sz}\" kern=\"1200\"><a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill><a:latin typeface=\"{font}\"/><a:ea typeface=\"{font}\"/><a:cs typeface=\"{font}\"/></a:defRPr></a:lvl{n}pPr>",
            if bullets { 1000 } else { 0 },
            n = i + 1
        ));
    }
    s
}

fn clr_map() -> &'static str {
    "<p:clrMap bg1=\"lt1\" tx1=\"dk1\" bg2=\"lt2\" tx2=\"dk2\" accent1=\"accent1\" accent2=\"accent2\" accent3=\"accent3\" accent4=\"accent4\" accent5=\"accent5\" accent6=\"accent6\" hlink=\"hlink\" folHlink=\"folHlink\"/>"
}

const GRP_ROOT: &str = "<p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr>";

fn ph_sp(id: u32, name: &str, ph: &str, r: (f64, f64, f64, f64), anchor: Option<&str>, prompt: &str) -> String {
    let body_pr = match anchor {
        Some(a) => format!("<a:bodyPr anchor=\"{a}\"/>"),
        None => "<a:bodyPr/>".to_string(),
    };
    format!(
        "<p:sp><p:nvSpPr><p:cNvPr id=\"{id}\" name=\"{name}\"/><p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr>{ph}</p:nvPr></p:nvSpPr><p:spPr>{}</p:spPr><p:txBody>{body_pr}<a:lstStyle/><a:p><a:r><a:rPr lang=\"en-US\"/><a:t>{prompt}</a:t></a:r></a:p></p:txBody></p:sp>",
        xfrm_xml("a:xfrm", r.0, r.1, r.2, r.3, 0.0)
    )
}

fn master_xml(w: f64, h: f64) -> String {
    let title = ph_sp(2, "Title Placeholder 1", "<p:ph type=\"title\"/>", (0.0625 * w, 0.05 * h, 0.875 * w, 0.18 * h), Some("ctr"), "Click to edit Master title style");
    let body = ph_sp(3, "Text Placeholder 2", "<p:ph type=\"body\" idx=\"1\"/>", (0.0625 * w, 0.26 * h, 0.875 * w, 0.64 * h), None, "Click to edit Master text styles");
    let mut ids = String::new();
    for i in 0..5u32 {
        ids.push_str(&format!("<p:sldLayoutId id=\"{}\" r:id=\"rId{}\"/>", 2_147_483_649u64 + i as u64, i + 1));
    }
    format!(
        "{XML_HEAD}<p:sldMaster {NS_DECL}><p:cSld><p:bg><p:bgRef idx=\"1001\"><a:schemeClr val=\"bg1\"/></p:bgRef></p:bg><p:spTree>{GRP_ROOT}{title}{body}</p:spTree></p:cSld>{}<p:sldLayoutIdLst>{ids}</p:sldLayoutIdLst><p:txStyles><p:titleStyle>{}</p:titleStyle><p:bodyStyle>{}</p:bodyStyle><p:otherStyle><a:defPPr><a:defRPr lang=\"en-US\"/></a:defPPr>{}</p:otherStyle></p:txStyles></p:sldMaster>",
        clr_map(),
        lvl_styles(&[4400], "+mj-lt", false, "l"),
        lvl_styles(&[2800, 2400, 2000, 1800], "+mn-lt", true, "l"),
        lvl_styles(&[1800], "+mn-lt", false, "l"),
    )
}

fn layout_xml(lay: Lay, w: f64, h: f64) -> String {
    let (typ, name, phs): (&str, &str, Vec<String>) = match lay {
        Lay::Title => (
            "title",
            "Title Slide",
            vec![
                ph_sp(2, "Title 1", "<p:ph type=\"ctrTitle\"/>", (0.083 * w, 0.30 * h, 0.834 * w, 0.22 * h), Some("b"), "Click to edit Master title style"),
                ph_sp(3, "Subtitle 2", "<p:ph type=\"subTitle\" idx=\"1\"/>", (0.125 * w, 0.555 * h, 0.75 * w, 0.15 * h), None, "Click to edit Master subtitle style"),
            ],
        ),
        Lay::Content => (
            "obj",
            "Title and Content",
            vec![
                ph_sp(2, "Title 1", "<p:ph type=\"title\"/>", (0.0625 * w, 0.05 * h, 0.875 * w, 0.18 * h), None, "Click to edit Master title style"),
                ph_sp(3, "Content Placeholder 2", "<p:ph idx=\"1\"/>", (0.0625 * w, 0.26 * h, 0.875 * w, 0.64 * h), None, "Click to edit Master text styles"),
            ],
        ),
        Lay::Section => (
            "secHead",
            "Section Header",
            vec![
                ph_sp(2, "Title 1", "<p:ph type=\"title\"/>", (0.0625 * w, 0.3 * h, 0.875 * w, 0.25 * h), Some("b"), "Click to edit Master title style"),
                ph_sp(3, "Text Placeholder 2", "<p:ph type=\"body\" idx=\"1\"/>", (0.0625 * w, 0.57 * h, 0.875 * w, 0.15 * h), None, "Click to edit Master text styles"),
            ],
        ),
        Lay::Two => (
            "twoObj",
            "Two Content",
            vec![
                ph_sp(2, "Title 1", "<p:ph type=\"title\"/>", (0.0625 * w, 0.05 * h, 0.875 * w, 0.18 * h), None, "Click to edit Master title style"),
                ph_sp(3, "Content Placeholder 2", "<p:ph sz=\"half\" idx=\"1\"/>", (0.0625 * w, 0.26 * h, 0.43 * w, 0.64 * h), None, "Click to edit Master text styles"),
                ph_sp(4, "Content Placeholder 3", "<p:ph sz=\"half\" idx=\"2\"/>", (0.5075 * w, 0.26 * h, 0.43 * w, 0.64 * h), None, "Click to edit Master text styles"),
            ],
        ),
        Lay::Blank => ("blank", "Blank", Vec::new()),
    };
    format!(
        "{XML_HEAD}<p:sldLayout {NS_DECL} type=\"{typ}\" preserve=\"1\"><p:cSld name=\"{name}\"><p:spTree>{GRP_ROOT}{}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>",
        phs.concat()
    )
}

fn theme_xml(t: &Theme) -> String {
    let c = |v: Option<&String>, d: &str| color_hex(v).unwrap_or_else(|| d.to_string());
    let font = |v: Option<&String>, d: &str| v.map(|s| s.trim()).filter(|s| !s.is_empty()).map(xml_attr).unwrap_or_else(|| d.to_string());
    let sys = |name: &str, _: &str, val: String| format!("<a:{name}><a:srgbClr val=\"{val}\"/></a:{name}>");
    let scheme = [
        sys("dk1", "", c(t.color.as_ref(), "000000")),
        sys("lt1", "", c(t.bg.as_ref(), "FFFFFF")),
        sys("dk2", "", "44546A".into()),
        sys("lt2", "", "E7E6E6".into()),
        sys("accent1", "", c(t.accent.as_ref(), "4472C4")),
        sys("accent2", "", "ED7D31".into()),
        sys("accent3", "", "A5A5A5".into()),
        sys("accent4", "", "FFC000".into()),
        sys("accent5", "", "5B9BD5".into()),
        sys("accent6", "", "70AD47".into()),
        sys("hlink", "", "0563C1".into()),
        sys("folHlink", "", "954F72".into()),
    ]
    .concat();
    let major = font(t.head_font.as_ref(), "Calibri Light");
    let minor = font(t.font.as_ref(), "Calibri");
    let solid = "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>";
    let ln = |w: u32| format!("<a:ln w=\"{w}\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/><a:miter lim=\"800000\"/></a:ln>");
    format!(
        "{XML_HEAD}<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Office Theme\"><a:themeElements><a:clrScheme name=\"Office\">{scheme}</a:clrScheme><a:fontScheme name=\"Office\"><a:majorFont><a:latin typeface=\"{major}\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:majorFont><a:minorFont><a:latin typeface=\"{minor}\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:minorFont></a:fontScheme><a:fmtScheme name=\"Office\"><a:fillStyleLst>{solid}<a:solidFill><a:schemeClr val=\"phClr\"><a:tint val=\"50000\"/></a:schemeClr></a:solidFill><a:solidFill><a:schemeClr val=\"phClr\"><a:shade val=\"80000\"/></a:schemeClr></a:solidFill></a:fillStyleLst><a:lnStyleLst>{}{}{}</a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst>{solid}<a:solidFill><a:schemeClr val=\"phClr\"><a:tint val=\"95000\"/></a:schemeClr></a:solidFill><a:solidFill><a:schemeClr val=\"phClr\"><a:shade val=\"90000\"/></a:schemeClr></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>",
        ln(6350),
        ln(12700),
        ln(19050)
    )
}

fn notes_master_xml() -> String {
    format!(
        "{XML_HEAD}<p:notesMaster {NS_DECL}><p:cSld><p:bg><p:bgRef idx=\"1001\"><a:schemeClr val=\"bg1\"/></p:bgRef></p:bg><p:spTree>{GRP_ROOT}<p:sp><p:nvSpPr><p:cNvPr id=\"2\" name=\"Slide Image Placeholder 1\"/><p:cNvSpPr><a:spLocks noGrp=\"1\" noRot=\"1\" noChangeAspect=\"1\"/></p:cNvSpPr><p:nvPr><p:ph type=\"sldImg\" idx=\"2\"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"381000\" y=\"685800\"/><a:ext cx=\"6096000\" cy=\"3429000\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:noFill/><a:ln w=\"12700\"><a:solidFill><a:prstClr val=\"black\"/></a:solidFill></a:ln></p:spPr></p:sp><p:sp><p:nvSpPr><p:cNvPr id=\"3\" name=\"Notes Placeholder 2\"/><p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr><p:ph type=\"body\" sz=\"quarter\" idx=\"3\"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"685800\" y=\"4400550\"/><a:ext cx=\"5486400\" cy=\"3600450\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert=\"horz\" lIns=\"91440\" tIns=\"45720\" rIns=\"91440\" bIns=\"45720\" rtlCol=\"0\"/><a:lstStyle/><a:p><a:pPr lvl=\"0\"/><a:r><a:rPr lang=\"en-US\"/><a:t>Click to edit Master text styles</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld>{}<p:notesStyle>{}</p:notesStyle></p:notesMaster>",
        clr_map(),
        lvl_styles(&[1200], "+mn-lt", false, "l")
    )
}

fn now_w3cdtf() -> String {
    let secs = std::time::SystemTime::now().duration_since(std::time::UNIX_EPOCH).map(|d| d.as_secs() as i64).unwrap_or(0);
    let z = secs.div_euclid(86_400) + 719_468;
    let era = z.div_euclid(146_097);
    let doe = z - era * 146_097;
    let yoe = (doe - doe / 1460 + doe / 36_524 - doe / 146_096) / 365;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = doy - (153 * mp + 2) / 5 + 1;
    let m = if mp < 10 { mp + 3 } else { mp - 9 };
    let y = yoe + era * 400 + if m <= 2 { 1 } else { 0 };
    let r = secs.rem_euclid(86_400);
    format!("{y:04}-{m:02}-{d:02}T{:02}:{:02}:{:02}Z", r / 3600, r % 3600 / 60, r % 60)
}

/// Write the slides model as a .pptx package. Returns bytes + notes about
/// anything that could not be written exactly.
pub fn write_pptx(doc: &SlidesDoc) -> Result<(Vec<u8>, Vec<String>), String> {
    let mut w = Writer { notes: Vec::new(), media: Vec::new(), media_index: HashMap::new(), exts: BTreeSet::new() };
    let sw = if doc.size.w.is_finite() && doc.size.w > 0.0 { doc.size.w } else { 960.0 };
    let sh = if doc.size.h.is_finite() && doc.size.h > 0.0 { doc.size.h } else { 540.0 };
    let cx = emu(sw).clamp(914_400, 51_206_400);
    let cy = emu(sh).clamp(914_400, 51_206_400);
    if cx != emu(sw) || cy != emu(sh) {
        w.note("The slide size was adjusted to the range PowerPoint supports (1 to 56 inches).");
    }
    let (wpt, hpt) = (cx as f64 / EMU_PER_PT, cy as f64 / EMU_PER_PT);

    let mut slides: Vec<&Slide> = doc.slides.iter().collect();
    if slides.len() > MAX_SLIDES {
        w.note(&format!("Only the first {MAX_SLIDES} slides were saved."));
        slides.truncate(MAX_SLIDES);
    }
    let blank;
    if slides.is_empty() {
        blank = Slide { layout: Some("blank".into()), ..Default::default() };
        slides.push(&blank);
    }
    let any_notes = slides.iter().any(|s| !s.notes.trim().is_empty());

    let mut parts: Vec<(String, Vec<u8>)> = Vec::new();
    let mut overrides: Vec<(String, String)> = Vec::new();
    let put = |parts: &mut Vec<(String, Vec<u8>)>, overrides: &mut Vec<(String, String)>, name: String, ct: Option<String>, data: String| {
        if let Some(ct) = ct {
            overrides.push((format!("/{name}"), ct));
        }
        parts.push((name, data.into_bytes()));
    };

    // presentation rels
    let mut prels = PartRels::new();
    prels.add("slideMaster", "slideMasters/slideMaster1.xml", false);
    let mut sld_ids = String::new();
    for (i, slide) in slides.iter().enumerate() {
        let n = i + 1;
        let rid = prels.add("slide", &format!("slides/slide{n}.xml"), false);
        sld_ids.push_str(&format!("<p:sldId id=\"{}\" r:id=\"{rid}\"/>", 256 + i));
        let lay = Lay::of(slide.layout.as_deref(), &slide.items);
        let mut out = write_slide(&mut w, slide, lay);
        if !slide.notes.trim().is_empty() {
            out.rels.add("notesSlide", &format!("../notesSlides/notesSlide{n}.xml"), false);
            let mut nrels = PartRels::new();
            nrels.add("notesMaster", "../notesMasters/notesMaster1.xml", false);
            nrels.add("slide", &format!("../slides/slide{n}.xml"), false);
            let nx = notes_slide_xml(&slide.notes);
            put(&mut parts, &mut overrides, format!("ppt/notesSlides/notesSlide{n}.xml"), Some(format!("{CT_P}.notesSlide+xml")), nx);
            put(&mut parts, &mut overrides, format!("ppt/notesSlides/_rels/notesSlide{n}.xml.rels"), None, rels_xml(&nrels.list));
        }
        put(&mut parts, &mut overrides, format!("ppt/slides/slide{n}.xml"), Some(format!("{CT_P}.slide+xml")), out.xml);
        put(&mut parts, &mut overrides, format!("ppt/slides/_rels/slide{n}.xml.rels"), None, rels_xml(&out.rels.list));
    }
    let notes_master_ref = if any_notes {
        let rid = prels.add("notesMaster", "notesMasters/notesMaster1.xml", false);
        format!("<p:notesMasterIdLst><p:notesMasterId r:id=\"{rid}\"/></p:notesMasterIdLst>")
    } else {
        String::new()
    };
    prels.add("presProps", "presProps.xml", false);
    prels.add("viewProps", "viewProps.xml", false);
    prels.add("theme", "theme/theme1.xml", false);
    prels.add("tableStyles", "tableStyles.xml", false);

    let deftext = lvl_styles(&[1800], "+mn-lt", false, "l");
    let pres = format!(
        "{XML_HEAD}<p:presentation {NS_DECL} saveSubsetFonts=\"1\"><p:sldMasterIdLst><p:sldMasterId id=\"2147483648\" r:id=\"rId1\"/></p:sldMasterIdLst>{notes_master_ref}<p:sldIdLst>{sld_ids}</p:sldIdLst><p:sldSz cx=\"{cx}\" cy=\"{cy}\"/><p:notesSz cx=\"6858000\" cy=\"9144000\"/><p:defaultTextStyle><a:defPPr><a:defRPr lang=\"en-US\"/></a:defPPr>{deftext}</p:defaultTextStyle></p:presentation>"
    );
    put(&mut parts, &mut overrides, "ppt/presentation.xml".into(), Some(format!("{CT_P}.presentation.main+xml")), pres);
    put(&mut parts, &mut overrides, "ppt/_rels/presentation.xml.rels".into(), None, rels_xml(&prels.list));

    // master + layouts
    put(&mut parts, &mut overrides, "ppt/slideMasters/slideMaster1.xml".into(), Some(format!("{CT_P}.slideMaster+xml")), master_xml(wpt, hpt));
    let mut mrels = PartRels::new();
    for i in 1..=5 {
        mrels.add("slideLayout", &format!("../slideLayouts/slideLayout{i}.xml"), false);
    }
    mrels.add("theme", "../theme/theme1.xml", false);
    put(&mut parts, &mut overrides, "ppt/slideMasters/_rels/slideMaster1.xml.rels".into(), None, rels_xml(&mrels.list));
    for lay in [Lay::Title, Lay::Content, Lay::Section, Lay::Two, Lay::Blank] {
        let n = lay.file_no();
        put(&mut parts, &mut overrides, format!("ppt/slideLayouts/slideLayout{n}.xml"), Some(format!("{CT_P}.slideLayout+xml")), layout_xml(lay, wpt, hpt));
        let mut lrels = PartRels::new();
        lrels.add("slideMaster", "../slideMasters/slideMaster1.xml", false);
        put(&mut parts, &mut overrides, format!("ppt/slideLayouts/_rels/slideLayout{n}.xml.rels"), None, rels_xml(&lrels.list));
    }
    let theme_ct = "application/vnd.openxmlformats-officedocument.theme+xml".to_string();
    put(&mut parts, &mut overrides, "ppt/theme/theme1.xml".into(), Some(theme_ct.clone()), theme_xml(&doc.theme));
    if any_notes {
        put(&mut parts, &mut overrides, "ppt/notesMasters/notesMaster1.xml".into(), Some(format!("{CT_P}.notesMaster+xml")), notes_master_xml());
        let mut nm = PartRels::new();
        nm.add("theme", "../theme/theme2.xml", false);
        put(&mut parts, &mut overrides, "ppt/notesMasters/_rels/notesMaster1.xml.rels".into(), None, rels_xml(&nm.list));
        put(&mut parts, &mut overrides, "ppt/theme/theme2.xml".into(), Some(theme_ct), theme_xml(&Theme::default()));
    }
    put(&mut parts, &mut overrides, "ppt/presProps.xml".into(), Some(format!("{CT_P}.presProps+xml")), format!("{XML_HEAD}<p:presentationPr {NS_DECL}/>"));
    put(
        &mut parts,
        &mut overrides,
        "ppt/viewProps.xml".into(),
        Some(format!("{CT_P}.viewProps+xml")),
        format!("{XML_HEAD}<p:viewPr {NS_DECL}><p:normalViewPr><p:restoredLeft sz=\"15620\"/><p:restoredTop sz=\"94660\"/></p:normalViewPr><p:gridSpacing cx=\"76200\" cy=\"76200\"/></p:viewPr>"),
    );
    put(
        &mut parts,
        &mut overrides,
        "ppt/tableStyles.xml".into(),
        Some(format!("{CT_P}.tableStyles+xml")),
        format!("{XML_HEAD}<a:tblStyleLst xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" def=\"{TBL_NO_STYLE}\"/>"),
    );

    // docProps
    let meta_str = |k: &str| doc.meta.get(k).and_then(Value::as_str).map(str::trim).filter(|s| !s.is_empty()).map(String::from);
    let now = now_w3cdtf();
    let core = format!(
        "{XML_HEAD}<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">{}{}{}<dcterms:created xsi:type=\"dcterms:W3CDTF\">{now}</dcterms:created><dcterms:modified xsi:type=\"dcterms:W3CDTF\">{now}</dcterms:modified></cp:coreProperties>",
        meta_str("title").map(|t| format!("<dc:title>{}</dc:title>", xml_text(&t))).unwrap_or_default(),
        meta_str("subject").map(|t| format!("<dc:subject>{}</dc:subject>", xml_text(&t))).unwrap_or_default(),
        meta_str("author").map(|t| format!("<dc:creator>{}</dc:creator>", xml_text(&t))).unwrap_or_default(),
    );
    put(&mut parts, &mut overrides, "docProps/core.xml".into(), Some("application/vnd.openxmlformats-package.core-properties+xml".into()), core);
    let app = format!(
        "{XML_HEAD}<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\"><Application>Microsoft Office PowerPoint</Application><PresentationFormat>Custom</PresentationFormat><Slides>{}</Slides><Notes>{}</Notes><HiddenSlides>{}</HiddenSlides><AppVersion>16.0000</AppVersion></Properties>",
        slides.len(),
        slides.iter().filter(|s| !s.notes.trim().is_empty()).count(),
        slides.iter().filter(|s| s.hidden).count()
    );
    put(&mut parts, &mut overrides, "docProps/app.xml".into(), Some("application/vnd.openxmlformats-officedocument.extended-properties+xml".into()), app);
    let mut root = PartRels::new();
    root.add("officeDocument", "ppt/presentation.xml", false);
    root.list.push(("rId2".into(), "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties".into(), "docProps/core.xml".into(), false));
    root.add("extended-properties", "docProps/app.xml", false);
    put(&mut parts, &mut overrides, "_rels/.rels".into(), None, rels_xml(&root.list));

    let mut ct = format!("{XML_HEAD}<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/>");
    for e in &w.exts {
        let mime = match *e {
            "jpeg" => "image/jpeg",
            "emf" => "image/x-emf",
            "wmf" => "image/x-wmf",
            other => mime_for_ext(other),
        };
        ct.push_str(&format!("<Default Extension=\"{e}\" ContentType=\"{mime}\"/>"));
    }
    for (n, t) in &overrides {
        ct.push_str(&format!("<Override PartName=\"{n}\" ContentType=\"{t}\"/>"));
    }
    ct.push_str("</Types>");

    let mut buf = Cursor::new(Vec::new());
    {
        let mut zip = ZipWriter::new(&mut buf);
        let opt = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
        let mut add = |name: &str, data: &[u8]| -> Result<(), String> {
            zip.start_file(name, opt).map_err(|e| e.to_string())?;
            zip.write_all(data).map_err(|e| e.to_string())
        };
        add("[Content_Types].xml", ct.as_bytes())?;
        for (n, d) in &parts {
            add(n, d)?;
        }
        for (n, d) in &w.media {
            add(n, d)?;
        }
        zip.finish().map_err(|e| e.to_string())?;
    }
    Ok((buf.into_inner(), w.notes))
}

// ============================ outlines ============================

fn para_plain(p: &Para) -> String {
    p.runs.iter().map(|r| r.t.as_str()).collect::<String>().replace('\n', " ")
}

fn md_escape(s: &str) -> String {
    let mut o = String::with_capacity(s.len());
    for c in s.chars() {
        if matches!(c, '\\' | '*' | '_' | '`' | '[' | ']' | '<' | '>' | '#') {
            o.push('\\');
        }
        o.push(c);
    }
    o
}

fn para_md(p: &Para) -> String {
    let mut s = String::new();
    for r in &p.runs {
        let t = r.t.replace('\n', " ");
        if t.trim().is_empty() {
            s.push_str(&t);
            continue;
        }
        let lead = &t[..t.len() - t.trim_start().len()];
        let trail = &t[t.trim_end().len()..];
        let mut core = md_escape(t.trim());
        if r.b && r.i {
            core = format!("***{core}***");
        } else if r.b {
            core = format!("**{core}**");
        } else if r.i {
            core = format!("*{core}*");
        }
        if r.s {
            core = format!("~~{core}~~");
        }
        if let Some(l) = r.link.as_deref().and_then(safe_link) {
            core = format!("[{core}]({})", l.replace(' ', "%20").replace(')', "%29"));
        }
        s.push_str(lead);
        s.push_str(&core);
        s.push_str(trail);
    }
    s.trim().to_string()
}

/// (title text, item index, para index) for a slide.
fn slide_title(s: &Slide) -> Option<(String, usize, usize)> {
    for (i, it) in s.items.iter().enumerate() {
        if let Item::Text(t) = it {
            if t.ph.as_deref() == Some("title") {
                let txt = t.paras.iter().map(para_plain).filter(|x| !x.trim().is_empty()).collect::<Vec<_>>().join(" ");
                if !txt.trim().is_empty() {
                    return Some((txt.trim().to_string(), i, usize::MAX));
                }
            }
        }
    }
    for (i, it) in s.items.iter().enumerate() {
        let paras = match it {
            Item::Text(t) => &t.paras,
            Item::Shape(sh) => &sh.paras,
            _ => continue,
        };
        if let Some((pi, p)) = paras.iter().enumerate().find(|(_, p)| !para_plain(p).trim().is_empty()) {
            return Some((para_plain(p).trim().to_string(), i, pi));
        }
    }
    None
}

fn cell_text(c: &SlideTableCell) -> String {
    c.paras.iter().map(para_plain).filter(|t| !t.trim().is_empty()).collect::<Vec<_>>().join(" ").trim().to_string()
}

/// Outline of the deck as Markdown (one "## n. title" section per slide,
/// bullets for body text, tables as GFM tables, speaker notes as a blockquote).
pub fn slides_to_markdown(doc: &SlidesDoc) -> String {
    let mut out = String::new();
    if let Some(t) = doc.meta.get("title").and_then(Value::as_str).map(str::trim).filter(|t| !t.is_empty()) {
        out.push_str(&format!("# {}\n\n", md_escape(t)));
    }
    for (n, s) in doc.slides.iter().enumerate() {
        let title = slide_title(s);
        let head = title.as_ref().map(|t| md_escape(&t.0)).unwrap_or_else(|| "Untitled slide".into());
        out.push_str(&format!("## {}. {head}{}\n\n", n + 1, if s.hidden { " (hidden)" } else { "" }));
        let mut body = String::new();
        for (i, it) in s.items.iter().enumerate() {
            match it {
                Item::Text(TextItem { paras, .. }) | Item::Shape(ShapeItem { paras, .. }) => {
                    let skip_all = title.as_ref().map(|t| t.1 == i && t.2 == usize::MAX).unwrap_or(false);
                    if skip_all {
                        continue;
                    }
                    let mut counters = [0usize; 9];
                    for (pi, p) in paras.iter().enumerate() {
                        if title.as_ref().map(|t| t.1 == i && t.2 == pi).unwrap_or(false) {
                            continue;
                        }
                        let txt = para_md(p);
                        if txt.is_empty() {
                            continue;
                        }
                        let lvl = p.level.min(8) as usize;
                        let indent = "  ".repeat(lvl);
                        if p.bullet.as_deref() == Some("ol") {
                            counters[lvl] += 1;
                            for c in counters.iter_mut().skip(lvl + 1) {
                                *c = 0;
                            }
                            body.push_str(&format!("{indent}{}. {txt}\n", counters[lvl]));
                        } else {
                            body.push_str(&format!("{indent}- {txt}\n"));
                        }
                    }
                }
                Item::Table(tb) => {
                    let ncols = tb.rows.iter().map(|r| r.len()).max().unwrap_or(0);
                    if ncols == 0 {
                        continue;
                    }
                    if !body.is_empty() && !body.ends_with("\n\n") {
                        body.push('\n');
                    }
                    let fmt_row = |r: &Vec<SlideTableCell>| {
                        let cells: Vec<String> = (0..ncols).map(|c| r.get(c).map(cell_text).unwrap_or_default().replace('|', "\\|")).collect();
                        format!("| {} |\n", cells.join(" | "))
                    };
                    for (ri, r) in tb.rows.iter().enumerate() {
                        body.push_str(&fmt_row(r));
                        if ri == 0 {
                            body.push_str(&format!("|{}\n", " --- |".repeat(ncols)));
                        }
                    }
                    body.push('\n');
                }
                Item::Image(im) => {
                    if !body.is_empty() && !body.ends_with("\n\n") {
                        body.push('\n');
                    }
                    let alt = im.alt.as_deref().map(|a| a.replace(['[', ']'], "")).unwrap_or_else(|| "Picture".into());
                    body.push_str(&format!("![{}](image)\n\n", alt.trim()));
                }
                Item::Other(_) => {}
            }
        }
        if !body.is_empty() {
            out.push_str(body.trim_end());
            out.push_str("\n\n");
        }
        if !s.notes.trim().is_empty() {
            for (k, line) in s.notes.trim().lines().enumerate() {
                if k == 0 {
                    out.push_str(&format!("> Notes: {}\n", line.trim_end()));
                } else {
                    out.push_str(&format!("> {}\n", line.trim_end()));
                }
            }
            out.push('\n');
        }
    }
    let t = out.trim_end().to_string();
    if t.is_empty() { t } else { t + "\n" }
}

/// Plain-text outline of the deck.
pub fn slides_to_text(doc: &SlidesDoc) -> String {
    let mut out = String::new();
    if let Some(t) = doc.meta.get("title").and_then(Value::as_str).map(str::trim).filter(|t| !t.is_empty()) {
        out.push_str(&format!("{t}\n\n"));
    }
    for (n, s) in doc.slides.iter().enumerate() {
        let title = slide_title(s);
        out.push_str(&format!(
            "Slide {}{}{}\n",
            n + 1,
            title.as_ref().map(|t| format!(": {}", t.0)).unwrap_or_default(),
            if s.hidden { " (hidden)" } else { "" }
        ));
        for (i, it) in s.items.iter().enumerate() {
            match it {
                Item::Text(TextItem { paras, .. }) | Item::Shape(ShapeItem { paras, .. }) => {
                    if title.as_ref().map(|t| t.1 == i && t.2 == usize::MAX).unwrap_or(false) {
                        continue;
                    }
                    let mut counters = [0usize; 9];
                    for (pi, p) in paras.iter().enumerate() {
                        if title.as_ref().map(|t| t.1 == i && t.2 == pi).unwrap_or(false) {
                            continue;
                        }
                        let txt = para_plain(p);
                        if txt.trim().is_empty() {
                            continue;
                        }
                        let lvl = p.level.min(8) as usize;
                        let indent = "  ".repeat(lvl + 1);
                        match p.bullet.as_deref() {
                            Some("ol") => {
                                counters[lvl] += 1;
                                for c in counters.iter_mut().skip(lvl + 1) {
                                    *c = 0;
                                }
                                out.push_str(&format!("{indent}{}. {}\n", counters[lvl], txt.trim()));
                            }
                            Some(_) => out.push_str(&format!("{indent}\u{2022} {}\n", txt.trim())),
                            None => out.push_str(&format!("{indent}{}\n", txt.trim())),
                        }
                    }
                }
                Item::Table(tb) => {
                    for r in &tb.rows {
                        out.push_str(&format!("  {}\n", r.iter().map(cell_text).collect::<Vec<_>>().join("\t")));
                    }
                }
                Item::Image(im) => out.push_str(&format!("  [Picture{}]\n", im.alt.as_deref().map(|a| format!(": {}", a.trim())).unwrap_or_default())),
                Item::Other(_) => {}
            }
        }
        if !s.notes.trim().is_empty() {
            out.push_str(&format!("  Notes: {}\n", s.notes.trim().replace('\n', "\n         ")));
        }
        out.push('\n');
    }
    let t = out.trim_end().to_string();
    if t.is_empty() { t } else { t + "\n" }
}


// ============================ tests ============================

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;
    use std::io::Read;

    /// 1x1 PNG.
    const PNG_1X1: &[u8] = &[
        0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52, 0x00, 0x00, 0x00, 0x01,
        0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4, 0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41,
        0x54, 0x78, 0x9c, 0x63, 0x60, 0x00, 0x00, 0x00, 0x02, 0x00, 0x01, 0x54, 0xa2, 0x4f, 0x5d, 0x00, 0x00, 0x00, 0x00, 0x49,
        0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
    ];

    fn fixture(name: &str) -> Vec<u8> {
        let p = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/pptx").join(name);
        std::fs::read(&p).unwrap_or_else(|e| panic!("{}: {e}", p.display()))
    }

    fn zip_parts(parts: &[(&str, &[u8])]) -> Vec<u8> {
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

    fn unzip(bytes: &[u8]) -> Vec<(String, Vec<u8>)> {
        let mut z = ZipArchive::new(Cursor::new(bytes)).unwrap();
        let mut out = Vec::new();
        for i in 0..z.len() {
            let mut f = z.by_index(i).unwrap();
            let mut v = Vec::new();
            f.read_to_end(&mut v).unwrap();
            out.push((f.name().map(|n| n.to_string()).unwrap_or_default(), v));
        }
        out
    }

    fn warnings(d: &SlidesDoc) -> Vec<String> {
        d.meta.get("warnings").and_then(Value::as_array).map(|a| a.iter().filter_map(|v| v.as_str().map(String::from)).collect()).unwrap_or_default()
    }

    fn text_of(it: &Item) -> String {
        let paras = match it {
            Item::Text(t) => &t.paras,
            Item::Shape(s) => &s.paras,
            _ => return String::new(),
        };
        paras.iter().map(para_plain).collect::<Vec<_>>().join("|")
    }

    fn deck(v: Value) -> SlidesDoc {
        serde_json::from_value(v).unwrap()
    }

    // ---------------------------------------------------------------- fixtures

    #[test]
    fn fixture_basic_deck() {
        let d = parse_pptx(&fixture("basic.pptx")).unwrap();
        assert_eq!((d.size.w, d.size.h), (720.0, 540.0));
        assert_eq!(d.meta["title"], "Quarterly Review");
        assert_eq!(d.meta["author"], "Ada Lovelace");
        assert_eq!(d.theme.accent.as_deref(), Some("#4F81BD"));
        assert_eq!(d.theme.color.as_deref(), Some("#000000"));
        assert_eq!(d.theme.bg.as_deref(), Some("#FFFFFF"));
        assert_eq!(d.theme.font.as_deref(), Some("Calibri"));
        assert_eq!(d.slides.len(), 6);
        let w = warnings(&d);
        assert!(w.iter().any(|x| x == W_SHAPES), "{w:?}");
        assert!(w.iter().any(|x| x == W_CHARTS), "{w:?}");

        // 1: title slide, inherited placeholder geometry + styles, notes, fade.
        let s = &d.slides[0];
        assert_eq!(s.layout.as_deref(), Some("title"));
        assert_eq!(s.transition.as_deref(), Some("fade"));
        assert_eq!(s.notes, "Welcome everyone.\nIntroduce the agenda.");
        let Item::Text(t) = &s.items[0] else { panic!() };
        assert_eq!(t.ph.as_deref(), Some("title"));
        assert_eq!((t.x, t.y, t.w), (54.0, 167.75, 612.0));
        assert_eq!(t.paras[0].align.as_deref(), Some("center"));
        assert_eq!(t.paras[0].runs[0].size, Some(44.0));
        assert_eq!(t.paras[0].runs[0].t, "Quarterly Review");
        let Item::Text(sub) = &s.items[1] else { panic!() };
        assert_eq!(sub.ph.as_deref(), Some("subtitle"));
        assert_eq!(sub.paras[0].bullet, None);
        assert_eq!(sub.paras[0].runs[0].color.as_deref(), Some("#404040")); // tx1 tint 75%

        // 2: bullets with levels and run formatting.
        let s = &d.slides[1];
        assert_eq!(s.layout.as_deref(), Some("content"));
        assert_eq!(s.transition.as_deref(), Some("push"));
        let Item::Text(body) = &s.items[1] else { panic!() };
        assert_eq!(body.ph.as_deref(), Some("body"));
        let p = &body.paras;
        assert_eq!(p[0].bullet.as_deref(), Some("ul"));
        assert_eq!(p[0].runs[0].size, Some(32.0));
        assert_eq!((p[1].level, p[1].runs[0].size), (1, Some(28.0)));
        assert_eq!(p[2].level, 2);
        assert!(p[2].runs[0].b && p[2].runs[0].t == "Bold ");
        assert!(p[2].runs[1].i && p[2].runs[1].size == Some(20.0));
        assert_eq!(p[2].runs[2].color.as_deref(), Some("#C00000"));
        assert_eq!(p[4].runs[0].link.as_deref(), Some("https://example.com/"));
        let acc = parse_hex(p[4].runs[1].color.as_deref().unwrap()).unwrap();
        // accent2 C0504D with lumMod 75% ~ 943634 in PowerPoint
        assert!(acc[0].abs_diff(0x94) <= 3 && acc[1].abs_diff(0x36) <= 3 && acc[2].abs_diff(0x34) <= 3, "{acc:?}");

        // 3: table (+ default table style approximation) and right-aligned text box.
        let s = &d.slides[2];
        let Item::Table(tb) = &s.items[1] else { panic!("{:?}", s.items[1]) };
        assert_eq!(tb.col_w, vec![216.0, 108.0, 108.0]);
        assert_eq!((tb.x, tb.y, tb.w, tb.h), (36.0, 108.0, 432.0, 108.0));
        assert_eq!(tb.rows.len(), 3);
        assert_eq!(cell_text(&tb.rows[2][0]), "South");
        assert_eq!(tb.rows[1][0].fill.as_deref(), Some("#FFF2CC"));
        assert_eq!(tb.rows[0][0].fill.as_deref(), Some("#4F81BD"));
        assert!(tb.rows[0][0].paras[0].runs[0].b);
        assert_eq!(tb.rows[0][0].paras[0].runs[0].color.as_deref(), Some("#FFFFFF"));
        let Item::Text(src) = &s.items[2] else { panic!() };
        assert_eq!(src.ph, None);
        assert_eq!(src.paras[0].align.as_deref(), Some("right"));
        assert_eq!(src.paras[0].runs[0].size, Some(12.0));

        // 4: shapes, theme-color fills, picture, scaled group, background.
        let s = &d.slides[3];
        assert_eq!(s.bg.as_deref(), Some("#F0F0F0"));
        let Item::Shape(rr) = &s.items[0] else { panic!() };
        assert_eq!(rr.shape, "roundRect");
        assert_eq!((rr.x, rr.y, rr.w, rr.h, rr.rot), (72.0, 72.0, 216.0, 108.0, 15.0));
        assert_eq!(rr.fill.as_deref(), Some("#4472C4"));
        assert_eq!((rr.line.as_deref(), rr.line_w), (Some("#1F3864"), Some(2.0)));
        assert_eq!(rr.valign.as_deref(), Some("bottom"));
        assert_eq!(rr.paras[0].runs[0].color.as_deref(), Some("#FFFFFF")); // style fontRef lt1
        let Item::Shape(el) = &s.items[1] else { panic!() };
        assert_eq!(el.shape, "ellipse");
        let f = parse_hex(el.fill.as_deref().unwrap()).unwrap(); // accent1 lumMod 60% lumOff 40% ~ 95B3D7
        assert!(f[0].abs_diff(0x95) <= 2 && f[1].abs_diff(0xB3) <= 2 && f[2].abs_diff(0xD7) <= 2, "{f:?}");
        let Item::Shape(cloud) = &s.items[2] else { panic!() };
        assert_eq!(cloud.shape, "rect");
        let Item::Image(im) = &s.items[3] else { panic!() };
        assert!(im.src.starts_with("data:image/png;base64,"));
        assert_eq!((im.x, im.y, im.w, im.h), (72.0, 252.0, 144.0, 108.0));
        let Item::Shape(g1) = &s.items[4] else { panic!() };
        let Item::Shape(g2) = &s.items[5] else { panic!() };
        assert_eq!((g1.x, g1.y, g1.w, g1.h), (360.0, 288.0, 144.0, 144.0));
        assert_eq!((g2.shape.as_str(), g2.x, g2.w), ("triangle", 576.0, 144.0));
        assert_eq!(g2.fill.as_deref(), Some("#800080"));

        // 5: hidden section header.
        let s = &d.slides[4];
        assert!(s.hidden);
        assert_eq!(s.layout.as_deref(), Some("section"));
        assert_eq!(text_of(&s.items[0]), "Appendix");
        // 6: numbered list in a two-content layout + chart placeholder.
        let s = &d.slides[5];
        assert_eq!(s.layout.as_deref(), Some("two"));
        let Item::Text(left) = &s.items[1] else { panic!() };
        assert!(left.paras.iter().all(|p| p.bullet.as_deref() == Some("ol")));
        let Item::Text(right) = &s.items[2] else { panic!() };
        assert!(right.x > left.x + left.w - 1.0);
        let Item::Shape(ch) = &s.items[3] else { panic!() };
        assert_eq!(text_of(&s.items[3]), "[Chart]");
        assert_eq!(ch.fill.as_deref(), Some("#F2F2F2"));
    }

    #[test]
    fn fixture_widescreen_alternate_content_transition() {
        let d = parse_pptx(&fixture("widescreen.pptx")).unwrap();
        assert_eq!((d.size.w, d.size.h), (960.0, 540.0));
        // p14:vortex (Choice) is not representable; the Fallback wipe is exact.
        assert_eq!(d.slides[0].transition.as_deref(), Some("wipe"));
        assert!(!warnings(&d).iter().any(|w| w == W_TRANSITION));
        assert_eq!(text_of(&d.slides[0].items[1]), "Sixteen by nine");
    }

    #[test]
    fn fixture_themed_master_decorations_and_color_maps() {
        let d = parse_pptx(&fixture("themed.pptx")).unwrap();
        assert_eq!(d.theme.accent.as_deref(), Some("#2E75B6"));
        assert_eq!((d.theme.font.as_deref(), d.theme.head_font.as_deref()), (Some("Verdana"), Some("Georgia")));
        let w = warnings(&d);
        assert!(w.iter().any(|x| x == W_ANIM) && w.iter().any(|x| x == W_COMMENTS), "{w:?}");
        // 1: master logo + rule copied behind the slide content; theme fonts stay implicit.
        let s = &d.slides[0].items;
        let Item::Image(logo) = &s[0] else { panic!("{:?}", s[0]) };
        assert_eq!((logo.alt.as_deref(), logo.x, logo.y, logo.w), (Some("Company logo"), 888.0, 12.0, 48.0));
        let Item::Shape(rule) = &s[1] else { panic!() };
        assert_eq!((rule.shape.as_str(), rule.line.as_deref(), rule.line_w, rule.w), ("line", Some("#2E75B6"), Some(1.5), 888.0));
        let Item::Text(title) = &s[2] else { panic!() };
        assert_eq!(title.paras[0].runs[0].font, None);
        let Item::Text(body) = &s[3] else { panic!() };
        assert_eq!(body.paras[0].runs[0].font, None);
        assert_eq!(body.paras[1].runs[0].font.as_deref(), Some("Georgia")); // +mj-lt in body text
        // 2: Blank layout has showMasterSp="0"; svgBlip picture uses its PNG blip.
        let s = &d.slides[1].items;
        assert_eq!(s.len(), 2);
        let Item::Image(im) = &s[0] else { panic!() };
        assert!(im.src.starts_with("data:image/png;base64,"));
        // 3: picture background becomes a full-slide picture at the bottom.
        let Item::Image(bg) = &d.slides[2].items[0] else { panic!() };
        assert_eq!((bg.x, bg.y, bg.w, bg.h), (0.0, 0.0, 960.0, 540.0));
        // 4: color-map override: dark background, white text (explicit, deck text color is black).
        let s = &d.slides[3];
        assert_eq!(s.bg.as_deref(), Some("#000000"));
        let Item::Text(t) = &s.items[3] else { panic!() };
        assert_eq!(t.paras[0].runs[0].color.as_deref(), Some("#FFFFFF"));
    }

    #[test]
    fn fixture_rewrite_keeps_content() {
        for f in ["basic.pptx", "widescreen.pptx", "themed.pptx"] {
            let a = parse_pptx(&fixture(f)).unwrap();
            let (bytes, _notes) = write_pptx(&a).unwrap();
            check_package(&bytes);
            let b = parse_pptx(&bytes).unwrap();
            assert_eq!(a.slides.len(), b.slides.len());
            assert_eq!((a.size.w, a.size.h), (b.size.w, b.size.h));
            for (sa, sb) in a.slides.iter().zip(&b.slides) {
                assert_eq!(sa.notes, sb.notes);
                assert_eq!(sa.hidden, sb.hidden);
                assert_eq!(sa.transition, sb.transition);
                assert_eq!(sa.bg, sb.bg);
                if sa.layout.is_some() {
                    assert_eq!(sa.layout, sb.layout);
                }
                let ta: Vec<String> = sa.items.iter().map(text_of).collect();
                let tb: Vec<String> = sb.items.iter().map(text_of).collect();
                assert_eq!(ta, tb, "{f}");
                assert_eq!(sa.items.len(), sb.items.len());
            }
        }
    }

    // ---------------------------------------------------------------- round trip

    fn full_deck() -> SlidesDoc {
        let png = data_url("image/png", PNG_1X1);
        let src = r##"{
            "kind": "slides",
            "meta": { "title": "Round <Trip> & Co", "author": "Grace Hopper" },
            "size": { "w": 960, "h": 540 },
            "theme": { "font": "Georgia", "headFont": "Verdana", "color": "#222222", "accent": "#C00000", "bg": "#FFFFF0" },
            "slides": [
                { "layout": "title", "transition": "fade", "notes": "First line\nSecond line",
                  "items": [
                    { "type": "text", "ph": "title", "x": 80, "y": 170, "w": 800, "h": 110,
                      "paras": [{ "align": "center", "runs": [{ "t": "Hello & <world>", "size": 44, "b": true }] }] },
                    { "type": "text", "ph": "subtitle", "x": 80, "y": 300, "w": 800, "h": 60, "valign": "middle",
                      "paras": [{ "align": "center", "runs": [{ "t": "", "size": 24 }] }] }
                  ] },
                { "layout": "content", "bg": "#EEF2FF", "transition": "push", "hidden": true,
                  "items": [
                    { "type": "text", "ph": "title", "x": 60, "y": 27, "w": 840, "h": 97, "paras": [{ "runs": [{ "t": "Agenda", "size": 40, "color": "#1F3864" }] }] },
                    { "type": "text", "ph": "body", "x": 60, "y": 140, "w": 840, "h": 345, "paras": [
                        { "bullet": "ul", "runs": [{ "t": "Point one", "size": 28 }] },
                        { "bullet": "ul", "level": 1, "runs": [{ "t": "Sub ", "size": 24 }, { "t": "italic", "i": true, "size": 24 }, { "t": " link", "size": 24, "link": "https://example.com/a b?x=1&y=2" }] },
                        { "bullet": "ol", "level": 2, "runs": [{ "t": "Numbered", "size": 20, "u": true, "s": true }] },
                        { "align": "justify", "runs": [{ "t": "x", "size": 18, "sup": true }, { "t": "2", "size": 18, "sub": true }, { "t": "font", "size": 18, "font": "Courier New", "hl": "#FFFF00" }] },
                        { "align": "right", "runs": [{ "t": "line1\nline2", "size": 18 }] }
                    ] }
                  ] },
                { "layout": "two", "transition": "wipe",
                  "items": [
                    { "type": "text", "ph": "body", "x": 60, "y": 140, "w": 410, "h": 345, "paras": [{ "bullet": "ul", "runs": [{ "t": "Left", "size": 24 }] }] },
                    { "type": "text", "ph": "body", "x": 490, "y": 140, "w": 410, "h": 345, "paras": [{ "runs": [{ "t": "Right", "size": 24 }] }] }
                  ] },
                { "layout": "blank", "transition": "zoom",
                  "items": [
                    { "type": "shape", "shape": "roundRect", "x": 100, "y": 100, "w": 200, "h": 100, "rot": 30, "fill": "#4472C4", "line": "#000000", "lineW": 2.5, "valign": "middle",
                      "paras": [{ "align": "center", "runs": [{ "t": "Box", "size": 20, "color": "#FFFFFF" }] }] },
                    { "type": "shape", "shape": "ellipse", "x": 350, "y": 100, "w": 100, "h": 100 },
                    { "type": "shape", "shape": "line", "x": 500, "y": 100, "w": 200, "h": 0, "line": "#FF0000", "lineW": 3 },
                    { "type": "shape", "shape": "star5", "x": 10, "y": 300, "w": 50, "h": 50, "fill": "#FFC000", "valign": "bottom" },
                    { "type": "shape", "shape": "triangle", "x": 70, "y": 300, "w": 50, "h": 50, "fill": "#00B050" },
                    { "type": "shape", "shape": "diamond", "x": 130, "y": 300, "w": 50, "h": 50, "fill": "#00B050" },
                    { "type": "shape", "shape": "rightArrow", "x": 190, "y": 300, "w": 50, "h": 50, "fill": "#00B050" },
                    { "type": "shape", "shape": "leftArrow", "x": 250, "y": 300, "w": 50, "h": 50, "fill": "#00B050" },
                    { "type": "shape", "shape": "hexagon", "x": 310, "y": 300, "w": 50, "h": 50, "fill": "#00B050" },
                    { "type": "shape", "shape": "chevron", "x": 370, "y": 300, "w": 50, "h": 50, "fill": "#00B050" },
                    { "type": "shape", "shape": "rect", "x": 430, "y": 300, "w": 50, "h": 50, "line": "#7F7F7F" },
                    { "type": "image", "x": 600, "y": 300, "w": 120, "h": 90, "rot": 10, "src": "PNG", "alt": "A \"pixel\"" },
                    { "type": "text", "x": 600, "y": 420, "w": 300, "h": 40, "fill": "#F2F2F2", "line": "#BFBFBF", "lineW": 1, "valign": "bottom",
                      "paras": [{ "runs": [{ "t": "Caption", "size": 14, "font": "Arial" }] }] },
                    { "type": "table", "x": 40, "y": 380, "w": 500, "h": 120, "colW": [200, 150, 150], "rows": [
                        [{ "paras": [{ "runs": [{ "t": "Name", "b": true, "size": 14 }] }], "fill": "#DDEBF7" }, { "paras": [{ "runs": [{ "t": "Q1", "size": 14 }] }] }, { "paras": [{ "align": "right", "runs": [{ "t": "Q2", "size": 14 }] }] }],
                        [{ "paras": [{ "runs": [{ "t": "A|B", "size": 14 }] }] }, { "paras": [{ "runs": [{ "t": "1", "size": 14 }] }] }, { "paras": [{ "runs": [{ "t": "2", "size": 14 }] }] }]
                    ] }
                  ] },
                { "layout": "section", "items": [
                    { "type": "text", "ph": "title", "x": 60, "y": 160, "w": 840, "h": 130, "valign": "bottom", "paras": [{ "runs": [{ "t": "Part 2", "size": 40 }] }] },
                    { "type": "text", "ph": "body", "x": 60, "y": 300, "w": 840, "h": 70, "paras": [{ "runs": [{ "t": "Details", "size": 20 }] }] }
                  ] }
            ]
        }"##;
        serde_json::from_str(&src.replace("\"PNG\"", &format!("\"{png}\""))).unwrap()
    }

    fn close(a: f64, b: f64) -> bool {
        (a - b).abs() < 0.02
    }

    fn valign_norm(v: &Option<String>) -> String {
        v.clone().unwrap_or_else(|| "top".into())
    }

    fn assert_paras(a: &[Para], b: &[Para], ctx: &str) {
        assert_eq!(a.len(), b.len(), "{ctx}: paragraph count");
        for (pa, pb) in a.iter().zip(b) {
            assert_eq!(pa.align, pb.align, "{ctx}");
            assert_eq!(pa.level, pb.level, "{ctx}");
            assert_eq!(pa.bullet, pb.bullet, "{ctx}");
            assert_eq!(pa.runs.len(), pb.runs.len(), "{ctx}: {:?} vs {:?}", pa.runs, pb.runs);
            for (ra, rb) in pa.runs.iter().zip(&pb.runs) {
                assert_eq!(ra.t, rb.t, "{ctx}");
                assert_eq!((ra.b, ra.i, ra.u, ra.s, ra.sup, ra.sub), (rb.b, rb.i, rb.u, rb.s, rb.sup, rb.sub), "{ctx}: {}", ra.t);
                assert_eq!(ra.color, rb.color, "{ctx}: {}", ra.t);
                assert_eq!(ra.size, rb.size, "{ctx}: {}", ra.t);
                assert_eq!(ra.font, rb.font, "{ctx}: {}", ra.t);
                assert_eq!(ra.hl, rb.hl, "{ctx}: {}", ra.t);
                let la = ra.link.as_deref().map(uri_escape);
                assert_eq!(la, rb.link, "{ctx}: {}", ra.t);
            }
        }
    }

    fn assert_decks_equal(a: &SlidesDoc, b: &SlidesDoc) {
        assert_eq!(a.meta.get("title"), b.meta.get("title"));
        assert_eq!(a.meta.get("author"), b.meta.get("author"));
        assert!(close(a.size.w, b.size.w) && close(a.size.h, b.size.h));
        assert_eq!(a.theme.font, b.theme.font);
        assert_eq!(a.theme.head_font, b.theme.head_font);
        assert_eq!(a.theme.color, b.theme.color);
        assert_eq!(a.theme.accent, b.theme.accent);
        assert_eq!(a.theme.bg, b.theme.bg);
        assert_eq!(a.slides.len(), b.slides.len());
        for (n, (sa, sb)) in a.slides.iter().zip(&b.slides).enumerate() {
            let ctx = format!("slide {}", n + 1);
            assert_eq!(sa.layout, sb.layout, "{ctx}");
            assert_eq!(sa.bg, sb.bg, "{ctx}");
            assert_eq!(sa.notes, sb.notes, "{ctx}");
            assert_eq!(sa.hidden, sb.hidden, "{ctx}");
            assert_eq!(sa.transition, sb.transition, "{ctx}");
            assert_eq!(sa.items.len(), sb.items.len(), "{ctx}");
            for (k, (ia, ib)) in sa.items.iter().zip(&sb.items).enumerate() {
                let ctx = format!("{ctx} item {k}");
                match (ia, ib) {
                    (Item::Text(x), Item::Text(y)) => {
                        assert!(close(x.x, y.x) && close(x.y, y.y) && close(x.w, y.w) && close(x.h, y.h) && close(x.rot, y.rot), "{ctx}: {x:?} vs {y:?}");
                        assert_eq!(x.ph, y.ph, "{ctx}");
                        assert_eq!((&x.fill, &x.line), (&y.fill, &y.line), "{ctx}");
                        assert_eq!(x.line_w, y.line_w, "{ctx}");
                        assert_eq!(valign_norm(&x.valign), valign_norm(&y.valign), "{ctx}");
                        assert_paras(&x.paras, &y.paras, &ctx);
                    }
                    (Item::Shape(x), Item::Shape(y)) => {
                        assert!(close(x.x, y.x) && close(x.y, y.y) && close(x.w, y.w) && close(x.h, y.h) && close(x.rot, y.rot), "{ctx}: {x:?} vs {y:?}");
                        assert_eq!(x.shape, y.shape, "{ctx}");
                        assert_eq!((&x.fill, &x.line), (&y.fill, &y.line), "{ctx}");
                        assert_eq!(x.line_w, y.line_w, "{ctx}");
                        assert_eq!(valign_norm(&x.valign), valign_norm(&y.valign), "{ctx}");
                        assert_paras(&x.paras, &y.paras, &ctx);
                    }
                    (Item::Image(x), Item::Image(y)) => {
                        assert!(close(x.x, y.x) && close(x.y, y.y) && close(x.w, y.w) && close(x.h, y.h) && close(x.rot, y.rot), "{ctx}");
                        assert_eq!(parse_data_url(&x.src).unwrap().1, parse_data_url(&y.src).unwrap().1, "{ctx}");
                        assert_eq!(x.alt, y.alt, "{ctx}");
                    }
                    (Item::Table(x), Item::Table(y)) => {
                        assert!(close(x.x, y.x) && close(x.y, y.y) && close(x.w, y.w) && close(x.h, y.h), "{ctx}: {x:?} vs {y:?}");
                        assert_eq!(x.col_w.len(), y.col_w.len());
                        assert!(x.col_w.iter().zip(&y.col_w).all(|(p, q)| close(*p, *q)), "{ctx}");
                        assert_eq!(x.rows.len(), y.rows.len());
                        for (ra, rb) in x.rows.iter().zip(&y.rows) {
                            assert_eq!(ra.len(), rb.len());
                            for (ca, cb) in ra.iter().zip(rb) {
                                assert_eq!(ca.fill, cb.fill, "{ctx}");
                                assert_paras(&ca.paras, &cb.paras, &ctx);
                            }
                        }
                    }
                    _ => panic!("{ctx}: item type changed: {ia:?} vs {ib:?}"),
                }
            }
        }
    }

    #[test]
    fn round_trip_keeps_every_contract_field() {
        let a = full_deck();
        let (bytes, notes) = write_pptx(&a).unwrap();
        assert!(notes.is_empty(), "{notes:?}");
        check_package(&bytes);
        let b = parse_pptx(&bytes).unwrap();
        assert!(warnings(&b).is_empty(), "{:?}", warnings(&b));
        assert_decks_equal(&a, &b);
        // Second generation is stable too.
        let (bytes2, _) = write_pptx(&b).unwrap();
        let c = parse_pptx(&bytes2).unwrap();
        assert_decks_equal(&b, &c);
    }

    #[test]
    fn round_trip_four_by_three_and_unset_fields() {
        let a = deck(json!({ "size": { "w": 720, "h": 540 }, "slides": [
            { "items": [{ "type": "text", "ph": "title", "x": 10, "y": 10, "w": 300, "h": 50, "paras": [{ "runs": [{ "t": "T" }] }] }] },
            { "items": [] }
        ] }));
        let (bytes, _) = write_pptx(&a).unwrap();
        let b = parse_pptx(&bytes).unwrap();
        assert_eq!((b.size.w, b.size.h), (720.0, 540.0));
        assert_eq!(b.slides.len(), 2);
        // Unset size falls back to the master title style; layout is chosen by content.
        let Item::Text(t) = &b.slides[0].items[0] else { panic!() };
        assert_eq!(t.paras[0].runs[0].size, Some(44.0));
        assert_eq!(t.paras[0].runs[0].font, None);
        assert_eq!(b.slides[0].layout.as_deref(), Some("content"));
        assert_eq!(b.slides[1].layout.as_deref(), Some("blank"));
        assert!(b.slides[1].items.is_empty());
    }

    // ---------------------------------------------------------------- package checks

    /// Every part is well-formed XML with a content type; every internal
    /// relationship target exists; slide/layout ids and cNvPr ids are valid.
    fn check_package(bytes: &[u8]) {
        let parts = unzip(bytes);
        let names: Vec<&str> = parts.iter().map(|(n, _)| n.as_str()).collect();
        let get = |n: &str| parts.iter().find(|(m, _)| m == n).map(|(_, d)| String::from_utf8(d.clone()).unwrap());
        let ct = get("[Content_Types].xml").expect("content types");
        for (n, d) in &parts {
            let ext = n.rsplit('.').next().unwrap();
            let has_default = ct.contains(&format!("Extension=\"{ext}\""));
            let has_override = ct.contains(&format!("PartName=\"/{n}\""));
            assert!(n == "[Content_Types].xml" || has_default || has_override, "{n} has no content type");
            if ext == "xml" || ext == "rels" {
                let s = std::str::from_utf8(d).unwrap();
                xml::parse(s, keep_ws_pptx).unwrap_or_else(|e| panic!("{n}: {e}"));
                let mut rdr = quick_xml::Reader::from_str(s);
                loop {
                    match rdr.read_event() {
                        Ok(quick_xml::events::Event::Eof) => break,
                        Ok(_) => {}
                        Err(e) => panic!("{n}: {e}"),
                    }
                }
                for c in s.chars() {
                    assert!(crate::docx::xml_char_ok(c), "{n}: bad char");
                }
                if ext == "xml" && n.starts_with("ppt/") && !n.contains("_rels") {
                    assert!(has_override, "{n} needs an Override");
                }
            }
        }
        for (n, _) in parts.iter().filter(|(n, _)| n.ends_with(".rels")) {
            let src = if n == "_rels/.rels" { String::new() } else { n.replace("_rels/", "").trim_end_matches(".rels").to_string() };
            let (dir, _) = split_part(&src);
            let node = xml::parse(&get(n).unwrap(), keep_ws_pptx).unwrap();
            for r in node.elems() {
                if r.attr("TargetMode") == Some("External") {
                    continue;
                }
                let t = resolve_target(&dir, r.attr("Target").unwrap());
                assert!(names.contains(&t.as_str()), "{n}: missing target {t}");
            }
        }
        // ids
        let pres = xml::parse(&get("ppt/presentation.xml").unwrap(), keep_ws_pptx).unwrap();
        for s in pres.child("sldIdLst").unwrap().elems() {
            assert!(num(s, "id").unwrap() >= 256.0);
        }
        let master = xml::parse(&get("ppt/slideMasters/slideMaster1.xml").unwrap(), keep_ws_pptx).unwrap();
        for l in master.child("sldLayoutIdLst").unwrap().elems() {
            assert!(num(l, "id").unwrap() >= 2_147_483_649.0);
        }
        for (n, d) in parts.iter().filter(|(n, _)| n.starts_with("ppt/slides/slide")) {
            let node = xml::parse(std::str::from_utf8(d).unwrap(), keep_ws_pptx).unwrap();
            let mut ids = Vec::new();
            node.find_all("cNvPr", &mut ids);
            let mut seen = BTreeSet::new();
            for c in ids {
                assert!(seen.insert(c.attr("id").unwrap().to_string()), "{n}: duplicate cNvPr id");
            }
        }
    }

    #[test]
    fn writer_strips_invalid_chars_and_escapes() {
        let a = deck(json!({ "meta": { "title": "A\u{1}B" }, "slides": [{ "items": [
            { "type": "text", "x": 0, "y": 0, "w": 100, "h": 20, "paras": [{ "runs": [{ "t": "bad\u{0}\u{1b}char <&> \"q\"", "font": "F\"o<o" }] }] }
        ], "notes": "n\u{2}otes & <stuff>" }] }));
        let (bytes, _) = write_pptx(&a).unwrap();
        check_package(&bytes);
        let b = parse_pptx(&bytes).unwrap();
        assert_eq!(text_of(&b.slides[0].items[0]), "badchar <&> \"q\"");
        assert_eq!(b.slides[0].notes, "notes & <stuff>");
        assert_eq!(b.meta["title"], "AB");
    }

    #[test]
    fn writer_skips_bad_images_with_notes() {
        let a = deck(json!({ "slides": [{ "items": [
            { "type": "image", "x": 0, "y": 0, "w": 10, "h": 10, "src": "data:image/png;base64,!!!notbase64" },
            { "type": "image", "x": 0, "y": 0, "w": 10, "h": 10, "src": "data:image/svg+xml;base64,PHN2Zy8+" },
            { "type": "image", "x": 0, "y": 0, "w": 10, "h": 10, "src": "https://example.com/x.png" },
            { "type": "image", "x": 0, "y": 0, "w": 10, "h": 10, "src": data_url("image/png", PNG_1X1) },
            { "type": "image", "x": 20, "y": 0, "w": 10, "h": 10, "src": data_url("image/png", PNG_1X1) },
            { "type": "weird", "foo": 1 },
            { "type": "table", "x": 0, "y": 0, "w": 10, "h": 10, "rows": [] }
        ] }] }));
        let (bytes, notes) = write_pptx(&a).unwrap();
        check_package(&bytes);
        assert!(notes.iter().any(|n| n.contains("invalid image")), "{notes:?}");
        assert!(notes.iter().any(|n| n.contains("SVG")), "{notes:?}");
        assert!(notes.iter().any(|n| n.contains("unknown type")), "{notes:?}");
        // identical pictures share one media part
        assert_eq!(unzip(&bytes).iter().filter(|(n, _)| n.starts_with("ppt/media/")).count(), 1);
        let b = parse_pptx(&bytes).unwrap();
        assert_eq!(b.slides[0].items.iter().filter(|i| matches!(i, Item::Image(_))).count(), 2);
    }

    #[test]
    fn writer_handles_empty_and_extreme_models() {
        let (bytes, _) = write_pptx(&SlidesDoc::default()).unwrap();
        check_package(&bytes);
        assert_eq!(parse_pptx(&bytes).unwrap().slides.len(), 1);
        let a = deck(json!({ "size": { "w": 1e9, "h": -5 }, "slides": [{ "items": [
            { "type": "shape", "shape": "cloud", "x": -1e30, "y": 1e30, "w": -50, "h": 1e20, "rot": -45, "lineW": 1e9, "line": "#000" },
            { "type": "text", "x": 0, "y": 0, "w": 10, "h": 10, "paras": [{ "level": 99, "runs": [{ "t": "x", "size": 99999 }] }, { "runs": [{ "t": "y", "size": 0.01 }] }] }
        ] }] }));
        let (bytes, notes) = write_pptx(&a).unwrap();
        check_package(&bytes);
        assert!(!notes.is_empty());
        let b = parse_pptx(&bytes).unwrap();
        let Item::Shape(s) = &b.slides[0].items[0] else { panic!() };
        assert_eq!((s.shape.as_str(), s.rot), ("rect", 315.0));
        let Item::Text(t) = &b.slides[0].items[1] else { panic!() };
        assert_eq!((t.paras[0].level, t.paras[0].runs[0].size), (8, Some(4000.0)));
        assert_eq!(t.paras[1].runs[0].size, Some(1.0));
    }

    // ---------------------------------------------------------------- hostile input

    const CT: &str = r#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>"#;
    const ROOT_RELS: &str = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#;

    fn pres_xml(n: usize) -> String {
        let ids: String = (0..n).map(|i| format!("<p:sldId id=\"{}\" r:id=\"rId{}\"/>", 256 + i, i + 1)).collect();
        format!("<p:presentation {NS_DECL}><p:sldIdLst>{ids}</p:sldIdLst><p:sldSz cx=\"12192000\" cy=\"6858000\"/></p:presentation>")
    }

    fn pres_rels(n: usize, target: impl Fn(usize) -> String) -> String {
        let r: String = (0..n).map(|i| format!("<Relationship Id=\"rId{}\" Type=\"{RT}/slide\" Target=\"{}\"/>", i + 1, target(i))).collect();
        format!("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">{r}</Relationships>")
    }

    fn slide_with(tree: &str) -> String {
        format!("<p:sld {NS_DECL}><p:cSld><p:spTree>{tree}</p:spTree></p:cSld></p:sld>")
    }

    fn sp(text: &str) -> String {
        format!("<p:sp><p:nvSpPr><p:cNvPr id=\"2\" name=\"t\"/><p:cNvSpPr txBox=\"1\"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"12700\" y=\"12700\"/><a:ext cx=\"127000\" cy=\"127000\"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:p><a:r><a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp>")
    }

    #[test]
    fn malformed_inputs_are_errors_not_panics() {
        let e = parse_pptx(b"hello world").unwrap_err();
        assert!(e.starts_with("This is not a valid PowerPoint file"), "{e}");
        assert!(parse_pptx(b"").is_err());
        assert!(parse_pptx(b"PK\x03\x04garbage").unwrap_err().contains("damaged"));
        let ole = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, 0, 0];
        assert!(parse_pptx(&ole).unwrap_err().contains(".ppt"));
        // zip without presentation.xml
        let z = zip_parts(&[("[Content_Types].xml", CT.as_bytes()), ("hello.txt", b"hi")]);
        assert!(parse_pptx(&z).unwrap_err().contains("ppt/presentation.xml"));
        // a Word document
        let z = zip_parts(&[("[Content_Types].xml", CT.as_bytes()), ("word/document.xml", b"<w:document/>")]);
        assert!(parse_pptx(&z).unwrap_err().contains("Word document"));
        // broken presentation.xml
        let z = zip_parts(&[("_rels/.rels", ROOT_RELS.as_bytes()), ("ppt/presentation.xml", b"<p:presentation><p:sldIdLst><<<")]);
        assert!(parse_pptx(&z).is_err());
        // listed slides all missing
        let z = zip_parts(&[("_rels/.rels", ROOT_RELS.as_bytes()), ("ppt/presentation.xml", pres_xml(2).as_bytes())]);
        assert!(parse_pptx(&z).unwrap_err().contains("slides are missing"));
        // every slide broken
        let z = zip_parts(&[
            ("_rels/.rels", ROOT_RELS.as_bytes()),
            ("ppt/presentation.xml", pres_xml(1).as_bytes()),
            ("ppt/_rels/presentation.xml.rels", pres_rels(1, |_| "slides/slide1.xml".into()).as_bytes()),
            ("ppt/slides/slide1.xml", b"<p:sld><unclosed"),
        ]);
        let e = parse_pptx(&z).unwrap_err();
        assert!(e.contains("none of its slides"), "{e}");
    }

    #[test]
    fn broken_rels_fall_back_to_file_order_and_skip_bad_slides() {
        // Damaged presentation rels: slides found by name, in numeric order.
        let s1 = slide_with(&sp("one"));
        let s2 = slide_with(&sp("two"));
        let s10 = slide_with(&sp("ten"));
        let z = zip_parts(&[
            ("_rels/.rels", ROOT_RELS.as_bytes()),
            ("ppt/presentation.xml", pres_xml(3).as_bytes()),
            ("ppt/_rels/presentation.xml.rels", b"<not-xml"),
            ("ppt/slides/slide10.xml", s10.as_bytes()),
            ("ppt/slides/slide2.xml", s2.as_bytes()),
            ("ppt/slides/slide1.xml", s1.as_bytes()),
        ]);
        let d = parse_pptx(&z).unwrap();
        let t: Vec<String> = d.slides.iter().map(|s| text_of(&s.items[0])).collect();
        assert_eq!(t, ["one", "two", "ten"]);
        assert!(warnings(&d).iter().any(|w| w == W_SLIDE_LIST));
        assert_eq!((d.size.w, d.size.h), (960.0, 540.0));

        // One damaged slide among good ones; rels with ../ and absolute targets.
        let z = zip_parts(&[
            ("_rels/.rels", ROOT_RELS.as_bytes()),
            ("ppt/presentation.xml", pres_xml(3).as_bytes()),
            ("ppt/_rels/presentation.xml.rels", pres_rels(3, |i| ["slides/a.xml", "../ppt/slides/b.xml", "/ppt/slides/c.xml"][i].into()).as_bytes()),
            ("ppt/slides/a.xml", s1.as_bytes()),
            ("ppt/slides/b.xml", b"<p:sld><p:cSld><p:spTree><p:sp></p:cSld>&bogus;<<"),
            ("ppt/slides/c.xml", s2.as_bytes()),
        ]);
        let d = parse_pptx(&z).unwrap();
        assert!(d.slides.len() >= 2);
        assert_eq!(text_of(&d.slides[0].items[0]), "one");
        assert_eq!(text_of(&d.slides.last().unwrap().items[0]), "two");
    }

    #[test]
    fn huge_counts_are_capped() {
        // 2500 slide ids pointing to one small slide.
        let n = 2500;
        let s = slide_with(&sp("x"));
        let z = zip_parts(&[
            ("_rels/.rels", ROOT_RELS.as_bytes()),
            ("ppt/presentation.xml", pres_xml(n).as_bytes()),
            ("ppt/_rels/presentation.xml.rels", pres_rels(n, |_| "slides/slide1.xml".into()).as_bytes()),
            ("ppt/slides/slide1.xml", s.as_bytes()),
        ]);
        let d = parse_pptx(&z).unwrap();
        assert_eq!(d.slides.len(), MAX_SLIDES);
        assert!(warnings(&d).iter().any(|w| w.contains("only the first 2000")));

        // Too many shapes on one slide + very deep group nesting.
        let many: String = (0..MAX_ITEMS + 50).map(|_| sp("y")).collect();
        let mut deep = sp("deep");
        for _ in 0..300 {
            deep = format!("<p:grpSp><p:nvGrpSpPr><p:cNvPr id=\"9\" name=\"g\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{deep}</p:grpSp>");
        }
        let z = zip_parts(&[
            ("_rels/.rels", ROOT_RELS.as_bytes()),
            ("ppt/presentation.xml", pres_xml(2).as_bytes()),
            ("ppt/_rels/presentation.xml.rels", pres_rels(2, |i| format!("slides/slide{}.xml", i + 1)).as_bytes()),
            ("ppt/slides/slide1.xml", slide_with(&many).as_bytes()),
            ("ppt/slides/slide2.xml", slide_with(&deep).as_bytes()),
        ]);
        let d = parse_pptx(&z).unwrap();
        assert_eq!(d.slides[0].items.len(), MAX_ITEMS);
        assert!(warnings(&d).iter().any(|w| w == W_ITEMS));
        assert!(d.slides[1].items.is_empty());
    }

    #[test]
    fn objects_become_placeholders_with_warnings() {
        let gf = |uri: &str, inner: &str| {
            format!("<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id=\"3\" name=\"g\"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"1270000\" cy=\"635000\"/></p:xfrm><a:graphic><a:graphicData uri=\"{uri}\">{inner}</a:graphicData></a:graphic></p:graphicFrame>")
        };
        let tree = [
            gf("http://schemas.openxmlformats.org/drawingml/2006/chart", "<c:chart xmlns:c=\"x\" r:id=\"rId9\"/>"),
            gf("http://schemas.openxmlformats.org/drawingml/2006/diagram", "<dgm:relIds xmlns:dgm=\"x\"/>"),
            gf("http://schemas.openxmlformats.org/presentationml/2006/ole", "<p:oleObj/>"),
            "<p:pic><p:nvPicPr><p:cNvPr id=\"5\" name=\"v\"/><p:cNvPicPr/><p:nvPr><a:videoFile r:link=\"rId7\"/></p:nvPr></p:nvPicPr><p:blipFill><a:blip r:embed=\"rId8\"/></p:blipFill><p:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"12700\" cy=\"12700\"/></a:xfrm></p:spPr></p:pic>".to_string(),
            "<p:sp><p:nvSpPr><p:cNvPr id=\"6\" name=\"s\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"12700\" cy=\"12700\"/></a:xfrm><a:prstGeom prst=\"rect\"/><a:gradFill><a:gsLst><a:gs pos=\"0\"><a:srgbClr val=\"000000\"/></a:gs><a:gs pos=\"100000\"><a:srgbClr val=\"FFFFFF\"/></a:gs></a:gsLst></a:gradFill></p:spPr></p:sp>".to_string(),
        ]
        .concat();
        let rels = format!("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId8\" Type=\"{RT}/image\" Target=\"../media/image1.png\"/></Relationships>");
        let z = zip_parts(&[
            ("_rels/.rels", ROOT_RELS.as_bytes()),
            ("ppt/presentation.xml", pres_xml(1).as_bytes()),
            ("ppt/_rels/presentation.xml.rels", pres_rels(1, |_| "slides/slide1.xml".into()).as_bytes()),
            ("ppt/slides/slide1.xml", slide_with(&tree).as_bytes()),
            ("ppt/slides/_rels/slide1.xml.rels", rels.as_bytes()),
            ("ppt/media/image1.png", PNG_1X1),
        ]);
        let d = parse_pptx(&z).unwrap();
        let items = &d.slides[0].items;
        assert_eq!(text_of(&items[0]), "[Chart]");
        assert_eq!(text_of(&items[1]), "[SmartArt]");
        assert_eq!(text_of(&items[2]), "[Embedded object]");
        assert!(matches!(&items[3], Item::Image(_)), "video poster frame");
        let Item::Shape(g) = &items[4] else { panic!() };
        assert_eq!(g.fill.as_deref(), Some("#7F7F7F"));
        let w = warnings(&d);
        for x in [W_CHARTS, W_SMARTART, W_OLE, W_MEDIA, W_GRADIENT] {
            assert!(w.iter().any(|y| y == x), "missing {x}: {w:?}");
        }
    }

    // ---------------------------------------------------------------- colors / helpers

    #[test]
    fn color_transforms() {
        let mut map = HashMap::new();
        map.insert("tx1".to_string(), "lt1".to_string());
        let cc = ColorCtx { scheme: Rc::new(office_scheme()), map };
        let n = |s: &str| xml::parse(&format!("<a:solidFill xmlns:a=\"a\">{s}</a:solidFill>"), keep_ws_pptx).unwrap();
        assert_eq!(cc.color_in(&n("<a:schemeClr val=\"tx1\"/>"), None), Some([255, 255, 255])); // clrMap applied
        assert_eq!(cc.color_in(&n("<a:schemeClr val=\"bg2\"/>"), None), Some([0xE7, 0xE6, 0xE6]));
        assert_eq!(cc.color_in(&n("<a:srgbClr val=\"FF0000\"><a:lumMod val=\"50000\"/></a:srgbClr>"), None), Some([128, 0, 0]));
        assert_eq!(cc.color_in(&n("<a:srgbClr val=\"000000\"><a:lumMod val=\"50000\"/><a:lumOff val=\"50000\"/></a:srgbClr>"), None), Some([128, 128, 128]));
        assert_eq!(cc.color_in(&n("<a:srgbClr val=\"FF0000\"><a:tint val=\"50000\"/></a:srgbClr>"), None), Some([255, 128, 128]));
        assert_eq!(cc.color_in(&n("<a:srgbClr val=\"FF0000\"><a:shade val=\"50000\"/></a:srgbClr>"), None), Some([128, 0, 0]));
        assert_eq!(cc.color_in(&n("<a:sysClr val=\"windowText\" lastClr=\"112233\"/>"), None), Some([0x11, 0x22, 0x33]));
        assert_eq!(cc.color_in(&n("<a:prstClr val=\"black\"/>"), None), Some([0, 0, 0]));
        assert_eq!(cc.color_in(&n("<a:schemeClr val=\"phClr\"/>"), Some([1, 2, 3])), Some([1, 2, 3]));
        assert_eq!(cc.color_in(&n("<a:scrgbClr r=\"100000\" g=\"0\" b=\"0\"/>"), None), Some([255, 0, 0]));
        assert_eq!(cc.color_in(&n("<a:srgbClr val=\"zz\"/>"), None), None);
    }

    #[test]
    fn helpers() {
        assert_eq!(resolve_target("ppt/slides/", "../media/a.png"), "ppt/media/a.png");
        assert_eq!(resolve_target("ppt/slides/", "/ppt/x.xml"), "ppt/x.xml");
        assert_eq!(resolve_target("ppt/", "../../../x.xml"), "x.xml");
        assert_eq!(uri_escape("https://e.com/a b?q=\u{fc}&x=%41%"), "https://e.com/a%20b?q=%C3%BC&x=%41%25");
        let r = map_rect(&[GroupXf { off: (100.0, 100.0), ext: (200.0, 200.0), ch_off: (0.0, 0.0), ch_ext: (100.0, 100.0), rot: 0.0 }], Rect { x: 10.0, y: 20.0, w: 30.0, h: 40.0, rot: 5.0 });
        assert_eq!((r.x, r.y, r.w, r.h, r.rot), (120.0, 140.0, 60.0, 80.0, 5.0));
        let r = map_rect(&[GroupXf { off: (0.0, 0.0), ext: (100.0, 100.0), ch_off: (0.0, 0.0), ch_ext: (100.0, 100.0), rot: 90.0 }], Rect { x: 0.0, y: 0.0, w: 10.0, h: 10.0, rot: 0.0 });
        assert!((r.x - 90.0).abs() < 1e-6 && r.y.abs() < 1e-6 && r.rot == 90.0, "{r:?}"); // top-left box -> top-right
        assert_eq!(norm_rot(-90.0), 270.0);
        assert_eq!(norm_rot(720.0), 0.0);
        assert_eq!(map_transition("vortex"), Some(("fade", false)));
        assert_eq!(map_transition("sndAc"), None);
    }

    // ---------------------------------------------------------------- outlines

    #[test]
    fn markdown_outline() {
        let mut d = full_deck();
        d.slides[1].items.push(Item::Image(ImageItem { src: "data:image/png;base64,AA".into(), alt: Some("Logo".into()), ..Default::default() }));
        let md = slides_to_markdown(&d);
        assert!(md.starts_with("# Round \\<Trip\\> & Co\n\n## 1. Hello & \\<world\\>\n\n"), "{md}");
        assert!(md.contains("> Notes: First line\n> Second line\n"), "{md}");
        assert!(md.contains("## 2. Agenda (hidden)\n\n- Point one\n  - Sub *italic* [link](https://example.com/a%20b?x=1&y=2)\n    1. ~~Numbered~~\n"), "{md}");
        assert!(md.contains("![Logo](image)"), "{md}");
        assert!(md.contains("## 3. Left\n\n- Right\n"), "{md}");
        assert!(md.contains("| Name | Q1 | Q2 |\n| --- | --- | --- |\n| A\\|B | 1 | 2 |"), "{md}");
        assert!(md.contains("## 5. Part 2\n\n- Details"), "{md}");
        assert!(md.ends_with('\n') && !md.ends_with("\n\n"));
        let empty = slides_to_markdown(&deck(json!({ "slides": [{ "items": [] }] })));
        assert_eq!(empty, "## 1. Untitled slide\n");
    }

    #[test]
    fn text_outline() {
        let txt = slides_to_text(&full_deck());
        assert!(txt.starts_with("Round <Trip> & Co\n\nSlide 1: Hello & <world>\n  Notes: First line\n         Second line\n"), "{txt}");
        assert!(txt.contains("Slide 2: Agenda (hidden)\n  \u{2022} Point one\n    \u{2022} Sub italic link\n      1. Numbered\n"), "{txt}");
        assert!(txt.contains("  Name\tQ1\tQ2\n"), "{txt}");
        assert!(txt.contains("[Picture: A \"pixel\"]"), "{txt}");
    }
}
