//! LaTeX support (see model.rs "latex" CONTRACT and the LaTeX <-> word
//! mapping shared with the extension's lib/latex.js).
//!
//! * `parse_latex` — a tolerant TeX-like tokenizer (comments, groups, control
//!   words/symbols, `\verb`, verbatim-like environments) with user macro
//!   expansion (`\newcommand`, `\renewcommand`, `\def`, `\let`,
//!   `\newenvironment`, `\DeclareMathOperator`; #1..#9, capped against
//!   recursion and blow-up), a tree builder that knows the argument shapes of
//!   common commands, and a converter to the word model: headings by document
//!   class, paragraphs, formatted runs, nested lists, description lists,
//!   quotes, code (one block per line), tables (with `\multicolumn` colspans),
//!   image placeholders and captions, footnotes (a final "Notes" section),
//!   `\label`/`\ref`, `\cite` with `thebibliography`, math as a Unicode
//!   approximation in "Cambria Math", theorem-like environments, abstract,
//!   beamer frames. Unknown commands keep their argument text and are listed
//!   in ONE warning. Parsing runs on its own thread with a large stack; depth,
//!   macro expansion and output size are capped and a panic degrades to a
//!   plain-text fallback, so no input can crash the server.
//! * `write_latex` — word model -> standalone pdflatex-compatible LaTeX. The
//!   output is canonical: writing `parse_latex(write_latex(d))` again yields
//!   the same text.
//! * `compiler` / `compile_pdf` — find a TeX engine (once) and run it in a
//!   private temp directory (no shell escape, paranoid file access, hard
//!   timeout, short error summary taken from the TeX log).

use crate::docx::{merge_runs, norm_hex, push_warning, resolve_block, safe_link};
use crate::model::{Block, Run, TableCell, WordDoc};
use serde_json::{Map, Value};
use std::borrow::Cow;
use std::collections::{BTreeSet, HashMap};
use std::ffi::OsString;
use std::fs;
use std::path::{Path, PathBuf};
use std::process::{Command, ExitStatus, Stdio};
use std::rc::Rc;
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::OnceLock;
use std::time::{Duration, Instant, SystemTime, UNIX_EPOCH};

// ============================ limits ============================

/// Source characters considered (larger inputs are truncated with a warning).
const MAX_SRC_CHARS: usize = 8 * 1024 * 1024;
/// Nesting depth of groups / environments / arguments in the tree.
const MAX_DEPTH: u32 = 48;
/// User macro expansions per document.
const MAX_EXPANSIONS: usize = 50_000;
/// Tokens produced by macro expansion per document.
const MAX_EXPANDED_TOKENS: usize = 4_000_000;
/// Tokens waiting in the push-back stack.
const MAX_PENDING: usize = 1_000_000;
/// Tokens kept in one macro argument / body.
const MAX_MACRO_TOKENS: usize = 200_000;
/// Text produced by the converter (bytes) and blocks produced.
const MAX_TEXT: usize = 32 * 1024 * 1024;
const MAX_BLOCKS: usize = 200_000;
/// Stack of the parser thread.
const PARSE_STACK: usize = 256 * 1024 * 1024;
/// compile_pdf limits.
const MAX_COMPILE_SRC: usize = 2 * 1024 * 1024;
const COMPILE_TIMEOUT: Duration = Duration::from_secs(60);
const MAX_ERROR_SUMMARY: usize = 1500;

const MATH_FONT: &str = "Cambria Math";
const MONO_FONT: &str = "Courier New";

const W_IMAGES: &str = "Images are not embedded in LaTeX files: each \\includegraphics was replaced by a placeholder such as \"[Image: file]\".";
const W_DRAWINGS: &str = "TikZ/picture drawings cannot be converted and were replaced by a placeholder.";
const W_INPUT: &str = "Files pulled in with \\input or \\include were not loaded; only this file was converted.";
const W_BIBTEX: &str = "The bibliography comes from a separate .bib file that is not available; citations show their keys.";
const W_LIMIT: &str = "The document is very large; only the first part was converted.";
const W_MACRO: &str = "Some macros expand too deeply (recursive definitions?) and were left out.";
const W_FALLBACK: &str = "The LaTeX source could not be fully converted; it is shown as plain text.";

// ============================ public API ============================

/// Convert LaTeX source into the word model (sections -> headings,
/// itemize/enumerate -> lists, tabular -> tables, \textbf etc. -> runs,
/// math kept as readable text). Notes go into `meta["warnings"]`.
pub fn parse_latex(src: &str) -> WordDoc {
    let owned = src.to_string();
    let spawned = std::thread::Builder::new()
        .name("latex-parse".into())
        .stack_size(PARSE_STACK)
        .spawn(move || parse_inner(&owned));
    match spawned {
        Ok(h) => match h.join() {
            Ok(doc) => doc,
            Err(_) => fallback_doc(src),
        },
        // Could not spawn a thread (resource exhaustion): parse inline; the
        // depth caps keep the recursion well inside a default stack.
        Err(_) => std::panic::catch_unwind(|| parse_inner(src)).unwrap_or_else(|_| fallback_doc(src)),
    }
}

/// Convert the word model into a standalone LaTeX document.
pub fn write_latex(doc: &WordDoc) -> String {
    Writer::default().write(doc)
}

/// Name (full path) of an installed TeX engine usable for PDF output
/// (tectonic, pdflatex, xelatex or lualatex), detected once and cached;
/// None when none is installed. `OFFICE_BACKEND_LATEX` (a path or a program
/// name) takes precedence.
pub fn compiler() -> Option<String> {
    static FOUND: OnceLock<Option<String>> = OnceLock::new();
    FOUND
        .get_or_init(|| {
            find_engine(std::env::var_os("OFFICE_BACKEND_LATEX"), std::env::var_os("PATH"), &default_engine_dirs())
                .map(|p| p.to_string_lossy().into_owned())
        })
        .clone()
}

/// Compile LaTeX source to PDF with `compiler()`. Err carries a short,
/// human-readable error with the relevant part of the TeX log.
pub fn compile_pdf(src: &str) -> Result<Vec<u8>, String> {
    if src.len() > MAX_COMPILE_SRC {
        return Err(too_large_msg());
    }
    let Some(engine) = compiler() else {
        return Err("No LaTeX engine (tectonic, pdflatex, xelatex or lualatex) is installed.".into());
    };
    let engine = PathBuf::from(engine);
    let kind = engine_kind(&engine);
    compile_with(&engine, kind, src, COMPILE_TIMEOUT, &std::env::temp_dir())
}

fn too_large_msg() -> String {
    "The LaTeX source is larger than 2 MB and was not compiled.".into()
}

// ============================ lexer ============================

const SYN: usize = usize::MAX;

#[derive(Clone, Debug, PartialEq)]
enum Tok {
    Cmd(String),
    Open,
    Close,
    Ch(char),
    /// Whitespace; 1 = a space, 2 = a paragraph break (blank line).
    Ws(u8),
    Tab,
    Sup,
    Sub,
    Tilde,
    Param(u8),
    Dollar,
}

#[derive(Clone, Debug)]
struct Token {
    t: Tok,
    /// Start offset in the source (chars); SYN for tokens made by expansion.
    pos: usize,
}

impl Token {
    fn syn(t: Tok) -> Self {
        Token { t, pos: SYN }
    }
}

struct Macro {
    nargs: usize,
    default: Option<Vec<Token>>,
    body: Vec<Token>,
}

struct Lexer {
    src: Vec<char>,
    pos: usize,
    /// Push-back stack (last = next token).
    back: Vec<Token>,
    macros: HashMap<String, Rc<Macro>>,
    expansions: usize,
    expanded: usize,
    limit_hit: bool,
    truncated: bool,
}

fn tok_text(t: &Tok, out: &mut String) {
    match t {
        Tok::Cmd(n) => {
            out.push('\\');
            out.push_str(n);
        }
        Tok::Open => out.push('{'),
        Tok::Close => out.push('}'),
        Tok::Ch(c) => out.push(*c),
        Tok::Ws(n) => out.push_str(if *n >= 2 { "\n\n" } else { " " }),
        Tok::Tab => out.push('&'),
        Tok::Sup => out.push('^'),
        Tok::Sub => out.push('_'),
        Tok::Tilde => out.push('~'),
        Tok::Param(n) => {
            out.push('#');
            out.push((b'0' + n) as char);
        }
        Tok::Dollar => out.push('$'),
    }
}

fn toks_text(ts: &[Token]) -> String {
    let mut s = String::new();
    for t in ts {
        tok_text(&t.t, &mut s);
    }
    s
}

impl Lexer {
    fn new(src: &str) -> Self {
        let mut v: Vec<char> = Vec::with_capacity(src.len().min(MAX_SRC_CHARS));
        let mut truncated = false;
        let mut it = src.chars().peekable();
        while let Some(c) = it.next() {
            if v.len() >= MAX_SRC_CHARS {
                truncated = true;
                break;
            }
            match c {
                '\r' => {
                    if it.peek() == Some(&'\n') {
                        it.next();
                    }
                    v.push('\n');
                }
                // Reserved for the converter's internal placeholders.
                '\u{E000}'..='\u{E0FF}' | '\u{FEFF}' => {}
                _ => v.push(c),
            }
        }
        Lexer { src: v, pos: 0, back: Vec::new(), macros: HashMap::new(), expansions: 0, expanded: 0, limit_hit: false, truncated }
    }

    fn skip_blanks(&mut self) {
        while matches!(self.src.get(self.pos), Some(' ' | '\t')) {
            self.pos += 1;
        }
    }

    fn lex(&mut self) -> Option<Token> {
        let mut after_comment = false;
        loop {
            let start = self.pos;
            let c = *self.src.get(self.pos)?;
            let t = match c {
                '%' => {
                    while self.pos < self.src.len() && self.src[self.pos] != '\n' {
                        self.pos += 1;
                    }
                    if self.pos < self.src.len() {
                        self.pos += 1;
                    }
                    self.skip_blanks();
                    after_comment = true;
                    continue;
                }
                ' ' | '\t' | '\n' => {
                    // A comment eats its own newline; a blank line after it
                    // still ends the paragraph.
                    let mut nl = u32::from(after_comment);
                    while let Some(&w) = self.src.get(self.pos) {
                        if w == '\n' {
                            nl += 1;
                        } else if w != ' ' && w != '\t' {
                            break;
                        }
                        self.pos += 1;
                    }
                    if after_comment && nl <= 1 {
                        after_comment = false;
                        continue;
                    }
                    Tok::Ws(if nl >= 2 { 2 } else { 1 })
                }
                '\\' => {
                    self.pos += 1;
                    match self.src.get(self.pos).copied() {
                        None => return None,
                        Some(n) if n.is_ascii_alphabetic() => {
                            let s = self.pos;
                            while self.pos < self.src.len() && self.src[self.pos].is_ascii_alphabetic() {
                                self.pos += 1;
                            }
                            let name: String = self.src[s..self.pos].iter().collect();
                            if name != "verb" && name != "lstinline" {
                                self.skip_blanks();
                            }
                            Tok::Cmd(name)
                        }
                        Some('\n') => {
                            self.pos += 1;
                            self.skip_blanks();
                            Tok::Cmd(" ".into())
                        }
                        Some(n) => {
                            self.pos += 1;
                            Tok::Cmd(n.to_string())
                        }
                    }
                }
                '{' => {
                    self.pos += 1;
                    Tok::Open
                }
                '}' => {
                    self.pos += 1;
                    Tok::Close
                }
                '&' => {
                    self.pos += 1;
                    Tok::Tab
                }
                '^' => {
                    self.pos += 1;
                    Tok::Sup
                }
                '_' => {
                    self.pos += 1;
                    Tok::Sub
                }
                '~' => {
                    self.pos += 1;
                    Tok::Tilde
                }
                '$' => {
                    self.pos += 1;
                    Tok::Dollar
                }
                '#' => {
                    self.pos += 1;
                    match self.src.get(self.pos).copied() {
                        Some(d @ '1'..='9') => {
                            self.pos += 1;
                            Tok::Param(d as u8 - b'0')
                        }
                        _ => Tok::Ch('#'),
                    }
                }
                c if (c as u32) < 0x20 || c == '\u{7f}' => {
                    self.pos += 1;
                    continue;
                }
                c => {
                    self.pos += 1;
                    Tok::Ch(c)
                }
            };
            return Some(Token { t, pos: start });
        }
    }

    fn next_raw(&mut self) -> Option<Token> {
        if let Some(t) = self.back.pop() {
            return Some(t);
        }
        self.lex()
    }

    fn push_back(&mut self, t: Token) {
        self.back.push(t);
    }

    /// Next token with user macros expanded.
    fn next(&mut self) -> Option<Token> {
        loop {
            let tok = self.next_raw()?;
            let m = match &tok.t {
                Tok::Cmd(name) => match self.macros.get(name) {
                    Some(m) => m.clone(),
                    None => return Some(tok),
                },
                _ => return Some(tok),
            };
            self.expand(&m);
        }
    }

    /// Read the arguments of `m` and push its substituted body back.
    fn expand(&mut self, m: &Macro) {
        if self.expansions >= MAX_EXPANSIONS || self.expanded >= MAX_EXPANDED_TOKENS {
            self.limit_hit = true;
            return;
        }
        self.expansions += 1;
        let mut args: Vec<Vec<Token>> = Vec::with_capacity(m.nargs);
        for i in 0..m.nargs {
            if i == 0 {
                if let Some(d) = &m.default {
                    let a = self.read_opt_raw().unwrap_or_else(|| d.clone());
                    args.push(a);
                    continue;
                }
            }
            args.push(self.read_arg_raw());
        }
        let mut out: Vec<Token> = Vec::with_capacity(m.body.len());
        for t in &m.body {
            match t.t {
                Tok::Param(n) if n >= 1 && (n as usize) <= args.len() => {
                    out.extend(args[n as usize - 1].iter().map(|a| Token::syn(a.t.clone())));
                }
                _ => out.push(Token::syn(t.t.clone())),
            }
            if out.len() > MAX_EXPANDED_TOKENS {
                break;
            }
        }
        self.expanded += out.len();
        if self.expanded > MAX_EXPANDED_TOKENS || self.back.len() + out.len() > MAX_PENDING {
            self.limit_hit = true;
            return;
        }
        self.back.extend(out.into_iter().rev());
    }

    /// An undelimited macro argument (no expansion): a balanced group's
    /// contents or a single token.
    fn read_arg_raw(&mut self) -> Vec<Token> {
        loop {
            let Some(t) = self.next_raw() else { return Vec::new() };
            match t.t {
                Tok::Ws(1) => continue,
                Tok::Ws(_) | Tok::Close => {
                    self.push_back(t);
                    return Vec::new();
                }
                Tok::Open => return self.read_balanced_raw(),
                _ => return vec![t],
            }
        }
    }

    /// Tokens up to the `}` matching an already consumed `{`.
    fn read_balanced_raw(&mut self) -> Vec<Token> {
        let mut depth = 0u32;
        let mut out = Vec::new();
        while let Some(t) = self.next_raw() {
            match t.t {
                Tok::Open => depth += 1,
                Tok::Close => {
                    if depth == 0 {
                        return out;
                    }
                    depth -= 1;
                }
                _ => {}
            }
            if out.len() < MAX_MACRO_TOKENS {
                out.push(t);
            }
        }
        out
    }

    /// `[...]` (no expansion), or None (nothing consumed).
    fn read_opt_raw(&mut self) -> Option<Vec<Token>> {
        let mut skipped = Vec::new();
        loop {
            let t = self.next_raw()?;
            match t.t {
                Tok::Ws(1) => skipped.push(t),
                Tok::Ch('[') => break,
                _ => {
                    self.push_back(t);
                    for s in skipped.into_iter().rev() {
                        self.push_back(s);
                    }
                    return None;
                }
            }
        }
        let mut depth = 0u32;
        let mut out = Vec::new();
        while let Some(t) = self.next_raw() {
            match t.t {
                Tok::Open => depth += 1,
                Tok::Close => {
                    if depth == 0 {
                        self.push_back(t);
                        return Some(out);
                    }
                    depth -= 1;
                }
                Tok::Ch(']') if depth == 0 => return Some(out),
                _ => {}
            }
            if out.len() < MAX_MACRO_TOKENS {
                out.push(t);
            }
        }
        Some(out)
    }

    // ---- raw (verbatim) access to the source ----

    /// Make the source position point at the next unread token so the raw
    /// readers below can be used. False inside macro expansions.
    fn raw_sync(&mut self) -> bool {
        if self.back.is_empty() {
            return true;
        }
        if self.back.iter().all(|t| t.pos != SYN) {
            if let Some(p) = self.back.iter().map(|t| t.pos).min() {
                self.pos = p;
            }
            self.back.clear();
            return true;
        }
        false
    }

    fn raw_collect(&self, a: usize, b: usize) -> String {
        self.src[a.min(self.src.len())..b.min(self.src.len())].iter().collect()
    }

    /// `{...}` read verbatim (backslash escapes do not count as braces).
    fn raw_braced(&mut self) -> Option<String> {
        let save = self.pos;
        while matches!(self.src.get(self.pos), Some(' ' | '\t' | '\n')) {
            self.pos += 1;
        }
        if self.src.get(self.pos) != Some(&'{') {
            self.pos = save;
            return None;
        }
        self.pos += 1;
        let start = self.pos;
        let mut depth = 0u32;
        while self.pos < self.src.len() {
            match self.src[self.pos] {
                '\\' => {
                    self.pos = (self.pos + 2).min(self.src.len());
                    continue;
                }
                '{' => depth += 1,
                '}' => {
                    if depth == 0 {
                        let s = self.raw_collect(start, self.pos);
                        self.pos += 1;
                        return Some(s);
                    }
                    depth -= 1;
                }
                _ => {}
            }
            self.pos += 1;
        }
        Some(self.raw_collect(start, self.src.len()))
    }

    /// Optional `[...]` read verbatim.
    fn raw_bracket(&mut self) -> Option<String> {
        let save = self.pos;
        self.skip_blanks();
        if self.src.get(self.pos) != Some(&'[') {
            self.pos = save;
            return None;
        }
        self.pos += 1;
        let start = self.pos;
        let mut depth = 0u32;
        while self.pos < self.src.len() {
            match self.src[self.pos] {
                '{' => depth += 1,
                '}' => depth = depth.saturating_sub(1),
                ']' if depth == 0 => {
                    let s = self.raw_collect(start, self.pos);
                    self.pos += 1;
                    return Some(s);
                }
                _ => {}
            }
            self.pos += 1;
        }
        Some(self.raw_collect(start, self.src.len()))
    }

    /// `\verb|...|`-style delimited text (stops at end of line).
    fn raw_delimited(&mut self) -> Option<String> {
        if self.src.get(self.pos) == Some(&'*') {
            self.pos += 1;
        }
        let d = *self.src.get(self.pos)?;
        if d.is_whitespace() || d.is_ascii_alphabetic() {
            return None;
        }
        self.pos += 1;
        let start = self.pos;
        while self.pos < self.src.len() && self.src[self.pos] != d && self.src[self.pos] != '\n' {
            self.pos += 1;
        }
        let s = self.raw_collect(start, self.pos);
        if self.src.get(self.pos) == Some(&d) {
            self.pos += 1;
        }
        Some(s)
    }

    /// Everything up to `end` (consumed); the rest of the file if missing.
    fn raw_until(&mut self, end: &str) -> String {
        let pat: Vec<char> = end.chars().collect();
        let n = self.src.len();
        let mut i = self.pos;
        while i + pat.len() <= n {
            if self.src[i] == pat[0] && self.src[i..i + pat.len()] == pat[..] {
                let s = self.raw_collect(self.pos, i);
                self.pos = i + pat.len();
                return s;
            }
            i += 1;
        }
        let s = self.raw_collect(self.pos, n);
        self.pos = n;
        s
    }
}

// ============================ tree ============================

#[derive(Clone, Debug)]
enum Node {
    Text(String),
    Space,
    Par,
    Group(Vec<Node>),
    Cmd(Box<CmdNode>),
    Env(Box<EnvNode>),
    Math { display: bool, env: Option<String>, body: Vec<Node> },
    Verb(String),
    Verbatim { text: String },
    Tab,
    Sup,
    Sub,
}

#[derive(Clone, Debug)]
struct CmdNode {
    name: String,
    star: bool,
    args: Vec<Option<Vec<Node>>>,
}

#[derive(Clone, Debug)]
struct EnvNode {
    name: String,
    args: Vec<Option<Vec<Node>>>,
    body: Vec<Node>,
}

fn push_char(out: &mut Vec<Node>, c: char) {
    if let Some(Node::Text(s)) = out.last_mut() {
        s.push(c);
    } else {
        out.push(Node::Text(c.to_string()));
    }
}

fn push_space(out: &mut Vec<Node>) {
    if !matches!(out.last(), Some(Node::Space)) {
        out.push(Node::Space);
    }
}

/// Approximate source text of nodes (keys, file names, colors, URLs).
fn nodes_src(nodes: &[Node]) -> String {
    fn rec(nodes: &[Node], s: &mut String, depth: u32) {
        if depth > MAX_DEPTH + 8 {
            return;
        }
        for n in nodes {
            match n {
                Node::Text(t) => s.extend(t.chars().map(|c| if c == '\u{a0}' { '~' } else { c })),
                Node::Space => s.push(' '),
                Node::Par => s.push_str("\n\n"),
                Node::Group(g) => {
                    s.push('{');
                    rec(g, s, depth + 1);
                    s.push('}');
                }
                Node::Cmd(c) => {
                    s.push('\\');
                    s.push_str(&c.name);
                    if c.star {
                        s.push('*');
                    }
                    for a in c.args.iter().flatten() {
                        s.push('{');
                        rec(a, s, depth + 1);
                        s.push('}');
                    }
                }
                Node::Env(e) => rec(&e.body, s, depth + 1),
                Node::Math { body, .. } => {
                    s.push('$');
                    rec(body, s, depth + 1);
                    s.push('$');
                }
                Node::Verb(t) => s.push_str(t),
                Node::Verbatim { text, .. } => s.push_str(text),
                Node::Tab => s.push('&'),
                Node::Sup => s.push('^'),
                Node::Sub => s.push('_'),
            }
        }
    }
    let mut s = String::new();
    rec(nodes, &mut s, 0);
    s
}

#[derive(PartialEq, Debug)]
enum Stop {
    Eof,
    Group,
    Bracket,
    Env(String),
    Dollar,
    DDollar,
    Paren,
    BrMath,
}

struct UserEnv {
    begin: Rc<Macro>,
    end: Vec<Token>,
}

const PROTECTED: &[&str] = &[
    "begin", "end", "par", "item", "def", "gdef", "let", "newcommand", "renewcommand", "providecommand",
    "newenvironment", "renewenvironment", "verb", "documentclass", "usepackage", "\\", "[", "]", "(", ")", "maketitle",
];

const VERBATIM_ENVS: &[&str] = &[
    "verbatim", "Verbatim", "BVerbatim", "LVerbatim", "lstlisting", "minted", "comment", "alltt", "filecontents", "code",
    "spverbatim", "boxedverbatim", "sourcecode", "pyglist",
];

fn is_math_env(base: &str) -> bool {
    matches!(
        base,
        "equation" | "align" | "alignat" | "flalign" | "gather" | "multline" | "eqnarray" | "displaymath" | "math"
            | "dmath" | "dgroup" | "IEEEeqnarray"
    )
}

struct Builder {
    lx: Lexer,
    depth: u32,
    open_groups: u32,
    flat_groups: u32,
    envs: Vec<String>,
    pending_end: Option<String>,
    done: bool,
    beamer: bool,
    math: u32,
    user_envs: HashMap<String, Rc<UserEnv>>,
    thm_envs: std::collections::HashSet<String>,
}

impl Builder {
    fn new(src: &str) -> Self {
        Builder {
            lx: Lexer::new(src),
            depth: 0,
            open_groups: 0,
            flat_groups: 0,
            envs: Vec::new(),
            pending_end: None,
            done: false,
            beamer: false,
            math: 0,
            user_envs: HashMap::new(),
            thm_envs: Default::default(),
        }
    }

    fn build(&mut self) -> Vec<Node> {
        let mut out = Vec::new();
        // Stray closing braces at top level are ignored by seq, so one call
        // reads the whole document.
        while !self.done {
            let part = self.seq(&Stop::Eof);
            let empty = part.is_empty();
            out.extend(part);
            if self.pending_end.take().is_none() && (empty || self.lx.back.is_empty() && self.lx.pos >= self.lx.src.len()) {
                break;
            }
        }
        out
    }

    fn seq(&mut self, stop: &Stop) -> Vec<Node> {
        let mut out: Vec<Node> = Vec::new();
        loop {
            if self.done || self.pending_end.is_some() {
                break;
            }
            let Some(tok) = self.lx.next() else { break };
            match tok.t {
                Tok::Close => {
                    if self.flat_groups > 0 {
                        self.flat_groups -= 1;
                        continue;
                    }
                    if *stop == Stop::Group {
                        break;
                    }
                    if self.open_groups > 0 {
                        self.lx.push_back(tok);
                        break;
                    }
                }
                Tok::Open => {
                    if self.depth >= MAX_DEPTH {
                        self.flat_groups += 1;
                        continue;
                    }
                    let g = self.group();
                    out.push(Node::Group(g));
                }
                Tok::Ch(']') if *stop == Stop::Bracket => break,
                Tok::Dollar => {
                    match stop {
                        Stop::Dollar => break,
                        Stop::DDollar => {
                            if let Some(t2) = self.lx.next() {
                                if t2.t != Tok::Dollar {
                                    self.lx.push_back(t2);
                                }
                            }
                            break;
                        }
                        _ => {}
                    }
                    let display = match self.lx.next() {
                        Some(t) if t.t == Tok::Dollar => true,
                        Some(t) => {
                            self.lx.push_back(t);
                            false
                        }
                        None => false,
                    };
                    let body = self.math_seq(if display { Stop::DDollar } else { Stop::Dollar });
                    out.push(Node::Math { display, env: None, body });
                }
                Tok::Ws(n) => {
                    if n >= 2 {
                        if matches!(stop, Stop::Dollar | Stop::DDollar | Stop::Paren | Stop::BrMath) {
                            self.lx.push_back(tok);
                            break;
                        }
                        if self.math > 0 {
                            push_space(&mut out);
                        } else {
                            out.push(Node::Par);
                        }
                    } else {
                        push_space(&mut out);
                    }
                }
                Tok::Tab => out.push(Node::Tab),
                Tok::Sup => out.push(Node::Sup),
                Tok::Sub => out.push(Node::Sub),
                Tok::Tilde => push_char(&mut out, '\u{a0}'),
                Tok::Ch(c) => push_char(&mut out, c),
                Tok::Param(n) => {
                    push_char(&mut out, '#');
                    push_char(&mut out, (b'0' + n) as char);
                }
                Tok::Cmd(name) => {
                    if self.command(name, stop, &mut out) {
                        break;
                    }
                }
            }
        }
        out
    }

    fn group(&mut self) -> Vec<Node> {
        self.depth += 1;
        self.open_groups += 1;
        let g = self.seq(&Stop::Group);
        self.open_groups -= 1;
        self.depth -= 1;
        g
    }

    fn math_seq(&mut self, stop: Stop) -> Vec<Node> {
        self.math += 1;
        self.depth += 1;
        let b = self.seq(&stop);
        self.depth -= 1;
        self.math -= 1;
        b
    }

    fn read_env_name(&mut self) -> String {
        let toks = self.lx.read_arg_raw();
        let mut s = String::new();
        for t in &toks {
            match &t.t {
                Tok::Ch(c) => s.push(*c),
                Tok::Cmd(_) | Tok::Ws(_) => {}
                other => tok_text(other, &mut s),
            }
        }
        s.trim().chars().take(64).collect()
    }

    /// Returns true when the current sequence must stop.
    fn command(&mut self, name: String, stop: &Stop, out: &mut Vec<Node>) -> bool {
        match name.as_str() {
            "begin" => {
                let env = self.read_env_name();
                if env.is_empty() {
                    return false;
                }
                self.begin_env(env, out)
            }
            "end" => {
                let env = self.read_env_name();
                if let Stop::Env(e) = stop {
                    if *e == env {
                        return true;
                    }
                }
                if env == "document" {
                    self.done = true;
                    return true;
                }
                if let Some(ue) = self.user_envs.get(&env).cloned() {
                    self.push_tokens(&ue.end);
                    return false;
                }
                if self.envs.iter().any(|e| *e == env) {
                    self.pending_end = Some(env);
                    return true;
                }
                false
            }
            "[" => {
                let body = self.math_seq(Stop::BrMath);
                out.push(Node::Math { display: true, env: None, body });
                false
            }
            "]" => *stop == Stop::BrMath,
            "(" => {
                let body = self.math_seq(Stop::Paren);
                out.push(Node::Math { display: false, env: None, body });
                false
            }
            ")" => *stop == Stop::Paren,
            "verb" => {
                let t = self.verb();
                out.push(Node::Verb(t));
                false
            }
            "lstinline" | "mintinline" | "inlinecode" => {
                let t = self.inline_code(&name);
                out.push(Node::Verb(t));
                false
            }
            "newcommand" | "renewcommand" | "providecommand" | "DeclareRobustCommand" | "newrobustcmd" | "renewrobustcmd" => {
                self.define(&name);
                false
            }
            "def" | "gdef" | "edef" | "xdef" => {
                self.def();
                false
            }
            "let" => {
                self.let_();
                false
            }
            "newenvironment" | "renewenvironment" | "provideenvironment" => {
                self.newenv();
                false
            }
            "DeclareMathOperator" => {
                self.declare_op();
                false
            }
            "\\" => {
                if let Some(t) = self.lx.next() {
                    if t.t != Tok::Ch('*') {
                        self.lx.push_back(t);
                    }
                }
                let _ = self.opt_arg();
                out.push(Node::Cmd(Box::new(CmdNode { name, star: false, args: Vec::new() })));
                false
            }
            _ => {
                let n = self.generic(name);
                out.push(n);
                false
            }
        }
    }

    fn push_tokens(&mut self, toks: &[Token]) {
        if self.lx.back.len() + toks.len() > MAX_PENDING {
            self.lx.limit_hit = true;
            return;
        }
        for t in toks.iter().rev() {
            self.lx.push_back(Token::syn(t.t.clone()));
        }
    }

    fn begin_env(&mut self, env: String, out: &mut Vec<Node>) -> bool {
        let base = env.trim_end_matches('*').to_string();
        if env == "document" {
            let body = self.seq(&Stop::Env(env.clone()));
            out.push(Node::Env(Box::new(EnvNode { name: env, args: Vec::new(), body })));
            self.done = true;
            return true;
        }
        if let Some(ue) = self.user_envs.get(&env).cloned() {
            let m = ue.begin.clone();
            self.lx.expand(&m);
            return false;
        }
        if VERBATIM_ENVS.contains(&base.as_str()) {
            let text = self.verbatim_body(&env, &base);
            if base != "comment" {
                out.push(Node::Verbatim { text });
            }
            return false;
        }
        if is_math_env(&base) {
            if base == "alignat" {
                let _ = self.req_arg();
            }
            self.envs.push(env.clone());
            let body = self.math_seq(Stop::Env(env.clone()));
            self.envs.pop();
            if self.pending_end.as_deref() == Some(env.as_str()) {
                self.pending_end = None;
            }
            out.push(Node::Math { display: base != "math", env: Some(env), body });
            return false;
        }
        if self.depth >= MAX_DEPTH {
            // Too deep: the environment's content continues flat.
            return false;
        }
        if self.beamer {
            self.skip_overlay();
        }
        let spec = if self.thm_envs.contains(&base) { "[" } else { env_spec(&base) };
        let (args, _) = self.read_args(spec);
        self.envs.push(env.clone());
        self.depth += 1;
        let body = self.seq(&Stop::Env(env.clone()));
        self.depth -= 1;
        self.envs.pop();
        if self.pending_end.as_deref() == Some(env.as_str()) {
            self.pending_end = None;
        }
        out.push(Node::Env(Box::new(EnvNode { name: env, args, body })));
        false
    }

    fn verbatim_body(&mut self, env: &str, base: &str) -> String {
        let end = format!("\\end{{{env}}}");
        if self.lx.raw_sync() {
            match base {
                "lstlisting" | "Verbatim" | "BVerbatim" | "LVerbatim" | "alltt" | "sourcecode" => {
                    let _ = self.lx.raw_bracket();
                }
                "minted" => {
                    let _ = self.lx.raw_bracket();
                    let _ = self.lx.raw_braced();
                }
                "filecontents" => {
                    let _ = self.lx.raw_bracket();
                    let _ = self.lx.raw_braced();
                }
                _ => {}
            }
            return self.lx.raw_until(&end);
        }
        // Inside a macro expansion: rebuild the text from tokens.
        let mut s = String::new();
        while let Some(t) = self.lx.next_raw() {
            if t.t == Tok::Cmd("end".into()) {
                let name = self.read_env_name();
                if name == env {
                    break;
                }
                s.push_str(&format!("\\end{{{name}}}"));
                continue;
            }
            tok_text(&t.t, &mut s);
            if s.len() > MAX_TEXT {
                break;
            }
        }
        s
    }

    fn verb(&mut self) -> String {
        if self.lx.raw_sync() {
            return self.lx.raw_delimited().unwrap_or_default();
        }
        let Some(mut d) = self.lx.next_raw() else { return String::new() };
        if d.t == Tok::Ch('*') {
            match self.lx.next_raw() {
                Some(t) => d = t,
                None => return String::new(),
            }
        }
        let mut s = String::new();
        let mut n = 0;
        while let Some(t) = self.lx.next_raw() {
            if t.t == d.t || n > 10_000 {
                break;
            }
            tok_text(&t.t, &mut s);
            n += 1;
        }
        s
    }

    fn inline_code(&mut self, name: &str) -> String {
        if self.lx.raw_sync() {
            let _ = self.lx.raw_bracket();
            if name == "mintinline" {
                let _ = self.lx.raw_braced();
            }
            if let Some(s) = self.lx.raw_braced() {
                return s;
            }
            return self.lx.raw_delimited().unwrap_or_default();
        }
        nodes_src(&self.req_arg())
    }

    fn read_cs_name(&mut self) -> Option<String> {
        loop {
            let t = self.lx.next_raw()?;
            match t.t {
                Tok::Ws(_) => continue,
                Tok::Cmd(n) => return Some(n),
                Tok::Open => {
                    let inner = self.lx.read_balanced_raw();
                    return inner.into_iter().find_map(|t| if let Tok::Cmd(n) = t.t { Some(n) } else { None });
                }
                _ => {
                    self.lx.push_back(t);
                    return None;
                }
            }
        }
    }

    fn skip_star_raw(&mut self) {
        if let Some(t) = self.lx.next_raw() {
            if t.t != Tok::Ch('*') {
                self.lx.push_back(t);
            }
        }
    }

    fn nargs_default(&mut self) -> (usize, Option<Vec<Token>>) {
        let mut nargs = 0;
        let mut default = None;
        if let Some(n) = self.lx.read_opt_raw() {
            nargs = toks_text(&n).trim().parse::<usize>().unwrap_or(0).min(9);
            if let Some(d) = self.lx.read_opt_raw() {
                default = Some(d);
            }
        }
        if nargs == 0 {
            default = None;
        }
        (nargs, default)
    }

    fn define(&mut self, kind: &str) {
        self.skip_star_raw();
        let Some(name) = self.read_cs_name() else { return };
        let (nargs, default) = self.nargs_default();
        let body = self.lx.read_arg_raw();
        if PROTECTED.contains(&name.as_str()) {
            return;
        }
        if kind == "providecommand" && self.lx.macros.contains_key(&name) {
            return;
        }
        self.lx.macros.insert(name, Rc::new(Macro { nargs, default, body }));
    }

    fn def(&mut self) {
        let Some(t) = self.lx.next_raw() else { return };
        let Tok::Cmd(name) = t.t else {
            self.lx.push_back(t);
            return;
        };
        let mut nargs = 0usize;
        let mut found = false;
        for _ in 0..32 {
            let Some(t) = self.lx.next_raw() else { return };
            match t.t {
                Tok::Open => {
                    found = true;
                    break;
                }
                Tok::Param(n) => nargs = nargs.max(n as usize),
                _ => {}
            }
        }
        if !found {
            return;
        }
        let body = self.lx.read_balanced_raw();
        if PROTECTED.contains(&name.as_str()) {
            return;
        }
        self.lx.macros.insert(name, Rc::new(Macro { nargs, default: None, body }));
    }

    fn let_(&mut self) {
        let Some(name) = self.read_cs_name() else { return };
        let mut t = self.lx.next_raw();
        while let Some(tok) = &t {
            if matches!(tok.t, Tok::Ws(_) | Tok::Ch('=')) {
                t = self.lx.next_raw();
            } else {
                break;
            }
        }
        let Some(tok) = t else { return };
        if PROTECTED.contains(&name.as_str()) {
            return;
        }
        match tok.t {
            Tok::Cmd(src) => {
                if let Some(m) = self.lx.macros.get(&src).cloned() {
                    self.lx.macros.insert(name, m);
                } else if src != name {
                    self.lx.macros.insert(name, Rc::new(Macro { nargs: 0, default: None, body: vec![Token::syn(Tok::Cmd(src))] }));
                }
            }
            other => {
                self.lx.macros.insert(name, Rc::new(Macro { nargs: 0, default: None, body: vec![Token::syn(other)] }));
            }
        }
    }

    fn newenv(&mut self) {
        self.skip_star_raw();
        let name = toks_text(&self.lx.read_arg_raw()).trim().to_string();
        let (nargs, default) = self.nargs_default();
        let begin = self.lx.read_arg_raw();
        let end = self.lx.read_arg_raw();
        let base = name.trim_end_matches('*');
        if name.is_empty() || name == "document" || VERBATIM_ENVS.contains(&base) || is_math_env(base) {
            return;
        }
        self.user_envs.insert(name, Rc::new(UserEnv { begin: Rc::new(Macro { nargs, default, body: begin }), end }));
    }

    fn declare_op(&mut self) {
        self.skip_star_raw();
        let Some(name) = self.read_cs_name() else { return };
        let text = self.lx.read_arg_raw();
        if PROTECTED.contains(&name.as_str()) {
            return;
        }
        let mut body = vec![Token::syn(Tok::Cmd("operatorname".into())), Token::syn(Tok::Open)];
        body.extend(text);
        body.push(Token::syn(Tok::Close));
        self.lx.macros.insert(name, Rc::new(Macro { nargs: 0, default: None, body }));
    }

    fn generic(&mut self, name: String) -> Node {
        if self.beamer {
            self.skip_overlay();
        }
        if let Some(spec) = cmd_spec(&name) {
            let (args, star) = self.read_args(spec);
            if name == "newtheorem" && self.thm_envs.len() < 1000 {
                self.thm_envs.insert(nodes_src(arg_of(&args, 0)).trim().to_string());
            }
            if name == "documentclass" {
                let class = args.get(1).and_then(|a| a.as_deref()).map(nodes_src).unwrap_or_default();
                if class.trim() == "beamer" {
                    self.beamer = true;
                }
            }
            return Node::Cmd(Box::new(CmdNode { name, star, args }));
        }
        let mut star = false;
        if !is_symbol(&name) {
            // Unknown command: an attached star / [options] are dropped.
            if let Some(t) = self.lx.next() {
                if t.t == Tok::Ch('*') {
                    star = true;
                } else {
                    self.lx.push_back(t);
                }
            }
            if let Some(t) = self.lx.next() {
                if t.t == Tok::Ch('[') {
                    if self.depth < MAX_DEPTH {
                        self.depth += 1;
                        let _ = self.seq(&Stop::Bracket);
                        self.depth -= 1;
                    }
                } else {
                    self.lx.push_back(t);
                }
            }
        }
        Node::Cmd(Box::new(CmdNode { name, star, args: Vec::new() }))
    }

    fn read_args(&mut self, spec: &str) -> (Vec<Option<Vec<Node>>>, bool) {
        let mut args = Vec::new();
        let mut star = false;
        for c in spec.chars() {
            match c {
                '*' => {
                    if let Some(t) = self.lx.next() {
                        if t.t == Tok::Ch('*') {
                            star = true;
                        } else {
                            self.lx.push_back(t);
                        }
                    }
                }
                '[' => args.push(self.opt_arg()),
                '{' => args.push(Some(self.req_arg())),
                'R' => args.push(Some(vec![Node::Text(self.raw_arg())])),
                'B' => args.push(self.brace_opt()),
                '<' => self.skip_overlay(),
                '(' => self.skip_paren(),
                _ => {}
            }
        }
        (args, star)
    }

    fn opt_arg(&mut self) -> Option<Vec<Node>> {
        let mut skipped = Vec::new();
        loop {
            let t = self.lx.next()?;
            match t.t {
                Tok::Ws(1) => skipped.push(t),
                Tok::Ch('[') => break,
                _ => {
                    self.lx.push_back(t);
                    for s in skipped.into_iter().rev() {
                        self.lx.push_back(s);
                    }
                    return None;
                }
            }
        }
        if self.depth >= MAX_DEPTH {
            return Some(Vec::new());
        }
        self.depth += 1;
        let b = self.seq(&Stop::Bracket);
        self.depth -= 1;
        Some(b)
    }

    fn req_arg(&mut self) -> Vec<Node> {
        loop {
            let Some(t) = self.lx.next() else { return Vec::new() };
            match t.t {
                Tok::Ws(1) => continue,
                Tok::Open => {
                    if self.depth >= MAX_DEPTH {
                        self.flat_groups += 1;
                        return Vec::new();
                    }
                    return self.group();
                }
                Tok::Ws(_) | Tok::Close | Tok::Tab | Tok::Dollar => {
                    self.lx.push_back(t);
                    return Vec::new();
                }
                Tok::Ch(c) => return vec![Node::Text(c.to_string())],
                Tok::Tilde => return vec![Node::Text("\u{a0}".into())],
                Tok::Sup => return vec![Node::Sup],
                Tok::Sub => return vec![Node::Sub],
                Tok::Param(n) => return vec![Node::Text(format!("#{n}"))],
                Tok::Cmd(n) => {
                    if matches!(n.as_str(), "begin" | "end" | "\\" | "]" | ")" | "[" | "(") || self.depth >= MAX_DEPTH {
                        self.lx.push_back(Token { t: Tok::Cmd(n), pos: t.pos });
                        return Vec::new();
                    }
                    let mut v = Vec::new();
                    self.depth += 1;
                    let _ = self.command(n, &Stop::Eof, &mut v);
                    self.depth -= 1;
                    return v;
                }
            }
        }
    }

    fn brace_opt(&mut self) -> Option<Vec<Node>> {
        let mut skipped = Vec::new();
        loop {
            let t = self.lx.next()?;
            match t.t {
                Tok::Ws(1) => skipped.push(t),
                Tok::Open if self.depth < MAX_DEPTH => return Some(self.group()),
                _ => {
                    self.lx.push_back(t);
                    for s in skipped.into_iter().rev() {
                        self.lx.push_back(s);
                    }
                    return None;
                }
            }
        }
    }

    fn raw_arg(&mut self) -> String {
        if self.lx.raw_sync() {
            return self.lx.raw_braced().unwrap_or_default();
        }
        nodes_src(&self.req_arg())
    }

    fn skip_overlay(&mut self) {
        let Some(t) = self.lx.next() else { return };
        if t.t != Tok::Ch('<') {
            self.lx.push_back(t);
            return;
        }
        let mut taken = vec![t];
        while taken.len() < 48 {
            let Some(u) = self.lx.next() else { break };
            let end = u.t == Tok::Ch('>');
            let ok = matches!(&u.t, Tok::Ws(1)) || matches!(&u.t, Tok::Ch(c) if c.is_ascii_alphanumeric() || "-+|,.@() ".contains(*c));
            taken.push(u);
            if end {
                return;
            }
            if !ok {
                break;
            }
        }
        for t in taken.into_iter().rev() {
            self.lx.push_back(t);
        }
    }

    fn skip_paren(&mut self) {
        let Some(t) = self.lx.next() else { return };
        if t.t != Tok::Ch('(') {
            self.lx.push_back(t);
            return;
        }
        for _ in 0..200 {
            match self.lx.next() {
                Some(u) if u.t == Tok::Ch(')') => return,
                Some(_) => {}
                None => return,
            }
        }
    }
}

// ============================ command tables ============================

/// Argument shape of known commands: `*` optional star, `[` optional
/// argument, `{` mandatory argument, `R` raw braced argument (URLs), `B`
/// optional braced argument, `<` beamer overlay, `(` optional (...) group.
fn cmd_spec(name: &str) -> Option<&'static str> {
    Some(match name {
        "part" | "chapter" | "section" | "subsection" | "subsubsection" | "paragraph" | "subparagraph" | "addchap"
        | "addsec" | "addpart" => "*[{",
        "minisec" => "{",
        "frametitle" | "framesubtitle" => "[{",
        "title" | "author" | "date" | "subtitle" | "institute" | "titlegraphic" | "subject" | "keywords" | "publishers"
        | "dedication" => "[{",
        "item" => "[",
        "bibitem" => "[{",
        "textbf" | "textit" | "textsl" | "emph" | "underline" | "uline" | "uuline" | "uwave" | "sout" | "st" | "xout"
        | "texttt" | "textsf" | "textrm" | "textsc" | "textmd" | "textup" | "textnormal" | "textsuperscript"
        | "textsubscript" | "mbox" | "hbox" | "text" | "textbox" | "hl" | "highlight" | "alert" | "structure"
        | "centerline" | "enquote" | "ensuremath" | "fbox" | "phantom" | "hphantom" | "vphantom" | "smash" | "boxed"
        | "thanks" | "emphasize" | "textsw" | "underbar" | "inst" | "titlecap" | "MakeUppercase" | "MakeLowercase"
        | "uppercase" | "lowercase" | "textlatin" | "textgreek" | "foreignquote" | "num" | "ang" => "{",
        "textcolor" => "[{{",
        "color" | "pagecolor" => "[{",
        "colorbox" => "[{{",
        "fcolorbox" => "[{[{{",
        "href" => "R{",
        "url" | "nolinkurl" | "path" => "R",
        "hyperref" => "[{",
        "hyperlink" | "hypertarget" => "{{",
        "footnote" | "marginpar" | "todo" | "sidenote" => "[{",
        "footnotemark" => "[",
        "footnotetext" => "[{",
        "label" | "ref" | "eqref" | "pageref" | "autoref" | "cref" | "Cref" | "nameref" | "vref" | "autopageref"
        | "index" | "glossary" | "nocite" | "subref" | "labelcref" => "*{",
        "cite" | "citep" | "citet" | "parencite" | "textcite" | "autocite" | "footcite" | "citeauthor" | "citeyear"
        | "Citep" | "Citet" | "Cite" | "supercite" | "citealp" | "citealt" | "Parencite" | "Textcite" | "Autocite" => {
            "*[[{"
        }
        "caption" => "*[{",
        "captionof" => "*{[{",
        "subcaption" => "[{",
        "includegraphics" => "*[[{",
        "includesvg" | "includepdf" | "includestandalone" => "[{",
        "input" | "include" | "subfile" | "includeonly" | "InputIfFileExists" => "{",
        "bibliography" | "bibliographystyle" | "graphicspath" | "hypersetup" | "geometry" | "newgeometry" | "pagestyle"
        | "thispagestyle" | "pagenumbering" | "linespread" | "stepcounter" | "refstepcounter" | "theoremstyle"
        | "lstset" | "setstretch" | "fontfamily" | "selectlanguage" | "addvspace" | "logo" | "usebeamerfont"
        | "minted" | "setminted" | "urlstyle" | "setmainfont" | "setsansfont" | "setmonofont" | "tcbuselibrary"
        | "usetikzlibrary" | "pgfplotsset" | "tikzset" | "sisetup" | "allowdisplaybreaks" | "setlist" | "address"
        | "signature" | "opening" | "closing" | "encl" | "cc" | "ps" | "location" | "telephone" | "excludecomment"
        | "includecomment" | "listfiles" | "pdfbookmark" | "bookmark" | "nopagecolor" => "{",
        "usepackage" | "RequirePackage" | "documentclass" => "[{[",
        "addbibresource" | "usetheme" | "usecolortheme" | "usefonttheme" | "useinnertheme" | "useoutertheme"
        | "usebeamercolor" | "captionsetup" | "printbibliography" | "printindex" | "makeindex" | "lipsum" | "blindtext"
        | "Blindtext" => "[{",
        "setbeamertemplate" => "{[[",
        "setbeamercolor" | "setbeamerfont" | "setbeamersize" | "setbeamercovered" => "*{{",
        "setlength" | "addtolength" | "setcounter" | "addtocounter" | "numberwithin" | "fontsize" | "settowidth"
        | "DeclareUnicodeCharacter" | "newfontfamily" | "renewbibmacro" | "DeclareFieldFormat" | "texorpdfstring"
        | "AtBeginEnvironment" | "AtEndEnvironment" | "patchcmd" => "{{",
        "newcounter" => "{[",
        "newtheorem" => "*{[{[",
        "definecolor" => "[{{{",
        "colorlet" => "[{{",
        "foreignlanguage" => "[{{",
        "addcontentsline" => "{{{",
        "vspace" | "hspace" => "*{",
        "rule" => "[{{",
        "raisebox" => "{[[{",
        "parbox" => "[[[{{",
        "makebox" | "framebox" => "[[{",
        "resizebox" => "*{{{",
        "scalebox" => "{[{",
        "rotatebox" => "[{{",
        "adjustbox" => "{{",
        "multicolumn" => "{{{",
        "multirow" => "[{[{[{",
        "cline" | "arrayrulecolor" | "hhline" => "{",
        "cmidrule" => "[({",
        "toprule" | "midrule" | "bottomrule" | "addlinespace" => "[",
        "rowcolor" => "[{[",
        "cellcolor" => "[{",
        "newcolumntype" => "{[{",
        "linebreak" | "nolinebreak" | "pagebreak" | "nopagebreak" | "enlargethispage" => "[",
        "newblock" => "",
        "frac" | "dfrac" | "tfrac" | "binom" | "dbinom" | "tbinom" | "stackrel" | "overset" | "underset" => {
            "{{"
        }
        "cfrac" => "[{{",
        "sqrt" => "[{",
        "mathbf" | "mathrm" | "mathit" | "mathsf" | "mathtt" | "mathcal" | "mathbb" | "mathfrak" | "mathscr"
        | "boldsymbol" | "bm" | "mathnormal" | "pmb" | "hat" | "widehat" | "bar" | "overline" | "vec" | "dot" | "ddot"
        | "dddot" | "tilde" | "widetilde" | "acute" | "grave" | "breve" | "check" | "mathring" | "overrightarrow"
        | "overleftarrow" | "underbrace" | "overbrace" | "pmod" | "substack" | "intertext" | "mathop" | "mathrel"
        | "mathbin" | "mathord" | "cancel" | "bcancel" | "xcancel" | "underleftarrow" | "underrightarrow"
        | "overleftrightarrow" => "{",
        "operatorname" | "tag" => "*{",
        "xrightarrow" | "xleftarrow" | "xLeftarrow" | "xRightarrow" | "xleftrightarrow" => "[{",
        "'" | "`" | "^" | "\"" | "~" | "=" | "." | "u" | "v" | "H" | "c" | "k" | "r" | "d" | "b" | "t" => "{",
        "only" | "uncover" | "visible" | "invisible" | "alt" | "temporal" | "onslide" | "action" => {
            if name == "alt" {
                "{{"
            } else if name == "onslide" {
                ""
            } else {
                "{"
            }
        }
        "SI" => "[{[{",
        "si" | "unit" => "[{",
        "qty" => "[{{",
        "lettrine" => "[{{",
        "textls" => "[{",
        "newpage" | "clearpage" | "cleardoublepage" | "maketitle" | "titlepage" | "tableofcontents" | "listoffigures"
        | "listoftables" | "appendix" | "frontmatter" | "mainmatter" | "backmatter" | "noindent" | "indent"
        | "centering" | "raggedright" | "raggedleft" | "par" | "smallskip" | "medskip" | "bigskip" | "vfill" | "hfill"
        | "hfil" | "vfil" | "null" | "relax" | "protect" | "makeatletter" | "makeatother" | "normalsize" | "small"
        | "footnotesize" | "scriptsize" | "tiny" | "large" | "Large" | "LARGE" | "huge" | "Huge" | "bf" | "bfseries"
        | "it" | "itshape" | "sl" | "slshape" | "em" | "tt" | "ttfamily" | "rm" | "rmfamily" | "sf" | "sffamily"
        | "sc" | "scshape" | "normalfont" | "upshape" | "mdseries" | "pause" | "hline" | "and" | "newline"
        | "tabularnewline" | "sloppy" | "fussy" | "frenchspacing" | "nonfrenchspacing" | "onehalfspacing"
        | "doublespacing" | "singlespacing" | "selectfont" | "ignorespaces" | "unskip" | "xspace" | "endhead"
        | "endfirsthead" | "endfoot" | "endlastfoot" | "nonumber" | "notag" | "displaystyle" | "textstyle"
        | "scriptstyle" | "scriptscriptstyle" | "limits" | "nolimits" | "left" | "right" | "middle" | "big" | "Big"
        | "bigg" | "Bigg" | "bigl" | "bigr" | "Bigl" | "Bigr" | "biggl" | "biggr" | "Biggl" | "Biggr" | "bigm"
        | "Bigm" | "hrule" | "hrulefill" | "dotfill" | "strut" | "leavevmode" | "break" | "allowbreak" | "cr"
        | "crcr" | "long" | "global" | "outer" | "expandafter" | "noexpand" | "ifdefined" | "fi" | "else"
        | "titlepageframe" | "qedhere" | "not" | "hfuzz" | "vfuzz" | "today" | "nobreak" | "displaybreak"
        | "printglossaries" | "makeglossaries" | "nocorr" | "boldmath" | "unboldmath" | "mathstrut" | "quad"
        | "qquad" | "enspace" | "thinspace" | "negthinspace" | "space" | "nobreakspace" | "slash" | "maketitlepage"
        | "flushbottom" | "raggedbottom" | "twocolumn" | "onecolumn" | "clearpage*" | "bmod" | "mod" | "S" | "P"
        | "inserttitle" | "insertauthor" | "insertdate" | "insertframenumber" | "inserttotalframenumber"
        | "insertsection" | "insertsubsection" | "framebreak" | "ttdefault" | "mdseriesdefault" => "",
        _ => return None,
    })
}

/// Argument shape of environments (see cmd_spec).
fn env_spec(base: &str) -> &'static str {
    match base {
        "tabular" | "array" | "longtable" | "supertabular" | "xtabular" | "mpsupertabular" => "[{",
        "tabularx" | "tabulary" | "tabu" | "longtabu" | "tblr" | "longtblr" => "{[{",
        "minipage" => "[[[{",
        "figure" | "table" | "sidewaysfigure" | "sidewaystable" | "marginfigure" | "margintable" | "SCfigure" => "[",
        "wrapfigure" | "wraptable" => "[{[{",
        "subfigure" | "subtable" => "[[[{",
        "thebibliography" | "multicols" | "letter" | "otherlanguage" | "spacing" | "adjustbox" | "lrbox"
        | "changemargin" | "addmargin" => "{",
        "column" => "[{",
        "columns" | "itemize" | "enumerate" | "description" | "proof" | "tcolorbox" | "mdframed" | "tikzpicture"
        | "compactitem" | "compactenum" | "compactdesc" | "asparaenum" | "inparaenum" | "displayquote" => "[",
        "block" | "alertblock" | "exampleblock" => "{",
        "frame" => "[BB",
        "list" => "{{",
        "picture" => "(",
        _ => {
            if default_theorem(base).is_some() {
                "["
            } else {
                ""
            }
        }
    }
}

/// Text-mode symbols (also accepted inside math via \text{...}).
fn text_sym(name: &str) -> Option<&'static str> {
    Some(match name {
        "ldots" | "dots" | "textellipsis" | "dotsc" => "…",
        "LaTeX" => "LaTeX",
        "LaTeXe" => "LaTeX2e",
        "TeX" => "TeX",
        "BibTeX" => "BibTeX",
        "XeTeX" => "XeTeX",
        "XeLaTeX" => "XeLaTeX",
        "LuaTeX" => "LuaTeX",
        "LuaLaTeX" => "LuaLaTeX",
        "pdfLaTeX" => "pdfLaTeX",
        "pdfTeX" => "pdfTeX",
        "AmS" => "AMS",
        "textbackslash" => "\\",
        "textasciitilde" => "~",
        "textasciicircum" => "^",
        "textless" => "<",
        "textgreater" => ">",
        "textbar" => "|",
        "textquotedbl" => "\"",
        "textquotesingle" => "'",
        "textasciigrave" => "`",
        "textunderscore" => "_",
        "textbraceleft" => "{",
        "textbraceright" => "}",
        "textdollar" => "$",
        "textpercent" => "%",
        "textampersand" => "&",
        "textnumbersign" => "#",
        "pounds" | "textsterling" => "£",
        "texteuro" | "euro" | "EUR" => "€",
        "textyen" => "¥",
        "textcent" => "¢",
        "copyright" | "textcopyright" => "©",
        "textregistered" => "®",
        "texttrademark" | "trademark" => "™",
        "textdegree" | "degree" => "°",
        "textcelsius" | "celsius" => "℃",
        "ohm" => "Ω",
        "micro" => "µ",
        "S" | "textsection" => "§",
        "P" | "textparagraph" | "textpilcrow" => "¶",
        "dag" | "textdagger" => "†",
        "ddag" | "textdaggerdbl" => "‡",
        "textbullet" => "•",
        "textperiodcentered" => "·",
        "textendash" => "–",
        "textemdash" => "—",
        "textexclamdown" => "¡",
        "textquestiondown" => "¿",
        "guillemotleft" | "guillemetleft" | "flqq" => "«",
        "guillemotright" | "guillemetright" | "frqq" => "»",
        "guilsinglleft" | "flq" => "‹",
        "guilsinglright" | "frq" => "›",
        "quotedblbase" | "glqq" => "„",
        "quotesinglbase" | "glq" => "‚",
        "textquotedblleft" | "grqq" => "“",
        "textquotedblright" => "”",
        "textquoteleft" | "grq" => "‘",
        "textquoteright" => "’",
        "texttimes" => "×",
        "textdiv" => "÷",
        "textpm" => "±",
        "textmu" => "µ",
        "textonesuperior" => "¹",
        "texttwosuperior" => "²",
        "textthreesuperior" => "³",
        "textonequarter" => "¼",
        "textonehalf" => "½",
        "textthreequarters" => "¾",
        "textlnot" => "¬",
        "textcurrency" => "¤",
        "textbrokenbar" => "¦",
        "textasciidieresis" => "¨",
        "textasciimacron" => "¯",
        "textasciiacute" => "´",
        "textordfeminine" => "ª",
        "textordmasculine" => "º",
        "textperthousand" => "‰",
        "textvisiblespace" => "␣",
        "textnumero" => "№",
        "textrightarrow" => "→",
        "textleftarrow" => "←",
        "checkmark" | "cmark" => "✓",
        "ss" => "ß",
        "SS" => "SS",
        "ae" => "æ",
        "AE" => "Æ",
        "oe" => "œ",
        "OE" => "Œ",
        "o" => "ø",
        "O" => "Ø",
        "aa" => "å",
        "AA" => "Å",
        "l" => "ł",
        "L" => "Ł",
        "i" => "ı",
        "j" => "ȷ",
        "dh" => "ð",
        "DH" => "Ð",
        "th" => "þ",
        "TH" => "Þ",
        "ng" => "ŋ",
        "NG" => "Ŋ",
        "dj" => "đ",
        "DJ" => "Đ",
        "ij" => "ĳ",
        "IJ" => "Ĳ",
        "qed" | "qedsymbol" => "∎",
        "quad" => "\t",
        "qquad" => "\t\t",
        "enspace" => "\u{2002}",
        "thinspace" | "," => "\u{2009}",
        "space" | " " | ";" | ":" | ">" | "hfill" | "hfil" | "hss" => " ",
        "nobreakspace" => "\u{a0}",
        "slash" => "/",
        "%" => "%",
        "&" => "&",
        "#" => "#",
        "$" => "$",
        "_" => "_",
        "{" => "{",
        "}" => "}",
        "|" => "‖",
        "-" | "/" | "@" | "!" | "noindent" | "indent" => "",
        _ => return None,
    })
}

/// Math symbols: LaTeX command -> Unicode. For the reverse direction the
/// first entry of a character wins.
const MATH_SYMS: &[(&str, &str)] = &[
    ("alpha", "α"), ("beta", "β"), ("gamma", "γ"), ("delta", "δ"), ("varepsilon", "ε"), ("epsilon", "ϵ"),
    ("zeta", "ζ"), ("eta", "η"), ("theta", "θ"), ("vartheta", "ϑ"), ("iota", "ι"), ("kappa", "κ"), ("lambda", "λ"),
    ("mu", "μ"), ("nu", "ν"), ("xi", "ξ"), ("pi", "π"), ("varpi", "ϖ"), ("rho", "ρ"), ("varrho", "ϱ"),
    ("sigma", "σ"), ("varsigma", "ς"), ("tau", "τ"), ("upsilon", "υ"), ("varphi", "φ"), ("phi", "ϕ"), ("chi", "χ"),
    ("psi", "ψ"), ("omega", "ω"), ("Gamma", "Γ"), ("Delta", "Δ"), ("Theta", "Θ"), ("Lambda", "Λ"), ("Xi", "Ξ"),
    ("Pi", "Π"), ("Sigma", "Σ"), ("Upsilon", "Υ"), ("Phi", "Φ"), ("Psi", "Ψ"), ("Omega", "Ω"),
    ("sum", "∑"), ("prod", "∏"), ("coprod", "∐"), ("int", "∫"), ("iint", "∬"), ("iiint", "∭"), ("oint", "∮"),
    ("bigcup", "⋃"), ("bigcap", "⋂"), ("bigoplus", "⨁"), ("bigotimes", "⨂"), ("bigvee", "⋁"), ("bigwedge", "⋀"),
    ("infty", "∞"), ("leq", "≤"), ("le", "≤"), ("leqslant", "⩽"), ("geq", "≥"), ("ge", "≥"), ("geqslant", "⩾"),
    ("neq", "≠"), ("ne", "≠"), ("approx", "≈"), ("equiv", "≡"), ("sim", "∼"), ("simeq", "≃"), ("cong", "≅"),
    ("propto", "∝"), ("ll", "≪"), ("gg", "≫"), ("prec", "≺"), ("succ", "≻"), ("preceq", "⪯"), ("succeq", "⪰"),
    ("lesssim", "≲"), ("gtrsim", "≳"), ("doteq", "≐"), ("asymp", "≍"), ("triangleq", "≜"), ("nleq", "≰"),
    ("ngeq", "≱"), ("ncong", "≇"), ("nsim", "≁"), ("lessgtr", "≶"),
    ("times", "×"), ("div", "÷"), ("cdot", "·"), ("pm", "±"), ("mp", "∓"), ("ast", "∗"), ("star", "⋆"),
    ("circ", "∘"), ("bullet", "•"), ("oplus", "⊕"), ("ominus", "⊖"), ("otimes", "⊗"), ("odot", "⊙"), ("cup", "∪"),
    ("cap", "∩"), ("setminus", "∖"), ("wedge", "∧"), ("land", "∧"), ("vee", "∨"), ("lor", "∨"), ("neg", "¬"),
    ("lnot", "¬"), ("sqcup", "⊔"), ("sqcap", "⊓"), ("uplus", "⊎"), ("amalg", "⨿"), ("dagger", "†"), ("ddagger", "‡"),
    ("rightarrow", "→"), ("to", "→"), ("leftarrow", "←"), ("gets", "←"), ("leftrightarrow", "↔"),
    ("Rightarrow", "⇒"), ("Leftarrow", "⇐"), ("Leftrightarrow", "⇔"), ("iff", "⟺"), ("implies", "⟹"),
    ("impliedby", "⟸"), ("longrightarrow", "⟶"), ("longleftarrow", "⟵"), ("Longrightarrow", "⟹"),
    ("Longleftarrow", "⟸"), ("longleftrightarrow", "⟷"), ("Longleftrightarrow", "⟺"), ("mapsto", "↦"),
    ("longmapsto", "⟼"), ("uparrow", "↑"), ("downarrow", "↓"), ("updownarrow", "↕"), ("Uparrow", "⇑"),
    ("Downarrow", "⇓"), ("nearrow", "↗"), ("searrow", "↘"), ("nwarrow", "↖"), ("swarrow", "↙"),
    ("hookrightarrow", "↪"), ("hookleftarrow", "↩"), ("rightharpoonup", "⇀"), ("leftharpoonup", "↼"),
    ("rightleftharpoons", "⇌"), ("leftrightarrows", "⇆"),
    ("in", "∈"), ("notin", "∉"), ("ni", "∋"), ("subset", "⊂"), ("subseteq", "⊆"), ("subsetneq", "⊊"),
    ("nsubseteq", "⊈"), ("supset", "⊃"), ("supseteq", "⊇"), ("supsetneq", "⊋"), ("sqsubseteq", "⊑"),
    ("sqsupseteq", "⊒"), ("emptyset", "∅"), ("varnothing", "∅"), ("forall", "∀"), ("exists", "∃"),
    ("nexists", "∄"), ("partial", "∂"), ("nabla", "∇"), ("square", "□"), ("Box", "□"), ("blacksquare", "∎"),
    ("Diamond", "◇"), ("diamond", "⋄"), ("triangle", "△"), ("triangleleft", "◁"), ("triangleright", "▷"),
    ("lhd", "⊲"), ("rhd", "⊳"), ("angle", "∠"), ("measuredangle", "∡"), ("perp", "⊥"), ("bot", "⊥"), ("top", "⊤"),
    ("parallel", "∥"), ("mid", "∣"), ("nmid", "∤"), ("vdash", "⊢"), ("dashv", "⊣"), ("models", "⊨"), ("vDash", "⊨"),
    ("ell", "ℓ"), ("hbar", "ℏ"), ("hslash", "ℏ"), ("Re", "ℜ"), ("Im", "ℑ"), ("aleph", "ℵ"), ("beth", "ℶ"),
    ("wp", "℘"), ("imath", "ı"), ("jmath", "ȷ"), ("langle", "⟨"), ("rangle", "⟩"), ("lfloor", "⌊"), ("rfloor", "⌋"),
    ("lceil", "⌈"), ("rceil", "⌉"), ("Vert", "‖"), ("lVert", "‖"), ("rVert", "‖"), ("ldots", "…"),
    ("cdots", "⋯"), ("dotsb", "⋯"), ("dotsm", "⋯"), ("vdots", "⋮"), ("ddots", "⋱"), ("prime", "′"),
    ("therefore", "∴"), ("because", "∵"), ("checkmark", "✓"), ("surd", "√"), ("smile", "⌣"), ("frown", "⌢"),
    ("flat", "♭"), ("natural", "♮"), ("sharp", "♯"), ("clubsuit", "♣"), ("diamondsuit", "♢"), ("heartsuit", "♡"),
    ("spadesuit", "♠"), ("complement", "∁"), ("mho", "℧"), ("bigstar", "★"), ("lozenge", "◊"),
    ("varpropto", "∝"), ("bowtie", "⋈"), ("dots", "…"), ("lvert", "|"), ("rvert", "|"), ("vert", "|"),
    ("backslash", "\\"), ("colon", ":"), ("lbrace", "{"), ("rbrace", "}"), ("lbrack", "["), ("rbrack", "]"),
    ("{", "{"), ("}", "}"), ("|", "‖"), ("#", "#"), ("$", "$"), ("%", "%"), ("&", "&"), ("_", "_"),
    ("degree", "°"), ("qed", "∎"),
];

/// Upright operator names (\sin x -> "sin x").
const FUNCS: &[&str] = &[
    "arccos", "arcsin", "arctan", "arg", "cos", "cosh", "cot", "coth", "csc", "deg", "det", "dim", "exp", "gcd", "hom",
    "inf", "ker", "lg", "lim", "liminf", "limsup", "ln", "log", "max", "min", "Pr", "sec", "sin", "sinh", "sup",
    "tan", "tanh",
];

fn math_sym(name: &str) -> Option<&'static str> {
    MATH_SYMS.iter().find(|(n, _)| *n == name).map(|(_, s)| *s)
}

/// Relations / arrows get spaces around them in the approximation.
fn is_rel(c: char) -> bool {
    matches!(
        c,
        '=' | '<' | '>' | '≤' | '≥' | '≠' | '≈' | '≡' | '∼' | '≃' | '≅' | '∝' | '→' | '←' | '⇒' | '⇐' | '↔' | '⇔'
            | '⟺' | '⟹' | '⟸' | '↦' | '⟶' | '⟵' | '⟷' | '⟼' | '∈' | '∉' | '∋' | '⊂' | '⊆' | '⊃' | '⊇' | '⊊' | '⊋'
            | '≪' | '≫' | '≺' | '≻' | '⊢' | '⊨' | '⩽' | '⩾' | '≲' | '≳' | '≐' | '≍' | '≜' | '≰' | '≱' | '≇' | '≁'
            | '⊈' | '⇀' | '↼' | '⇌' | '⇆' | '↪' | '↩' | '≮' | '≯' | '≢' | '⊄' | '∄'
    )
}

/// Characters next to which spaces carry no meaning in math output.
fn is_math_op(c: char) -> bool {
    is_rel(c)
        || "+-=<>/,;:()[]|!*".contains(c)
        || matches!(c, '×' | '÷' | '·' | '±' | '∓' | '∘' | '∪' | '∩' | '∧' | '∨' | '⊕' | '⊗' | '∖' | '⋅' | '∗' | '⋆' | '⟨' | '⟩' | '‖')
}

fn is_symbol(name: &str) -> bool {
    text_sym(name).is_some() || math_sym(name).is_some() || FUNCS.contains(&name)
}

const SUPS: &[(char, char)] = &[
    ('0', '⁰'), ('1', '¹'), ('2', '²'), ('3', '³'), ('4', '⁴'), ('5', '⁵'), ('6', '⁶'), ('7', '⁷'), ('8', '⁸'),
    ('9', '⁹'), ('+', '⁺'), ('-', '⁻'), ('=', '⁼'), ('(', '⁽'), (')', '⁾'), ('n', 'ⁿ'), ('i', 'ⁱ'), ('a', 'ᵃ'),
    ('b', 'ᵇ'), ('c', 'ᶜ'), ('d', 'ᵈ'), ('e', 'ᵉ'), ('f', 'ᶠ'), ('g', 'ᵍ'), ('h', 'ʰ'), ('j', 'ʲ'), ('k', 'ᵏ'),
    ('l', 'ˡ'), ('m', 'ᵐ'), ('o', 'ᵒ'), ('p', 'ᵖ'), ('r', 'ʳ'), ('s', 'ˢ'), ('t', 'ᵗ'), ('u', 'ᵘ'), ('v', 'ᵛ'),
    ('w', 'ʷ'), ('x', 'ˣ'), ('y', 'ʸ'), ('z', 'ᶻ'), ('A', 'ᴬ'), ('B', 'ᴮ'), ('D', 'ᴰ'), ('E', 'ᴱ'), ('G', 'ᴳ'),
    ('H', 'ᴴ'), ('I', 'ᴵ'), ('J', 'ᴶ'), ('K', 'ᴷ'), ('L', 'ᴸ'), ('M', 'ᴹ'), ('N', 'ᴺ'), ('O', 'ᴼ'), ('P', 'ᴾ'),
    ('R', 'ᴿ'), ('T', 'ᵀ'), ('U', 'ᵁ'), ('V', 'ⱽ'), ('W', 'ᵂ'), ('α', 'ᵅ'), ('β', 'ᵝ'), ('γ', 'ᵞ'), ('δ', 'ᵟ'),
    ('ε', 'ᵋ'), ('θ', 'ᶿ'), ('φ', 'ᵠ'), ('χ', 'ᵡ'),
];

const SUBS: &[(char, char)] = &[
    ('0', '₀'), ('1', '₁'), ('2', '₂'), ('3', '₃'), ('4', '₄'), ('5', '₅'), ('6', '₆'), ('7', '₇'), ('8', '₈'),
    ('9', '₉'), ('+', '₊'), ('-', '₋'), ('=', '₌'), ('(', '₍'), (')', '₎'), ('a', 'ₐ'), ('e', 'ₑ'), ('o', 'ₒ'),
    ('x', 'ₓ'), ('h', 'ₕ'), ('k', 'ₖ'), ('l', 'ₗ'), ('m', 'ₘ'), ('n', 'ₙ'), ('p', 'ₚ'), ('s', 'ₛ'), ('t', 'ₜ'),
    ('i', 'ᵢ'), ('r', 'ᵣ'), ('u', 'ᵤ'), ('v', 'ᵥ'), ('j', 'ⱼ'), ('β', 'ᵦ'), ('γ', 'ᵧ'), ('ρ', 'ᵨ'), ('φ', 'ᵩ'),
    ('χ', 'ᵪ'),
];

fn sup_of(c: char) -> Option<char> {
    SUPS.iter().find(|(b, _)| *b == c).map(|(_, s)| *s)
}
fn sub_of(c: char) -> Option<char> {
    SUBS.iter().find(|(b, _)| *b == c).map(|(_, s)| *s)
}
fn sup_base(c: char) -> Option<char> {
    SUPS.iter().find(|(_, s)| *s == c).map(|(b, _)| *b)
}
fn sub_base(c: char) -> Option<char> {
    SUBS.iter().find(|(_, s)| *s == c).map(|(b, _)| *b)
}

/// Accent command -> (bases, composed) tables and the combining fallback.
const ACCENTS: &[(&str, &str, &str, char)] = &[
    ("'", "aeiouyAEIOUYcnszlrgCNSZLRGı", "áéíóúýÁÉÍÓÚÝćńśźĺŕǵĆŃŚŹĹŔǴí", '\u{301}'),
    ("`", "aeiouAEIOUı", "àèìòùÀÈÌÒÙì", '\u{300}'),
    ("^", "aeiouAEIOUcghjswyCGHJSWYı", "âêîôûÂÊÎÔÛĉĝĥĵŝŵŷĈĜĤĴŜŴŶî", '\u{302}'),
    ("\"", "aeiouyAEIOUYı", "äëïöüÿÄËÏÖÜŸï", '\u{308}'),
    ("~", "anoANOiuIU", "ãñõÃÑÕĩũĨŨ", '\u{303}'),
    ("=", "aeiouAEIOU", "āēīōūĀĒĪŌŪ", '\u{304}'),
    (".", "zZcCeEgGI", "żŻċĊėĖġĠİ", '\u{307}'),
    ("u", "aAgGuUeEiIoO", "ăĂğĞŭŬĕĔĭĬŏŎ", '\u{306}'),
    ("v", "cCsSzZeErRnNdDtTlL", "čČšŠžŽěĚřŘňŇďĎťŤľĽ", '\u{30C}'),
    ("H", "oOuU", "őŐűŰ", '\u{30B}'),
    ("c", "cCsStTgGkKlLnNrR", "çÇşŞţŢģĢķĶļĻņŅŗŖ", '\u{327}'),
    ("k", "aAeEiIuU", "ąĄęĘįĮųŲ", '\u{328}'),
    ("r", "aAuU", "åÅůŮ", '\u{30A}'),
    ("d", "", "", '\u{323}'),
    ("b", "", "", '\u{331}'),
    ("t", "", "", '\u{361}'),
];

fn accent(cmd: &str, text: &str) -> String {
    let Some((_, bases, composed, comb)) = ACCENTS.iter().find(|(c, ..)| *c == cmd) else {
        return text.to_string();
    };
    let mut chars = text.chars();
    let Some(first) = chars.next() else {
        // \'{} -> the accent itself
        return match cmd {
            "'" => "´".into(),
            "`" => "`".into(),
            "^" => "^".into(),
            "\"" => "¨".into(),
            "~" => "~".into(),
            "=" => "¯".into(),
            _ => String::new(),
        };
    };
    let rest: String = chars.collect();
    if let Some(i) = bases.chars().position(|b| b == first) {
        if let Some(c) = composed.chars().nth(i) {
            return format!("{c}{rest}");
        }
    }
    format!("{first}{comb}{rest}")
}

const NAMED_COLORS: &[(&str, &str)] = &[
    ("black", "000000"), ("white", "FFFFFF"), ("red", "FF0000"), ("green", "00FF00"), ("blue", "0000FF"),
    ("cyan", "00FFFF"), ("magenta", "FF00FF"), ("yellow", "FFFF00"), ("gray", "808080"), ("grey", "808080"),
    ("darkgray", "404040"), ("lightgray", "BFBFBF"), ("brown", "BF8040"), ("lime", "BFFF00"), ("olive", "808000"),
    ("orange", "FF8000"), ("pink", "FFBFBF"), ("purple", "BF0040"), ("teal", "008080"), ("violet", "800080"),
    ("navyblue", "006EB8"), ("royalblue", "0071BC"), ("skyblue", "46C5DD"), ("forestgreen", "009B55"),
    ("junglegreen", "00A99A"), ("seagreen", "3FBC9D"), ("limegreen", "8DC73E"), ("maroon", "AF3235"),
    ("brickred", "B6321C"), ("bittersweet", "C04F17"), ("goldenrod", "FFDF42"), ("dandelion", "FDBC42"),
    ("darkorchid", "A4538A"), ("plum", "92268F"), ("orchid", "AF72B0"), ("cerulean", "00A2E3"), ("emerald", "00A99D"),
    ("apricot", "FBB982"), ("peach", "F7965A"), ("melon", "F89E7B"), ("tan", "DA9D76"), ("sepia", "671800"),
    ("mahogany", "A9341F"), ("redorange", "F26035"), ("orangered", "ED135A"), ("rubinered", "ED017D"),
    ("violetred", "EF58A0"), ("periwinkle", "7977B8"), ("cadetblue", "74729A"), ("cornflowerblue", "41B0E4"),
    ("midnightblue", "006795"), ("processblue", "00B0F0"), ("turquoise", "00B4CE"), ("aquamarine", "00B5BE"),
    ("bluegreen", "00B3B8"), ("tealblue", "00AEB3"), ("pinegreen", "008B72"), ("olivegreen", "3C8031"),
    ("yellowgreen", "98CC70"), ("springgreen", "C6DC67"), ("greenyellow", "DFE674"), ("lavender", "F49EC4"),
    ("thistle", "D883B7"), ("mulberry", "A93C93"), ("fuchsia", "8C368C"), ("royalpurple", "613F99"),
    ("blueviolet", "473992"), ("darkblue", "00008B"), ("darkred", "8B0000"), ("darkgreen", "006400"),
    ("navy", "000080"), ("gold", "FFD700"), ("silver", "C0C0C0"), ("crimson", "DC143C"), ("indigo", "4B0082"),
    ("coral", "FF7F50"), ("salmon", "FA8072"), ("khaki", "F0E68C"), ("chocolate", "D2691E"), ("steelblue", "4682B4"),
    ("tomato", "FF6347"), ("darkorange", "FF8C00"), ("lightblue", "ADD8E6"), ("lightgreen", "90EE90"),
];

fn hex_rgb(h: &str) -> Option<[f64; 3]> {
    let h = norm_hex(h)?;
    let v = u32::from_str_radix(&h, 16).ok()?;
    Some([((v >> 16) & 255) as f64, ((v >> 8) & 255) as f64, (v & 255) as f64])
}

fn rgb_hex(c: [f64; 3]) -> String {
    let b = |x: f64| x.round().clamp(0.0, 255.0) as u8;
    format!("#{:02X}{:02X}{:02X}", b(c[0]), b(c[1]), b(c[2]))
}

fn floats(s: &str) -> Vec<f64> {
    s.split([',', ' ']).filter(|p| !p.trim().is_empty()).filter_map(|p| p.trim().parse::<f64>().ok()).filter(|f| f.is_finite()).collect()
}

/// Resolve an xcolor color: explicit model values or a (mixed) name.
fn color_value(model: Option<&str>, spec: &str, user: &HashMap<String, String>) -> Option<String> {
    let spec = spec.trim();
    match model.map(str::trim) {
        Some("HTML") | Some("html") => norm_hex(spec).map(|h| format!("#{h}")),
        Some("RGB") => {
            let f = floats(spec);
            (f.len() >= 3).then(|| rgb_hex([f[0], f[1], f[2]]))
        }
        Some("rgb") => {
            let f = floats(spec);
            (f.len() >= 3).then(|| rgb_hex([f[0] * 255.0, f[1] * 255.0, f[2] * 255.0]))
        }
        Some("gray") | Some("Gray") => {
            let f = floats(spec);
            let g = *f.first()?;
            let g = if model == Some("Gray") { g * 255.0 / 15.0 } else { g * 255.0 };
            Some(rgb_hex([g, g, g]))
        }
        Some("cmyk") => {
            let f = floats(spec);
            if f.len() < 4 {
                return None;
            }
            let k = 1.0 - f[3];
            Some(rgb_hex([255.0 * (1.0 - f[0]) * k, 255.0 * (1.0 - f[1]) * k, 255.0 * (1.0 - f[2]) * k]))
        }
        Some(m) if !m.is_empty() && !m.eq_ignore_ascii_case("named") => None,
        _ => color_expr(spec, user).map(rgb_hex),
    }
}

fn color_expr(e: &str, user: &HashMap<String, String>) -> Option<[f64; 3]> {
    let lookup = |n: &str| -> Option<[f64; 3]> {
        let n = n.trim();
        if let Some(h) = user.get(n) {
            return hex_rgb(h);
        }
        let l = n.to_ascii_lowercase();
        NAMED_COLORS.iter().find(|(k, _)| *k == l).and_then(|(_, h)| hex_rgb(h))
    };
    let parts: Vec<&str> = e.split('!').collect();
    let first = parts.first()?.trim();
    let (neg, first) = match first.strip_prefix('-') {
        Some(r) => (true, r),
        None => (false, first),
    };
    let mut cur = lookup(first)?;
    let mut i = 1;
    while i < parts.len() && i < 16 {
        let p = parts[i].trim().parse::<f64>().ok()?.clamp(0.0, 100.0) / 100.0;
        let other = if i + 1 < parts.len() { lookup(parts[i + 1])? } else { [255.0; 3] };
        for k in 0..3 {
            cur[k] = cur[k] * p + other[k] * (1.0 - p);
        }
        i += 2;
    }
    if neg {
        cur = [255.0 - cur[0], 255.0 - cur[1], 255.0 - cur[2]];
    }
    Some(cur)
}

// ============================ converter ============================

#[derive(Clone, Copy, PartialEq, Debug)]
enum Class {
    Article,
    Chapters,
    Beamer,
}

#[derive(Clone, Default)]
struct St {
    run: Run,
    align: Option<&'static str>,
}

impl St {
    fn mono(&self) -> bool {
        self.run.font.as_deref() == Some(MONO_FONT)
    }
}

struct ListCtx {
    kind: &'static str,
    desc: bool,
    fresh: bool,
    n: u32,
    quote_at: u32,
}

struct Label {
    num: String,
    kind: String,
    title: String,
}

struct Thm {
    label: String,
    counter: Option<String>,
    within: Option<String>,
}

fn arg_of(args: &[Option<Vec<Node>>], i: usize) -> &[Node] {
    args.get(i).and_then(|a| a.as_deref()).unwrap_or(&[])
}
fn opt_of(args: &[Option<Vec<Node>>], i: usize) -> Option<&[Node]> {
    args.get(i).and_then(|a| a.as_deref())
}
fn arg(c: &CmdNode, i: usize) -> &[Node] {
    arg_of(&c.args, i)
}
fn opt(c: &CmdNode, i: usize) -> Option<&[Node]> {
    opt_of(&c.args, i)
}
fn last_arg(c: &CmdNode) -> &[Node] {
    c.args.iter().rev().find_map(|a| a.as_deref()).unwrap_or(&[])
}

fn plain(runs: &[Run]) -> String {
    runs.iter().map(|r| r.t.as_str()).collect()
}

fn is_blank(nodes: &[Node]) -> bool {
    nodes.iter().all(|n| match n {
        Node::Space | Node::Par => true,
        Node::Text(t) => t.trim().is_empty(),
        Node::Cmd(c) => matches!(
            c.name.as_str(),
            "hline" | "toprule" | "midrule" | "bottomrule" | "cline" | "cmidrule" | "hhline" | "addlinespace" | "rowcolor"
                | "noalign" | "nonumber" | "notag"
        ),
        _ => false,
    })
}

/// TeX input ligatures (quotes and dashes).
fn ligatures(s: &str) -> String {
    if !s.contains(['-', '`', '\'']) {
        return s.to_string();
    }
    let c: Vec<char> = s.chars().collect();
    let mut out = String::with_capacity(s.len() + 4);
    let mut i = 0;
    while i < c.len() {
        match c[i] {
            '-' if c.get(i + 1) == Some(&'-') => {
                if c.get(i + 2) == Some(&'-') {
                    out.push('—');
                    i += 3;
                } else {
                    out.push('–');
                    i += 2;
                }
                continue;
            }
            '`' if c.get(i + 1) == Some(&'`') => {
                out.push('“');
                i += 2;
                continue;
            }
            '\'' if c.get(i + 1) == Some(&'\'') => {
                out.push('”');
                i += 2;
                continue;
            }
            '!' if c.get(i + 1) == Some(&'`') => {
                out.push('¡');
                i += 2;
                continue;
            }
            '?' if c.get(i + 1) == Some(&'`') => {
                out.push('¿');
                i += 2;
                continue;
            }
            '`' => out.push('‘'),
            '\'' => out.push('’'),
            ch => out.push(ch),
        }
        i += 1;
    }
    out
}

/// Canonical whitespace of a paragraph (shared by reader and writer so the
/// round trip is stable): no leading/trailing spaces or line breaks, single
/// spaces, no spaces around line breaks; adjacent equal runs merged.
fn normalize_runs(runs: Vec<Run>) -> Vec<Run> {
    let mut out: Vec<Run> = Vec::with_capacity(runs.len());
    let mut prev: Option<char> = None;
    for mut r in runs {
        let mut t = String::with_capacity(r.t.len());
        for c in r.t.chars() {
            let c = match c {
                '\r' | '\u{2028}' | '\u{2029}' => '\n',
                '\u{b}' | '\u{c}' => ' ',
                c => c,
            };
            match c {
                ' ' => {
                    if matches!(prev, None | Some(' ') | Some('\n')) {
                        continue;
                    }
                }
                '\n' => {
                    if prev.is_none() {
                        continue;
                    }
                    while t.ends_with(' ') {
                        t.pop();
                    }
                    if t.is_empty() {
                        while let Some(last) = out.last_mut() {
                            while last.t.ends_with(' ') {
                                last.t.pop();
                            }
                            if last.t.is_empty() {
                                out.pop();
                            } else {
                                break;
                            }
                        }
                    }
                }
                c if (c as u32) < 0x20 && c != '\t' => continue,
                _ => {}
            }
            t.push(c);
            prev = Some(c);
        }
        r.t = t;
        if !r.t.is_empty() {
            out.push(r);
        }
    }
    while let Some(last) = out.last_mut() {
        let keep = last.t.trim_end_matches([' ', '\n']).len();
        last.t.truncate(keep);
        if last.t.is_empty() {
            out.pop();
        } else {
            break;
        }
    }
    merge_runs(out)
}

fn heading_runs(runs: Vec<Run>) -> Vec<Run> {
    normalize_runs(runs.into_iter().map(|mut r| {
        r.t = r.t.replace('\n', " ");
        r
    }).collect())
}

fn heading_block(level: u32, runs: Vec<Run>, align: Option<String>) -> Block {
    match level {
        0 | 1 => Block::H1 { runs, align },
        2 => Block::H2 { runs, align },
        3 => Block::H3 { runs, align },
        4 => Block::H4 { runs, align },
        5 => Block::H5 { runs, align },
        _ => Block::H6 { runs, align },
    }
}

fn runs_mut(b: &mut Block) -> Option<&mut Vec<Run>> {
    match b {
        Block::P { runs, .. } | Block::H1 { runs, .. } | Block::H2 { runs, .. } | Block::H3 { runs, .. }
        | Block::H4 { runs, .. } | Block::H5 { runs, .. } | Block::H6 { runs, .. } | Block::Quote { runs, .. }
        | Block::Code { runs, .. } | Block::List { runs, .. } => Some(runs),
        _ => None,
    }
}

fn roman(mut n: u32) -> String {
    const T: &[(u32, &str)] = &[(1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"), (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")];
    let mut s = String::new();
    for (v, r) in T {
        while n >= *v {
            s.push_str(r);
            n -= v;
        }
    }
    s
}

fn letter(n: u32) -> String {
    if n == 0 {
        return "0".into();
    }
    let mut n = n;
    let mut s = Vec::new();
    while n > 0 {
        n -= 1;
        s.push((b'A' + (n % 26) as u8) as char);
        n /= 26;
    }
    s.iter().rev().collect()
}

fn today_string() -> String {
    const MONTHS: [&str; 12] = [
        "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December",
    ];
    let secs = SystemTime::now().duration_since(UNIX_EPOCH).map(|d| d.as_secs()).unwrap_or(0);
    let z = (secs / 86_400) as i64 + 719_468;
    let era = if z >= 0 { z } else { z - 146_096 } / 146_097;
    let doe = (z - era * 146_097) as u64;
    let yoe = (doe - doe / 1460 + doe / 36_524 - doe / 146_096) / 365;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = doy - (153 * mp + 2) / 5 + 1;
    let m = if mp < 10 { mp + 3 } else { mp - 9 };
    let y = yoe as i64 + era * 400 + i64::from(m <= 2);
    format!("{} {}, {}", MONTHS[(m - 1) as usize], d, y)
}

fn default_theorem(name: &str) -> Option<(&'static str, bool)> {
    Some(match name {
        "theorem" | "thm" => ("Theorem", true),
        "lemma" | "lem" => ("Lemma", true),
        "proposition" | "prop" => ("Proposition", true),
        "corollary" | "cor" => ("Corollary", true),
        "definition" | "defn" | "dfn" | "definitionx" => ("Definition", true),
        "example" | "ex" | "exmp" => ("Example", true),
        "remark" | "rem" => ("Remark", true),
        "note" => ("Note", true),
        "exercise" => ("Exercise", true),
        "problem" | "prob" => ("Problem", true),
        "question" => ("Question", true),
        "claim" => ("Claim", true),
        "conjecture" | "conj" => ("Conjecture", true),
        "fact" => ("Fact", true),
        "observation" => ("Observation", true),
        "assumption" => ("Assumption", true),
        "notation" => ("Notation", true),
        "hypothesis" => ("Hypothesis", true),
        "solution" | "sol" => ("Solution", false),
        "answer" => ("Answer", false),
        _ => return None,
    })
}

const PLACE_OPEN: char = '\u{E000}';
const PLACE_CLOSE: char = '\u{E001}';
const PLACE_SEP: char = '\u{E002}';
const PLACE_SUB: char = '\u{E003}';

fn placeholder(op: &str, arg: &str) -> String {
    format!("{PLACE_OPEN}{op}{PLACE_SEP}{arg}{PLACE_CLOSE}")
}

/// Unescape a raw URL argument (\% \# \& \_ \~ -> the character).
fn unescape_url(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    let mut it = s.trim().chars().peekable();
    while let Some(c) = it.next() {
        if c == '\\' {
            if let Some(&n) = it.peek() {
                if !n.is_ascii_alphabetic() {
                    out.push(n);
                    it.next();
                    continue;
                }
            }
        }
        if c == '\n' || c == '\t' {
            continue;
        }
        out.push(c);
    }
    out
}

fn unit_symbol(name: &str) -> Option<&'static str> {
    Some(match name {
        "meter" | "metre" => "m",
        "second" => "s",
        "kilogram" => "kg",
        "gram" => "g",
        "kilo" => "k",
        "milli" => "m",
        "micro" => "µ",
        "nano" => "n",
        "centi" => "c",
        "mega" => "M",
        "giga" => "G",
        "newton" => "N",
        "joule" => "J",
        "watt" => "W",
        "volt" => "V",
        "ampere" => "A",
        "ohm" => "Ω",
        "hertz" => "Hz",
        "kelvin" => "K",
        "celsius" | "degreeCelsius" => "°C",
        "per" => "/",
        "squared" => "²",
        "cubed" => "³",
        "percent" => "%",
        "mole" => "mol",
        "pascal" => "Pa",
        "coulomb" => "C",
        "litre" | "liter" => "L",
        "minute" => "min",
        "hour" => "h",
        "day" => "d",
        "byte" => "B",
        "bit" => "bit",
        "degree" => "°",
        _ => return None,
    })
}

struct Conv {
    class: Class,
    out: Vec<Block>,
    para: Vec<Run>,
    lists: Vec<ListCtx>,
    quote: u32,
    footnotes: Vec<(u32, Vec<Run>)>,
    marks: Vec<u32>,
    labels: HashMap<String, Label>,
    target: (String, String),
    target_title: String,
    sec: [u32; 7],
    appendix: bool,
    fig: u32,
    tab: u32,
    eq: u32,
    subfig: u32,
    float: Option<&'static str>,
    theorems: HashMap<String, Thm>,
    thm_counts: HashMap<String, (u32, String)>,
    bib: HashMap<String, String>,
    bib_n: u32,
    colors: HashMap<String, String>,
    unknown: Vec<String>,
    unknown_seen: std::collections::HashSet<String>,
    warnings: Vec<String>,
    title: Option<Vec<Node>>,
    author: Option<Vec<Node>>,
    date: Option<Vec<Node>>,
    subtitle: Option<Vec<Node>>,
    preamble: bool,
    text_len: usize,
    limit: bool,
    today: String,
}

impl Conv {
    fn new() -> Self {
        Conv {
            class: Class::Article,
            out: Vec::new(),
            para: Vec::new(),
            lists: Vec::new(),
            quote: 0,
            footnotes: Vec::new(),
            marks: Vec::new(),
            labels: HashMap::new(),
            target: (String::new(), String::new()),
            target_title: String::new(),
            sec: [0; 7],
            appendix: false,
            fig: 0,
            tab: 0,
            eq: 0,
            subfig: 0,
            float: None,
            theorems: HashMap::new(),
            thm_counts: HashMap::new(),
            bib: HashMap::new(),
            bib_n: 0,
            colors: HashMap::new(),
            unknown: Vec::new(),
            unknown_seen: Default::default(),
            warnings: Vec::new(),
            title: None,
            author: None,
            date: None,
            subtitle: None,
            preamble: false,
            text_len: 0,
            limit: false,
            today: today_string(),
        }
    }

    fn chapters(&self) -> bool {
        self.class == Class::Chapters
    }

    fn warn(&mut self, w: &str) {
        if !self.warnings.iter().any(|x| x == w) {
            self.warnings.push(w.to_string());
        }
    }

    fn unknown_cmd(&mut self, name: &str) {
        if self.preamble || name.trim().is_empty() {
            return;
        }
        if self.unknown_seen.len() > 10_000 || !self.unknown_seen.insert(name.to_string()) {
            return;
        }
        self.unknown.push(name.to_string());
    }

    fn hit_limit(&mut self) {
        if !self.limit {
            self.limit = true;
            self.warn(W_LIMIT);
        }
    }

    fn push_block(&mut self, b: Block) {
        if self.preamble {
            return;
        }
        if self.out.len() >= MAX_BLOCKS {
            self.hit_limit();
            return;
        }
        self.out.push(b);
    }

    fn push_run(&mut self, t: impl Into<String>, st: &St) {
        if self.preamble || self.limit {
            return;
        }
        let t = t.into();
        if t.is_empty() {
            return;
        }
        self.text_len += t.len();
        if self.text_len > MAX_TEXT {
            self.hit_limit();
            return;
        }
        let mut r = st.run.clone();
        r.t = t;
        self.para.push(r);
    }

    fn push_runs(&mut self, runs: Vec<Run>) {
        if self.preamble || self.limit {
            return;
        }
        self.para.extend(runs);
    }

    fn space(&mut self, st: &St) {
        match self.para.last() {
            None => {}
            Some(r) if r.t.ends_with(' ') || r.t.ends_with('\n') => {}
            _ => self.push_run(" ", st),
        }
    }

    fn text(&mut self, s: &str, st: &St) {
        if self.preamble {
            return;
        }
        let t = if st.mono() { s.to_string() } else { ligatures(s) };
        self.push_run(t, st);
    }

    fn flush(&mut self, align: Option<&'static str>) {
        let runs = normalize_runs(std::mem::take(&mut self.para));
        let q = self.quote;
        let in_list = match self.lists.last_mut() {
            Some(l) if q <= l.quote_at => {
                let fresh = l.fresh;
                l.fresh = false;
                Some((fresh, l.kind))
            }
            _ => None,
        };
        match in_list {
            Some((true, kind)) => {
                let level = (self.lists.len() - 1).min(8) as u32;
                self.push_block(Block::List { list: kind.into(), level, runs });
            }
            Some((false, _)) => {
                if !runs.is_empty() {
                    self.push_block(Block::P { runs, align: None });
                }
            }
            None => {
                if runs.is_empty() {
                    return;
                }
                if q > 0 {
                    self.push_block(Block::Quote { runs, align: None });
                } else {
                    self.push_block(Block::P { runs, align: align.map(String::from) });
                }
            }
        }
    }

    fn par(&mut self, st: &St) {
        let q = self.quote;
        if self.lists.last().is_some_and(|l| q <= l.quote_at && l.fresh) {
            self.space(st);
            return;
        }
        self.flush(st.align);
    }

    // ---- driving ----

    fn run(&mut self, nodes: &[Node]) {
        let doc = nodes.iter().position(|n| matches!(n, Node::Env(e) if e.name == "document"));
        let mut st = St::default();
        match doc {
            Some(i) => {
                self.preamble = true;
                self.walk(&nodes[..i], &mut st);
                self.preamble = false;
                if let Node::Env(e) = &nodes[i] {
                    let mut st = St::default();
                    self.walk(&e.body, &mut st);
                }
            }
            None => self.walk(nodes, &mut st),
        }
    }

    fn walk(&mut self, nodes: &[Node], st: &mut St) {
        for n in nodes {
            if self.limit {
                return;
            }
            self.node(n, st);
        }
    }

    fn node(&mut self, n: &Node, st: &mut St) {
        match n {
            Node::Text(s) => self.text(s, st),
            Node::Space => {
                if !self.preamble {
                    self.space(st)
                }
            }
            Node::Par => {
                if !self.preamble {
                    self.par(st)
                }
            }
            Node::Group(g) => {
                let mut s2 = st.clone();
                self.walk(g, &mut s2);
            }
            Node::Cmd(c) => self.cmd(c, st),
            Node::Env(e) => {
                if !self.preamble {
                    self.env(e, st)
                }
            }
            Node::Math { display, env, body } => {
                if self.preamble {
                    return;
                }
                if *display {
                    self.display_math(env.as_deref(), body, st);
                } else {
                    let t = self.math(body);
                    if !t.is_empty() {
                        let mut s2 = st.clone();
                        s2.run.font = Some(MATH_FONT.into());
                        self.push_run(t, &s2);
                    }
                }
            }
            Node::Verb(t) => {
                let mut s2 = st.clone();
                s2.run.font = Some(MONO_FONT.into());
                self.push_run(t.replace('\n', " "), &s2);
            }
            Node::Verbatim { text, .. } => self.verbatim(text, st),
            Node::Tab => self.space(st),
            Node::Sup => self.text("^", st),
            Node::Sub => self.text("_", st),
        }
    }

    fn styled(&mut self, nodes: &[Node], st: &St, f: impl FnOnce(&mut St)) {
        let mut s = st.clone();
        f(&mut s);
        self.walk(nodes, &mut s);
    }

    /// Convert nodes in isolation and flatten the result to runs.
    fn inline_runs(&mut self, nodes: &[Node], st: &St) -> Vec<Run> {
        let blocks = self.sub_blocks(nodes, st);
        let mut runs = Vec::new();
        for b in &blocks {
            let r = b.para_runs();
            if r.is_empty() {
                continue;
            }
            if !runs.is_empty() {
                runs.push(Run::text(" "));
            }
            runs.extend(r.iter().cloned());
        }
        normalize_runs(runs)
    }

    fn sub_blocks(&mut self, nodes: &[Node], st: &St) -> Vec<Block> {
        let saved_out = std::mem::take(&mut self.out);
        let saved_para = std::mem::take(&mut self.para);
        let saved_lists = std::mem::take(&mut self.lists);
        let saved_quote = std::mem::replace(&mut self.quote, 0);
        let saved_pre = std::mem::replace(&mut self.preamble, false);
        let mut s = st.clone();
        s.align = None;
        self.walk(nodes, &mut s);
        self.flush(None);
        let blocks = std::mem::replace(&mut self.out, saved_out);
        self.para = saved_para;
        self.lists = saved_lists;
        self.quote = saved_quote;
        self.preamble = saved_pre;
        blocks
    }

    fn plain_of(&mut self, nodes: &[Node]) -> String {
        let runs = self.inline_runs(nodes, &St::default());
        plain(&runs).replace('\n', " ")
    }

    fn verbatim(&mut self, text: &str, st: &St) {
        if self.preamble {
            return;
        }
        self.flush(st.align);
        let mut t = text;
        match t.find('\n') {
            Some(i) if t[..i].trim().is_empty() => t = &t[i + 1..],
            None if t.trim().is_empty() => return,
            _ => {}
        }
        let mut lines: Vec<&str> = t.split('\n').collect();
        if lines.last().is_some_and(|l| l.trim().is_empty()) {
            lines.pop();
        }
        for l in lines {
            if self.limit {
                return;
            }
            self.text_len += l.len();
            if self.text_len > MAX_TEXT {
                self.hit_limit();
                return;
            }
            let runs = if l.is_empty() { Vec::new() } else { vec![Run::text(l)] };
            self.push_block(Block::Code { runs, align: None });
        }
    }

    // ---- commands ----

    fn preamble_cmd(&mut self, c: &CmdNode) {
        match c.name.as_str() {
            "documentclass" => self.set_class(c),
            "title" | "author" | "date" | "subtitle" => self.meta_cmd(c),
            "definecolor" | "colorlet" => self.define_color(c),
            "newtheorem" => self.newtheorem(c),
            _ => {}
        }
    }

    fn set_class(&mut self, c: &CmdNode) {
        let class = nodes_src(arg(c, 1)).trim().to_string();
        self.class = match class.as_str() {
            "report" | "book" | "scrreprt" | "scrbook" | "memoir" | "extreport" | "extbook" | "amsbook" | "thesis"
            | "ociamthesis" | "tufte-book" | "kaobook" => Class::Chapters,
            "beamer" => Class::Beamer,
            _ => Class::Article,
        };
    }

    fn meta_cmd(&mut self, c: &CmdNode) {
        let v = Some(arg(c, 1).to_vec());
        match c.name.as_str() {
            "title" => self.title = v,
            "author" => self.author = v,
            "date" => self.date = v,
            "subtitle" => self.subtitle = v,
            _ => {}
        }
    }

    fn define_color(&mut self, c: &CmdNode) {
        let name = nodes_src(arg(c, 1)).trim().to_string();
        let value = if c.name == "definecolor" {
            let model = nodes_src(arg(c, 2));
            color_value(Some(model.trim()), &nodes_src(arg(c, 3)), &self.colors)
        } else {
            color_value(None, &nodes_src(arg(c, 2)), &self.colors)
        };
        if let Some(v) = value {
            if !name.is_empty() && self.colors.len() < 10_000 {
                self.colors.insert(name, v);
            }
        }
    }

    fn newtheorem(&mut self, c: &CmdNode) {
        let name = nodes_src(arg(c, 0)).trim().to_string();
        if name.is_empty() {
            return;
        }
        let label = self.plain_of(arg(c, 2));
        let shared = opt(c, 1).map(|o| nodes_src(o).trim().to_string()).filter(|s| !s.is_empty());
        let within = opt(c, 3).map(|o| nodes_src(o).trim().to_string()).filter(|s| !s.is_empty());
        let counter = if c.star { None } else { Some(shared.unwrap_or_else(|| name.clone())) };
        self.theorems.insert(name, Thm { label, counter, within });
    }

    fn color_arg(&self, c: &CmdNode, model_i: usize, spec_i: usize) -> Option<String> {
        let model = opt(c, model_i).map(nodes_src);
        color_value(model.as_deref(), &nodes_src(arg(c, spec_i)), &self.colors)
    }

    fn cmd(&mut self, c: &CmdNode, st: &mut St) {
        let name = c.name.as_str();
        if self.preamble {
            self.preamble_cmd(c);
            return;
        }
        match name {
            "part" | "chapter" | "section" | "subsection" | "subsubsection" | "paragraph" | "subparagraph" | "addchap"
            | "addsec" | "addpart" | "minisec" => self.heading(c, st),
            "frametitle" | "framesubtitle" => {
                self.flush(st.align);
                let runs = heading_runs(self.inline_runs(arg(c, 1), &St::default()));
                if !runs.is_empty() {
                    let lvl = if name == "frametitle" { 2 } else { 3 };
                    self.push_block(heading_block(lvl, runs, None));
                }
            }
            "documentclass" => self.set_class(c),
            "title" | "author" | "date" | "subtitle" => self.meta_cmd(c),
            "definecolor" | "colorlet" => self.define_color(c),
            "newtheorem" => self.newtheorem(c),
            "maketitle" | "titlepage" => self.title_block(st),
            "item" => self.item(c, st),
            "bibitem" => self.bibitem(c, st),
            "\\" | "newline" | "tabularnewline" | "linebreak" | "cr" => self.push_run("\n", st),
            "par" => self.par(st),
            "smallskip" | "medskip" | "bigskip" => self.flush(st.align),
            "newpage" | "clearpage" | "cleardoublepage" | "pagebreak" | "framebreak" => {
                if self.class != Class::Beamer {
                    self.flush(st.align);
                    self.push_block(Block::PageBreak);
                }
            }
            "centering" => st.align = Some("center"),
            "raggedleft" => st.align = Some("right"),
            "raggedright" => st.align = None,
            "centerline" => {
                self.flush(st.align);
                let mut s = st.clone();
                s.align = Some("center");
                self.walk(arg(c, 0), &mut s);
                self.flush(Some("center"));
            }
            "textbf" => self.styled(arg(c, 0), st, |s| s.run.b = true),
            "textit" | "textsl" => self.styled(arg(c, 0), st, |s| s.run.i = true),
            "emph" | "emphasize" => self.styled(arg(c, 0), st, |s| s.run.i = !s.run.i),
            "textmd" => self.styled(arg(c, 0), st, |s| s.run.b = false),
            "textup" => self.styled(arg(c, 0), st, |s| s.run.i = false),
            "textnormal" => self.styled(arg(c, 0), st, |s| {
                s.run.b = false;
                s.run.i = false;
                s.run.font = None;
            }),
            "underline" | "uline" | "uuline" | "uwave" | "underbar" => self.styled(arg(c, 0), st, |s| s.run.u = true),
            "sout" | "st" | "xout" => self.styled(arg(c, 0), st, |s| s.run.s = true),
            "texttt" => self.styled(arg(c, 0), st, |s| s.run.font = Some(MONO_FONT.into())),
            "textsuperscript" => self.styled(arg(c, 0), st, |s| s.run.sup = true),
            "textsubscript" => self.styled(arg(c, 0), st, |s| s.run.sub = true),
            "textsf" | "textrm" | "textsc" | "mbox" | "hbox" | "text" | "textbox" | "textlatin" | "textgreek"
            | "titlecap" | "smash" | "inst" | "only" | "uncover" | "visible" | "temporal" | "action" | "alt" => {
                if name == "inst" {
                    return;
                }
                let mut s = st.clone();
                self.walk(arg(c, 0), &mut s);
            }
            "textcolor" => {
                let col = self.color_arg(c, 0, 1);
                self.styled(arg(c, 2), st, |s| {
                    if col.is_some() {
                        s.run.color = col
                    }
                });
            }
            "color" => {
                if let Some(col) = self.color_arg(c, 0, 1) {
                    st.run.color = Some(col);
                }
            }
            "colorbox" => {
                let col = self.color_arg(c, 0, 1);
                self.styled(arg(c, 2), st, |s| {
                    if col.is_some() {
                        s.run.hl = col
                    }
                });
            }
            "fcolorbox" => {
                let col = self.color_arg(c, 2, 3);
                self.styled(arg(c, 4), st, |s| {
                    if col.is_some() {
                        s.run.hl = col
                    }
                });
            }
            "hl" | "highlight" => self.styled(arg(c, 0), st, |s| s.run.hl = Some("#FFFF00".into())),
            "alert" => self.styled(arg(c, 0), st, |s| s.run.color = Some("#FF0000".into())),
            "structure" => self.styled(arg(c, 0), st, |s| s.run.color = Some("#3333B3".into())),
            "href" => {
                let url = unescape_url(&nodes_src(arg(c, 0)));
                let link = safe_link(&url);
                self.styled(arg(c, 1), st, |s| s.run.link = link);
            }
            "url" | "path" | "nolinkurl" => {
                let url = unescape_url(&nodes_src(arg(c, 0)));
                let mut s = st.clone();
                if name == "url" {
                    s.run.link = safe_link(&url);
                }
                self.push_run(url, &s);
            }
            "hyperref" => {
                let mut s = st.clone();
                self.walk(arg(c, 1), &mut s);
            }
            "hyperlink" | "hypertarget" => {
                let mut s = st.clone();
                self.walk(arg(c, 1), &mut s);
            }
            "bf" | "bfseries" => st.run.b = true,
            "it" | "itshape" | "sl" | "slshape" => st.run.i = true,
            "em" => st.run.i = !st.run.i,
            "tt" | "ttfamily" => st.run.font = Some(MONO_FONT.into()),
            "upshape" => st.run.i = false,
            "mdseries" => st.run.b = false,
            "normalfont" | "rm" | "rmfamily" => {
                st.run.b = false;
                st.run.i = false;
                st.run.font = None;
            }
            "footnote" | "sidenote" | "marginpar" => {
                let body = last_arg(c).to_vec();
                let num = if name == "footnote" { opt(c, 0).map(nodes_src) } else { None };
                self.footnote(&body, num, true);
            }
            "thanks" => {
                let body = arg(c, 0).to_vec();
                self.footnote(&body, None, true);
            }
            "footnotemark" => {
                let n = opt(c, 0).and_then(|o| nodes_src(o).trim().parse::<u32>().ok()).unwrap_or_else(|| self.next_note());
                self.marks.push(n);
                let mut s = St::default();
                s.run.sup = true;
                self.push_run(format!("[{n}]"), &s);
            }
            "footnotetext" => {
                let body = arg(c, 1).to_vec();
                let num = opt(c, 0).map(nodes_src).or_else(|| (!self.marks.is_empty()).then(|| self.marks.remove(0).to_string()));
                self.footnote(&body, num, false);
            }
            "label" => self.label(c),
            "ref" | "vref" | "subref" | "labelcref" => self.ref_(c, "ref", st),
            "eqref" => self.ref_(c, "eqref", st),
            "pageref" | "autopageref" => self.ref_(c, "pageref", st),
            "autoref" | "cref" | "Cref" => self.ref_(c, "autoref", st),
            "nameref" => self.ref_(c, "nameref", st),
            "cite" | "citep" | "citet" | "parencite" | "textcite" | "autocite" | "footcite" | "citeauthor" | "citeyear"
            | "Citep" | "Citet" | "Cite" | "supercite" | "citealp" | "citealt" | "Parencite" | "Textcite" | "Autocite" => {
                self.cite(c, st)
            }
            "caption" | "captionof" | "subcaption" => self.caption(c, st),
            "includegraphics" | "includesvg" | "includepdf" | "includestandalone" => {
                let file = nodes_src(last_arg(c));
                self.image(&file, st);
            }
            "input" | "include" | "subfile" | "InputIfFileExists" => self.warn(W_INPUT),
            "bibliography" | "printbibliography" => self.warn(W_BIBTEX),
            "today" => {
                let t = self.today.clone();
                self.push_run(t, st);
            }
            "and" => {
                while let Some(last) = self.para.last_mut() {
                    let keep = last.t.trim_end_matches(' ').len();
                    last.t.truncate(keep);
                    if last.t.is_empty() {
                        self.para.pop();
                    } else {
                        break;
                    }
                }
                self.push_run(", ", st);
            }
            "ensuremath" => {
                let t = self.math(arg(c, 0));
                self.push_run(t, st);
            }
            "fbox" | "framebox" | "boxed" => {
                let runs = self.inline_runs(last_arg(c), st);
                let text = plain(&runs);
                match text.strip_prefix("Image: ") {
                    Some(rest) => {
                        let t = format!("[Image: {rest}]");
                        self.push_run(t, &St::default());
                    }
                    None => self.push_runs(runs),
                }
            }
            "makebox" | "raisebox" | "parbox" | "resizebox" | "scalebox" | "rotatebox" | "adjustbox" | "textls"
            | "lettrine" | "multicolumn" | "multirow" | "foreignlanguage" => {
                if name == "lettrine" {
                    // \lettrine{T}{he} -> "The"
                    let mut s = st.clone();
                    self.walk(arg(c, 1), &mut s);
                    self.walk(arg(c, 2), &mut s);
                    return;
                }
                let mut s = st.clone();
                self.walk(last_arg(c), &mut s);
            }
            "texorpdfstring" => {
                let mut s = st.clone();
                self.walk(arg(c, 0), &mut s);
            }
            "enquote" | "foreignquote" => {
                self.push_run("“", st);
                let mut s = st.clone();
                self.walk(last_arg(c), &mut s);
                self.push_run("”", st);
            }
            "MakeUppercase" | "uppercase" | "MakeLowercase" | "lowercase" => {
                let up = name.contains("pper");
                let mut runs = self.inline_runs(arg(c, 0), st);
                for r in &mut runs {
                    r.t = if up { r.t.to_uppercase() } else { r.t.to_lowercase() };
                }
                self.push_runs(runs);
            }
            "lipsum" | "blindtext" | "Blindtext" => {
                self.push_run("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.", st);
            }
            "hspace" | "hfill" | "quad" | "qquad" | "enspace" | "hfil" => {
                if let Some(t) = text_sym(name) {
                    if t != " " {
                        self.push_run(t, st);
                        return;
                    }
                }
                self.space(st);
            }
            "opening" | "closing" | "signature" | "encl" | "cc" | "ps" => {
                self.flush(st.align);
                let prefix = match name {
                    "encl" => "Enclosure: ",
                    "cc" => "cc: ",
                    _ => "",
                };
                self.push_run(prefix, st);
                let mut s = st.clone();
                self.walk(arg(c, 0), &mut s);
                self.flush(st.align);
            }
            "appendix" => {
                self.appendix = true;
                let top = if self.chapters() { 1 } else { 2 };
                for k in top..7 {
                    self.sec[k] = 0;
                }
            }
            "num" => {
                let mut s = st.clone();
                self.walk(arg(c, 0), &mut s);
            }
            "ang" => {
                let mut s = st.clone();
                self.walk(arg(c, 0), &mut s);
                self.push_run("°", st);
            }
            "SI" | "qty" | "si" | "unit" => {
                let (num, unit) = match name {
                    "SI" => (Some(arg(c, 1)), arg(c, 3)),
                    "qty" => (Some(arg(c, 1)), arg(c, 2)),
                    _ => (None, arg(c, 1)),
                };
                if let Some(n) = num {
                    let mut s = st.clone();
                    self.walk(n, &mut s);
                    self.push_run("\u{a0}", st);
                }
                let u = self.unit_text(unit);
                self.push_run(u, st);
            }
            "'" | "`" | "^" | "\"" | "~" | "=" | "." | "u" | "v" | "H" | "c" | "k" | "r" | "d" | "b" | "t" => {
                let inner = self.math_text(arg(c, 0));
                self.push_run(accent(name, &inner), st);
            }
            _ => {
                if let Some(s) = text_sym(name) {
                    self.push_run(s, st);
                    return;
                }
                if let Some(s) = math_sym(name) {
                    self.push_run(s, st);
                    return;
                }
                if FUNCS.contains(&name) {
                    self.push_run(name, st);
                    return;
                }
                if matches!(
                    name,
                    "frac" | "dfrac" | "tfrac" | "cfrac" | "sqrt" | "binom" | "dbinom" | "tbinom" | "mathbf" | "mathrm"
                        | "mathit" | "mathcal" | "mathbb" | "mathfrak" | "mathsf" | "mathtt" | "boldsymbol" | "bm"
                        | "hat" | "bar" | "vec" | "dot" | "ddot" | "tilde" | "overline" | "widehat" | "widetilde"
                        | "operatorname"
                ) {
                    let t = self.math(std::slice::from_ref(&Node::Cmd(Box::new(c.clone()))));
                    self.push_run(t, st);
                    return;
                }
                if cmd_spec(name).is_none() {
                    self.unknown_cmd(name);
                }
            }
        }
    }

    fn unit_text(&mut self, nodes: &[Node]) -> String {
        let mut s = String::new();
        for n in nodes {
            match n {
                Node::Cmd(c) => {
                    if let Some(u) = unit_symbol(&c.name) {
                        if u == "/" || u == "²" || u == "³" {
                            s.push_str(u);
                        } else {
                            if !s.is_empty() && !s.ends_with(['/', 'k', 'm', 'µ', 'n', 'c', 'M', 'G']) {
                                s.push('·');
                            }
                            s.push_str(u);
                        }
                    } else if let Some(t) = text_sym(&c.name).or_else(|| math_sym(&c.name)) {
                        s.push_str(t);
                    } else {
                        s.push_str(&c.name);
                    }
                }
                Node::Text(t) => s.push_str(t),
                Node::Group(g) => s.push_str(&self.unit_text(g)),
                Node::Space => s.push(' '),
                Node::Math { body, .. } => s.push_str(&self.math(body)),
                Node::Sup => s.push('^'),
                _ => {}
            }
        }
        s
    }

    fn next_note(&self) -> u32 {
        let used = self.footnotes.iter().map(|(n, _)| *n).chain(self.marks.iter().copied()).max().unwrap_or(0);
        used + 1
    }

    fn footnote(&mut self, body: &[Node], num: Option<String>, mark: bool) {
        let n = num.and_then(|s| s.trim().parse::<u32>().ok()).unwrap_or_else(|| self.next_note());
        let saved = std::mem::replace(&mut self.target, (n.to_string(), "footnote".into()));
        let runs = self.inline_runs(body, &St::default());
        self.target = saved;
        if self.footnotes.len() < 100_000 {
            self.footnotes.push((n, runs));
        }
        if mark {
            let mut s = St::default();
            s.run.sup = true;
            self.push_run(format!("[{n}]"), &s);
        }
    }

    fn label(&mut self, c: &CmdNode) {
        let key = nodes_src(arg(c, 0)).trim().to_string();
        if key.is_empty() || self.labels.len() > 100_000 {
            return;
        }
        let l = Label { num: self.target.0.clone(), kind: self.target.1.clone(), title: self.target_title.clone() };
        self.labels.entry(key).or_insert(l);
    }

    fn ref_(&mut self, c: &CmdNode, op: &str, st: &St) {
        let key = nodes_src(arg(c, 0)).trim().to_string();
        let op = if op == "autoref" && c.name == "Cref" { "Autoref" } else { op };
        self.push_run(placeholder(op, &key), st);
    }

    fn cite(&mut self, c: &CmdNode, st: &St) {
        let keys = nodes_src(arg(c, 2)).trim().to_string();
        let (pre, post) = match (opt(c, 0), opt(c, 1)) {
            (Some(a), Some(b)) => (self.plain_of(a), self.plain_of(b)),
            (Some(a), None) => (String::new(), self.plain_of(a)),
            _ => (String::new(), String::new()),
        };
        let arg = format!("{keys}{PLACE_SUB}{pre}{PLACE_SUB}{post}");
        self.push_run(placeholder("cite", &arg), st);
    }

    fn sec_number(&self, idx: usize) -> String {
        if idx == 0 {
            return roman(self.sec[0]);
        }
        let top = if self.chapters() { 1 } else { 2 };
        if idx < top {
            return String::new();
        }
        (top..=idx.min(6))
            .map(|k| if k == top && self.appendix { letter(self.sec[k]) } else { self.sec[k].to_string() })
            .collect::<Vec<_>>()
            .join(".")
    }

    fn chap_num(&self, n: u32) -> String {
        if self.chapters() && self.sec[1] > 0 {
            format!("{}.{}", self.sec_number(1), n)
        } else {
            n.to_string()
        }
    }

    fn heading(&mut self, c: &CmdNode, st: &St) {
        self.flush(st.align);
        let name = c.name.as_str();
        let ch = self.chapters();
        let (idx, level) = match name {
            "part" | "addpart" => (0, 1),
            "chapter" | "addchap" => (1, 1),
            "section" | "addsec" => (2, if ch { 2 } else { 1 }),
            "subsection" => (3, if ch { 3 } else { 2 }),
            "subsubsection" => (4, if ch { 4 } else { 3 }),
            "paragraph" => (5, if ch { 5 } else { 4 }),
            "subparagraph" => (6, if ch { 6 } else { 5 }),
            _ => (6, if ch { 6 } else { 5 }),
        };
        let numbered = !c.star && !name.starts_with("add") && name != "minisec";
        let title_nodes = if name == "minisec" { arg(c, 0) } else { arg(c, 1) };
        if numbered {
            self.sec[idx] = self.sec[idx].saturating_add(1);
            if idx > 0 {
                for k in idx + 1..7 {
                    self.sec[k] = 0;
                }
            }
            if idx == 1 {
                self.fig = 0;
                self.tab = 0;
                self.eq = 0;
            }
            let kind = match idx {
                0 => "Part",
                1 if self.appendix => "Appendix",
                1 => "Chapter",
                2 if self.appendix && !ch => "Appendix",
                _ => "Section",
            };
            self.target = (self.sec_number(idx), kind.into());
        }
        let runs = heading_runs(self.inline_runs(title_nodes, &St::default()));
        self.target_title = plain(&runs);
        if !runs.is_empty() {
            self.push_block(heading_block(level, runs, None));
        }
    }

    fn title_block(&mut self, st: &St) {
        self.flush(st.align);
        let base = St::default();
        if let Some(t) = self.title.clone() {
            let runs = heading_runs(self.inline_runs(&t, &base));
            if !runs.is_empty() {
                self.push_block(Block::H1 { runs, align: Some("center".into()) });
            }
        }
        if let Some(t) = self.subtitle.clone() {
            let runs = self.inline_runs(&t, &base);
            if !runs.is_empty() {
                self.push_block(Block::P { runs, align: Some("center".into()) });
            }
        }
        let mut runs = match self.author.clone() {
            Some(a) => self.inline_runs(&a, &base),
            None => Vec::new(),
        };
        let date = match self.date.clone() {
            Some(d) => self.inline_runs(&d, &base),
            None => vec![Run::text(self.today.clone())],
        };
        if !date.is_empty() {
            if !runs.is_empty() {
                runs.push(Run::text("\n"));
            }
            runs.extend(date);
        }
        let runs = normalize_runs(runs);
        if !runs.is_empty() {
            self.push_block(Block::P { runs, align: Some("center".into()) });
        }
    }

    fn item(&mut self, c: &CmdNode, st: &St) {
        self.flush(st.align);
        let label = opt(c, 0).map(|l| l.to_vec());
        let q = self.quote;
        let info = match self.lists.last_mut() {
            Some(l) if q <= l.quote_at => {
                l.fresh = true;
                l.n += 1;
                Some((l.desc, l.kind, l.n))
            }
            _ => None,
        };
        let desc = match info {
            Some((desc, kind, n)) => {
                if kind == "ol" {
                    self.target = (n.to_string(), "item".into());
                }
                desc
            }
            None => {
                if label.is_none() {
                    self.push_run("• ", &St::default());
                }
                false
            }
        };
        if let Some(l) = label {
            let mut runs = self.inline_runs(&l, &St::default());
            if desc {
                for r in &mut runs {
                    r.b = true;
                }
            }
            if !runs.is_empty() {
                runs.push(Run::text(" "));
                self.push_runs(runs);
            }
        }
    }

    fn bibitem(&mut self, c: &CmdNode, st: &St) {
        self.flush(st.align);
        self.bib_n += 1;
        let key = nodes_src(arg(c, 1)).trim().to_string();
        let label = opt(c, 0).map(|o| self.plain_of(o)).filter(|s| !s.trim().is_empty()).unwrap_or_else(|| self.bib_n.to_string());
        if !key.is_empty() && self.bib.len() < 100_000 {
            self.bib.entry(key).or_insert_with(|| label.clone());
        }
        self.push_run(format!("[{label}] "), &St::default());
    }

    fn caption(&mut self, c: &CmdNode, st: &St) {
        self.flush(st.align);
        let (kind, body, star) = match c.name.as_str() {
            "captionof" => (nodes_src(arg(c, 0)).trim().to_string(), arg(c, 2), c.star),
            "subcaption" => ("sub".to_string(), arg(c, 1), false),
            _ => (self.float.unwrap_or("figure").to_string(), arg(c, 1), c.star),
        };
        let prefix = match kind.as_str() {
            _ if star => None,
            "table" => {
                self.tab += 1;
                let n = self.chap_num(self.tab);
                self.target = (n.clone(), "Table".into());
                Some(format!("Table {n}: "))
            }
            "sub" | "subfigure" | "subtable" => {
                self.subfig += 1;
                let l = letter(self.subfig).to_lowercase();
                self.target = (l.clone(), "Figure".into());
                Some(format!("({l}) "))
            }
            _ => {
                self.fig += 1;
                let n = self.chap_num(self.fig);
                self.target = (n.clone(), "Figure".into());
                Some(format!("Figure {n}: "))
            }
        };
        let mut runs: Vec<Run> = prefix.map(Run::text).into_iter().collect();
        runs.extend(self.inline_runs(body, &St::default()));
        let runs = normalize_runs(runs);
        if !runs.is_empty() {
            self.push_block(Block::P { runs, align: st.align.map(String::from) });
        }
    }

    fn image(&mut self, file: &str, st: &St) {
        self.flush(st.align);
        let f = file.trim();
        let f = if f.is_empty() { "image" } else { f };
        self.push_block(Block::P { runs: vec![Run::text(format!("[Image: {f}]"))], align: st.align.map(String::from) });
        self.warn(W_IMAGES);
    }

    // ---- environments ----

    fn env(&mut self, e: &EnvNode, st: &St) {
        let base = e.name.trim_end_matches('*');
        match base {
            "document" => {
                let mut s = st.clone();
                self.walk(&e.body, &mut s);
            }
            "itemize" | "compactitem" | "asparaitem" | "inparaitem" | "dinglist" => self.list_env(e, st, "ul", false),
            "enumerate" | "compactenum" | "asparaenum" | "inparaenum" => self.list_env(e, st, "ol", false),
            "description" | "compactdesc" | "asparadesc" => self.list_env(e, st, "ul", true),
            "quote" | "quotation" | "verse" | "displayquote" | "chapquote" => {
                self.flush(st.align);
                self.quote += 1;
                let mut s = st.clone();
                self.walk(&e.body, &mut s);
                self.flush(None);
                self.quote -= 1;
            }
            "center" | "flushright" | "flushleft" => {
                self.flush(st.align);
                let a = match base {
                    "center" => Some("center"),
                    "flushright" => Some("right"),
                    _ => None,
                };
                let mut s = st.clone();
                s.align = a;
                self.walk(&e.body, &mut s);
                self.flush(s.align);
            }
            "abstract" => {
                self.flush(st.align);
                self.push_block(Block::H2 { runs: vec![Run::text("Abstract")], align: Some("center".into()) });
                let mut s = st.clone();
                self.walk(&e.body, &mut s);
                self.flush(s.align);
            }
            "figure" | "wrapfigure" | "SCfigure" | "sidewaysfigure" | "marginfigure" => self.float_env(e, st, "figure"),
            "table" | "wraptable" | "sidewaystable" | "margintable" => self.float_env(e, st, "table"),
            "subfigure" | "subtable" => self.float_env(e, st, "subfigure"),
            "tabular" | "tabularx" | "tabulary" | "longtable" | "tabu" | "longtabu" | "supertabular" | "xtabular"
            | "tblr" | "longtblr" | "array" | "mpsupertabular" => self.table(e, st),
            "thebibliography" => {
                self.flush(st.align);
                let title = if self.chapters() { "Bibliography" } else { "References" };
                self.push_block(Block::H1 { runs: vec![Run::text(title)], align: None });
                let mut s = st.clone();
                self.walk(&e.body, &mut s);
                self.flush(s.align);
            }
            "frame" => {
                self.flush(st.align);
                for (i, lvl) in [(1usize, 2u32), (2, 3)] {
                    if let Some(t) = opt_of(&e.args, i) {
                        let runs = heading_runs(self.inline_runs(t, &St::default()));
                        if !runs.is_empty() {
                            self.push_block(heading_block(lvl, runs, None));
                        }
                    }
                }
                let mut s = st.clone();
                self.walk(&e.body, &mut s);
                self.flush(s.align);
            }
            "block" | "alertblock" | "exampleblock" => {
                self.flush(st.align);
                let mut runs = self.inline_runs(arg_of(&e.args, 0), &St::default());
                for r in &mut runs {
                    r.b = true;
                }
                if !runs.is_empty() {
                    self.push_block(Block::P { runs, align: None });
                }
                let mut s = st.clone();
                self.walk(&e.body, &mut s);
                self.flush(s.align);
            }
            "proof" => {
                self.flush(st.align);
                let mut head = St::default();
                head.run.i = true;
                match opt_of(&e.args, 0) {
                    Some(t) => {
                        let mut runs = self.inline_runs(t, &head);
                        for r in &mut runs {
                            r.i = true;
                        }
                        self.push_runs(runs);
                        self.push_run(".", &head);
                    }
                    None => self.push_run("Proof.", &head),
                }
                self.push_run(" ", st);
                let mut s = st.clone();
                self.walk(&e.body, &mut s);
                self.push_run(" ∎", st);
                self.flush(s.align);
            }
            "tikzpicture" | "pgfpicture" | "picture" | "circuitikz" | "forest" | "pspicture" | "tikzcd" => {
                self.flush(st.align);
                self.push_block(Block::P { runs: vec![Run::text("[Image: drawing]")], align: st.align.map(String::from) });
                self.warn(W_DRAWINGS);
            }
            "letter" => {
                self.flush(st.align);
                let mut s = st.clone();
                let addr = arg_of(&e.args, 0).to_vec();
                self.walk(&addr, &mut s);
                self.flush(None);
                let mut s = st.clone();
                self.walk(&e.body, &mut s);
                self.flush(s.align);
            }
            "titlepage" | "minipage" | "multicols" | "columns" | "column" | "landscape" | "spacing" | "onehalfspace"
            | "doublespace" | "singlespace" | "otherlanguage" | "samepage" | "framed" | "mdframed" | "tcolorbox"
            | "boxedminipage" | "small" | "footnotesize" | "scriptsize" | "large" | "Large" | "normalsize" | "adjustbox"
            | "appendices" | "subappendices" | "sloppypar" | "raggedright" | "raggedleft" | "fullwidth" | "addmargin"
            | "leftbar" | "shaded" | "trivlist" | "list" | "english" | "frontmatter" | "mainmatter" | "backmatter"
            | "nopagebreak" | "minipage*" | "wide" | "widepage" | "changemargin" | "itemize*" | "lrbox" | "comment*"
            | "linenumbers" | "flushleft*" | "bfseries" | "itshape" | "onlyenv" | "uncoverenv" | "visibleenv"
            | "actionenv" | "overlayarea" | "overprint" | "noindent" | "figure*" | "algorithm" | "algorithmic"
            | "subequations" | "keywords" | "IEEEkeywords" | "acknowledgments" | "acknowledgements" => {
                if matches!(base, "minipage" | "column" | "titlepage" | "algorithm" | "algorithmic" | "multicols") {
                    self.flush(st.align);
                }
                let mut s = st.clone();
                if base == "bfseries" {
                    s.run.b = true;
                }
                if base == "itshape" {
                    s.run.i = true;
                }
                self.walk(&e.body, &mut s);
                if matches!(base, "minipage" | "column" | "titlepage" | "algorithm" | "algorithmic") {
                    self.flush(s.align);
                }
            }
            _ => {
                if let Some((label, counter, within)) = self.theorem_info(base) {
                    self.theorem(e, st, label, counter, within);
                    return;
                }
                self.unknown_cmd(&format!("begin{{{}}}", e.name));
                let mut s = st.clone();
                self.walk(&e.body, &mut s);
            }
        }
    }

    fn theorem_info(&self, base: &str) -> Option<(String, Option<String>, Option<String>)> {
        if let Some(t) = self.theorems.get(base) {
            let within = match &t.counter {
                Some(cn) if cn != base => self.theorems.get(cn).and_then(|x| x.within.clone()).or(t.within.clone()),
                _ => t.within.clone(),
            };
            return Some((t.label.clone(), t.counter.clone(), within));
        }
        default_theorem(base).map(|(l, numbered)| (l.to_string(), numbered.then(|| base.to_string()), None))
    }

    fn thm_number(&mut self, counter: &str, within: Option<&str>) -> String {
        let w = match within {
            Some("chapter") => self.sec_number(1),
            Some("section") => self.sec_number(2),
            Some("subsection") => self.sec_number(3),
            _ => String::new(),
        };
        let e = self.thm_counts.entry(counter.to_string()).or_insert((0, w.clone()));
        if e.1 != w {
            e.0 = 0;
            e.1 = w.clone();
        }
        e.0 += 1;
        if w.is_empty() {
            e.0.to_string()
        } else {
            format!("{w}.{}", e.0)
        }
    }

    fn theorem(&mut self, e: &EnvNode, st: &St, label: String, counter: Option<String>, within: Option<String>) {
        self.flush(st.align);
        let mut head = label.clone();
        if let Some(cn) = &counter {
            let n = self.thm_number(cn, within.as_deref());
            head = format!("{head} {n}").trim().to_string();
            self.target = (n, label.clone());
        }
        let mut b = St::default();
        b.run.b = true;
        self.push_run(head, &b);
        if let Some(note) = opt_of(&e.args, 0) {
            let runs = self.inline_runs(note, &St::default());
            if !runs.is_empty() {
                self.push_run(" (", &St::default());
                self.push_runs(runs);
                self.push_run(")", &St::default());
            }
        }
        self.push_run(".", &b);
        self.push_run(" ", &St::default());
        let mut s = st.clone();
        self.walk(&e.body, &mut s);
        self.flush(s.align);
    }

    fn list_env(&mut self, e: &EnvNode, st: &St, kind: &'static str, desc: bool) {
        self.flush(st.align);
        if self.lists.len() >= 64 {
            let mut s = st.clone();
            self.walk(&e.body, &mut s);
            return;
        }
        self.lists.push(ListCtx { kind, desc, fresh: false, n: 0, quote_at: self.quote });
        let mut s = st.clone();
        self.walk(&e.body, &mut s);
        self.flush(s.align);
        self.lists.pop();
    }

    fn float_env(&mut self, e: &EnvNode, st: &St, kind: &'static str) {
        self.flush(st.align);
        let saved = self.float.replace(kind);
        if kind != "subfigure" {
            self.subfig = 0;
        }
        let mut s = st.clone();
        self.walk(&e.body, &mut s);
        self.flush(s.align);
        self.float = saved;
    }

    fn table(&mut self, e: &EnvNode, st: &St) {
        self.flush(st.align);
        // rows -> cells -> nodes; longtable header/footer segments.
        let mut segs: Vec<(String, Vec<Vec<Vec<Node>>>)> = Vec::new();
        let mut rows: Vec<Vec<Vec<Node>>> = vec![vec![Vec::new()]];
        for n in &e.body {
            match n {
                Node::Tab => {
                    if let Some(r) = rows.last_mut() {
                        if r.len() < 1000 {
                            r.push(Vec::new());
                        }
                    }
                }
                Node::Cmd(c) => match c.name.as_str() {
                    "\\" | "tabularnewline" | "cr" => rows.push(vec![Vec::new()]),
                    "hline" | "cline" | "toprule" | "midrule" | "bottomrule" | "cmidrule" | "hhline" | "addlinespace"
                    | "specialrule" | "morecmidrules" | "rowcolor" | "noalign" | "arrayrulecolor" => {}
                    "endfirsthead" | "endhead" | "endfoot" | "endlastfoot" => {
                        let taken = std::mem::replace(&mut rows, vec![vec![Vec::new()]]);
                        segs.push((c.name.clone(), taken));
                    }
                    _ => {
                        if let Some(cell) = rows.last_mut().and_then(|r| r.last_mut()) {
                            cell.push(n.clone());
                        }
                    }
                },
                _ => {
                    if let Some(cell) = rows.last_mut().and_then(|r| r.last_mut()) {
                        cell.push(n.clone());
                    }
                }
            }
        }
        let mut all: Vec<Vec<Vec<Node>>> = Vec::new();
        let has_first = segs.iter().any(|(k, _)| k == "endfirsthead");
        let mut tail: Vec<Vec<Vec<Node>>> = Vec::new();
        let mut foot: Vec<Vec<Vec<Node>>> = Vec::new();
        let mut has_last = false;
        for (k, r) in segs {
            match k.as_str() {
                "endfirsthead" => all.extend(r),
                "endhead" => {
                    if !has_first {
                        all.extend(r)
                    }
                }
                "endfoot" => foot = r,
                _ => {
                    has_last = true;
                    tail = r;
                }
            }
        }
        all.extend(rows);
        all.extend(if has_last { tail } else { foot });

        let mut out_rows: Vec<Vec<TableCell>> = Vec::new();
        for row in all {
            if row.len() == 1 && is_blank(&row[0]) {
                continue;
            }
            // A longtable caption row becomes a caption paragraph.
            if row.len() == 1 {
                if let Some(Node::Cmd(c)) = row[0].iter().find(|n| !matches!(n, Node::Space | Node::Par)) {
                    if c.name == "caption" {
                        let saved = self.float.replace("table");
                        self.caption(c, st);
                        self.float = saved;
                        for n in &row[0] {
                            if let Node::Cmd(l) = n {
                                if l.name == "label" {
                                    self.label(l);
                                }
                            }
                        }
                        continue;
                    }
                }
            }
            let mut cells = Vec::new();
            for cell in row {
                let first = cell.iter().find(|n| !matches!(n, Node::Space | Node::Par));
                let (span, content): (u32, Vec<Node>) = match first {
                    Some(Node::Cmd(c)) if c.name == "multicolumn" => {
                        let n = nodes_src(arg(c, 0)).trim().parse::<u32>().unwrap_or(1).clamp(1, 1000);
                        (n, arg(c, 2).to_vec())
                    }
                    _ => (1, cell),
                };
                let mut blocks = self.sub_blocks(&content, st);
                if blocks.is_empty() {
                    blocks.push(Block::P { runs: Vec::new(), align: None });
                }
                cells.push(TableCell { blocks, colspan: (span > 1).then_some(span) });
            }
            out_rows.push(cells);
            if self.limit {
                break;
            }
        }
        if !out_rows.is_empty() {
            self.push_block(Block::Table { rows: out_rows });
        }
    }

    // ---- math ----

    fn display_math(&mut self, env: Option<&str>, body: &[Node], st: &St) {
        self.flush(st.align);
        let name = env.unwrap_or("");
        let base = name.trim_end_matches('*');
        let starred = name.ends_with('*');
        let numbered = !starred && matches!(base, "equation" | "align" | "gather" | "multline" | "eqnarray" | "flalign" | "alignat" | "dmath");
        let per_line = matches!(base, "align" | "gather" | "eqnarray" | "flalign" | "alignat");
        let inner = body.iter().filter(|n| !matches!(n, Node::Space | Node::Par)).collect::<Vec<_>>();
        let content: &[Node] = match inner.as_slice() {
            [Node::Env(e)] if matches!(e.name.as_str(), "split" | "aligned" | "gathered" | "alignedat" | "multlined") => &e.body,
            _ => body,
        };
        let lines = split_lines(content);
        let lines: Vec<&Vec<Node>> = lines.iter().filter(|l| !is_blank(l)).collect();
        let count = lines.len();
        let mut pending_keys: Vec<String> = Vec::new();
        for (k, line) in lines.into_iter().enumerate() {
            let (tag, nonum, keys) = line_meta(line);
            pending_keys.extend(keys);
            let text = self.math(line);
            let last = k + 1 == count;
            let num = match tag {
                Some(t) => Some(t),
                None if numbered && !nonum && (per_line || last) => {
                    self.eq += 1;
                    Some(self.chap_num(self.eq))
                }
                None => None,
            };
            if num.is_some() || per_line || last {
                let n = num.clone().unwrap_or_else(|| self.target.0.clone());
                if num.is_some() {
                    self.target = (n.clone(), "Equation".into());
                }
                for key in pending_keys.drain(..) {
                    self.labels.entry(key).or_insert(Label { num: n.clone(), kind: "Equation".into(), title: String::new() });
                }
            }
            if text.is_empty() && num.is_none() {
                continue;
            }
            let mut runs = Vec::new();
            if !text.is_empty() {
                runs.push(Run { t: text, font: Some(MATH_FONT.into()), ..Default::default() });
            }
            if let Some(n) = num {
                runs.push(Run::text(format!(" ({n})")));
            }
            self.push_block(Block::P { runs: normalize_runs(runs), align: Some("center".into()) });
        }
    }

    fn math(&mut self, nodes: &[Node]) -> String {
        let mut m = MOut::default();
        self.math_into(nodes, &mut m, 0);
        m.finish()
    }

    fn math_str(&mut self, nodes: &[Node], depth: u32) -> String {
        let mut m = MOut::default();
        self.math_into(nodes, &mut m, depth + 1);
        m.finish()
    }

    fn math_into(&mut self, nodes: &[Node], m: &mut MOut, depth: u32) {
        if depth > MAX_DEPTH + 24 || m.s.len() > MAX_TEXT {
            return;
        }
        let mut i = 0;
        while i < nodes.len() {
            let skip_dot = std::mem::take(&mut m.skip_dot);
            match &nodes[i] {
                Node::Text(s) => {
                    let mut first = true;
                    for c in s.chars() {
                        if first && skip_dot && c == '.' {
                            first = false;
                            continue;
                        }
                        first = false;
                        match c {
                            '\'' => m.push('′'),
                            '\u{a0}' => m.push(' '),
                            '`' => {}
                            _ => m.push(c),
                        }
                    }
                }
                Node::Space => {}
                Node::Par | Node::Tab => m.push(' '),
                Node::Group(g) => self.math_into(g, m, depth + 1),
                Node::Sup | Node::Sub => {
                    let sup = matches!(nodes[i], Node::Sup);
                    let mut j = i + 1;
                    while j < nodes.len() && matches!(nodes[j], Node::Space) {
                        j += 1;
                    }
                    if j >= nodes.len() {
                        i = j;
                        continue;
                    }
                    match &nodes[j] {
                        Node::Text(t) => {
                            let mut cs = t.chars();
                            let first = cs.next().unwrap_or(' ');
                            let a = if first == '\'' { "′".to_string() } else { first.to_string() };
                            script(m, sup, &a);
                            for c in cs {
                                match c {
                                    '\'' => m.push('′'),
                                    '\u{a0}' => m.push(' '),
                                    '`' => {}
                                    _ => m.push(c),
                                }
                            }
                        }
                        other => {
                            let a = self.math_str(std::slice::from_ref(other), depth);
                            script(m, sup, &a);
                        }
                    }
                    i = j + 1;
                    continue;
                }
                Node::Cmd(c) => self.math_cmd(c, m, depth),
                Node::Env(e) => self.math_env(e, m, depth),
                Node::Math { body, .. } => self.math_into(body, m, depth + 1),
                Node::Verb(t) => m.push_str(t),
                Node::Verbatim { .. } => {}
            }
            i += 1;
        }
    }

    fn math_cmd(&mut self, c: &CmdNode, m: &mut MOut, depth: u32) {
        let name = c.name.as_str();
        match name {
            "frac" | "dfrac" | "tfrac" | "cfrac" => {
                let (a, b) = if name == "cfrac" { (arg(c, 1), arg(c, 2)) } else { (arg(c, 0), arg(c, 1)) };
                let n = self.math_str(a, depth);
                let d = self.math_str(b, depth);
                m.push_str(&paren_simple(&n));
                m.push('/');
                m.push_str(&paren_simple(&d));
            }
            "binom" | "dbinom" | "tbinom" => {
                let n = self.math_str(arg(c, 0), depth);
                let k = self.math_str(arg(c, 1), depth);
                m.push_str(&format!("C({n},{k})"));
            }
            "sqrt" => {
                let x = self.math_str(arg(c, 1), depth);
                let root = opt(c, 0).map(|o| self.math_str(o, depth)).unwrap_or_default();
                match root.as_str() {
                    "3" => m.push('∛'),
                    "4" => m.push('∜'),
                    "" => m.push('√'),
                    r => {
                        script(m, true, r);
                        m.push('√');
                    }
                }
                if is_simple_root(&x) {
                    m.push_str(&x);
                } else {
                    m.push('(');
                    m.push_str(&x);
                    m.push(')');
                }
            }
            "text" | "textrm" | "mbox" | "hbox" | "textnormal" | "textit" | "textbf" | "emph" | "textsf" | "texttt"
            | "textup" | "intertext" | "textsl" | "textsc" => {
                let t = self.math_text(arg(c, 0));
                m.raw(&t);
            }
            "operatorname" => {
                let t = self.math_str(arg(c, 0), depth);
                m.word(&t);
            }
            "mathrm" | "mathit" | "mathbf" | "mathsf" | "mathtt" | "boldsymbol" | "bm" | "mathnormal" | "pmb"
            | "mathop" | "mathrel" | "mathbin" | "mathord" | "underbrace" | "overbrace" | "boxed" | "cancel"
            | "bcancel" | "xcancel" | "ensuremath" | "displaystyle" | "substack" => self.math_into(arg(c, 0), m, depth + 1),
            "mathbb" | "mathcal" | "mathscr" | "mathfrak" => {
                let style = match name {
                    "mathbb" => "bb",
                    "mathfrak" => "frak",
                    _ => "cal",
                };
                let t = self.math_str(arg(c, 0), depth);
                for ch in t.chars() {
                    m.push(styled_letter(style, ch));
                }
            }
            "hat" | "widehat" | "bar" | "overline" | "vec" | "overrightarrow" | "dot" | "ddot" | "dddot" | "tilde"
            | "widetilde" | "acute" | "grave" | "breve" | "check" | "mathring" | "underline" | "overleftarrow"
            | "overleftrightarrow" | "underleftarrow" | "underrightarrow" => {
                let mark = match name {
                    "hat" | "widehat" => '\u{302}',
                    "bar" => '\u{304}',
                    "overline" => '\u{305}',
                    "vec" | "overrightarrow" => '\u{20D7}',
                    "overleftarrow" => '\u{20D6}',
                    "overleftrightarrow" => '\u{20E1}',
                    "dot" => '\u{307}',
                    "ddot" => '\u{308}',
                    "dddot" => '\u{20DB}',
                    "tilde" | "widetilde" => '\u{303}',
                    "acute" => '\u{301}',
                    "grave" => '\u{300}',
                    "breve" => '\u{306}',
                    "check" => '\u{30C}',
                    "mathring" => '\u{30A}',
                    _ => '\u{332}',
                };
                let x = self.math_str(arg(c, 0), depth);
                let every = matches!(name, "overline" | "underline");
                let n = x.chars().filter(|c| *c != ' ').count();
                let mut k = 0;
                for ch in x.chars() {
                    if ch == ' ' {
                        continue;
                    }
                    k += 1;
                    m.push(ch);
                    if every || k == n {
                        m.s.push(mark);
                    }
                }
            }
            "stackrel" | "overset" | "underset" => self.math_into(arg(c, 1), m, depth + 1),
            "xrightarrow" | "xRightarrow" | "xleftarrow" | "xLeftarrow" | "xleftrightarrow" => {
                m.push(match name {
                    "xrightarrow" => '→',
                    "xRightarrow" => '⇒',
                    "xleftarrow" => '←',
                    "xLeftarrow" => '⇐',
                    _ => '↔',
                });
            }
            "pmod" => {
                m.raw(" (mod ");
                self.math_into(arg(c, 0), m, depth + 1);
                m.push(')');
            }
            "bmod" | "mod" => m.raw(" mod "),
            "left" | "right" | "middle" | "big" | "Big" | "bigg" | "Bigg" | "bigl" | "bigr" | "Bigl" | "Bigr"
            | "biggl" | "biggr" | "Biggl" | "Biggr" | "bigm" | "Bigm" => m.skip_dot = true,
            "not" => m.negate = true,
            "quad" | "qquad" | ";" | ":" | ">" | " " | "enspace" | "space" => m.push(' '),
            "," | "!" | "thinspace" | "negthinspace" | "nobreakspace" | "-" | "/" => {}
            "\\" | "cr" | "newline" => m.raw("; "),
            "textcolor" | "colorbox" => self.math_into(arg(c, 2), m, depth + 1),
            "ref" | "eqref" | "pageref" | "autoref" | "cref" | "Cref" | "nameref" => {
                let key = nodes_src(arg(c, 0)).trim().to_string();
                let op = match name {
                    "eqref" => "eqref",
                    "pageref" => "pageref",
                    "autoref" | "cref" | "Cref" => "autoref",
                    "nameref" => "nameref",
                    _ => "ref",
                };
                m.raw(&placeholder(op, &key));
            }
            "label" | "nonumber" | "notag" | "tag" | "limits" | "nolimits" | "textstyle" | "scriptstyle"
            | "scriptscriptstyle" | "color" | "qedhere" | "hline" | "vphantom" | "hphantom" | "phantom" | "strut"
            | "mathstrut" | "allowbreak" | "nobreak" | "displaybreak" | "boldmath" | "unboldmath" | "relax"
            | "protect" | "noindent" | "centering" | "vspace" | "hspace" | "rule" | "kern" | "mkern" | "mskip" => {}
            "today" => {
                let t = self.today.clone();
                m.raw(&t);
            }
            "'" | "`" | "^" | "\"" | "~" | "=" | "." | "u" | "v" | "H" | "c" | "k" | "r" | "d" | "b" | "t" => {
                let inner = self.math_text(arg(c, 0));
                m.push_str(&accent(name, &inner));
            }
            _ => {
                if let Some(s) = math_sym(name) {
                    m.push_str(s);
                    return;
                }
                if FUNCS.contains(&name) {
                    m.word(name);
                    return;
                }
                if let Some(s) = text_sym(name) {
                    m.push_str(if s.contains('\t') { " " } else { s });
                    return;
                }
                if cmd_spec(name).is_none() {
                    self.unknown_cmd(name);
                }
                for a in c.args.iter().flatten() {
                    self.math_into(a, m, depth + 1);
                }
            }
        }
    }

    fn math_env(&mut self, e: &EnvNode, m: &mut MOut, depth: u32) {
        let base = e.name.trim_end_matches('*');
        let rows = split_rows(&e.body);
        let mut parts = Vec::new();
        for row in rows {
            if row.iter().all(|c| is_blank(c)) {
                continue;
            }
            let cells: Vec<String> = row.iter().map(|c| self.math_str(c, depth)).collect();
            parts.push(cells.join(", "));
        }
        let body = parts.join("; ");
        let (l, r) = match base {
            "pmatrix" | "psmallmatrix" => ("(", ")"),
            "bmatrix" | "bsmallmatrix" => ("[", "]"),
            "Bmatrix" => ("{", "}"),
            "vmatrix" => ("|", "|"),
            "Vmatrix" => ("‖", "‖"),
            "cases" | "dcases" | "rcases" => ("{", ""),
            _ => ("", ""),
        };
        m.push_str(l);
        m.push_str(&body);
        m.push_str(r);
    }

    /// Text inside \text{...} in math (spaces kept).
    fn math_text(&mut self, nodes: &[Node]) -> String {
        let mut s = String::new();
        for n in nodes {
            match n {
                Node::Text(t) => s.push_str(&ligatures(t)),
                Node::Space | Node::Par | Node::Tab => s.push(' '),
                Node::Group(g) => {
                    let t = self.math_text(g);
                    s.push_str(&t);
                }
                Node::Math { body, .. } => {
                    let t = self.math(body);
                    s.push_str(&t);
                }
                Node::Verb(t) => s.push_str(t),
                Node::Cmd(c) => {
                    if let Some(t) = text_sym(&c.name) {
                        s.push_str(if t.contains('\t') { " " } else { t });
                    } else if ACCENTS.iter().any(|(a, ..)| *a == c.name) {
                        let inner = self.math_text(arg(c, 0));
                        s.push_str(&accent(&c.name, &inner));
                    } else if c.name == "\\" {
                        s.push(' ');
                    } else if let Some(t) = math_sym(&c.name) {
                        s.push_str(t);
                    } else if c.name == "today" {
                        s.push_str(&self.today.clone());
                    } else {
                        if cmd_spec(&c.name).is_none() {
                            self.unknown_cmd(&c.name);
                        }
                        let t = self.math_text(last_arg(c));
                        s.push_str(&t);
                    }
                }
                _ => {}
            }
        }
        s
    }

    // ---- finishing ----

    fn resolve_one(&self, op: &str, arg: &str) -> String {
        if op == "cite" {
            let mut parts = arg.split(PLACE_SUB);
            let keys = parts.next().unwrap_or("");
            let pre = parts.next().unwrap_or("").trim();
            let post = parts.next().unwrap_or("").trim();
            let labels: Vec<String> = keys
                .split(',')
                .map(str::trim)
                .filter(|k| !k.is_empty())
                .map(|k| self.bib.get(k).cloned().unwrap_or_else(|| k.to_string()))
                .collect();
            let mut s = String::from("[");
            if !pre.is_empty() {
                s.push_str(pre);
                s.push(' ');
            }
            s.push_str(if labels.is_empty() { "?" } else { "" });
            s.push_str(&labels.join(", "));
            if !post.is_empty() {
                s.push_str(", ");
                s.push_str(post);
            }
            s.push(']');
            return s;
        }
        let parts: Vec<String> = arg
            .split(',')
            .map(str::trim)
            .filter(|k| !k.is_empty())
            .map(|k| {
                let l = self.labels.get(k);
                let num = l.map(|l| l.num.clone()).filter(|n| !n.is_empty()).unwrap_or_else(|| "??".into());
                match op {
                    "eqref" => format!("({num})"),
                    "pageref" => "??".into(),
                    "nameref" => l.map(|l| l.title.clone()).filter(|t| !t.is_empty()).unwrap_or(num),
                    "autoref" | "Autoref" => match l.map(|l| l.kind.as_str()) {
                        Some("Equation") => format!("Equation ({num})"),
                        Some(kind) if !kind.is_empty() => format!("{kind} {num}"),
                        _ => num,
                    },
                    _ => num,
                }
            })
            .collect();
        if parts.is_empty() {
            "??".into()
        } else {
            parts.join(", ")
        }
    }

    fn resolve_text(&self, s: &str) -> String {
        let mut out = String::with_capacity(s.len());
        let mut rest = s;
        while let Some(i) = rest.find(PLACE_OPEN) {
            out.push_str(&rest[..i]);
            let after = &rest[i + PLACE_OPEN.len_utf8()..];
            let Some(j) = after.find(PLACE_CLOSE) else {
                rest = after;
                continue;
            };
            let inner = &after[..j];
            rest = &after[j + PLACE_CLOSE.len_utf8()..];
            let (op, a) = inner.split_once(PLACE_SEP).unwrap_or((inner, ""));
            out.push_str(&self.resolve_one(op, a));
        }
        out.push_str(rest);
        out.retain(|c| !matches!(c, PLACE_OPEN | PLACE_CLOSE | PLACE_SEP | PLACE_SUB));
        out
    }

    fn resolve_blocks(&self, blocks: &mut [Block], depth: u32) {
        if depth > MAX_DEPTH + 8 {
            return;
        }
        for b in blocks {
            if let Block::Table { rows } = b {
                for row in rows {
                    for cell in row {
                        self.resolve_blocks(&mut cell.blocks, depth + 1);
                    }
                }
            } else if let Some(runs) = runs_mut(b) {
                for r in runs.iter_mut() {
                    if r.t.contains(PLACE_OPEN) {
                        r.t = self.resolve_text(&r.t);
                    }
                }
            }
        }
    }

    fn meta_text(&mut self, nodes: &[Node]) -> String {
        let filtered: Vec<Node> = nodes
            .iter()
            .filter(|n| !matches!(n, Node::Cmd(c) if matches!(c.name.as_str(), "thanks" | "footnote" | "footnotemark" | "inst")))
            .cloned()
            .collect();
        let t = self.plain_of(&filtered);
        let t = self.resolve_text(&t);
        t.split_whitespace().collect::<Vec<_>>().join(" ")
    }

    fn finish(mut self, macro_limit: bool, truncated: bool) -> WordDoc {
        self.flush(None);
        self.lists.clear();
        self.quote = 0;
        if !self.footnotes.is_empty() {
            let notes = std::mem::take(&mut self.footnotes);
            self.out.push(Block::H1 { runs: vec![Run::text("Notes")], align: None });
            for (n, runs) in notes {
                let mut rs = vec![Run { t: format!("[{n}]"), sup: true, ..Default::default() }, Run::text(" ")];
                rs.extend(runs);
                self.out.push(Block::P { runs: normalize_runs(rs), align: None });
            }
        }
        let mut meta = Map::new();
        if let Some(t) = self.title.clone() {
            let s = self.meta_text(&t);
            if !s.is_empty() {
                meta.insert("title".into(), Value::String(s));
            }
        }
        if let Some(a) = self.author.clone() {
            let s = self.meta_text(&a);
            if !s.is_empty() {
                meta.insert("author".into(), Value::String(s));
            }
        }
        let mut blocks = std::mem::take(&mut self.out);
        self.resolve_blocks(&mut blocks, 0);
        if macro_limit {
            self.warn(W_MACRO);
        }
        if truncated {
            self.warn(W_LIMIT);
        }
        if !self.unknown.is_empty() {
            let shown: Vec<String> = self
                .unknown
                .iter()
                .take(20)
                .map(|n| if n.starts_with("begin{") { format!("\\{n}") } else { format!("\\{n}") })
                .collect();
            let more = self.unknown.len().saturating_sub(20);
            let tail = if more > 0 { format!(" and {more} more") } else { String::new() };
            let msg = format!(
                "Some LaTeX commands are not supported and were dropped (their text was kept): {}{tail}.",
                shown.join(", ")
            );
            self.warnings.insert(0, msg);
        }
        for w in &self.warnings {
            push_warning(&mut meta, w);
        }
        if blocks.is_empty() {
            blocks.push(Block::P { runs: Vec::new(), align: None });
        }
        WordDoc { kind: "word".into(), meta, blocks }
    }
}

fn split_lines(nodes: &[Node]) -> Vec<Vec<Node>> {
    let mut out = vec![Vec::new()];
    for n in nodes {
        match n {
            Node::Cmd(c) if matches!(c.name.as_str(), "\\" | "cr") => out.push(Vec::new()),
            _ => {
                if let Some(l) = out.last_mut() {
                    l.push(n.clone());
                }
            }
        }
    }
    out
}

fn split_rows(nodes: &[Node]) -> Vec<Vec<Vec<Node>>> {
    let mut rows: Vec<Vec<Vec<Node>>> = vec![vec![Vec::new()]];
    for n in nodes {
        match n {
            Node::Cmd(c) if matches!(c.name.as_str(), "\\" | "cr") => rows.push(vec![Vec::new()]),
            Node::Cmd(c) if matches!(c.name.as_str(), "hline" | "noalign") => {}
            Node::Tab => {
                if let Some(r) = rows.last_mut() {
                    r.push(Vec::new());
                }
            }
            _ => {
                if let Some(c) = rows.last_mut().and_then(|r| r.last_mut()) {
                    c.push(n.clone());
                }
            }
        }
    }
    rows
}

/// \tag, \nonumber/\notag and \label keys of one display line.
fn line_meta(line: &[Node]) -> (Option<String>, bool, Vec<String>) {
    let mut tag = None;
    let mut nonum = false;
    let mut keys = Vec::new();
    fn rec(nodes: &[Node], tag: &mut Option<String>, nonum: &mut bool, keys: &mut Vec<String>, depth: u32) {
        if depth > 8 {
            return;
        }
        for n in nodes {
            match n {
                Node::Cmd(c) => match c.name.as_str() {
                    "nonumber" | "notag" => *nonum = true,
                    "tag" => *tag = Some(nodes_src(arg(c, 0)).trim().to_string()),
                    "label" => keys.push(nodes_src(arg(c, 0)).trim().to_string()),
                    _ => {}
                },
                Node::Group(g) => rec(g, tag, nonum, keys, depth + 1),
                _ => {}
            }
        }
    }
    rec(line, &mut tag, &mut nonum, &mut keys, 0);
    keys.retain(|k| !k.is_empty());
    (tag, nonum, keys)
}

#[derive(Default)]
struct MOut {
    s: String,
    need_space: bool,
    negate: bool,
    skip_dot: bool,
}

impl MOut {
    fn push(&mut self, c: char) {
        if self.negate && c != ' ' {
            self.negate = false;
            let n = match c {
                '=' => Some('≠'),
                '∈' => Some('∉'),
                '<' => Some('≮'),
                '>' => Some('≯'),
                '≡' => Some('≢'),
                '⊂' => Some('⊄'),
                '⊆' => Some('⊈'),
                '∃' => Some('∄'),
                '∼' => Some('≁'),
                '≤' => Some('≰'),
                '≥' => Some('≱'),
                '≅' => Some('≇'),
                _ => None,
            };
            if let Some(n) = n {
                self.push_inner(n);
            } else {
                self.push_inner(c);
                self.s.push('\u{338}');
            }
            return;
        }
        self.push_inner(c);
    }

    fn push_inner(&mut self, c: char) {
        if self.need_space && c.is_alphanumeric() {
            self.s.push(' ');
        }
        self.need_space = false;
        if is_rel(c) {
            if !self.s.ends_with(' ') {
                self.s.push(' ');
            }
            self.s.push(c);
            self.s.push(' ');
            return;
        }
        self.s.push(c);
    }

    fn push_str(&mut self, t: &str) {
        for c in t.chars() {
            self.push(c);
        }
    }

    /// Text pushed as is (no relation spacing).
    fn raw(&mut self, t: &str) {
        if self.need_space && t.starts_with(|c: char| c.is_alphanumeric()) {
            self.s.push(' ');
        }
        self.need_space = false;
        self.s.push_str(t);
    }

    fn word(&mut self, w: &str) {
        self.raw(w);
        self.need_space = true;
    }

    fn finish(self) -> String {
        let mut out = String::with_capacity(self.s.len());
        let mut prev_space = true;
        for c in self.s.chars() {
            if c == ' ' {
                if !prev_space {
                    out.push(' ');
                }
                prev_space = true;
            } else {
                out.push(c);
                prev_space = false;
            }
        }
        while out.ends_with(' ') {
            out.pop();
        }
        out
    }
}

fn script(m: &mut MOut, sup: bool, s: &str) {
    let s = s.trim();
    if s.is_empty() {
        return;
    }
    if sup && s == "∘" {
        m.push('°');
        return;
    }
    if s.chars().all(|c| c == '′') {
        m.raw(s);
        return;
    }
    let compact: String = s.chars().filter(|c| *c != ' ').collect();
    let mapped: Option<String> = compact.chars().map(|c| if sup { sup_of(c) } else { sub_of(c) }).collect();
    if let Some(mm) = mapped {
        // keep a pending operator space: "min" + subscript + "f" -> "minₓ f"
        m.s.push_str(&mm);
        return;
    }
    m.need_space = false;
    m.s.push(if sup { '^' } else { '_' });
    if compact.chars().count() == 1 {
        m.s.push_str(&compact);
    } else {
        m.s.push('(');
        m.s.push_str(s);
        m.s.push(')');
    }
}

fn paren_simple(s: &str) -> String {
    if s.chars().count() <= 1 || s.chars().all(|c| c.is_alphanumeric()) {
        s.to_string()
    } else {
        format!("({s})")
    }
}

fn is_simple_root(s: &str) -> bool {
    let n = s.chars().count();
    n == 1 || (n > 0 && (s.chars().all(|c| c.is_ascii_digit()) || s.chars().all(|c| c.is_alphabetic())))
}

fn styled_letter(style: &str, c: char) -> char {
    let off = |base: u32, from: char| char::from_u32(base + (c as u32 - from as u32)).unwrap_or(c);
    match style {
        "bb" => match c {
            'C' => 'ℂ',
            'H' => 'ℍ',
            'N' => 'ℕ',
            'P' => 'ℙ',
            'Q' => 'ℚ',
            'R' => 'ℝ',
            'Z' => 'ℤ',
            'A'..='Z' => off(0x1D538, 'A'),
            _ => c,
        },
        "frak" => match c {
            'C' => 'ℭ',
            'H' => 'ℌ',
            'I' => 'ℑ',
            'R' => 'ℜ',
            'Z' => 'ℨ',
            'A'..='Z' => off(0x1D504, 'A'),
            'a'..='z' => off(0x1D51E, 'a'),
            _ => c,
        },
        _ => match c {
            'B' => 'ℬ',
            'E' => 'ℰ',
            'F' => 'ℱ',
            'H' => 'ℋ',
            'I' => 'ℐ',
            'L' => 'ℒ',
            'M' => 'ℳ',
            'R' => 'ℛ',
            'A'..='Z' => off(0x1D49C, 'A'),
            _ => c,
        },
    }
}

/// Reverse of styled_letter plus Word's math italic/bold letters:
/// Some((style, base)) with style "" for plain.
fn unstyle_letter(c: char) -> Option<(&'static str, char)> {
    let specials: &[(char, &str, char)] = &[
        ('ℂ', "bb", 'C'), ('ℍ', "bb", 'H'), ('ℕ', "bb", 'N'), ('ℙ', "bb", 'P'), ('ℚ', "bb", 'Q'), ('ℝ', "bb", 'R'),
        ('ℤ', "bb", 'Z'), ('ℭ', "frak", 'C'), ('ℌ', "frak", 'H'), ('ℨ', "frak", 'Z'), ('ℬ', "cal", 'B'),
        ('ℰ', "cal", 'E'), ('ℱ', "cal", 'F'), ('ℋ', "cal", 'H'), ('ℐ', "cal", 'I'), ('ℒ', "cal", 'L'),
        ('ℳ', "cal", 'M'), ('ℛ', "cal", 'R'), ('ℎ', "", 'h'), ('ℯ', "", 'e'), ('ℊ', "", 'g'), ('ℴ', "", 'o'),
    ];
    if let Some((_, s, b)) = specials.iter().find(|(ch, ..)| *ch == c) {
        return Some((s, *b));
    }
    let u = c as u32;
    let az = |base: u32| char::from_u32(u - base + 'A' as u32);
    let lz = |base: u32| char::from_u32(u - base + 'a' as u32);
    match u {
        0x1D538..=0x1D551 => az(0x1D538).map(|b| ("bb", b)),
        0x1D504..=0x1D51D => az(0x1D504).map(|b| ("frak", b)),
        0x1D51E..=0x1D537 => lz(0x1D51E).map(|b| ("frak", b)),
        0x1D49C..=0x1D4B5 => az(0x1D49C).map(|b| ("cal", b)),
        // bold, italic, bold italic, script lowercase, bb lowercase, sans... -> plain
        0x1D400..=0x1D6A3 => {
            let k = (u - 0x1D400) % 52;
            let b = if k < 26 { char::from_u32('A' as u32 + k) } else { char::from_u32('a' as u32 + k - 26) };
            b.map(|b| ("", b))
        }
        0x1D7CE..=0x1D7FF => char::from_u32('0' as u32 + (u - 0x1D7CE) % 10).map(|b| ("", b)),
        _ => None,
    }
}

fn parse_inner(src: &str) -> WordDoc {
    let mut b = Builder::new(src);
    let nodes = b.build();
    let mut conv = Conv::new();
    conv.run(&nodes);
    conv.finish(b.lx.limit_hit, b.lx.truncated)
}

fn fallback_doc(src: &str) -> WordDoc {
    let mut blocks = Vec::new();
    for para in src.split("\n\n").take(100_000) {
        let t = para.trim();
        if !t.is_empty() {
            blocks.push(Block::P { runs: vec![Run::text(t.chars().take(1_000_000).collect::<String>())], align: None });
        }
    }
    if blocks.is_empty() {
        blocks.push(Block::P { runs: Vec::new(), align: None });
    }
    let mut meta = Map::new();
    push_warning(&mut meta, W_FALLBACK);
    WordDoc { kind: "word".into(), meta, blocks }
}

// ============================ writer ============================

#[derive(Default)]
struct Writer {
    math: bool,
    color: bool,
    ulem: bool,
    longtable: bool,
    textcomp: bool,
    uni: BTreeSet<char>,
}

#[derive(Clone, Copy, PartialEq)]
enum Ctx {
    Para,
    Heading,
    Cell,
}

/// Latin letters pdflatex (utf8 + T1) typesets directly.
fn latin_ok(c: char) -> bool {
    let u = c as u32;
    ((0xC0..=0xFF).contains(&u) && u != 0xD7 && u != 0xF7)
        || ((0x100..=0x17F).contains(&u) && !matches!(u, 0x126 | 0x127 | 0x138 | 0x13F | 0x140 | 0x149 | 0x166 | 0x167 | 0x17F))
}

/// Text-mode command for a character: (LaTeX, needs textcomp).
fn text_cmd_for(c: char, mono: bool) -> Option<(&'static str, bool)> {
    Some(match c {
        '“' => (if mono { "\\textquotedblleft{}" } else { "``" }, false),
        '”' => (if mono { "\\textquotedblright{}" } else { "''" }, false),
        '‘' => (if mono { "\\textquoteleft{}" } else { "`" }, false),
        '’' => (if mono { "\\textquoteright{}" } else { "'" }, false),
        '–' => (if mono { "\\textendash{}" } else { "--" }, false),
        '—' => (if mono { "\\textemdash{}" } else { "---" }, false),
        '…' => ("\\ldots{}", false),
        '\u{a0}' => ("~", false),
        '\u{2009}' | '\u{202f}' => ("\\,", false),
        '\u{2002}' => ("\\enspace{}", false),
        '\u{2003}' | '\t' => ("\\quad{}", false),
        '€' => ("\\texteuro{}", true),
        '™' => ("\\texttrademark{}", true),
        '©' => ("\\textcopyright{}", true),
        '®' => ("\\textregistered{}", true),
        '°' => ("\\textdegree{}", true),
        '℃' => ("\\textcelsius{}", true),
        '№' => ("\\textnumero{}", true),
        '§' => ("\\S{}", false),
        '¶' => ("\\P{}", false),
        '†' => ("\\dag{}", false),
        '‡' => ("\\ddag{}", false),
        '•' => ("\\textbullet{}", true),
        '£' => ("\\pounds{}", false),
        '«' => ("\\guillemotleft{}", false),
        '»' => ("\\guillemotright{}", false),
        '‹' => ("\\guilsinglleft{}", false),
        '›' => ("\\guilsinglright{}", false),
        '„' => ("\\quotedblbase{}", false),
        '‚' => ("\\quotesinglbase{}", false),
        '¡' => ("\\textexclamdown{}", false),
        '¿' => ("\\textquestiondown{}", false),
        '×' => ("\\texttimes{}", true),
        '÷' => ("\\textdiv{}", true),
        '±' => ("\\textpm{}", true),
        'µ' => ("\\textmu{}", true),
        '¹' => ("\\textonesuperior{}", true),
        '²' => ("\\texttwosuperior{}", true),
        '³' => ("\\textthreesuperior{}", true),
        '¼' => ("\\textonequarter{}", true),
        '½' => ("\\textonehalf{}", true),
        '¾' => ("\\textthreequarters{}", true),
        '¬' => ("\\textlnot{}", true),
        '¢' => ("\\textcent{}", true),
        '¥' => ("\\textyen{}", true),
        '¤' => ("\\textcurrency{}", true),
        '¦' => ("\\textbrokenbar{}", true),
        '¨' => ("\\textasciidieresis{}", true),
        '¯' => ("\\textasciimacron{}", true),
        '´' => ("\\textasciiacute{}", true),
        'ª' => ("\\textordfeminine{}", true),
        'º' => ("\\textordmasculine{}", true),
        '·' => ("\\textperiodcentered{}", true),
        '‰' => ("\\textperthousand{}", true),
        '␣' => ("\\textvisiblespace{}", false),
        'ı' => ("\\i{}", false),
        'ȷ' => ("\\j{}", false),
        _ => return None,
    })
}

fn math_cmd_for(c: char) -> Option<&'static str> {
    if c.is_ascii() {
        return None;
    }
    let mut buf = [0u8; 4];
    let s: &str = c.encode_utf8(&mut buf);
    MATH_SYMS.iter().find(|(_, u)| *u == s).map(|(n, _)| *n)
}

fn comb_cmd(c: char) -> Option<&'static str> {
    Some(match c {
        '\u{302}' => "hat",
        '\u{304}' => "bar",
        '\u{305}' => "overline",
        '\u{20D7}' => "vec",
        '\u{307}' => "dot",
        '\u{308}' => "ddot",
        '\u{303}' => "tilde",
        '\u{301}' => "acute",
        '\u{300}' => "grave",
        '\u{306}' => "breve",
        '\u{30C}' => "check",
        '\u{30A}' => "mathring",
        '\u{332}' => "underline",
        '\u{20D6}' => "overleftarrow",
        '\u{20E1}' => "overleftrightarrow",
        '\u{20DB}' => "dddot",
        '\u{338}' => "not",
        _ => return None,
    })
}

fn ends_with_control_word(s: &str) -> bool {
    let b = s.as_bytes();
    let mut i = b.len();
    // our control words are short; bounded scan keeps long words O(1)
    let stop = b.len().saturating_sub(48);
    while i > stop && b[i - 1].is_ascii_alphabetic() {
        i -= 1;
    }
    if i == stop && stop > 0 && b[i - 1].is_ascii_alphabetic() {
        return false;
    }
    if i == b.len() || i == 0 || b[i - 1] != b'\\' {
        return false;
    }
    // the backslash must not itself be escaped
    let mut k = i - 1;
    let mut n = 0;
    while k > 0 && b[k - 1] == b'\\' {
        k -= 1;
        n += 1;
    }
    n % 2 == 0
}

/// Append LaTeX text, keeping TeX from forming ligatures or swallowing the
/// next letter into a control word.
fn append_guarded(out: &mut String, piece: &str) {
    if let (Some(last), Some(first)) = (out.chars().last(), piece.chars().next()) {
        let lig = (last == '-' && first == '-')
            || (last == '\'' && first == '\'')
            || (last == '`' && first == '`')
            || ((last == '!' || last == '?') && first == '`')
            || (last == ',' && first == ',')
            || (last == '<' && first == '<')
            || (last == '>' && first == '>');
        if lig {
            out.push_str("{}");
        } else if first.is_ascii_alphabetic() && ends_with_control_word(out) {
            out.push(' ');
        }
    }
    out.push_str(piece);
}

fn match_paren(chars: &[char], open: usize) -> Option<usize> {
    let mut depth = 0i32;
    for (k, &c) in chars.iter().enumerate().skip(open) {
        match c {
            '(' => depth += 1,
            ')' => {
                depth -= 1;
                if depth == 0 {
                    return Some(k);
                }
            }
            _ => {}
        }
    }
    None
}

fn root_operand(chars: &[char], k: usize) -> Option<(String, usize)> {
    let c = *chars.get(k)?;
    if c == '(' {
        let j = match_paren(chars, k)?;
        return Some((chars[k + 1..j].iter().collect(), j + 1));
    }
    if c.is_ascii_digit() {
        let mut e = k;
        while e < chars.len() && chars[e].is_ascii_digit() {
            e += 1;
        }
        return Some((chars[k..e].iter().collect(), e));
    }
    if c.is_alphabetic() {
        let mut e = k;
        while e < chars.len() && chars[e].is_alphabetic() {
            e += 1;
        }
        return Some((chars[k..e].iter().collect(), e));
    }
    if c == ' ' {
        return None;
    }
    Some((c.to_string(), k + 1))
}

/// Unicode approximation of a LaTeX math snippet (the reader's view).
fn approx_math_str(tex: &str) -> String {
    let src = format!("\\({tex}\\)");
    let mut b = Builder::new(&src);
    let nodes = b.build();
    let mut c = Conv::new();
    for n in &nodes {
        if let Node::Math { body, .. } = n {
            return c.math(body);
        }
    }
    String::new()
}

fn meta_str(doc: &WordDoc, key: &str) -> Option<String> {
    let s = doc.meta.get(key)?.as_str()?;
    let s = s.split_whitespace().collect::<Vec<_>>().join(" ");
    (!s.is_empty()).then_some(s)
}

fn is_math_run(r: &Run) -> bool {
    r.font.as_deref() == Some(MATH_FONT)
}

fn is_plain_run(r: &Run) -> bool {
    crate::docx::same_format(r, &Run::default())
}

fn image_placeholder(runs: &[Run]) -> Option<&str> {
    match runs {
        [r] if is_plain_run(r) => r.t.strip_prefix("[Image: ").and_then(|t| t.strip_suffix(']')),
        _ => None,
    }
}

fn is_display_math(runs: &[Run]) -> bool {
    runs.iter().any(|r| is_math_run(r) && !r.t.trim().is_empty())
        && runs.iter().all(|r| is_math_run(r) || r.t.trim().is_empty())
}

impl Writer {
    fn write(mut self, doc: &WordDoc) -> String {
        let blocks: Vec<Cow<Block>> = doc
            .blocks
            .iter()
            .filter_map(resolve_block)
            .filter(|b| !matches!(b.as_ref(), Block::Other(_)))
            .collect();
        let title = meta_str(doc, "title");
        let author = meta_str(doc, "author");
        let mut i = 0;
        let mut date: Option<String> = None;
        if let Some(t) = &title {
            if let Some(Block::H1 { runs, align: Some(a) }) = blocks.first().map(|b| b.as_ref()) {
                let no_marks: Vec<Run> = runs.iter().filter(|r| !r.sup).cloned().collect();
                if a == "center" && plain(&heading_runs(no_marks)) == *t {
                    i = 1;
                    if let (Some(au), Some(Block::P { runs, align: Some(a) })) = (&author, blocks.get(1).map(|b| b.as_ref())) {
                        if a == "center" {
                            let p = plain(&normalize_runs(runs.clone()));
                            if p == *au {
                                i = 2;
                            } else if let Some(rest) = p.strip_prefix(&format!("{au}\n")) {
                                date = Some(rest.replace('\n', " "));
                                i = 2;
                            }
                        }
                    }
                }
            }
        }

        let mut chunks: Vec<String> = Vec::new();
        while i < blocks.len() {
            let b = blocks[i].as_ref();
            match b {
                Block::List { .. } => {
                    let start = i;
                    while i < blocks.len() && matches!(blocks[i].as_ref(), Block::List { .. }) {
                        i += 1;
                    }
                    let c = self.lists(&blocks[start..i]);
                    if !c.is_empty() {
                        chunks.push(c);
                    }
                    continue;
                }
                Block::Code { .. } => {
                    let mut lines: Vec<String> = Vec::new();
                    while i < blocks.len() {
                        let Block::Code { runs, .. } = blocks[i].as_ref() else { break };
                        let t: String = runs.iter().map(|r| r.t.as_str()).collect();
                        for l in t.replace("\r\n", "\n").replace('\r', "\n").split('\n') {
                            lines.push(self.verb_line(l));
                        }
                        i += 1;
                    }
                    chunks.push(format!("\\begin{{verbatim}}\n{}\n\\end{{verbatim}}", lines.join("\n")));
                    continue;
                }
                Block::Quote { .. } => {
                    let mut paras = Vec::new();
                    while i < blocks.len() {
                        let Block::Quote { runs, .. } = blocks[i].as_ref() else { break };
                        let t = self.para_tex(&normalize_runs(runs.clone()), Ctx::Para);
                        if !t.is_empty() {
                            paras.push(t);
                        }
                        i += 1;
                    }
                    if !paras.is_empty() {
                        chunks.push(format!("\\begin{{quote}}\n{}\n\\end{{quote}}", paras.join("\n\n")));
                    }
                    continue;
                }
                Block::P { runs, align: Some(a) } if a == "center" || a == "right" => {
                    let runs = normalize_runs(runs.clone());
                    if a == "center" && is_display_math(&runs) {
                        let t: String = runs.iter().map(|r| r.t.as_str()).collect();
                        let tex = self.math_tex(t.trim());
                        chunks.push(format!("\\[ {tex} \\]"));
                        i += 1;
                        continue;
                    }
                    let al = a.clone();
                    let mut paras = Vec::new();
                    while i < blocks.len() {
                        let Block::P { runs, align: Some(a2) } = blocks[i].as_ref() else { break };
                        if *a2 != al {
                            break;
                        }
                        let runs = normalize_runs(runs.clone());
                        if al == "center" && is_display_math(&runs) {
                            break;
                        }
                        if let Some(t) = self.para_chunk(&runs) {
                            paras.push(t);
                        }
                        i += 1;
                    }
                    if !paras.is_empty() {
                        let env = if al == "center" { "center" } else { "flushright" };
                        chunks.push(format!("\\begin{{{env}}}\n{}\n\\end{{{env}}}", paras.join("\n\n")));
                    }
                    continue;
                }
                Block::P { runs, .. } => {
                    if let Some(t) = self.para_chunk(&normalize_runs(runs.clone())) {
                        chunks.push(t);
                    }
                }
                Block::H1 { runs, align } | Block::H2 { runs, align } | Block::H3 { runs, align } | Block::H4 { runs, align }
                | Block::H5 { runs, align } | Block::H6 { runs, align } => {
                    let hr = heading_runs(runs.clone());
                    if align.as_deref() == Some("center") && plain(&hr) == "Abstract" {
                        let mut paras = Vec::new();
                        i += 1;
                        while i < blocks.len() {
                            let Block::P { runs, align: None } = blocks[i].as_ref() else { break };
                            if let Some(t) = self.para_chunk(&normalize_runs(runs.clone())) {
                                paras.push(t);
                            }
                            i += 1;
                        }
                        chunks.push(format!("\\begin{{abstract}}\n{}\n\\end{{abstract}}", paras.join("\n\n")));
                        continue;
                    }
                    let level = match b {
                        Block::H1 { .. } => "section",
                        Block::H2 { .. } => "subsection",
                        Block::H3 { .. } => "subsubsection",
                        Block::H4 { .. } => "paragraph",
                        _ => "subparagraph",
                    };
                    let t = self.para_tex(&hr, Ctx::Heading);
                    if !t.is_empty() {
                        chunks.push(format!("\\{level}{{{t}}}"));
                    }
                }
                Block::Table { rows } => {
                    if let Some(t) = self.table(rows) {
                        chunks.push(t);
                    }
                }
                Block::Img { alt, .. } => {
                    let alt = alt.as_deref().map(|a| a.split_whitespace().collect::<Vec<_>>().join(" ")).filter(|a| !a.is_empty());
                    let alt: String = alt.unwrap_or_else(|| "picture".into()).chars().take(200).collect();
                    chunks.push(self.image_tex(&alt));
                }
                Block::PageBreak => chunks.push("\\newpage".into()),
                Block::Other(_) => {}
            }
            i += 1;
        }

        let mut out = String::new();
        out.push_str("% Created by Office Editor\n\\documentclass{article}\n\\usepackage[utf8]{inputenc}\n\\usepackage[T1]{fontenc}\n\\usepackage{lmodern}\n");
        let head_title = title.as_ref().map(|t| self.esc(t, false));
        let head_author = author.as_ref().map(|a| self.esc(a, false)).unwrap_or_default();
        let head_date = date.as_ref().map(|d| self.esc(d, false)).unwrap_or_default();
        if self.textcomp {
            out.push_str("\\usepackage{textcomp}\n");
        }
        if self.math {
            out.push_str("\\usepackage{amsmath}\n\\usepackage{amssymb}\n");
        }
        if self.color {
            out.push_str("\\usepackage{xcolor}\n");
        }
        if self.ulem {
            out.push_str("\\usepackage[normalem]{ulem}\n");
        }
        if self.longtable {
            out.push_str("\\usepackage{longtable}\n");
        }
        out.push_str("\\usepackage{hyperref}\n\\setcounter{secnumdepth}{0}\n");
        if !self.uni.is_empty() {
            out.push_str("% Characters pdfLaTeX cannot typeset print as \"?\" (they are kept in this file).\n\\ifdefined\\DeclareUnicodeCharacter\n");
            for c in &self.uni {
                out.push_str(&format!("\\DeclareUnicodeCharacter{{{:04X}}}{{?}}\n", *c as u32));
            }
            out.push_str("\\fi\n");
        }
        if let Some(t) = &head_title {
            out.push_str(&format!("\n\\title{{{t}}}\n\\author{{{head_author}}}\n\\date{{{head_date}}}\n"));
        }
        out.push_str("\n\\begin{document}\n");
        if head_title.is_some() {
            out.push_str("\\maketitle\n");
        }
        for c in &chunks {
            out.push('\n');
            out.push_str(c);
            out.push('\n');
        }
        out.push_str("\n\\end{document}\n");
        out
    }

    fn verb_line(&mut self, l: &str) -> String {
        let mut s: String = l
            .chars()
            .filter(|c| !c.is_control() || *c == '\t')
            .inspect(|c| {
                if !c.is_ascii() && !latin_ok(*c) {
                    self.uni.insert(*c);
                }
            })
            .collect();
        while let Some(p) = s.find("\\end{verbatim}") {
            s.replace_range(p..p + 14, "\\end {verbatim}");
        }
        s
    }

    fn image_tex(&mut self, alt: &str) -> String {
        let e = self.esc(alt, false);
        format!("% image omitted: {e}\n\\fbox{{Image: {e}}}")
    }

    /// A normal paragraph; None when it has no visible content.
    fn para_chunk(&mut self, runs: &[Run]) -> Option<String> {
        if let Some(alt) = image_placeholder(runs) {
            return Some(self.image_tex(alt));
        }
        let t = self.para_tex(runs, Ctx::Para);
        (!t.is_empty()).then_some(t)
    }

    fn para_tex(&mut self, runs: &[Run], ctx: Ctx) -> String {
        let runs: Vec<Run> = if ctx == Ctx::Para {
            runs.to_vec()
        } else {
            heading_runs(runs.to_vec())
        };
        let mut lines: Vec<Vec<Run>> = vec![Vec::new()];
        for r in &runs {
            for (k, part) in r.t.split('\n').enumerate() {
                if k > 0 {
                    lines.push(Vec::new());
                }
                if !part.is_empty() {
                    let mut p = r.clone();
                    p.t = part.to_string();
                    if let Some(l) = lines.last_mut() {
                        l.push(p);
                    }
                }
            }
        }
        let mut out = String::new();
        for (k, line) in lines.iter().enumerate() {
            let t = self.line_tex(line);
            if k > 0 {
                out.push_str("\\\\");
                if t.starts_with('[') || t.starts_with('*') {
                    out.push_str("{}");
                }
                out.push('\n');
            }
            out.push_str(&t);
        }
        out
    }

    fn line_tex(&mut self, runs: &[Run]) -> String {
        let mut out = String::new();
        for r in runs {
            if is_math_run(r) {
                let core = r.t.trim_matches(' ');
                if r.t.starts_with(' ') {
                    append_guarded(&mut out, " ");
                }
                if !core.is_empty() {
                    let tex = self.math_tex(core);
                    if !tex.is_empty() {
                        let w = self.wrap(r, format!("${tex}$"));
                        append_guarded(&mut out, &w);
                    }
                }
                if r.t.ends_with(' ') && !core.is_empty() {
                    append_guarded(&mut out, " ");
                }
                continue;
            }
            let mono = r.font.as_deref().is_some_and(crate::docx::is_mono_font);
            let inner = self.esc(&r.t, mono);
            if inner.is_empty() {
                continue;
            }
            let w = self.wrap(r, inner);
            append_guarded(&mut out, &w);
        }
        out
    }

    fn wrap(&mut self, r: &Run, inner: String) -> String {
        let mut s = inner;
        if !is_math_run(r) && r.font.as_deref().is_some_and(crate::docx::is_mono_font) {
            s = format!("\\texttt{{{s}}}");
        }
        if r.sub {
            s = format!("\\textsubscript{{{s}}}");
        }
        if r.sup {
            s = format!("\\textsuperscript{{{s}}}");
        }
        if r.s {
            self.ulem = true;
            s = format!("\\sout{{{s}}}");
        }
        if r.u {
            s = format!("\\underline{{{s}}}");
        }
        if r.i {
            s = format!("\\textit{{{s}}}");
        }
        if r.b {
            s = format!("\\textbf{{{s}}}");
        }
        if let Some(c) = r.color.as_deref().and_then(norm_hex) {
            self.color = true;
            s = format!("\\textcolor[HTML]{{{c}}}{{{s}}}");
        }
        if let Some(h) = r.hl.as_deref().and_then(norm_hex) {
            self.color = true;
            s = format!("\\colorbox[HTML]{{{h}}}{{{s}}}");
        }
        if let Some(l) = r.link.as_deref().and_then(safe_link) {
            s = format!("\\href{{{}}}{{{s}}}", url_tex(&l));
        }
        s
    }

    /// Escape text for LaTeX text mode.
    fn esc(&mut self, s: &str, mono: bool) -> String {
        let mut out = String::with_capacity(s.len() + 8);
        for c in s.chars() {
            let piece: Cow<str> = match c {
                '\\' => "\\textbackslash{}".into(),
                '{' => "\\{".into(),
                '}' => "\\}".into(),
                '#' => "\\#".into(),
                '$' => "\\$".into(),
                '%' => "\\%".into(),
                '&' => "\\&".into(),
                '_' => "\\_".into(),
                '~' => "\\textasciitilde{}".into(),
                '^' => "\\textasciicircum{}".into(),
                '<' => "\\textless{}".into(),
                '>' => "\\textgreater{}".into(),
                '|' => "\\textbar{}".into(),
                '"' => "\\textquotedbl{}".into(),
                '`' => "\\textasciigrave{}".into(),
                '\'' if mono => "\\textquotesingle{}".into(),
                '\n' => " ".into(),
                '\t' => "\\quad{}".into(),
                '\u{ad}' | '\u{200b}' | '\u{200c}' | '\u{200d}' | '\u{2060}' | '\u{feff}' => continue,
                c if c.is_control() && c != '\t' => continue,
                c if c.is_ascii() => c.to_string().into(),
                c => {
                    if let Some((cmd, tc)) = text_cmd_for(c, mono) {
                        if tc {
                            self.textcomp = true;
                        }
                        cmd.into()
                    } else if latin_ok(c) {
                        c.to_string().into()
                    } else if let Some(cmd) = math_cmd_for(c) {
                        self.math = true;
                        format!("\\ensuremath{{\\{cmd}}}").into()
                    } else if let Some(b) = sup_base(c) {
                        self.math = true;
                        let t = self.math_char(b);
                        format!("\\ensuremath{{^{{{t}}}}}").into()
                    } else if let Some(b) = sub_base(c) {
                        self.math = true;
                        let t = self.math_char(b);
                        format!("\\ensuremath{{_{{{t}}}}}").into()
                    } else if let Some((style, b)) = unstyle_letter(c) {
                        match style {
                            "bb" | "cal" | "frak" => {
                                self.math = true;
                                format!("\\ensuremath{{\\math{style}{{{b}}}}}").into()
                            }
                            _ => b.to_string().into(),
                        }
                    } else {
                        self.uni.insert(c);
                        c.to_string().into()
                    }
                }
            };
            append_guarded(&mut out, &piece);
        }
        out
    }

    /// Canonical LaTeX for a math run: iterate to a fixpoint of
    /// tex -> reader approximation -> tex so the round trip is stable.
    fn math_tex(&mut self, t: &str) -> String {
        self.math = true;
        let mut cur = self.math_to_tex(t);
        for _ in 0..4 {
            let approx = approx_math_str(&cur);
            let next = self.math_to_tex(&approx);
            if next == cur {
                break;
            }
            cur = next;
        }
        cur
    }

    fn math_char(&mut self, c: char) -> String {
        match c {
            '0'..='9' | '+' | '-' | '=' | '(' | ')' | '[' | ']' | '/' | '.' | ',' | ';' | ':' | '!' | '?' | '<' | '>'
            | '|' | '*' | '@' => c.to_string(),
            'a'..='z' | 'A'..='Z' => c.to_string(),
            '\'' | '′' => "'".into(),
            '#' => "\\#".into(),
            '$' => "\\$".into(),
            '%' => "\\%".into(),
            '&' => "\\&".into(),
            '{' => "\\{".into(),
            '}' => "\\}".into(),
            '\\' => "\\backslash".into(),
            '~' => "\\sim".into(),
            '^' => "\\wedge".into(),
            '_' => "\\_".into(),
            '"' => "\\text{\\textquotedbl}".into(),
            '`' => "\\text{\\textasciigrave}".into(),
            '\u{a0}' => "~".into(),
            '\t' => "\\quad".into(),
            '°' => "^{\\circ}".into(),
            c if c.is_control() => String::new(),
            c => {
                if let Some(cmd) = math_cmd_for(c) {
                    return format!("\\{cmd}");
                }
                if let Some((style, b)) = unstyle_letter(c) {
                    return match style {
                        "bb" | "cal" | "frak" => format!("\\math{style}{{{b}}}"),
                        _ => b.to_string(),
                    };
                }
                let t = self.esc(&c.to_string(), false);
                format!("\\text{{{t}}}")
            }
        }
    }

    fn math_to_tex(&mut self, t: &str) -> String {
        let chars: Vec<char> = t.chars().filter(|c| !c.is_control() || *c == '\t').collect();
        let mut out = String::new();
        let mut i = 0;
        let mut after_func = false;
        let mut guard = 0usize;
        while i < chars.len() {
            guard += 1;
            if guard > 1_000_000 {
                break;
            }
            let c = chars[i];
            if let Some(&next) = chars.get(i + 1) {
                if let Some(acc) = comb_cmd(next) {
                    let inner = self.math_char(c);
                    let piece = if acc == "not" { format!("\\not{inner}") } else { format!("\\{acc}{{{inner}}}") };
                    append_guarded(&mut out, &piece);
                    after_func = false;
                    i += 2;
                    continue;
                }
            }
            if c == ' ' {
                let prev = if i > 0 { chars[i - 1] } else { ' ' };
                let next = chars.get(i + 1).copied().unwrap_or(' ');
                let skip = i == 0 || i + 1 == chars.len() || prev == ' ' || is_math_op(prev) || is_math_op(next) || after_func;
                if !skip {
                    out.push_str("\\ ");
                }
                after_func = false;
                i += 1;
                continue;
            }
            if c.is_ascii_alphabetic() {
                let start = i;
                while i < chars.len() && chars[i].is_ascii_alphabetic() {
                    i += 1;
                }
                let w: String = chars[start..i].iter().collect();
                if FUNCS.contains(&w.as_str()) {
                    append_guarded(&mut out, &format!("\\{w}"));
                    after_func = true;
                } else {
                    append_guarded(&mut out, &w);
                    after_func = false;
                }
                continue;
            }
            after_func = false;
            if c == '^' || c == '_' {
                match chars.get(i + 1) {
                    Some('(') => {
                        if let Some(j) = match_paren(&chars, i + 1) {
                            let inner: String = chars[i + 2..j].iter().collect();
                            let tex = self.math_to_tex(&inner);
                            out.push_str(&format!("{c}{{{tex}}}"));
                            i = j + 1;
                            continue;
                        }
                        out.push_str(&format!("{c}{{(}}"));
                        i += 2;
                        continue;
                    }
                    Some(&n) if n != ' ' => {
                        let tex = self.math_to_tex(&n.to_string());
                        out.push_str(&format!("{c}{{{tex}}}"));
                        i += 2;
                        continue;
                    }
                    _ => {
                        append_guarded(&mut out, if c == '^' { "\\wedge" } else { "\\_" });
                        i += 1;
                        continue;
                    }
                }
            }
            if sup_base(c).is_some() || sub_base(c).is_some() {
                let sup = sup_base(c).is_some();
                let mut bases = String::new();
                while i < chars.len() {
                    let b = if sup { sup_base(chars[i]) } else { sub_base(chars[i]) };
                    match b {
                        Some(b) => bases.push(b),
                        None => break,
                    }
                    i += 1;
                }
                let tex = self.math_to_tex(&bases);
                out.push_str(&format!("{}{{{tex}}}", if sup { '^' } else { '_' }));
                continue;
            }
            if matches!(c, '√' | '∛' | '∜') {
                let root = match c {
                    '∛' => "[3]",
                    '∜' => "[4]",
                    _ => "",
                };
                match root_operand(&chars, i + 1) {
                    Some((op, end)) => {
                        let tex = self.math_to_tex(&op);
                        append_guarded(&mut out, &format!("\\sqrt{root}{{{tex}}}"));
                        i = end;
                    }
                    None => {
                        append_guarded(&mut out, if root.is_empty() { "\\surd" } else { "\\sqrt{root}{}" });
                        i += 1;
                    }
                }
                continue;
            }
            let piece = self.math_char(c);
            append_guarded(&mut out, &piece);
            i += 1;
        }
        out
    }

    fn lists(&mut self, items: &[Cow<Block>]) -> String {
        let mut out = String::new();
        let mut stack: Vec<&'static str> = Vec::new();
        for b in items {
            let Block::List { list, level, runs } = b.as_ref() else { continue };
            let env = if list == "ol" { "enumerate" } else { "itemize" };
            let lvl = (*level as usize).min(3).min(stack.len());
            while stack.len() > lvl + 1 {
                let e = stack.pop().unwrap_or("itemize");
                out.push_str(&format!("{}\\end{{{e}}}\n", "  ".repeat(stack.len())));
            }
            if stack.len() == lvl + 1 && stack[lvl] != env {
                let e = stack.pop().unwrap_or("itemize");
                out.push_str(&format!("{}\\end{{{e}}}\n", "  ".repeat(stack.len())));
            }
            if stack.len() == lvl {
                out.push_str(&format!("{}\\begin{{{env}}}\n", "  ".repeat(stack.len())));
                stack.push(env);
            }
            let t = self.para_tex(&normalize_runs(runs.clone()), Ctx::Para);
            out.push_str(&"  ".repeat(stack.len()));
            out.push_str("\\item");
            if !t.is_empty() {
                out.push(' ');
                if t.starts_with('[') {
                    out.push_str("{}");
                }
                out.push_str(&t);
            }
            out.push('\n');
        }
        while let Some(e) = stack.pop() {
            out.push_str(&format!("{}\\end{{{e}}}\n", "  ".repeat(stack.len())));
        }
        out.trim_end().to_string()
    }

    fn cell_tex(&mut self, blocks: &[Block], depth: u32) -> String {
        let mut parts = Vec::new();
        for b in blocks {
            let Some(b) = resolve_block(b) else { continue };
            match b.as_ref() {
                Block::Table { rows } if depth < 8 => {
                    for row in rows {
                        for cell in row {
                            let t = self.cell_tex(&cell.blocks, depth + 1);
                            if !t.is_empty() {
                                parts.push(t);
                            }
                        }
                    }
                }
                Block::Img { alt, .. } => {
                    let alt = alt.as_deref().map(|a| a.split_whitespace().collect::<Vec<_>>().join(" ")).filter(|a| !a.is_empty());
                    let t = format!("[Image: {}]", alt.unwrap_or_else(|| "picture".into()));
                    parts.push(self.esc(&t, false));
                }
                other => {
                    let t = self.para_tex(&normalize_runs(other.para_runs().to_vec()), Ctx::Cell);
                    if !t.is_empty() {
                        parts.push(t);
                    }
                }
            }
        }
        parts.join(" ")
    }

    fn table(&mut self, rows: &[Vec<TableCell>]) -> Option<String> {
        let rows: Vec<&Vec<TableCell>> = rows.iter().filter(|r| !r.is_empty()).collect();
        if rows.is_empty() {
            return None;
        }
        let span = |c: &TableCell| c.colspan.unwrap_or(1).clamp(1, 64) as usize;
        let ncols = rows.iter().map(|r| r.iter().map(span).sum::<usize>()).max().unwrap_or(1).clamp(1, 64);
        let long = rows.len() > 30;
        if long {
            self.longtable = true;
        }
        let env = if long { "longtable" } else { "tabular" };
        let mut out = format!("\\begin{{{env}}}{{|{}}}\n\\hline\n", "l|".repeat(ncols));
        for row in rows {
            let mut cells = Vec::new();
            let mut used = 0usize;
            for cell in row {
                if used >= ncols {
                    break;
                }
                let s = span(cell).min(ncols - used);
                let t = self.cell_tex(&cell.blocks, 0);
                if s > 1 {
                    let spec = if used == 0 { "|l|" } else { "l|" };
                    cells.push(format!("\\multicolumn{{{s}}}{{{spec}}}{{{t}}}"));
                } else {
                    cells.push(t);
                }
                used += s;
            }
            while used < ncols {
                cells.push(String::new());
                used += 1;
            }
            out.push_str(&cells.join(" & "));
            out.push_str(" \\\\\n\\hline\n");
        }
        out.push_str(&format!("\\end{{{env}}}"));
        Some(out)
    }
}

/// URL for \href: escape TeX specials, percent-encode what hyperref cannot take.
fn url_tex(u: &str) -> String {
    let mut s = String::with_capacity(u.len());
    for c in u.chars() {
        match c {
            '%' => s.push_str("\\%"),
            '#' => s.push_str("\\#"),
            '\\' | '{' | '}' | ' ' | '"' | '^' | '<' | '>' | '|' | '`' | '$' => s.push_str(&format!("%{:02X}", c as u32)),
            c if c.is_control() => {}
            c if !c.is_ascii() => {
                let mut buf = [0u8; 4];
                for b in c.encode_utf8(&mut buf).bytes() {
                    s.push_str(&format!("%{b:02X}"));
                }
            }
            c => s.push(c),
        }
    }
    s
}

// ============================ compiling ============================

#[derive(Clone, Copy, PartialEq, Debug)]
enum EngineKind {
    Tectonic,
    Tex,
}

fn engine_kind(p: &Path) -> EngineKind {
    let stem = p.file_stem().map(|s| s.to_string_lossy().to_ascii_lowercase()).unwrap_or_default();
    if stem.contains("tectonic") {
        EngineKind::Tectonic
    } else {
        EngineKind::Tex
    }
}

const ENGINES: [&str; 4] = ["tectonic", "pdflatex", "xelatex", "lualatex"];

fn exe_candidates(name: &str) -> Vec<String> {
    if cfg!(windows) && Path::new(name).extension().is_none() {
        vec![format!("{name}.exe"), name.to_string()]
    } else {
        vec![name.to_string()]
    }
}

fn search_dirs(name: &str, dirs: &[PathBuf]) -> Option<PathBuf> {
    for d in dirs {
        for n in exe_candidates(name) {
            let p = d.join(&n);
            if p.is_file() {
                return Some(p);
            }
        }
    }
    None
}

/// Engine lookup: override (path or program name), then PATH, then the
/// usual install folders.
fn find_engine(over: Option<OsString>, path: Option<OsString>, extra: &[PathBuf]) -> Option<PathBuf> {
    let path_dirs: Vec<PathBuf> =
        path.map(|p| std::env::split_paths(&p).filter(|d| !d.as_os_str().is_empty()).collect()).unwrap_or_default();
    if let Some(o) = over.filter(|o| !o.is_empty()) {
        let p = PathBuf::from(&o);
        if p.is_absolute() || p.components().count() > 1 {
            if p.is_file() {
                return Some(p);
            }
            if cfg!(windows) && p.extension().is_none() {
                let e = p.with_extension("exe");
                if e.is_file() {
                    return Some(e);
                }
            }
        } else {
            let name = o.to_string_lossy().into_owned();
            if let Some(f) = search_dirs(&name, &path_dirs).or_else(|| search_dirs(&name, extra)) {
                return Some(f);
            }
        }
    }
    for name in ENGINES {
        if let Some(f) = search_dirs(name, &path_dirs) {
            return Some(f);
        }
    }
    for name in ENGINES {
        if let Some(f) = search_dirs(name, extra) {
            return Some(f);
        }
    }
    None
}

fn texlive_dirs(root: &Path, subs: &[&str], out: &mut Vec<PathBuf>) {
    let Ok(rd) = fs::read_dir(root) else { return };
    let mut years: Vec<String> = rd
        .flatten()
        .map(|e| e.file_name().to_string_lossy().into_owned())
        .filter(|n| !n.is_empty() && n.chars().all(|c| c.is_ascii_digit()))
        .collect();
    years.sort();
    years.reverse();
    for y in years.iter().take(10) {
        for s in subs {
            out.push(root.join(y).join(s));
        }
    }
}

fn default_engine_dirs() -> Vec<PathBuf> {
    let mut v = Vec::new();
    if cfg!(windows) {
        if let Some(l) = std::env::var_os("LOCALAPPDATA") {
            v.push(PathBuf::from(l).join("Programs").join("MiKTeX").join("miktex").join("bin").join("x64"));
        }
        let mut pfs: Vec<PathBuf> = std::env::var_os("ProgramFiles").map(PathBuf::from).into_iter().collect();
        pfs.push(PathBuf::from("C:\\Program Files"));
        for pf in pfs {
            v.push(pf.join("MiKTeX").join("miktex").join("bin").join("x64"));
        }
        texlive_dirs(Path::new("C:\\texlive"), &["bin\\windows", "bin\\win64", "bin\\win32"], &mut v);
        if let Some(h) = std::env::var_os("USERPROFILE") {
            v.push(PathBuf::from(h).join(".cargo").join("bin"));
        }
    } else {
        for d in ["/Library/TeX/texbin", "/usr/bin", "/usr/local/bin", "/opt/homebrew/bin"] {
            v.push(PathBuf::from(d));
        }
        texlive_dirs(
            Path::new("/usr/local/texlive"),
            &["bin/x86_64-linux", "bin/aarch64-linux", "bin/universal-darwin", "bin/x86_64-darwin"],
            &mut v,
        );
        if let Some(h) = std::env::var_os("HOME") {
            v.push(PathBuf::from(h).join(".cargo").join("bin"));
        }
    }
    v
}

/// A private temporary folder, removed (with retries) when dropped.
struct WorkDir(PathBuf);

impl WorkDir {
    fn new(base: &Path) -> Result<Self, String> {
        static N: AtomicU64 = AtomicU64::new(0);
        let nanos = SystemTime::now().duration_since(UNIX_EPOCH).map(|d| d.subsec_nanos()).unwrap_or(0);
        for _ in 0..64 {
            let p = base.join(format!("office-latex-{}-{}-{}", std::process::id(), nanos, N.fetch_add(1, Ordering::Relaxed)));
            match fs::create_dir(&p) {
                Ok(()) => return Ok(WorkDir(p)),
                Err(e) if e.kind() == std::io::ErrorKind::AlreadyExists => continue,
                Err(e) => return Err(format!("Could not create a temporary folder for LaTeX: {e}")),
            }
        }
        Err("Could not create a temporary folder for LaTeX.".into())
    }
}

impl Drop for WorkDir {
    fn drop(&mut self) {
        for _ in 0..40 {
            if fs::remove_dir_all(&self.0).is_ok() || !self.0.exists() {
                return;
            }
            std::thread::sleep(Duration::from_millis(100));
        }
    }
}

fn read_lossy(p: &Path) -> String {
    fs::read(p).map(|b| String::from_utf8_lossy(&b[..b.len().min(4 << 20)]).into_owned()).unwrap_or_default()
}

fn needs_rerun(log: &str) -> bool {
    log.contains("Rerun to get") || log.contains("There were undefined references") || log.contains("Label(s) may have changed") || log.contains("Rerun LaTeX")
}

/// Compile in a fresh folder under `base` with `engine` (testable core of
/// compile_pdf).
fn compile_with(engine: &Path, kind: EngineKind, src: &str, timeout: Duration, base: &Path) -> Result<Vec<u8>, String> {
    if src.len() > MAX_COMPILE_SRC {
        return Err(too_large_msg());
    }
    let dir = WorkDir::new(base)?;
    fs::write(dir.0.join("main.tex"), src).map_err(|e| format!("Could not write the LaTeX file: {e}"))?;
    let deadline = Instant::now() + timeout;
    let runs = if kind == EngineKind::Tex { 2 } else { 1 };
    for run in 0..runs {
        let status = run_engine(engine, kind, &dir.0, deadline, timeout)?;
        let log = read_lossy(&dir.0.join("main.log"));
        if !status.success() {
            let out = read_lossy(&dir.0.join("engine-output.txt"));
            let s = summarize_log(&log, &out, &dir.0);
            return Err(if s.is_empty() { "The TeX engine reported an error but wrote no details.".into() } else { s });
        }
        if run + 1 < runs && !needs_rerun(&log) {
            break;
        }
    }
    match fs::read(dir.0.join("main.pdf")) {
        Ok(b) if !b.is_empty() => Ok(b),
        _ => {
            let log = read_lossy(&dir.0.join("main.log"));
            let out = read_lossy(&dir.0.join("engine-output.txt"));
            let s = summarize_log(&log, &out, &dir.0);
            Err(if s.is_empty() {
                "The TeX engine finished without producing a PDF.".into()
            } else {
                format!("The TeX engine finished without producing a PDF.\n{s}")
            })
        }
    }
}

fn run_engine(engine: &Path, kind: EngineKind, dir: &Path, deadline: Instant, timeout: Duration) -> Result<ExitStatus, String> {
    let prep = |e: std::io::Error| format!("Could not prepare the LaTeX run: {e}");
    let out = fs::File::create(dir.join("engine-output.txt")).map_err(prep)?;
    let err = out.try_clone().map_err(prep)?;
    let mut cmd = Command::new(engine);
    match kind {
        EngineKind::Tectonic => {
            cmd.arg("--untrusted").arg("--keep-logs").arg("--outdir").arg(dir).arg("main.tex");
        }
        EngineKind::Tex => {
            if engine.to_string_lossy().to_ascii_lowercase().contains("miktex") {
                cmd.arg("--disable-installer");
            }
            cmd.args(["-interaction=nonstopmode", "-halt-on-error", "-no-shell-escape", "-file-line-error", "-output-directory"])
                .arg(dir)
                .arg("main.tex");
        }
    }
    cmd.current_dir(dir)
        .stdin(Stdio::null())
        .stdout(Stdio::from(out))
        .stderr(Stdio::from(err))
        .env("openout_any", "p")
        .env("openin_any", "p")
        .env("TEXMFOUTPUT", dir)
        .env("shell_escape", "f");
    #[cfg(windows)]
    {
        use std::os::windows::process::CommandExt;
        cmd.creation_flags(0x0800_0000);
    }
    let mut child = cmd.spawn().map_err(|e| format!("Could not start the TeX engine ({}): {e}", engine.display()))?;
    loop {
        match child.try_wait() {
            Ok(Some(s)) => return Ok(s),
            Ok(None) => {
                if Instant::now() >= deadline {
                    kill_tree(&mut child);
                    return Err(format!(
                        "LaTeX compilation took longer than {} seconds and was stopped. Simplify the document (very large images, long loops or complex drawings) and try again.",
                        timeout.as_secs().max(1)
                    ));
                }
                std::thread::sleep(Duration::from_millis(40));
            }
            Err(e) => {
                kill_tree(&mut child);
                return Err(format!("The TeX engine could not be monitored: {e}"));
            }
        }
    }
}

fn kill_tree(child: &mut std::process::Child) {
    #[cfg(windows)]
    {
        use std::os::windows::process::CommandExt;
        let _ = Command::new("taskkill")
            .args(["/PID", &child.id().to_string(), "/T", "/F"])
            .stdin(Stdio::null())
            .stdout(Stdio::null())
            .stderr(Stdio::null())
            .creation_flags(0x0800_0000)
            .status();
    }
    let _ = child.kill();
    let _ = child.wait();
}

fn is_file_line_error(l: &str) -> bool {
    let Some(p) = l.find(".tex:") else { return false };
    let rest = &l[p + 5..];
    let digits = rest.chars().take_while(|c| c.is_ascii_digit()).count();
    digits > 0 && rest[digits..].starts_with(':')
}

fn is_l_line(l: &str) -> bool {
    l.strip_prefix("l.").is_some_and(|r| r.starts_with(|c: char| c.is_ascii_digit()))
}

/// Short human summary of a failed run: the "! ..." / file:line errors with
/// their "l.<n>" context lines.
fn summarize_log(log: &str, out: &str, dir: &Path) -> String {
    fn collect(src: &str) -> Vec<String> {
        let lines: Vec<&str> = src.lines().collect();
        let mut res: Vec<String> = Vec::new();
        let mut total = 0;
        let mut i = 0;
        while i < lines.len() && total < MAX_ERROR_SUMMARY {
            let l = lines[i].trim_end();
            if l.starts_with("! ") || is_file_line_error(l) || l.starts_with("error:") {
                if res.last().map(String::as_str) != Some(l) {
                    total += l.len();
                    res.push(l.to_string());
                }
                for j in i + 1..(i + 10).min(lines.len()) {
                    let lj = lines[j].trim_end();
                    if lj.starts_with("! ") || is_file_line_error(lj) {
                        break;
                    }
                    if is_l_line(lj) {
                        total += lj.len();
                        res.push(lj.to_string());
                        i = j;
                        break;
                    }
                }
            }
            i += 1;
        }
        res
    }
    let mut lines = collect(log);
    if lines.is_empty() {
        lines = collect(out);
    }
    if lines.is_empty() {
        let src = if out.trim().is_empty() { log } else { out };
        let all: Vec<&str> = src.lines().map(str::trim_end).filter(|l| !l.is_empty()).collect();
        lines = all[all.len().saturating_sub(12)..].iter().map(|s| s.to_string()).collect();
    }
    let mut s = lines.join("\n");
    let d = dir.to_string_lossy().to_string();
    for v in [format!("{d}\\"), format!("{d}/"), d.replace('\\', "/") + "/", d.clone()] {
        if !v.is_empty() {
            s = s.replace(&v, "");
        }
    }
    s = s.replace("./main.tex", "main.tex").replace(".\\main.tex", "main.tex");
    if s.chars().count() > MAX_ERROR_SUMMARY {
        s = s.chars().take(MAX_ERROR_SUMMARY).collect::<String>() + "…";
    }
    s
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Compact text view of a document for assertions.
    fn dump(doc: &WordDoc) -> String {
        fn runs_s(runs: &[Run]) -> String {
            runs.iter()
                .map(|r| {
                    let mut f = String::new();
                    if r.b { f.push('b'); }
                    if r.i { f.push('i'); }
                    if r.u { f.push('u'); }
                    if r.s { f.push('s'); }
                    if r.sup { f.push('^'); }
                    if r.sub { f.push('_'); }
                    if r.font.as_deref() == Some(MATH_FONT) { f.push('M'); }
                    if r.font.as_deref() == Some(MONO_FONT) { f.push('T'); }
                    if let Some(c) = &r.color { f.push_str(&format!("c{c}")); }
                    if let Some(c) = &r.hl { f.push_str(&format!("h{c}")); }
                    if let Some(l) = &r.link { f.push_str(&format!("@{l}")); }
                    if f.is_empty() { r.t.clone() } else { format!("<{f}:{}>", r.t) }
                })
                .collect::<String>()
                .replace('\n', "\\n")
        }
        fn rec(blocks: &[Block], out: &mut Vec<String>) {
            for b in blocks {
                let s = match b {
                    Block::P { runs, align } => format!("p{}: {}", align.as_deref().map(|a| format!("[{a}]")).unwrap_or_default(), runs_s(runs)),
                    Block::H1 { runs, align } => format!("h1{}: {}", align.as_deref().map(|a| format!("[{a}]")).unwrap_or_default(), runs_s(runs)),
                    Block::H2 { runs, align } => format!("h2{}: {}", align.as_deref().map(|a| format!("[{a}]")).unwrap_or_default(), runs_s(runs)),
                    Block::H3 { runs, .. } => format!("h3: {}", runs_s(runs)),
                    Block::H4 { runs, .. } => format!("h4: {}", runs_s(runs)),
                    Block::H5 { runs, .. } => format!("h5: {}", runs_s(runs)),
                    Block::H6 { runs, .. } => format!("h6: {}", runs_s(runs)),
                    Block::Quote { runs, .. } => format!("quote: {}", runs_s(runs)),
                    Block::Code { runs, .. } => format!("code: {}", runs_s(runs)),
                    Block::List { list, level, runs } => format!("{list}{level}: {}", runs_s(runs)),
                    Block::Table { rows } => {
                        let rs: Vec<String> = rows
                            .iter()
                            .map(|r| {
                                r.iter()
                                    .map(|c| {
                                        let mut v = Vec::new();
                                        rec(&c.blocks, &mut v);
                                        let t = v.join(" ").replace("p: ", "");
                                        match c.colspan { Some(n) => format!("{t}{{{n}}}"), None => t }
                                    })
                                    .collect::<Vec<_>>()
                                    .join(" | ")
                            })
                            .collect();
                        format!("table: {}", rs.join(" / "))
                    }
                    Block::Img { alt, .. } => format!("img: {}", alt.clone().unwrap_or_default()),
                    Block::PageBreak => "pagebreak".into(),
                    Block::Other(v) => format!("other: {v}"),
                };
                out.push(s);
            }
        }
        let mut v = Vec::new();
        rec(&doc.blocks, &mut v);
        v.join("\n")
    }

    fn warnings(doc: &WordDoc) -> Vec<String> {
        doc.meta.get("warnings").and_then(Value::as_array).map(|a| a.iter().filter_map(|v| v.as_str().map(String::from)).collect()).unwrap_or_default()
    }

    #[test]
    #[ignore]
    fn show_fixtures() {
        for f in ["article", "report", "beamer", "homework"] {
            let src = std::fs::read_to_string(format!("tests/fixtures/latex/{f}.tex")).unwrap();
            let d = parse_latex(&src);
            println!("==== {f} meta={:?}\n{}\n", d.meta, dump(&d));
            println!("---- written:\n{}", write_latex(&d));
        }
    }

    fn p(src: &str) -> WordDoc {
        parse_latex(src)
    }

    fn d(src: &str) -> String {
        dump(&p(src))
    }

    fn body(src: &str) -> String {
        d(&format!("\\documentclass{{article}}\n\\begin{{document}}\n{src}\n\\end{{document}}\n"))
    }

    fn fixture(name: &str) -> String {
        let path = format!("{}/tests/fixtures/latex/{name}.tex", env!("CARGO_MANIFEST_DIR"));
        std::fs::read_to_string(path).unwrap()
    }

    fn run(t: &str) -> Run {
        Run::text(t)
    }

    fn para(runs: Vec<Run>) -> Block {
        Block::P { runs, align: None }
    }

    fn doc(blocks: Vec<Block>) -> WordDoc {
        WordDoc { kind: "word".into(), meta: Map::new(), blocks }
    }

    /// Unescaped braces balance and never go negative.
    fn tex_balanced(s: &str) -> bool {
        // verbatim bodies are not tokenized by TeX
        let mut stripped = String::new();
        let mut rest = s;
        while let Some(a) = rest.find("\\begin{verbatim}") {
            stripped.push_str(&rest[..a]);
            match rest[a..].find("\\end{verbatim}") {
                Some(b) => rest = &rest[a + b + 14..],
                None => return false,
            }
        }
        stripped.push_str(rest);
        let s = stripped.as_str();
        let mut depth = 0i64;
        let mut chars = s.chars().peekable();
        let mut in_comment = false;
        while let Some(c) = chars.next() {
            if in_comment {
                if c == '\n' {
                    in_comment = false;
                }
                continue;
            }
            match c {
                '\\' => {
                    chars.next();
                }
                '%' => in_comment = true,
                '{' => depth += 1,
                '}' => {
                    depth -= 1;
                    if depth < 0 {
                        return false;
                    }
                }
                _ => {}
            }
        }
        depth == 0
    }

    fn assert_stable(doc: &WordDoc) {
        let w1 = write_latex(doc);
        let w2 = write_latex(&parse_latex(&w1));
        if w1 != w2 {
            let a: Vec<&str> = w1.lines().collect();
            let b: Vec<&str> = w2.lines().collect();
            for k in 0..a.len().max(b.len()) {
                if a.get(k) != b.get(k) {
                    panic!("round trip not stable at line {k}:\n W1: {:?}\n W2: {:?}\n--- W1 ---\n{w1}", a.get(k), b.get(k));
                }
            }
        }
        assert!(tex_balanced(&w1), "unbalanced braces:\n{w1}");
    }

    // ---------------- parse: structure ----------------

    #[test]
    fn headings_follow_the_article_class() {
        let s = body("\\section{A}\\subsection*{B}\\subsubsection{C}\\paragraph{D}\\subparagraph{E}\\part{P}");
        assert_eq!(s, "h1: A\nh2: B\nh3: C\nh4: D\nh5: E\nh1: P");
    }

    #[test]
    fn headings_shift_with_chapters() {
        let s = d("\\documentclass{report}\\begin{document}\\chapter{A}\\section{B}\\subsection{C}\\subsubsection{D}\\paragraph{E}\\end{document}");
        assert_eq!(s, "h1: A\nh2: B\nh3: C\nh4: D\nh5: E");
        let s = d("\\documentclass[12pt]{book}\\begin{document}\\part{P}\\chapter*{Pref}\\end{document}");
        assert_eq!(s, "h1: P\nh1: Pref");
    }

    #[test]
    fn paragraphs_split_on_blank_lines_and_comments_are_dropped() {
        let s = body("First line\ncontinues % a comment\nhere.\n\n  \nSecond 50\\% para.%\nGlued\n\n\nThird.");
        assert_eq!(s, "p: First line continues here.\np: Second 50% para.Glued\np: Third.");
        // a blank line right after a comment still ends the paragraph
        assert_eq!(body("a%c\n\nb"), "p: a\np: b");
        // \par and text without a document environment
        assert_eq!(d("x\\par y"), "p: x\np: y");
    }

    #[test]
    fn preamble_text_is_ignored_and_meta_is_read() {
        let doc = p("\\documentclass{article}\n\\title{My \\textbf{Title}}\\author{A. Author \\and B. Writer}\nstray preamble text\n\\begin{document}\nBody\n\\end{document}\ntrailing text");
        assert_eq!(dump(&doc), "p: Body");
        assert_eq!(doc.meta["title"], "My Title");
        assert_eq!(doc.meta["author"], "A. Author, B. Writer");
    }

    #[test]
    fn maketitle_emits_title_and_author_date() {
        let s = d("\\title{T}\\author{Me}\\date{2024-01-02}\\begin{document}\\maketitle\nText\\end{document}");
        assert_eq!(s, "h1[center]: T\np[center]: Me\\n2024-01-02\np: Text");
        let s = d("\\title{T}\\author{Me}\\date{}\\begin{document}\\maketitle\\end{document}");
        assert_eq!(s, "h1[center]: T\np[center]: Me");
    }

    #[test]
    fn today_is_a_real_date() {
        let s = body("\\today");
        let year = 1970 + SystemTime::now().duration_since(UNIX_EPOCH).unwrap().as_secs() / 31_556_952;
        assert!(s.contains(&year.to_string()) || s.contains(&(year + 1).to_string()), "{s}");
        assert!(["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"].iter().any(|m| s.contains(m)));
    }

    // ---------------- parse: inline formatting ----------------

    #[test]
    fn formatting_nests_and_declarations_are_scoped() {
        assert_eq!(body("\\textbf{a \\textit{b} c} d"), "p: <b:a ><bi:b><b: c> d");
        assert_eq!(body("{\\bf x} {\\itshape y \\bfseries z} w"), "p: <b:x> <i:y ><bi:z> w");
        assert_eq!(body("\\emph{a \\emph{b}}"), "p: <i:a >b");
        assert_eq!(body("\\underline{u}\\sout{s}\\texttt{t}H\\textsubscript{2}O x\\textsuperscript{n}"), "p: <u:u><s:s><T:t>H<_:2>O x<^:n>");
        assert_eq!(body("\\textbf{\\textnormal{n}}"), "p: n");
    }

    #[test]
    fn colors_and_highlights() {
        assert_eq!(body("\\textcolor{red}{r} {\\color{blue} b}"), "p: <c#FF0000:r> <c#0000FF:b>");
        assert_eq!(body("\\textcolor[HTML]{00ff00}{g} \\textcolor{red!50}{pink} \\textcolor{red!50!black}{dark}"), "p: <c#00FF00:g> <c#FF8080:pink> <c#800000:dark>");
        assert_eq!(body("\\colorbox{yellow}{h} \\hl{y}"), "p: <h#FFFF00:h> <h#FFFF00:y>");
        let s = d("\\definecolor{mine}{rgb}{0,0.5,1}\\colorlet{alias}{mine!50}\\begin{document}\\textcolor{mine}{a}\\textcolor{alias}{b}\\textcolor{nosuch}{c}\\end{document}");
        assert_eq!(s, "p: <c#0080FF:a><c#80C0FF:b>c");
    }

    #[test]
    fn links() {
        assert_eq!(body("\\href{https://a.b/c?d=1\\&e=2\\#top}{the \\textbf{site}}"), "p: <@https://a.b/c?d=1&e=2#top:the ><b@https://a.b/c?d=1&e=2#top:site>");
        assert_eq!(body("\\url{http://x.org/a%20b_c}"), "p: <@http://x.org/a%20b_c:http://x.org/a%20b_c>");
        assert_eq!(body("\\href{javascript:alert(1)}{bad}"), "p: bad");
    }

    #[test]
    fn special_characters_quotes_dashes_and_accents() {
        assert_eq!(body("\\% \\& \\# \\_ \\$ \\{ \\} \\textbackslash{} a~b"), "p: % & # _ $ { } \\ a\u{a0}b");
        assert_eq!(body("``Hi'' -- `x' --- y\\ldots{} \\LaTeX{} and \\TeX"), "p: “Hi” – ‘x’ — y… LaTeX and TeX");
        assert_eq!(body("\\'e\\`a\\^o\\\"u\\~n\\c{c}\\v{s}\\H{o}\\'{\\i}\\ss\\o"), "p: éàôüñçšőíßø");
        assert_eq!(body("\\texttt{--x 'y'}"), "p: <T:--x 'y'>");
        assert_eq!(body("\\textless{}tag\\textgreater{} \\textbar{} \\textquotedbl{}"), "p: <tag> | \"");
    }

    // ---------------- parse: blocks ----------------

    #[test]
    fn nested_lists_and_description() {
        let s = body("\\begin{itemize}\n\\item One\n\\item Two\n\\begin{enumerate}\\item A\n\n continued \\item B\\end{enumerate}\n\\item Three\\end{itemize}");
        assert_eq!(s, "ul0: One\nul0: Two\nol1: A continued\nol1: B\nul0: Three");
        let s = body("\\begin{description}\\item[Key] Value\\item[\\emph{K2}] V2\\end{description}");
        assert_eq!(s, "ul0: <b:Key> Value\nul0: <bi:K2> V2");
        let s = body("\\begin{itemize}\\item a\\begin{itemize}\\item b\\begin{itemize}\\item c\\end{itemize}\\end{itemize}\\item[--] d\\end{itemize}");
        assert_eq!(s, "ul0: a\nul1: b\nul2: c\nul0: – d");
        // empty item and text after a nested list
        let s = body("\\begin{enumerate}\\item\\item x\\begin{itemize}\\item y\\end{itemize} tail\\end{enumerate}");
        assert_eq!(s, "ol0: \nol0: x\nul1: y\np: tail");
    }

    #[test]
    fn quote_center_flushright_and_pagebreak() {
        let s = body("\\begin{quote}Q1\n\nQ2\\end{quote}\\begin{center}C\\\\D\\end{center}\\newpage\\begin{flushright}R\\end{flushright}{\\centering X\\par}");
        assert_eq!(s, "quote: Q1\nquote: Q2\np[center]: C\\nD\npagebreak\np[right]: R\np[center]: X");
    }

    #[test]
    fn tables_with_multicolumn() {
        let s = body("\\begin{tabular}{|l|c|r|}\\hline\nA & \\textbf{B} & C \\\\ \\hline\n\\multicolumn{2}{|c|}{wide} & z \\\\\n & & \\\\\n1 & 2 \\\\\n\\hline\n\\end{tabular}");
        assert_eq!(s, "table: A | <b:B> | C / wide{2} | z /  |  |  / 1 | 2");
        let doc = p("\\begin{tabular}{ll}a&b\\end{tabular}");
        let Block::Table { rows } = &doc.blocks[0] else { panic!() };
        assert_eq!(rows.len(), 1);
        assert_eq!(rows[0][1].colspan, None);
        // tabularx, booktabs rules and a longtable caption row
        let s = body("\\begin{tabularx}{\\textwidth}{lX}\\toprule a & b\\\\\\midrule c & d\\\\\\bottomrule\\end{tabularx}");
        assert_eq!(s, "table: a | b / c | d");
    }

    #[test]
    fn figures_images_and_captions() {
        let doc = p("\\begin{document}\\begin{figure}\\centering\\includegraphics[width=3cm]{img/cat.png}\\caption{A cat.}\\label{f}\\end{figure}See Fig.~\\ref{f}.\\begin{tikzpicture}\\draw (0,0)--(1,1);\\end{tikzpicture}\\end{document}");
        assert_eq!(dump(&doc), "p[center]: [Image: img/cat.png]\np[center]: Figure 1: A cat.\np: See Fig.\u{a0}1.\np: [Image: drawing]");
        let w = warnings(&doc);
        assert!(w.iter().any(|x| x.contains("not embedded")), "{w:?}");
        assert!(w.iter().any(|x| x.contains("TikZ")), "{w:?}");
    }

    #[test]
    fn verbatim_is_one_code_block_per_line() {
        let s = body("Text\n\\begin{verbatim}\n  if (a < b) { x = 100%; } \\end\n\n\\textbf{not bold}\n\\end{verbatim}\nAfter \\verb|a_b{}%| and \\verb+x+.");
        assert_eq!(s, "p: Text\ncode:   if (a < b) { x = 100%; } \\end\ncode: \ncode: \\textbf{not bold}\np: After <T:a_b{}%> and <T:x>.");
        let s = body("\\begin{lstlisting}[language=Python, caption={A}]\nprint(1)\n\\end{lstlisting}\\begin{minted}{rust}\nfn main() {}\n\\end{minted}\\begin{comment}\nhidden\n\\end{comment}");
        assert_eq!(s, "code: print(1)\ncode: fn main() {}");
    }

    #[test]
    fn footnotes_become_notes() {
        let s = body("A\\footnote{First \\emph{note}.} B\\footnote{Second.}");
        assert_eq!(s, "p: A<^:[1]> B<^:[2]>\nh1: Notes\np: <^:[1]> First <i:note>.\np: <^:[2]> Second.");
    }

    #[test]
    fn labels_and_refs_resolve() {
        let src = "See \\ref{sec:b}, \\ref{fig:x}, \\ref{tab:t}, \\eqref{eq:e}, \\autoref{fig:x}, \\ref{nope}.\n\n\\section{A}\\section{B}\\label{sec:b}\\subsection{C}\\label{sec:c}\\ref{sec:c}\n\\begin{figure}\\caption{X}\\label{fig:x}\\end{figure}\\begin{table}\\caption{T}\\label{tab:t}\\end{table}\\begin{equation}a=b\\label{eq:e}\\end{equation}\\nameref{sec:b}";
        let s = body(src);
        assert!(s.starts_with("p: See 2, 1, 1, (1), Figure 1, ??."), "{s}");
        assert!(s.contains("p: 2.1"), "{s}");
        assert!(s.contains("p[center]: <M:a = b> (1)"), "{s}");
        assert!(s.ends_with("p: B"), "{s}");
        // report: chapter-relative numbers and appendix letters
        let s = d("\\documentclass{report}\\begin{document}\\chapter{X}\\section{Y}\\label{y}\\begin{figure}\\caption{F}\\label{f}\\end{figure}\\appendix\\chapter{Z}\\label{z}\\ref{y} \\ref{f} \\ref{z}\\end{document}");
        assert!(s.ends_with("p: 1.1 1.1 A"), "{s}");
    }

    #[test]
    fn citations_and_bibliography() {
        let s = body("As \\cite{b,a} and \\cite[p.~5]{a} and \\citep{zz}.\\begin{thebibliography}{9}\\bibitem{a} Alpha.\\bibitem[Bx]{b} Beta.\\end{thebibliography}");
        assert_eq!(s, "p: As [Bx, 1] and [1, p.\u{a0}5] and [zz].\nh1: References\np: [1] Alpha.\np: [Bx] Beta.");
        let doc = p("\\begin{document}\\cite{k}\\bibliography{refs}\\end{document}");
        assert!(warnings(&doc).iter().any(|w| w.contains(".bib")));
    }

    #[test]
    fn math_unicode_approximation() {
        let cases = [
            ("$x^2 + y_i$", "x²+yᵢ"),
            ("$\\alpha\\beta\\Gamma$", "αβΓ"),
            ("$\\frac{a}{b}$", "a/b"),
            ("$\\frac{a+b}{c}$", "(a+b)/c"),
            ("$\\sqrt{x} + \\sqrt{x+1} + \\sqrt[3]{8}$", "√x+√(x+1)+∛8"),
            ("$\\sum_{i=1}^{n} x_i \\leq \\infty$", "∑ᵢ₌₁ⁿxᵢ ≤ ∞"),
            ("$a \\neq b \\approx c \\times d \\cdot e \\pm f$", "a ≠ b ≈ c×d·e±f"),
            ("$x \\to y \\Rightarrow z \\in A, \\forall x \\exists y$", "x → y ⇒ z ∈ A,∀x∃y"),
            ("$\\partial f / \\nabla g$", "∂f/∇g"),
            ("$\\sin x + \\log(y)$", "sin x+log(y)"),
            ("$\\mathbb{R}^n \\mathcal{L}$", "ℝⁿℒ"),
            ("$x^{10} + e^{i\\pi}$", "x¹⁰+e^(iπ)"),
            ("$\\hat{x}\\bar{y}$", "x\u{302}y\u{304}"),
            ("$\\text{if } x > 0$", "if x > 0"),
            ("$f'(x)$", "f′(x)"),
            ("$90^\\circ$", "90°"),
            ("$\\left( \\frac{1}{2} \\right)$", "(1/2)"),
            ("$\\left. x \\right|$", "x|"),
            ("$a \\not= b$", "a ≠ b"),
            ("\\(\\infty\\)", "∞"),
            ("$\\binom{n}{k}$", "C(n,k)"),
            ("$\\operatorname{rank} A$", "rank A"),
        ];
        let mut bad = Vec::new();
        for (src, want) in cases {
            let s = body(src);
            if s != format!("p: <M:{want}>") {
                bad.push(format!("{src} => {s:?} (want {want:?})"));
            }
        }
        assert!(bad.is_empty(), "{}", bad.join("\n"));
    }

    #[test]
    fn display_math_is_centered_one_paragraph_per_line() {
        let s = body("Before\n\\[ E = mc^2 \\]\nafter $$a+b$$ end\n\\begin{align}\nx &= 1 \\\\\ny &= 2 \\nonumber \\\\\nz &= 3 \\tag{*}\n\\end{align}\\begin{equation*}q\\end{equation*}");
        assert_eq!(
            s,
            "p: Before\np[center]: <M:E = mc²>\np: after\np[center]: <M:a+b>\np: end\np[center]: <M:x = 1> (1)\np[center]: <M:y = 2>\np[center]: <M:z = 3> (*)\np[center]: <M:q>"
        );
        let s = body("\\begin{equation}\\begin{split}a &= b\\\\ &= c\\end{split}\\end{equation}");
        assert_eq!(s, "p[center]: <M:a = b>\np[center]: <M:= c> (1)");
    }

    #[test]
    fn macros_expand_with_arguments_and_defaults() {
        let s = d("\\newcommand{\\hi}[2][Hello]{#1, #2!}\\newcommand\\x{X}\\def\\pair#1#2{(#1;#2)}\\let\\y\\x\\renewcommand{\\x}{Z}\\begin{document}\\hi{Bob} \\hi[Bye]{Al} \\pair{a}{b} \\x\\y\\end{document}");
        assert_eq!(s, "p: Hello, Bob! Bye, Al! (a;b) ZX");
        let s = d("\\newenvironment{boxed}[1]{\\textbf{#1:} \\begin{itshape}}{\\end{itshape}}\\begin{document}\\begin{boxed}{Note}body text\\end{boxed}\\end{document}");
        assert_eq!(s, "p: <b:Note:> <i:body text>");
        // DeclareMathOperator
        assert_eq!(d("\\DeclareMathOperator{\\Tr}{Tr}\\begin{document}$\\Tr A$\\end{document}"), "p: <M:Tr A>");
    }

    #[test]
    fn recursive_macros_are_capped() {
        let t = Instant::now();
        let doc = p("\\newcommand{\\a}{\\a\\a}\\def\\b{x\\b}\\begin{document}Start \\a \\b End\\end{document}");
        assert!(t.elapsed() < Duration::from_secs(20));
        let s = dump(&doc);
        assert!(s.starts_with("p: Start"), "{}", &s[..s.len().min(200)]);
        assert!(warnings(&doc).iter().any(|w| w.contains("expand too deeply")), "{:?}", warnings(&doc));
    }

    #[test]
    fn unknown_commands_keep_text_and_give_one_warning() {
        let doc = p("\\begin{document}\\foo{kept} \\baz[opt]{also} \\foo{again} \\begin{weird}inside\\end{weird}\\end{document}");
        assert_eq!(dump(&doc), "p: kept also again inside");
        let w = warnings(&doc);
        let unk: Vec<&String> = w.iter().filter(|x| x.contains("not supported")).collect();
        assert_eq!(unk.len(), 1, "{w:?}");
        assert!(unk[0].contains("\\foo") && unk[0].contains("\\baz") && unk[0].contains("\\begin{weird}"), "{}", unk[0]);
        // known commands do not warn
        let doc = p("\\begin{document}\\noindent\\textbf{x}\\vspace{1cm}\\hfill\\end{document}");
        assert!(warnings(&doc).is_empty(), "{:?}", warnings(&doc));
    }

    #[test]
    fn theorems_proofs_and_abstract() {
        let s = d("\\newtheorem{thm}{Theorem}\\newtheorem{lem}[thm]{Lemma}\\begin{document}\\begin{abstract}Short.\\end{abstract}\\begin{thm}[Big]\\label{t}True.\\end{thm}\\begin{lem}Also.\\end{lem}\\begin{definition}Def.\\end{definition}\\begin{proof}Easy.\\end{proof}See \\ref{t}.\\end{document}");
        assert_eq!(
            s,
            "h2[center]: Abstract\np: Short.\np: <b:Theorem 1> (Big)<b:.> True.\np: <b:Lemma 2.> Also.\np: <b:Definition 1.> Def.\np: <i:Proof.> Easy. ∎\np: See 1."
        );
    }

    #[test]
    fn beamer_frames() {
        let s = d("\\documentclass{beamer}\\title{Deck}\\begin{document}\\begin{frame}\\titlepage\\end{frame}\\begin{frame}<1->[t]{Title}{Sub}\\begin{itemize}\\item<2-> A\\item<+-| alert@+> B\\end{itemize}\\end{frame}\\frame{\\frametitle{Old}x}\\begin{frame}\\frametitle<2>{T2}\\only<3>{y}\\pause z\\end{frame}\\end{document}");
        let today = today_string();
        assert_eq!(s, format!("h1[center]: Deck\np[center]: {today}\nh2: Title\nh3: Sub\nul0: A\nul0: B\nh2: Old\np: x\nh2: T2\np: yz"));
    }

    #[test]
    fn verb_and_url_inside_macros_do_not_break() {
        let s = d("\\newcommand{\\code}[1]{\\texttt{#1}}\\newcommand{\\site}{\\url{http://a.b/x_y}}\\begin{document}\\code{a_b} \\site\\end{document}");
        assert_eq!(s, "p: <T:a_b> <@http://a.b/x_y:http://a.b/x_y>");
    }

    #[test]
    fn siunitx_and_misc_text_commands() {
        assert_eq!(body("\\SI{3}{\\kilo\\meter\\per\\second} \\enquote{q} \\MakeUppercase{up}"), "p: 3\u{a0}km/s “q” UP");
    }

    // ---------------- write ----------------

    #[test]
    fn write_escapes_specials() {
        let w = write_latex(&doc(vec![para(vec![run("# $ % & _ { } ~ ^ \\ < > | \" ` a--b x''y")])]));
        assert!(w.contains("\\# \\$ \\% \\& \\_ \\{ \\} \\textasciitilde{} \\textasciicircum{} \\textbackslash{} \\textless{} \\textgreater{} \\textbar{} \\textquotedbl{} \\textasciigrave{} a-{}-b x'{}'y"), "{w}");
        assert!(tex_balanced(&w));
        // round trip gives back the text
        let back = parse_latex(&w);
        assert_eq!(dump(&back), "p: # $ % & _ { } ~ ^ \\ < > | \" ` a--b x’’y");
    }

    #[test]
    fn write_preamble_only_has_needed_packages() {
        let w = write_latex(&doc(vec![para(vec![run("plain")])]));
        for pkg in ["inputenc", "fontenc", "hyperref"] {
            assert!(w.contains(pkg), "{pkg}");
        }
        for pkg in ["xcolor", "ulem", "amsmath", "longtable", "textcomp"] {
            assert!(!w.contains(pkg), "{pkg} should not be loaded:\n{w}");
        }
        let r = Run { t: "x".into(), s: true, color: Some("#112233".into()), ..Default::default() };
        let m = Run { t: "α²".into(), font: Some(MATH_FONT.into()), ..Default::default() };
        let w = write_latex(&doc(vec![para(vec![r, m, run(" 5 €")])]));
        for pkg in ["xcolor", "[normalem]{ulem}", "amsmath", "amssymb", "textcomp"] {
            assert!(w.contains(pkg), "{pkg} missing:\n{w}");
        }
        assert!(w.contains("\\textcolor[HTML]{112233}{\\sout{x}}$\\alpha^{2}$ 5 \\texteuro{}"), "{w}");
    }

    #[test]
    fn write_nested_formatting_links_and_colors() {
        let r = Run { t: "x y".into(), b: true, i: true, u: true, link: Some("https://e.com/a%20b#c".into()), hl: Some("#ffff00".into()), ..Default::default() };
        let w = write_latex(&doc(vec![para(vec![run("Go "), r])]));
        assert!(w.contains("Go \\href{https://e.com/a\\%20b\\#c}{\\colorbox[HTML]{FFFF00}{\\textbf{\\textit{\\underline{x y}}}}}"), "{w}");
        let back = parse_latex(&w);
        assert_eq!(dump(&back), "p: Go <biuh#FFFF00@https://e.com/a%20b#c:x y>");
        let bad = Run { t: "evil".into(), link: Some("javascript:alert(1)".into()), ..Default::default() };
        assert!(!write_latex(&doc(vec![para(vec![bad])])).contains("javascript"));
    }

    #[test]
    fn write_lists_open_and_close_across_levels() {
        let li = |l: &str, lvl: u32, t: &str| Block::List { list: l.into(), level: lvl, runs: vec![run(t)] };
        let w = write_latex(&doc(vec![li("ul", 0, "a"), li("ul", 2, "deep"), li("ol", 1, "n"), li("ul", 0, "b"), li("ol", 0, "[x] c"), para(vec![run("after")])]));
        let body = &w[w.find("\\begin{document}").unwrap()..];
        assert!(body.contains("\\begin{itemize}\n  \\item a\n  \\begin{itemize}\n    \\item deep\n  \\end{itemize}\n  \\begin{enumerate}\n    \\item n\n  \\end{enumerate}\n  \\item b\n\\end{itemize}\n\\begin{enumerate}\n  \\item {}[x] c\n\\end{enumerate}\n\nafter"), "{body}");
        assert_eq!(dump(&parse_latex(&w)), "ul0: a\nul1: deep\nol1: n\nul0: b\nol0: [x] c\np: after");
    }

    #[test]
    fn write_tables_pad_rows_and_use_multicolumn() {
        let cell = |t: &str, span: Option<u32>| TableCell { blocks: vec![para(vec![run(t)])], colspan: span };
        let t = Block::Table { rows: vec![vec![cell("a", None), cell("b & c", None), cell("d", None)], vec![cell("wide", Some(2))], vec![cell("x", None)]] };
        let w = write_latex(&doc(vec![t]));
        assert!(w.contains("\\begin{tabular}{|l|l|l|}\n\\hline\na & b \\& c & d \\\\\n\\hline\n\\multicolumn{2}{|l|}{wide} &  \\\\\n\\hline\nx &  &  \\\\\n\\hline\n\\end{tabular}"), "{w}");
        assert_eq!(dump(&parse_latex(&w)), "table: a | b & c | d / wide{2} |  / x |  | ");
        let rows: Vec<Vec<TableCell>> = (0..40).map(|i| vec![cell(&i.to_string(), None)]).collect();
        let w = write_latex(&doc(vec![Block::Table { rows }]));
        assert!(w.contains("\\usepackage{longtable}") && w.contains("\\begin{longtable}{|l|}"), "{w}");
    }

    #[test]
    fn write_title_code_quote_images_and_math() {
        let mut d0 = doc(vec![
            Block::H1 { runs: vec![run("My Title")], align: Some("center".into()) },
            Block::P { runs: vec![run("Me")], align: Some("center".into()) },
            Block::H1 { runs: vec![run("Intro")], align: None },
            Block::Code { runs: vec![run("let x = 1;\n  \\end{verbatim} y")], align: None },
            Block::Quote { runs: vec![run("Quoted")], align: None },
            Block::Img { src: "data:image/png;base64,AAAA".into(), alt: Some("A chart".into()), w: None, h: None },
            Block::P { runs: vec![Run { t: "x^(a+b) = √(y+1)".into(), font: Some(MATH_FONT.into()), ..Default::default() }], align: Some("center".into()) },
            Block::PageBreak,
            Block::H5 { runs: vec![run("Small")], align: None },
        ]);
        d0.meta.insert("title".into(), Value::String("My Title".into()));
        d0.meta.insert("author".into(), Value::String("Me".into()));
        let w = write_latex(&d0);
        assert!(w.contains("\\title{My Title}\n\\author{Me}\n\\date{}"), "{w}");
        assert!(w.contains("\\maketitle\n\n\\section{Intro}"), "{w}");
        assert!(w.contains("\\begin{verbatim}\nlet x = 1;\n  \\end {verbatim} y\n\\end{verbatim}"), "{w}");
        assert!(w.contains("\\begin{quote}\nQuoted\n\\end{quote}"), "{w}");
        assert!(w.contains("% image omitted: A chart\n\\fbox{Image: A chart}"), "{w}");
        assert!(w.contains("\\[ x^{a+b}=\\sqrt{y+1} \\]"), "{w}");
        assert!(w.contains("\\newpage") && w.contains("\\subparagraph{Small}"), "{w}");
        assert_stable(&d0);
    }

    #[test]
    fn write_unicode_fallbacks() {
        let w = write_latex(&doc(vec![para(vec![run("Café – α → 中文 😀 ²")])]));
        assert!(w.contains("Café -- \\ensuremath{\\alpha} \\ensuremath{\\rightarrow} 中文 😀 \\texttwosuperior{}"), "{w}");
        assert!(w.contains("\\DeclareUnicodeCharacter{4E2D}{?}") && w.contains("\\DeclareUnicodeCharacter{1F600}{?}"), "{w}");
        assert!(!w.contains("DeclareUnicodeCharacter{E9}") && !w.contains("{00E9}"), "é is native");
        assert_eq!(dump(&parse_latex(&w)), "p: Café – α → 中文 😀 ²");
    }

    // ---------------- round trip ----------------

    #[test]
    fn round_trip_is_stable_for_fixtures() {
        for f in ["article", "report", "beamer", "homework"] {
            let d0 = parse_latex(&fixture(f));
            assert_stable(&d0);
        }
    }

    #[test]
    fn round_trip_is_stable_for_a_rich_document() {
        let fmt = |t: &str, f: fn(&mut Run)| {
            let mut r = run(t);
            f(&mut r);
            r
        };
        let d0 = doc(vec![
            Block::H1 { runs: vec![run("Title with # and %")], align: None },
            para(vec![
                run("Plain "),
                fmt("bold", |r| r.b = true),
                run(" and "),
                fmt("it", |r| r.i = true),
                fmt(" mono -- 'q' ", |r| r.font = Some("Consolas".into())),
                fmt("red", |r| r.color = Some("#FF0000".into())),
                run(" “quoted” — dash… 50% off\nnext line"),
                fmt("x", |r| r.sup = true),
                fmt("2", |r| r.sub = true),
            ]),
            Block::List { list: "ul".into(), level: 0, runs: vec![run("item one")] },
            Block::List { list: "ol".into(), level: 1, runs: vec![fmt("[bracket] start", |r| r.u = true)] },
            Block::List { list: "ol".into(), level: 1, runs: vec![] },
            Block::Quote { runs: vec![run("A quote\nwith a break")], align: None },
            Block::P { runs: vec![run("Centered")], align: Some("center".into()) },
            Block::P { runs: vec![run("[starts with bracket]")], align: Some("right".into()) },
            Block::P { runs: vec![Run { t: "∑ᵢ xᵢ² ≤ ∞, sin θ + αβ".into(), font: Some(MATH_FONT.into()), ..Default::default() }], align: Some("center".into()) },
            para(vec![run("Inline "), Run { t: "a/b + √2 · π".into(), font: Some(MATH_FONT.into()), ..Default::default() }, run(" done")]),
            Block::Code { runs: vec![run("\tindented { code } % not a comment")], align: None },
            Block::Code { runs: vec![], align: None },
            Block::Code { runs: vec![run("last")], align: None },
            para(vec![run("[Image: logo.png]")]),
            Block::H2 { runs: vec![run("Abstract")], align: Some("center".into()) },
            para(vec![run("Abstract text.")]),
            Block::H6 { runs: vec![run("Deep")], align: None },
        ]);
        assert_stable(&d0);
    }

    /// Tiny deterministic PRNG for the generated tests.
    struct Rng(u64);
    impl Rng {
        fn next(&mut self) -> u64 {
            self.0 ^= self.0 << 13;
            self.0 ^= self.0 >> 7;
            self.0 ^= self.0 << 17;
            self.0
        }
        fn below(&mut self, n: usize) -> usize {
            (self.next() % n as u64) as usize
        }
        fn pick<'a, T>(&mut self, v: &'a [T]) -> &'a T {
            &v[self.below(v.len())]
        }
    }

    fn random_text(rng: &mut Rng, pool: &[&str], n: usize) -> String {
        (0..n).map(|_| *rng.pick(pool)).collect()
    }

    fn random_runs(rng: &mut Rng) -> Vec<Run> {
        const TEXT: &[&str] = &[
            "a", "b", "word", " ", " ", "  ", "#", "$", "%", "&", "_", "{", "}", "~", "^", "\\", "<", ">", "|", "\"", "'",
            "`", "-", "--", "[", "]", "*", "é", "ü", "ß", "α", "→", "“", "”", "–", "—", "…", "€", "✓", "中", "😀", "\t",
            "\n", "!", "?", ",", ".", "x²", "\u{a0}", "ı", "ﬁ",
        ];
        const MATH: &[&str] = &["x", "y", "2", "+", "-", "=", "(", ")", "α", "β", "²", "₁", "≤", "→", "∑", "√", "∞", "′", " ", "sin", "/", "ℝ", ",", "·"];
        let n = 1 + rng.below(6);
        (0..n)
            .map(|_| {
                let math = rng.below(6) == 0;
                let len = 1 + rng.below(if math { 6 } else { 8 });
                let t = if math { random_text(rng, MATH, len) } else { random_text(rng, TEXT, len) };
                let mut r = run(&t);
                if math {
                    r.font = Some(MATH_FONT.into());
                } else if rng.below(6) == 0 {
                    r.font = Some(MONO_FONT.into());
                }
                r.b = rng.below(4) == 0;
                r.i = rng.below(4) == 0;
                r.u = rng.below(8) == 0;
                r.s = rng.below(10) == 0;
                r.sup = rng.below(12) == 0;
                r.sub = !r.sup && rng.below(12) == 0;
                if rng.below(8) == 0 {
                    r.color = Some(rng.pick(&["#FF0000", "#00AA00", "#123456"]).to_string());
                }
                if rng.below(10) == 0 {
                    r.hl = Some("#FFFF00".into());
                }
                if rng.below(10) == 0 {
                    r.link = Some(rng.pick(&["https://a.b/c?d=1&e=2", "http://x.y/%41#z", "mailto:me@x.org"]).to_string());
                }
                r
            })
            .collect()
    }

    fn random_doc(rng: &mut Rng) -> WordDoc {
        let n = 1 + rng.below(12);
        let mut blocks = Vec::new();
        for _ in 0..n {
            let b = match rng.below(10) {
                0 => heading_block(1 + rng.below(6) as u32, random_runs(rng), None),
                1 => Block::Quote { runs: random_runs(rng), align: None },
                2 => Block::Code { runs: vec![run(&random_text(rng, &["x", " ", "{", "}", "\\", "%", "é", "→", "\t"], 6))], align: None },
                3 | 4 => Block::List { list: rng.pick(&["ul", "ol"]).to_string(), level: rng.below(6) as u32, runs: random_runs(rng) },
                5 => Block::Table {
                    rows: (0..1 + rng.below(3))
                        .map(|_| {
                            (0..1 + rng.below(3))
                                .map(|_| TableCell { blocks: vec![para(random_runs(rng))], colspan: (rng.below(4) == 0).then_some(2) })
                                .collect()
                        })
                        .collect(),
                },
                6 => Block::P { runs: random_runs(rng), align: Some(rng.pick(&["center", "right", "justify", "left"]).to_string()) },
                7 if rng.below(3) == 0 => Block::PageBreak,
                _ => para(random_runs(rng)),
            };
            blocks.push(b);
        }
        doc(blocks)
    }

    #[test]
    fn round_trip_is_stable_for_generated_documents() {
        let mut rng = Rng(0x9E37_79B9_7F4A_7C15);
        for _ in 0..1000 {
            let d0 = random_doc(&mut rng);
            assert_stable(&d0);
        }
    }

    #[test]
    fn written_fixtures_parse_back_to_the_same_structure() {
        let d0 = parse_latex(&fixture("article"));
        let d1 = parse_latex(&write_latex(&d0));
        let kinds = |d: &WordDoc| d.blocks.iter().map(|b| std::mem::discriminant(b)).collect::<Vec<_>>();
        assert_eq!(kinds(&d0), kinds(&d1));
        assert_eq!(d1.meta["title"], d0.meta["title"]);
    }

    // ---------------- robustness ----------------

    #[test]
    fn truncated_fixtures_never_panic() {
        for f in ["article", "report", "beamer", "homework"] {
            let src = fixture(f);
            let chars: Vec<char> = src.chars().collect();
            let step = (chars.len() / 60).max(1);
            let mut k = 0;
            while k <= chars.len() {
                let cut: String = chars[..k].iter().collect();
                let d0 = parse_inner(&cut);
                let w = write_latex(&d0);
                assert!(tex_balanced(&w), "{f} cut at {k}");
                let tail: String = chars[k..].iter().collect();
                let _ = parse_inner(&tail);
                k += step;
            }
        }
    }

    #[test]
    fn garbage_and_bombs_never_panic_or_hang() {
        let start = Instant::now();
        let deep_braces = "{".repeat(100_000) + "x" + &"}".repeat(50_000);
        let deep_envs = "\\begin{itemize}\\item x".repeat(20_000);
        let deep_math = "$\\frac{".repeat(5_000) + "1";
        let mut rng = Rng(12345);
        let noise: String = (0..20_000).map(|_| *rng.pick(&['\\', '{', '}', '$', '&', '#', '^', '_', '%', '[', ']', 'a', ' ', '\n', '~', 'é'])).collect();
        let inputs = vec![
            String::new(),
            "\\".into(),
            "\\begin".into(),
            "\\begin{".into(),
            "\\end{document}\\end{itemize}}}}".into(),
            "$$$$$".into(),
            "\\verb".into(),
            "\\verb|unterminated".into(),
            "\\begin{verbatim}never closed".into(),
            "\\newcommand".into(),
            "\\newcommand{\\a}[9]{#9#9#9#9}\\a".into(),
            "\\def\\a{\\a}\\a".into(),
            "\\def\\a#1{\\a{#1#1}}\\a x".into(),
            "\\newenvironment{e}{\\begin{e}}{}\\begin{e}".into(),
            "\\let\\a\\a\\a".into(),
            "\\href{".into(),
            "\\url{%".into(),
            "\\begin{tabular}{".into(),
            "\\multicolumn{999999999999}{l}{x}".into(),
            "\\footnote{\\footnote{\\footnote{x}}}".into(),
            "\\section{\\section{\\section{x}}}".into(),
            "\\item\\item\\bibitem".into(),
            "\\begin{frame}<9999999>".into(),
            "\\begin{thebibliography}\\bibitem".into(),
            "\\\\\\\\\\\\[".into(),
            "^_^_&&&#1#2".into(),
            "\u{E000}\u{E001}\u{E002}\\ref{\u{E000}}".into(),
            deep_braces,
            deep_envs,
            deep_math,
            noise,
            "x".repeat(2_000_000),
        ];
        for (k, s) in inputs.iter().enumerate() {
            let t0 = Instant::now();
            let d0 = parse_latex(s);
            let t1 = t0.elapsed();
            let w = write_latex(&d0);
            assert!(t1 < Duration::from_secs(30), "input {k} parse took {t1:?}");
            assert!(tex_balanced(&w), "input {:?}", &s[..s.len().min(40)]);
        }
        assert!(start.elapsed() < Duration::from_secs(120), "took {:?}", start.elapsed());
    }

    #[test]
    fn big_documents_are_capped() {
        let src = "\\documentclass{article}\\newcommand{\\w}{word word word word }\\begin{document}".to_string() + &"\\w\\w\\w\\w\n\n".repeat(20_000) + "\\end{document}";
        let d0 = parse_latex(&src);
        assert!(d0.blocks.len() >= 12_000, "{}", d0.blocks.len());
        assert!(warnings(&d0).iter().any(|w| w.contains("too deeply")), "expansion cap reached: {:?}", warnings(&d0));
    }

    // ---------------- engine detection and compilation ----------------

    fn fake_engine(dir: &Path, name: &str, script: &str) -> PathBuf {
        #[cfg(windows)]
        {
            let p = dir.join(format!("{name}.cmd"));
            std::fs::write(&p, format!("@echo off\r\n{}\r\n", script.replace('\n', "\r\n"))).unwrap();
            p
        }
        #[cfg(not(windows))]
        {
            use std::os::unix::fs::PermissionsExt;
            let p = dir.join(name);
            std::fs::write(&p, format!("#!/bin/sh\n{script}\n")).unwrap();
            std::fs::set_permissions(&p, std::fs::Permissions::from_mode(0o755)).unwrap();
            p
        }
    }

    #[test]
    fn engine_detection_prefers_the_override() {
        let tmp = tempfile::tempdir().unwrap();
        let bin = tmp.path().join("bin");
        std::fs::create_dir(&bin).unwrap();
        let exe = if cfg!(windows) { "pdflatex.exe" } else { "pdflatex" };
        let xe = if cfg!(windows) { "xelatex.exe" } else { "xelatex" };
        std::fs::write(bin.join(exe), b"x").unwrap();
        std::fs::write(bin.join(xe), b"x").unwrap();
        let path = std::env::join_paths([tmp.path().join("nothing"), bin.clone()]).unwrap();
        // PATH order: tectonic, pdflatex, xelatex, lualatex
        assert_eq!(find_engine(None, Some(path.clone()), &[]), Some(bin.join(exe)));
        // override by program name and by full path
        assert_eq!(find_engine(Some("xelatex".into()), Some(path.clone()), &[]), Some(bin.join(xe)));
        let custom = tmp.path().join("my-tex");
        std::fs::write(&custom, b"x").unwrap();
        assert_eq!(find_engine(Some(custom.clone().into_os_string()), Some(path.clone()), &[]), Some(custom));
        // a broken override falls back to the normal search
        assert_eq!(find_engine(Some(tmp.path().join("missing").into_os_string()), Some(path), &[]), Some(bin.join(exe)));
        // install folders are searched after PATH
        assert_eq!(find_engine(None, None, &[bin.clone()]), Some(bin.join(exe)));
        assert_eq!(find_engine(None, Some(OsString::new()), &[tmp.path().join("none")]), None);
        assert_eq!(engine_kind(Path::new("/x/tectonic.exe")), EngineKind::Tectonic);
        assert_eq!(engine_kind(Path::new("C:/x/lualatex.exe")), EngineKind::Tex);
    }

    #[test]
    fn compile_with_a_fake_engine_succeeds() {
        let tmp = tempfile::tempdir().unwrap();
        let work = tmp.path().join("work");
        std::fs::create_dir(&work).unwrap();
        let script = if cfg!(windows) { "echo %%PDF-1.4 fake> main.pdf\necho done> main.log\nexit /b 0" } else { "echo '%PDF-1.4 fake' > main.pdf\necho done > main.log\nexit 0" };
        let eng = fake_engine(tmp.path(), "fakelatex", script);
        let pdf = compile_with(&eng, EngineKind::Tex, "\\documentclass{article}\\begin{document}x\\end{document}", Duration::from_secs(30), &work).unwrap();
        assert!(pdf.starts_with(b"%PDF-1.4"), "{:?}", String::from_utf8_lossy(&pdf));
        assert_eq!(std::fs::read_dir(&work).unwrap().count(), 0, "temp dir removed");
    }

    #[test]
    fn compile_reruns_when_the_log_asks() {
        let tmp = tempfile::tempdir().unwrap();
        let work = tmp.path().join("work");
        std::fs::create_dir(&work).unwrap();
        let script = if cfg!(windows) {
            "if exist run1.txt (\n  echo %%PDF second> main.pdf\n  echo ok> main.log\n) else (\n  echo 1> run1.txt\n  echo %%PDF first> main.pdf\n  echo LaTeX Warning: Rerun to get cross-references right.> main.log\n)\nexit /b 0"
        } else {
            "if [ -f run1.txt ]; then echo '%PDF second' > main.pdf; echo ok > main.log; else echo 1 > run1.txt; echo '%PDF first' > main.pdf; echo 'Rerun to get cross-references right.' > main.log; fi\nexit 0"
        };
        let eng = fake_engine(tmp.path(), "rerunlatex", script);
        let pdf = compile_with(&eng, EngineKind::Tex, "x", Duration::from_secs(30), &work).unwrap();
        assert!(String::from_utf8_lossy(&pdf).contains("second"));
    }

    #[test]
    fn compile_errors_are_summarized_with_line_numbers() {
        let tmp = tempfile::tempdir().unwrap();
        let work = tmp.path().join("work");
        std::fs::create_dir(&work).unwrap();
        let script = if cfg!(windows) {
            "echo This is pdfTeX> main.log\necho ./main.tex:7: Undefined control sequence.>> main.log\necho l.7 \\foo>> main.log\necho                  bar>> main.log\necho ! Emergency stop.>> main.log\nexit /b 1"
        } else {
            "printf 'This is pdfTeX\\n./main.tex:7: Undefined control sequence.\\nl.7 \\\\foo\\n                 bar\\n! Emergency stop.\\n' > main.log\nexit 1"
        };
        let eng = fake_engine(tmp.path(), "badlatex", script);
        let err = compile_with(&eng, EngineKind::Tex, "x", Duration::from_secs(30), &work).unwrap_err();
        assert!(err.contains("main.tex:7: Undefined control sequence."), "{err}");
        assert!(err.contains("l.7 \\foo"), "{err}");
        assert!(err.contains("! Emergency stop."), "{err}");
        assert!(!err.contains("./main.tex"), "{err}");
        assert!(err.len() <= MAX_ERROR_SUMMARY + 4);
        assert_eq!(std::fs::read_dir(&work).unwrap().count(), 0);
    }

    #[test]
    fn compile_times_out_and_cleans_up() {
        let tmp = tempfile::tempdir().unwrap();
        let work = tmp.path().join("work");
        std::fs::create_dir(&work).unwrap();
        let script = if cfg!(windows) { "ping -n 30 127.0.0.1 > nul\necho %%PDF> main.pdf" } else { "sleep 30\necho '%PDF' > main.pdf" };
        let eng = fake_engine(tmp.path(), "slowlatex", script);
        let t = Instant::now();
        let err = compile_with(&eng, EngineKind::Tex, "x", Duration::from_secs(1), &work).unwrap_err();
        assert!(t.elapsed() < Duration::from_secs(15), "{:?}", t.elapsed());
        assert!(err.contains("longer than 1 seconds"), "{err}");
        assert_eq!(std::fs::read_dir(&work).unwrap().count(), 0, "temp dir removed after kill");
    }

    #[test]
    fn compile_rejects_big_sources_and_missing_engines() {
        let big = "x".repeat(MAX_COMPILE_SRC + 1);
        assert!(compile_pdf(&big).unwrap_err().contains("2 MB"));
        let tmp = tempfile::tempdir().unwrap();
        let err = compile_with(&tmp.path().join("no-such-engine"), EngineKind::Tex, "x", Duration::from_secs(5), tmp.path()).unwrap_err();
        assert!(err.contains("Could not start the TeX engine"), "{err}");
    }

    #[test]
    fn log_summary_picks_error_lines() {
        let log = "junk\n! LaTeX Error: File `nope.sty' not found.\n\nType X to quit\nl.3 \\usepackage\n          {nope}\nmore junk\n! Undefined control sequence.\nl.9 \\bad\n";
        let s = summarize_log(log, "", Path::new("/tmp/x"));
        assert_eq!(s, "! LaTeX Error: File `nope.sty' not found.\nl.3 \\usepackage\n! Undefined control sequence.\nl.9 \\bad");
        let s = summarize_log("", "error: main.tex:4: Missing $ inserted\n", Path::new("/tmp/x"));
        assert_eq!(s, "error: main.tex:4: Missing $ inserted");
        let long = "! x\n".repeat(2000);
        assert!(summarize_log(&long, "", Path::new("/t")).chars().count() <= MAX_ERROR_SUMMARY + 1);
    }
}
